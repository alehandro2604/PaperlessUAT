// ============================================================
// HR TASK MAPPER
// Shapes raw SharePoint list items into the HR task objects the
// App component renders. Extracted from app.ts  the logic here
// is independent of component state and reads the signed-in user
// from UserService.
// ============================================================
import { Injectable } from '@angular/core';
import { normalizeSharePointFileUrl } from '../file-utils';
import { isTodoProcurementTaskList, isHrSourceEFormList } from '../hr-task-lists.config';
import {
  isCommentEntryTitle, isStatuslessCommentTitle, extractAttachmentFromCommentHtml,
  extractEFormUrlFromComment,
} from '../utils/comment-utils';
import { isDocumentLibraryTaskList } from '../utils/hr-task-matching';
import {
  cleanParsedPersonName, extractPersonName, collectTaskAssigneeValues,
} from '../utils/person-name-matching';
import { UserService } from './user.service';
import { SubordinaryTaskService } from './subordinaryTask.service';

@Injectable({ providedIn: 'root' })
export class HrTaskMapperService {
  constructor(
    private readonly userService: UserService,
    private readonly subordinaryTaskService: SubordinaryTaskService,
  ) { }

  /** Pull the assigner/submitter name from workflow comment HTML or plain text. */
  private parseAssignerFromComment(comment: unknown): string {
    if (!comment) return '';
    const plain = String(comment)
      .replace(/<[^>]*>/g, ' ')
      .replace(/&nbsp;|&#160;/g, ' ')
      .replace(/&#58;/g, ':').replace(/&#46;/g, '.')
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
      .replace(/\s+/g, ' ')
      .trim();

    const stopBefore = '(?:\\s+please click|\\.|$)';
    const patterns = [
      new RegExp(`was assigned to you by\\s+(.+?)${stopBefore}`, 'i'),
      new RegExp(`assigned to you by\\s+(.+?)${stopBefore}`, 'i'),
      new RegExp(`was assigned.*?by\\s+(.+?)${stopBefore}`, 'i'),
      new RegExp(`requested by:\\s*(.+?)${stopBefore}`, 'i'),
    ];
    for (const pattern of patterns) {
      const match = plain.match(pattern);
      if (match?.[1]) return cleanParsedPersonName(match[1]);
    }
    return '';
  }


  /** Match assignee person/group fields even when Graph omits LookupValue (email-only entries). */
  isAssignedToUserFromFields(
    fields: Record<string, any>,
    name: string,
    email: string,
  ): boolean {
    const assigneeValues = collectTaskAssigneeValues(fields);
    if (assigneeValues.some(value => this.userService.matchesAssigneeField(value, name, email))) {
      return true;
    }

    const emailLower = email.toLowerCase();
    const personFields = [fields['AssignedTo'], fields['Assigned'], fields['AssignedTo0']];
    for (const field of personFields) {
      if (!field) continue;
      const entries = Array.isArray(field) ? field : [field];
      for (const entry of entries) {
        if (!entry || typeof entry !== 'object') continue;
        const person = entry as Record<string, unknown>;
        const personEmail = String(
          person['EMail'] ?? person['Email'] ?? person['email'] ??
          person['UserPrincipalName'] ?? person['userPrincipalName'] ?? '',
        ).toLowerCase().trim();
        const personName = extractPersonName(entry);
        if (personEmail && emailLower && personEmail === emailLower) return true;
        if (personName && this.userService.matchesAssigneeField(personName, name, email)) return true;
      }
    }

    return false;
  }

  resolveTaskSubmitter(item: any, fields: any, listName: string): string {
    // ProcTasks / ProcTasksArchive (and similar): Live uses CustomCreatedBy, not Author/Created By.
    const customCreatedBy = extractPersonName(fields.CustomCreatedBy);
    if (customCreatedBy) return customCreatedBy;

    const sharePointTitle = String(fields.Title ?? '').trim();
    if (isCommentEntryTitle(sharePointTitle)) {
      return this.resolveCommentEntrySubmitter(item, fields);
    }

    if (isDocumentLibraryTaskList(listName) || isTodoProcurementTaskList(listName)) {
      const commentText = fields.Comment ?? fields.Notes ?? fields.field_10 ?? '';
      return (
        extractPersonName(fields.Requestor) ||
        extractPersonName(fields.SubmittedBy) ||
        this.parseAssignerFromComment(commentText) ||
        extractPersonName(fields.CreatedBy) ||
        extractPersonName(fields.Author) ||
        extractPersonName(item?.createdBy?.user) ||
        extractPersonName(item?.createdBy) ||
        ''
      );
    }

    // Sick Certificate Upload and other source eForm lists often store the person
    // on Employee / EmployeeName rather than Author.
    if (isHrSourceEFormList(listName)) {
      return (
        extractPersonName(fields.EmployeeName) ||
        extractPersonName(fields.Employee) ||
        extractPersonName(fields.Requestor) ||
        extractPersonName(fields.SubmittedBy) ||
        extractPersonName(fields.Submitter) ||
        extractPersonName(fields.Author) ||
        extractPersonName(fields.CreatedBy) ||
        extractPersonName(item?.createdBy?.user) ||
        extractPersonName(item?.createdBy) ||
        sharePointTitle
      );
    }

    return (
      extractPersonName(fields.Requestor) ||
      extractPersonName(fields.SubmittedBy) ||
      extractPersonName(fields.Submitter) ||
      extractPersonName(fields.Author) ||
      extractPersonName(fields.CreatedBy) ||
      extractPersonName(item?.createdBy?.user) ||
      extractPersonName(item?.createdBy) ||
      String(fields.Title ?? '').trim()
    );
  }

  private resolveCommentEntrySubmitter(item: any, fields: Record<string, unknown>): string {
    const candidates = [
      fields['commentSubmittedBy'],
      fields['SubmittedBy'],
      fields['Submitter'],
      fields['Requestor'],
      fields['Author'],
      item?.createdBy?.user?.displayName,
      item?.createdBy?.user?.email,
      item?.createdBy?.user?.title,
      item?.lastModifiedBy?.user?.displayName,
    ];

    for (const candidate of candidates) {
      const name = extractPersonName(candidate);
      if (name && !isCommentEntryTitle(name)) {
        return name;
      }
    }

    return this.userService.getCurrentUserName();
  }

  // ============================================================
  // SHAPE ONE HR TASK ITEM
  // Pure mapping - returns null when the item doesn't belong to
  // the current user.
  // ============================================================
  mapSharePointItemToHrTask(
    item: any,
    listName: string,
    list: any,
    siteWebUrl: string,
    userEmail: string,
    userUpn: string,
    skipUserFilter = false,
    subordinateTasksOnly = false
  ): any | null {
    const currentUser = this.userService.getCurrentUser();
    const f = item.fields ?? {};

    // ?? User matching ??????????????????????????????????????
    const submitter = this.resolveTaskSubmitter(item, f, listName);
    const submitterLower = submitter.toLowerCase();
    const userDisplayLower = (currentUser?.username ?? '').toLowerCase();
    const emailLocal = userEmail.includes('@') ? userEmail.split('@')[0] : userEmail;
    const displayParts = userDisplayLower.split(/\s+/).filter(p => p.length > 2);
    const emailParts = emailLocal.replace(/[-_.]/g, ' ').split(' ').filter(p => p.length > 2);
    const allNameParts = [...new Set([...displayParts, ...emailParts])];

    const extractPerson = (val: any): string => extractPersonName(val);

    const exactMatch = submitterLower === userDisplayLower;
    const partialMatch = allNameParts.length >= 2 &&
      allNameParts.filter(p => submitterLower.includes(p)).length >= 2;
    const isCurrentUser = exactMatch || partialMatch ||
      submitterLower.includes(userEmail) || submitterLower.includes(userUpn);

    const assigneeValues = collectTaskAssigneeValues(f);
    const assignedToValue = assigneeValues.join(', ') || extractPersonName(f.AssignedTo ?? f.Assigned);
    const matchName = currentUser?.username ?? currentUser?.userPrincipalName ?? '';
    const matchEmail = userEmail || userUpn;
    const isAssignedToCurrentUser =
      skipUserFilter || this.isAssignedToUserFromFields(f, matchName, matchEmail);

    // Keep tasks where the current user is Superior 1 or Superior 2, so the
    // "Subordinate Tasks" view in the To Do list has data to show.
    const superior1Value = extractPerson(f.AssignedToSuperior1);
    const superior2Value = extractPerson(f.AssignedToSuperior2);
    const isSuperiorForCurrentUser =
      this.userService.matchesAssigneeField(superior1Value, matchName, matchEmail) ||
      this.userService.matchesAssigneeField(superior2Value, matchName, matchEmail);
    const isAssignedToSubordinate = this.subordinaryTaskService.isTaskAssignedToAnyAssigneeValue(
      assigneeValues.length ? assigneeValues : [assignedToValue]
    );

    const sharePointTitle = String(f.Title ?? '').trim();
    const sharePointCategory = String(f.Category ?? f.Type ?? f.field_9 ?? '').trim();
    const isCommentEntry = isCommentEntryTitle(sharePointTitle) || isCommentEntryTitle(sharePointCategory);
    const isStatuslessComment =
      isStatuslessCommentTitle(sharePointTitle) ||
      isStatuslessCommentTitle(sharePointCategory);

    if (subordinateTasksOnly) {
      if (!isAssignedToSubordinate) return null;
    } else if (
      !skipUserFilter &&
      !isCommentEntry &&
      !isCurrentUser &&
      !isAssignedToCurrentUser &&
      !isSuperiorForCurrentUser &&
      !isAssignedToSubordinate &&
      submitterLower !== ''
    ) {
      return null;
    }

    // ?? Type flags ?????????????????????????????????????????
    const isChangeOfShift = listName.toLowerCase().includes('changeofshift');
    const isTeleworkReports = listName.toLowerCase().includes('teleworkreports');
    const isTelework = listName.toLowerCase().includes('telework') && !isTeleworkReports;

    // ?? Status ?????????????????????????????????????????????
    // Procurement lists: Progress is the live gate (Pending/Complete). LastState/Status
    // often still say "Approved..." from an earlier step - using those hid open Pr- tasks.
    let status = isTodoProcurementTaskList(listName)
      ? String(f.Progress ?? f.LastState ?? f.Status ?? 'Pending').trim()
      : String(
          f.LastState ?? f.Status ?? f.ApprovalStatus ?? f.WorkflowStatus ??
          f.TaskOutcome ?? f.Outcome ?? f.CurrentStage ?? f.Stage ??
          f.field_3 ?? f.field_4 ?? f.field_5 ?? f.Completed ?? f.IsCompleted ??
          f.State ?? f.Progress ?? f.Modified ?? 'Pending'
        ).trim();

    // Hide status on General Comment, New Attachment, and Request for Action cards
    if (isStatuslessComment) {
      status = '';
    }

    // ?? Approver chain ?????????????????????????????????????
    const approver1 = String(f.Approver1 ?? f.Approver ?? f.field_12 ?? '').trim();
    const approver2 = String(f.Approver2 ?? f.field_15 ?? '').trim();
    const approver3 = String(f.Approver3 ?? f.field_28 ?? '').trim();
    const sharePointCompletedBy =
      extractPersonName(f.CompletedBy) ||
      extractPersonName(f.CompletedBy0) ||
      '';
    const primaryApprover = sharePointCompletedBy || approver1 || approver2 || approver3;

    const approver1Date = String(f.Approver1Date ?? f.field_13 ?? '').trim();
    const approver1Comment = String(f.Approver1Comment ?? f.field_29 ?? '').trim();
    const approver2Date = String(f.Approver2Date ?? '').trim();
    const approver2Comment = String(f.Approver2Comment ?? '').trim();
    const approver3Date = String(f.Approver3Date ?? '').trim();
    const approver3Comment = String(f.Approver3Comment ?? '').trim();

    // ?? Dates & time ???????????????????????????????????????
    const missedDate = String(f.MissedDate ?? f.field_5 ?? '').trim();
    // Proc archive rows: Created/Modified are often archive-move dates; Live uses custom columns.
    const customCreatedDate = String(f.CustomCreatedDate ?? '').trim();
    const customModifiedBy =
      extractPersonName(f.CustomModifiedBy) ||
      extractPersonName(f.Editor) ||
      extractPersonName(f.ModifiedBy) ||
      extractPersonName(item?.lastModifiedBy?.user) ||
      extractPersonName(item?.lastModifiedBy);
    const customModifiedDate = String(
      f.CustomModifiedDate ?? f.CustomModified ?? ''
    ).trim();
    // SubmittedDate is the primary sort key for HR Files progressive loading.
    // Some lists omit Created/Submitted fields, so we fall back to last modified
    // to avoid "0" dates that cause items to appear at the bottom first and
    // then jump to the top later.
    const submittedDate = String(
      customCreatedDate ||
        f.Created ||
        f.SubmittedDate ||
        item?.createdDateTime ||
        item?.fields?.Created ||
        item?.fields?.SubmittedDate ||
        f.Modified ||
        item?.lastModifiedDateTime ||
        ''
    ).trim();
    const statusDate = String(f.Approver1Date ?? f.field_13 ?? f.Modified ?? '').trim();
    const timeIn = String(f.TimeIn ?? f.TimeInText ?? f.field_7 ?? '').trim();
    const timeOut = String(f.TimeOut ?? f.TimeOutText ?? f.field_8 ?? '').trim();
    const missed = String(f.Missed ?? f.field_6 ?? '').trim();

    // ?? Common fields ??????????????????????????????????????
    const reason = String(f.Reason ?? f.ReasonOther ?? f.field_11 ?? '').trim();
    const stage = String(f.CurrentStage ?? f.field_4 ?? '').trim();
    const section = String(f.Section ?? f.field_34 ?? '').trim();
    const fromDate = String(f.FromDate ?? f.StartDate ?? f.SickLeaveStartDate ?? f.DateFrom ?? f.field_12 ?? f.field_13 ?? '').trim();
    const toDate = String(f.ToDate ?? f.EndDate ?? f.SickLeaveEndDate ?? f.DateTo ?? f.field_14 ?? f.field_15 ?? '').trim();

    // ?? TransferOfVL fields ????????????????????????????????
    const hoursForTransfer = String(f.HoursForTransfer ?? f.HoursTransferred ?? '').trim();
    const fromYear = String(f.FromYear ?? '').trim();
    const toYear = String(f.ToYear ?? '').trim();
    const needEngineer = String(f.NeedEngineerApproval ?? f.NeedEngineerAp ?? '').trim();
    const requestorPin = String(f.RequestorPin ?? '').trim();
    const approver1Pin = String(f.Approver1Pin ?? '').trim();
    const approver2Pin = String(f.Approver2Pin ?? '').trim();
    const approver3Pin = String(f.Approver3Pin ?? '').trim();
    const spToken = String(f.Token ?? '').trim();

    // ?? Extra columns ??????????????????????????????????????
    const progress = String(f.Progress ?? f.field_5 ?? '').trim();
    const taskOutcome = String(f.TaskOutcome ?? f.Outcome ?? f.field_6 ?? '').trim();
    const taskOutcomeField = String(f.TaskOutcome ?? '').trim();
    const eFormCategory = String(f.Category ?? '').trim();
    const eFormProgress = String(f.Progress ?? '').trim();
    const createdBy = String(f.AuthorLookupId ?? '').trim();
    const modifiedBy = String(f.EditorLookupId ?? '').trim();
    const createdDate = String(f.Created ?? '').trim();
    const modifiedDate = String(f.Modified ?? '').trim();
    const category = String(f.Category ?? f.Type ?? f.field_9 ?? '').trim();
    // SharePoint field that todo-list uses as the grouping id (see tasksAssignedToMe / buildGroupKey).
    const eFormListId = String(f.eFormListId ?? f.ListId ?? f.field_11 ?? '').trim();

    const assignedTo = assignedToValue || extractPerson(f.AssignedTo ?? f.Assigned);
    const assignedToSuperior = extractPerson(f.AssignedToSuperior ?? f.Superior ?? f.field_8);
    const assignedTo2 = extractPerson(f.AssignedTo);
    const assignedToSuperior1 = extractPerson(f.AssignedToSuperior1);
    const assignedToSuperior2field = extractPerson(f.AssignedToSuperior2);

    // ?? Task routing / SLA columns (shown in AllFiles ? Comments) ??
    const assignedAtStepNo = String(
      f.AssignedAtStepNo ?? f.AssignedAtStep ?? f.AssignedAtStepNumber ?? ''
    ).trim();
    const assignedToStepNo = String(
      f.AssignedToStepNo ?? f.AssignedToStep ?? f.AssignedToStepNumber ?? ''
    ).trim();
    const dueDate = String(
      f.DueDate ?? f.DueDate0 ?? f.Due ?? f.Due_x0020_Date ?? ''
    ).trim();

    // ?? Comment field ??????????????????????????????????????
    let comment = '';
    let commentHtml = '';
    if (f.Comment ?? f.Notes ?? f.field_10) {
      let raw = String(f.Comment ?? f.Notes ?? f.field_10).trim()
        .replace(/&nbsp;|&#160;/g, ' ')
        .replace(/&#58;/g, ':').replace(/&#46;/g, '.')
        .replace(/&quot;|&#34;/g, '"').replace(/&#39;|&apos;/g, "'")
        .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
        .replace(/<(?!\/?a\b)[^>]*>/gi, '');
      comment = raw.replace(/<[^>]*>/g, '').trim()
        .replace(/Please Approve ([\w\s]+) eForm for - ([\w\s]+)/gi, '$1 Request\n\nRequested by: $2')
        .replace(/Date From[\s]*:[\s]*([\d\/]+)/gi, 'From: $1')
        .replace(/Date To[\s]*:[\s]*([\d\/]+)/gi, '\nTo: $1')
        .replace(/Requested Days[\s]*:[\s]*(\d+)/gi, '\nDuration: $1 days')
        .replace(/Please click.*HERE.*to update eForm[\.] ID=([\d]+)/gi, '\n\neForm List ID: $1')
        .replace(/\s+/g, ' ').replace(/\n\s*\n/g, '\n\n').trim();
      commentHtml = raw
        // Match quoted AND unquoted hrefs - SharePoint/PowerApps comments are
        // stored with mixed encodings, so quotes may be missing after decoding.
        .replace(/<a\b[^>]*\bhref\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))[^>]*>/gi, (_match, dq, sq, unq) => {
          const rawHref = String(dq ?? sq ?? unq ?? '').replace(/^["']+|["']+$/g, '').trim();
          const fixedHref = normalizeSharePointFileUrl(rawHref);
          return `<a class="comment-link" href="${fixedHref}" target="_blank" rel="noopener noreferrer">`;
        })
        .replace(/(Date From\s*:)/gi, '<br><br>$1')
        .replace(/(Date To\s*:)/gi, '<br>$1')
        .replace(/(Requested Days\s*:)/gi, '<br>$1')
        .replace(/(Please click)/gi, '<br>$1')
        .replace(/(<br\s*\/?>\s*){3,}/gi, '<br><br>').trim();
    }

    // insert line breaks before numbered items (e.g. "39. ...40. ...")
    // Only touch text nodes - running this over the whole string used to inject
    // "<br>" inside href="..." attributes and break links to files whose URL
    // contains digits followed by a dot (e.g. ".../2022.pdf").
    commentHtml = commentHtml
      .split(/(<[^>]*>)/)
      .map(part => part.startsWith('<') ? part : part.replace(/(\d+)\.\s*/g, '<br>$1. '))
      .join('');

    // ensure adjacent anchor tags are separated onto their own lines
    commentHtml = commentHtml.replace(/<\/a>\s*<a /gi, '</a><br><a ');

    const attachmentFromHtml = extractAttachmentFromCommentHtml(commentHtml);
    const attachmentUrl = normalizeSharePointFileUrl(attachmentFromHtml.url);
    const attachmentFileName = attachmentFromHtml.fileName;
    const updateEFormUrl = normalizeSharePointFileUrl(extractEFormUrlFromComment(commentHtml));
    // ?? Body text ??????????????????????????????????????????
    const taskLabel = listName.replace('HRTask', '');
    const displayName = isCommentEntry ? sharePointTitle : taskLabel;
    const statusLower = status.toLowerCase();
    const isApproved = statusLower.includes('approv') || statusLower.includes('complet');
    let body = '';

    if (isChangeOfShift) {
      body = submitter && primaryApprover && isApproved
        ? `Change of Shift Request for ${submitter} was Approved by ${primaryApprover}`
        : submitter ? `Change of Shift Request for ${submitter}` : '';
      if (requestorPin) body += `\nRequestor PIN: ${requestorPin}`;
      if (section) body += `\nSection: ${section}`;
      if (reason) body += `\nReason: ${reason}`;
      if (stage) body += `\nCurrent Stage: ${stage}`;
      if (approver1) {
        body += `\n${isApproved ? '(approved)' : '(pending)'} Approver 1: ${approver1}`;
        if (approver1Pin) body += ` (PIN: ${approver1Pin})`;
        if (approver1Date) body += ` - ${approver1Date}`;
        if (approver1Comment) body += `\n  Comment: ${approver1Comment}`;
      }
      if (approver2) {
        body += `\n(pending) Approver 2: ${approver2}`;
        if (approver2Pin) body += ` (PIN: ${approver2Pin})`;
        if (approver2Date) body += ` - ${approver2Date}`;
        if (approver2Comment) body += `\n  Comment: ${approver2Comment}`;
      }
      if (approver3) {
        body += `\n(pending) Approver 3: ${approver3}`;
        if (approver3Pin) body += ` (PIN: ${approver3Pin})`;
        if (approver3Date) body += ` - ${approver3Date}`;
        if (approver3Comment) body += `\n  Comment: ${approver3Comment}`;
      }
      if (spToken) body += `\nToken: ${spToken}`;
    } else if (isTelework || isTeleworkReports) {
      body = submitter && primaryApprover && isApproved
        ? `Telework Request for ${submitter} was Approved by ${primaryApprover}`
        : submitter ? `Telework Request for ${submitter}` : '';
      if (fromDate && toDate) body += `\nFrom: ${fromDate}  To: ${toDate}`;
      if (reason && reason !== 'Enter value here') body += `\nReason: ${reason}`;
      if (primaryApprover && primaryApprover !== 'Enter value here') {
        body += `\n${isApproved ? '(approved)' : '(pending)'} ${primaryApprover}`;
        if (approver1Date) body += ` - ${approver1Date}`;
      }
    } else {
      body = submitter && primaryApprover && isApproved
        ? `${taskLabel} Task for ${submitter} was Approved by ${primaryApprover}`
        : submitter ? `${taskLabel} Task for ${submitter}` : '';
      if (reason && reason !== 'Enter value here') body += `\nReason: ${reason}`;
      if (stage) body += `\nCurrent Stage: ${stage}`;
      if (section) body += `\nSection: ${section}`;
      if (approver1) {
        body += `\n${isApproved ? '(approved)' : '(pending)'} Approver 1: ${approver1}`;
        if (approver1Date) body += ` - ${approver1Date}`;
        if (approver1Comment) body += `\n  Comment: ${approver1Comment}`;
      }
      if (approver2) {
        body += `\n(pending) Approver 2: ${approver2}`;
        if (approver2Date) body += ` - ${approver2Date}`;
      }
      if (approver3) {
        body += `\n(pending) Approver 3: ${approver3}`;
        if (approver3Date) body += ` - ${approver3Date}`;
      }
    }
    body += `\nPlease click HERE to update HR Task. eFormListId=${eFormListId}`;

    const itemWebUrl = list.webUrl
      ? `${list.webUrl}/DispForm.aspx?ID=${item.id}`
      : `${siteWebUrl}/Lists/${listName}/DispForm.aspx?ID=${item.id}`;

    return {
      id: item.id,
      name: displayName,
      webUrl: itemWebUrl,
      lastModifiedDateTime: f.Modified ?? item?.lastModifiedDateTime ?? '',
      isFolder: false,
      status, statusDate, submittedBy: submitter, submittedDate,
      completedBy: primaryApprover,
      description: body,
      listName,
      isAssignedToCurrentUser,
      isContentLoaded: true,
      eFormDetails: {
        type: displayName, status, submitter, listName,
        completedBy: primaryApprover,
        approver1, approver1Date, approver1Comment,
        approver2, approver2Date, approver2Comment,
        approver3, approver3Date, approver3Comment,
        submittedDate, statusDate, missedDate, timeIn, timeOut, missed,
        hoursForTransfer, fromYear, toYear, needEngineer,
        requestorPin, approver1Pin, approver2Pin, approver3Pin, token: spToken,
        fromDate, toDate, reason, stage, section, body,
        progress, taskOutcome, assignedTo, assignedToSuperior,
        category, comment, commentHtml: isCommentEntry ? (commentHtml || comment) : commentHtml, eFormListId,
        attachmentUrl, attachmentFileName,
        assignedAtStepNo, assignedToStepNo, dueDate,
        taskOutcomeField, assignedTo2, assignedToSuperior1, assignedToSuperior2field,
        createdBy, modifiedBy, createdDate, modifiedDate, eFormCategory, eFormProgress,
        customCreatedBy: submitter,
        customCreatedDate: customCreatedDate || submittedDate,
        customModifiedBy,
        customModifiedDate: customModifiedDate || String(f.Modified ?? item.lastModifiedDateTime ?? '').trim(),
        rawFields: f,
        updateEFormUrl,
      },
    };
  }
}
