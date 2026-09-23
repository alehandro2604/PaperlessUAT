import { Component, EventEmitter, Input, Output, ViewEncapsulation } from '@angular/core';

export type MobileCardView = 'task' | 'todo' | 'allFiles' | 'comments' | 'attachments' | 'files';

export function getNextMobileView(current: MobileCardView): MobileCardView {
  if (current === 'task') return 'comments';
  if (current === 'todo') return 'allFiles';
  if (current === 'allFiles') return 'comments';
  if (current === 'comments') return 'attachments';
  if (current === 'attachments') return 'files';
  return 'task';
}

export function getMobileViewButtonLabel(view: MobileCardView): string {
  if (view === 'task') return 'Tasks';
  if (view === 'comments') return 'Comments';
  if (view === 'allFiles') return 'All Files';
  if (view === 'todo') return 'To Do';
  if (view === 'files') return 'Files';
  return 'Attachments';
}

export const MOBILE_NAV_INDEX = {
  task: 0,
  todo: 1,
  allFiles: 2,
  comments: 3,
  attachments: 4,
} as const;

export function getTaskTabMobileState(): {
  mobileCardView: MobileCardView;
  activeNavIndex: number;
} {
  return { mobileCardView: 'task', activeNavIndex: MOBILE_NAV_INDEX.task };
}

export function getTodoTabMobileState(): {
  mobileCardView: MobileCardView;
  activeNavIndex: number;
} {
  return { mobileCardView: 'todo', activeNavIndex: MOBILE_NAV_INDEX.todo };
}

export function getAllFilesTabMobileState(): {
  mobileCardView: MobileCardView;
  activeNavIndex: number;
} {
  return { mobileCardView: 'allFiles', activeNavIndex: MOBILE_NAV_INDEX.allFiles };
}

/** Detail screens keep the section the user came from highlighted in the bar. */
export function getCommentsTabMobileState(): {
  mobileCardView: MobileCardView;
} {
  return { mobileCardView: 'comments' };
}

export function getAttachmentsTabMobileState(): {
  mobileCardView: MobileCardView;
} {
  return { mobileCardView: 'attachments' };
}

export type MobileNavClickAction =
  | { type: 'none' }
  | { type: 'goHome' }
  | { type: 'goTodo' }
  | { type: 'goAllFiles' }
  | { type: 'goComments'; syncSearch: boolean }
  | { type: 'goAttachments'; shouldLoadFiles: boolean };

export function resolveMobileNavClick(
  index: number,
  activeIndex: number,
  hasCurrentUser: boolean
): MobileNavClickAction {
  if (index === activeIndex) {
    return { type: 'none' };
  }

  if (index === MOBILE_NAV_INDEX.task) {
    return { type: 'goHome' };
  }

  if (index === MOBILE_NAV_INDEX.todo) {
    return { type: 'goTodo' };
  }

  if (index === MOBILE_NAV_INDEX.allFiles) {
    return { type: 'goAllFiles' };
  }

  if (index === MOBILE_NAV_INDEX.comments) {
    return { type: 'goComments', syncSearch: true };
  }

  if (index === MOBILE_NAV_INDEX.attachments) {
    return {
      type: 'goAttachments',
      shouldLoadFiles: hasCurrentUser,
    };
  }

  return { type: 'none' };
}

export function handleMobileNavTap(
  index: number,
  activeIndex: number,
  mobileCardView: MobileCardView,
  currentUser?: { email?: string },
  openMyFiles?: () => void,
  openPowerAppsModal?: () => void
): {
  newActiveIndex: number;
  newMobileCardView?: MobileCardView;
  shouldLoadFiles: boolean;
  shouldOpenModal: boolean;
} {
  if (index === activeIndex) {
    return {
      newActiveIndex: activeIndex,
      shouldLoadFiles: false,
      shouldOpenModal: false,
    };
  }

  let newMobileCardView: MobileCardView | undefined;
  let shouldLoadFiles = false;
  const shouldOpenModal = false;

  switch (index) {
    case MOBILE_NAV_INDEX.task:
      newMobileCardView = 'task';
      break;
    case MOBILE_NAV_INDEX.todo:
      newMobileCardView = 'todo';
      break;
    case MOBILE_NAV_INDEX.allFiles:
      newMobileCardView = 'allFiles';
      break;
    case MOBILE_NAV_INDEX.comments:
      newMobileCardView = 'comments';
      break;
    case MOBILE_NAV_INDEX.attachments:
      newMobileCardView = 'attachments';
      shouldLoadFiles = !!currentUser?.email;
      break;
  }

  return {
    newActiveIndex: index,
    newMobileCardView,
    shouldLoadFiles,
    shouldOpenModal,
  };
}

@Component({
  selector: 'app-mobile-navigation',
  standalone: true,
  templateUrl: './mobile-navigation.html',
  styleUrls: ['./mobile-navigation-bar.css'],
  encapsulation: ViewEncapsulation.None,
})
export class MobileNavigationComponent {
  @Input() activeIndex = 0;
  @Output() navItemClick = new EventEmitter<number>();

  onTap(index: number): void {
    this.navItemClick.emit(index);
  }
}
