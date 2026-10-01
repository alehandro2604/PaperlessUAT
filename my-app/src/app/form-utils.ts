/**
 * Turns a SharePoint form/list name into a readable title:
 * "SickCertificateUploader_Tasks" -> "Sick Certificate Uploader", "SickLeaveByAppointment" -> "Sick Leave By Appointment".
 */
export function formatFormTitle(title: string | null | undefined): string {
    return (title ?? '')
      .replace(/_Tasks?$/i, '') // drop the "_Tasks" list suffix
      .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
      .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2')
      .replace(/([a-zA-Z])(\d)/g, '$1 $2') // space between words and numbers
      .trim();
  }
