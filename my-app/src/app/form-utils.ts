/** Turns PascalCase/camelCase (e.g. "SickLeaveByAppointment") into spaced words. */
export function formatFormTitle(title: string | null | undefined): string {
    return (title ?? '')
      .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
      .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2')
      .trim();
  }