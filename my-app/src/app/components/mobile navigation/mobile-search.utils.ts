export type MobileSearchSection = 'hr' | 'todo' | 'allFiles';

export interface MobileSearchSyncInput {
  showToDoSection: boolean;
  showAllFilesSection: boolean;
  searchQuery: string;
  todoSearchTerm?: string;
  allFilesSearchTerm?: string;
}

export interface MobileSearchChangeResult {
  searchTerm: string;
  section: MobileSearchSection;
  searchQuery?: string;
  navigateToComments: boolean;
  commentsNavIndex: number;
}

export const MOBILE_COMMENTS_NAV_INDEX = 3;

export function getMobileSearchSection(
  showToDoSection: boolean,
  showAllFilesSection: boolean
): MobileSearchSection {
  if (showAllFilesSection) return 'allFiles';
  if (showToDoSection) return 'todo';
  return 'hr';
}

export function getMobileSearchPlaceholder(
  showToDoSection: boolean,
  showAllFilesSection: boolean
): string {
  const section = getMobileSearchSection(showToDoSection, showAllFilesSection);
  if (section === 'todo') return 'Search to do...';
  if (section === 'allFiles') return 'Search files...';
  return 'Search HR tasks...';
}

export function syncMobileSearchTerm(input: MobileSearchSyncInput): string {
  if (input.showAllFilesSection) {
    return input.allFilesSearchTerm ?? '';
  }
  if (input.showToDoSection) {
    return input.todoSearchTerm ?? '';
  }
  return input.searchQuery;
}

export function resolveMobileSearchChange(
  term: string,
  showToDoSection: boolean,
  showAllFilesSection: boolean
): MobileSearchChangeResult {
  const searchTerm = term ?? '';
  const section = getMobileSearchSection(showToDoSection, showAllFilesSection);

  if (section === 'allFiles' || section === 'todo') {
    return {
      searchTerm,
      section,
      navigateToComments: false,
      commentsNavIndex: MOBILE_COMMENTS_NAV_INDEX,
    };
  }

  return {
    searchTerm,
    section: 'hr',
    searchQuery: searchTerm,
    navigateToComments: searchTerm.trim().length > 0,
    commentsNavIndex: MOBILE_COMMENTS_NAV_INDEX,
  };
}
