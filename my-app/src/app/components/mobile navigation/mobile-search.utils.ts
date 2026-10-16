interface MobileSearchSyncInput {
  showToDoSection: boolean;
  showAllFilesSection: boolean;
  searchQuery: string;
  todoSearchTerm?: string;
  allFilesSearchTerm?: string;
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
