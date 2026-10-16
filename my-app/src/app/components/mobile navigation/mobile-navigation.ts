import { Component, EventEmitter, Input, Output, ViewEncapsulation } from '@angular/core';

export type MobileCardView = 'task' | 'todo' | 'allFiles' | 'comments' | 'attachments' | 'files';

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
