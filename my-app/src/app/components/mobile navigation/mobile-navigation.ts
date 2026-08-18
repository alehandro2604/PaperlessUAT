import {AfterViewInit,Component,EventEmitter, Input, OnChanges, Output, SimpleChanges, ViewEncapsulation } from '@angular/core';

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

export const mobileNavIconPaths = [
  'account.png',
  'todo.png',
  'allfiles.png',
  'speech-bubble.png',
  'attach-file.png',
  'file.png',
];

export function moveNavBubbleToItem(index: number): void {
  setTimeout(() => {
    const navBar = document.getElementById('navBar');
    const items = navBar?.querySelectorAll('.nav-item') ?? [];
    const bubble = document.getElementById('bubble');
    const glow = document.getElementById('glow');

    if (navBar && items[index] && bubble && glow) {
      const barRect = navBar.getBoundingClientRect();
      const itemRect = items[index].getBoundingClientRect();
      const cx = itemRect.left - barRect.left + itemRect.width / 2;

      bubble.style.left = cx - 25 + 'px';
      glow.style.left = cx - 33 + 'px';
    }
  });
}

export function setNavBubbleIcon(index: number): void {
  const bubbleIcon = document.getElementById('bubbleSvg');
  if (bubbleIcon) {
    bubbleIcon.innerHTML = `<img src="${mobileNavIconPaths[index]}" style="width:24px;height:24px;display:block;filter:brightness(0) invert(1);" />`;
  }
}

export function setupMobileBottomNav(): void {
  setTimeout(() => {
    moveNavBubbleToItem(0);
    setNavBubbleIcon(0);
  });
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

export function getCommentsTabMobileState(): {
  mobileCardView: MobileCardView;
  activeNavIndex: number;
} {
  return { mobileCardView: 'comments', activeNavIndex: MOBILE_NAV_INDEX.comments };
}

export function getAttachmentsTabMobileState(): {
  mobileCardView: MobileCardView;
  activeNavIndex: number;
} {
  return { mobileCardView: 'attachments', activeNavIndex: MOBILE_NAV_INDEX.attachments };
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
export class MobileNavigationComponent implements AfterViewInit, OnChanges {
  @Input() activeIndex = 0;
  @Output() navItemClick = new EventEmitter<number>();

  ngAfterViewInit(): void {
    this.updateBubble(this.activeIndex);
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['activeIndex'] && !changes['activeIndex'].firstChange) {
      this.updateBubble(this.activeIndex);
    }
  }

  onTap(index: number): void {
    this.navItemClick.emit(index);
  }

  private updateBubble(index: number): void {
    moveNavBubbleToItem(index);
    setNavBubbleIcon(index);
  }
}
