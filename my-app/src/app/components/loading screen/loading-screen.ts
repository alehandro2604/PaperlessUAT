import { CommonModule } from '@angular/common';
import { Component, Input, ViewEncapsulation } from '@angular/core';

@Component({
  selector: 'app-loading-screen',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './loading-screen.html',
  styleUrls: ['./loading-screen.css'],
  encapsulation: ViewEncapsulation.None,
})
export class LoadingScreenComponent {
  private static nextId = 0;

  /** Unique prefix so two loaders on the page don't clash on SVG mask/filter ids. */
  readonly uid = `loader-${LoadingScreenComponent.nextId++}`;

  @Input() loading = false;
  @Input() message = 'Loading...';
  /** Smaller centered loader for the Comments panel (avoids competing with tall cards). */
  @Input() compact = false;
}
