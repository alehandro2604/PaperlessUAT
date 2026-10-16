# Angular: Study + Code (short lessons)

Each lesson = **Study** (1 minute) → **See it in your app** → **Your task** (5–15 minutes) → **Done when**.
Answers are hidden under "Solution". Try first, **no AI, no copy-paste**. Stuck 10 minutes? Open the solution, read it, close it, type it again from memory.

**Setup (once):** work in a practice project, not the real app.
```bash
ng new practice --style=css --ssr=false
cd practice
ng serve        # leave running, open http://localhost:4200
```
Put every task in `practice/src/app/app.ts` and `app.html` (replace what is there) unless it says otherwise.

---

## Lesson 1: Show a value
**Study:** `{{ }}` in HTML prints a variable from the class. A `signal` is a box; read it with `()`.

**See it in your app:** `components/loading screen/loading-screen.html` prints `{{ message }}` style values, and `loading-screen.ts` holds `message = 'Loading...'`.

**Task:** Show `Hello, <your name>!` in an `<h1>`, with the name stored in a signal.

**Done when:** the page shows your name. Change the name in the `.ts` file and watch it update.

<details><summary>Solution</summary>

```ts
import { Component, signal } from '@angular/core';
@Component({ selector: 'app-root', templateUrl: './app.html' })
export class App {
  name = signal('Camilleri');
}
```
```html
<h1>Hello, {{ name() }}!</h1>
```
</details>

---

## Lesson 2: Click a button
**Study:** `(click)="method()"` runs a method when clicked. `signal.update(old => new)` changes the box.

**See it in your app:** `components/manual-refresh/refresh.html`: `(click)="refreshTasks()"` calls a method in `refresh.ts`.

**Task:** A counter: a number, a `+` button, a `-` button, and a `Reset` button. The number must never go below 0.

**Done when:** `-` at 0 does nothing.

<details><summary>Solution</summary>

```ts
count = signal(0);
add()    { this.count.update(c => c + 1); }
remove() { this.count.update(c => Math.max(0, c - 1)); }
reset()  { this.count.set(0); }
```
```html
<p>{{ count() }}</p>
<button (click)="remove()">-</button>
<button (click)="add()">+</button>
<button (click)="reset()">Reset</button>
```
</details>

---

## Lesson 3: Show or hide with `@if`
**Study:** `@if (condition) { ... } @else { ... }` shows parts of the page only when something is true.

**See it in your app:** `todo-list.component.html` line 57: `*ngIf="!isTodoLoading && submitterNames.length === 0"` shows "You're all caught up!". That's the **old** way to write `@if`; the idea is identical.

**Task:** Add a `isLoggedIn` signal and a button that toggles it. Show `Welcome back` when true and `Please sign in` when false.

**Done when:** the button flips the message. (Hint: `this.isLoggedIn.update(v => !v)`.)

<details><summary>Solution</summary>

```ts
isLoggedIn = signal(false);
toggle() { this.isLoggedIn.update(v => !v); }
```
```html
@if (isLoggedIn()) { <p>Welcome back</p> } @else { <p>Please sign in</p> }
<button (click)="toggle()">Toggle</button>
```
</details>

---

## Lesson 4: Lists with `@for`
**Study:** `@for (item of list; track item.id) { ... }` repeats HTML for each item. `track` tells Angular how to tell items apart.

**See it in your app:** `todo-list.component.html` line 84: `*ngFor="let group of filteredGroupedBySubmitter"` draws one block per person. `app-dropdown.component.html` draws one button per option.

**Task:** Make this list and show each title in an `<li>`:
```ts
tasks = signal([
  { id: 1, title: 'Approve leave', done: false },
  { id: 2, title: 'Sign invoice', done: true },
  { id: 3, title: 'Review contract', done: false },
]);
```
**Done when:** three lines appear. Then add `[class.done]="t.done"` and a CSS rule `.done { text-decoration: line-through; }` so finished tasks are crossed out.

<details><summary>Solution</summary>

```html
<ul>
  @for (t of tasks(); track t.id) {
    <li [class.done]="t.done">{{ t.title }}</li>
  }
</ul>
```
</details>

---

## Lesson 5: TypeScript shapes (`interface`)
**Study:** An `interface` describes what an object must look like. TypeScript then catches typos for you.

**See it in your app:** `services/site-metadata.service.ts` lines 7–11 define `SiteMetadata`. `app-dropdown.component.ts` defines `AppDropdownOption` with an optional `tone?`.

**Task:** Create `interface Task { id: number; title: string; done: boolean; dueDate?: string }`. Use it as the type of the signal in Lesson 4: `signal<Task[]>([...])`. Then deliberately write `titel:` in one item and read the red error.

**Done when:** you can explain why `dueDate?` doesn't have to be filled in.

<details><summary>Solution</summary>

```ts
interface Task { id: number; title: string; done: boolean; dueDate?: string }
tasks = signal<Task[]>([ { id: 1, title: 'Approve leave', done: false } ]);
```
`?` means optional: the property may be missing.
</details>

---

## Lesson 6: Mark done (change one item in a list)
**Study:** Don't change the old array. Make a **new** one with `.map()`; signals notice the change that way.

**Task:** Add a `Done` button inside each `<li>`. Clicking it flips `done` for that task only.

**Done when:** clicking Done on task 2 crosses out only task 2.

<details><summary>Solution</summary>

```ts
toggle(id: number) {
  this.tasks.update(list => list.map(t => t.id === id ? { ...t, done: !t.done } : t));
}
```
```html
<li [class.done]="t.done">{{ t.title }} <button (click)="toggle(t.id)">Done</button></li>
```
`{ ...t, done: !t.done }` = copy of `t`, but with `done` flipped.
</details>

---

## Lesson 7: `computed` (values worked out from other values)
**Study:** `computed(() => ...)` recalculates itself when the signals inside it change.

**See it in your app:** `todo-list.component.ts` has filtered/grouped lists worked out from the raw tasks (`filteredGroupedBySubmitter`), the same idea, written the older way.

**Task:** Show `2 of 3 tasks left` using a `computed` for the number of undone tasks.

**Done when:** the text changes when you press Done.

<details><summary>Solution</summary>

```ts
remaining = computed(() => this.tasks().filter(t => !t.done).length);
```
```html
<p>{{ remaining() }} of {{ tasks().length }} tasks left</p>
```
</details>

---

## Lesson 8: Typing in a text box
**Study:** `#box` gives an element a name inside the template. `box.value` reads what the user typed.

**Task:** Add an input and an `Add` button. Clicking Add creates a new task with that title (new id = current list length + 1) and clears the box. Ignore empty text.

**Done when:** you can add 3 tasks and they appear in the list.

<details><summary>Solution</summary>

```ts
add(title: string) {
  const clean = title.trim();
  if (!clean) return;
  this.tasks.update(list => [...list, { id: list.length + 1, title: clean, done: false }]);
}
```
```html
<input #box placeholder="New task">
<button (click)="add(box.value); box.value = ''">Add</button>
```
</details>

---

## Lesson 9: Child component with `input()` and `output()`
**Study:** A child gets data with `input()` and tells the parent about events with `output()`.

**See it in your app:**
- `loading-screen.ts`: `@Input() message = 'Loading...'` (data **in**).
- `manual-refresh/refresh.ts`: `@Output() refreshRequested = new EventEmitter<void>()` (event **out**).
- Those are the old spellings of `input()` and `output()`.

**Task:** Run `ng generate component task-item`. Move one `<li>` into it. It receives a `task` and emits `toggled` with the task id when the button is clicked. The parent uses `<app-task-item [task]="t" (toggled)="toggle($event)" />`. Remember to put `TaskItem` in the parent's `imports: []`.

**Done when:** it behaves exactly like Lesson 6, but through the child.

<details><summary>Solution</summary>

```ts
// task-item.ts
import { Component, input, output } from '@angular/core';
@Component({ selector: 'app-task-item', templateUrl: './task-item.html' })
export class TaskItem {
  task = input.required<Task>();
  toggled = output<number>();
}
```
```html
<!-- task-item.html -->
<span [class.done]="task().done">{{ task().title }}</span>
<button (click)="toggled.emit(task().id)">Done</button>
```
```ts
// parent: imports: [TaskItem]
```
Note `task()` has brackets: an `input()` is a signal too.
</details>

---

## Lesson 10: Service (shared data and logic)
**Study:** A service is a class that holds data/logic for many components. Get it with `inject()`.

**See it in your app:** `services/site-metadata.service.ts`: `@Injectable({ providedIn: 'root' })` marks it as shared. It keeps the answer in `cached` so the app doesn't ask SharePoint over and over. `todo-task-store.service.ts` does the same for tasks.

**Task:** Run `ng generate service task-store`. Move the `tasks` signal plus `add()` and `toggle()` into it. In the app, use `store = inject(TaskStore)` and call `store.tasks()`, `store.add(...)`.

**Done when:** the app works the same, but the parent no longer owns the data.

<details><summary>Solution</summary>

```ts
import { Injectable, signal } from '@angular/core';
@Injectable({ providedIn: 'root' })
export class TaskStore {
  tasks = signal<Task[]>([]);
  add(title: string) { this.tasks.update(l => [...l, { id: l.length + 1, title, done: false }]); }
  toggle(id: number) { this.tasks.update(l => l.map(t => t.id === id ? { ...t, done: !t.done } : t)); }
}
```
```ts
// in the component
store = inject(TaskStore);
```
```html
@for (t of store.tasks(); track t.id) { ... }
```
</details>

---

## Lesson 11: Filter with a dropdown (like your app)
**Study:** A `<select>` can drive a signal; a `computed` then filters the list.

**See it in your app:** `components/app-dropdown/`: your custom dropdown (`options`, `value`, `selectOption()`) is used in the To-Do list to filter by status.

**Task:** Add `<select>` with `all`, `open`, `done`. Show only the matching tasks.

**Done when:** choosing `done` shows only done tasks.

<details><summary>Solution</summary>

```ts
filter = signal<'all' | 'open' | 'done'>('all');
visible = computed(() => {
  const f = this.filter();
  return this.store.tasks().filter(t => f === 'all' || (f === 'done') === t.done);
});
setFilter(v: string) { this.filter.set(v as 'all' | 'open' | 'done'); }
```
```html
<select #s (change)="setFilter(s.value)">
  <option value="all">All</option><option value="open">Open</option><option value="done">Done</option>
</select>
@for (t of visible(); track t.id) { ... }
```
`'all' | 'open' | 'done'` is a **union type**: only these three words are allowed.
</details>

---

## Lesson 12: Load data from the internet
**Study:** `HttpClient.get()` returns an Observable (a future answer). `toSignal()` turns it into a signal.

**See it in your app:** `app.config.ts` has `provideHttpClient()`. That line must exist or `inject(HttpClient)` fails. `site-metadata.service.ts` shows a service that fetches and caches.

**Task:**
1. In `practice/src/app/app.config.ts` add `provideHttpClient()` to `providers`.
2. Fetch `https://jsonplaceholder.typicode.com/todos?_limit=10` and show the titles.
3. Show `Loading...` until the data arrives.

**Done when:** you see 10 titles.

<details><summary>Solution</summary>

```ts
import { inject } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { toSignal } from '@angular/core/rxjs-interop';

http = inject(HttpClient);
todos = toSignal(this.http.get<{ id: number; title: string }[]>(
  'https://jsonplaceholder.typicode.com/todos?_limit=10'));   // undefined until it arrives
```
```html
@if (todos(); as list) {
  @for (t of list; track t.id) { <p>{{ t.title }}</p> }
} @else {
  <p>Loading...</p>
}
```
</details>

---

## Final boss (30–60 minutes, no solution)
Build a mini "To-Do" page that:
1. Loads the 10 todos from Lesson 12 into the `TaskStore` (Lesson 10).
2. Lets you add tasks and mark them done (Lessons 6, 8).
3. Filters by status (Lesson 11).
4. Shows "X tasks left" (Lesson 7).
5. Uses a `TaskItem` child component (Lesson 9).

Then open `my-app/src/app/components/todo-list/todo-list.component.ts` in the real app and see how many of these pieces you now recognise.

## Quick fixes when something breaks
| Symptom | Likely cause |
|---------|--------------|
| `NG8001: 'app-task-item' is not a known element` | Add `TaskItem` to the parent's `imports: []` |
| Page shows nothing / `[object Object]` / function text | You wrote `count` instead of `count()` |
| `NullInjectorError: No provider for HttpClient` | Add `provideHttpClient()` in `app.config.ts` |
| List doesn't update after changing it | You changed the old array; use `update(l => [...l, x])` or `.map()` |
| `Property 'x' does not exist on type` | Typo, or the `interface` is missing that field |
