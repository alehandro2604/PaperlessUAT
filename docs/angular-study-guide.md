# Angular + TypeScript Study Guide (for someone who already knows JS/TS)

Goal: sit down, open an empty file, and **write Angular code without AI**.

## Why you get stuck (and it's normal)

Next.js lets you start with "a function that returns UI". Angular makes you hold
~8 concepts in your head before the first useful line. You're not bad at it; you
just haven't memorized the **skeleton**, so every file starts with a blank page.

The fix is not understanding more. It is **typing the skeleton from memory until
your fingers know it**, then adding one concept at a time.

Your app (Angular 21, `my-app/`) is a good *reference* but a bad *textbook*:
`src/app/app.ts` is 7,700 lines and mixes everything, and it uses the older style
(`@Input()`, `*ngIf`, constructor injection). Learn the modern style below, and
use your app to practice **reading** real code.

---

## The 8 things that make up all of Angular

| # | Concept | Next.js/React equivalent | One-line mental model |
|---|---------|--------------------------|-----------------------|
| 1 | **Component** | Component function | A class + an HTML template + CSS |
| 2 | **Template syntax** | JSX | HTML with `{{ }}`, `[prop]`, `(event)`, `@if`, `@for` |
| 3 | **Signals** | `useState` | `count = signal(0)`; read `count()`, write `count.set(1)` |
| 4 | **Inputs / Outputs** | props / callback props | `input()` in, `output()` out |
| 5 | **Services + `inject()`** | Context / custom hooks / modules | A class that holds shared state or does HTTP |
| 6 | **Routing** | file-based routes | A `routes` array you write yourself |
| 7 | **Forms** | react-hook-form | Template-driven (`ngModel`) or Reactive (`FormGroup`) |
| 8 | **HttpClient + RxJS** | `fetch` / SWR | `http.get()` returns an *Observable*, not a Promise |

Learn them **in this order**. Don't skip ahead.

---

## The skeleton to memorize (type this 10 times)

Think of a component like a **LEGO brick**: it has a *brain* (the TypeScript file),
a *face* (the HTML file) and *clothes* (the CSS file). Angular snaps bricks together to make a page.

```ts
// hello.component.ts  <- the BRAIN of the brick: it holds data and the things the brick can DO

// "import" = borrow tools from Angular's toolbox so we can use them below.
//   Component -> the tool that says "this class is a brick"
//   signal    -> a tool for making a box that holds a value AND tells the page when it changes
//   computed  -> a tool for a value that is worked out from other values (like a formula in Excel)
import { Component, signal, computed } from '@angular/core';

// "@Component" is a label (called a decorator) stuck on top of the class.
// It tells Angular: "the class below is a component, and here are its settings".
@Component({
  // selector = the NAME of the brick as an HTML tag.
  // Anywhere you write <app-hello></app-hello>, this brick shows up.
  selector: 'app-hello',

  // imports = other bricks you want to use INSIDE this brick's face (HTML).
  // Empty [] means "I'm not using any other bricks". Forget to add one you use = error!
  imports: [],

  // templateUrl = which HTML file is the FACE of this brick.
  templateUrl: './hello.component.html',

  // styleUrl = which CSS file is the CLOTHES (colors, sizes) of this brick.
  styleUrl: './hello.component.css',
})
// "class" = a blueprint. "export" = lets other files use it. This is the brick's brain.
export class HelloComponent {
  // signal('World') makes a box that holds the text "World".
  // Later you can change what's inside, and the page updates by itself.
  name = signal('World');

  // Another box, this one holds the number 0 (it's how many times we clicked).
  count = signal(0);

  // computed = "always keep this equal to count x 2".
  // When count changes, doubled changes too. You never update it by hand.
  // Notice we READ a signal by adding () -> this.count()  (like opening the box to look inside).
  doubled = computed(() => this.count() * 2);

  // A method = an action the brick can do. The button below will call this.
  increment() {
    // update() = "take what's in the box (c), and put the new value in".
    // So: new count = old count + 1.
    this.count.update(c => c + 1);
  }
}
```

```html
<!-- hello.component.html  <- the FACE of the brick: what the user actually sees -->

<!-- {{ }} = "show a value here". name() opens the box and prints what's inside: Hello World -->
<h1>Hello {{ name() }}</h1>

<!-- (click)="increment()" = "WHEN this button is clicked, run the increment action".
     Round brackets ( ) always mean "listen for something happening".
     The text inside the button shows the count and the doubled number. -->
<button (click)="increment()">Clicked {{ count() }} (x2 = {{ doubled() }})</button>

<!-- @if = "only show this part IF the question is true", like an if statement in JavaScript.
     @else = "otherwise show this other part". -->
@if (count() > 5) {
  <p>That's a lot of clicks</p>
} @else {
  <p>Keep going</p>
}

<!-- @for = "repeat this part for every item in the list", like a loop.
     "n of [1, 2, 3]" = take each number one at a time and call it n.
     "track n" = give each item a name tag (its own value) so Angular can tell them apart.
     This makes three <li> lines: 1, 2, 3. -->
@for (n of [1, 2, 3]; track n) {
  <li>{{ n }}</li>
}
```

**The four bracket types, memorize these:**

- `{{ value }}`: show a value
- `[prop]="value"`: send data **into** an element/component
- `(event)="doSomething()"`: react to something happening
- `[(ngModel)]="value"`: both ways (forms)

---

## Study plan (6 weeks, ~1 hour/day)

Do the exercises **without AI and without copying**. Docs (angular.dev) are allowed.
If you're stuck for 15 minutes, look at the docs, close them, and type it again.

### Week 0: Setup (1 evening)
Create a **practice project outside this repo** so you can't break the real app:

```bash
# Install the Angular helper tool ("ng") on your computer. You only do this once.
npm install -g @angular/cli

# Make a brand-new Angular project in a folder called "practice".
# --style=css  -> use plain CSS for colors/layout
# --ssr=false  -> skip an advanced feature we don't need yet
ng new practice --style=css --ssr=false

# Go into the new folder (cd = "change directory").
cd practice

# Start the app. Open http://localhost:4200 in your browser.
# It reloads by itself every time you save a file.
ng serve
```

Open `src/app/app.ts` and `app.html`. Delete the boilerplate. This is your sandbox.
Useful commands: `ng generate component name`, `ng generate service name`.

### Week 1: TypeScript you actually need
You know TS basics; Angular leans on these specifically:
- `interface` / `type` for data shapes (see `SiteMetadata` in `src/app/services/site-metadata.service.ts`)
- Classes, `private`/`readonly`, constructor parameter properties
- Generics (`Array<T>`, `Promise<T>`, `Observable<T>`, `signal<T>()`)
- Decorators (`@Component`, `@Injectable`): just functions that attach metadata
- Optional chaining `?.`, nullish `??`, union types `string | null`
- `async/await`

**Exercise:** model a `Task { id: number; title: string; done: boolean; dueDate?: string }`
and write functions: `filterDone(tasks)`, `groupByDueDate(tasks)`, `toggle(task)` (immutably).

### Week 2: Components + templates + signals
- The skeleton above, `@if`, `@for`, `@switch`
- `signal`, `computed`, `.set()`, `.update()`
- Event binding, property binding, class binding `[class.active]="isOn()"`

**Exercises (in order):**
1. Counter with +, -, reset
2. Toggle that shows/hides a paragraph
3. A list of tasks rendered with `@for`; click to mark done (strike-through via `[class.done]`)
4. Add a text input + "Add" button to create tasks (hint: `#box` template ref, `box.value`)

### Week 3: Parent/child communication + services
- `input()` and `output()` (modern replacements for `@Input()` / `@Output()`)

A child brick can't see its parent's data. The parent has to **pass it down** (input),
and the child **shouts back up** when something happens (output). Like a kid (child) getting a
note from a teacher (parent), and putting a hand up when they have a question.

```ts
export class TaskItem {
  // input = a slot where the PARENT hands this brick some data.
  // <Task> = the data must have the shape of our Task type (TypeScript checks this).
  // ".required" = the parent MUST give it, or Angular shows an error.
  task = input.required<Task>();

  // output = a "raise hand" button. When we fire it, the parent hears about it.
  // <number> = along with the shout we send a number (the task's id).
  toggled = output<number>();
}
```
```html
<!-- This is in the PARENT's HTML. It uses the child brick like an HTML tag.

     [task]="t"            -> square brackets = SEND data IN. Give the child the task "t".
     (toggled)="onToggle($event)" -> round brackets = LISTEN. When the child shouts "toggled",
                                     run onToggle in the parent. $event is whatever the child
                                     sent with the shout (the task id). -->
<app-task-item [task]="t" (toggled)="onToggle($event)" />
```
- Services: `@Injectable({ providedIn: 'root' })` + `private store = inject(TaskStore)`

**Exercises:**
1. Split week 2's list into `TaskList` + `TaskItem` components
2. Move the tasks array into a `TaskStore` service with `add/toggle/remove`; both components read from it
3. Persist to `localStorage` inside the service

### Week 4: Routing + forms
- `app.routes.ts`: `{ path: 'tasks/:id', component: TaskDetail }`, `<router-outlet />`, `routerLink`
- Read params with `inject(ActivatedRoute)` or `input()` + `withComponentInputBinding()`
- Reactive forms: `FormGroup`, `FormControl`, `Validators.required`, show errors

**Exercises:** a 2-page app (list → detail), and an "Add task" form with validation.

### Week 5: HTTP + RxJS (the part that scares people)
Only learn these RxJS pieces; ignore the other 100 operators:
`Observable`, `subscribe`, `pipe`, `map`, `switchMap`, `catchError`, `debounceTime`,
and `toSignal()` to convert to a signal.

```ts
// inject(HttpClient) = "Angular, please hand me the tool that can talk to the internet".
// "private" = only this class can use it.
private http = inject(HttpClient);

// this.http.get<Task[]>(url) = ask the website at that address for a list of tasks.
//   It does NOT answer straight away (the internet is slow), so it gives back an
//   "Observable" = like a pizza order: the pizza arrives LATER.
// toSignal(...) = turns that "pizza order" into a signal box, so we can use it in our
//   HTML like any other signal: tasks().
// initialValue: [] = while we wait for the pizza, the box holds an empty list,
//   so the page doesn't crash.
tasks = toSignal(this.http.get<Task[]>('https://jsonplaceholder.typicode.com/todos'), { initialValue: [] });
```

**Exercises:** list todos from jsonplaceholder; add a search box that filters with
`debounceTime`; show loading + error states.

### Week 6: Build something small, alone
Rebuild a mini version of **your own app's To-Do list**: list, filter by status,
detail view, comment box, mock service. No AI. Then compare with the real code.

---

## Using YOUR app as reading practice

Read these in order. For each, answer: *what does the selector do, what are the inputs/outputs, where is state kept?*

| File | What to learn |
|------|---------------|
| `src/main.ts` + `src/app/app.config.ts` | How an Angular app boots; `provideHttpClient`, `provideRouter` |
| `components/loading screen/loading-screen.ts` (22 lines) | Smallest real component: `@Input()`s + template |
| `components/manual-refresh/refresh.ts` (36 lines) | `@Output()` + `EventEmitter` |
| `components/app-dropdown/` (130 lines) | A reusable UI component |
| `services/site-metadata.service.ts` (125 lines) | Service, `@Injectable`, caching a promise |
| `services/cache.ts` + `cache.spec.ts` | Plain TS logic + how tests are written |
| `components/todo-list/` (990 lines) | A full feature; read after week 4 |
| `app.ts` (7,700 lines) | Don't start here. Treat as a refactoring target |

Things in this codebase you'll see that differ from modern Angular (so you're not confused):
- `@Input()` / `@Output() EventEmitter` → now `input()` / `output()`
- `*ngIf` / `*ngFor` → now `@if` / `@for`
- `constructor(private http: HttpClient)` → now `private http = inject(HttpClient)`
- `standalone: true` is now the default and can be omitted
- No signals yet (0 files use `signal(`). New code could

---

## Habits that make the "blank page" problem disappear

1. **Type, don't paste.** Muscle memory is the whole game.
2. **Daily 10-minute drill:** open an empty folder, write a component + template + a service from memory. Check against the skeleton. Repeat until zero mistakes.
3. **Read the compiler error first.** Angular's errors are specific (`NG8002: Can't bind to 'x' since it isn't a known property`, usually means a missing `imports: []` entry).
4. **Know the 3 most common errors:**
   - `NG0303 / NG8002 / NG8001`: forgot to add a component/module to `imports: []`
   - `NullInjectorError: No provider for HttpClient`: missing `provideHttpClient()`
   - Template shows nothing: you wrote `count` instead of `count()` (signals are functions)
5. **One feature at a time.** Component first, hard-code data, *then* service, *then* HTTP.
6. **Use `ng generate`** for files, but write the contents yourself.
7. **Console + devtools:** install the *Angular DevTools* browser extension to inspect the component tree.

## Reference links
- Official tutorial (do this first): https://angular.dev/tutorials/learn-angular
- Docs: https://angular.dev
- Signals: https://angular.dev/guide/signals
- TypeScript handbook: https://www.typescriptlang.org/docs/handbook/intro.html
- RxJS (only the basics): https://rxjs.dev/guide/overview
