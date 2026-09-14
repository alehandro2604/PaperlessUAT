<#
  Component naming + folder cleanup for the Paperless Angular app.

  Fixes: spaces in folder names, inconsistent casing, and files missing
  Angular's `.component` suffix - then updates every import/templateUrl/
  styleUrls reference that pointed at the old names.

  SAFE TO RUN: everything happens on a new git branch
  (codebase-cleanup/component-naming). Nothing on main is touched.

  Don't like the result?
      git checkout main
      git branch -D codebase-cleanup/component-naming   # optional, deletes the branch

  Ran it and it errored partway through?
      git reset --hard HEAD   # discards the partial renames on this branch
      git checkout main

  After a successful run, review before merging:
      git diff main..codebase-cleanup/component-naming --stat
      cd my-app; ng build      # fastest way to confirm nothing broke

  Then, if you're happy with it:
      git checkout main
      git merge codebase-cleanup/component-naming
#>

$ErrorActionPreference = 'Stop'
$RepoRoot = $PSScriptRoot
$AppRoot  = Join-Path $RepoRoot 'my-app\src\app'
$Comp     = Join-Path $AppRoot 'components'

function Replace-InFile($Path, $Old, $New) {
    if (-not (Test-Path -LiteralPath $Path)) { throw "Expected file not found: $Path" }
    $content = Get-Content -Raw -LiteralPath $Path
    if (-not $content.Contains($Old)) {
        Write-Warning "Pattern not found in $Path`: $Old"
        return
    }
    $content = $content.Replace($Old, $New)
    Set-Content -NoNewline -LiteralPath $Path -Value $content
}

Set-Location $RepoRoot

# --- safety checks ----------------------------------------------------------
if (-not (Test-Path '.git')) { throw "No .git found at $RepoRoot - run this from the Paperless repo root." }

$dirty = git status --porcelain
if ($dirty) {
    Write-Host "You have uncommitted changes. Commit or stash them first so this cleanup" -ForegroundColor Yellow
    Write-Host "stays on its own clean branch and is easy to review/revert." -ForegroundColor Yellow
    $go = Read-Host "Continue anyway? (y/N)"
    if ($go -ne 'y') { exit 1 }
}

git checkout -b codebase-cleanup/component-naming

# --- 1. All-files -> all-files (case-only rename needs a temp hop on Windows) ---
git mv "$Comp\All-files" "$Comp\all-files-tmp"
git mv "$Comp\all-files-tmp" "$Comp\all-files"
git mv "$Comp\all-files\all-files.ts"   "$Comp\all-files\all-files.component.ts"
git mv "$Comp\all-files\all-files.html" "$Comp\all-files\all-files.component.html"
git mv "$Comp\all-files\all-files.css"  "$Comp\all-files\all-files.component.css"
Replace-InFile "$Comp\all-files\all-files.component.ts" "./all-files.html" "./all-files.component.html"
Replace-InFile "$Comp\all-files\all-files.component.ts" "./all-files.css"  "./all-files.component.css"

# --- 2. Claimable Task -> claimable-task ---
git mv "$Comp\Claimable Task" "$Comp\claimable-task"
git mv "$Comp\claimable-task\claim.ts"   "$Comp\claimable-task\claimable-task.component.ts"
git mv "$Comp\claimable-task\claim.html" "$Comp\claimable-task\claimable-task.component.html"
git mv "$Comp\claimable-task\claim.css"  "$Comp\claimable-task\claimable-task.component.css"
Replace-InFile "$Comp\claimable-task\claimable-task.component.ts" "./claim.html" "./claimable-task.component.html"
Replace-InFile "$Comp\claimable-task\claimable-task.component.ts" "./claim.css"  "./claimable-task.component.css"

# --- 3. delegate-component -> delegate ---
git mv "$Comp\delegate-component" "$Comp\delegate"
git mv "$Comp\delegate\delegate.ts"   "$Comp\delegate\delegate.component.ts"
git mv "$Comp\delegate\delegate.html" "$Comp\delegate\delegate.component.html"
git mv "$Comp\delegate\delegate.css"  "$Comp\delegate\delegate.component.css"
Replace-InFile "$Comp\delegate\delegate.component.ts" "./delegate.html" "./delegate.component.html"
Replace-InFile "$Comp\delegate\delegate.component.ts" "./delegate.css"  "./delegate.component.css"

# --- 4. loading screen -> loading-screen ---
git mv "$Comp\loading screen" "$Comp\loading-screen"
git mv "$Comp\loading-screen\loading-screen.ts"   "$Comp\loading-screen\loading-screen.component.ts"
git mv "$Comp\loading-screen\loading-screen.html" "$Comp\loading-screen\loading-screen.component.html"
git mv "$Comp\loading-screen\loading-screen.css"  "$Comp\loading-screen\loading-screen.component.css"
Replace-InFile "$Comp\loading-screen\loading-screen.component.ts" "./loading-screen.html" "./loading-screen.component.html"
Replace-InFile "$Comp\loading-screen\loading-screen.component.ts" "./loading-screen.css"  "./loading-screen.component.css"

# --- 5. manual-refresh/refresh.* -> manual-refresh.component.* (folder name already fine) ---
git mv "$Comp\manual-refresh\refresh.ts"   "$Comp\manual-refresh\manual-refresh.component.ts"
git mv "$Comp\manual-refresh\refresh.html" "$Comp\manual-refresh\manual-refresh.component.html"
git mv "$Comp\manual-refresh\refresh.css"  "$Comp\manual-refresh\manual-refresh.component.css"
Replace-InFile "$Comp\manual-refresh\manual-refresh.component.ts" "./refresh.html" "./manual-refresh.component.html"
Replace-InFile "$Comp\manual-refresh\manual-refresh.component.ts" "./refresh.css"  "./manual-refresh.component.css"

# --- 6. mobile navigation -> mobile-navigation ---
git mv "$Comp\mobile navigation" "$Comp\mobile-navigation"
git mv "$Comp\mobile-navigation\mobile-navigation.ts"   "$Comp\mobile-navigation\mobile-navigation.component.ts"
git mv "$Comp\mobile-navigation\mobile-navigation.html" "$Comp\mobile-navigation\mobile-navigation.component.html"
Replace-InFile "$Comp\mobile-navigation\mobile-navigation.component.ts" "./mobile-navigation.html" "./mobile-navigation.component.html"
# mobile-navigation-bar.css and mobile-navigation.css keep their names - only the folder moved.

# --- 7. new-comment/new-comment.* -> new-comment.component.* ---
git mv "$Comp\new-comment\new-comment.ts"   "$Comp\new-comment\new-comment.component.ts"
git mv "$Comp\new-comment\new-comment.html" "$Comp\new-comment\new-comment.component.html"
git mv "$Comp\new-comment\new-comment.css"  "$Comp\new-comment\new-comment.component.css"
Replace-InFile "$Comp\new-comment\new-comment.component.ts" "./new-comment.html" "./new-comment.component.html"
Replace-InFile "$Comp\new-comment\new-comment.component.ts" "./new-comment.css"  "./new-comment.component.css"

# --- 8. new-comment/complete -> new-comment/mark-complete ---
git mv "$Comp\new-comment\complete" "$Comp\new-comment\mark-complete"
git mv "$Comp\new-comment\mark-complete\complete.ts"   "$Comp\new-comment\mark-complete\mark-complete.component.ts"
git mv "$Comp\new-comment\mark-complete\complete.html" "$Comp\new-comment\mark-complete\mark-complete.component.html"
git mv "$Comp\new-comment\mark-complete\complete.css"  "$Comp\new-comment\mark-complete\mark-complete.component.css"
Replace-InFile "$Comp\new-comment\mark-complete\mark-complete.component.ts" "./complete.html" "./mark-complete.component.html"
Replace-InFile "$Comp\new-comment\mark-complete\mark-complete.component.ts" "./complete.css"  "./mark-complete.component.css"

# --- 9. powerApps -> power-apps (files already well-named) ---
git mv "$Comp\powerApps" "$Comp\power-apps"

# --- 10. todo-list/change task date -> todo-list/change-task-date ---
git mv "$Comp\todo-list\change task date" "$Comp\todo-list\change-task-date"
git mv "$Comp\todo-list\change-task-date\date.ts"   "$Comp\todo-list\change-task-date\change-task-date.component.ts"
git mv "$Comp\todo-list\change-task-date\date.html" "$Comp\todo-list\change-task-date\change-task-date.component.html"
git mv "$Comp\todo-list\change-task-date\date.css"  "$Comp\todo-list\change-task-date\change-task-date.component.css"
Replace-InFile "$Comp\todo-list\change-task-date\change-task-date.component.ts" "./date.html" "./change-task-date.component.html"
Replace-InFile "$Comp\todo-list\change-task-date\change-task-date.component.ts" "./date.css"  "./change-task-date.component.css"

# --- fix every cross-file reference to the old paths -------------------------
$AppTs = Join-Path $AppRoot 'app.ts'
Replace-InFile $AppTs "./components/mobile navigation/mobile-navigation'" "./components/mobile-navigation/mobile-navigation.component'"
Replace-InFile $AppTs "./components/mobile navigation/mobile-search.utils'" "./components/mobile-navigation/mobile-search.utils'"
Replace-InFile $AppTs "./components/mobile navigation/mobile-navigation.css'" "./components/mobile-navigation/mobile-navigation.css'"
Replace-InFile $AppTs "./components/All-files/all-files'" "./components/all-files/all-files.component'"
Replace-InFile $AppTs "./components/powerApps/powerapps-modal.component'" "./components/power-apps/powerapps-modal.component'"
Replace-InFile $AppTs "./components/new-comment/new-comment'" "./components/new-comment/new-comment.component'"
Replace-InFile $AppTs "./components/loading screen/loading-screen'" "./components/loading-screen/loading-screen.component'"
Replace-InFile $AppTs "./components/delegate-component/delegate'" "./components/delegate/delegate.component'"
Replace-InFile $AppTs "./components/manual-refresh/refresh'" "./components/manual-refresh/manual-refresh.component'"

$TodoTs = Join-Path $Comp 'todo-list\todo-list.component.ts'
Replace-InFile $TodoTs "../new-comment/complete/complete'" "../new-comment/mark-complete/mark-complete.component'"
Replace-InFile $TodoTs "./change task date/date'" "./change-task-date/change-task-date.component'"
Replace-InFile $TodoTs "../Claimable Task/claim'" "../claimable-task/claimable-task.component'"

# --- wrap up ------------------------------------------------------------------
git add -A
git commit -m "Rename component folders/files: no spaces, consistent kebab-case, .component suffix"

Write-Host ""
Write-Host "Done. You're on branch 'codebase-cleanup/component-naming'." -ForegroundColor Green
Write-Host "Review it:      git diff main..codebase-cleanup/component-naming --stat"
Write-Host "Build check:    cd my-app; ng build"
Write-Host "Keep it:        git checkout main; git merge codebase-cleanup/component-naming"
Write-Host "Don't want it:  git checkout main   (nothing on main changed; delete the branch whenever with git branch -D codebase-cleanup/component-naming)"
