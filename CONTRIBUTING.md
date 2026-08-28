# Contributing to IguanaTeX

This document describes the developer-facing repository model introduced in
this fork. It is written for both human contributors and AI coding agents.
End-user installation and usage remain in [README.md](README.md).

## The core rule

The Git-tracked files are the source of truth. A `.pptm` or `.ppam` is a
generated or synchronization container, never the canonical implementation.

- VBA source lives in `src/`.
- VBProject metadata and explicit reference declarations live in
  `office/project/`.
- RibbonX lives in `office/ribbon/`.
- Generated Office artifacts belong under the ignored `.build/` directory or
  another explicitly selected disposable output directory.

Do not use an upstream, release, or previously generated PPTM as the template
for a fresh build. The fresh-build workflow starts with
`PowerPoint.Presentations.Add()` and reconstructs the VBA and Ribbon state from
canonical files.

## Fork-specific history

All changes beginning with commit `0d56666` are developer infrastructure for
this fork. They do not intentionally change IguanaTeX application behavior.

| Commit | Developer-facing change | Consequence |
| --- | --- | --- |
| `0d56666` | Added `scripts/vba-sync.ps1`. | Existing PPTM containers can be imported from, exported to, or compared with canonical VBA files. |
| `a41a394` | Moved VBA modules, classes, forms, and form binaries from the repository root into `src/`; reserialized form snapshots; removed `ExportVBA.bas`. | `src/` is canonical. Documentation or automation must not assume root-level VBA files or the old exporter macro. |
| `d3991fa` | Added canonical project metadata, reference declarations, and both RibbonX versions under `office/`. | Office state that cannot be represented by `.bas`, `.cls`, or `.frm` files is now tracked explicitly. |
| `feb21f9` | Added the fresh PPTM/PPAM build pipeline, isolated VBE compile validation, package/Ribbon validation, PowerPoint round trips, and shared VBA import helpers. | A validated macro-enabled Office artifact can be produced without using any PPTM as build input. |
| `ec95079` | Moved the existing-container sync lookup from the repository root to `.build/office/`. | Build and sync now use the same ignored artifact directory; root-level PPTM files are not sync inputs. |

When changing this infrastructure, preserve the separation between the sync
workflow and the fresh-build workflow. They share low-level import helpers but
have different user-facing purposes.

## Repository map

| Path | Role |
| --- | --- |
| `src/` | Canonical VBA standard modules (`.bas`), class modules (`.cls`), UserForms (`.frm`), and opaque UserForm binaries (`.frx`). |
| `office/project/project.json` | Canonical VBProject name, description, protection state, and conditional-compilation-argument expectation. |
| `office/project/references.json` | References that must be added explicitly to a fresh VBProject. |
| `office/ribbon/customUI.xml` | Canonical Office 2007 RibbonX. |
| `office/ribbon/customUI14.xml` | Canonical Office 2010+ RibbonX. |
| `scripts/vba-sync.ps1` | Existing PPTM ↔ canonical VBA source workflow. |
| `scripts/office-build.ps1` | Canonical source → fresh PPTM, artifact validation, and optional PPAM workflow. |
| `scripts/lib/IguanaTex.Office.psm1` | Shared source-closure, component-import, and COM-release helpers. |
| `scripts/lib/IguanaTex.Compile.psm1` | Compile controller and exact-process recovery boundary. |
| `scripts/lib/IguanaTex.Compile.Worker.ps1` | Isolated PowerPoint/VBE compile worker. |
| `scripts/lib/IguanaTex.Compile.Watcher.ps1` | PID-scoped compile-dialog detector and dismissor. |
| `scripts/lib/IguanaTex.Package.psm1` | OPC/OOXML Ribbon injection and static package validation. |
| `AppleScript/` | macOS PowerPoint integration source. |
| `IguanaTexHelper/` | Swift package for the macOS native helper; see its own `README.md` and `DEVELOP.md`. |
| `.build/` | Ignored disposable outputs. Never commit its contents. |

There is deliberately no canonical PPTM, PPAM, unpacked Office package, or
generated `vbaProject.bin` in the repository.

## Prerequisites for Windows Office automation

The sync and build scripts require:

- Windows desktop PowerPoint;
- Windows PowerShell 5.1 running in an STA apartment;
- PowerPoint's **Trust access to the VBA project object model** setting;
- permission to instantiate PowerPoint and access VBIDE automation;
- the Microsoft Scripting Runtime type library registered on the machine.

Close unrelated PowerPoint windows before running the scripts. PowerPoint is
effectively single-instance in several automation scenarios. The build and
compile code refuses ambiguous COM ownership, but closing other instances is
the most reliable way to obtain a successful isolated run.

The scripts set `AutomationSecurity` to force-disable macros before opening an
artifact. They never need to execute a macro inside the target project to
compile or validate it.

## Choose the correct workflow

| Goal | Command or source area |
| --- | --- |
| Edit canonical VBA directly | Change files under `src/`. |
| Synchronize an existing PPTM under `.build/office/` with `src/` | `scripts/vba-sync.ps1` |
| Build a new PPTM without a PPTM input | `scripts/office-build.ps1 build` |
| Validate an existing generated PPTM or PPAM without modifying the original | `scripts/office-build.ps1 validate` |
| Convert a validated PPTM to PPAM | `scripts/office-build.ps1 ppam` |
| Change project metadata or explicit references | Edit `office/project/`. |
| Change Ribbon controls or callbacks | Edit both canonical files in `office/ribbon/`. |
| Change macOS native helper behavior | Work in `IguanaTexHelper/` and follow its development guide. |

Do not use `vba-sync.ps1` as a substitute for the fresh build. Conversely, do
not expect `office-build.ps1` to update canonical source from an edited PPTM.

## Canonical Office state

### VBA components and UserForms

The current source tree contains 23 importable components: 9 standard modules,
5 classes, and 9 UserForms with 9 matching FRX files. The closure validator
requires:

- exactly one `Attribute VB_Name` per `.bas`, `.cls`, or `.frm`;
- the declared component name to match the filename;
- one matching `.frx` and one matching `OleObjectBlob` reference per form;
- no orphan `.frx` files;
- no duplicate component names across supported source types.

FRX files are accepted opaque snapshots. Do not attempt to normalize or
regenerate them mechanically. PowerPoint can serialize forms and FRX data
nondeterministically, so an export may differ even when the visible form is
unchanged.

PowerPoint does not create a `ThisPresentation` host component for this
project. The build rejects VBComponent type 100 rather than fabricating or
importing a host module.

### VBProject metadata and references

`office/project/project.json` currently declares:

- name: `IguanaTeX`;
- description: `LaTeX Add-In for PowerPoint`;
- protection: `0` (unprotected);
- conditional compilation arguments: empty.

PowerPoint VBIDE does not expose a reliable project-level setter for
conditional compilation arguments. The build therefore accepts only the
canonical empty value and validates the fresh default.

A fresh PowerPoint project is expected to contain one healthy reference each
to `VBA`, `PowerPoint`, `stdole`, and `Office`. Importing a real UserForm adds
`MSForms`; do not inject it explicitly. The only current explicit dependency is
Microsoft Scripting Runtime:

```text
GUID    {420B2830-E718-11CF-893D-00A0C9054228}
Version 1.0
```

The build rechecks GUID, major/minor version, uniqueness, and `IsBroken` after
adding the reference. Removing it causes compilation to fail at
`FileSystemObject`, which is an intentional negative test for the compile gate.

### RibbonX

Both Ribbon files are authoritative and currently describe the same IguanaTeX
tab: three groups, eight buttons, and eight callbacks. They use built-in
`imageMso` icons only. There are no custom images, `customUI/*.rels`, or outgoing
Ribbon relationships.

The package locations and root relationship types are:

| Part | Root relationship type |
| --- | --- |
| `customUI/customUI.xml` | `http://schemas.microsoft.com/office/2006/relationships/ui/extensibility` |
| `customUI/customUI14.xml` | `http://schemas.microsoft.com/office/2007/relationships/ui/extensibility` |

An `onAction` name must resolve to exactly one public or default-public Sub or
Function in a standard `.bas` module. Do not rename callbacks in only one side
of the Ribbon/VBA boundary.

## VBA sync workflow

`vba-sync.ps1` operates on exactly one non-lock `.pptm` in `.build/office/`.
It ignores `~$` lock files and fails if the directory is absent or contains
zero or multiple candidates. The default fresh-build output
`.build/office/IguanaTeX.pptm` is therefore discovered automatically;
root-level PPTM files are ignored.

```powershell
# Canonical source -> existing PPTM in .build\office
.\scripts\vba-sync.ps1 import -Prune

# Existing PPTM in .build\office -> canonical text source
.\scripts\vba-sync.ps1 export

# Compare an exported staging snapshot with canonical text source
.\scripts\vba-sync.ps1 verify
```

Important options:

| Option | Meaning |
| --- | --- |
| `-Prune` | On import, remove supported components absent from `src/`. On export, remove canonical text files absent from the PPTM and the matching FRX for a removed form. |
| `-DryRun` | Report import/export mutations without applying them. |
| `-Visible` | Show the PowerPoint window used by the sync operation. |
| `-UpdateFrx FormName` | During export, replace the named form's canonical FRX. Accepts multiple names or `all`. Existing FRX files are otherwise preserved. |
| `-VerifyFrx` | Include byte-level FRX comparison in `verify`; this can report nondeterministic differences. |

Use `export -DryRun` before updating a broad set of canonical files. Review
every `.frm` change and update FRX only when the form binary state was
intentionally changed. The sync workflow handles VBA components only; it does
not rebuild project metadata, references, RibbonX, or the OOXML package.

## Fresh Office artifact workflow

### CLI

```powershell
# Fresh canonical source -> default .build\office\IguanaTeX.pptm
.\scripts\office-build.ps1 build

# Build PPTM and PPAM in one validated run
.\scripts\office-build.ps1 build `
    -PpamOutputPath .\.build\office\IguanaTeX.ppam

# Validate without changing the supplied file
.\scripts\office-build.ps1 validate `
    -InputPath .\.build\office\IguanaTeX.pptm

# Validated PPTM -> PPAM
.\scripts\office-build.ps1 ppam `
    -InputPath .\.build\office\IguanaTeX.pptm `
    -OutputPath .\.build\office\IguanaTeX.ppam
```

The three actions are intentionally distinct:

- `build` rejects `-InputPath`; a PPTM must never become a fresh-build input.
- `validate` copies its input to a temporary directory before compile, save, or
  add-in operations, and rejects output/overwrite options.
- `ppam` first validates its PPTM input, then creates a staged add-in.

`-OutputPath` selects a nondefault destination. `-Force` is required to replace
an existing output. `-CompileTimeoutSeconds` and `-OfficeTimeoutSeconds` default
to 60 and accept values from 10 through 600. `-Visible` is for diagnostics, not
normal unattended builds.

Outputs are written to short random staging names such as `IT-<token>.pptm` and
published only after all gates pass. A failed build does not intentionally
replace the requested destination.

### PPTM pipeline

The `build` action performs these stages:

1. Validate canonical VBA/form closure and read project JSON.
2. Explicitly launch a hidden PowerPoint `/AUTOMATION` process and bind COM only
   when its HWND maps to that exact PID and process start time.
3. Create a fresh presentation and `SaveAs` macro-enabled PPTM format 25.
4. Apply project metadata, add explicit references, import canonical
   components, and validate components and references including auto-added
   MSForms.
5. Save and completely close PowerPoint. The initial slide/master/theme and
   document properties are disposable fresh scaffolding.
6. Run the VBE compile gate in a separate helper process.
7. Confirm that PowerPoint produced `ppt/vbaProject.bin`, its presentation
   relationship, and the `.bin` VBA content type.
8. Inject both RibbonX parts by direct OPC editing, then run static package,
   relationship, XML, and callback validation.
9. Open/save/close the PPTM in PowerPoint, validate the package again, reopen it
   read-only, and recheck metadata, components, references, and package state.

The final PPTM has never used another PPTM as source or template.

### Validated implementation baseline

When commit `feb21f9` introduced this pipeline, it was exercised against the
canonical source in this repository on Windows PowerPoint. The verified
baseline included:

- a canonical-only fresh PPTM build with project metadata, Scripting Runtime,
  all modules/classes/forms, and a passing VBE compile;
- package closure, both Ribbon versions, callback resolution, and Ribbon
  persistence across PowerPoint save/reopen cycles;
- a PPAM conversion followed by package checks and two add-in load/unload/remove
  cycles;
- a missing-Scripting case that failed compilation at `FileSystemObject`, had
  its compiler dialog detected and dismissed, and exited nonzero;
- an intentionally undefined VBA type with the same controlled compile failure
  and no permanent modal hang;
- an unrelated PowerPoint sentinel process remaining alive while ambiguous COM
  ownership failed safely.

This is a regression baseline, not a substitute for rerunning the relevant
tests after changing Office automation, package handling, references, or VBA.

### PPAM pipeline

PPAM support uses PowerPoint macro-enabled add-in format 30:

1. Validate the source PPTM on a temporary copy.
2. Open the validated source and `SaveAs` a staged PPAM.
3. Reinject RibbonX after the final PowerPoint save if Office removed or rewrote
   it.
4. Repeat VBA, Ribbon, callback, content-type, and internal-target validation.
5. Register, load, inspect, unload, remove, and repeat the add-in cycle to prove
   that it can be reloaded without leaving a registration behind.

PowerPoint does not provide a safe way to target the VBE Compile command at a
specific loaded add-in project. The PPAM therefore inherits the successfully
compiled PPTM source and is validated by package inspection and two complete
add-in load cycles; it is not separately VBE-compiled after `SaveAs`.

## Compile validation design

The verified VBE compile command is CommandBars control ID `578`. A successful
fresh compile has this state transition:

```text
Enabled before = True
Execute compile
Enabled after  = False
```

Compile errors do not propagate as a COM exception. `Execute()` returns while a
modal **Microsoft Visual Basic for Applications** dialog remains open, and
additional COM calls may be rejected until that dialog is dismissed. Therefore
ordinary `try`/`catch` is not a compiler-error detector.

The implementation uses three processes/roles:

- `IguanaTex.Compile.psm1` controls timeout, protocol files, and recovery.
- `IguanaTex.Compile.Worker.ps1` explicitly launches PowerPoint, proves exact
  PID/start-time/COM-HWND ownership, opens the target, and invokes control 578.
- `IguanaTex.Compile.Watcher.ps1` watches only that PowerPoint PID for a new VBE
  dialog, dismisses it, and records the outcome.

The worker publishes its exact ownership record before COM activation so the
controller can recover from an activation hang. It separates ownership of the
explicitly launched process from proof that COM bound to that process. A bind
mismatch is `IsolationFailed`: the target file is not opened and an unrelated
PowerPoint instance is never sent `Quit` or killed.

Expected result classes include `Passed`, `AlreadyCompiled` when explicitly
allowed by validation, `CompileErrorDialog`, `CompileCommandStillEnabled`,
`Timeout`, `ComRecoveryFailed`, and `IsolationFailed`. Any result other than an
allowed success causes the build entry point to exit nonzero.

COM references are released leaf-to-root. The important compile chain is:

```text
compile command
-> CommandBars
-> VBE MainWindow
-> VBE
-> Presentation
-> Presentations
-> PowerPoint Application
```

After final release and garbage collection, only the exact process whose PID,
name, start time, and explicit-launch record still match is eligible for
fallback termination. Never replace this with `Stop-Process -Name POWERPNT`,
`taskkill /IM POWERPNT.EXE`, or any other name-wide cleanup.

Noncompile PowerPoint operations use the same exact-process boundary plus a
PID-scoped modal watchdog. A repair, corruption, add-in, save, close, or quit
dialog is a validation failure rather than an unattended hang.

## Package and Ribbon validation design

Ribbon injection occurs only after PowerPoint and VBIDE have completely closed
the PPTM or PPAM. `IguanaTex.Package.psm1` merges into `_rels/.rels` and
`[Content_Types].xml`; it does not replace either package-wide file. Relationship
IDs are generated collision-free and are not canonical source state.

The VBA closure gate requires:

- nonempty `ppt/vbaProject.bin`;
- exactly one logical presentation relationship of type
  `http://schemas.microsoft.com/office/2006/relationships/vbaProject` resolving
  to `ppt/vbaProject.bin`;
- exactly one `.bin` Default with content type
  `application/vnd.ms-office.vbaProject`;
- the correct macro-enabled main content type for `.pptm` or `.ppam`.

The full package gate also requires:

- both canonical Ribbon parts and both root extensibility relationships;
- parseable XML and byte-for-byte canonical Ribbon content;
- no duplicate logical Ribbon relationship;
- every Ribbon relationship target to exist;
- every internal relationship in the package to resolve to an existing part;
- every `onAction` callback to resolve unambiguously in canonical VBA.

If a package already has the canonical Ribbon state, injection is a no-op. When
mutation is necessary, the package writer preserves Office content and repairs
ZIP entry ordering so `[Content_Types].xml` and `_rels/.rels` appear first. This
ordering is not meaningful source state, but the PowerPoint AddIns loader in the
validated Office environment rejected otherwise valid PPAM archives when those
entries were moved to the end.

## Making common changes

### Change VBA code

Edit the relevant `.bas` or `.cls` under `src/`, preserve its `VB_Name`, then run
a fresh build. Do not edit application behavior merely to simplify automation.

### Change a UserForm

Use PowerPoint/VBE when binary form state changes. Export to a disposable stage,
review the `.frm`, and update only the intended `.frx` with `-UpdateFrx`. Keep the
form filename, `VB_Name`, `OleObjectBlob`, and FRX basename aligned.

### Add a VBA component

Add the canonical file under `src/`. The fresh build imports it automatically.
If it introduces an external type-library dependency, declare that dependency
in `office/project/references.json` and add a negative test proving why it is
required.

### Change project metadata or references

Edit the JSON source, not `vbaProject.bin`. Do not add default Office references
or MSForms to `references.json`; only dependencies that a fresh PowerPoint
project cannot acquire implicitly belong there.

### Change the Ribbon

Edit both Ribbon versions unless the schema difference is intentional. Use
built-in `imageMso` icons unless the task explicitly introduces and validates
custom image relationships. Add or rename the public VBA callback in the same
change.

### Change build or package automation

Keep public roles separate, reuse `IguanaTex.Office.psm1` for common import
logic, and preserve exact-PID ownership and staged publishing. Package changes
must be tested in PowerPoint, not only with an XML/ZIP parser.

## Validation expectations

### Fast checks

At minimum, parse every PowerShell entry point/module with the Windows
PowerShell parser and validate source closure:

```powershell
Import-Module .\scripts\lib\IguanaTex.Office.psm1 `
    -Force -DisableNameChecking
Assert-VbaSourceClosure -SourceDirectory .\src
```

Documentation-only changes do not require Office artifact regeneration, but
commands and paths in the documentation must be checked against the actual
parameter declarations and current tree.

### Full happy path

For changes affecting VBA, metadata, references, Ribbon, PowerPoint automation,
or packaging, run:

```powershell
.\scripts\office-build.ps1 build -Force `
    -PpamOutputPath .\.build\office\IguanaTeX.ppam

.\scripts\office-build.ps1 validate `
    -InputPath .\.build\office\IguanaTeX.pptm

.\scripts\office-build.ps1 validate `
    -InputPath .\.build\office\IguanaTeX.ppam
```

The acceptance chain is canonical-only input → fresh PPTM → metadata and
references → component/form import → real VBE compile → VBA package closure →
Ribbon injection → static package validation → PowerPoint save/reopen → optional
PPAM conversion and add-in reload.

### Negative tests for compile or ownership changes

Use only disposable copies and verify all of the following:

- remove Microsoft Scripting Runtime and confirm a detected compile-error
  dialog, `EnabledAfter = True`, nonzero validation, and successful cleanup;
- add a module with an intentionally undefined type and confirm the same
  controlled failure without a persistent dialog;
- keep a separate PowerPoint presentation/process alive and confirm an
  ambiguous validator returns `IsolationFailed` without closing or killing the
  sentinel process;
- exercise timeout/recovery logic when the implementation being changed affects
  it;
- confirm no validator-owned `POWERPNT.EXE` or protocol directory remains.

Never inject broken test source into `src/` or commit negative-test artifacts.

### Ribbon and PPAM smoke tests

Automation intentionally does not click the Ribbon. Before a release, load the
generated PPAM manually and confirm that the IguanaTeX tab contains three groups
and eight buttons. Exercise the buttons in an environment with the normal
LaTeX, Ghostscript, ImageMagick, and optional helper dependencies as applicable.

## Rules for human and AI contributors

Before editing:

1. Read `README.md`, this file, the target script, and relevant JSON/XML source.
2. Run `git status` and preserve unrelated tracked changes and untracked files.
3. Inspect history when a behavior or serialization choice is unclear.
4. Decide explicitly whether the task belongs to sync, fresh build, package
   validation, runtime VBA, or a platform helper.

While editing:

- Treat `src/` and `office/` as authoritative; never reverse-engineer source
  state from an arbitrary binary artifact when canonical input exists.
- Keep generated PPTM, PPAM, unpacked packages, protocol files, screenshots, and
  diagnostic dumps out of the repository root and source directories.
- Do not broaden a developer-infrastructure task into an IguanaTeX behavior,
  UI, callback-name, or Ribbon-layout refactor.
- Preserve opaque FRX data unless a form change explicitly requires updating it.
- Do not terminate PowerPoint by executable name. Exact PID and start-time
  verification are mandatory.
- Do not weaken validation merely to make a generated artifact pass.
- Prefer shared helpers over copying import or COM cleanup logic into another
  entry point.
- Fail safely if ownership, package state, metadata, or reference identity is
  ambiguous.

Before handing off:

1. Run validation proportional to the affected layer.
2. Check that no PowerPoint process or temporary Office directory owned by the
   test remains.
3. Review `git diff --check` and `git status`.
4. Confirm `.build/` and user-supplied local files are not staged.
5. Report implemented behavior, reused code, commands actually run, passed
   criteria, manual checks, and remaining Office/VBE risks.

## Commit and review guidance

Keep a commit focused on one coherent change. A useful commit body explains the
canonical state affected, the safety/validation behavior, and the tests run.
Do not commit generated Office artifacts or temporary unpacked packages.

Reviewers should pay particular attention to:

- accidental use of a PPTM as fresh-build input;
- source/component/FRX closure;
- explicit versus implicit references;
- exact COM process ownership and cleanup paths;
- compiler-dialog and timeout behavior;
- root relationship merging rather than replacement;
- internal OPC target closure and callback resolution;
- PPAM unload/removal and Ribbon persistence;
- destructive overwrite behavior and staging cleanup.

## Known Office/VBE caveats

- PowerPoint COM activation may attach to an existing instance; automation must
  prove exact ownership or stop safely.
- A compile-error modal is not surfaced by `CommandBars.Execute()` as an
  exception.
- COM can reject calls while a modal dialog exists.
- UserForm text and FRX serialization can change across export/import cycles.
- The PPTM is first injected only after its build session closes, then reopened
  and saved to prove Ribbon persistence. PPAM conversion reinjects after its
  final `SaveAs` because PowerPoint may remove or rewrite Ribbon parts.
- PPAM is not separately compiled after conversion because VBE does not safely
  target a specific loaded add-in project.
- Static OPC validity is necessary but not sufficient; real PowerPoint reopen or
  add-in load validation remains required.

These are design constraints, not reasons to bypass a gate.
