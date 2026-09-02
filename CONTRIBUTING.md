# Contributing to IguanaTeX

This document describes the developer-facing repository model introduced in
this fork. It is written for both human contributors and AI coding agents.
End-user installation and usage remain in [README.md](README.md).

## The core rule

The Git-tracked files are the source of truth. A `.pptm` or `.ppam` is a
generated or synchronization container, never the canonical implementation.

- Standard-module and class source lives in `src/*.bas` and `src/*.cls`.
- Each UserForm is represented by `src/<FormName>.json`, the adjacent
  `src/<FormName>.vba`, and any assets referenced by relative `file://` URIs in
  that JSON under `src/<FormName>/`. `office/forms/manifest.json` is the closed
  list of canonical forms.
- VBProject metadata and explicit reference declarations live in
  `office/project/`.
- RibbonX lives in `office/ribbon/`.
- Native `.frm/.frx` pairs are generated with the pinned FrxEdit submodule into
  `.build/vba-source/`; they are never edited or committed as source.
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
| `7a2b8de` | Last commit before the UserForm source transition. | This is the previous binary-canonical baseline to consult when migration history is required. |
| `3bbfb96` (2026-09-02) | Replaced canonical `.frm/.frx` pairs with JSON templates, separate VBA, and extracted assets. | This is the exact point where authority moved to `src/*.json`, matching `src/*.vba`, and referenced assets. Native form pairs are generated artifacts from this commit onward. |

When changing this infrastructure, preserve the separation between the sync
workflow and the fresh-build workflow. They share low-level import helpers but
have different user-facing purposes.

## Repository responsibilities and dependency direction

FrxEdit and IguanaTeX deliberately have different responsibilities:

- `frx-edit` owns the generic MS-OFORMS reader/writer, JSON edit contract,
  structural validation, binary invariants, and minimal generic regression
  fixtures. It must not acquire IguanaTeX-specific workarounds or dependencies.
- IguanaTeX consumes that codec and owns the nine-form canonical corpus, VBA and
  image assets, the full-corpus round-trip matrix, Office packaging, and native
  PowerPoint acceptance.

The dependency is one-way: IguanaTeX depends on FrxEdit through
`tools/frx-edit`; FrxEdit does not depend on IguanaTeX. The submodule gitlink is
the codec version lock, the publish script is the build-contract lock, and the
generated provenance record is execution evidence. Do not introduce a second
lock file containing the same revision.

## Repository map

| Path | Role |
| --- | --- |
| `src/` | Canonical standard modules (`.bas`), classes (`.cls`), UserForm templates (`.json`), matching form code (`.vba`), and JSON-referenced assets. |
| `office/forms/manifest.json` | Versioned, explicit list of the nine canonical UserForms and their JSON/VBA paths. |
| `office/project/project.json` | Canonical VBProject name, description, protection state, and conditional-compilation-argument expectation. |
| `office/project/references.json` | References that must be added explicitly to a fresh VBProject. |
| `office/ribbon/customUI.xml` | Canonical Office 2007 RibbonX. |
| `office/ribbon/customUI14.xml` | Canonical Office 2010+ RibbonX. |
| `tools/frx-edit/` | Git submodule containing the pinned generic MSForms codec. |
| `scripts/frxedit-build.ps1` | Validate/publish the codec and record build provenance. |
| `scripts/userforms-build.ps1` | Canonical JSON/VBA/assets to a complete generated VBA import tree. |
| `scripts/userforms-export.ps1` | Explicit, transactional native-form export back to selected canonical forms. |
| `scripts/userforms-verify.ps1` | Pinned nine-form create/rebuild/watch/template semantic matrix. |
| `scripts/vba-sync.ps1` | Existing PPTM and canonical VBA/UserForm synchronization workflow. |
| `scripts/office-build.ps1` | Canonical source to fresh PPTM, artifact validation, and optional PPAM workflow. |
| `scripts/lib/IguanaTex.UserForms.psm1` | Shared manifest, submodule, asset, codec, and semantic-comparison helpers. |
| `scripts/lib/IguanaTex.Office.psm1` | Shared source-closure, component-import, and COM-release helpers. |
| `scripts/lib/IguanaTex.Compile.psm1` | Compile controller and exact-process recovery boundary. |
| `scripts/lib/IguanaTex.Compile.Worker.ps1` | Isolated PowerPoint/VBE compile worker. |
| `scripts/lib/IguanaTex.Compile.Watcher.ps1` | PID-scoped compile-dialog detector and dismissor. |
| `scripts/lib/IguanaTex.Package.psm1` | OPC/OOXML Ribbon injection and static package validation. |
| `AppleScript/` | macOS PowerPoint integration source. |
| `IguanaTexHelper/` | Swift package for the macOS native helper; see its own `README.md` and `DEVELOP.md`. |
| `.build/frxedit/` | Published `frxedit.exe` and provenance for the current run. |
| `.build/vba-source/` | Generated `.frm/.frx` plus copied `.bas/.cls`; the complete Office import tree. |
| `.build/userforms-verify/` | Disposable UserForm matrix reports and retained diagnostics. |
| `.build/office/` | Generated PPTM/PPAM artifacts used by build and sync workflows. |

There is deliberately no canonical `.frm`, `.frx`, PPTM, PPAM, unpacked Office
package, or generated `vbaProject.bin` in the repository.

## Clone and initialize the codec submodule

Clone with submodules when possible:

```powershell
git clone --recurse-submodules <repository-url>
```

For an existing checkout, after switching branches, or after pulling a commit
that advances the gitlink, run:

```powershell
git submodule update --init --recursive
git submodule status --recursive
```

Do not independently pull or float `tools/frx-edit` during normal IguanaTeX
work. The default `Pinned` script mode requires the submodule to be initialized,
at the recorded gitlink commit, and clean. `WorkingTree` mode exists only for
deliberate local codec development and does not qualify as pinned acceptance.
This integration tracks `https://github.com/mirinae3145/frx-edit.git` at
`9cf15bbbde76e0251dc6d6b988911fc1cda9af4b`; the gitlink, rather than this
explanatory sentence, remains authoritative after a future reviewed upgrade.

## Development prerequisites and Windows Office automation

Publishing and generating forms require Git and the .NET 8 SDK. The sync and
Office build scripts additionally require:

- Windows desktop PowerPoint;
- Windows PowerShell 5.1 running in an STA apartment;
- PowerPoint's **Trust access to the VBA project object model** setting;
- permission to instantiate PowerPoint and access VBIDE automation.

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
| Edit a standard module or class | Change `src/*.bas` or `src/*.cls`. |
| Edit a UserForm | Change its manifest-listed JSON, matching `.vba`, or referenced assets, then run the UserForm build and matrix. |
| Publish the pinned codec | `scripts/frxedit-build.ps1` |
| Generate the complete Office import tree | `scripts/userforms-build.ps1` |
| Verify all nine canonical forms | `scripts/userforms-verify.ps1` |
| Synchronize an existing PPTM under `.build/office/` with `src/` | `scripts/vba-sync.ps1` |
| Build a new PPTM without a PPTM input | `scripts/office-build.ps1 build` |
| Validate an existing generated PPTM or PPAM without modifying the original | `scripts/office-build.ps1 validate` |
| Convert a PPTM to PPAM, validating by default | `scripts/office-build.ps1 ppam` |
| Change project metadata or explicit references | Edit `office/project/`. |
| Change Ribbon controls or callbacks | Edit both canonical files in `office/ribbon/`. |
| Change macOS native helper behavior | Work in `IguanaTexHelper/` and follow its development guide. |

Do not use `vba-sync.ps1` as a substitute for the fresh build. Conversely, do
not expect `office-build.ps1` to update canonical source from an edited PPTM.

## Canonical Office state

### VBA components and UserForms

The current source tree contains 24 importable components: 10 standard modules,
5 classes, and these 9 UserForms: `AboutBox`, `BatchEditForm`, `ErrorForm`,
`ExternalEditorForm`, `LatexForm`, `LoadVectorGraphicsForm`, `LogFileViewer`,
`RegenerateForm`, and `SetTempForm`.

`office/forms/manifest.json` has a schema version and exactly one name/JSON/VBA
entry for each form. The UserForm closure validator requires:

- exactly the manifest-listed form set, with no duplicate names or paths;
- each JSON and same-form VBA path to use repository-relative forward slashes,
  resolve to top-level sibling files in `src/`, and use the manifest name as its
  basename;
- each relative `file://` asset reference to resolve from its JSON directory,
  remain within that form's owned `src/<FormName>/` asset directory, and name an
  existing file;
- every file in an owned asset directory to be referenced, with no orphan asset;
- no unlisted top-level UserForm JSON/VBA pair;
- generated `.frm` names, `Attribute VB_Name`, `OleObjectBlob` references, and
  `.frx` basenames to agree;
- no duplicate component names across the generated `.frm`, copied `.bas`, and
  copied `.cls` files.

`userforms-build.ps1` publishes FrxEdit if necessary, performs strict `create`
and `validate` for every manifest entry, and writes the native pairs only to the
selected staging directory (default `.build/vba-source/`). It also copies the
canonical `.bas/.cls` files there so Office import and closure checks consume one
complete, generated tree. A byte difference between two FRX files is not a
semantic failure by itself; comparisons are based on strict MSForms semantics
and form VBA.

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
`MSForms`; do not inject it explicitly. There are currently no explicit project
references in `references.json`. Windows runtime use of FileSystemObject and
Dictionary COM objects is late-bound, and the build rejects a Microsoft
Scripting Runtime project reference if one is present.

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

## Codec and UserForm generation

The normal entry points are:

```powershell
# Validate and publish the exact submodule revision
.\scripts\frxedit-build.ps1

# Generate .frm/.frx plus copied .bas/.cls for Office import
.\scripts\userforms-build.ps1

# Run the complete nine-form semantic matrix
.\scripts\userforms-verify.ps1
```

`frxedit-build.ps1` defaults to `-Mode Pinned`. It rejects an uninitialized
submodule, a submodule HEAD that does not equal the recorded gitlink, or a dirty
submodule. It publishes `src/FrxEdit.Cli/FrxEdit.Cli.csproj` as
Release/win-x64/self-contained/single-file/untrimmed to
`.build/frxedit/frxedit.exe`. The command reports the codec commit, .NET SDK
version, and executable SHA-256 and records them in a provenance JSON file.

`userforms-build.ps1` uses the same pinned mode by default and writes to
`.build/vba-source/`; use `-OutputDirectory` only for a disposable alternate
stage. `-Mode WorkingTree` on these two commands is an explicit development
escape hatch for testing an uncommitted codec fix. Do not cite its output as a
pinned or release result.

`userforms-verify.ps1` always requires the pinned, clean codec. For every form it
performs strict creation and validation, no-op rebuilding, patch reapplication,
bounded watch regeneration, recreation-template generation, and semantic
comparison. It also proves that canonical input hashes did not change. The
matrix uses `.build/userforms-verify/` for reports; `-KeepArtifacts` retains
detailed successful intermediates for diagnosis, while failures are always
retained and each successful run updates `summary.json`. `-SkipWatch` is a
focused diagnostic option, not a complete acceptance run.

### Default UserForm editing sequence

This is the fixed workflow for an AI agent or a contributor without prior
repository knowledge:

1. Confirm the submodule and both working-tree statuses, then locate the form in
   `office/forms/manifest.json`.
1. Modify only the listed canonical JSON, matching VBA, and referenced assets.
1. Run `userforms-build.ps1` to produce and inspect a strict native
   reconstruction.
1. Run `userforms-verify.ps1` without `-SkipWatch` for the full nine-form matrix.
1. Build the PPTM/PPAM and pass the native Office import, VBE compile,
   save/reopen, package, and add-in-load gates.
1. For a release or UI/MSForms change, perform the representative GUI/runtime
   smoke test.

Do not shortcut this sequence by editing a generated `.frm/.frx`. If a stage
fails, use the ownership rules in "Diagnose or fix a codec defect" below before
deciding which repository to change.

## VBA sync workflow

`vba-sync.ps1` operates on exactly one non-lock `.pptm` in `.build/office/`.
It ignores `~$` lock files and fails if the directory is absent or contains
zero or multiple candidates. The default fresh-build output
`.build/office/IguanaTeX.pptm` is therefore discovered automatically;
root-level PPTM files are ignored.

```powershell
# Generated canonical source -> existing PPTM in .build\office
.\scripts\vba-sync.ps1 import -Prune

# Preview module/class export without changing canonical files
.\scripts\vba-sync.ps1 export -DryRun

# Explicitly promote one edited native form to JSON/VBA/assets
.\scripts\vba-sync.ps1 export -UpdateUserForm LatexForm -DryRun

# Compare modules/classes textually and UserForms semantically
.\scripts\vba-sync.ps1 verify
```

Important options:

| Option | Meaning |
| --- | --- |
| `-Prune` | On import, remove supported components absent from the generated canonical tree. On export, pruning applies to standard modules/classes; it never authorizes an implicit canonical UserForm deletion. |
| `-DryRun` | Report import/export mutations without applying them. |
| `-Visible` | Show the PowerPoint window used by the sync operation. |
| `-UpdateUserForm FormName` | During export, promote only the named forms to canonical JSON/VBA/assets. Accepts multiple explicit names. All selected forms are validated before any of their canonical files are replaced. |

`import` automatically regenerates the pinned `.build/vba-source/` tree before
opening PowerPoint. `verify` compares `.bas/.cls` text as before, then compares
the canonical generated forms with the PPTM export using form VBA and strict
MSForms semantic equality; it never uses FRX byte equality. VBA text comparison
keeps the historical byte-exact check as its first path, then permits only line
ending, final-newline, and VBE identifier-casing normalization. String literals,
apostrophe or `Rem` comment text, whitespace, and actual token changes still
fail verification. `export` continues to synchronize standard modules/classes,
but a UserForm is ignored unless it is named with `-UpdateUserForm`. Multiple
selected forms are converted and checked in temporary storage before they are
applied together.

The retired `-UpdateFrx` and `-VerifyFrx` options intentionally produce a
migration error. Use `-UpdateUserForm` and the always-semantic `verify` action
instead. Always run `export -DryRun` first and inspect the JSON, VBA, and asset
changes. Sync still does not author project metadata, references, RibbonX, or
the OOXML package.

For a directory that already contains explicitly exported native pairs, the
lower-level equivalent is:

```powershell
.\scripts\userforms-export.ps1 `
    -InputDirectory C:\path\to\native-export `
    -Form LatexForm `
    -DryRun
```

`-Form` is mandatory; there is no implicit "all" mode. The command extracts a
strict recreation template, VBA, and referenced assets into temporary storage,
validates every requested form, and only then replaces canonical files unless
`-DryRun` was supplied.

## Fresh Office artifact workflow

### CLI

```powershell
# Fresh canonical source -> default .build\office\IguanaTeX.pptm
.\scripts\office-build.ps1 build

# Build PPTM and PPAM in one validated run
.\scripts\office-build.ps1 build `
    -PpamOutputPath .\.build\office\IguanaTeX.ppam

# Fast local build with artifact validation deferred
.\scripts\office-build.ps1 build -NoValidation

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
- `ppam` validates its PPTM input by default, then creates a staged add-in.

Every action first regenerates a pinned `.build/vba-source/` tree. Source
closure, callback resolution, import, and comparison therefore all operate on
the same native representation derived from the manifest-listed canonical
JSON/VBA/assets. A stale `.frm/.frx` export is never accepted as implicit input.

`-OutputPath` selects a nondefault destination. `-Force` is required to replace
an existing output. `-CompileTimeoutSeconds` and `-OfficeTimeoutSeconds` default
to 60 and accept values from 10 through 600. `-Visible` is for diagnostics, not
normal unattended builds.

`-NoValidation` is accepted by `build` and `ppam` for fast local iteration. It
skips VBE compile validation, explicit post-build package gates, PowerPoint
save/reopen validation, and PPAM load/unload validation. With
`build -PpamOutputPath`, it applies to both outputs. Canonical input checks, exact
PowerPoint process ownership, staged publishing, and transactional Ribbon
injection still run because they are required to construct an artifact safely.
The resulting files are explicitly reported as unverified; run the `validate`
action without `-NoValidation` before distributing or relying on them.

Outputs are written to short random staging names such as `IT-<token>.pptm` and
published only after all enabled gates pass. A failed build does not
intentionally replace the requested destination.

### PPTM pipeline

The `build` action performs these stages:

1. Publish the pinned codec, generate all native forms into
   `.build/vba-source/`, validate the complete source closure, and read project
   JSON.
1. Explicitly launch a hidden PowerPoint `/AUTOMATION` process and bind COM only
   when its HWND maps to that exact PID and process start time.
1. Create a fresh presentation and `SaveAs` macro-enabled PPTM format 25.
1. Apply project metadata, import the generated staging components, and validate
   components
   and references including auto-added MSForms and the absence of Microsoft
   Scripting Runtime.
1. Save and completely close PowerPoint. The initial slide/master/theme and
   document properties are disposable fresh scaffolding.
1. Run the VBE compile gate in a separate helper process.
1. Confirm that PowerPoint produced `ppt/vbaProject.bin`, its presentation
   relationship, and the `.bin` VBA content type.
1. Inject both RibbonX parts by direct OPC editing, then run static package,
   relationship, XML, and callback validation.
1. Open/save/close the PPTM in PowerPoint, validate the package again, reopen it
   read-only, and recheck metadata, components, references, and package state.

The final PPTM has never used another PPTM as source or template.

### Validated implementation baseline

When commit `feb21f9` introduced this pipeline, it was exercised against the
canonical source in this repository on Windows PowerPoint. The verified
baseline included:

- a canonical-only fresh PPTM build with project metadata, all
  modules/classes/forms, and a passing VBE compile;
- package closure, both Ribbon versions, callback resolution, and Ribbon
  persistence across PowerPoint save/reopen cycles;
- a PPAM conversion followed by package checks and two add-in load/unload/remove
  cycles;
- an intentionally undefined VBA type with a controlled compile failure
  and no permanent modal hang;
- an unrelated PowerPoint sentinel process remaining alive while ambiguous COM
  ownership failed safely.

This is a regression baseline, not a substitute for rerunning the relevant
tests after changing Office automation, package handling, references, or VBA.

### PPAM pipeline

PPAM support uses PowerPoint macro-enabled add-in format 30:

1. Validate the source PPTM on a temporary copy.
1. Open the validated source and `SaveAs` a staged PPAM.
1. Reinject RibbonX after the final PowerPoint save if Office removed or rewrote
   it.
1. Repeat VBA, Ribbon, callback, content-type, and internal-target validation.
1. Register, load, inspect, unload, remove, and repeat the add-in cycle to prove
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

Start in IguanaTeX and edit only the form's manifest-listed `.json`, matching
`.vba`, and referenced assets. Then run:

```powershell
.\scripts\userforms-build.ps1
.\scripts\userforms-verify.ps1
.\scripts\office-build.ps1 build -Force
```

If a change must originate in PowerPoint/VBE, export it to disposable native
files and use `userforms-export.ps1 -Form <name> -DryRun`, or use
`vba-sync.ps1 export -UpdateUserForm <name> -DryRun` for the single PPTM under
`.build/office/`. Review the regenerated JSON, VBA, and assets before running the
command without `-DryRun`. Never directly edit the generated `.frm/.frx` in
`.build/vba-source/` and treat it as source.

### Diagnose or fix a codec defect

First isolate ownership of the failure:

- Incorrect canonical intent, invalid form data, form VBA, or an IguanaTeX asset
  is fixed in IguanaTeX.
- Valid JSON that FrxEdit reads or writes incorrectly, or native persistence that
  PowerPoint rejects for a general MSForms reason, is a FrxEdit defect.
- PPTM/PPAM packaging, VBProject references, RibbonX, and Office automation remain
  IguanaTeX defects.

For a codec defect, work in `tools/frx-edit` on a clean `fix-<issue>` branch.
Reduce the IguanaTeX failure to the smallest generic MSForms fixture or invariant,
add a permanent FrxEdit regression, fix and test FrxEdit, and publish it locally.
Use `-Mode WorkingTree` only while exercising that candidate against the complete
IguanaTeX matrix and native Office validation. Examples of generic invariants
include FormStreamData/FormSiteData boundaries, `GuidAndStdFont` versus
`GuidAndTextProps`, MultiPage/TabStrip allocation, and persisted
FormDesignExData.

Merge and push the FrxEdit fix first. In a separate IguanaTeX integration change,
advance `tools/frx-edit` to that exact commit and rerun the pinned validation
layers below. Do not combine an uncommitted submodule worktree and a gitlink
update, and do not make FrxEdit depend on the IguanaTeX corpus to implement a
product-specific workaround.

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

At minimum, parse every changed PowerShell entry point/module with the Windows
PowerShell 5.1 parser. Then generate the native staging tree and validate source
closure:

```powershell
.\scripts\userforms-build.ps1

Import-Module .\scripts\lib\IguanaTex.Office.psm1 `
    -Force -DisableNameChecking
Assert-VbaSourceClosure -SourceDirectory .\.build\vba-source
```

Documentation-only changes do not require Office artifact regeneration, but
commands and paths in the documentation must be checked against the actual
parameter declarations and current tree.

### Validation layers and release gate

Validation is intentionally layered so a failure is routed to the repository
that owns it:

1. **FrxEdit unit/conformance:** exercise the generic MS-OFORMS structures and
   generated-container invariants in `tools/frx-edit`.
1. **IguanaTeX canonical matrix:** run all nine forms through strict create,
   no-op, patch, bounded watch, template recreation, and semantic comparison.
1. **Native Office build:** import the generated tree, compile in the VBE,
   validate/package, and prove PowerPoint save/reopen (plus PPAM load cycles when
   producing an add-in).
1. **GUI/runtime smoke:** open representative forms, inspect expected Pages and
   controls, and exercise representative interactions manually.

Layers 1 through 3 are the ordinary release gate. Layer 4 is required for a
release or when an MSForms/UI change warrants it; it complements rather than
replaces the automated gates. For the full automated path, run:

```powershell
.\scripts\frxedit-build.ps1

Push-Location .\tools\frx-edit
dotnet run --project .\tests\FrxEdit.Tests\FrxEdit.Tests.csproj -c Release
.\scripts\test-generated-container-pipeline.ps1 -Configuration Release
Pop-Location

.\scripts\userforms-verify.ps1

.\scripts\office-build.ps1 build -Force `
    -PpamOutputPath .\.build\office\IguanaTeX.ppam

.\scripts\office-build.ps1 validate `
    -InputPath .\.build\office\IguanaTeX.pptm

.\scripts\office-build.ps1 validate `
    -InputPath .\.build\office\IguanaTeX.ppam
```

The acceptance chain is pinned codec plus canonical JSON/VBA/assets, generated
native source, fresh PPTM, metadata/references, component/form import, real VBE
compile, VBA package closure, Ribbon injection, static package validation,
PowerPoint save/reopen, and optional PPAM conversion/add-in reload.

### Negative tests for compile or ownership changes

Use only disposable copies and verify all of the following:

- add a module with an intentionally undefined type and confirm a
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
1. Initialize the submodule, then run `git status` in both repositories and
   preserve unrelated tracked changes and untracked files.
1. Inspect history when a behavior or serialization choice is unclear; use
   `3bbfb96` and `7a2b8de` when investigating the source transition.
1. Route the task explicitly to canonical form data, the generic codec, sync,
   fresh build, package validation, runtime VBA, or a platform helper.

While editing:

- Treat manifest-listed JSON/VBA/assets, `src/*.bas`, `src/*.cls`, and other
  `office/` metadata as authoritative; never reverse-engineer source state from
  an arbitrary binary artifact when canonical input exists.
- Keep generated `.frm/.frx`, published codecs, PPTM, PPAM, unpacked packages,
  protocol files, screenshots, and diagnostic dumps under `.build/` or another
  disposable directory.
- Do not broaden a developer-infrastructure task into an IguanaTeX behavior,
  UI, callback-name, or Ribbon-layout refactor.
- Never edit or preserve an opaque FRX as canonical state. Export an explicitly
  selected native form through the strict transactional workflow when needed.
- Do not terminate PowerPoint by executable name. Exact PID and start-time
  verification are mandatory.
- Do not weaken validation merely to make a generated artifact pass.
- Prefer shared helpers over copying import or COM cleanup logic into another
  entry point.
- Fail safely if ownership, package state, metadata, or reference identity is
  ambiguous.

Before handing off:

1. Run validation proportional to the affected layer.
1. Check that no PowerPoint process or temporary Office directory owned by the
   test remains.
1. Review `git diff --check`, the IguanaTeX status, and the submodule status.
1. Confirm `.build/` and user-supplied local files are not staged.
1. Report implemented behavior, reused code, codec revision, commands actually
   run, criteria passed, manual checks, and remaining Office/VBE risks.

## Commit and review guidance

Keep a commit focused on one coherent change. A useful commit body explains the
canonical state affected, the safety/validation behavior, and the tests run.
Do not commit generated Office artifacts or temporary unpacked packages.

Reviewers should pay particular attention to:

- accidental use of a PPTM as fresh-build input;
- manifest, asset, component, and generated form closure;
- an uninitialized, floating, mismatched, or dirty codec submodule;
- direct edits to generated `.frm/.frx` or byte-level FRX assertions;
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
- Native UserForm text and FRX serialization can change across export/import
  cycles, so canonical verification uses form VBA and MSForms semantics rather
  than FRX byte identity.
- The PPTM is first injected only after its build session closes, then reopened
  and saved to prove Ribbon persistence. PPAM conversion reinjects after its
  final `SaveAs` because PowerPoint may remove or rewrite Ribbon parts.
- PPAM is not separately compiled after conversion because VBE does not safely
  target a specific loaded add-in project.
- Static OPC validity is necessary but not sufficient; real PowerPoint reopen or
  add-in load validation remains required.

These are design constraints, not reasons to bypass a gate.
