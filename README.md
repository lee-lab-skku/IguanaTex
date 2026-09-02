<!-- markdownlint-disable MD033 -->

# A Canonical IguanaTeX for VizVault

(C) Modified by Taeha Shin (This fork); [Jonathan Le Roux](https://www.jonathanleroux.org/) and Zvika Ben-Haim (Windows), [Tsung-Ju Chiang](https://github.com/tsung-ju) and Jonathan Le Roux (Mac)

IguanaTeX is a PowerPoint add-in which allows you to insert LaTeX equations into your PowerPoint presentation on Windows and Mac. It is distributed completely for free, along with its source code.

This fork stores canonical VBA source under `src/`, Office project and RibbonX metadata under `office/`, and developer automation under `scripts/`. Generated `.pptm` and `.ppam` files are disposable build artifacts and are not source files. See [CONTRIBUTING.md](CONTRIBUTING.md) for the repository model and development workflows.

This README describes the Docker-only fork. Upstream `.ppam` and `.pptm` artifacts do not contain these changes. Use an artifact explicitly built from this fork, or create one from the canonical source with the [fresh Office artifact workflow](CONTRIBUTING.md#fresh-office-artifact-workflow).

## Table of Contents

- [A Canonical IguanaTeX for VizVault](#a-canonical-iguanatex-for-vizvault)
  - [Table of Contents](#table-of-contents)
  - [System Requirements](#system-requirements)
    - [Shared requirements](#shared-requirements)
    - [Windows](#windows)
    - [Mac](#mac)
  - [Download and Install](#download-and-install)
    - [Add-in artifact](#add-in-artifact)
    - [Shared Docker setup](#shared-docker-setup)
    - [Windows Installation](#windows-installation)
    - [Mac Installation](#mac-installation)
    - [Optional LaTeXiT metadata](#optional-latexit-metadata)
    - [Other installation settings](#other-installation-settings)
  - [Development](#development)
  - [Tips, Bugs, and Known Issues](#tips-bugs-and-known-issues)
    - [What to do if something does not work, or does not work as you expected](#what-to-do-if-something-does-not-work-or-does-not-work-as-you-expected)
    - [Debugging an issue](#debugging-an-issue)
    - [Keyboard shortcuts](#keyboard-shortcuts)
    - [Known Issues](#known-issues)
  - [Project support and upstream news](#project-support-and-upstream-news)
  - [AI disclosure](#ai-disclosure)
  - [License](#license)

## System Requirements

### Shared requirements

- A desktop PowerPoint release that can load VBA `.ppam` add-ins and custom
  RibbonX. Shape output additionally requires PowerPoint support for inserting
  and converting SVG; Picture output remains available when SVG conversion is
  unavailable.
- Docker Desktop, or another local Docker engine reachable through the `docker`
  CLI. Every Generate/ReGenerate render and every PDF/DVI/XDV/PS/EPS-to-SVG
  conversion runs in a network-disabled container. The default image is
  `danteev/texlive:latest` and can be changed in Main Settings.
- The configured image must provide `sh`, `tar`, `timeout`, `awk`, the selected
  LaTeX engine, and the tools used by the selected output path: `latexmk`,
  `dvisvgm`, `dvipng`, `dvipdfmx`, Ghostscript including `ps2pdf`, ImageMagick's
  `magick`, `pdfcrop`, and `pdftocairo` as a PDF-to-SVG fallback when dvisvgm
  cannot process the installed Ghostscript version.
- Optionally, the platform-specific [LaTeXiT-metadata](#optional-latexit-metadata)
  helper can convert displays created with
  [LaTeXiT](https://www.chachatelier.fr/latexit/) into IguanaTeX displays.

### Windows

- OS: a Windows release supported by both the installed PowerPoint version and the chosen Docker engine.
- Picture output uses PNG.

### Mac

- Mac support in this fork is best-effort. Compatibility has been observed on
  some PowerPoint and macOS combinations, but no specific combination is
  guaranteed or part of an assured Mac regression baseline.
- An Intel or Apple Silicon Mac running a macOS release supported by both
  PowerPoint and the chosen Docker engine.
- Picture output can use PDF or PNG. Shape output requires SVG insertion and
  conversion support in the installed PowerPoint release.

## Download and Install

### Add-in artifact

Use a Docker-only `.ppam` built on Windows from this fork. Contributors can
create `.build\office\IguanaTeX.ppam` with the
[fresh Office artifact workflow](CONTRIBUTING.md#fresh-office-artifact-workflow).
The project does not yet provide a standalone Mac distribution workflow; a
Windows-built fork artifact can be used with the Mac integration files below for
best-effort testing.

An upstream `.ppam`, `.pptm`, or `iguanatexmac` Homebrew installation is not a
substitute: those artifacts use the upstream runtime rather than this fork's
Docker-only implementation.

### Shared Docker setup

1. Install and start Docker Desktop, or another local Docker engine with the
   `docker` CLI, and confirm that `docker version` can reach it.
1. Before first use, run `docker pull danteev/texlive:latest`, or pull another
   image and enter that image reference in the **Docker image** field in Main
   Settings.
1. Verify that `docker image inspect danteev/texlive:latest` succeeds,
   substituting the configured image when it differs. Runtime jobs use
   `--pull never`, so IguanaTeX never downloads an image implicitly.

IguanaTeX starts one `--rm -i --network none --pull never` container per
Generate/ReGenerate or vector-file conversion operation. Each job receives a
private workspace below the configured temporary root. TeX source, the generated
render job, and supported root-level TeX/graphics auxiliary inputs are staged
there and sent as a tar stream; only the final PDF, PNG, or SVG bytes are returned
on standard output.

**Keep Temp. files** preserves the host-side Docker job workspace, payload, log,
and returned artifact. Intermediate files created only inside the disposable
container are not copied back. The configured temporary root itself is never
used as the tar boundary. When it is the operating system's shared `%TEMP%`/`TMP`
root, pre-existing files are not treated as auxiliary inputs. To use custom
`.sty`, image, bibliography, font, or data files, select a dedicated Temp folder
and place those root-level inputs there.

Host LaTeX, Ghostscript, ImageMagick, TeX2img, pdfiumdraw, and dvisvgm
executables are not used for rendering. Main Settings exposes a Docker image
instead of host renderer paths.

### Windows Installation

1. **Authorize the individual add-in file** only after verifying that it came
   from a trusted build of this fork. In File Explorer, right-click the `.ppam`,
   select **Properties**, select **Unblock** on the **General** tab when that
   option is present, and then select **Apply** and **OK**. This removes the
   file's Mark of the Web as described in Microsoft's
   [guidance for trusted macro files](https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked#remove-mark-of-the-web-from-a-file).
   If organizational policy still blocks the add-in, contact the administrator;
   do not weaken global macro settings or broadly trust a folder.
1. **Load the add-in**: in **File > Options > Add-Ins**, choose **PowerPoint
   Add-Ins** from the **Manage** list. Select **Go...**, choose **Add New**,
   select the Docker-only `.ppam`, and close the dialog.
1. **Create and set a temporary file folder**: IguanaTeX needs access to a folder with read/write permissions to store temporary files.
   - The default is "C:\Temp\". If you have write permissions under "C:\", create the folder "C:\Temp\". You're all set.
   - If you cannot create this folder, choose or create a folder with write permission at any other location. In the IguanaTeX tab, choose "Main Settings" and put the path to the folder of your choice. You can also use a relative path under the presentation's folder (e.g., ".\" for the presentation folder itself).

### Mac Installation

1. **Install the Mac integration files**. The upstream project's [prebuilt Mac files](https://github.com/Jonathan-LeRoux/IguanaTex/releases) may be used for `IguanaTex.scpt` and `libIguanaTexHelper.dylib`; use the Docker-only `.ppam` described above instead of the bundled upstream add-in.

   There are three files to install:
   - `IguanaTex.scpt`: AppleScript file for handling file and folder access
   - `libIguanaTexHelper.dylib`: library for creating native text views; source code included in the git repo, under "IguanaTexHelper/"
   - the Docker-only `.ppam`: main add-in file
1. **Install `IguanaTex.scpt`**

    ```bash
    mkdir -p ~/Library/Application\ Scripts/com.microsoft.Powerpoint
    cp ./IguanaTex.scpt ~/Library/Application\ Scripts/com.microsoft.Powerpoint/IguanaTex.scpt
    ```

1. **Install `libIguanaTexHelper.dylib`**

    ```bash
    sudo mkdir -p '/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized'
    sudo cp ./libIguanaTexHelper.dylib '/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib'
    ```

1. **Load the add-in**: Start PowerPoint (restart if it was running when installing the dylib). From the menu bar, select Tools > PowerPoint Add-ins... > '+', and choose the Docker-only `.ppam`.
   - If macOS blocks `libIguanaTexHelper.dylib`, first confirm that the helper
     came from the linked upstream release and has not been replaced. Then open
     **System Settings > Privacy & Security**, scroll to **Security**, select
     **Open**, and then select **Open Anyway** and authenticate when prompted.
     Do not override the warning for an unverified file. See Apple's
     [security guidance](https://support.apple.com/guide/mac-help/mh40616/mac).

1. **Verify the runtime settings**:
   Click on "Main Settings" in the IguanaTeX ribbon tab:
   - Set the Temp folder used for file conversions in one of the following ways:
     - (Recommended) The simplest is to let IguanaTeX pick the Temp folder by selecting "Absolute" and leaving the path empty. The Temp folder will be inside the PowerPoint sandbox and everything will work without having to give permissions.
     - (If needed) If you need the Temp folder to be a specific folder outside the sandbox, or if you'd rather it be relative to your PowerPoint presentation (e.g., because you are using specific macros in an external file), you will need to give access permission to that folder to IguanaTeX. The best thing to do is to drag and drop that folder from the Finder on top of the PowerPoint application. This will allow you to give permission to the whole folder at least for the current session.

### Optional LaTeXiT metadata

Pierre Chatelier, LaTeXiT's author, prepared the metadata helpers at Jonathan Le
Roux's request. Their source is available for
[Windows](https://github.com/LaTeXiT-metadata/LaTeXiT-metadata-Win) and
[macOS](https://github.com/LaTeXiT-metadata/LaTeXiT-metadata-MacOS).

- On Windows, download
  [`LaTeXiT-metadata-Win.zip`](https://github.com/Jonathan-LeRoux/IguanaTex/releases/download/v1.60.3/LaTeXiT-metadata-Win.zip)
  from the upstream Releases page, unzip it, and set the path to
  `LaTeXiT-metadata.exe` in Main Settings.
- On macOS, download
  [`LaTeXiT-metadata-macos`](https://github.com/Jonathan-LeRoux/IguanaTex/releases/download/v1.60.3/LaTeXiT-metadata-macos)
  from the upstream Releases page, add executable permission, and either set its
  path in Main Settings or copy it to the secure add-in folder:

  ```bash
  chmod 755 ./LaTeXiT-metadata-macos
  sudo cp ./LaTeXiT-metadata-macos '/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/'
  ```

  If macOS blocks the helper after its first invocation, verify the downloaded
  file, then use the **System Settings > Privacy & Security** Open/Open Anyway
  procedure described above.

### Other installation settings

- The Main Settings field that previously displayed the host TeX executable path is now the **Docker image** field. It controls every generated PDF, PNG, and SVG, including ReGenerate and supported vector-file conversions. Retired host-renderer registry values are no longer read; the corresponding keys in old settings XML are ignored during import.
- If you select Tectonic, the configured Docker image must provide `tectonic` on its internal `PATH`.
- If you would like to have the option of using an external editor, e.g., when debugging LaTeX source code, you can specify the path to that editor in Main Settings. If you would like to use that editor by default over the IguanaTeX edit window, check the "use as default" checkbox.

## Development

The repository does not use a checked-in `.pptm` as source. VBA, project
metadata, references, and RibbonX are maintained in `src/` and `office/`.
UserForms are canonical JSON templates, matching `.vba` files, and referenced
assets; native `.frm/.frx` pairs are generated and never checked in.

Before building, make sure the `tools/frx-edit` submodule is initialized and at
the revision recorded by this repository.

The pinned FrxEdit submodule is published automatically and all nine UserForms
are staged under `.build/vba-source/` by the Office workflow. Windows developers
can create disposable Office artifacts with:

```powershell
.\scripts\office-build.ps1 build
.\scripts\office-build.ps1 validate -InputPath .\.build\office\IguanaTeX.pptm
```

The .NET 8 SDK is required to publish the codec. PowerPoint and Windows
PowerShell 5.1 are required for Office automation, and **Trust access to the VBA
project object model** must be enabled. The separate `vba-sync.ps1` workflow is
only for synchronizing an existing PPTM in `.build/office/` with canonical
source; UserForms are compared semantically and require explicit selection
before an export can update their JSON/VBA/assets.

Runtime rendering has one external-tool setting: `DockerImage`. Its default is
`DEFAULT_DOCKER_IMAGE` in `src/Defaults.bas` (`danteev/texlive:latest`). Host
renderer paths are not part of the runtime configuration, and no generated
Office artifact is the source of that setting.

Read [CONTRIBUTING.md](CONTRIBUTING.md) before changing canonical sources or
automation. It documents the fork-specific history, complete repository map,
sync and fresh-build CLI, Office/VBE safety boundaries, validation matrix,
canonical UserForm workflow, FrxEdit escalation rules, and instructions for both
human and AI contributors.

## Tips, Bugs, and Known Issues

### What to do if something does not work, or does not work as you expected

For Docker-only rendering, canonical-source, or fork-built artifact problems,
search or open an issue in this fork's
[issue tracker](https://github.com/lee-lab-skku/iguanatex-canonical-vizvault/issues).
The upstream [Frequently Asked Questions](https://www.jonathanleroux.org/software/iguanatex/faq.html)
and [issue tracker](https://github.com/Jonathan-LeRoux/IguanaTex/issues?q=is%3Aissue)
may help with behavior shared with upstream IguanaTeX, but upstream host-renderer
instructions do not apply to this Docker-only fork.

### Debugging an issue

When a display fails to generate, select **Debug** in the Editor before choosing
**Generate**. The current runtime shows the single host-side Docker command in
the Error/Debug window before executing it. Use **Copy** to copy that command for
inspection or manual execution in Terminal or Command Prompt. Debug mode also
preserves the private host-side Docker workspace, payload, log, and returned
artifact. It does not step through the individual LaTeX and conversion commands
inside the container; inspect the preserved generated job and log to identify
the failing stage.

If this does not solve the issue, or the issue does not occur during the generation process, the next step is to try to debug in the VBA Editor. To do so:

- build a fresh disposable Docker-only `.pptm` from this checkout by following
  the [fresh Office artifact workflow](CONTRIBUTING.md#fresh-office-artifact-workflow).
- open the VBA Editor (`Alt+F11` on Windows, `Tools > Macro > Visual Basic Editor` on Mac).
- search for "Macros" under "Module" in the exploration pane on the left.
- place a breakpoint, for example at Line 7 (`Load LatexForm` under `NewLatexEquation()`) by clicking in the margin.
- Launch the display generation process:
  - on Windows, click on the "New LaTeX Display" button in the IguanaTeX ribbon (if the add-in is loaded, you will likely have two IguanaTeX tabs in the ribbon, one for the loaded add-in and the other for the `.pptm` file: just try one, and if the IguanaTeX window appears without hitting the breakpoint, try the other), or in the VBA Editor click `Tools >  Macros...` and select `NewLatexEquation` and `Run`.
  - on Mac, clicking on the buttons in the ribbon does not work for `.pptm` files, so instead click `Tools > Macro > Macros...` and select `NewLatexEquation` and `Run`.
- The code will stop at the breakpoint.
- Step Over (Shift+F8 on Windows, Shift+⌘+O on Mac) until hitting the bug. If the bug occurs on a line calling another function, you can run again and then Step Into (F8 on Windows, Shift+⌘+I on Mac) when you reach that line.
- Eventually, you'll reach the actual line causing the bug. Either fix it or
  report it in the fork issue tracker linked above.

### Keyboard shortcuts

Accelerator keys (i.e., keyboard shortcuts): many of IguanaTeX's commands ("Generate", "Cancel", etc) can be accessed by using a combination of modifier keys and a single letter. Look for the underlined letter in the corresponding button's text/label.

- Windows: Alt + letter. For example, instead of clicking on the "<ins>G</ins>enerate" button, you can use `Alt + g`. (This is the standard Office behavior on Windows)
- Mac: Ctrl + Cmd + letter. For example, instead of clicking on the "<ins>G</ins>enerate" button, you can use `Ctrl + ⌘ + g`. (Accelerator keys are not available in the standard Office for Mac, this was specially coded by Tsung-Ju for IguanaTex)

### Known Issues

- "Picture" displays created on Mac (which are inserted PDFs) appear cropped on Windows ([Issue #32](https://github.com/Jonathan-LeRoux/IguanaTex/issues/32)). Regenerating them on Windows fixes the issue. This seems to be a bug with the way PowerPoint handles some PDFs on Mac, internally storing them as EMF files. The PDFs created by LaTeXiT do not have that issue, however, so there may be a way to circumvent this bug in a future version of IguanaTeX.
- There may be some scaling issues when changing a display between Picture and Shape or between the two SVG generation modes. The best way to handle this is to use the "Convert to Shape"/"Convert to Picture" functions, which regenerate the display in the desired format while keeping the size fixed. One can then further modify the content if needed, and the scaling will be correct.
- For Shape (i.e., vector graphics) displays, the default "SVG via DVI in Docker" is recommended because PDF-derived SVG can have symbols or parts of symbols missing. Certain lines are represented in PDF by open paths with a line width instead of closed paths and are handled differently when PowerPoint converts the SVG to a Shape. See [this discussion](https://github.com/mgieseki/dvisvgm/issues/166) for more details.

## Project support and upstream news

Use this fork's [issue tracker](https://github.com/lee-lab-skku/iguanatex-canonical-vizvault/issues)
for fork development and Docker-only behavior. The
[IguanaTex Google Group](https://groups.google.com/d/forum/iguanatex) announces
upstream project news and releases; it is not a release channel for this fork or
its best-effort Mac testing artifacts.

## AI disclosure

This fork uses AI-assisted coding and documentation tools. AI-assisted changes
are reviewed and validated through the repository's development workflows before
acceptance.

## License

[![CC BY 3.0][cc-by-image]][cc-by]

This work is licensed under a
[Creative Commons Attribution 3.0 Unported License][cc-by].

[cc-by]: http://creativecommons.org/licenses/by/3.0/
[cc-by-image]: https://i.creativecommons.org/l/by/3.0/88x31.png
