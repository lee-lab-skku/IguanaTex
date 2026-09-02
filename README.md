# IguanaTex

(C) [Jonathan Le Roux](https://www.jonathanleroux.org/) and Zvika Ben-Haim (Windows), [Tsung-Ju Chiang](https://github.com/tsung-ju) and Jonathan Le Roux (Mac)

IguanaTex is a PowerPoint add-in which allows you to insert LaTeX equations into your PowerPoint presentation on Windows and Mac. It is distributed completely for free, along with its source code.

This fork stores canonical VBA source under `src/`, Office project and RibbonX metadata under `office/`, and developer automation under `scripts/`. Generated `.pptm` and `.ppam` files are disposable build artifacts and are not source files. See [CONTRIBUTING.md](CONTRIBUTING.md) for the repository model and development workflows.

This README describes the Docker-only fork. Upstream `.ppam` and `.pptm` artifacts do not contain these changes. Use an artifact explicitly built from this fork, or create one from the canonical source with the [fresh Office artifact workflow](CONTRIBUTING.md#fresh-office-artifact-workflow).

## Table of Contents

- [System Requirements](#system-requirements)
  - [Windows](#windows)
  - [Mac](#mac)
- [Download and Install](#download-and-install)
  - [Windows Installation](#windows-installation)
  - [Mac Installation](#mac-installation)
  - [Other installation settings](#other-installation-settings)
- [Development](#development)
- [Tips, Bugs, and Known Issues](#tips-bugs-and-known-issues)
  - [What to do if something does not work, or does not work as you expected](#what-to-do-if-something-does-not-work-or-does-not-work-as-you-expected)
  - [Debugging an issue](#debugging-an-issue)
  - [Keyboard shortcuts](#keyboard-shortcuts)
  - [Known Issues](#known-issues)
- [Stay up to date: IguanaTex Google Group](#stay-up-to-date-iguanatex-google-group)
- [License](#license)

## System Requirements

### Windows

- OS: a Windows release supported by both the installed PowerPoint version and the chosen Docker engine.
- PowerPoint:
  - The current Windows regression baseline is Microsoft 365 PowerPoint. Picture output uses PNG; Shape output requires a PowerPoint release that can insert and convert SVG.
  - Older PowerPoint releases from the upstream compatibility history are not part of the Docker-only regression baseline.
- Docker Desktop (or another local Docker engine with the `docker` CLI). Every Generate/ReGenerate render and every PDF/DVI/XDV/PS/EPS-to-SVG conversion runs in a network-disabled container. The default image is `danteev/texlive:latest` and can be changed in Main Settings.
- The container image must provide `sh`, `tar`, `timeout`, `awk`, the selected LaTeX engine, and the rendering tools used by the selected output (`latexmk`, `dvisvgm`, `dvipng`, `dvipdfmx`, Ghostscript including `ps2pdf`, ImageMagick's `magick`, `pdfcrop`, and `pdftocairo` as a PDF-to-SVG fallback when dvisvgm cannot process the installed Ghostscript version).
- (Optional) [LaTeXiT-metadata](https://github.com/LaTeXiT-metadata/LaTeXiT-metadata-Win), used to convert displays generated with [LaTeXiT](https://www.chachatelier.fr/latexit/) on Mac into IguanaTex displays

### Mac

- An Intel or Apple Silicon Mac running a macOS release supported by both PowerPoint and the chosen Docker engine.
- PowerPoint for Mac:
  - Office 365, Office 2021 (including LTSC version), PowerPoint 2019, PowerPoint 2016 (Version 16.16.7 190210 or later)
  - SVG support is available for Office 365 and recent retail versions of PowerPoint, including 2019 and 2021. Volume-licensed releases without SVG conversion support can use Picture output but not Shape output.
- Docker Desktop (or another local Docker engine with the `docker` CLI). Every generated PDF, PNG, or SVG and every PDF/DVI/XDV/PS/EPS-to-SVG conversion uses the image configured in Main Settings (`danteev/texlive:latest` by default). A host MacTeX, Ghostscript, or ImageMagick installation is not used for rendering.
- (Optional) [LaTeXiT-metadata](https://github.com/LaTeXiT-metadata/LaTeXiT-metadata-MacOS), used to convert [LaTeXiT](https://www.chachatelier.fr/latexit/) displays into IguanaTex displays


## Download and Install

### Windows Installation

1. **Obtain a Docker-only .ppam add-in** explicitly built from this fork. Contributors can produce `.build\office\IguanaTeX.ppam` by following the [fresh Office artifact workflow](CONTRIBUTING.md#fresh-office-artifact-workflow). Do not substitute an upstream release artifact; it uses the upstream runtime rather than this Docker-only implementation.
1. **Place the .ppam in a Trusted Location**, such as `%appdata%\Microsoft\Addins` (for example, `C:\Users\user_name\Appdata\Roaming\Microsoft\Addins`). See Microsoft's guidance for [Trusted Locations](https://learn.microsoft.com/en-us/DeployOffice/security/trusted-locations) and [trusted macro files](https://learn.microsoft.com/en-us/DeployOffice/security/internet-macros-blocked#guidance-on-allowing-vba-macros-in-files-you-trust).
1. **Load the add-in**: in "File" > "Options" > "Add-Ins" > "Manage:" (lower part of the window), choose "PowerPoint Add-Ins" in the selection box. Then press "Go...", click "Add New", select the Docker-only `.ppam`, and click "Close".
1. **Create and set a temporary file folder**: IguanaTex needs access to a folder with read/write permissions to store temporary files.
   - The default is "C:\Temp\". If you have write permissions under "C:\", create the folder "C:\Temp\". You're all set.
   - If you cannot create this folder, choose or create a folder with write permission at any other location. In the IguanaTex tab, choose "Main Settings" and put the path to the folder of your choice. You can also use a relative path under the presentation's folder (e.g., ".\" for the presentation folder itself).
1. **Install Docker and make the rendering image available locally**:
   - Ensure `docker version` can reach the local engine. Before first use, run `docker pull danteev/texlive:latest`, or pull another image and enter that image reference in the **Docker image** field in Main Settings.
   - Verify that `docker image inspect danteev/texlive:latest` succeeds, substituting the configured image when it differs. Runtime jobs use `--pull never`, so IguanaTex never downloads an image implicitly.
   - IguanaTex starts one `--rm -i --network none --pull never` container per Generate/ReGenerate or vector-file conversion operation. Each job gets a private workspace below the configured temporary root. TeX source, the generated render job, and supported root-level TeX/graphics auxiliary inputs are staged there and sent as a tar stream; only the final PDF, PNG, or SVG bytes are returned on standard output.
   - **Keep Temp. files** preserves the host-side Docker job workspace, payload, log, and returned artifact. Intermediate files created only inside the disposable container are not copied back.
   - The configured temporary root itself is never used as the tar boundary. When the configured folder is the operating system's shared `%TEMP%`/`TMP` root, pre-existing files are not treated as auxiliary inputs at all. To use custom `.sty`, image, bibliography, font, or data files, select a dedicated Temp folder and place those root-level inputs there.
   - Host LaTeX, Ghostscript, ImageMagick, TeX2img, and pdfiumdraw executables are not used. Main Settings exposes a Docker image instead of host renderer paths.
1. (Optional) **Install LaTeXiT-metadata**:
   - Needed to convert displays generated with [LaTeXiT](https://www.chachatelier.fr/latexit/) on Mac into IguanaTex displays
   - Download [`LaTeXiT-metadata-Win.zip`](https://github.com/Jonathan-LeRoux/IguanaTex/releases/download/v1.60.3/LaTeXiT-metadata-Win.zip) from the upstream Releases page, unzip, and set the path to `LaTeXiT-metadata.exe` in the "Main Settings" window.
   - LaTeXiT-metadata was kindly prepared by Pierre Chatelier, [LaTeXiT](https://www.chachatelier.fr/latexit/)'s author, at my request. Many thanks to him!
   - [Source code is now public](https://github.com/LaTeXiT-metadata/LaTeXiT-metadata-Win).

### Mac Installation

1. **Obtain a Docker-only .ppam add-in** explicitly built from this fork. The upstream release and `iguanatexmac` Homebrew cask install the upstream add-in and are not substitutes for this fork's `.ppam`.
1. **Obtain the Mac integration files**. The upstream project's [prebuilt Mac files](https://github.com/Jonathan-LeRoux/IguanaTex/releases) may be used for `IguanaTex.scpt` and `libIguanaTexHelper.dylib`; use the Docker-only `.ppam` from the preceding step instead of the bundled upstream add-in.

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
   - The first time you click on one of the add-in buttons, you may be notified that `libIguanaTexHelper.dylib` was blocked. Go to the Mac's Settings, then Security and Privacy, and click "Allow Anyway".

1. **Verify the runtime settings**:
   Click on "Main Settings" in the IguanaTex ribbon tab:
   - Install and start Docker Desktop (or another local Docker engine with the `docker` CLI), then verify that `docker version` can reach it.
   - Set the Temp folder used for file conversions in one of the following ways:
     - (Recommended) The simplest is to let IguanaTex pick the Temp folder by selecting "Absolute" and leaving the path empty. The Temp folder will be inside the PowerPoint sandbox and everything will work without having to give permissions.
     - (If needed) If you need the Temp folder to be a specific folder outside the sandbox, or if you'd rather it be relative to your PowerPoint presentation (e.g., because you are using specific macros in an external file), you will need to give access permission to that folder to IguanaTex. The best thing to do is to drag and drop that folder from the Finder on top of the PowerPoint application. This will allow you to give permission to the whole folder at least for the current session.
   - Before first use, run `docker pull danteev/texlive:latest`, or pull another image and set the **Docker image** field to that reference. The default is `danteev/texlive:latest`.
   - Verify `docker image inspect danteev/texlive:latest`, substituting the configured image if it differs. Runtime jobs use `--pull never`; IguanaTex does not download images implicitly or use host LaTeX, Ghostscript, ImageMagick, or libgs paths.

1. (Optional) **Install LaTeXiT-metadata**:
   - Needed to convert displays generated with [LaTeXiT](https://www.chachatelier.fr/latexit/) on Mac into IguanaTex displays
   - Download [`LaTeXiT-metadata-macos`](https://github.com/Jonathan-LeRoux/IguanaTex/releases/download/v1.60.3/LaTeXiT-metadata-macos) from the upstream Releases page, add executable permission, and either set the path to its location in the "Main Settings" window or copy it to the secure add-in folder:

     ```bash
     chmod 755 ./LaTeXiT-metadata-macos
     sudo cp ./LaTeXiT-metadata-macos '/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/'
     ```
   - The first time LaTeXiT-metadata-macos is called by IguanaTex, Mac OS may block it. Go to the Mac's Settings, then Security and Privacy, and click "Allow Anyway".
   - The executable was compiled on Mac OS 10.13 but should work on all versions. Please let me know if you have any issue.
   - LaTeXiT-metadata was kindly prepared by Pierre Chatelier, [LaTeXiT](https://www.chachatelier.fr/latexit/)'s author, at my request. Many thanks to him!
   - [Source code is now public](https://github.com/LaTeXiT-metadata/LaTeXiT-metadata-MacOS).

### Other installation settings

- The Main Settings field that previously displayed the host TeX executable path is now the **Docker image** field. It controls every generated PDF, PNG, and SVG, including ReGenerate and supported vector-file conversions. Retired host-renderer registry values are no longer read; the corresponding keys in old settings XML are ignored during import.
- If you select Tectonic, the configured Docker image must provide `tectonic` on its internal `PATH`.
- If you would like to have the option of using an external editor, e.g., when debugging LaTeX source code, you can specify the path to that editor in Main Settings. If you would like to use that editor by default over the IguanaTex edit window, check the "use as default" checkbox.

## Development

The repository does not use a checked-in `.pptm` as source. VBA, project
metadata, references, and RibbonX are maintained in `src/` and `office/`.
UserForms are canonical JSON templates, matching `.vba` files, and referenced
assets; native `.frm/.frx` pairs are generated and never checked in.

Clone with `--recurse-submodules`, or initialize an existing checkout before
building:

```powershell
git submodule update --init --recursive
```

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

Most issues originate from some steps of the installation process described above not being followed: please double-check you went through all the steps. A reboot also often helps after a first installation.

If you are having trouble installing or using IguanaTex, please see the [Frequently Asked Questions](https://www.jonathanleroux.org/software/iguanatex/faq.html) and check the [upstream issue tracker](https://github.com/Jonathan-LeRoux/IguanaTex/issues?q=is%3Aissue).

### Debugging an issue

When running into an issue while trying to generate a display, the first thing to do is to check the "Debug" box in the Editor window prior to clicking "Generate". This will step through the process of generating the display, so that we can know where the error occurred, and it will give the option to copy each command so that they can be run in a Terminal or Command Prompt.

If this does not solve the issue, or the issue does not occur during the generation process, the next step is to try to debug in the VBA Editor. To do so:

- build a fresh disposable Docker-only `.pptm` from this checkout by following
  the [fresh Office artifact workflow](CONTRIBUTING.md#fresh-office-artifact-workflow).
- open the VBA Editor (`Alt+F11` on Windows, `Tools > Macro > Visual Basic Editor` on Mac).
- search for "Macros" under "Module" in the exploration pane on the left.
- place a breakpoint, for example at Line 7 (`Load LatexForm` under `NewLatexEquation()`) by clicking in the margin.
- Launch the display generation process:
  - on Windows, click on the "New LaTeX Display" button in the IguanaTex ribbon (if the add-in is loaded, you will likely have two IguanaTex tabs in the ribbon, one for the loaded add-in and the other for the `.pptm` file: just try one, and if the IguanaTex window appears without hitting the breakpoint, try the other), or in the VBA Editor click `Tools >  Macros...` and select `NewLatexEquation` and `Run`.
  - on Mac, clicking on the buttons in the ribbon does not work for `.pptm` files, so instead click `Tools > Macro > Macros...` and select `NewLatexEquation` and `Run`.
- The code will stop at the breakpoint.
- Step Over (Shift+F8 on Windows, Shift+⌘+O on Mac) until hitting the bug. If the bug occurs on a line calling another function, you can run again and then Step Into (F8 on Windows, Shift+⌘+I on Mac) when you reach that line.
- Eventually, you'll reach the actual line causing the bug. Now, either try to fix it, or open an issue.

### Keyboard shortcuts

Accelerator keys (i.e., keyboard shortcuts): many of IguanaTex's commands ("Generate", "Cancel", etc) can be accessed by using a combination of modifier keys and a single letter. Look for the underlined letter in the corresponding button's text/label.

- Windows: Alt + letter. For example, instead of clicking on the "<ins>G</ins>enerate" button, you can use `Alt + g`. (This is the standard Office behavior on Windows)
- Mac: Ctrl + Cmd + letter. For example, instead of clicking on the "<ins>G</ins>enerate" button, you can use `Ctrl + ⌘ + g`. (Accelerator keys are not available in the standard Office for Mac, this was specially coded by Tsung-Ju for IguanaTex)

### Known Issues

- "Picture" displays created on Mac (which are inserted PDFs) appear cropped on Windows ([Issue #32](https://github.com/Jonathan-LeRoux/IguanaTex/issues/32)). Regenerating them on Windows fixes the issue. This seems to be a bug with the way PowerPoint handles some PDFs on Mac, internally storing them as EMF files. The PDFs created by LaTeXiT do not have that issue, however, so there may be a way to circumvent this bug in a future version of IguanaTex.
- IguanaTex macros cannot be added to the Quick Access Toolbar on Mac ([Issue #23](https://github.com/Jonathan-LeRoux/IguanaTex/issues/23)): this is a [known bug](https://answers.microsoft.com/en-us/msoffice/forum/all/can-add-in-commands-be-added-to-the-quick-access/6872187f-3c17-40ee-8620-80a4068edc82) on which Microsoft is allegedly working, although there has been no progress for multiple years.
- There may be some scaling issues when changing a display between Picture and Shape or between the two SVG generation modes. The best way to handle this is to use the "Convert to Shape"/"Convert to Picture" functions, which regenerate the display in the desired format while keeping the size fixed. One can then further modify the content if needed, and the scaling will be correct.
- For Shape (i.e., vector graphics) displays, the default "SVG via DVI in Docker" is recommended because PDF-derived SVG can have symbols or parts of symbols missing. Certain lines are represented in PDF by open paths with a line width instead of closed paths and are handled differently when PowerPoint converts the SVG to a Shape. See [this discussion](https://github.com/mgieseki/dvisvgm/issues/166) for more details.

## Stay up to date: IguanaTex Google Group

To be informed of the release of new versions, you can subscribe to the [IguanaTex Google Group](https://groups.google.com/d/forum/iguanatex).

## License

[![CC BY 3.0][cc-by-image]][cc-by]

This work is licensed under a
[Creative Commons Attribution 3.0 Unported License][cc-by].

[cc-by]: http://creativecommons.org/licenses/by/3.0/
[cc-by-image]: https://i.creativecommons.org/l/by/3.0/88x31.png
[cc-by-shield]: https://img.shields.io/badge/License-CC%20BY%203.0-lightgrey.svg
