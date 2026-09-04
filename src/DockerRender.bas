Attribute VB_Name = "DockerRender"
Option Explicit

Private Const TAR_BLOCK_SIZE As Long = 512
Private Const DOCKER_JOB_FILENAME As String = "job.sh"
Private Const DOCKER_PAYLOAD_SUFFIX As String = "_docker_payload.tar"
Private Const DOCKER_ARTIFACT_SUFFIX As String = "_docker_artifact.tmp"
Private Const DOCKER_ENUM_ATTRIBUTES As Long = vbNormal Or vbReadOnly Or vbHidden

Public Type DockerRenderRequest
    FilePrefix As String
    LatexCommand As String
    LatexDviOptions As String
    LatexmkPdfOptions As String
    LatexmkDviOptions As String
    ShellEscapeBatchOptions As String
    UseLatexmk As Boolean
    UseDvi As Boolean
    UsePdf As Boolean
    UseVector As Boolean
    PictureOutputType As String
    OutputDpi As Long
    TimeoutSeconds As Long
End Type

Public Function ExecuteDockerRenderJob(ByRef Request As DockerRenderRequest, _
                                       ByVal TempPath As String, _
                                       ByVal debugMode As Boolean, _
                                       ByVal WaitTime As Long, _
                                       ByVal PreserveWorkspace As Boolean, _
                                       ByRef FinalFilename As String, _
                                       ByRef OutputType As String, _
                                       ByRef RunCommand As String, _
                                       ByRef FailureStage As String) As Long
    On Error GoTo BuildError
    Dim JobText As String
    Dim StageCount As Long
    JobText = BuildDockerJob(Request, FinalFilename, OutputType, StageCount)

    ExecuteDockerRenderJob = ExecuteDockerPayloadJob(JobText, StageCount, _
        TempPath, TempPath & Request.FilePrefix & ".tex", _
        Request.FilePrefix & ".tex", TempPath, Request.FilePrefix, _
        Request.FilePrefix, debugMode, WaitTime, PreserveWorkspace, _
        FinalFilename, RunCommand, FailureStage)
    Exit Function

BuildError:
    ExecuteDockerRenderJob = RecordDockerBuildError(TempPath, Request.FilePrefix, _
        "build Docker render job", Err.Number, Err.Source, Err.Description, FailureStage)
End Function

Public Function ExecuteDockerVectorConversionJob(ByVal InputPath As String, _
                                                  ByVal TempPath As String, _
                                                  ByVal debugMode As Boolean, _
                                                  ByVal WaitTime As Long, _
                                                  ByVal PreserveWorkspace As Boolean, _
                                                  ByRef FinalFilename As String, _
                                                  ByRef RunCommand As String, _
                                                  ByRef FailureStage As String) As Long
    On Error GoTo BuildError

    Dim InputExtension As String
    Dim InputFilename As String
    Dim JobText As String
    Dim StageCount As Long
    InputExtension = GetExtension(InputPath)
    InputFilename = "vector-input." & InputExtension
    FinalFilename = DefaultFilePrefix & "_vector.svg"
    JobText = BuildDockerVectorConversionJob(InputFilename, InputExtension, _
        FinalFilename, CLng(val(GetITSetting("TimeOutTime", "20"))), StageCount)

    ExecuteDockerVectorConversionJob = ExecuteDockerPayloadJob(JobText, StageCount, _
        TempPath, InputPath, InputFilename, vbNullString, vbNullString, _
        DefaultFilePrefix & "_vector", debugMode, _
        WaitTime, PreserveWorkspace, FinalFilename, RunCommand, FailureStage)
    Exit Function

BuildError:
    ExecuteDockerVectorConversionJob = RecordDockerBuildError(TempPath, _
        DefaultFilePrefix & "_vector", "build Docker vector conversion job", _
        Err.Number, Err.Source, Err.Description, FailureStage)
End Function

Private Function ExecuteDockerPayloadJob(ByVal JobText As String, _
                                         ByVal StageCount As Long, _
                                         ByVal TempPath As String, _
                                         ByVal SourcePath As String, _
                                         ByVal SourceArchiveName As String, _
                                         ByVal AuxiliarySourcePath As String, _
                                         ByVal AuxiliaryFilePrefix As String, _
                                         ByVal JobFilePrefix As String, _
                                         ByVal debugMode As Boolean, _
                                         ByVal WaitTime As Long, _
                                         ByVal PreserveWorkspace As Boolean, _
                                         ByVal FinalFilename As String, _
                                         ByRef RunCommand As String, _
                                         ByRef FailureStage As String) As Long
    On Error GoTo HostError

    RunCommand = vbNullString
    FailureStage = vbNullString

    Dim HostOperation As String
    Dim HostPath As String
    Dim DockerWasExecuted As Boolean
    Dim RootLogPath As String
    RootLogPath = TempPath & JobFilePrefix & ".log"

    Dim WorkspacePath As String
    Dim PayloadPath As String
    Dim ArtifactPath As String
    Dim FinalPath As String
    Dim WorkspaceLogPath As String

    HostOperation = "create private render workspace"
    HostPath = TempPath
    WorkspacePath = CreateDockerWorkspace(TempPath)
    PayloadPath = WorkspacePath & JobFilePrefix & DOCKER_PAYLOAD_SUFFIX
    ArtifactPath = WorkspacePath & JobFilePrefix & DOCKER_ARTIFACT_SUFFIX
    FinalPath = TempPath & FinalFilename
    WorkspaceLogPath = WorkspacePath & JobFilePrefix & ".log"

    DeleteDockerFile FinalPath

    HostOperation = "write generated Docker job"
    HostPath = WorkspacePath & DOCKER_JOB_FILENAME
    WriteToFile WorkspacePath, "job", ".sh", JobText

    HostOperation = "copy Docker job input"
    HostPath = SourcePath
    FileCopy HostPath, WorkspacePath & SourceArchiveName

    If AuxiliarySourcePath <> vbNullString Then
        StageDockerAuxiliaryFiles AuxiliarySourcePath, WorkspacePath, _
            AuxiliaryFilePrefix, HostOperation, HostPath
    End If

    WriteDockerPayload PayloadPath, WorkspacePath, HostOperation, HostPath

    Dim DockerCommand As String
    DockerCommand = BuildDockerCommand(GetDockerImageReference())
    RunCommand = DockerCommand & " < " & ShellEscape(PayloadPath) & _
        " > " & ShellEscape(ArtifactPath) & " 2> " & ShellEscape(WorkspaceLogPath)

    If debugMode Then
        ShowError vbNullString, RunCommand, "Debug mode", "Docker render job:", "Continue"
    End If

    Dim OverallWaitTime As Long
    OverallWaitTime = DockerOverallWaitTime(WaitTime, StageCount)
    HostOperation = "execute Docker render command"
    HostPath = WorkspacePath
    DockerWasExecuted = True
    ExecuteDockerPayloadJob = ExecuteRedirected(DockerCommand, WorkspacePath, _
        PayloadPath, ArtifactPath, WorkspaceLogPath, False, OverallWaitTime)

    FailureStage = ReadDockerFailureStage(WorkspaceLogPath)
    PublishDockerLog WorkspaceLogPath, RootLogPath
    If ExecuteDockerPayloadJob <> 0 Then Exit Function

    HostOperation = "validate Docker output artifact"
    HostPath = ArtifactPath
    If Not FileExists(ArtifactPath) Then
        FailureStage = "artifact"
        ExecuteDockerPayloadJob = 1
        Exit Function
    End If
    If FileLen(ArtifactPath) = 0 Then
        FailureStage = "artifact"
        ExecuteDockerPayloadJob = 1
        Exit Function
    End If

    HostOperation = "publish Docker output artifact"
    HostPath = ArtifactPath
    Name ArtifactPath As FinalPath
    If Not PreserveWorkspace Then DeleteDockerWorkspace WorkspacePath
    Exit Function

HostError:
    Dim HostErrorNumber As Long
    Dim HostErrorSource As String
    Dim HostErrorDescription As String
    Dim HostLogText As String
    HostErrorNumber = Err.Number
    HostErrorSource = Err.Source
    HostErrorDescription = Err.Description

    FailureStage = "host-payload"
    ExecuteDockerPayloadJob = 1
    HostLogText = "IguanaTex could not prepare or complete the host side of the Docker render job." & vbLf & _
        "Failure stage: host-payload" & vbLf & _
        "Host operation: " & HostOperation & vbLf
    If HostPath <> vbNullString Then
        HostLogText = HostLogText & "File/path: " & HostPath & vbLf
    End If
    HostLogText = HostLogText & _
        "VBA error number: " & CStr(HostErrorNumber) & vbLf & _
        "VBA error source: " & HostErrorSource & vbLf & _
        "VBA error description: " & HostErrorDescription & vbLf
    If DockerWasExecuted Then
        HostLogText = HostLogText & "Docker command execution had already started."
    Else
        HostLogText = HostLogText & "Docker command was not executed."
    End If

    On Error Resume Next
    WriteToFile TempPath, JobFilePrefix, ".log", HostLogText
    On Error GoTo 0
End Function

Private Function RecordDockerBuildError(ByVal TempPath As String, _
                                        ByVal JobFilePrefix As String, _
                                        ByVal HostOperation As String, _
                                        ByVal ErrorNumber As Long, _
                                        ByVal ErrorSource As String, _
                                        ByVal ErrorDescription As String, _
                                        ByRef FailureStage As String) As Long
    FailureStage = "host-payload"
    RecordDockerBuildError = 1
    On Error Resume Next
    WriteToFile TempPath, JobFilePrefix, ".log", _
        "IguanaTex could not prepare the host side of the Docker job." & vbLf & _
        "Failure stage: host-payload" & vbLf & _
        "Host operation: " & HostOperation & vbLf & _
        "VBA error number: " & CStr(ErrorNumber) & vbLf & _
        "VBA error source: " & ErrorSource & vbLf & _
        "VBA error description: " & ErrorDescription & vbLf & _
        "Docker command was not executed."
    On Error GoTo 0
End Function

Public Function GetDockerImageReference() As String
    Dim ImageReference As String
    ImageReference = Trim$(CStr(GetITSetting("DockerImage", DEFAULT_DOCKER_IMAGE)))
    If ImageReference = vbNullString Then ImageReference = DEFAULT_DOCKER_IMAGE
    GetDockerImageReference = ImageReference
End Function

Public Function DockerRenderErrorMessage(ByVal FailureStage As String, _
                                         ByVal TimeOutTimeString As String) As String
    Select Case FailureStage
        Case "latex"
            DockerRenderErrorMessage = "LaTeX failed inside the Docker render job. " & _
                "Please make sure your code compiles with the selected engine."
        Case "dvipdfmx"
            DockerRenderErrorMessage = "Error while using dvipdfmx inside the Docker render job."
        Case "dvisvgm"
            DockerRenderErrorMessage = "Error while using dvisvgm to create SVG inside the Docker render job."
        Case "pdf-to-svg"
            DockerRenderErrorMessage = "Error while converting PDF to SVG inside the Docker render job. " & _
                "The image must provide dvisvgm or the pdftocairo fallback."
        Case "pdfcrop"
            DockerRenderErrorMessage = "Error while cropping PDF inside the Docker render job."
        Case "ps2pdf"
            DockerRenderErrorMessage = "Error while converting PostScript to PDF inside the Docker render job."
        Case "bbox"
            DockerRenderErrorMessage = "Error while using Ghostscript to compute the bounding box inside the Docker render job."
        Case "ghostscript-png"
            DockerRenderErrorMessage = "Error while using Ghostscript to convert from PDF to PNG inside the Docker render job."
        Case "imagemagick"
            DockerRenderErrorMessage = "Error while using ImageMagick to set the PNG DPI inside the Docker render job."
        Case "dvipng"
            DockerRenderErrorMessage = "dvipng failed inside the Docker render job, or exceeded the " & _
                TimeOutTimeString & " second stage timeout."
        Case "artifact"
            DockerRenderErrorMessage = "The Docker render job completed without returning a final artifact."
        Case "host-payload"
            DockerRenderErrorMessage = "IguanaTex could not prepare or finish the host side of the Docker render job. " & _
                "The temporary log records whether Docker started, plus the host operation, path, and VBA error."
        Case Else
            DockerRenderErrorMessage = "The Docker render job failed before producing an artifact. " & _
                "Make sure Docker is running and image " & GetDockerImageReference() & " is available locally."
    End Select
End Function

Private Function BuildDockerJob(ByRef Request As DockerRenderRequest, _
                                ByRef FinalFilename As String, _
                                ByRef OutputType As String, _
                                ByRef StageCount As Long) As String
    Dim LF As String
    Dim JobText As String
    Dim InputFilename As String
    Dim CurrentType As String
    Dim CurrentExtension As String
    LF = vbLf
    InputFilename = Request.FilePrefix & ".tex"

    JobText = BuildDockerJobPreamble(Request.TimeoutSeconds, _
        Request.FilePrefix & ".log")

    If Request.UseDvi Then
        If Request.LatexCommand = "xelatex" Or Request.LatexCommand = "tectonic" Then
            CurrentType = "XDV"
            CurrentExtension = ".xdv"
        Else
            CurrentType = "DVI"
            CurrentExtension = ".dvi"
        End If

        If Request.UseLatexmk Then
            AppendDockerStage JobText, StageCount, "latex", "latexmk " & _
                DockerOptionsForShell(Request.LatexmkDviOptions) & _
                Request.ShellEscapeBatchOptions & DockerShellQuote(InputFilename)
        Else
            AppendDockerStage JobText, StageCount, "latex", Request.LatexCommand & " " & _
                DockerOptionsForShell(Request.LatexDviOptions) & _
                Request.ShellEscapeBatchOptions & DockerShellQuote(InputFilename)
        End If
        JobText = JobText & "require_file latex " & _
            DockerShellQuote(Request.FilePrefix & CurrentExtension) & LF

        If Request.UsePdf Then
            AppendDockerStage JobText, StageCount, "dvipdfmx", _
                "dvipdfmx -o " & DockerShellQuote(Request.FilePrefix & ".pdf") & " " & _
                DockerShellQuote(Request.FilePrefix & CurrentExtension)
            JobText = JobText & "require_file dvipdfmx " & _
                DockerShellQuote(Request.FilePrefix & ".pdf") & LF
            CurrentType = "PDF"
            CurrentExtension = ".pdf"
        End If
    Else
        CurrentType = "PDF"
        CurrentExtension = ".pdf"
        If Request.UseLatexmk Then
            AppendDockerStage JobText, StageCount, "latex", "latexmk " & _
                DockerOptionsForShell(Request.LatexmkPdfOptions) & _
                Request.ShellEscapeBatchOptions & DockerShellQuote(InputFilename)
        Else
            AppendDockerStage JobText, StageCount, "latex", Request.LatexCommand & _
                Request.ShellEscapeBatchOptions & DockerShellQuote(InputFilename)
        End If
        JobText = JobText & "require_file latex " & _
            DockerShellQuote(Request.FilePrefix & CurrentExtension) & LF
    End If

    If Request.UsePdf And CurrentType <> "PDF" Then
        AppendDockerStage JobText, StageCount, "dvipdfmx", _
            "dvipdfmx -o " & DockerShellQuote(Request.FilePrefix & ".pdf") & " " & _
            DockerShellQuote(Request.FilePrefix & CurrentExtension)
        JobText = JobText & "require_file dvipdfmx " & _
            DockerShellQuote(Request.FilePrefix & ".pdf") & LF
        CurrentType = "PDF"
        CurrentExtension = ".pdf"
    End If

    If Request.UseVector Then
        FinalFilename = Request.FilePrefix & ".svg"
        OutputType = "SVG"
        If CurrentType = "PDF" Then
            AppendDockerPdfToSvgStage JobText, StageCount, _
                Request.FilePrefix & CurrentExtension, FinalFilename
        Else
            AppendDockerStage JobText, StageCount, "dvisvgm", _
                "dvisvgm --no-fonts -o " & DockerShellQuote(FinalFilename) & _
                " " & DockerShellQuote(Request.FilePrefix & CurrentExtension)
        End If
        JobText = JobText & "require_file svg " & DockerShellQuote(FinalFilename) & LF
    ElseIf Request.PictureOutputType = "PDF" Then
        FinalFilename = Request.FilePrefix & "_tmp.pdf"
        OutputType = "PDF"
        JobText = JobText & BuildPdfToPdfJob(Request.FilePrefix & ".pdf", _
            FinalFilename, StageCount)
    ElseIf CurrentType = "PDF" Then
        FinalFilename = Request.FilePrefix & ".png"
        OutputType = "PNG"
        JobText = JobText & BuildPdfToPngJob(Request, StageCount)
    Else
        FinalFilename = Request.FilePrefix & ".png"
        OutputType = "PNG"
        JobText = JobText & "if [ ! -s " & DockerShellQuote(FinalFilename) & " ]; then" & LF
        AppendDockerStage JobText, StageCount, "dvipng", _
            "dvipng -q -D " & CStr(Request.OutputDpi) & _
            " -T tight -bg Transparent -o " & DockerShellQuote(FinalFilename) & _
            " " & DockerShellQuote(Request.FilePrefix & CurrentExtension), "  "
        JobText = JobText & "fi" & LF & _
            "require_file dvipng " & DockerShellQuote(FinalFilename) & LF
    End If

    JobText = JobText & "cat " & DockerShellQuote(FinalFilename)
    BuildDockerJob = JobText & LF
End Function

Private Function BuildDockerJobPreamble(ByVal TimeoutSeconds As Long, _
                                        ByVal LatexLogFilename As String) As String
    Dim LF As String
    LF = vbLf
    BuildDockerJobPreamble = "#!/bin/sh" & LF & _
        "set -u" & LF & _
        "IGUANATEX_TIMEOUT=" & DockerShellQuote(CStr(TimeoutSeconds)) & LF & _
        "IGUANATEX_LATEX_LOG=" & DockerShellQuote(LatexLogFilename) & LF & _
        "run_stage() {" & LF & _
        "  iguanatex_stage=$1" & LF & _
        "  shift" & LF & _
        "  printf '%s\n' ""IGUANATEX_STAGE:${iguanatex_stage}"" >&2" & LF & _
        "  timeout ""$IGUANATEX_TIMEOUT"" ""$@"" >&2" & LF & _
        "  iguanatex_status=$?" & LF & _
        "  if [ ""$iguanatex_status"" -ne 0 ]; then" & LF & _
        "    if [ ""$iguanatex_stage"" = latex ] && [ -f ""$IGUANATEX_LATEX_LOG"" ]; then" & LF & _
        "      cat ""$IGUANATEX_LATEX_LOG"" >&2" & LF & _
        "    fi" & LF & _
        "    printf '%s\n' ""IGUANATEX_ERROR:${iguanatex_stage}:${iguanatex_status}"" >&2" & LF & _
        "    exit ""$iguanatex_status""" & LF & _
        "  fi" & LF & _
        "}" & LF & _
        "require_file() {" & LF & _
        "  if [ ! -s ""$2"" ]; then" & LF & _
        "    printf '%s\n' ""IGUANATEX_ERROR:missing-output:$1:$2"" >&2" & LF & _
        "    exit 66" & LF & _
        "  fi" & LF & _
        "}" & LF
End Function

Private Function BuildDockerVectorConversionJob(ByVal InputFilename As String, _
                                                ByVal InputExtension As String, _
                                                ByVal FinalFilename As String, _
                                                ByVal TimeoutSeconds As Long, _
                                                ByRef StageCount As Long) As String
    Dim JobText As String
    Dim DvisvgmOptions As String
    Dim DvisvgmInput As String
    Dim UsePdfToSvg As Boolean
    JobText = BuildDockerJobPreamble(TimeoutSeconds, "vector-conversion.log")
    DvisvgmInput = InputFilename

    Select Case InputExtension
        Case "pdf"
            UsePdfToSvg = True
        Case "ps"
            DvisvgmInput = "vector-input.pdf"
            AppendDockerStage JobText, StageCount, "ps2pdf", _
                "ps2pdf " & DockerShellQuote(InputFilename) & " " & _
                DockerShellQuote(DvisvgmInput)
            JobText = JobText & "require_file ps2pdf " & _
                DockerShellQuote(DvisvgmInput) & vbLf
            UsePdfToSvg = True
        Case "eps"
            DvisvgmOptions = "--eps"
        Case "dvi", "xdv"
            DvisvgmOptions = "--no-fonts"
        Case Else
            Err.Raise vbObjectError + 2107, "DockerRender", _
                "Unsupported Docker vector input extension: " & InputExtension
    End Select

    If UsePdfToSvg Then
        AppendDockerPdfToSvgStage JobText, StageCount, DvisvgmInput, FinalFilename
    Else
        AppendDockerStage JobText, StageCount, "dvisvgm", _
            "dvisvgm " & DvisvgmOptions & " -o " & DockerShellQuote(FinalFilename) & _
            " " & DockerShellQuote(DvisvgmInput)
    End If
    JobText = JobText & "require_file svg " & DockerShellQuote(FinalFilename) & vbLf & _
        "cat " & DockerShellQuote(FinalFilename) & vbLf
    BuildDockerVectorConversionJob = JobText
End Function

Private Sub AppendDockerPdfToSvgStage(ByRef JobText As String, _
                                      ByRef StageCount As Long, _
                                      ByVal PdfFilename As String, _
                                      ByVal SvgFilename As String)
    Dim ConversionCommand As String
    ConversionCommand = "dvisvgm --pdf -o " & DockerShellQuote(SvgFilename) & _
        " " & DockerShellQuote(PdfFilename) & _
        " || { rm -f " & DockerShellQuote(SvgFilename) & _
        "; exec pdftocairo -svg " & DockerShellQuote(PdfFilename) & _
        " " & DockerShellQuote(SvgFilename) & "; }"
    AppendDockerStage JobText, StageCount, "pdf-to-svg", _
        "sh -c " & DockerShellQuote(ConversionCommand)
End Sub

Private Function BuildPdfToPdfJob(ByVal PdfFilename As String, _
                                  ByVal FinalPdf As String, _
                                  ByRef StageCount As Long) As String
    Dim JobText As String
    AppendDockerStage JobText, StageCount, "pdfcrop", _
        "pdfcrop --margins 0.1 " & DockerShellQuote(PdfFilename) & _
        " " & DockerShellQuote(FinalPdf)
    BuildPdfToPdfJob = JobText & "require_file pdfcrop " & _
        DockerShellQuote(FinalPdf) & vbLf
End Function

Private Function BuildPdfToPngJob(ByRef Request As DockerRenderRequest, _
                                  ByRef StageCount As Long) As String
    Dim LF As String
    Dim JobText As String
    Dim PdfFilename As String
    Dim BbxFilename As String
    Dim TemporaryPng As String
    Dim FinalPng As String
    LF = vbLf
    PdfFilename = Request.FilePrefix & ".pdf"
    BbxFilename = Request.FilePrefix & ".bbx"
    TemporaryPng = Request.FilePrefix & "_tmp.png"
    FinalPng = Request.FilePrefix & ".png"

    StageCount = StageCount + 1
    JobText = "printf '%s\n' 'IGUANATEX_STAGE:bbox' >&2" & LF & _
        "timeout ""$IGUANATEX_TIMEOUT"" gs -q -dBATCH -dNOPAUSE -sDEVICE=bbox " & _
        DockerShellQuote(PdfFilename) & " >/dev/null 2>" & DockerShellQuote(BbxFilename) & LF & _
        "iguanatex_status=$?" & LF & _
        "cat " & DockerShellQuote(BbxFilename) & " >&2" & LF & _
        "if [ ""$iguanatex_status"" -ne 0 ]; then" & LF & _
        "  printf '%s\n' ""IGUANATEX_ERROR:bbox:${iguanatex_status}"" >&2" & LF & _
        "  exit ""$iguanatex_status""" & LF & _
        "fi" & LF & _
        "iguanatex_bbox=$(awk -v dpi=" & CStr(Request.OutputDpi) & _
        " '/^%%HiResBoundingBox:/ {llx=$2-0.1; lly=$3-0.1; urx=$4+0.1; ury=$5+0.1; found=1} " & _
        "END {if (found) printf ""%d %d %.10g %.10g"", int((urx-llx)/72*dpi+0.5), int((ury-lly)/72*dpi+0.5), -llx, -lly}' " & _
        DockerShellQuote(BbxFilename) & ")" & LF & _
        "set -- $iguanatex_bbox" & LF & _
        "if [ ""$#"" -ne 4 ]; then" & LF & _
        "  printf '%s\n' 'IGUANATEX_ERROR:bbox:missing-hires-bounding-box' >&2" & LF & _
        "  exit 65" & LF & _
        "fi" & LF

    AppendDockerStage JobText, StageCount, "ghostscript-png", _
        "gs -q -dBATCH -dNOPAUSE -sDEVICE=pngalpha -r" & CStr(Request.OutputDpi) & _
        " -sOutputFile=" & DockerShellQuote(TemporaryPng) & _
        " ""-g${1}x${2}"" -c ""<</Install {${3} ${4} translate}>> setpagedevice"" -f " & _
        DockerShellQuote(PdfFilename)
    JobText = JobText & "require_file ghostscript-png " & DockerShellQuote(TemporaryPng) & LF

    AppendDockerStage JobText, StageCount, "imagemagick", _
        "magick -units PixelsPerInch " & DockerShellQuote(TemporaryPng) & _
        " -density 96 " & DockerShellQuote(FinalPng)
    JobText = JobText & "require_file imagemagick " & DockerShellQuote(FinalPng) & LF
    BuildPdfToPngJob = JobText
End Function

Private Sub AppendDockerStage(ByRef JobText As String, _
                              ByRef StageCount As Long, _
                              ByVal StageName As String, _
                              ByVal CommandText As String, _
                              Optional ByVal Prefix As String = vbNullString)
    JobText = JobText & Prefix & "run_stage " & DockerShellQuote(StageName) & _
        " " & CommandText & vbLf
    StageCount = StageCount + 1
End Sub

Private Function DockerOptionsForShell(ByVal OptionsText As String) As String
    DockerOptionsForShell = Replace(OptionsText, "$", "\$")
    If Len(DockerOptionsForShell) > 0 Then
        DockerOptionsForShell = DockerOptionsForShell & " "
    End If
End Function

Private Function DockerShellQuote(ByVal value As String) As String
    DockerShellQuote = "'" & Replace(value, "'", "'" & Chr$(34) & "'" & Chr$(34) & "'") & "'"
End Function

Private Function BuildDockerCommand(ByVal DockerImage As String) As String
    Dim BootstrapCommand As String
    BootstrapCommand = "mkdir -p /tmp/iguanatex && cd /tmp/iguanatex && tar -xf - && exec sh job.sh"
    #If Mac Then
        BuildDockerCommand = ShellEscape(DEFAULT_DOCKER_COMMAND) & _
            " run --rm -i --network none --pull never " & ShellEscape(DockerImage) & _
            " sh -c " & ShellEscape(BootstrapCommand)
    #Else
        BuildDockerCommand = ShellEscape(DEFAULT_DOCKER_COMMAND) & _
            " run --rm -i --network none --pull never " & ShellEscape(DockerImage) & _
            " sh -c """ & BootstrapCommand & """"
    #End If
End Function

Private Function DockerOverallWaitTime(ByVal StageWaitTime As Long, _
                                       ByVal StageCount As Long) As Long
    If StageWaitTime <= 0 Then
        DockerOverallWaitTime = -1
        Exit Function
    End If

    Dim TotalWait As Double
    TotalWait = CDbl(StageWaitTime) * CDbl(StageCount + 2)
    If TotalWait > 2147483000# Then TotalWait = 2147483000#
    DockerOverallWaitTime = CLng(TotalWait)
End Function

Private Function ReadDockerFailureStage(ByVal LogPath As String) As String
    ReadDockerFailureStage = "docker"
    If Not FileExists(LogPath) Then Exit Function

    Dim LogText As String
    Dim MarkerPosition As Long
    Dim MarkerEnd As Long
    Const Marker As String = "IGUANATEX_STAGE:"
    LogText = ReadAll(LogPath)
    MarkerPosition = InStrRev(LogText, Marker, -1, vbBinaryCompare)
    If MarkerPosition = 0 Then Exit Function
    MarkerPosition = MarkerPosition + Len(Marker)
    MarkerEnd = InStr(MarkerPosition, LogText, vbLf, vbBinaryCompare)
    If MarkerEnd = 0 Then MarkerEnd = Len(LogText) + 1
    ReadDockerFailureStage = Trim$(Replace(Mid$(LogText, MarkerPosition, _
        MarkerEnd - MarkerPosition), vbCr, vbNullString))
End Function

Private Function CreateDockerWorkspace(ByVal TempPath As String) As String
    Dim WorkspaceStem As String
    Dim WorkspacePath As String
    Dim Attempt As Long
    WorkspaceStem = "IguanaTex-render-" & Format$(Now, "yyyymmdd-hhnnss") & _
        "-" & Right$("00000000" & Hex$(CLng(Timer * 1000)), 8)

    For Attempt = 0 To 999
        WorkspacePath = TempPath & WorkspaceStem & "-" & CStr(Attempt)
        If Not DockerFolderExists(WorkspacePath) And Not FileExists(WorkspacePath) Then
            MkDir WorkspacePath
            CreateDockerWorkspace = WorkspacePath & PathSep
            Exit Function
        End If
    Next

    Err.Raise vbObjectError + 2105, "DockerRender", _
        "Could not allocate a unique private Docker render workspace under " & TempPath
End Function

Private Function DockerFolderExists(ByVal FolderPath As String) As Boolean
    On Error GoTo FolderMissing
    DockerFolderExists = ((GetAttr(FolderPath) And vbDirectory) <> 0)
    Exit Function
FolderMissing:
    DockerFolderExists = False
End Function

Private Sub StageDockerAuxiliaryFiles(ByVal SourcePath As String, _
                                      ByVal WorkspacePath As String, _
                                      ByVal FilePrefix As String, _
                                      ByRef HostOperation As String, _
                                      ByRef HostPath As String)
    HostOperation = "enumerate configured auxiliary source"
    HostPath = SourcePath

    ' A shared OS temp root is never an auxiliary source. Users who need custom
    ' root-level inputs can keep using the existing Temp folder setting with a
    ' dedicated directory.
    If IsSystemTemporaryRoot(SourcePath) Then
        HostOperation = "skip shared system temporary root auxiliary staging"
        Exit Sub
    End If

    Dim Candidate As String
    Dim FullPath As String
    Dim Attributes As Long
    Candidate = Dir(SourcePath & "*", DOCKER_ENUM_ATTRIBUTES)
    Do While Candidate <> vbNullString
        FullPath = SourcePath & Candidate
        Attributes = GetAttr(FullPath)
        If (Attributes And vbDirectory) = 0 Then
            If IsDockerAuxiliaryInput(Candidate, FilePrefix) Then
                HostOperation = "copy auxiliary input into private render workspace"
                HostPath = FullPath
                FileCopy FullPath, WorkspacePath & Candidate
            End If
        End If
        Candidate = Dir()
    Loop
End Sub

Private Function IsSystemTemporaryRoot(ByVal SourcePath As String) As Boolean
    Dim NormalizedSource As String
    NormalizedSource = NormalizeDockerHostPath(SourcePath)

    #If Mac Then
        IsSystemTemporaryRoot = _
            (NormalizedSource = NormalizeDockerHostPath(MacTempPath()))
    #Else
        Dim SystemTempPath As String
        SystemTempPath = Environ$("TEMP")
        If SystemTempPath <> vbNullString Then
            If NormalizedSource = NormalizeDockerHostPath(SystemTempPath) Then
                IsSystemTemporaryRoot = True
                Exit Function
            End If
        End If

        SystemTempPath = Environ$("TMP")
        If SystemTempPath <> vbNullString Then
            IsSystemTemporaryRoot = _
                (NormalizedSource = NormalizeDockerHostPath(SystemTempPath))
        End If
    #End If
End Function

Private Function NormalizeDockerHostPath(ByVal HostPath As String) As String
    HostPath = Trim$(Replace(HostPath, WrongPathSep, PathSep))
    If HostPath = vbNullString Then Exit Function
    If Right$(HostPath, 1) <> PathSep Then HostPath = HostPath & PathSep
    #If Mac Then
        NormalizeDockerHostPath = HostPath
    #Else
        NormalizeDockerHostPath = LCase$(HostPath)
    #End If
End Function

Private Function IsDockerAuxiliaryInput(ByVal FileName As String, _
                                        ByVal FilePrefix As String) As Boolean
    If LCase$(Left$(FileName, Len(FilePrefix))) = LCase$(FilePrefix) Then Exit Function
    If LCase$(FileName) = LCase$(DOCKER_JOB_FILENAME) Then Exit Function

    Dim Extension As String
    Extension = LCase$(GetExtension(FileName))
    Select Case Extension
        Case "tex", "ltx", "sty", "cls", "clo", "cfg", "cnf", "def", "fd", _
             "aux", "bbl", "bib", "bst", "bbx", "cbx", "lbx", "idx", "ind", _
             "ist", "xdy", "gls", "glo", "acr", "acn", "tikz", "pgf", _
             "map", "enc", "tfm", "vf", "afm", "pfm", "pfb", "otf", "ttf", "ttc", _
             "png", "jpg", "jpeg", "jbig2", "jp2", "pdf", "eps", "epsi", "ps", _
             "svg", "mps", "tif", "tiff", "bmp", "gif", "pdf_tex", "eps_tex", _
             "csv", "tsv", "dat", "txt", "json", "xml", "yaml", "yml", _
             "lua", "py", "pl", "r", "m", "sh", "latexmkrc"
            IsDockerAuxiliaryInput = True
    End Select
End Function

Private Sub WriteDockerPayload(ByVal PayloadPath As String, _
                               ByVal WorkspacePath As String, _
                               ByRef HostOperation As String, _
                               ByRef HostPath As String)
    On Error GoTo PayloadError

    Dim WorkspaceNames As New Collection
    CollectDockerWorkspaceFiles WorkspacePath, WorkspaceNames, HostOperation, HostPath

    Dim TarFile As Integer
    Dim TarIsOpen As Boolean
    HostOperation = "create Docker payload tar"
    HostPath = PayloadPath
    TarFile = FreeFile()
    Open PayloadPath For Binary Access Write As #TarFile
    TarIsOpen = True

    Dim WorkspaceName As Variant
    For Each WorkspaceName In WorkspaceNames
        HostOperation = "write private workspace file to Docker payload"
        HostPath = WorkspacePath & CStr(WorkspaceName)
        WriteTarFileEntry TarFile, CStr(WorkspaceName), HostPath
    Next

    HostOperation = "finalize Docker payload tar"
    HostPath = PayloadPath
    Dim EndBlocks(0 To TAR_BLOCK_SIZE * 2 - 1) As Byte
    Put #TarFile, , EndBlocks
    Close #TarFile
    TarIsOpen = False
    Exit Sub

PayloadError:
    Dim PayloadErrorNumber As Long
    Dim PayloadErrorSource As String
    Dim PayloadErrorDescription As String
    PayloadErrorNumber = Err.Number
    PayloadErrorSource = Err.Source
    PayloadErrorDescription = Err.Description
    On Error Resume Next
    If TarIsOpen Then Close #TarFile
    On Error GoTo 0
    If PayloadErrorNumber = 0 Then
        PayloadErrorNumber = vbObjectError + 2106
        PayloadErrorSource = "DockerRender"
        PayloadErrorDescription = "Unknown error while writing the Docker payload tar."
    End If
    Err.Raise PayloadErrorNumber, PayloadErrorSource, PayloadErrorDescription
End Sub

Private Sub CollectDockerWorkspaceFiles(ByVal WorkspacePath As String, _
                                        ByRef WorkspaceNames As Collection, _
                                        ByRef HostOperation As String, _
                                        ByRef HostPath As String)
    HostOperation = "enumerate private render workspace"
    HostPath = WorkspacePath

    Dim Candidate As String
    Dim FullPath As String
    Dim Attributes As Long
    Candidate = Dir(WorkspacePath & "*", DOCKER_ENUM_ATTRIBUTES)
    Do While Candidate <> vbNullString
        FullPath = WorkspacePath & Candidate
        Attributes = GetAttr(FullPath)
        If (Attributes And vbDirectory) = 0 Then
            WorkspaceNames.Add Candidate
        End If
        Candidate = Dir()
    Loop
End Sub

Private Sub PublishDockerLog(ByVal WorkspaceLogPath As String, _
                             ByVal RootLogPath As String)
    On Error Resume Next
    If FileExists(WorkspaceLogPath) Then
        DeleteDockerFile RootLogPath
        FileCopy WorkspaceLogPath, RootLogPath
    End If
    On Error GoTo 0
End Sub

Private Sub DeleteDockerWorkspace(ByVal WorkspacePath As String)
    On Error Resume Next
    #If Mac Then
        Dim MacFs As New MacFileSystemObject
        If MacFs.FolderExists(WorkspacePath) Then MacFs.FindDelete WorkspacePath, "*"
        Set MacFs = Nothing
    #Else
        Dim fs As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.FolderExists(WorkspacePath) Then fs.DeleteFolder WorkspacePath, True
        Set fs = Nothing
    #End If
    On Error GoTo 0
End Sub

Private Sub WriteTarFileEntry(ByVal TarFile As Integer, _
                              ByVal ArchiveName As String, _
                              ByVal SourcePath As String)
    Dim FileData() As Byte
    Dim FileSize As Long
    FileData = ReadAllBytes(SourcePath)
    FileSize = ArrayLength(FileData)

    Dim Header(0 To TAR_BLOCK_SIZE - 1) As Byte
    PutTarName Header, ArchiveName
    PutTarText Header, 100, 8, TarOctalField(420, 8)
    PutTarText Header, 108, 8, TarOctalField(0, 8)
    PutTarText Header, 116, 8, TarOctalField(0, 8)
    PutTarText Header, 124, 12, TarOctalField(FileSize, 12)
    PutTarText Header, 136, 12, TarOctalField(0, 12)
    PutTarText Header, 148, 8, Space$(8)
    Header(156) = Asc("0")
    PutTarText Header, 257, 6, "ustar" & Chr$(0)
    PutTarText Header, 263, 2, "00"
    PutTarText Header, 265, 32, "iguanatex"
    PutTarText Header, 297, 32, "iguanatex"

    Dim Checksum As Long
    Dim i As Long
    For i = 0 To TAR_BLOCK_SIZE - 1
        Checksum = Checksum + CLng(Header(i))
    Next
    PutTarText Header, 148, 8, TarChecksumField(Checksum)

    Put #TarFile, , Header
    If FileSize > 0 Then Put #TarFile, , FileData

    Dim PaddingSize As Long
    PaddingSize = (TAR_BLOCK_SIZE - (FileSize Mod TAR_BLOCK_SIZE)) Mod TAR_BLOCK_SIZE
    If PaddingSize > 0 Then
        Dim Padding() As Byte
        ReDim Padding(0 To PaddingSize - 1)
        Put #TarFile, , Padding
    End If
End Sub

Private Sub PutTarName(ByRef Header() As Byte, ByVal ArchiveName As String)
    Dim NormalizedName As String
    Dim NameBytes() As Byte
    Dim NameLength As Long
    NormalizedName = Replace(ArchiveName, "\", "/")
    NameBytes = DockerUtf8Bytes(NormalizedName)
    NameLength = ArrayLength(NameBytes)
    If NameLength = 0 Then
        Err.Raise vbObjectError + 2101, "DockerRender", "Docker payload filename is empty."
    End If

    If NameLength <= 100 Then
        PutTarBytes Header, 0, NameBytes
        Exit Sub
    End If

    Dim SlashPosition As Long
    Dim PrefixText As String
    Dim LeafText As String
    Dim PrefixBytes() As Byte
    Dim LeafBytes() As Byte
    SlashPosition = InStrRev(NormalizedName, "/")
    Do While SlashPosition > 0
        PrefixText = Left$(NormalizedName, SlashPosition - 1)
        LeafText = Mid$(NormalizedName, SlashPosition + 1)
        PrefixBytes = DockerUtf8Bytes(PrefixText)
        LeafBytes = DockerUtf8Bytes(LeafText)
        If ArrayLength(PrefixBytes) <= 155 And ArrayLength(LeafBytes) <= 100 Then
            PutTarBytes Header, 0, LeafBytes
            PutTarBytes Header, 345, PrefixBytes
            Exit Sub
        End If
        SlashPosition = InStrRev(PrefixText, "/")
    Loop

    Err.Raise vbObjectError + 2101, "DockerRender", _
        "Docker payload path does not fit in a ustar header: " & ArchiveName
End Sub

Private Sub PutTarBytes(ByRef Header() As Byte, _
                        ByVal Offset As Long, _
                        ByRef value() As Byte)
    Dim i As Long
    For i = 0 To ArrayLength(value) - 1
        Header(Offset + i) = value(i)
    Next
End Sub

Private Sub PutTarText(ByRef Header() As Byte, _
                       ByVal Offset As Long, _
                       ByVal FieldLength As Long, _
                       ByVal value As String)
    If Len(value) > FieldLength Then
        Err.Raise vbObjectError + 2102, "DockerRender", "Tar header field is too long."
    End If

    Dim i As Long
    For i = 1 To Len(value)
        Header(Offset + i - 1) = AscW(Mid$(value, i, 1)) And &HFF&
    Next
End Sub

Private Function TarOctalField(ByVal value As Long, ByVal FieldLength As Long) As String
    Dim Digits As String
    Digits = Oct$(value)
    If Len(Digits) > FieldLength - 1 Then
        Err.Raise vbObjectError + 2103, "DockerRender", "Value is too large for tar header."
    End If
    TarOctalField = String$(FieldLength - Len(Digits) - 1, "0") & Digits & Chr$(0)
End Function

Private Function TarChecksumField(ByVal value As Long) As String
    Dim Digits As String
    Digits = Oct$(value)
    If Len(Digits) > 6 Then
        Err.Raise vbObjectError + 2104, "DockerRender", "Tar checksum is too large."
    End If
    TarChecksumField = String$(6 - Len(Digits), "0") & Digits & Chr$(0) & " "
End Function

Private Function DockerUtf8Bytes(ByVal value As String) As Byte()
    #If Mac Then
        DockerUtf8Bytes = StringToUtf8(value)
    #Else
        Dim TextStream As Object
        Dim BinaryStream As Object
        Dim Result() As Byte
        Set TextStream = CreateObject("ADODB.Stream")
        Set BinaryStream = CreateObject("ADODB.Stream")

        TextStream.Type = 2
        TextStream.Charset = "utf-8"
        TextStream.Open
        TextStream.WriteText value
        TextStream.Position = 3

        BinaryStream.Type = 1
        BinaryStream.Open
        TextStream.CopyTo BinaryStream
        BinaryStream.Position = 0
        Result = BinaryStream.Read

        TextStream.Close
        BinaryStream.Close
        Set TextStream = Nothing
        Set BinaryStream = Nothing
        DockerUtf8Bytes = Result
    #End If
End Function

Private Sub DeleteDockerFile(ByVal FilePath As String)
    On Error Resume Next
    If FileExists(FilePath) Then
        SetAttr FilePath, vbNormal
        Kill FilePath
    End If
    On Error GoTo 0
End Sub
