Attribute VB_Name = "Defaults"
Option Explicit

#If Mac Then
Public Const DEFAULT_TEMP_DIR As String = vbNullString
Public Const DEFAULT_VECTOR_OUTPUT_TYPE As String = "dvisvgm"
Public Const DEFAULT_PICTURE_OUTPUT_TYPE As String = "PDF"
Public Const DEFAULT_EDITOR As String = "open -b 'texstudio'"
Public Const DEFAULT_ADDIN_FOLDER As String = "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/"
Public Const DEFAULT_LATEXIT_METADATA_COMMAND As String = DEFAULT_ADDIN_FOLDER & "LaTeXiT-metadata-macos"
Public Const NEWLINE As String = vbLf
Public Const PathSep As String = "/"
Public Const WrongPathSep As String = "\"

#Else
Public Const DEFAULT_TEMP_DIR As String = "c:\temp\"
Public Const DEFAULT_VECTOR_OUTPUT_TYPE As String = "dvisvgm"
Public Const DEFAULT_PICTURE_OUTPUT_TYPE As String = "PNG"
Public Const DEFAULT_EDITOR As String = "C:\Program Files (x86)\TeXstudio\texstudio.exe"
Public Const DEFAULT_LATEXIT_METADATA_COMMAND As String = "%USERPROFILE%\Downloads\LaTeXiT-metadata\LaTeXiT-metadata-win.exe"
Public Const NEWLINE As String = vbCrLf
Public Const PathSep As String = "\"
Public Const WrongPathSep As String = "/"

#End If

Public Const DEFAULT_DOCKER_COMMAND As String = "docker"
Public Const DEFAULT_DOCKER_IMAGE As String = "danteev/texlive:latest"

Public Const IGUANATEX_VERSION As Integer = 162

Public Const DEFAULT_LATEX_CODE As String = "\documentclass{article}" & NEWLINE & "\usepackage{amsmath}" & NEWLINE & "\pagestyle{empty}" & NEWLINE & _
                                            "\begin{document}" & NEWLINE & NEWLINE & NEWLINE & NEWLINE & NEWLINE & "\end{document}"
Public Const DEFAULT_LATEX_CODE_PRE As String = "\documentclass{article}" & NEWLINE & "\usepackage{amsmath}" & NEWLINE & "\pagestyle{empty}" & NEWLINE & _
                                                "\begin{document}" & NEWLINE & NEWLINE
Public Const DEFAULT_LATEX_CODE_POST As String = NEWLINE & NEWLINE & "\end{document}"
