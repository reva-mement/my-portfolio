Option Explicit

' === [1] User Config ===
Const PICS_PER_ROW = 4
Const SLOT_WIDTH_CM = 6.5
Const MARGIN_LEFT_CM = 1
Const MARGIN_TOP_CM = 2
Const GAP_H_CM = 0.4
Const GAP_V_CM = 0.3
Const DEFAULT_HEIGHT_CM = 8
Const INLINE_EXCLUDE = True
Const AUTO_SAVE = False
Const AUTO_CLOSE = False
Const FILENAME_MODE = 1 ' 0=msg name, 1=anon+timestamp+seq, 2=subject

' === [2] Utilities ===
Function CmToPt(cm)
  CmToPt = cm * 28.3464566929
End Function

Function IsImageFile(ext)
  ext = LCase(ext)
  IsImageFile = (ext = "jpg" Or ext = "jpeg" Or ext = "png" Or ext = "bmp" Or ext = "gif")
End Function

Function LooksAutoInlineName(fname)
  LooksAutoInlineName = LCase(Left(fname, 5)) = "image" And InStr(fname, ".png") > 0
End Function

Function SafeName(s)
  Dim badChars, i
  badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
  For i = 0 To UBound(badChars)
    s = Replace(s, badChars(i), "_")
  Next
  SafeName = s
End Function

Function TimeStampStr()
  Dim t: t = Now
  TimeStampStr = Year(t) & Right("0" & Month(t),2) & Right("0" & Day(t),2) & "_" & _
                 Right("0" & Hour(t),2) & Right("0" & Minute(t),2) & Right("0" & Second(t),2)
End Function

Function UniqueSavePath(baseFolder, baseName, ext)
  Dim fso, i, path
  Set fso = CreateObject("Scripting.FileSystemObject")
  i = 1
  Do
    path = baseFolder & "\" & baseName & "_" & TimeStampStr() & "_" & i & "." & ext
    If Not fso.FileExists(path) Then Exit Do
    i = i + 1
  Loop
  UniqueSavePath = path
End Function

' === [3] Input Height ===
Dim heightInput, heightCm
heightInput = InputBox("Enter image height in cm (1–100):", "Image Height", DEFAULT_HEIGHT_CM)
If Not IsNumeric(heightInput) Then WScript.Quit
heightCm = CDbl(heightInput)
If heightCm < 1 Or heightCm > 100 Then WScript.Quit

' === [4] Select Template ===
Dim excel, fso, shell, templatePath
Set excel = CreateObject("Excel.Application")
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("Shell.Application")
templatePath = excel.Application.GetOpenFilename("Excel Templates (*.xlt;*.xltx), *.xlt;*.xltx", , "Select Excel Template")
If templatePath = "False" Then WScript.Quit

' === [5] Process .msg Files ===
Dim args, i, msgPath, outlook, ns, msg, attach, tempFolder, tempPath
Dim book, sheet, row, col, imgCount, mailCount
Set args = WScript.Arguments
If args.Count = 0 Then
  MsgBox "Please drag and drop .msg files onto this script.", vbInformation
  WScript.Quit
End If

Set outlook = CreateObject("Outlook.Application")
Set ns = outlook.GetNamespace("MAPI")
Set tempFolder = fso.GetSpecialFolder(2)
imgCount = 0 : mailCount = 0

For i = 0 To args.Count - 1
  msgPath = args(i)
  If LCase(fso.GetExtensionName(msgPath)) <> "msg" Then Continue For

  Set msg = ns.OpenSharedItem(msgPath)
  If msg Is Nothing Then Continue For

  Set book = excel.Workbooks.Add(templatePath)
  Set sheet = book.Sheets(1)
  excel.Visible = True
  excel.DisplayAlerts = False

  row = 0 : col = 0
  Dim attachCount, j, imgPath, shape, slotW, slotH, xPt, yPt, slotsUsed
  attachCount = msg.Attachments.Count

  For j = 1 To attachCount
    Set attach = msg.Attachments(j)
    Dim fname, ext
    fname = attach.FileName
    ext = fso.GetExtensionName(fname)
    If Not IsImageFile(ext) Then Continue For
    If INLINE_EXCLUDE Then
      If LooksAutoInlineName(fname) Then Continue For
      On Error Resume Next
      If Not attach.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E") = "" Then
        Err.Clear
        Continue For
      End If
      On Error GoTo 0
    End If

    imgPath = tempFolder & "\" & fso.GetTempName
    attach.SaveAsFile imgPath

    slotH = CmToPt(heightCm)
    slotW = CmToPt(SLOT_WIDTH_CM)
    xPt = CmToPt(MARGIN_LEFT_CM + col * (SLOT_WIDTH_CM + GAP_H_CM))
    yPt = CmToPt(MARGIN_TOP_CM + row * (heightCm + GAP_V_CM))

    Set shape = sheet.Shapes.AddPicture(imgPath, False, True, xPt, yPt, -1, -1)
    shape.LockAspectRatio = True
    shape.Height = slotH
    Dim wPt: wPt = shape.Width
    slotsUsed = Int((wPt + CmToPt(GAP_H_CM)) / CmToPt(SLOT_WIDTH_CM))
    If slotsUsed < 1 Then slotsUsed = 1

    col = col + slotsUsed
    If col >= PICS_PER_ROW Then
      col = 0
      row = row + 1
    End If

    fso.DeleteFile imgPath, True
    imgCount = imgCount + 1
  Next

  Dim savePath
  Select Case FILENAME_MODE
    Case 0
      savePath = fso.GetBaseName(msgPath)
    Case 1
      savePath = "Export_" & TimeStampStr() & "_" & i + 1
    Case 2
      savePath = SafeName(msg.Subject)
  End Select
  savePath = fso.GetParentFolderName(msgPath) & "\" & savePath & ".xlsx"

  If AUTO_SAVE Then
    book.SaveAs savePath
    If AUTO_CLOSE Then book.Close False
  End If

  mailCount = mailCount + 1
Next

' === [6] Cleanup ===
excel.DisplayAlerts = True
Set sheet = Nothing
Set book = Nothing
Set excel = Nothing
Set outlook = Nothing
Set ns = Nothing
Set fso = Nothing
Set shell = Nothing

' === [7] Summary ===
MsgBox mailCount & " message(s) processed." & vbCrLf & imgCount & " image(s) extracted.", vbInformation, "Done"
