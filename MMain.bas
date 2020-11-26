Attribute VB_Name = "MMain"
Option Compare Database
Option Explicit

Private Const TEST_FILE1 As String = "c:\temp\textfiles\notepad_text.txt"

' Chaining all tests

Public Sub TestAll()
  InitMD5
  Test_ReadTextFileByLine
  Test_ReadTextFileByLine_BadIdea
  Test_ReadNotepadTextFileWithByteArray
  Test_Normalization
  Test_ReadSampleEncodings
End Sub

' Read a file line by line with VBA statements and opening
' the file as a text file:

Public Sub ReadTextFileByLine(ByVal psFileName As String)
  Dim hFile     As Integer
  Dim sLine     As String
  Dim iLineCt   As Integer
  
  Debug.Print "---- ReadTextFileByLine()"
  
  hFile = FreeFile
  Open psFileName For Input Access Read As #hFile
  Debug.Print "[File: " & psFileName & "]"
  While Not EOF(hFile)
    iLineCt = iLineCt + 1
    Line Input #hFile, sLine
    Debug.Print iLineCt & ":" & sLine
  Wend
  Close hFile
  Debug.Print "[EOF]"
End Sub

Public Sub Test_ReadTextFileByLine()
  ReadTextFileByLine TEST_FILE1
End Sub

' Trying to UTF8 decode lines read when the file is open
' as text doesn't work:

Public Sub ReadTextFileByLine_BadIdea(ByVal psFileName As String)
  Dim hFile     As Integer
  Dim sLine     As String
  Dim iLineCt   As Integer
  
  Debug.Print "---- ReadTextFileByLine_BadIdea()"
  
  hFile = FreeFile
  Open psFileName For Input Access Read As #hFile
  Debug.Print "[File: " & psFileName & "]"
  While Not EOF(hFile)
    iLineCt = iLineCt + 1
    Line Input #hFile, sLine
    If iLineCt = 1 Then DumpStringBytes sLine
    Debug.Print iLineCt & ":" & UTF8DecodeString(sLine)
  Wend
  Close hFile
  Debug.Print "[EOF]"
End Sub

Public Sub Test_ReadTextFileByLine_BadIdea()
  ReadTextFileByLine_BadIdea TEST_FILE1
End Sub

' Opening an UTF8 file in binary mode and converting to UTF8
' the contents read into a byte array works:

Public Sub ReadNotepadTextFileWithByteArray(ByVal psFileName As String)
  Dim abString()  As Byte
  Dim sDecoded    As String
  Dim hFile       As Integer
  
  Debug.Print "---- ReadNotepadTextFile() in a byte array and convert to UTF8:"
  
  Debug.Print "[File: " & psFileName & "]"
  hFile = FreeFile
  Open psFileName For Binary Access Read As #hFile
  ReDim abString(1 To LOF(hFile)) As Byte
  Get #hFile, 1, abString
  Close hFile
  
  sDecoded = UTF8DecodeByteArrayToString(abString)
  Debug.Print sDecoded & vbCrLf & "[UTF8] (len=" & Len(sDecoded) & ")"
  
  Debug.Print "[EOF]"
End Sub

Public Sub Test_ReadNotepadTextFileWithByteArray()
  ReadNotepadTextFileWithByteArray TEST_FILE1
End Sub
  
  'ReadNotepadTextFile TEST_FILE1, "utf8"
  'Open with word
  'Shell """C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"" " & TEST_FILE1

' Normalizing a string to "C" form with the Win32 API:

Private Function TryNormalization(ByVal psText As String) As String
  If UcIsNormalizedString(psText) Then
    TryNormalization = "Already normalized in this form"
    Exit Function
  End If
  TryNormalization = UcNormalizeString(psText)
  If (UcGetLastError() <> 0) Or (Len(UcGetLastErrorText()) > 0) Then
    TryNormalization = "FAILED: " & UcGetLastErrorText()
  End If
End Function

Public Sub Test_Normalization()
  Dim sInput        As String
  Dim sNormalized   As String
  
  sInput = ChrW$(&HE8) & "st string " & ChrW$(&HFF54) & ChrW$(&HFF4F) & " n" & ChrW$(&HF8) & "rm" & ChrW$(&HE4) & "lize"
  Debug.Print "Comparison of Normalization Forms"
  Debug.Print "input string: " & sInput
  Debug.Print "normalized: " & TryNormalization(sInput)
  Debug.Print "Attempt to normalize illegal lone surrogate:"
  sInput = "Bad surrogate is here: '" & ChrW$(&HD800) & "'"
  Debug.Print sInput
  Debug.Print "normalized: " & TryNormalization(sInput)
  Debug.Print "Attempt to normalize a string that expands beyond the initial guess"
  sInput = String$(48, ChrW$(&H958))
  sInput = sInput & String$(24, ChrW$(&HFB2C))
  sNormalized = TryNormalization(sInput)
  Debug.Print "normalized: start length=" & Len(sInput) & ", end length=" & Len(sNormalized)
  Debug.Print "normalized: " & sNormalized & "<<"
  Debug.Print "normalized last char code=" & AscW(Right$(sNormalized, 1))
End Sub

' Read other encodings with GetFileText()

Public Sub Test_ReadSampleEncodings()
  Const TESTFILES_SUBFOLDER As String = "text_files_samples"
  Dim sTestFilesPath    As String
  Dim sTestFileName     As String
  Dim sText             As String
  Dim sHashCheck        As String
  Dim sFileHash         As String
  
  sTestFilesPath = CurrentProject.Path & "\" & TESTFILES_SUBFOLDER & "\"
  
  'Read the UTF16 LE (No BOM) version (ie VBA encoding when writing a text file)
  sTestFileName = sTestFilesPath & "UCS2 (UTF16) LE BOM.txt"
  sText = GetFileText(sTestFileName, "utf16")
  Debug.Print "---- [" & sTestFileName & "]:"
  Debug.Print sText
  Debug.Print "[EOF]"
  'Get its hash for reference
  sHashCheck = MD5Hex(sText)
  Debug.Print "REFERENCE HASH: " & sHashCheck & " (MD5)"
  Debug.Print
  
  'Check with UTF16 BE BOM file
  sTestFileName = sTestFilesPath & "UCS2 (UTF16) BE BOM.txt"
  sText = GetFileText(sTestFileName, "utf16")
  sFileHash = MD5Hex(sText)
  Debug.Print "[" & sTestFileName & "] MD5: " & sFileHash
  If sFileHash = sHashCheck Then
    Debug.Print "OK - hashes match"
  Else
    Debug.Print "FAILED - hash mismatch"
  End If
  
  'Check with UTF16 LE BOM file
  sTestFileName = sTestFilesPath & "UCS2 (UTF16) LE BOM.txt"
  sText = GetFileText(sTestFileName, "utf16")
  sFileHash = MD5Hex(sText)
  Debug.Print "[" & sTestFileName & "] MD5: " & sFileHash
  If sFileHash = sHashCheck Then
    Debug.Print "OK - hashes match"
  Else
    Debug.Print "FAILED - hash mismatch"
  End If
  
  'Check with UTF8 BOM file
  sTestFileName = sTestFilesPath & "UTF8 BOM.txt"
  sText = GetFileText(sTestFileName, "utf8")
  sFileHash = MD5Hex(sText)
  Debug.Print "[" & sTestFileName & "] MD5: " & sFileHash
  If sFileHash = sHashCheck Then
    Debug.Print "OK - hashes match"
  Else
    Debug.Print "FAILED - hash mismatch"
  End If
  
  'Check with UTF8 *NO* BOM file
  sTestFileName = sTestFilesPath & "UTF8.txt"
  sText = GetFileText(sTestFileName, "utf8")
  sFileHash = MD5Hex(sText)
  Debug.Print "[" & sTestFileName & "] MD5: " & sFileHash
  If sFileHash = sHashCheck Then
    Debug.Print "OK - hashes match"
  Else
    Debug.Print "FAILED - hash mismatch"
  End If
  
End Sub

