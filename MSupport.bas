Attribute VB_Name = "MSupport"
Option Compare Database
Option Explicit

Public Sub DumpStringBytes(ByVal psText As String)
  Dim I           As Integer
  Dim abBytes()   As Byte
  Dim sOut        As String
  
  abBytes = psText
  For I = 0 To UBound(abBytes) / 2
    Debug.Print Format$(I + 1, "00000"); "|";
  Next I
  Debug.Print
  
  For I = 0 To UBound(abBytes)
    Debug.Print Format$(I + 1, "00"); "|";
  Next I
  Debug.Print
  
  For I = 0 To UBound(abBytes)
    Debug.Print "--+";
  Next I
  Debug.Print
  
  For I = 0 To UBound(abBytes)
    sOut = Hex$(abBytes(I))
    If Len(sOut) < 2 Then sOut = " " & sOut
    Debug.Print sOut; "|";
  Next I
  Debug.Print
  
  For I = 0 To UBound(abBytes)
    Debug.Print "--+";
  Next I
  Debug.Print
    
  For I = 0 To UBound(abBytes)
    sOut = Chr$(abBytes(I))
    If Len(sOut) < 1 Then sOut = "?"
    Debug.Print sOut; " |";
  Next I
  Debug.Print
    
End Sub

Public Function SplitString(ByRef asRetItems() As String, _
  ByVal sToSplit As String, _
  Optional sSep As String = " ", _
  Optional lMaxItems As Long = 0&, _
  Optional eCompare As VbCompareMethod = vbBinaryCompare) _
  As Long

  Dim lPos        As Long
  Dim lDelimLen   As Long
  Dim lRetCount   As Long
  
  On Error Resume Next
  Erase asRetItems
  On Error GoTo SplitString_Err
  
  If Len(sToSplit) Then
    lDelimLen = Len(sSep)
    If lDelimLen Then
      lPos = InStr(1, sToSplit, sSep, eCompare)
      Do While lPos
        lRetCount = lRetCount + 1&
        ReDim Preserve asRetItems(1& To lRetCount)
        asRetItems(lRetCount) = Left$(sToSplit, lPos - 1&)
        sToSplit = Mid$(sToSplit, lPos + lDelimLen)
        If lMaxItems Then
          If lRetCount = lMaxItems - 1& Then Exit Do
        End If
        lPos = InStr(1, sToSplit, sSep, eCompare)
      Loop
    End If
    lRetCount = lRetCount + 1&
    ReDim Preserve asRetItems(1& To lRetCount)
    asRetItems(lRetCount) = sToSplit
  End If
  SplitString = lRetCount
SplitString_Err:
End Function

' Source: Hardcore Visual Basic (book), http://vb.mvps.org/hardweb/mckinney.htm
Public Function ExistFile(psSpec As String) As Boolean
  On Error Resume Next
  Call FileLen(psSpec)
  ExistFile = (Err.Number = 0&)
End Function

#If ACCEPT_CRASH_ONWIN64 Then
  'https://gist.github.com/ken-itakura/c35455bccbae1544189e37b713698b75
  Public Function MD5Hex(textString As String) As String
    Dim enc
    Dim textBytes() As Byte
    Dim bytes
    Dim outstr As String
    Dim pos As Integer
    
    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    textBytes = textString
    bytes = enc.ComputeHash_2((textBytes))
      
    For pos = 1 To LenB(bytes)
      outstr = outstr & LCase(Right("0" & Hex(AscB(MidB(bytes, pos, 1))), 2))
    Next
    MD5Hex = outstr
    Set enc = Nothing
  End Function
#Else
  'https://stackoverflow.com/questions/492523/calculating-md5-of-string-from-microsoft-access
  Public Function MD5Hex(psTextString As String) As String
    Dim sOutStr As String
    Dim iPos    As Integer
    
    psTextString = MD5(psTextString)
    For iPos = 1 To Len(psTextString)
      sOutStr = sOutStr & LCase(Right("0" & Hex(AscB(Mid$(psTextString, iPos, 1))), 2))
    Next
    MD5Hex = sOutStr
  End Function
#End If

