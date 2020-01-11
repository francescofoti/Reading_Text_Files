Attribute VB_Name = "MSupport"
Option Compare Database
Option Explicit

Public Sub DumpStringBytes(ByVal psText As String)
  Dim i           As Integer
  Dim abBytes()   As Byte
  Dim sOut        As String
  
  abBytes = psText
  For i = 0 To UBound(abBytes) / 2
    Debug.Print Format$(i + 1, "00000"); "|";
  Next i
  Debug.Print
  
  For i = 0 To UBound(abBytes)
    Debug.Print Format$(i + 1, "00"); "|";
  Next i
  Debug.Print
  
  For i = 0 To UBound(abBytes)
    Debug.Print "--+";
  Next i
  Debug.Print
  
  For i = 0 To UBound(abBytes)
    sOut = Hex$(abBytes(i))
    If Len(sOut) < 2 Then sOut = " " & sOut
    Debug.Print sOut; "|";
  Next i
  Debug.Print
  
  For i = 0 To UBound(abBytes)
    Debug.Print "--+";
  Next i
  Debug.Print
    
  For i = 0 To UBound(abBytes)
    sOut = Chr$(abBytes(i))
    If Len(sOut) < 1 Then sOut = "?"
    Debug.Print sOut; " |";
  Next i
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
