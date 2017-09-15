Attribute VB_Name = "Module_Main"
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (ByVal lpSystemName As String, ByVal lpAccountName As String, ByVal Sid As String, ByRef cbSid As Long, ByVal ReferencedDomainName As String, ByRef cbReferencedDomainName As Long, ByRef peUse As Long) As Long
Private Declare Function IsValidSid Lib "advapi32.dll" (ByRef pSid As Any) As Long
Private Declare Function ConvertSidToStringSid Lib "advapi32.dll" Alias "ConvertSidToStringSidA" (ByVal Sid As String, ByRef lpStringSid As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
 
Private Function get_current_computername() As String
  On Error Resume Next
  get_current_computername = ""
  
  Dim sBuffer As String
  sBuffer = String$(255, 0&)
  Call GetComputerName(sBuffer, 255)
  
  get_current_computername = Left$(sBuffer, InStr(vbNull, sBuffer, vbNullChar, vbBinaryCompare) - 1)

End Function


Private Function get_current_username() As String
  On Error Resume Next
  get_current_username = ""
  
  Dim sBuffer As String
  sBuffer = String$(255, 0&)
  Call GetUserName(sBuffer, 255)
  
  get_current_username = Left$(sBuffer, InStr(vbNull, sBuffer, vbNullChar, vbBinaryCompare) - 1)
End Function
 
Private Function get_sid(lpSystemName As String, lpAccountName As String) As String
  On Error Resume Next
  get_sid = ""

  Dim Sid As String
  Dim Domain As String
  Dim peUse As Long
  
  Dim sResult_DOMAIN As String
  Dim sResult_SID As String
   
  Sid = String$(255, 0&)
  Domain = String$(255, 0&)
  sResult_DOMAIN = String$(255, 0&)
  sResult_SID = String$(255, 0&)
  
  Call LookupAccountName(lpSystemName, lpAccountName, Sid, 255, Domain, 255, peUse)
  If (IsValidSid(ByVal Sid) = 0&) Then Exit Function
  
  Call ConvertSidToStringSid(Sid, peUse)
  Call CopyMemory(ByVal sResult_SID, ByVal peUse, 255)
  
  sResult_DOMAIN = Left$(Domain, InStr(vbNull, Domain, vbNullChar, vbBinaryCompare) - vbNull)
  sResult_SID = Left$(sResult_SID, InStr(vbNull, sResult_SID, vbNullChar, vbBinaryCompare) - vbNull)
  Call GlobalFree(peUse)
  
  get_sid = sResult_SID ' & vbNewLine & sResult_DOMAIN
End Function

Private Function get_sid_noarg()
  On Error Resume Next
  get_sid_noarg = ""

  Dim computer_name As String
  Dim username As String
  
  computer_name = get_current_computername()
  username = get_current_username()
  
  If computer_name = "" Then Exit Function
  If username = "" Then Exit Function

  get_sid_noarg = get_sid(computer_name, username)
End Function


Public Sub Main()
    WriteStdOut get_sid_noarg()
End Sub

