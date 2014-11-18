Attribute VB_Name = "mProtect"
'Pink/Danyfirex
'All Credits <!-- m --><a class="postlink" href="http://waleedassar.blogspot.com/2013/02/kernel-bug-1-processiopriority.html" onclick="window.open(this.href);return false;">http://waleedassar.blogspot.com/2013/02 ... ority.html</a><!-- m -->
 
 
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function ZwSetInformationProcess Lib "ntdll.dll" (ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long
 
 
 
Public Sub Protect()
ZwSetInformationProcess GetCurrentProcess(), &H21&, VarPtr(&H8000F129), &H4&
End Sub