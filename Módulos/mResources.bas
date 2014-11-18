Attribute VB_Name = "mResources"
'mAddResource By Slek Thx to Cobein
'Indetectables.net
'15/12/2011
'Ex: bRet = AddResource("C:\1.exe", RT_RCDATA, 101, Buff())
Option Explicit

Public Enum RT

    RT_CURSOR = 1&
    RT_BITMAP = 2&
    RT_ICON = 3&
    RT_MENU = 4&
    RT_DIALOG = 5&
    RT_STRING = 6&
    RT_FONTDIR = 7&
    RT_FONT = 8&
    RT_ACCELERATOR = 9&
    RT_RCDATA = 10&
    RT_MESSAGETABLE = 11&
    RT_GROUP_CURSOR = 12&
    RT_GROUP_ICON = 14&
    RT_VERSION = 16&
    RT_DLGINCLUDE = 17&
    RT_PLUGPLAY = 19&
    RT_VXD = 20&
    RT_ANICURSOR = 21&
    RT_ANIICON = 22&
    RT_HTML = 23&
    RT_MANIFEST = 24&

End Enum

Private Const PADDING As String = "PADDINGXXPADDING"

'Kernel32.Dll
Private Declare Function BeginUpdateResource _
                Lib "KERNEL32" _
                Alias "BeginUpdateResourceA" (ByVal pFileName As String, _
                                              ByVal bDeleteExistingResources As Long) As Long

Private Declare Function EndUpdateResource _
                Lib "KERNEL32" _
                Alias "EndUpdateResourceA" (ByVal hUpdate As Long, _
                                            ByVal fDiscard As Long) As Boolean

Private Declare Function UpdateResource _
                Lib "KERNEL32" _
                Alias "UpdateResourceA" (ByVal hUpdate As Long, _
                                         ByVal lpType As Long, _
                                         ByVal lpName As Long, _
                                         ByVal wLanguage As Long, _
                                         lpData As Any, _
                                         ByVal cbData As Long) As Boolean
										 
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
				(hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)										 

'Version.Dll
Private Declare Function GetFileVersionInfo _
                Lib "Version.dll" _
                Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, _
                                             ByVal dwhandle As Long, _
                                             ByVal dwlen As Long, _
                                             lpData As Any) As Long

Private Declare Function GetFileVersionInfoSize _
                Lib "Version.dll" _
                Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
                                                 lpdwHandle As Long) As Long

Private Declare Function VerQueryValue _
                Lib "Version.dll" _
                Alias "VerQueryValueA" (pBlock As Any, _
                                        ByVal lpSubBlock As String, _
                                        lplpBuffer As Any, _
                                        puLen As Long) As Long

Public Function AddResource(ByVal sFileName As String, _
                            ByVal lType As RT, _
                            ByVal lID As Long, _
                            ByRef Buff() As Byte, _
                            Optional bRepalcePadd As Boolean = True) As Boolean
    Dim lUpdate As Long
    Dim lLangId As Long
    lLangId = GetLangID(sFileName)
    'If Not lLangId = 0 Then
    lUpdate = BeginUpdateResource(sFileName, False)

    If Not lUpdate = 0 Then
        If UpdateResource(lUpdate, lType, lID, lLangId, Buff(0), UBound(Buff) + 1) Then
            If EndUpdateResource(lUpdate, False) Then
                If bRepalcePadd Then Call ReplacePadd(sFileName)
                AddResource = True
                Exit Function

            End If

        End If

        Call EndUpdateResource(lUpdate, True)

    End If

    'End If
End Function

Private Function GetLangID(ByVal sFileName As String) As Long 'By Cobein
    Dim lLen        As Long
    Dim lHandle     As Long
    Dim bvBuffer()  As Byte
    Dim lVerPointer As Long
    Dim iVal        As Integer
    lLen = GetFileVersionInfoSize(sFileName, lHandle)

    If Not lLen = 0 Then
        ReDim bvBuffer(lLen)

        If Not GetFileVersionInfo(sFileName, 0&, lLen, bvBuffer(0)) = 0 Then
            If Not VerQueryValue(bvBuffer(0), "\VarFileInfo\Translation", lVerPointer, lLen) = 0 Then
                CopyMemory iVal, ByVal lVerPointer, 2
                GetLangID = iVal

            End If

        End If

    End If

End Function

Public Sub ReplacePadd(ByVal sFileName As String) 'By Cobein
    Dim iFile    As Integer
    Dim sBuff    As String
    Dim sReplace As String
    sReplace = String$(Len(PADDING), vbNullChar)
    iFile = FreeFile
    Open sFileName For Binary Access Read Write As iFile
    sBuff = Space$(LOF(iFile))
    Get iFile, , sBuff
    sBuff = Replace$(sBuff, PADDING, sReplace)
    Put iFile, 1, sBuff
    Close iFile

End Sub
