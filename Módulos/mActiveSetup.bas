Attribute VB_Name = "ActiveSetup"
Option Explicit

Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Const HKEY_LOCAL_MACHINE            As Long = &H80000002
Private Const HKEY_CURRENT_USER             As Long = &H80000001
Private Const REG_SZ                        As Long = 1
Private Const KEY_QUERY_VALUE               As Long = &H1
Private Const KEY_SET_VALUE                 As Long = &H2
Private Const KEY_ENUMERATE_SUB_KEYS        As Long = &H8
Private Const ERROR_NO_MORE_ITEMS           As Long = 259&

Private Const InstalledKey As String = "Software\Microsoft\Active Setup\Installed Components"


Private Function RandomKey() As String
    Dim i As Long
    Dim sLong As String
    
    RandomKey = "{"
    Randomize
xx:
    sLong = Hex(Rnd * 65536)
    Do While Len(sLong) < 4:        sLong = "0" & sLong:        Loop
    RandomKey = RandomKey & sLong
    
    i = Len(RandomKey)
    Select Case (i)
        Case 5, 29, 33:                                         GoTo xx
        Case 9, 14, 19, 24:     RandomKey = RandomKey & "-":    GoTo xx
        Case 37:                RandomKey = RandomKey & "}"
    End Select
End Function



Public Function ActiveStartUp(ByVal sPath As String) As Boolean
    Dim hKey As Long
    Dim hKey2 As Long
    Dim index As Long
    Dim bufKey As String
    Dim bufValue As String
    Dim bufSize As Long
    Dim ActiveKeyName As String
    
    'Se espera un poco para que windows escriba la clave en hkcu en caso de estar reiniciando la aplicación
    Sleep (200)
    'Sacamos un handle de la clave "Installed Components" de HKLM
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, InstalledKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey) Then Exit Function
    'y buscamos la ruta de nuestro exe entre las subclaves
    bufKey = Space(255): bufValue = Space(255): bufSize = 255
    Do While RegEnumKeyEx(hKey, index, bufKey, bufSize, ByVal 0, vbNullString, ByVal 0, 0) <> ERROR_NO_MORE_ITEMS
        bufKey = Left$(bufKey, bufSize)
        bufSize = 255
        If RegOpenKeyEx(hKey, bufKey, 0, KEY_QUERY_VALUE, hKey2) = 0 Then
            If RegQueryValueEx(hKey2, "StubPath", 0, REG_SZ, ByVal bufValue, bufSize) = 0 Then
                bufValue = Left$(bufValue, bufSize - 1)
                If LCase$(sPath) = LCase$(bufValue) Then
                    'si encontramos nuestra subclave se guarda y dejamos de buscar
                    ActiveKeyName = InstalledKey & "\" & bufKey
                    Exit Do
                End If
            End If
        End If
        RegCloseKey hKey2
        index = index + 1
        bufKey = Space(255): bufValue = Space(255): bufSize = 255
        DoEvents
    Loop
    RegCloseKey hKey
    RegCloseKey hKey2
    
    'si no existe subclave con nuestro proceso en hklm creamos una y salimos
    If ActiveKeyName = vbNullString Then
        ActiveKeyName = InstalledKey & "\" & RandomKey
        If RegCreateKeyEx(HKEY_LOCAL_MACHINE, ActiveKeyName, 0, vbNullString, 0, KEY_SET_VALUE, 0, hKey, 0) = 0 Then
            sPath = sPath & Chr(0)
            If RegSetValueEx(hKey, "StubPath", 0, REG_SZ, ByVal sPath, Len(sPath)) = 0 Then
                ActiveStartUp = True
            End If
            RegCloseKey hKey
        End If
        Exit Function
    End If
    
    'si existía subclave en hklm, comprobamos si existe en hkcu y la borramos
    If RegOpenKeyEx(HKEY_CURRENT_USER, ActiveKeyName, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        RegCloseKey hKey
        If RegDeleteKey(HKEY_CURRENT_USER, ActiveKeyName) = 0 Then
            ActiveStartUp = True
            RegCloseKey hKey
        End If
    Else        'si no existe en HKCU hay que reiniciarse
        ShellExecute 0, "open", sPath, vbNullString, vbNullString, 0
        End
    End If
    
End Function




