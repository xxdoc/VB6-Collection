'---------------------------------------------------------------------------------------
' Modulo   : mUACForcer
' Autor     : chequinho
' Creditos : Slek, rudeboy1991
' Finalidad : Se ejecuta a si mismo como administrador hasta que el usuario de clic en "Si"
' Uso       : Call RunAsAdmin
' Notas    :
'  - Necesaria referencia a Microsoft Scripting Runtime (scrrun.dll)
'---------------------------------------------------------------------------------------


'SHELL32
Public Declare Function IsUserAnAdmin Lib "SHELL32" () As Long


Public Sub RunAsAdmin()
    On Error GoTo Err
        If IsUserAnAdmin = 0 Then
            MsgBox "No soy admin D:"
            Dim numberOfMe As Integer
            numberOfMe = getNumberOfProcess(App.EXEName & ".exe")
            Set objShell = CreateObject("Shell.Application")
            objShell.ShellExecute App.Path & "\" & App.EXEName & ".exe", "", "", "runas", 0
            Set objShell = Nothing
            While getNumberOfProcess("consent.exe") > 0
                'No hacer nada
           Wend
            If Not getNumberOfProcess(App.EXEName & ".exe") > numberOfMe Then
                Call RunAsAdmin
            Else
                End
            End If
        Else
            MsgBox "Soy admin :B"
        End If
        Exit Sub
Err:
End Sub


Private Function getNumberOfProcess(ByVal Process As String) As Integer
    Dim objWMIService, colProcesses
    Set objWMIService = GetObject("winmgmts:")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name='" & Process & "'")
    getNumberOfProcess = colProcesses.Count
End Function