'' Launcher do monitoramento de relatórios - roda em segundo plano (sem janela)
'' Caminho automático: funciona no servidor e na máquina local

Dim WshShell, fso, pastaScript, batPath

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

'' Detecta a pasta onde este .vbs está (funciona em qualquer caminho)
pastaScript = fso.GetParentFolderName(WScript.ScriptFullName)
batPath = pastaScript & "\executar_loop.bat"

If Not fso.FileExists(batPath) Then
    MsgBox "ERRO: executar_loop.bat não encontrado em:" & vbCrLf & pastaScript, vbCritical, "Relatorio Andre"
    WScript.Quit 1
End If

'' 0 = sem janela visível, False = não bloqueia (roda em background)
WshShell.Run "cmd.exe /c """ & batPath & """", 0, False

Set WshShell = Nothing
Set fso = Nothing
