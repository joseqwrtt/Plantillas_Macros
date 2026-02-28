Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
basePath = fso.GetParentFolderName(WScript.ScriptFullName)
ajustesPath = basePath & "\ajustes"

' Instalar dependencias
shell.Run "cmd /c python """ & ajustesPath & "\instalar_dependencias.py""", 1, True

' Comprobar si el servidor ya está corriendo
Set wmi = GetObject("winmgmts:\\.\root\cimv2")
query = "SELECT * FROM Win32_Process WHERE Name='python.exe' OR Name='pythonw.exe'"
Set processes = wmi.ExecQuery(query)

serverRunning = False
For Each process In processes
    If InStr(LCase(process.CommandLine), "app.py") > 0 Then
        serverRunning = True
        Exit For
    End If
Next

' Iniciar servidor solo si no está corriendo
If Not serverRunning Then
    ' Usamos cmd /c start para ejecutar python en segundo plano correctamente
    shell.Run "cmd /c start ""FlaskServer"" python """ & basePath & "\app.py""", 0, False
End If

' Esperar a que el servidor Flask esté listo
serverReady = False
Do Until serverReady
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "http://127.0.0.1:5000", False
    http.Send
    If Err.Number = 0 And http.Status = 200 Then
        serverReady = True
    Else
        WScript.Sleep 500 ' esperar medio segundo antes de reintentar
        Err.Clear
    End If
    On Error GoTo 0
Loop

' Abrir navegador
shell.Run "http://127.0.0.1:5000"
