Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
basePath = fso.GetParentFolderName(WScript.ScriptFullName)
ajustesPath = basePath & "\ajustes"
dependenciasMarker = ajustesPath & "\ult_ejec_dependencias.txt"

' Comprobar si instalar_dependencias.py ya se ejecutó hoy
ejecutarDependencias = True

If fso.FileExists(dependenciasMarker) Then
    Set file = fso.OpenTextFile(dependenciasMarker, 1)
    fechaUltEjec = CDate(file.ReadLine)
    file.Close

    ' Comparar solo la fecha, ignorando la hora
    If DateValue(fechaUltEjec) = Date Then
        ejecutarDependencias = False
    End If
End If

' Ejecutar script de dependencias solo si es necesario
If ejecutarDependencias Then
    shell.Run "cmd /c python """ & ajustesPath & "\instalar_dependencias.py""", 1, True

    ' Actualizar el archivo marcador con la fecha actual
    Set file = fso.CreateTextFile(dependenciasMarker, True)
    file.WriteLine Now
    file.Close
End If

' --- Comprobar si el servidor ya está corriendo ---
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
    shell.Run "cmd /c start /min ""FlaskServer"" pythonw """ & basePath & "\app.py""", 0, False
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
