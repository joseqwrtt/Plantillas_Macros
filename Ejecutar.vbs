Option Explicit

Dim objShell, objFSO, pythonCheck, installerURL, installerFile, ret, ajustesFolder, plantillasFolder
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' --------------------------
' COMPROBAR SI PYTHON ESTÁ INSTALADO
' --------------------------
On Error Resume Next
pythonCheck = objShell.Exec("cmd /c where python").StdOut.ReadLine
On Error GoTo 0

If pythonCheck = "" Then
    MsgBox "Python no encontrado. Se procederá a instalarlo automáticamente.", vbInformation, "Instalador de Python"

    installerURL = "https://www.python.org/ftp/python/3.13.2/python-3.13.2-amd64.exe"
    installerFile = objFSO.GetAbsolutePathName(".") & "\python-installer.exe"

    ' Descargar el instalador de Python
    DownloadFile installerURL, installerFile

    ' Ejecutar instalación silenciosa con Python agregado al PATH
    ret = objShell.Run("cmd /c title Instalando Python... && """ & installerFile & """ /quiet InstallAllUsers=0 PrependPath=1 Include_test=0", 1, True)

    ' Comprobar si se instaló correctamente
    On Error Resume Next
    pythonCheck = objShell.Exec("cmd /c where python").StdOut.ReadLine
    On Error GoTo 0

    If pythonCheck = "" Then
        MsgBox "No se pudo instalar Python automáticamente. Instálalo manualmente desde python.org.", vbCritical, "Error"
        WScript.Quit
    End If

    ' Borrar instalador
    If objFSO.FileExists(installerFile) Then objFSO.DeleteFile installerFile
End If

' --------------------------
' CREAR CARPETA "Ajustes" SI NO EXISTE
' --------------------------
ajustesFolder = objFSO.GetAbsolutePathName(".") & "\Ajustes"
If Not objFSO.FolderExists(ajustesFolder) Then
    objFSO.CreateFolder ajustesFolder
End If

' --------------------------
' CREAR CARPETA "Plantillas" SI NO EXISTE
' --------------------------
plantillasFolder = objFSO.GetAbsolutePathName(".") & "\Plantillas"
If Not objFSO.FolderExists(plantillasFolder) Then
    objFSO.CreateFolder plantillasFolder
End If

' --------------------------
' CREAR ARCHIVO config.json SI NO EXISTE
' --------------------------
Dim configFilePath, configFile
configFilePath = ajustesFolder & "\config.json"

If Not objFSO.FileExists(configFilePath) Then
    Set configFile = objFSO.CreateTextFile(configFilePath, True)
    configFile.WriteLine "{""modo_oscuro"": true}"
    configFile.Close
End If

' --------------------------
' CREAR ARCHIVO primer_uso.json SI NO EXISTE
' --------------------------
Dim primerUsoFile
primerUsoFile = ajustesFolder & "\primer_uso.json"

If Not objFSO.FileExists(primerUsoFile) Then
    ' Crear script Python temporal para mostrar ventana FAQ interactiva
    Dim faqPyFile, file
    faqPyFile = ajustesFolder & "\faq_temp.py"
    Set file = objFSO.CreateTextFile(faqPyFile, True)

    file.WriteLine "import tkinter as tk"
    file.WriteLine "from tkinter import scrolledtext"
    file.WriteLine "import os"
    file.WriteLine ""
    file.WriteLine "root = tk.Tk()"
    file.WriteLine "root.title('Bienvenido - Guía de uso')"
    file.WriteLine "root.geometry('650x500')"
    file.WriteLine ""
    file.WriteLine "txt = scrolledtext.ScrolledText(root, wrap=tk.WORD, font=('Segoe UI', 10))"
    file.WriteLine "txt.pack(expand=True, fill='both', padx=10, pady=10)"
    file.WriteLine ""
    file.WriteLine "faq_text = '''🎉 Bienvenido a la aplicación de Plantillas\n\n" & _
                   "Esta aplicación permite copiar plantillas ODT y DOCX al portapapeles en formato HTML para pegar en emails o documentos compatibles.\n\n" & _
                   "📌 Instrucciones de uso:\n" & _
                   "1️⃣ Coloca tus plantillas en la carpeta 'Plantillas' dentro de la carpeta principal de la aplicación.\n" & _
                   "   - Solo se reconocerán archivos con extensión .ODT o .DOCX.\n" & _
                   "2️⃣ Haz clic en el botón de la plantilla que desees copiar al portapapeles.\n" & _
                   "3️⃣ Si agregas nuevas plantillas mientras la aplicación está abierta, pulsa 'Refrescar' para actualizar los botones.\n" & _
                   "4️⃣ Alterna entre modo claro y oscuro desde el menú ⚙️.\n" & _
                   "5️⃣ La posición de la ventana y el modo claro/oscuro se guardarán automáticamente.\n\n" & _
                   "🛠️ Depuración y logs:\n" & _
                   "- Los logs de errores y procesos se guardan en la carpeta 'Ajustes/log'.\n" & _
                   "- Si algo falla, revisa los logs para información detallada sobre el error.\n\n" & _
                   "💡 Buenas prácticas al crear Word/ODT:\n" & _
                   "- Usa estilos de párrafo para títulos, listas y texto normal.\n" & _
                   "- Para listas, usa bullets o numeración de Word/ODT (no guiones manuales).\n" & _
                   "- Para enlaces, usa la función de hipervínculo; se copiarán correctamente.\n" & _
                   "- Las imágenes se detectan siempre que estén insertadas dentro del documento, preferiblemente en línea con el texto.\n" & _
                   "- Evita formatos complejos como tablas anidadas o campos especiales que no sean texto, imágenes o enlaces.\n\n" & _
                   "💡 Consejos adicionales:\n" & _
                   "- Mantén nombres de plantillas claros y descriptivos.\n" & _
                   "- Organiza la carpeta 'Plantillas' para encontrar rápidamente lo que necesitas.\n" & _
                   "- Evita duplicados con nombres idénticos.\n\n" & _
                   "¡Disfruta de la aplicación y optimiza tu flujo de trabajo con plantillas!'''" 
    file.WriteLine "txt.insert(tk.END, faq_text)"
    file.WriteLine "txt.configure(state='disabled')"
    file.WriteLine ""
    file.WriteLine "def abrir_carpeta():" 
    file.WriteLine "    plantillas = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'Plantillas')"
    file.WriteLine "    if os.path.exists(plantillas):"
    file.WriteLine "        os.startfile(plantillas)"
    file.WriteLine ""
    file.WriteLine "btn = tk.Button(root, text='Abrir carpeta Plantillas', command=abrir_carpeta, font=('Segoe UI',10))"
    file.WriteLine "btn.pack(pady=8)"
    file.WriteLine ""
    file.WriteLine "root.mainloop()"

    file.Close

    ' Ejecutar script FAQ temporal y esperar que el usuario lo cierre
    ret = objShell.Run("python """ & faqPyFile & """", 1, True)

    ' Crear archivo primer_uso.json para marcar que ya se mostró
    Dim primerUsoObj
    Set primerUsoObj = objFSO.CreateTextFile(primerUsoFile, True)
    primerUsoObj.WriteLine "{""mostrado"": true}"
    primerUsoObj.Close

    ' Borrar script temporal FAQ
    If objFSO.FileExists(faqPyFile) Then objFSO.DeleteFile(faqPyFile)
End If

' --------------------------
' COMPROBAR FECHA DE ÚLTIMA ACTUALIZACIÓN DE DEPENDENCIAS
' -------------------------- '<--- MODIFICADO
' --------------------------
Dim depFilePath, depFile, lastRunDate, daysSinceLastRun, runDependencies
depFilePath = ajustesFolder & "\dependencias.json"
runDependencies = True ' Por defecto, ejecutar

If objFSO.FileExists(depFilePath) Then
    ' Leer fecha almacenada
    Dim depContent, jsonDate
    Set depFile = objFSO.OpenTextFile(depFilePath, 1)
    depContent = depFile.ReadAll
    depFile.Close

    ' Extraer fecha en formato yyyy-mm-dd
    jsonDate = Mid(depContent, InStr(depContent, ":") + 3, 10)
    lastRunDate = CDate(jsonDate)
    
    ' Calcular días desde la última ejecución
    daysSinceLastRun = DateDiff("d", lastRunDate, Date)
    
    If daysSinceLastRun < 30 Then
        runDependencies = False
    End If
End If

If runDependencies Then
    ' --------------------------
    ' CREAR ARCHIVO PYTHON TEMPORAL PARA INSTALAR DEPENDENCIAS
    ' --------------------------
    Dim depPyFile
    depPyFile = ajustesFolder & "\instalar_dependencias_temp.py"

    Set file = objFSO.CreateTextFile(depPyFile, True)
    file.WriteLine "import sys"
    file.WriteLine "import subprocess"
    file.WriteLine "deps = ['pywin32', 'python-docx', 'win10toast']"
    file.WriteLine "print('=== Instalando / actualizando dependencias ===')"
    file.WriteLine "for d in deps:"
    file.WriteLine "    print(f'📦 Actualizando {d} ...')"
    file.WriteLine "    subprocess.run([sys.executable, '-m', 'pip', 'install', '--upgrade', d], check=False)"
    file.WriteLine "print('✅ Dependencias actualizadas correctamente.')"
    file.Close

    ' --------------------------
    ' EJECUTAR SCRIPT DE DEPENDENCIAS
    ' --------------------------
    ret = objShell.Run("cmd /c title Instalando dependencias... && python """ & depPyFile & """", 1, True)

    ' Borrar script temporal
    If objFSO.FileExists(depPyFile) Then objFSO.DeleteFile depPyFile

    ' --------------------------
    ' GUARDAR FECHA DE LA ÚLTIMA EJECUCIÓN
    ' --------------------------
    Set depFile = objFSO.CreateTextFile(depFilePath, True)
    depFile.WriteLine "{""ultima_ejecucion"": """ & FormatDateTime(Date, vbShortDate) & """}"
    depFile.Close
Else
'    MsgBox "✅ Dependencias ya actualizadas en los últimos 30 días. No se ejecuta de nuevo.", vbInformation, "Info"
End If

' --------------------------
' EJECUTAR SCRIPT PRINCIPAL CopiarPlantillasWord_ODT_3.0.py (oculto)
' --------------------------
Dim mainScript
mainScript = objFSO.GetAbsolutePathName("CopiarPlantillasWord_ODT_3.0.py")
If objFSO.FileExists(mainScript) Then
    ret = objShell.Run("python """ & mainScript & """", 0, True)
Else
    MsgBox "❌ No se encontró el archivo CopiarPlantillasWord_ODT_3.0.py en esta carpeta.", vbCritical, "Error"
End If

' --------------------------
' FUNCIÓN PARA DESCARGAR ARCHIVOS
' --------------------------
Sub DownloadFile(URL, LocalFile)
    Dim objHTTP, objStream
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    objHTTP.Open "GET", URL, False
    objHTTP.Send
    If objHTTP.Status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1 ' Binario
        objStream.Open
        objStream.Write objHTTP.ResponseBody
        objStream.SaveToFile LocalFile, 2 ' Sobrescribir
        objStream.Close
    Else
        MsgBox "Error al descargar " & URL & vbCrLf & "Código: " & objHTTP.Status, vbCritical, "Error"
        WScript.Quit
    End If
End Sub
