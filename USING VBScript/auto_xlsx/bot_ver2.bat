@echo off 
:minimized
if not "%minimized%"=="" goto :minimized
set minimized=true
start /min cmd /C "%~dpnx0"
goto :EOF
:minimized
::---------------------------------------
::-Get & Store {main file{dir}}
set "main_file=%~dp0"
::---------------------------------------
taskkill /IM cscript.exe /F
::del main.vbs
if exist main.vbs del main.vbs
if exist Bot_ver1.bat del Bot_ver1.bat
if exist %temp%\main.vbs del %temp%\main.vbs
cls
::---------------------------------------------
::-Create VBScript
echo Option Explicit >> %temp%\main.vbs
echo Dim fso, dataFile, file, data, line, objExcel, objWorkbook, objWorksheet, lngColumn, lngRow, shell, blnEmptyLine, stream, textdata, WshShell, intButton >> %temp%\main.vbs
echo Set fso = CreateObject("Scripting.FileSystemObject") >> %temp%\main.vbs
echo Set stream = CreateObject("ADODB.Stream") >> %temp%\main.vbs
echo Set objExcel = CreateObject("Excel.Application") >> %temp%\main.vbs
echo Set shell = CreateObject("WScript.Shell") >> %temp%\main.vbs
echo Do While True >> %temp%\main.vbs
echo    If Not fso.FileExists("%main_file%data.txt") Then >> %temp%\main.vbs
echo        Set WshShell = WScript.CreateObject("WScript.Shell") >> %temp%\main.vbs
echo        intButton = WshShell.Popup("%main_file%data.txt does not exist.", 1.5, "Alert", 48) >> %temp%\main.vbs
echo        WScript.Quit >> %temp%\main.vbs
echo    End If >> %temp%\main.vbs
echo    If fso.FileExists("%main_file%data.txt") And fso.GetFile("%main_file%data.txt").Size > 0 Then >> %temp%\main.vbs
echo        Set dataFile = fso.OpenTextFile("%main_file%data.txt", 1) >> %temp%\main.vbs
echo        data = dataFile.ReadAll >> %temp%\main.vbs
echo        dataFile.Close >> %temp%\main.vbs
echo        If fso.FileExists("%main_file%backup_data.txt") Then >> %temp%\main.vbs
echo            fso.DeleteFile("%main_file%backup_data.txt") >> %temp%\main.vbs
echo        End If >> %temp%\main.vbs
echo        Set file = fso.CreateTextFile("%main_file%backup_data.txt", True) >> %temp%\main.vbs
echo        file.Write data >> %temp%\main.vbs
echo        file.Close >> %temp%\main.vbs
echo        If fso.FileExists(fso.GetAbsolutePathName("%main_file%data.xlsx")) Then >> %temp%\main.vbs
echo            Set objWorkbook = objExcel.Workbooks.Open(fso.GetAbsolutePathName("%main_file%data.xlsx")) >> %temp%\main.vbs
echo        Else >> %temp%\main.vbs
echo            Set objWorkbook = objExcel.Workbooks.Add >> %temp%\main.vbs
echo        End If >> %temp%\main.vbs
echo        Set objWorksheet = objWorkbook.Worksheets(1) >> %temp%\main.vbs
echo        stream.Open >> %temp%\main.vbs
echo        stream.Type = 2 'Specify stream type - we want To save text/string data. >> %temp%\main.vbs
echo        stream.Charset = "utf-8" 'Specify charset For the source text data. >> %temp%\main.vbs
echo        stream.LoadFromFile "%main_file%data.txt" 'Load the data from the file >> %temp%\main.vbs
echo        textdata = Split(stream.ReadText, vbCrLf) 'Read the text data >> %temp%\main.vbs
echo        stream.Close >> %temp%\main.vbs
echo        lngColumn = 1 >> %temp%\main.vbs
echo        lngRow = objWorksheet.Cells(objWorksheet.Rows.Count, 1).End(-4162).Row + 1 >> %temp%\main.vbs
echo        blnEmptyLine = False >> %temp%\main.vbs
echo        For Each line In textdata >> %temp%\main.vbs
echo            If Trim(line) = "" Then >> %temp%\main.vbs
echo                If blnEmptyLine = False Then >> %temp%\main.vbs
echo                    lngRow = lngRow + 1 >> %temp%\main.vbs
echo                    lngColumn = 1 >> %temp%\main.vbs
echo                    blnEmptyLine = True >> %temp%\main.vbs
echo                End If >> %temp%\main.vbs
echo            Else >> %temp%\main.vbs
echo                objWorksheet.Cells(lngRow, lngColumn).Value = line >> %temp%\main.vbs
echo                lngColumn = lngColumn + 1 >> %temp%\main.vbs
echo                blnEmptyLine = False >> %temp%\main.vbs
echo            End If >> %temp%\main.vbs
echo        Next >> %temp%\main.vbs
echo        Dim fileToClear >> %temp%\main.vbs
echo        Set fileToClear = fso.OpenTextFile("%main_file%data.txt", 2) >> %temp%\main.vbs
echo        fileToClear.Write "" >> %temp%\main.vbs
echo        fileToClear.Close >> %temp%\main.vbs
echo        If fso.FileExists(fso.GetAbsolutePathName("%main_file%data.xlsx")) Then >> %temp%\main.vbs
echo            objWorkbook.Save >> %temp%\main.vbs
echo        Else >> %temp%\main.vbs
echo            objWorkbook.SaveAs fso.GetAbsolutePathName("%main_file%data.xlsx"), 51 ' 51 corresponds to xlsx format >> %temp%\main.vbs
echo        End If >> %temp%\main.vbs
echo        objWorkbook.Close >> %temp%\main.vbs
echo        Set objWorksheet = Nothing >> %temp%\main.vbs
echo        Set objWorkbook = Nothing >> %temp%\main.vbs
echo        shell.Run "powershell -c ""[console]::beep(500, 300)""", 0 >> %temp%\main.vbs
echo    End If >> %temp%\main.vbs
echo Loop >> %temp%\main.vbs
echo Set fso = Nothing >> %temp%\main.vbs
echo Set objExcel = Nothing >> %temp%\main.vbs
echo Set shell = Nothing >> %temp%\main.vbs
::run the script
cls
cscript %temp%\main.vbs