"use strict";

if (WScript.Arguments.length !== 2) {
  WScript.Echo('Please drop two Excel files.');
  WScript.Quit();
}

var shell = WScript.CreateObject('WScript.Shell');
var fso = WScript.CreateObject('Scripting.FileSystemObject');

shell.CurrentDirectory = fso.GetParentFolderName(WScript.ScriptFullName);

var temporaryFile = fso.GetTempName();

var file = fso.CreateTextFile(temporaryFile, true);
file.WriteLine(WScript.Arguments(0));
file.WriteLine(WScript.Arguments(1));
file.Close();

// If AppVLP.exe (Microsoft Application Virtualization) is used.
shell.Run('"C:\\Program Files\\Microsoft Office\\root\\Client\\AppVLP.exe" "C:\\Program Files (x86)\\Microsoft Office\\Office16\\DCF\\SPREADSHEETCOMPARE.EXE"' + temporaryFile);

// If not AppVLP.exe (Microsoft Application Virtualization) is used.
shell.Run('"C:\\Program Files (x86)\\Microsoft Office\\Office16\\DCF\\SPREADSHEETCOMPARE.EXE"' + temporaryFile);