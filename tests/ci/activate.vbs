' Cree l'objet COM du complement comme le fait Word (client COM natif), puis demande le ruban.
' Usage : cscript //nologo activate.vbs <fichier resultat>
Option Explicit
Dim outFile, addin, ribbon, result, fso, f
outFile = WScript.Arguments(0)
On Error Resume Next
Set addin = CreateObject("WordTableToExcel.Connect")
If Err.Number <> 0 Then
    result = "ECHEC CreateObject : 0x" & Hex(Err.Number) & " " & Err.Description
Else
    ribbon = addin.GetCustomUI("Microsoft.Word.Document")
    If Err.Number <> 0 Then
        result = "ECHEC GetCustomUI : 0x" & Hex(Err.Number) & " " & Err.Description
    ElseIf InStr(ribbon, "OnExportClick") > 0 And InStr(ribbon, "OnImportClick") > 0 Then
        result = "OK"
    Else
        result = "ECHEC ruban : " & ribbon
    End If
End If
Err.Clear
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile(outFile, True)
f.Write result
f.Close
