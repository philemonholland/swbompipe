Option Explicit

Const AddInProgId = "AFCA.PipingBom.Generator"
Const AddInClassId = "{E1DB31B7-4F08-4B0D-99E4-69AFC34C5B1A}"
Const swSuccess = 0
Const swAddinAlreadyLoaded = 2

Dim swApp
Dim addinObject
Dim loadResult
Dim fso
Dim comHostPath
Dim attempt

On Error Resume Next

Set swApp = GetObject(, "SldWorks.Application")
If Err.Number <> 0 Then
    WScript.Quit 2
End If

Err.Clear
Set addinObject = swApp.GetAddInObject(AddInProgId)
If Err.Number = 0 Then
    If Not (addinObject Is Nothing) Then
        WScript.Quit 0
    End If
End If

Err.Clear
Set addinObject = swApp.GetAddInObject(AddInClassId)
If Err.Number = 0 Then
    If Not (addinObject Is Nothing) Then
        WScript.Quit 0
    End If
End If

Set fso = CreateObject("Scripting.FileSystemObject")
comHostPath = fso.BuildPath(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "SolidWorksBOMAddin"), "AFCA.PipingBom.Generator.comhost.dll")
If Not fso.FileExists(comHostPath) Then
    WScript.Quit 1
End If

Err.Clear
loadResult = swApp.LoadAddIn(comHostPath)
If Err.Number <> 0 Then
    WScript.Quit 1
End If

If loadResult <> swSuccess And loadResult <> swAddinAlreadyLoaded Then
    WScript.Quit 1
End If

For attempt = 1 To 10
    Err.Clear
    Set addinObject = swApp.GetAddInObject(AddInProgId)
    If Err.Number = 0 Then
        If Not (addinObject Is Nothing) Then
            WScript.Quit 0
        End If
    End If

    Err.Clear
    Set addinObject = swApp.GetAddInObject(AddInClassId)
    If Err.Number = 0 Then
        If Not (addinObject Is Nothing) Then
            WScript.Quit 0
        End If
    End If

    WScript.Sleep 1000
Next

WScript.Quit 1
