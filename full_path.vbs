' Copy the first argument to clipboard

Option Explicit

Dim arg, i
Set arg = WScript.Arguments

If arg.Count > 0 Then
   SetClip arg(0)
End If


'
' Copy 'str' to clipboard.  It can't contain vbCrLf in str.
'
Sub SetClip(str)
    Dim cmd, objShell

    'strA= "some character string"
    Set objShell = WScript.CreateObject("WScript.Shell")
    objShell.Run "cmd /C echo . | set /p x=" & str & "| clip", 2
End Sub

'
' Copy text from clipboard.
'
Function GetClip
    GetClip = createobject("htmlfile").parentwindow.clipboarddata.getdata("text")
End Function
