Attribute VB_Name = "Module2"
Function WriteToTextFile(sPath As String, sData As String)
    Open sPath For Output As #1
    Print #1, sData
    Close #1
End Function
Public Sub WriteRegister()
Dim sAppPath As String

sAppPath = App.Path & "\" & App.EXEName & ".exe"

Dim reg As New Regidit
reg.CreateKey "HKEY_CURRENT_USER\Software\Classes\Directory\Background\Shell\ºÙ«–∞Â", "command"

reg.SetKeyValueREG_SZ "HKEY_CURRENT_USER\Software\Classes\Directory\Background\Shell\ºÙ«–∞Â\command", "", sAppPath & " %V"
End Sub
