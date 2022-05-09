VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

Call WriteRegister
If Command = "" Then
    MsgBox "生成右键菜单成功", vbOKCancel, "提示"
End
End If
If IsString() Then
     Call WriteToTextFile(CreateTextFileName(Command), GetTextClipboard())
ElseIf IsClipboardFormatAvailable(CF_BITMAP) = 1 Then
    Call WriteBitmapToFile(CreateImageFileName(Command))
End If
End
End Sub


Private Function CreateTextFileName(ByVal sRoot As String)
CreateTextFileName = CreateFileName(sRoot, ".txt")
End Function
Private Function CreateImageFileName(ByVal sRoot As String)
CreateImageFileName = CreateFileName(sRoot, ".png")
End Function

Private Function CreateFileName(ByVal sRoot As String, ByVal sType As String) As String
Dim sName As String
sName = sRoot & "\" & Replace(Date, "/", "-") & "-" & Replace(Time, ":", "-")

If Dir(sName & ".txt") = "" Then
    CreateFileName = sName & sType
Else
    Dim iCount As Integer
    Do
        iCount = iCount + 1
    Loop While Dir(sName & "(" & iCount & ")" & sType) <> ""
    CreateFileName = sName & "(" & iCount & ")" & sType
End If

End Function
