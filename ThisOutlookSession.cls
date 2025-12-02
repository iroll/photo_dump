VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' --- ThisOutlookSession: email-only context menu hook ---
Option Explicit
Private Const PDUMP_TAG As String = "PDUMP_CTX_BTN"

Private Sub Application_ItemContextMenuDisplay( _
    ByVal CommandBar As Office.CommandBar, _
    ByVal Selection As Selection)

    On Error Resume Next

    ' Remove any prior instance so we don't duplicate
    DeleteControlIfExists CommandBar, PDUMP_TAG

    ' Only show if at least one MailItem is selected
    If CountMailInSelection_(Selection) = 0 Then Exit Sub

    Dim btn As Office.CommandBarButton
    Set btn = CommandBar.Controls.Add(Type:=msoControlButton, Temporary:=True)
    With btn
        .Caption = "Picture Dump"
        .Style = msoButtonIconAndCaption
        .FaceId = 108            ' built-in icon; swap if you prefer another
        .tag = PDUMP_TAG
        .OnAction = "PictureDump_EmailsOnly"   ' calls your simple macro
    End With
End Sub

' -- helpers --
Private Sub DeleteControlIfExists(ByVal CommandBar As Office.CommandBar, ByVal tag As String)
    On Error Resume Next
    Dim ctl As Office.CommandBarControl
    For Each ctl In CommandBar.Controls
        If ctl.tag = tag Then ctl.Delete
    Next
End Sub

Private Function CountMailInSelection_(sel As Selection) As Long
    On Error Resume Next
    Dim i As Long
    If sel Is Nothing Then Exit Function
    For i = 1 To sel.Count
        If TypeName(sel.Item(i)) = "MailItem" Then CountMailInSelection_ = CountMailInSelection_ + 1
    Next
End Function


