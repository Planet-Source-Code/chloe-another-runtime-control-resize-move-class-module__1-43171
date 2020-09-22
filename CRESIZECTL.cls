VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRESIZECTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const SWP_DRAWFRAME As Long = &H20
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FLAGS As Long = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
Private Const GWL_STYLE As Long = (-16)
Private Const WS_THICKFRAME As Long = &H40000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private mOBJ As Control
Attribute mOBJ.VB_VarHelpID = -1
Private mStyle As Long

'*************************************************
' Add more to the logic below for more controls
' as needed
'*************************************************
Private WithEvents mtxt As TextBox
Attribute mtxt.VB_VarHelpID = -1
Private WithEvents mcmd As CommandButton
Attribute mcmd.VB_VarHelpID = -1
Private WithEvents mpic As PictureBox
Attribute mpic.VB_VarHelpID = -1
Private Sub CheckType()
    On Error GoTo ErrorCheckType
    Set mtxt = Nothing
    Set mcmd = Nothing
    Set mpic = Nothing
    If Not mOBJ Is Nothing Then
        If TypeOf mOBJ Is TextBox Then
            Set mtxt = mOBJ
        ElseIf TypeOf mOBJ Is CommandButton Then
            Set mcmd = mOBJ
        ElseIf TypeOf mOBJ Is PictureBox Then
            Set mpic = mOBJ
        End If
    End If
    Exit Sub
ErrorCheckType:
    MsgBox Err & ":Error in call to CheckType()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub mtxt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage mtxt.hwnd, &HA1, 2, 0&
End Sub
Private Sub mcmd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage mcmd.hwnd, &HA1, 2, 0&
End Sub
Private Sub mpic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage mpic.hwnd, &HA1, 2, 0&
End Sub


'*************************************************
' Code below is for resizing
'*************************************************
Public Sub AllowResize(ctl As Object)
    On Error GoTo ErrorAllowResize
    Dim Style As Long
    If Not mOBJ Is Nothing Then
        If mOBJ.Name = ctl.Name And mOBJ.Parent.Name = ctl.Parent.Name Then Exit Sub
        DisallowResize
    End If
     Style = GetWindowLong(ctl.hwnd, GWL_STYLE)
     mStyle = Style
     Style = Style Or WS_THICKFRAME
     SetControlStyle Style, ctl
     ctl.ZOrder 0
     Set mOBJ = ctl
     CheckType
    Exit Sub
ErrorAllowResize:
    MsgBox Err & ":Error in call to AllowResize()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Public Sub DisallowResize()
    On Error GoTo ErrorDisallowResize
    If Not mOBJ Is Nothing Then
        SetControlStyle mStyle, mOBJ
        Set mOBJ = Nothing
        CheckType
    End If
    Exit Sub
ErrorDisallowResize:
    MsgBox Err & ":Error in call to DisallowResize()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub SetControlStyle(Style, X As Control)
    If Style Then
        Call SetWindowLong(X.hwnd, GWL_STYLE, Style)
        Call SetWindowPos(X.hwnd, X.Parent.hwnd, 0, 0, 0, 0, SWP_FLAGS)
    End If
End Sub
Private Sub Class_Terminate()
    On Error Resume Next
    DisallowResize
End Sub
