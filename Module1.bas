' VBA - BOM 2 pcbnew selector
' 
' Notes :
' ensure "Do not warp mouse pointer" in pcbnew find box is UNCHECKED.

Option Explicit

Public Const ViewOpenGL = True ' don't change this


Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal HWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal HWnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal HWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal HWnd As Long) As Boolean
Private Declare Function SetForegroundWindow Lib "user32" (ByVal HWnd As Long) As Long
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4

Private Const GW_HWNDNEXT = 2

Private working As Boolean
Private ref_idx As Integer


Private Sub Main()
    Application.EnableEvents = True
    Application.OnKey "p", "comp_placed"
    Application.OnKey "u", "comp_unplaced"
    Application.OnKey ",", "comp_prev"
    Application.OnKey ".", "comp_next"
    working = False
    ref_idx = 0
End Sub



Private Sub pcbnew_find(ByVal ref As String)
    Dim selflhWndP As Long
    Dim lhWndP As Long
    
    If GetHandleFromPartialCaption(lhWndP, "Pcbnew") = True _
    And GetHandleFromPartialCaption(selflhWndP, "Microsoft Excel") Then
        'If IsWindowVisible(lhWndP) = True Then
        '  MsgBox "Found VISIBLE Window Handle: " & lhWndP, vbOKOnly + vbInformation
        'Else
        '  MsgBox "Found INVISIBLE Window Handle: " & lhWndP, vbOKOnly + vbInformation
        'End If
        Sleep 200
        SetForegroundWindow (lhWndP)
        Sleep 50
        SendKeys "^f", True ' Find
        Sleep 50
        SendKeys ref, True ' The component refference.
        Sleep 50
        SendKeys vbCr, True ' ENTER - DOIT!
        Sleep 100
        SendKeys Chr(27), True ' Escape : Close Find Dialog
        Sleep 100
        ' Just in case the component is not found do it again.
        ' kicad will go ding.
        SendKeys Chr(27), True
        If ViewOpenGL = False Then
            ' Click where ever we ended up to crossprobe in Eeschema
            ' Not needed in OpenGL mode, and I seems not to work in
            ' the default viewer so yeah.
            Sleep 100
            mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
            Sleep 100
            mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        End If
        Sleep 200
        ' Back to Excel
        SetForegroundWindow (selflhWndP)
        
    Else
        MsgBox "Could not get handles for Excel or pcbnew?", vbOKOnly + vbExclamation
    End If

End Sub

Private Function GetHandleFromPartialCaption(ByRef lWnd As Long, ByVal sCaption As String) As Boolean

    Dim lhWndP As Long
    Dim sStr As String
    GetHandleFromPartialCaption = False
    lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
    Do While lhWndP <> 0
        sStr = String(GetWindowTextLength(lhWndP) + 1, Chr$(0))
        GetWindowText lhWndP, sStr, Len(sStr)
        sStr = Left$(sStr, Len(sStr) - 1)
        If InStr(1, sStr, sCaption) > 0 Then
            GetHandleFromPartialCaption = True
            lWnd = lhWndP
            Exit Do
        End If
        lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
    Loop

End Function

Sub select_reff(ByVal direction As Integer, ByVal highlight_mode As Integer)
    If working Then Return
    Dim cell_reffs() As String
    Dim col_start As Integer
    Dim col_length As Integer
    Dim a As Integer
    Dim cell_reffs_len As Integer
    working = True
    ActiveCell.Font.Bold = False
    ActiveCell.Font.Underline = False
    cell_reffs = Split(Trim(ActiveCell.Text), " ")
    cell_reffs_len = UBound(cell_reffs) - LBound(cell_reffs) + 1
    ref_idx = ref_idx + direction
    If ref_idx > cell_reffs_len Then ref_idx = 1
    If ref_idx <= 0 Then ref_idx = cell_reffs_len
    col_start = 0
    col_length = 0
    For a = 1 To cell_reffs_len
        col_length = Len(cell_reffs(a - 1)) + 1
        If a = ref_idx Then Exit For
        col_start = col_start + col_length
    Next
    ActiveCell.Characters(Start:=col_start, Length:=col_length).Font.Bold = True
    ActiveCell.Characters(Start:=col_start, Length:=col_length).Font.Underline = True
    If (highlight_mode = 1) Then ' Placed
        ActiveCell.Characters(Start:=col_start, Length:=col_length).Font.Color = RGB(0, 255, 0)
    ElseIf (highlight_mode = 2) Then ' unplaced
        ActiveCell.Characters(Start:=col_start, Length:=col_length).Font.Color = RGB(0, 0, 0)
    End If
    If highlight_mode = 0 Then pcbnew_find (cell_reffs(ref_idx - 1))
    working = False
End Sub

Sub comp_prev()
    select_reff -1, 0
End Sub

Sub comp_next()
    select_reff 1, 0
End Sub

Sub comp_placed()
    select_reff 0, 1
End Sub

Sub comp_unplaced()
    select_reff 0, 2
End Sub
