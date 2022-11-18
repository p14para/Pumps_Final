Attribute VB_Name = "mod_general"
Option Explicit

Private last_nfc_shown As String
Public mem_last_selected_printer As String

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Const LVM_FIRST = &H1000
'Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
'Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
'Private Const LVSCW_AUTOSIZE As Long = -1
'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'Public Sub ShowForm(frm As Form)
'    If frm.WindowState <> vbMaximized Then
'        frm.Move (mdi_main.ScaleWidth - frm.Width) / 2, (mdi_main.ScaleHeight - frm.Height) / 2
'    End If
'    frm.Visible = True
'    frm.ZOrder 0
'    frm.Initialize
'End Sub

'Public Sub ListView_AutoSizeColumn(lv As ListView, Optional Column As ColumnHeader = Nothing)
'    Dim c As ColumnHeader
'    If Column Is Nothing Then
'        For Each c In lv.ColumnHeaders
'            Call SendMessage(lv.hWnd, LVM_SETCOLUMNWIDTH, c.Index, ByVal LVSCW_AUTOSIZE_USEHEADER)
'        Next
'    Else
'        Call SendMessage(lv.hWnd, LVM_SETCOLUMNWIDTH, Column.Index - 1, ByVal LVSCW_AUTOSIZE_USEHEADER)
'    End If
'    lv.Refresh
'End Sub
