VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_pumpuse 
   Caption         =   "Χρήση αντλιών"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12015
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   457
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   801
   WindowState     =   2  'Maximized
   Begin VB.Frame fr_top_frame 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8895
      Begin VB.TextBox txt_pump_id 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6480
         TabIndex        =   28
         Tag             =   "field=pump_id;"
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmd_exit 
         Caption         =   "Έξοδος"
         Height          =   435
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   975
      End
      Begin VB.TextBox txt_id 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3000
         TabIndex        =   9
         Tag             =   "type=key;field=id;"
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_producer_id 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4560
         TabIndex        =   8
         Tag             =   "field=producer_id;"
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_nfc_id 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5520
         TabIndex        =   7
         Tag             =   "field=nfc_id;"
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Εκτύπωση"
         Height          =   435
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   30
         Width           =   1335
      End
   End
   Begin VB.Frame fr_list 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   11
      Top             =   360
      Width           =   10815
      Begin MSComctlLib.ListView lv 
         Height          =   975
         Left            =   1680
         TabIndex        =   12
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1720
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame fr_data 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   480
      TabIndex        =   13
      Top             =   2760
      Width           =   7815
      Begin VB.CommandButton cmd_moves_pump 
         Caption         =   "Κινήσεις"
         Height          =   315
         Left            =   1920
         TabIndex        =   26
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   70
         TabIndex        =   4
         Tag             =   "field=pump_name;"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txt_fullname 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   70
         TabIndex        =   2
         Tag             =   "field=fullname;text=1;filter=1;need=1;"
         Top             =   720
         Width           =   5415
      End
      Begin VB.CommandButton cmd_moves_producer 
         Caption         =   "Κινήσεις"
         Height          =   315
         Left            =   6720
         TabIndex        =   19
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmd_insert 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Καταχώρηση"
         Height          =   435
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmd_delete 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Διαγραφή"
         Height          =   435
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txt_nfc_uid 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   70
         TabIndex        =   3
         Tag             =   "field=nfc_uid;"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmd_moves_card 
         Caption         =   "Κινήσεις"
         Height          =   315
         Left            =   3120
         TabIndex        =   16
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmd_all_moves 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Κινήσεις"
         Height          =   435
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txt_date_in 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txt_date_out 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txt_sumofmoves 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "λεπτά"
         Height          =   315
         Left            =   6120
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Αντλία"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Όνομα"
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Κάρτα"
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Από"
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "έως"
         Height          =   315
         Left            =   2280
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Σύνολο"
         Height          =   315
         Left            =   4200
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm_pumpuse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub tool_print_callback(ByVal CallBackValue As Long)
    '
End Sub

Public Sub Initialize(Optional ByVal arg As String = vbNullString)
    With lv.ColumnHeaders
        If .Count = 0 Then
            .Add , , "Ημερομηνία", 130 * Screen.TwipsPerPixelX
            .Add , , "Ώρες", 50 * Screen.TwipsPerPixelX, lvwColumnRight
            .Add , , "Παραγωγός", 250 * Screen.TwipsPerPixelX
            .Add , , "Αντλία", 150 * Screen.TwipsPerPixelX
            .Add , , "Κάρτα", 100 * Screen.TwipsPerPixelX
            .Add , , "Status", 100 * Screen.TwipsPerPixelX
        End If
    End With
    
    PopulateLVData arg
End Sub

Private Sub cmd_all_moves_Click()
    Dim bdate As String
    bdate = Replace(DBDateWhereFilter("pump_use.date_time_in", txt_date_in, txt_date_out), "'", "#")
    If bdate <> vbNullString Then bdate = "where " & bdate
    PopulateLVData bdate
End Sub

Private Sub cmd_exit_Click()
    Unload Me
End Sub

Private Sub PopulateLVData(ByVal data_filter As String)
    Dim sql As String, bwhere As String
    
    sql = "select pump_use.id, pump_use.date_time_in, pump_use.hours, producer.fullname, pump.pump_name, nfc.nfc_uid, pump_use.status, " & _
          "pump_use.producer_id, pump_use.nfc_id, pump_use.pump_id " & _
          "from producer inner join (nfc inner join (pump inner join pump_use on pump.id = pump_use.pump_id) on nfc.id = pump_use.nfc_id) on producer.id = pump_use.producer_id " & _
          data_filter & _
          " order by pump_use.date_time_in desc"

    LockWindowUpdate lv.hWnd
    DBFillListview lv, sql, False, True
    ListView_AutoSizeColumn lv
    LockWindowUpdate 0
    
    Dim i As Long, sum As Double
    For i = 1 To lv.ListItems.Count
        sum = sum + Val(lv.ListItems(i).ListSubItems(1).Text)
    Next i
    txt_sumofmoves = Format(sum, "#0.00")
    
    CheckCaptions
End Sub

Private Sub CheckCaptions()
    If Val(txt_id) = 0 Then
        cmd_insert.Caption = "Καταχώρηση"
        cmd_moves_producer.Enabled = False
        cmd_moves_card.Enabled = False
        cmd_delete.Enabled = False
    Else
        cmd_insert.Caption = "Ανανέωση"
        cmd_moves_producer.Enabled = True
        cmd_moves_card.Enabled = True
        cmd_delete.Enabled = True
    End If
End Sub

Private Sub cmd_insert_Click()
    Dim response As Long
End Sub

Private Sub cmd_moves_card_Click()
    If Val(txt_nfc_id) = 0 Then Exit Sub
    Dim bdate As String
    bdate = Replace(DBDateWhereFilter("pump_use.date_time_in", txt_date_in, txt_date_out), "'", "#")
    PopulateLVData "where pump_use.nfc_id=" & txt_nfc_id & IIf(bdate <> vbNullString, " and " & bdate, vbNullString)
End Sub

Private Sub cmd_moves_producer_Click()
    If Val(txt_producer_id) = 0 Then Exit Sub
    Dim bdate As String
    bdate = Replace(DBDateWhereFilter("pump_use.date_time_in", txt_date_in, txt_date_out), "'", "#")
    PopulateLVData "where pump_use.producer_id=" & txt_producer_id & IIf(bdate <> vbNullString, " and " & bdate, vbNullString)
End Sub

Private Sub cmd_moves_pump_Click()
    If Val(txt_pump_id) = 0 Then Exit Sub
    Dim bdate As String
    bdate = Replace(DBDateWhereFilter("pump_use.date_time_in", txt_date_in, txt_date_out), "'", "#")
    PopulateLVData "where pump_use.pump_id=" & txt_pump_id & IIf(bdate <> vbNullString, " and " & bdate, vbNullString)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fr_top_frame.Move 0, 0, Me.ScaleWidth
    fr_data.Move (Me.ScaleWidth - fr_data.Width) / 2, Me.ScaleHeight - fr_data.Height - 3
    
    fr_list.Move Me.ScaleWidth / 2 - fr_list.Width / 2, fr_list.Top, fr_list.Width, Me.ScaleHeight - fr_data.Height - fr_list.Top
    lv.Move 45, lv.Top, fr_list.Width * Screen.TwipsPerPixelX - 90, fr_list.Height * Screen.TwipsPerPixelY - lv.Top - 45
End Sub

Private Sub lv_Click()
    If lv.ListItems.Count = 0 Then Exit Sub
    
    Dim sql     As String

    sql = "select pump_use.id, pump_use.date_time_in, pump_use.hours, producer.fullname, pump.pump_name, nfc.nfc_uid, pump_use.status, " & _
          "pump_use.producer_id, pump_use.nfc_id, pump_use.pump_id " & _
          "from producer inner join (nfc inner join (pump inner join pump_use on pump.id = pump_use.pump_id) on nfc.id = pump_use.nfc_id) on producer.id = pump_use.producer_id " & _
          " where pump_use.id = " & Mid(lv.SelectedItem.Key, 2)
    DBFillFields Me, sql
    
    CheckCaptions
End Sub


Private Sub txt_date_in_LostFocus()
    txt_date_in = Date_Complete(txt_date_in)
End Sub

Private Sub txt_date_out_LostFocus()
    txt_date_out = Date_Complete(txt_date_out)
End Sub

Private Sub cmd_print_Click()
    printIt
End Sub

Private Sub printIt()
    Dim bstr    As String
    
    Dim idx     As Long
    Dim i       As Long
    Dim j       As Long
       
    tool_print.PrintInitialize frm_paymoves, vbPRORPortrait
    tool_print.SetPrinter tool_print.pic
    tool_print.sg_records_header_posy = 1100
    tool_print.sg_records_per_page = 45
    
    bstr = "ΕΚΤΥΠΩΣΗ ΚΙΝΗΣΕΩΝ - ΑΝΤΛΙΕΣ"
    If txt_date_in <> vbNullString Then bstr = bstr & " ΑΠΟ " & txt_date_in
    If txt_date_out <> vbNullString Then bstr = bstr & " ΕΩΣ " & txt_date_out
    
    tool_print.AddText bstr, 500, 500, 0, vbCenter, "", 10, True, False, False, False
    tool_print.AddText Format(Date, "ddd dd mmmm yyyy"), 500, 250, 0, vbRightJustify, "", 10, False, False, False, False
    
    idx = tool_print.AddColumn("Α/Α", 700, vbRightJustify, False, False)
    idx = tool_print.AddColumn("Ημερομηνία", 1800, vbCenter, False, False)
    idx = tool_print.AddColumn("Ώρες", 800, vbRightJustify, True, True)
    idx = tool_print.AddColumn("Παραγωγός", 4500, vbLeftJustify, False, False)
    idx = tool_print.AddColumn("Αντλία", 700, vbRightJustify, False, False)
    idx = tool_print.AddColumn("Κάρτα", 1400, vbLeftJustify, False, False)
    idx = tool_print.AddColumn("Status", 900, vbLeftJustify, False, False)
    
    For i = 1 To lv.ListItems.Count
        bstr = i & vbTab & _
               lv.ListItems(i).Text & vbTab & _
               lv.ListItems(i).ListSubItems(1).Text & vbTab & _
               lv.ListItems(i).ListSubItems(2).Text & vbTab & _
               lv.ListItems(i).ListSubItems(3).Text & vbTab & _
               lv.ListItems(i).ListSubItems(4).Text & vbTab & _
               lv.ListItems(i).ListSubItems(5).Text
        tool_print.AddRecord bstr
    Next i
    
    tool_print.PreviewPrint
End Sub























