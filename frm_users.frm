VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_producers 
   Caption         =   "Μητρώο αγροτών"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11055
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   12
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   593
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   737
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame fr_top_frame 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton cmd_clear 
         Caption         =   "Νέα εγγραφή"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   1455
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
         Left            =   5280
         TabIndex        =   16
         Tag             =   "type=key;field=id;"
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmd_exit 
         Caption         =   "Έξοδος"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   0
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
      Height          =   3135
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   11055
      Begin MSComctlLib.ListView lv 
         Height          =   975
         Left            =   1950
         TabIndex        =   15
         Top             =   195
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
   Begin VB.Frame fr_cards 
      Caption         =   "Χρεωμένες κάρτες"
      Height          =   4335
      Left            =   8160
      TabIndex        =   18
      Top             =   3720
      Width           =   2895
      Begin VB.CommandButton cmd_chargecard 
         Caption         =   "Πίστωση λεπτών"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   3720
         Width           =   2655
      End
      Begin VB.CommandButton cmd_removecard 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   22
         Top             =   3240
         Width           =   495
      End
      Begin VB.CommandButton cmd_newcard 
         Caption         =   "Παράδοση κάρτας"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   3240
         Width           =   2175
      End
      Begin VB.ListBox lst_cards 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   2550
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   2655
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
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   7935
      Begin VB.CommandButton cmd_delete 
         Caption         =   "Διαγραφή"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmd_insert 
         Caption         =   "Καταχώρηση"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         TabIndex        =   4
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmd_moves 
         Caption         =   "Κινήσεις"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   5
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txt_user_phone 
         Appearance      =   0  'Flat
         Height          =   1275
         Left            =   1680
         MaxLength       =   44
         MultiLine       =   -1  'True
         TabIndex        =   3
         Tag             =   "field=phone;"
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox txt_user_address 
         Appearance      =   0  'Flat
         Height          =   915
         Left            =   1680
         MaxLength       =   70
         TabIndex        =   2
         Tag             =   "field=address;"
         Top             =   1200
         Width           =   6135
      End
      Begin VB.TextBox txt_user_taxnumber 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   1
         Tag             =   "field=taxnumber;"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txt_fullname 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   1680
         MaxLength       =   70
         TabIndex        =   0
         Tag             =   "field=fullname;text=1;filter=1;need=1;"
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Τηλέφωνο"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Διεύθυνση"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "ΑΦΜ"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Όνομα"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frm_producers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Initialize(Optional ByVal arg As String = vbNullString)
    With lv.ColumnHeaders
        If .Count = 0 Then
            .Add , , "Όνομα", 300 * Screen.TwipsPerPixelX
            .Add , , "ΑΦΜ", 90 * Screen.TwipsPerPixelX
            .Add , , "Διεύθυνση", 250 * Screen.TwipsPerPixelX
        End If
    End With
    
    PopulateLVData arg
End Sub

Private Sub cmd_chargecard_Click()
    If lst_cards.ListIndex = -1 Then Exit Sub
    ShowForm frm_cards
    
    frm_cards.txt_producer_id = txt_id
    frm_cards.txt_nfc_id = lst_cards.ItemData(lst_cards.ListIndex)
    
    frm_cards.txt_owner = tool_nfc.lb_owner
    frm_cards.txt_card_uid = tool_nfc.lb_card_uid
    frm_cards.txt_credits = Val(tool_nfc.lb_credits)
    frm_cards.txt_newcredits = "0"
End Sub

Private Sub cmd_clear_Click()
    ClearAllFields Me
    CheckCaptions
End Sub

Private Sub cmd_exit_Click()
    Unload Me
End Sub

Private Sub cmd_insert_Click()
    Dim response As Long
    
    response = DBInsert(Me, "producer")
    If response > 0 Then
        ClearAllFields Me
        PopulateLVData vbNullString
    End If
End Sub

Private Sub cmd_moves_Click()
    ShowForm frm_paymoves, "where move.producer_id=" & txt_id
End Sub

Private Sub cmd_newcard_Click()
    Dim sql As String, rs As ADODB.Recordset
    
    sql = "select id, producer_id from nfc where nfc_uid = '" & tool_nfc.lb_card_uid.Caption & "'"
    Set rs = de.conn.Execute(sql)
    
    If rs.EOF = False Then
        sql = "select fullname from producer where id = " & rs.Fields("producer_id").Value
        Set rs = de.conn.Execute(sql)
        MsgBox "Η κάρτα είναι ήδη χρεωμένη στον " & vbCrLf & rs.Fields("fullname").Value, vbOKOnly Or vbExclamation, "Χρέωση κάρτας"
        Exit Sub
    End If
    
    sql = "insert into nfc (producer_id, nfc_uid) values (" & txt_id & ", '" & tool_nfc.lb_card_uid.Caption & "')"
    Set rs = de.conn.Execute(sql)
    
    PopulateCards
End Sub

Private Sub cmd_removecard_Click()
    If lst_cards.ListCount = 0 Then Exit Sub
    If lst_cards.ListIndex = -1 Then Exit Sub
    
    Dim response As String, sql As String, rs As ADODB.Recordset
    
    sql = "select count(id) as moves from move where nfc_id = " & lst_cards.ItemData(lst_cards.ListIndex)
    Set rs = de.conn.Execute(sql)
    
    If rs.Fields("moves").Value > 0 Then
        MsgBox "Υπάρχουν κινήσεις για την κάρτα " & lst_cards.List(lst_cards.ListIndex) & "." & vbCrLf & "Η αφαίρεση της κάρτας δεν είναι δυνατή.", vbOKOnly Or vbCritical
        Exit Sub
    End If
    
    response = MsgBox("θέλετε να αφαιρέσετε την κάρτα " & lst_cards.List(lst_cards.ListIndex) & vbCrLf & " του " & txt_fullname & ";", vbYesNo Or vbCritical, "Αφαίρεση κάρτας")
    If response = vbYes Then
        sql = "delete from nfc where id = " & lst_cards.ItemData(lst_cards.ListIndex)
        de.conn.Execute sql
        
        PopulateCards
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fr_top_frame.Move 0, 0, Me.ScaleWidth
    fr_data.Move Me.ScaleWidth / 2 - (fr_data.Width + fr_cards.Width) / 2, Me.ScaleHeight - fr_data.Height - 3
    fr_cards.Move fr_data.Left + fr_data.Width, fr_data.Top
    fr_list.Move Me.ScaleWidth / 2 - fr_list.Width / 2, fr_list.Top, fr_list.Width, Me.ScaleHeight - fr_data.Height - fr_list.Top
    lv.Move 45, lv.Top, fr_list.Width * Screen.TwipsPerPixelX - 90, fr_list.Height * Screen.TwipsPerPixelY - lv.Top - 45
End Sub

Private Sub PopulateLVData(ByVal data_filter As String)
    Dim sql As String
    
    sql = "select id, fullname, taxnumber, address from producer order by fullname"
    LockWindowUpdate lv.hWnd
    DBFillListview lv, sql, False, True
    ListView_AutoSizeColumn lv
    LockWindowUpdate 0
    
    CheckCaptions
End Sub

Private Sub PopulateCards()
    If Val(txt_id) = 0 Then Exit Sub
    
    Dim sql As String
    sql = "select id, nfc_uid from nfc where producer_id = " & txt_id
    ListBox_DBFill lst_cards, sql, tool_nfc.lb_card_uid, True
End Sub

Private Sub lst_cards_Click()
    If lst_cards.ListCount = 0 Then Exit Sub
    If lst_cards.List(lst_cards.ListIndex) <> tool_nfc.lb_card_uid Then Exit Sub
    
    cmd_removecard.Enabled = True
    cmd_chargecard.Enabled = True
End Sub

Private Sub lv_Click()
    If lv.ListItems.Count = 0 Then Exit Sub
    
    Dim sql     As String
    sql = "select id, fullname, taxnumber, address, phone from producer where id = " & Mid(lv.SelectedItem.Key, 2)
    DBFillFields Me, sql
    
    PopulateCards
    
    CheckCaptions
End Sub

Private Sub CheckCaptions()
    If Val(txt_id) = 0 Then
        cmd_insert.Caption = "Καταχώρηση"
        cmd_moves.Enabled = False
        cmd_delete.Enabled = False
        cmd_newcard.Enabled = False
        cmd_removecard.Enabled = False
    Else
        cmd_insert.Caption = "Ανανέωση"
        cmd_moves.Enabled = True
        cmd_delete.Enabled = True
        If tool_nfc.lb_card_uid <> vbNullString Then
            cmd_newcard.Enabled = True
            cmd_removecard.Enabled = True
        End If
    End If
End Sub

Public Sub Populate_ProducerId(ByVal producer_id As String)
    Dim sql     As String
    sql = "select id, fullname, taxnumber, address, phone from producer where id = " & producer_id
    DBFillFields Me, sql
    
    PopulateCards
    
    CheckCaptions
End Sub









