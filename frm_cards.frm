VERSION 5.00
Begin VB.Form frm_cards 
   Caption         =   "Πίστωση ωρών στην κάρτα"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9030
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
   ScaleHeight     =   5820
   ScaleWidth      =   9030
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
      Height          =   3735
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   8415
      Begin VB.ComboBox Pn 
         Height          =   435
         ItemData        =   "frm_cards.frx":0000
         Left            =   2160
         List            =   "frm_cards.frx":000A
         TabIndex        =   19
         Text            =   "0105"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txt_credits_sum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   13
         Tag             =   "field=power_consumption;"
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox txt_card_uid 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txt_owner 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   6
         Tag             =   "field=place;"
         Top             =   720
         Width           =   6015
      End
      Begin VB.TextBox txt_credits 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "field=water_supply;"
         Top             =   1680
         Width           =   855
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
         Left            =   6960
         TabIndex        =   4
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txt_newcredits 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   4200
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "field=power_consumption;"
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label nfc_pump 
         Caption         =   "Πομώνα"
         Height          =   495
         Left            =   1200
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Νέο υπόλοιπο"
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   2910
         Width           =   1935
      End
      Begin VB.Label lb_pump_code 
         Alignment       =   1  'Right Justify
         Caption         =   "Κωδικός κάρτας"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lb_pump_area 
         Alignment       =   1  'Right Justify
         Caption         =   "Κάτοχος"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lb_pump_quota 
         Alignment       =   1  'Right Justify
         Caption         =   "Λεπτά στην κάρτα"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1710
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Αγορά λεπτών"
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   2190
         Width           =   1935
      End
   End
   Begin VB.Frame fr_top_frame 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.TextBox txt_nfc_id 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5640
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_producer_id 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3600
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
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
         TabIndex        =   1
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "txt_nfc_id"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   17
         Top             =   128
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "txt_producer_id"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   128
         Visible         =   0   'False
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frm_cards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Initialize(Optional ByVal arg As String = vbNullString)
    '
End Sub

Private Sub cmd_exit_Click()
    Unload Me
End Sub

Private Sub cmd_insert_Click()
    If Val(txt_newcredits) <= 0 And Left$(txt_newcredits, 1) <> "$" Then
        MsgBox "Λάθος στην εισαγωγή ωρών!", vbOKOnly Or vbCritical, "Πίστωση ωρών"
        txt_newcredits.SelStart = 0
        txt_newcredits.SelLength = Len(txt_newcredits)
        txt_newcredits.SetFocus
        Exit Sub
    End If

    Dim bstr As String
    
    If Left$(txt_newcredits, 1) = "$" Then
        txt_credits_sum = Mid$(txt_newcredits, 2)
    End If
    
    'bstr = Format(Date, "yyyyMMDD")
    bstr = Format(Pn, "00000000")
    bstr = bstr & Format(Val(txt_credits_sum), "0000")
    bstr = bstr & Format(txt_card_uid, "00000000")
    'bstr = bstr & Format(Pn, "0000")
    
    tool_nfc.CommSend bstr
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub txt_newcredits_Change()
    txt_credits_sum = Val(txt_credits) + Val(txt_newcredits)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fr_top_frame.Move 0, 0, Me.ScaleWidth
    fr_data.Move Me.ScaleWidth / 2 - fr_data.Width / 2, Me.ScaleHeight / 2 - fr_data.Height / 2 - 3
End Sub

Public Sub Confirm_Entry()
    If Val(txt_producer_id) = 0 Or Val(txt_nfc_id) = 0 Then Exit Sub
    Dim sql As String
    
    sql = "insert into move (producer_id, nfc_id, date_in, amount) values (" & txt_producer_id & "," & txt_nfc_id & ",#" & Format(Date, "yyyy/MM/DD") & "#,'" & txt_newcredits & "')"
    de.conn.Execute sql
    
    
    MsgBox "Προστέθηκαν τα λεπτά στην κάρτα...", vbOKOnly Or vbInformation, "Πίστωση ωρών"
    Unload Me
End Sub

Public Sub Card_Fail()
    MsgBox "Ανεπιτυχής εγγραφή των λεπτών στην κάρτα...", vbOKOnly Or vbInformation, "Πίστωση ωρών"
End Sub

Private Sub Πομώνα_Click()

End Sub

