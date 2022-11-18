VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_pumps 
   Caption         =   "Μητρώο Αντλιών"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11175
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
   ScaleHeight     =   484
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   745
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame fr_top_frame 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6255
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
         TabIndex        =   8
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
      TabIndex        =   9
      Top             =   480
      Width           =   11055
      Begin MSComctlLib.ListView lv 
         Height          =   975
         Left            =   180
         TabIndex        =   14
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
      Height          =   2775
      Left            =   1080
      TabIndex        =   10
      Top             =   3720
      Width           =   8415
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "field=power_consumption;"
         Top             =   1680
         Width           =   855
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
         Left            =   3480
         TabIndex        =   5
         Top             =   2160
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
         Left            =   6960
         TabIndex        =   4
         Top             =   2160
         Width           =   1335
      End
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txt_pump_quota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "field=water_supply;"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txt_pump_area 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   1
         Tag             =   "field=place;"
         Top             =   720
         Width           =   6015
      End
      Begin VB.TextBox txt_pump_code 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   0
         Tag             =   "field=pump_name;"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Κατανάλωση KW"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lb_pump_quota 
         Alignment       =   1  'Right Justify
         Caption         =   "Παροχή (m3/h)"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lb_pump_area 
         Alignment       =   1  'Right Justify
         Caption         =   "Περιοχή"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lb_pump_code 
         Alignment       =   1  'Right Justify
         Caption         =   "Κωδικός αντλίας"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frm_pumps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Initialize(Optional ByVal arg As String = vbNullString)
    With lv.ColumnHeaders
        If .Count = 0 Then
            .Add , , "Κωδικός", 80 * Screen.TwipsPerPixelX
            .Add , , "Περιοχή", 300 * Screen.TwipsPerPixelX
            .Add , , "Παροχή", 80 * Screen.TwipsPerPixelX
            .Add , , "Κατανάλωση", 80 * Screen.TwipsPerPixelX
        End If
    End With
    
    PopulateLVData arg
End Sub

Private Sub cmd_exit_Click()
    Unload Me
End Sub

Private Sub cmd_insert_Click()
    Dim response As Long
    
    response = DBInsert(Me, "pump")
    If response > 0 Then
        ClearAllFields Me
        PopulateLVData vbNullString
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fr_top_frame.Move 0, 0, Me.ScaleWidth
    fr_data.Move Me.ScaleWidth / 2 - fr_data.Width / 2, Me.ScaleHeight - fr_data.Height - 3
    
    fr_list.Move Me.ScaleWidth / 2 - fr_list.Width / 2, fr_list.Top, fr_list.Width, Me.ScaleHeight - fr_data.Height - fr_list.Top
    lv.Move 45, lv.Top, fr_list.Width * Screen.TwipsPerPixelX - 90, fr_list.Height * Screen.TwipsPerPixelY - lv.Top - 45
End Sub

Private Sub PopulateLVData(ByVal data_filter As String)
    Dim sql As String
    
    sql = "select id, pump_name, place, water_supply, power_consumption from pump order by pump_name"
    LockWindowUpdate lv.hWnd
    DBFillListview lv, sql, False, True
    ListView_AutoSizeColumn lv
    LockWindowUpdate 0
End Sub

Private Sub lv_Click()
    If lv.ListItems.Count = 0 Then Exit Sub
    
    Dim sql As String
    sql = "select id, pump_name, place, water_supply, power_consumption from pump where id = " & Mid(lv.SelectedItem.Key, 2)
    DBFillFields Me, sql
    
'    CheckCaptions
End Sub
