VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form tool_nfc 
   BackColor       =   &H00F2E4D7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NFC"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8145
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   543
   Begin VB.ComboBox cb_com 
      Height          =   345
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton cmd_connect 
      Caption         =   "ΣΥΝΔΕΣΗ"
      Height          =   345
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txt_log 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2280
      Width           =   8055
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7080
      Top             =   0
   End
   Begin MSCommLib.MSComm comm 
      Left            =   7440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
   End
   Begin VB.Label lb_check 
      BackColor       =   &H00F2F4E7&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   12
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Έλεγχος εγγραφής"
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lb_credits 
      BackColor       =   &H00F2F4E7&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Πιστωμένα λεπτά"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lb_owner 
      BackColor       =   &H00F2F4E7&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label lb_card_uid 
      BackColor       =   &H00F2F4E7&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lb_card_type 
      BackColor       =   &H00F2F4E7&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Τύπος κάρτας"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Κάτοχος"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lb_pump_code 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Κωδικός κάρτας"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "tool_nfc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    CommSend "201705020140FFFF"
End Sub

Private Sub Command2_Click()
    CommSend "Theodore0234AAAA"
End Sub

Private Sub Form_Load()
    Enumerate_Com cb_com, "arduino"
    If cb_com.ListIndex > -1 Then cmd_connect_Click
End Sub

Private Sub cmd_connect_Click()
    On Error GoTo errHandle
    
    Dim get_comport As Long
    Dim gstr As String
    
    If comm.PortOpen = False Then
        If cb_com.ListIndex > -1 Then
            get_comport = Val(getCmd(cb_com.List(cb_com.ListIndex), "COM", " "))
            If get_comport = 0 Then
                MsgBox "Αδύνατη η σύνδεση με το NFC...", vbOKOnly Or vbCritical, "Σύνδεση με NFC"
                Exit Sub
            End If
            comm.CommPort = get_comport
            tmr.Enabled = True
        Else
            Exit Sub
        End If
    End If
111
    If comm.PortOpen = False Then
        txt_log = vbNullString
        comm.PortOpen = True
        tmr.Enabled = True
        cmd_connect.Caption = "ΑΠΟΣΥΝΔΕΣΗ"
    Else
        tmr.Enabled = False
        comm.PortOpen = False
        cmd_connect.Caption = "ΣΥΝΔΕΣΗ"
    End If
    
    Exit Sub
errHandle:
    tmr.Enabled = False
    MsgBox "Πρόβλημα σύνδεσης με τη συσκευή ανάγνωσης.", vbOKOnly Or vbCritical, "NFC reader"
End Sub

Private Sub lb_card_uid_DblClick()
    Me.Height = 6030
End Sub

Private Sub tmr_Timer()
    Dim gstr    As String
    Dim dstr    As String
    Dim prid    As String
    Dim mcuid   As String
    
    gstr = comm.Input
    
    If gstr <> vbNullString Then
        If Len(txt_log) > 2048 Then txt_log = vbNullString
        txt_log = txt_log & gstr
        txt_log.SelStart = Len(txt_log)
        ' Debug.Print gstr
    End If
    
    Dim p As Long
    
    p = InStr(txt_log, "@")
    If p Then
        txt_log = Mid(txt_log, p + 1)
    End If
    
    p = InStr(txt_log, "END#")
    If p Then
        txt_log = Replace(txt_log, vbLf, vbNullString)
        txt_log = Replace(txt_log, vbCr, vbNullString)
        Debug.Print txt_log
        lb_card_type = getCmd(txt_log, "PICC#", "#")
        mcuid = lb_card_uid
        lb_card_uid = getCmd(txt_log, "CUID#", "#")
        If mcuid <> lb_card_uid Then
            If SilentExists("frm_cards") Then Unload frm_cards
        End If
        dstr = getCmd(txt_log, "DATA#", "#")
        
        If Len(dstr) >= 12 Then
            lb_credits = Val(Mid(dstr, 9, 4))
        Else
            lb_credits = "0"
        End If
        
        dstr = getCmd(txt_log, "DATAW#", "#")
        If Len(dstr) >= 12 Then
            lb_credits = Val(Mid(dstr, 9, 4))
        End If
        
        If InStr(txt_log, "#CHOK") Then
            lb_check = "OK"
            frm_cards.Confirm_Entry
        End If
        If InStr(txt_log, "#CHFA") Then
            lb_check = "FAIL"
            frm_cards.Card_Fail
        End If
        
        
        Dim sql As String, rs As ADODB.Recordset
        sql = "select producer.id, producer.fullname from producer inner join nfc on producer.id = nfc.producer_id where nfc.nfc_uid = '" & lb_card_uid & "'"
        Set rs = de.conn.Execute(sql)
        
        If rs.EOF = False Then
            lb_owner = IIf(IsNull(rs.Fields("fullname").Value) = False, rs.Fields("fullname").Value, "")
            prid = rs.Fields("id").Value
        Else
            lb_owner = "Κανένας"
        End If
        
        mdi_main.sbar.Panels(1).Text = "Κωδικός κάρτας: " & lb_card_uid & ", Χρόνος: " & lb_credits & ", Κάτοχος: " & lb_owner
        If SilentExists("frm_producers") = True Then
            frm_producers.ZOrder 0
        Else
            ShowForm frm_producers
        End If
        
        If prid <> vbNullString Then
            frm_producers.Populate_ProducerId prid
            frm_producers.cmd_newcard.Enabled = False
            frm_producers.cmd_chargecard.Enabled = True
        Else
            If Val(frm_producers.txt_id) > 0 Then
                frm_producers.cmd_newcard.Enabled = True
            End If
        End If
        
        txt_log = vbNullString
    End If
End Sub

Public Sub CommSend(ByVal sstr As String)
    If comm.PortOpen = False Then Exit Sub
    sstr = sstr & vbLf
    comm.Output = sstr
    Debug.Print "OUTPUT> " & sstr
End Sub

Private Sub Enumerate_Com(ByRef cb As ComboBox, Optional ByVal search_for As String = vbNullString)
    Dim strComputer As String
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim point_this As Long

    strComputer = "."
    point_this = -1
    search_for = LCase(search_for)
    
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_SerialPort", , 48)
    For Each objItem In colItems
        cb.AddItem objItem.DeviceID & " " & objItem.Description
        If InStr(LCase(objItem.Description), search_for) Then
            point_this = cb.NewIndex
        End If
    Next
    cb.ListIndex = point_this
    
    Set objItem = Nothing
    Set colItems = Nothing
    Set objWMIService = Nothing
End Sub
