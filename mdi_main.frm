VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdi_main 
   BackColor       =   &H8000000C&
   Caption         =   "Αντλίες"
   ClientHeight    =   8685
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12930
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar sbar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8310
      Width           =   12930
      _ExtentX        =   22807
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   22278
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnu_file 
      Caption         =   "Αρχείο"
      Begin VB.Menu mnu_file_pumps 
         Caption         =   "Αντλίες"
      End
      Begin VB.Menu mnu_file_producers 
         Caption         =   "Παραγωγοί"
      End
      Begin VB.Menu mnu_file_sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_file_paymoves 
         Caption         =   "Αγορές ωρών"
      End
      Begin VB.Menu mnu_file_pumpuse 
         Caption         =   "Χρήση αντλιών"
      End
      Begin VB.Menu mnu_file_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_file_exit 
         Caption         =   "Έξοδος"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnu_action 
      Caption         =   "Εργασίες"
      Begin VB.Menu mnu_action_pump_data 
         Caption         =   "Εισαγωγή δεδομένων"
      End
   End
End
Attribute VB_Name = "mdi_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    tool_nfc.Visible = True
    de.conn.Open
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    de.conn.Close
End Sub

Private Sub mnu_action_pump_data_Click()
    import_pump_data
End Sub

Private Sub mnu_file_exit_Click()
    Unload Me
End Sub

Private Sub mnu_file_paymoves_Click()
    ShowForm frm_paymoves
End Sub

Private Sub mnu_file_pumpuse_Click()
    ShowForm frm_pumpuse
End Sub

Private Sub mnu_file_producers_Click()
    ShowForm frm_producers
End Sub

Private Sub mnu_file_pumps_Click()
    ShowForm frm_pumps
End Sub

Private Sub import_pump_data()
    Dim getf As String
    
    getf = BrowseForFile("C:\", "Pump data (*.dat);*.dat", "’νοιγμα αρχείου καταγραφής", Me.hWnd)
    If getf = vbNullString Then Exit Sub
    
    Dim ffree As Long, gstr As String, sp() As String
    Dim get_pump_id As Long, sql As String, rs As ADODB.Recordset
    
    ffree = FreeFile
    Open getf For Input As #ffree
    Line Input #ffree, gstr
    sql = "select id from pump where pump_name='" & gstr & "'"
    Set rs = de.conn.Execute(sql)
    If rs.EOF = True Then
        MsgBox "’γνωστη αντλία...", vbOKOnly Or vbCritical, "Εισαγωγή κινήσεων αντλίας"
        Exit Sub
    Else
        get_pump_id = rs.Fields("id").Value
    End If
    
    Dim get_last_date_in As String
    sql = "select max(date_time_in) as maxdate from pump_use where pump_id = " & get_pump_id
    Set rs = de.conn.Execute(sql)
    
    If IsNull(rs.Fields("maxdate").Value) Then
        get_last_date_in = "1900/01/01 00:00:00"
    Else
        get_last_date_in = Format(rs.Fields("maxdate").Value, "yyyy/mm/dd hh:mm:ss")
    End If
    
    Dim get_producer_id As Long, get_nfc_id As Long
    Dim cnt_not_found As Long, count_found As Long, count_out_of_date As Long
    
    Do Until EOF(ffree)
        Line Input #ffree, gstr
        sp() = Split(gstr, ";")
        If UBound(sp()) = 2 Then
            sql = "select id, producer_id, nfc_uid from nfc where nfc.nfc_uid = '" & sp(0) & "'"
            Set rs = de.conn.Execute(sql)
            
            If rs.EOF = True Then
                cnt_not_found = cnt_not_found + 1
            Else
                If Format(sp(2), "yyyy/mm/dd hh:mm:ss") <= get_last_date_in Then
                    count_out_of_date = count_out_of_date + 1
                Else
                    count_found = count_found + 1
                    sql = "insert into pump_use (producer_id, nfc_id, pump_id, date_time_in, hours, status) values (" & _
                           rs.Fields("producer_id").Value & "," & _
                           rs.Fields("id").Value & "," & _
                           get_pump_id & ",#" & _
                           Format(sp(2), "yyyy/mm/dd hh:mm:ss") & "#," & _
                           sp(1) & "," & _
                           "0)"
                    Set rs = de.conn.Execute(sql)
                End If
            End If
            
        End If
    Loop
    Close #ffree
    
    MsgBox "Έγινε εισαγωγή " & count_found & ", απορρίφθηκαν " & cnt_not_found & vbCrLf & _
           "Βρέθηκαν " & count_out_of_date & " παλαιότερες εγγραφές που δεν έγιναν εισαγωγή.", vbOKOnly Or vbInformation, "Εισαγωγή κινήσεων αντλίας"
    
    
End Sub





























