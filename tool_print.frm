VERSION 5.00
Begin VB.Form tool_print 
   BorderStyle     =   0  'None
   Caption         =   "Εκτύπωση"
   ClientHeight    =   9150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14790
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
   ScaleHeight     =   610
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   986
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fr_top_frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00408040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14655
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Εκτύπωση"
         Height          =   495
         Left            =   10800
         TabIndex        =   6
         Top             =   60
         Width           =   1215
      End
      Begin VB.ComboBox cbPrinter 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   60
         Width           =   6255
      End
      Begin VB.CommandButton cmd_previous_page 
         Caption         =   "Προηγούμενη"
         Height          =   495
         Left            =   1320
         TabIndex        =   4
         Top             =   60
         Width           =   1455
      End
      Begin VB.CommandButton cmd_exit 
         Cancel          =   -1  'True
         Caption         =   "Έξοδος"
         Height          =   495
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   1095
      End
      Begin VB.CommandButton cmd_next_page 
         Caption         =   "Επόμενη"
         Height          =   495
         Left            =   2880
         TabIndex        =   1
         Top             =   60
         Width           =   1455
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5340
      Left            =   0
      ScaleHeight     =   5310
      ScaleWidth      =   4980
      TabIndex        =   3
      Top             =   720
      Width           =   5010
   End
End
Attribute VB_Name = "tool_print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private px As Single, py As Single, mouseisdown As Boolean

Private Type Printer_SubHeader_Type
    pCaption        As String
    px              As Long
    pWidth          As Long
    pAlign          As AlignmentConstants
    isSum           As Boolean
    cSum            As Double
    isPageContinued As Boolean
    cx              As Long
End Type

Private Type Printer_Column_Type
    pCaption        As String
    px              As Long
    pWidth          As Long
    pAlign          As AlignmentConstants
    pSubHeader()    As Printer_SubHeader_Type
    isSum           As Boolean
    cSum            As Double
    isPageContinued As Boolean
    cx              As Long
End Type

Private Type Text_Type
    pCaption        As String
    px              As Long
    py              As Long
    pWidth          As Long
    pAlign          As AlignmentConstants
    pFont           As String
    pFontsize       As Long
    pBold           As Boolean
    pItalic         As Boolean
    onlyInFirstPage As Boolean
    onlyInFooter    As Boolean
End Type

Private Type Line_Type
    sX              As Long
    sY              As Long
    eX              As Long
    eY              As Long
    pDrawStyle      As DrawStyleConstants
End Type

Private ptxt()      As Text_Type
Private pcol()      As Printer_Column_Type
Private plin()      As Line_Type
Private records()   As String

Private prn_font        As String
Private prn_font_size   As Double
Private prn_font_bold   As Boolean
Private prn_font_italic As Boolean

Private last_page_record        As Long
Private prev_page_record        As Long
Private footer_max_height       As Long

Private prn             As Object
Private prn_orienation  As PrinterObjectConstants
Private prn_papersize   As PrinterObjectConstants

Private call_back_form      As Form
Public print_command_back   As Long

Public sg_cell_padding          As Long
Public sg_records_header_height As Long
Public sg_records_header_posy   As Long
Public sg_records_per_page      As Long

Private Sub EnumeratePrinters(ByRef cb As ComboBox)
    Dim p   As Printer
    Dim i   As Long
    For Each p In Printers
        cb.AddItem p.DeviceName
        If Printer.DeviceName = p.DeviceName Then cb.ListIndex = cb.NewIndex
    Next
    cb.AddItem "Εξαγωγή σε αρχείο"
End Sub

Public Sub PrintInitialize(ByVal CallBackForm As Form, ByVal pOrientation As PrinterObjectConstants, Optional ByVal pPaperSize As PrinterObjectConstants = vbPRPSA4)
    ReDim ptxt(0)
    ReDim pcol(0)
    ReDim plin(0)
    ReDim records(0)
    prn_orienation = pOrientation
    prn_papersize = pPaperSize
    sg_records_header_height = 720
    sg_records_header_posy = 2200
    sg_records_per_page = 26
    last_page_record = 1
    sg_cell_padding = 60
    EnumeratePrinters cbPrinter
    Set call_back_form = CallBackForm
    print_command_back = 0
End Sub

Public Sub AddText(ByVal pCaption As String, ByVal px As Long, ByVal py As Long, ByVal pWidth As Long, ByVal pAlign As AlignmentConstants, _
                   ByVal pFont As String, ByVal pFontsize As Long, _
                   ByVal pBold As Boolean, ByVal pItalic As Boolean, ByVal onlyInFirstPage As Boolean, ByVal onlyInFooter As Boolean)
    ReDim Preserve ptxt(UBound(ptxt()) + 1)
    
    With ptxt(UBound(ptxt()))
        .pCaption = pCaption
        .px = px
        .py = py
        .pWidth = pWidth
        .pAlign = pAlign
        .pFont = pFont
        .pFontsize = pFontsize
        .pBold = pBold
        .pItalic = pItalic
        .onlyInFirstPage = onlyInFirstPage
        .onlyInFooter = onlyInFooter
    End With
End Sub

Public Sub AddLine(ByVal sX As Long, ByVal sY As Long, ByVal eX As Long, ByVal eY As Long, ByVal pDrawStyle As DrawStyleConstants)
    ReDim Preserve plin(UBound(plin()) + 1)
    
    With plin(UBound(plin()))
        .sX = sX
        .sY = sY
        .eX = eX
        .eY = eY
        .pDrawStyle = pDrawStyle
    End With
End Sub

Public Function AddColumn(ByVal vCaption As String, ByVal vWidth As Long, ByVal vAlignment As AlignmentConstants, _
                          ByVal vIsSum As Boolean, ByVal vIsPageContinued As Boolean) As Long
    ReDim Preserve pcol(UBound(pcol()) + 1)
    
    With pcol(UBound(pcol()))
        .pCaption = vCaption
        .pWidth = vWidth
        .pAlign = vAlignment
        .isSum = vIsSum
        .isPageContinued = vIsPageContinued
        ReDim .pSubHeader(0)
    End With
    
    AddColumn = UBound(pcol())
End Function

Public Function AddSubColumn(ByVal col_idx As Long, ByVal vCaption As String, ByVal vWidth As Long, ByVal vAlignment As AlignmentConstants, _
                             ByVal vIsSum As Boolean, ByVal vIsPageContinued As Boolean) As Long
    ReDim Preserve pcol(col_idx).pSubHeader(UBound(pcol(col_idx).pSubHeader()) + 1)
    
    With pcol(col_idx).pSubHeader(UBound(pcol(col_idx).pSubHeader()))
        .pCaption = vCaption
        .pWidth = vWidth
        .pAlign = vAlignment
        .isSum = vIsSum
        .isPageContinued = vIsPageContinued
    End With
End Function

Public Sub AddRecord(ByVal ptxt As String)
    ReDim Preserve records(UBound(records()) + 1)
    
    records(UBound(records)) = ptxt
End Sub

Public Sub PreviewPrint()
    Me.Move 0, 0, Screen.Width, Screen.Height - 50
    
    CalculateColumns 0, vbCenter
    
    last_page_record = 1
    
    PrintTexts False
    PrintRecords sg_records_header_posy, 300
    PrintLines
    Me.Show vbModal
End Sub

Public Sub DirectPrint(ByVal includeEndDoc As Boolean)
    SetPrinter Printer
    
    Dim p As Printer
    
    For Each p In Printers
        If p.DeviceName = mem_last_selected_printer Then
            Set Printer = p
            Exit For
        End If
    Next
    
    Printer.Orientation = prn_orienation
    
    CalculateColumns 0, vbCenter
    
    last_page_record = 1
    PrintTexts False
    PrintRecords sg_records_header_posy, 300
    PrintLines
    
    If includeEndDoc Then
        Printer.EndDoc
    End If
    
    Unload Me
End Sub

Public Sub SetPrinter(ByRef obj As Object)
    Set prn = obj
    
    If TypeName(obj) = "PictureBox" Then
        If prn_orienation = vbPRORPortrait Then
            If prn_papersize = vbPRPSA4 Then obj.Width = 780: obj.Height = 1110
            If prn_papersize = vbPRPSA3 Then obj.Width = 1110: obj.Height = 1570
        Else
            If prn_papersize = vbPRPSA4 Then obj.Width = 1110: obj.Height = 780
            If prn_papersize = vbPRPSA3 Then obj.Width = 1570: obj.Height = 1110
        End If
    End If
End Sub

Private Sub cbPrinter_Click()
    mem_last_selected_printer = cbPrinter.List(cbPrinter.ListIndex)
End Sub

Private Sub cmd_next_page_Click()
    If last_page_record >= UBound(records()) Then Exit Sub
    
    If TypeName(prn) = "PictureBox" Then prn.Cls
    
    PrintTexts True
    PrintRecords sg_records_header_posy, 300
    PrintLines
End Sub

Private Sub cmd_previous_page_Click()
    If last_page_record - sg_records_per_page < 2 Then Exit Sub
    
    If TypeName(prn) = "PictureBox" Then prn.Cls
    
    If last_page_record >= UBound(records()) Then
        last_page_record = last_page_record - sg_records_per_page - last_page_record Mod sg_records_per_page + 1
        If last_page_record < 1 Then last_page_record = 1
    Else
        last_page_record = last_page_record - sg_records_per_page * 2
    End If
    
    If last_page_record = 1 Then
        PrintTexts False
    Else
        PrintTexts True
    End If
    PrintRecords sg_records_header_posy, 300
    PrintLines
End Sub


Private Sub cmd_print_Click()
    If cbPrinter.ListIndex = cbPrinter.ListCount - 1 Then
        ExportData
        Exit Sub
    End If

    If print_command_back = 1 Then
        call_back_form.tool_print_callback 1
        Unload Me
        Exit Sub
    End If
    SetPrinter Printer
    
    Dim p As Printer
    
    For Each p In Printers
        If p.DeviceName = cbPrinter.List(cbPrinter.ListIndex) Then
            Set Printer = p
            Exit For
        End If
    Next
    
    Printer.Orientation = prn_orienation
    
    last_page_record = 1
    PrintTexts False
    PrintRecords sg_records_header_posy, 300
    PrintLines
    
    Printer.EndDoc
    Unload Me
End Sub

Private Sub fr_top_frame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    px = X: py = Y: mouseisdown = True
End Sub
Private Sub fr_top_frame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mouseisdown Then Me.Move Me.Left + X - px, Me.Top + Y - py
End Sub
Private Sub fr_top_frame_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseisdown = False
End Sub

Private Sub cmd_exit_Click()
    call_back_form.tool_print_callback -1
    Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fr_top_frame.Move 0, 0, Me.ScaleWidth
    pic.Move Me.ScaleWidth / 2 - pic.Width / 2, fr_top_frame.Height + 5 ' , Me.ScaleWidth - 20, Me.ScaleHeight - fr_top_frame.Height - 15
End Sub

Private Sub Form_Paint()
    Me.Line (0, 0)-(8, Me.ScaleHeight), &H408040, BF
    Me.Line (8, Me.ScaleHeight - 8)-(Me.ScaleWidth, Me.ScaleHeight), &H408040, BF
    Me.Line (Me.ScaleWidth - 8, fr_top_frame.Height)-(Me.ScaleWidth, Me.ScaleHeight), &H408040, BF
End Sub


Public Sub PrintTexts(ByVal isSecondPage As Boolean)
    Dim i           As Long
    Dim memStyle    As Long
    
    memStyle = prn.DrawStyle
    footer_max_height = 0
    For i = 1 To UBound(ptxt())
        With ptxt(i)
            If isSecondPage = True Then
                If .onlyInFirstPage = False And .onlyInFooter = False Then
                    MultilinePrint .pCaption, .px, .py, .pFontsize, .pAlign, .pBold, .pItalic
                End If
            Else
                If .onlyInFooter = False Then
                    MultilinePrint .pCaption, .px, .py, .pFontsize, .pAlign, .pBold, .pItalic
                End If
            End If
            
            If .onlyInFooter = True Then
                If footer_max_height < .py Then footer_max_height = .py
            End If
        End With
    Next i
    prn.DrawStyle = memStyle
End Sub

Public Sub PrintFooterTexts(ByVal dy As Long)
    Dim i           As Long
    Dim memStyle    As Long
    
    memStyle = prn.DrawStyle
    For i = 1 To UBound(ptxt())
        With ptxt(i)
            If .onlyInFooter = True Then
                MultilinePrint .pCaption, .px, .py + dy, .pFontsize, .pAlign, .pBold, .pItalic
            End If
        End With
    Next i
    prn.DrawStyle = memStyle
End Sub

Public Sub MultilinePrint(ByVal ptxt As String, ByVal px As Long, ByVal py As Long, ByVal f_size As Long, _
                          ByVal t_align As AlignmentConstants, Optional ByVal f_bold As Boolean = False, Optional ByVal f_italic As Boolean = False)
    On Error GoTo err_prn
    
    Dim sp()    As String
    Dim i       As Long
    Dim dy      As Long
    
    PrinterStoreSettings
    prn.FontSize = f_size
    prn.FontBold = f_bold
    prn.FontItalic = f_italic

    sp() = Split(ptxt, vbCrLf)
    For i = 0 To UBound(sp())
        Select Case t_align
            Case vbRightJustify: prn.CurrentX = prn.ScaleWidth - prn.TextWidth(sp(i)) - px
            Case vbLeftJustify: prn.CurrentX = px
            Case vbCenter: prn.CurrentX = prn.ScaleWidth / 2 - prn.TextWidth(sp(i)) / 2
        End Select
        prn.CurrentY = py
        prn.Print sp(i)
        py = py + prn.TextHeight(sp(i))
    Next i
    PrinterRestoreSettings
    Exit Sub
    
err_prn:
    Exit Sub
End Sub

Public Sub PrintText(ByVal ptxt As String, ByVal px As Long, ByVal py As Long, ByVal pw As Long, ByVal ph As Long, _
                     ByVal f_size As Long, ByVal t_orientation As AlignmentConstants, ByVal f_bold As Boolean, Optional ByVal pCrop As Boolean = False)
    Dim sp()    As String
    Dim i       As Long
    Dim dy      As Long
    Dim ch      As Long
    Dim j       As Long
    
    PrinterStoreSettings
    prn.FontSize = f_size
    prn.FontBold = f_bold
    sp() = Split(ptxt, vbCrLf)
    
    If pCrop = True Then
        For i = 0 To UBound(sp())
            For j = Len(ptxt) To 1 Step -1
                If prn.TextWidth(Left(sp(i), j)) <= pw Then
                    sp(i) = Left(sp(i), j)
                    Exit For
                End If
            Next j
        Next i
    End If
    
    If UBound(sp()) > 0 Then
        For i = 0 To UBound(sp())
            ch = ch + prn.TextHeight(sp(i))
        Next i
        
        py = py - ch / (5 - UBound(sp()))
    End If
    
    For i = 0 To UBound(sp())
        Select Case t_orientation
            Case vbLeftJustify: prn.CurrentX = px + sg_cell_padding
            Case vbRightJustify: prn.CurrentX = px + pw - prn.TextWidth(sp(i)) - sg_cell_padding
            Case vbCenter: prn.CurrentX = px + pw / 2 - prn.TextWidth(sp(i)) / 2
        End Select
        prn.CurrentY = py + ph / 2 - prn.TextHeight(sp(i)) / 2
        prn.Print sp(i)
        py = py + prn.TextHeight(sp(i))
    Next i
    PrinterRestoreSettings
End Sub

Public Sub PrintLines()
    Dim i       As Long
    
    For i = 1 To UBound(plin())
        With plin(i)
            prn.DrawStyle = .pDrawStyle
            prn.Line (.sX, .sY)-(.eX, .eY)
        End With
    Next i
End Sub

Public Sub CalculateColumns(ByVal px As Long, ByVal vAlignment As AlignmentConstants)
    Dim i   As Long
    Dim j   As Long
    Dim dx  As Long
    Dim sdx As Long
    Dim smx As Long
    
    For i = 1 To UBound(pcol())
        smx = smx + pcol(i).pWidth
    Next i
    
    If vAlignment = vbCenter Then
        px = prn.ScaleWidth / 2 - smx / 2
    End If
    If vAlignment = vbRightJustify Then
        px = prn.ScaleWidth - px - smx
    End If
    If px < 0 Then px = 0
    dx = px
    
    For i = 1 To UBound(pcol())
        pcol(i).cx = dx
        If UBound(pcol(i).pSubHeader()) > 0 Then
            sdx = dx
            For j = 1 To UBound(pcol(i).pSubHeader())
                pcol(i).pSubHeader(j).cx = sdx
                sdx = sdx + pcol(i).pSubHeader(j).pWidth
            Next j
        End If
        dx = dx + pcol(i).pWidth
    Next i
End Sub

Public Sub PrinterStoreSettings()
    prn_font = prn.Font
    prn_font_size = prn.FontSize
    prn_font_bold = prn.FontBold
    prn_font_italic = prn.FontItalic
End Sub

Public Sub PrinterRestoreSettings()
    prn.Font = prn_font
    prn.FontSize = prn_font_size
    prn.FontBold = prn_font_bold
    prn.FontItalic = prn_font_italic
End Sub

Public Sub PrintHeader(ByVal py As Long)
    Dim i   As Long
    Dim j   As Long
    Dim dx  As Long
    Dim sdx As Long
    Dim nh  As Long
    
    prn.DrawStyle = 0
    
    For i = 1 To UBound(pcol())
    
        nh = sg_records_header_height
        
        If UBound(pcol(i).pSubHeader()) > 0 Then
            nh = sg_records_header_height / 2
        End If
        
        prn.Line (pcol(i).cx, py)-(pcol(i).cx + pcol(i).pWidth, py + sg_records_header_height), , B
        PrintText pcol(i).pCaption, pcol(i).cx, py, pcol(i).pWidth, nh, 9, pcol(i).pAlign, False
        
        If UBound(pcol(i).pSubHeader()) > 0 Then
            For j = 1 To UBound(pcol(i).pSubHeader())
                With pcol(i).pSubHeader(j)
                    prn.Line (.cx, py + sg_records_header_height / 2)-(.cx + .pWidth, py + sg_records_header_height), , B
                    PrintText .pCaption, .cx, py + sg_records_header_height / 2, .pWidth, sg_records_header_height / 2, 9, .pAlign, False
                End With
            Next j
        End If
    Next i
End Sub

Public Sub PrintRecords(ByVal py As Long, ByVal cell_height As Long)
    Dim i               As Long
    Dim j               As Long
    Dim k               As Long
    Dim sp()            As String
    Dim my              As Long
    Dim cl              As Long
    Dim maxPage         As Long
    Dim cntPage         As Long
    Dim count_records   As Long
    
    If (UBound(records()) + (footer_max_height \ cell_height)) Mod sg_records_per_page = 0 Then
        maxPage = (UBound(records()) + (footer_max_height \ cell_height)) \ sg_records_per_page
    Else
        maxPage = (UBound(records()) + (footer_max_height \ cell_height)) \ sg_records_per_page + 1
    End If
    
    cntPage = last_page_record \ sg_records_per_page + 1
    
    PrintHeader py
    
    my = py + sg_records_header_height
    prn.DrawStyle = vbDot
    
    ' Preview & Print: Reset sums
    For i = 1 To UBound(pcol())
        pcol(i).cSum = 0
        For j = 1 To UBound(pcol(i).pSubHeader())
            pcol(i).pSubHeader(j).cSum = 0
        Next j
    Next i
    
    ' Preview: Recalculate sums
    If TypeName(prn) = "PictureBox" Then
        For i = 1 To last_page_record - 1
            sp() = Split(records(i), vbTab)
            
            j = 0
            cl = 1
            Do Until j > UBound(sp())
                If UBound(pcol(cl).pSubHeader()) = 0 Then
                    If pcol(cl).isSum = True Then pcol(cl).cSum = pcol(cl).cSum + Val(Replace(sp(j), ",", "."))
                    j = j + 1
                Else
                    For k = 1 To UBound(pcol(cl).pSubHeader())
                        If pcol(cl).pSubHeader(k).isSum = True Then
                            pcol(cl).pSubHeader(k).cSum = pcol(cl).pSubHeader(k).cSum + Val(Replace(sp(j), ",", "."))
                        End If
                        j = j + 1
                    Next k
                End If
                cl = cl + 1
            Loop
        Next i
    End If
    
    For i = last_page_record To UBound(records())
        sp() = Split(records(i), vbTab)
        
        j = 0
        cl = 1
        Do Until j > UBound(sp())
            If UBound(pcol(cl).pSubHeader()) = 0 Then
                PrintText sp(j), pcol(cl).cx, my, pcol(cl).pWidth, cell_height, 9, pcol(cl).pAlign, False, True
                If pcol(cl).isSum = True Then pcol(cl).cSum = pcol(cl).cSum + Val(Replace(sp(j), ",", "."))
                j = j + 1
            Else
                For k = 1 To UBound(pcol(cl).pSubHeader())
                    PrintText sp(j), pcol(cl).pSubHeader(k).cx, my, pcol(cl).pSubHeader(k).pWidth, cell_height, 9, pcol(cl).pSubHeader(k).pAlign, False, True
                    If pcol(cl).pSubHeader(k).isSum = True Then
                        pcol(cl).pSubHeader(k).cSum = pcol(cl).pSubHeader(k).cSum + Val(Replace(sp(j), ",", "."))
                    End If
                    j = j + 1
                Next k
            End If
            cl = cl + 1
        Loop
        
        my = my + cell_height
        
        prn.DrawStyle = vbDot
        prn.Line (pcol(1).cx, my)-(pcol(UBound(pcol())).cx + pcol(UBound(pcol())).pWidth, my)
        count_records = count_records + 1
        last_page_record = last_page_record + 1
        
        If count_records >= sg_records_per_page Or i = UBound(records()) Then
            j = 0
            cl = 1
            Do Until j > UBound(sp())
                If UBound(pcol(cl).pSubHeader()) = 0 Then
                    If pcol(cl).isSum = True Then
                        PrintText Format(pcol(cl).cSum, "#0.00"), pcol(cl).cx, my, pcol(cl).pWidth, cell_height, 9, pcol(cl).pAlign, True
                    End If
                    j = j + 1
                Else
                    For k = 1 To UBound(pcol(cl).pSubHeader())
                        If pcol(cl).pSubHeader(k).isSum = True Then
                            PrintText Format(pcol(cl).pSubHeader(k).cSum, "#0.00"), pcol(cl).pSubHeader(k).cx, my, pcol(cl).pSubHeader(k).pWidth, cell_height, 9, pcol(cl).pSubHeader(k).pAlign, True
                        End If
                        j = j + 1
                    Next k
                End If
                cl = cl + 1
            Loop
        
            MultilinePrint "Σελίδα " & cntPage & " / " & maxPage, 1000, prn.ScaleHeight - 1000, 10, vbRightJustify, False
            If i = UBound(records()) Then Exit For
            cntPage = cntPage + 1
            
            If TypeName(prn) = "PictureBox" Then
                Exit Sub
            Else
                prn.NewPage
                PrintTexts True
                PrintHeader py
                PrintLines
                my = py + sg_records_header_height
                count_records = 0
            End If
            
            ' PrintHeader records_header_posy
            ' my = py
        End If
    Next i
    
    If my + footer_max_height > prn.Height - 1000 Then
        my = py + sg_records_header_height
        If TypeName(prn) <> "PictureBox" Then prn.NewPage
        PrintTexts True
        PrintHeader py
        PrintFooterTexts my
    Else
        PrintFooterTexts my
    End If
    
End Sub




Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    px = X: py = Y: mouseisdown = True
End Sub
Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mouseisdown Then pic.Move pic.Left + (X - px) / Screen.TwipsPerPixelX, pic.Top + (Y - py) / Screen.TwipsPerPixelY
End Sub
Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseisdown = False
End Sub



Private Sub ExportData()
    Dim i As Long, j As Long
    Dim gfile As String, ffree As Long
    Dim bstr As String, sstr As String
    Dim sp() As String
    
    gfile = BrowseForFile("c:\", "Text file (*.txt);*.txt", "Αρχείο κειμένου")
    If gfile = vbNullString Then Exit Sub
    If LCase(Right(gfile, 4)) <> ".txt" Then gfile = gfile & ".txt"
    
    ffree = FreeFile
    Open gfile For Output As #ffree
    For i = 1 To UBound(pcol())
        bstr = bstr & IIf(bstr <> vbNullString, vbTab, vbNullString) & Replace(pcol(i).pCaption, vbCrLf, " ")
        If UBound(pcol(i).pSubHeader()) > 0 Then
            For j = 1 To UBound(pcol(i).pSubHeader())
                sstr = sstr & IIf(sstr <> vbNullString And j <> 1, vbTab, vbNullString) & Replace(pcol(i).pSubHeader(j).pCaption, vbCrLf, " ")
                If j > 1 Then bstr = bstr & vbTab
            Next j
            sstr = sstr & vbTab
        Else
            sstr = sstr & vbTab
        End If
    Next i
    Print #ffree, bstr
    Print #ffree, sstr
    
    For i = 1 To UBound(records())
        sp() = Split(records(i), vbTab)
        bstr = vbNullString
        For j = 0 To UBound(sp())
            bstr = bstr & IIf(bstr <> vbNullString, vbTab, vbNullString) & sp(j)
        Next j
        Print #ffree, bstr
    Next i
    Close #ffree
    
    MsgBox "Η εξαγωγή των στοιχείων ολοκληρώθηκε στο αρχείο " & vbCrLf & gfile, vbOKOnly Or vbInformation, "Εκτύπωση"
End Sub















