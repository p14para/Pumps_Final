Attribute VB_Name = "mod_db_utils"
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_FIRST = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
Private Const LVSCW_AUTOSIZE As Long = -1
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long

Private Type OpenFileName
    lStructSize         As Long
    hwndOwner           As Long
    hInstance           As Long
    lpstrFilter         As String
    lpstrCustomFilter   As String
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type

Public Sub ListView_AutoSizeColumn(lv As ListView, Optional Column As ColumnHeader = Nothing)
    Dim c As ColumnHeader
    If Column Is Nothing Then
        For Each c In lv.ColumnHeaders
            Call SendMessage(lv.hWnd, LVM_SETCOLUMNWIDTH, c.Index, ByVal LVSCW_AUTOSIZE_USEHEADER)
        Next
    Else
        Call SendMessage(lv.hWnd, LVM_SETCOLUMNWIDTH, Column.Index - 1, ByVal LVSCW_AUTOSIZE_USEHEADER)
    End If
    lv.Refresh
End Sub

Public Sub ListView_RowColor(lv As ListView, Row As ListItem, Optional setColor As Long = 0, Optional setBold As Boolean = False)
    Dim i   As Long
    
    Row.Bold = setBold
    Row.ForeColor = setColor
    
    For i = 1 To Row.ListSubItems.Count
        Row.ListSubItems(i).Bold = setBold
        Row.ListSubItems(i).ForeColor = setColor
    Next i
End Sub

Public Sub ListView_FixDecimals(lv As ListView, cColumn As Integer)
    If lv.ListItems.Count = 0 Then Exit Sub
    Dim i   As Long
    
    For i = 1 To lv.ListItems.Count
        lv.ListItems(i).ListSubItems(cColumn).Text = Replace(lv.ListItems(i).ListSubItems(cColumn).Text, ",", ".")
    Next i
End Sub

Public Sub ShowForm(frm As Form, Optional ByVal arg As String = vbNullString)
    If frm.WindowState <> vbMaximized Then
        frm.Move (mdi_main.ScaleWidth - frm.Width) / 2, (mdi_main.ScaleHeight - frm.Height) / 2
    End If
    frm.Visible = True
    frm.ZOrder 0
    frm.Initialize arg
End Sub

Public Function SilentExists(ByVal frm As String) As Boolean
    Dim fr As Form
    
    For Each fr In Forms
        If fr.Name = frm Then
            SilentExists = True
            Exit For
        End If
    Next
End Function

Public Sub DBFillListview(ByRef lv As ListView, ByVal sql As String, ByVal BuildHeaders As Boolean, ByVal FirstRecordIsKey As Boolean)
    Dim rs          As ADODB.Recordset
    Dim i           As Long
    Dim p           As Long
    
    Set rs = de.conn.Execute(sql)
    
    If FirstRecordIsKey = True Then p = 1
    
    lv.Sorted = False
    
    ' Build headers
    If BuildHeaders Then
        lv.ColumnHeaders.Clear
        For i = p To rs.Fields.Count - 1
            If rs.Fields(i).Type = adDouble Or rs.Fields(i).Type = adInteger And lv.ColumnHeaders.Count > 0 Then
                lv.ColumnHeaders.Add , , rs.Fields(i).Name, 50, lvwColumnRight
            Else
                lv.ColumnHeaders.Add , , rs.Fields(i).Name, 50, lvwColumnLeft
            End If
        Next i
    End If
    
    
    ' Fill data
    lv.ListItems.Clear
    Do Until rs.EOF
        If FirstRecordIsKey = True Then
            lv.ListItems.Add , "k" & rs.Fields(0).Value, IIf(IsNull(rs.Fields(1).Value) = False, rs.Fields(1).Value, " ")
        Else
            lv.ListItems.Add , , IIf(IsNull(rs.Fields(0).Value) = False, rs.Fields(0).Value, " ")
        End If
        
        With lv.ListItems(lv.ListItems.Count).ListSubItems
        For i = p + 1 To rs.Fields.Count - 1
            Select Case rs.Fields(i).Type
                Case adBoolean     ' 11. Boolean or bit
                    If IsNull(rs.Fields(i).Value) = True Then
                        .Add , , "ΟΧΙ"
                    Else
                        .Add , , IIf(rs.Fields(i).Value = True, "ΝΑΙ", "ΟΧΙ")
                    End If
                
                Case adDouble
                    If IsNull(rs.Fields(i).Value) = True Then
                        .Add , , "0.00"
                    Else
                        .Add , , Replace(Format(rs.Fields(i).Value, "#0.00"), ",", ".")
                    End If
                
                Case adDBTimeStamp
                    If IsNull(rs.Fields(i).Value) = True Then
                        .Add , , " "
                    Else
                        .Add , , Format(rs.Fields(i).Value, "dd/mm/yyyy HH:mm")
                    End If
                    
                Case Else
                    .Add , , IIf(IsNull(rs.Fields(i).Value) = False, rs.Fields(i).Value, " ")
            End Select
        Next i
        End With
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

' type=key; -> dbo.table key (update)
' type=time; -> time textbox
' need=1; -> strict field to has a value (combobox)
' ui=null; -> include in (insert, update)
' check<>null; -> validate data (insert, update)
' field=field name; -> dbo.field (insert, update)


Public Function DBInsert(ByRef F As Form, ByVal DBTable As String, Optional ByVal tag_include As String = vbNullString) As Long
    Dim obj         As Object
    Dim bfld        As String
    Dim bVal        As String
    Dim bins        As String
    Dim bupd        As String
    Dim sql         As String
    Dim rs          As ADODB.Recordset
    Dim i           As Long
    
    Dim getPK       As Long
    Dim getPKField  As String
    
    Dim chk()   As String
    Dim cval()  As String
    
    ReDim chk(0)
    ReDim cval(0)
    
    For Each obj In F
        If tag_include = vbNullString Or InStr(obj.Tag, tag_include) Then
            If getCmd(obj.Tag, "type=", ";") = "key" Then
                getPK = Val(obj.Text)
                getPKField = getCmd(obj.Tag, "field=", ";")
            End If
        End If
    Next
    
    For Each obj In F
        If tag_include = vbNullString Or InStr(obj.Tag, tag_include) Then
            If getCmd(obj.Tag, "ui=", ";") = vbNullString Then
                If getCmd(obj.Tag, "field=", ";") <> vbNullString Then
                    bfld = bfld & IIf(bfld <> vbNullString, ", ", vbNullString) & getCmd(obj.Tag, "field=", ";")
                End If
                If getCmd(obj.Tag, "check=", ";") <> vbNullString Then
                    ReDim Preserve chk(UBound(chk()) + 1)
                    chk(UBound(chk())) = getCmd(obj.Tag, "field=", ";")
                    ReDim Preserve cval(UBound(cval()) + 1)
                    cval(UBound(cval())) = Trim(obj.Text)
                End If
            End If
        End If
    Next
    
    For i = 1 To UBound(chk())
        sql = "select count(id) as cntEx from " & DBTable & " where " & chk(i) & "='" & cval(i) & "'"
        Set rs = de.conn.Execute(sql)
        If rs.Fields("cntEx").Value > 0 Then
            MsgBox "Ο κωδικός υπάρχει ήδη. Η καταχώρηση δεν έγινε.", vbOKOnly Or vbCritical, "Καταχώρηση"
            Exit Function
        End If
        rs.Close
        Set rs = Nothing
    Next i
    
    sql = "select top 1 " & bfld & " from " & DBTable
    Set rs = de.conn.Execute(sql)
    
    For i = 0 To rs.Fields.Count - 1
        For Each obj In F
            If tag_include = vbNullString Or InStr(obj.Tag, tag_include) Then
                If LCase(getCmd(obj.Tag, "field=", ";")) = LCase(rs.Fields(i).Name) And LCase(rs.Fields(i).Name) <> LCase(getPKField) Then
                
                    Select Case TypeName(obj)
                        Case "TextBox"
                            If getCmd(obj.Tag, "need=", ";") = "1" And Trim(obj.Text) = vbNullString Then
                                MsgBox "Δεν έχετε καταχωρήσει στοιχείο. Η καταχώρηση δεν έγινε.", vbOKOnly Or vbCritical, "Καταχώρηση"
                                DBInsert = -1
                                obj.SetFocus
                                Exit Function
                            End If
                            
                            ' Check TimeStamp
                            If rs.Fields(i).Type = adDBTimeStamp Then
                                If getCmd(obj.Tag, "need=", ";") = "1" Then
                                    If IsDate(obj.Text) = False Or Trim(obj.Text) = vbNullString Then
                                        MsgBox "Δεν έχετε εισάγει ημερομηνία. Η καταχώρηση δεν έγινε.", vbOKOnly Or vbCritical, "Καταχώρηση"
                                        DBInsert = -1
                                        obj.SetFocus
                                        Exit Function
                                    End If
                                End If
                                
                                If IsDate(obj.Text) = False And Trim(obj.Text) <> vbNullString Then
                                    MsgBox "Λάθος στην εισαγωγή της ημερομηνίας. Η καταχώρηση δεν έγινε.", vbOKOnly Or vbCritical, "Καταχώρηση"
                                    DBInsert = -1
                                    obj.SetFocus
                                    Exit Function
                                Else
                                    bins = bins & IIf(bins <> vbNullString, ", ", vbNullString) & rs.Fields(i).Name
                                    bupd = bupd & IIf(bupd <> vbNullString, ", ", vbNullString) & rs.Fields(i).Name & "='"
                                    If getCmd(obj.Tag, "type=", ";") = "time" Then
                                        bVal = bVal & IIf(bVal <> vbNullString, ", ", vbNullString) & "'" & Format(obj.Text, "hh:mm:ss") & "'"
                                        bupd = bupd & Format(obj.Text, "hh:mm:ss") & "'"
                                    Else
                                        bVal = bVal & IIf(bVal <> vbNullString, ", ", vbNullString) & "'" & Format(obj.Text, "yyyy/mm/dd HH:mm") & "'"
                                        bupd = bupd & Format(obj.Text, "yyyy/mm/dd HH:mm") & "'"
                                        End If
                                    End If
                                End If
                            
                            ' Check and build double, integer
                            If rs.Fields(i).Type = adDouble Or rs.Fields(i).Type = adInteger Or rs.Fields(i).Type = adSingle Then
                                If IsNumeric(obj.Text) = False Then
                                    If Trim(obj.Text) <> "" Then
                                        MsgBox "Λάθος στην εισαγωγή αριθμού. Η καταχώρηση δεν έγινε.", vbOKOnly Or vbCritical, "Καταχώρηση"
                                        DBInsert = -1
                                        obj.SetFocus
                                        Exit Function
                                    End If
                                Else
                                    bins = bins & IIf(bins <> vbNullString, ", ", vbNullString) & rs.Fields(i).Name
                                    bupd = bupd & IIf(bupd <> vbNullString, ", ", vbNullString) & rs.Fields(i).Name & "='"
                                    bVal = bVal & IIf(bVal <> vbNullString, ", ", vbNullString) & "'" & Replace(obj.Text, ",", ".") & "'"
                                    bupd = bupd & Replace(obj.Text, ",", ".") & "'"
                                End If
                            End If
                            
                            ' Build text by fixing text length
                            If rs.Fields(i).Type = adVarWChar Then
                                bins = bins & IIf(bins <> vbNullString, ", ", vbNullString) & rs.Fields(i).Name
                                bupd = bupd & IIf(bupd <> vbNullString, ", ", vbNullString) & rs.Fields(i).Name & "='"
                                bVal = bVal & IIf(bVal <> vbNullString, ", ", vbNullString) & "'" & Replace(Left(obj.Text, rs.Fields(i).DefinedSize), "'", "''") & "'"
                                bupd = bupd & Replace(Left(obj.Text, rs.Fields(i).DefinedSize), "'", "''") & "'"
                            End If
                    
                        Case "ComboBox"
                            If getCmd(obj.Tag, "need=", ";") = "1" Then
                                If obj.ListIndex = -1 Then
                                    MsgBox "Δεν έχετε επιλέξει τιμή. Η καταχώρηση δεν έγινε.", vbOKOnly Or vbCritical, "Καταχώρηση"
                                    DBInsert = -1
                                    obj.SetFocus
                                    Exit Function
                                End If
                            End If
                            
                            bins = bins & IIf(bins <> vbNullString, ", ", vbNullString) & rs.Fields(i).Name
                            bupd = bupd & IIf(bupd <> vbNullString, ", ", vbNullString) & rs.Fields(i).Name & "='"
                            If obj.ListIndex > -1 Then
                                bVal = bVal & IIf(bVal <> vbNullString, ", ", vbNullString) & "'" & Replace(obj.ItemData(obj.ListIndex), ",", ".") & "'"
                                bupd = bupd & Replace(obj.ItemData(obj.ListIndex), ",", ".") & "'"
                            Else
                                bVal = bVal & IIf(bVal <> vbNullString, ", ", vbNullString) & "'0'"
                                bupd = bupd & "0'"
                            End If
                            
                        Case "CheckBox"
                            If TypeName(obj) = "CheckBox" Then
                                bins = bins & IIf(bins <> vbNullString, ", ", vbNullString) & rs.Fields(i).Name
                                bupd = bupd & IIf(bupd <> vbNullString, ", ", vbNullString) & rs.Fields(i).Name & "='"
                                bVal = bVal & IIf(bVal <> vbNullString, ", ", vbNullString) & "'" & IIf(obj.Value = vbChecked, 1, 0) & "'"
                                bupd = bupd & IIf(obj.Value = vbChecked, 1, 0) & "'"
                            End If
                    End Select
                    
                End If
            End If
        Next
    Next i

    rs.Close
    Set rs = Nothing

    If getPK = 0 Then
        sql = "insert into " & DBTable & " (" & bins & ") values (" & bVal & ")"
    Else
        sql = "update " & DBTable & " set " & bupd & " where " & getPKField & "=" & getPK
        DBInsert = getPK
    End If
    ' MsgBox sql
    de.conn.Execute sql
    
    If getPK = 0 Then
        sql = "select @@identity as newId"
        Set rs = de.conn.Execute(sql)
        DBInsert = rs.Fields("newId").Value
    End If
    
    MsgBox "Η καταχώρηση έγινε με επιτυχία", vbOKOnly Or vbInformation, "Καταχώρηση"
End Function

Public Function DBWhereFilter(ByRef F As Form) As String
    Dim obj     As Object
    Dim bstr    As String
    Dim getfld  As String
    
    For Each obj In F
        If getCmd(obj.Tag, "filter=", ";") <> vbNullString Then
            If TypeName(obj) = "TextBox" Then
                If Trim(obj.Text) <> vbNullString Then
                    getfld = getCmd(obj.Tag, "field=", ";")
                    bstr = bstr & IIf(bstr <> vbNullString, " and ", vbNullString) & getfld & " like '%" & obj.Text & "%'"
                End If
            End If
            ' If TypeName(obj) = "CheckBox" Then obj.Value = vbUnchecked
        End If
    Next
    
    If bstr <> vbNullString Then DBWhereFilter = "where " & bstr
End Function

Public Function DBDateWhereFilter(ByVal dbfield As String, ByVal dateIn As String, ByVal dateOut As String) As String
    dateIn = Trim(dateIn)
    dateOut = Trim(dateOut)
    
    If dateIn = vbNullString And dateOut = vbNullString Then Exit Function
    
    If dateIn <> vbNullString And IsDate(dateIn) = False Then
        MsgBox "Λάθος στην αρχική ημερομηνία", vbOKOnly Or vbCritical, "Ημερομηνία"
        Exit Function
    End If
    If dateOut <> vbNullString And IsDate(dateOut) = False Then
        MsgBox "Λάθος στην τελική ημερομηνία", vbOKOnly Or vbCritical, "Ημερομηνία"
        Exit Function
    End If
    
    If dateIn <> vbNullString Then dateIn = Format(dateIn, "yyyy/mm/dd 00:00")
    If dateOut <> vbNullString Then dateOut = Format(dateOut, "yyyy/mm/dd 23:59")
    
    If (dateIn <> vbNullString And dateOut = vbNullString) Then
        DBDateWhereFilter = " " & dbfield & " >= '" & dateIn & "' "
        Exit Function
    End If
    If dateIn = vbNullString And dateOut <> vbNullString Then
        DBDateWhereFilter = " " & dbfield & " <= '" & dateOut & "' "
        Exit Function
    End If
    DBDateWhereFilter = " " & dbfield & " between '" & dateIn & "' and '" & dateOut & "' "
End Function

Public Sub ClearAllFields(ByRef F As Form)
    Dim obj     As Object
    
    For Each obj In F
        If getCmd(obj.Tag, "field=", ";") <> vbNullString Then
            If TypeName(obj) = "TextBox" Then obj.Text = vbNullString
            If TypeName(obj) = "CheckBox" Then obj.Value = vbUnchecked
            If TypeName(obj) = "ComboBox" Then obj.ListIndex = -1
        End If
    Next
End Sub


Function getCmd(txt As String, quoteOpen As String, quoteClose As String) As String
    Dim psOpen As Integer
    Dim psClose As Integer

    If quoteOpen <> "" Then psOpen = InStr(txt, quoteOpen) Else psOpen = 0
    If quoteClose <> "" And psOpen > 0 Then psClose = InStr(psOpen + Len(quoteOpen), txt, quoteClose) Else psClose = 0
    If psOpen > 0 And psClose > 0 Then
        getCmd = Mid(txt, psOpen + Len(quoteOpen), psClose - psOpen - Len(quoteOpen))
    End If
End Function

Function remCmd(txt As String, quoteOpen As String, quoteClose As String) As String
    Dim psOpen As Integer
    Dim psClose As Integer

    If quoteOpen <> "" Then psOpen = InStr(txt, quoteOpen) Else psOpen = 0
    If quoteClose <> "" And psOpen > 0 Then psClose = InStr(psOpen, txt, quoteClose) Else psClose = 0
    If psOpen > 0 And psClose > 0 Then
        remCmd = Left(txt, psOpen - 1) & Mid(txt, psClose + 1)
    Else
        remCmd = txt
    End If
End Function

Public Sub DBFillCombobox(ByRef cb As ComboBox, ByVal sql As String, Optional ByVal do_clear As Boolean = False)
    Dim rs          As ADODB.Recordset

    Set rs = de.conn.Execute(sql)
    If do_clear Then cb.Clear
    Do Until rs.EOF
        cb.AddItem IIf(IsNull(rs.Fields(1).Value) = False, rs.Fields(1).Value, " ")
        cb.ItemData(cb.NewIndex) = rs.Fields(0).Value
        rs.MoveNext
    Loop
    rs.Close

    Set rs = Nothing
End Sub

Public Sub DB_Combobox_Fill(ByRef cb As ComboBox, ByVal sql As String, ByRef s() As String, Optional ByVal do_clear As Boolean = False)
    Dim rs          As ADODB.Recordset
    Dim i           As Long

    Set rs = de.conn.Execute(sql)
    ReDim s(rs.RecordCount)
    If do_clear Then cb.Clear
    Do Until rs.EOF
        cb.AddItem IIf(IsNull(rs.Fields(1).Value) = False, rs.Fields(1).Value, " ")
        s(i) = rs.Fields(0).Value
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close

    Set rs = Nothing
End Sub

Public Sub ListBox_DBFill(ByRef lst As ListBox, ByVal sql As String, ByVal select_value As String, Optional ByVal do_clear As Boolean = False)
    Dim rs As ADODB.Recordset, pick_index As Long
    
    Set rs = de.conn.Execute(sql)
    
    pick_index = -1
    If do_clear Then lst.Clear
    Do Until rs.EOF
        If IsNull(rs.Fields(1).Value) = False Then
            lst.AddItem rs.Fields(1).Value
            lst.ItemData(lst.NewIndex) = rs.Fields(0).Value
            If rs.Fields(1).Value = select_value Then
                pick_index = lst.NewIndex
            End If
        End If
        rs.MoveNext
    Loop
    If pick_index > -1 Then lst.ListIndex = pick_index
    
    rs.Close
    Set rs = Nothing
End Sub


Public Sub DBFillFields(ByRef F As Form, ByVal sql As String)
    Dim rs          As ADODB.Recordset
    Dim obj         As Object
    Dim i           As Long
    Dim getDecimals As Long
    Dim bformat     As String

    Set rs = de.conn.Execute(sql)
    
    If rs.EOF Then Exit Sub
    
    For i = 0 To rs.Fields.Count - 1
        For Each obj In F
            If LCase(rs.Fields(i).Name) = LCase(getCmd(obj.Tag, "field=", ";")) Then
                
                Select Case TypeName(obj)
                    Case "TextBox"
                        ' Debug.Print obj.Name, obj.Text, rs.Fields(i).Value, obj.Tag
                        If rs.Fields(i).Type = adDouble Or rs.Fields(i).Type = adSingle Then
                            getDecimals = Val(getCmd(obj.Tag, "decimals=", ";"))
                            If getDecimals Then
                                bformat = "#0." & String(getDecimals, "0")
                            Else
                                bformat = "#0"
                            End If
                            obj.Text = IIf(IsNull(rs.Fields(i).Value), Format(0, bformat), Format(rs.Fields(i).Value, bformat))
                        ElseIf rs.Fields(i).Type = adDBTimeStamp Then
                            If IsNull(rs.Fields(i).Value) = False Then
                                If Year(rs.Fields(i).Value) = 1900 Then
                                    obj.Text = vbNullString
                                Else
                                    obj.Text = Format(rs.Fields(i).Value, "dd/mm/yyyy")
                                End If
                            Else
                                obj.Text = vbNullString
                            End If
                        Else
                            obj.Text = IIf(IsNull(rs.Fields(i).Value), vbNullString, rs.Fields(i).Value)
                        End If
                    Case "ComboBox"
                        ' Debug.Print obj.Name, obj.Text, rs.Fields(i).Value, obj.Tag
                        If IsNull(rs.Fields(i).Value) = True Then
                            obj.ListIndex = -1
                        Else
                            DBComboLookUp obj, rs.Fields(i).Value
                        End If
                    Case "CheckBox"
                        ' Debug.Print obj.Name, obj.Caption, rs.Fields(i).Value, obj.Tag
                        If IsNull(rs.Fields(i).Value) Then
                            obj.Value = vbUnchecked
                        Else
                            If rs.Fields(i).Value = True Then
                                obj.Value = vbChecked
                            Else
                                obj.Value = vbUnchecked
                            End If
                        End If
                End Select
                Exit For
            End If
        Next
    Next i

    rs.Close
    Set rs = Nothing
End Sub

Public Sub DBComboLookUp(ByRef cb As ComboBox, ByVal id As Long)
    Dim i           As Long
    
    For i = 0 To cb.ListCount - 1
        If cb.ItemData(i) = id Then
            cb.ListIndex = i
            Exit Sub
        End If
    Next i
    
    cb.ListIndex = -1
End Sub

Public Function ComboLookUpValue(ByRef cb As ComboBox, ByVal id As Long) As String
    Dim i           As Long
    
    For i = 0 To cb.ListCount - 1
        If cb.ItemData(i) = id Then
            ComboLookUpValue = cb.List(i)
            Exit For
        End If
    Next i
End Function

Public Function ComboBox_ShowByValue(ByRef cb As ComboBox, ByVal vstr As String)
    Dim i   As Long
    
    For i = 0 To cb.ListCount - 1
        If cb.List(i) = vstr Then
            cb.ListIndex = i
            Exit For
        End If
    Next i
End Function

Public Function ExportToFile(ByVal sql As String, ByVal pfile As String, ByVal pdelimiter As String, ByVal plinesep As String, Optional pheaders As String) As Boolean
    Dim rs      As ADODB.Recordset
    Dim i       As Long
    Dim bstr    As String
    Dim ffree   As Long
    Dim gfield  As String
    
    Set rs = de.conn.Execute(sql)
    If rs.EOF Then
        MsgBox "Δεν υπάρχουν δεδομένα προς εξαγωγή", vbOKOnly, "Εξαγωγή στοιχείων"
        Exit Function
    End If
    
    ffree = FreeFile
    Open pfile For Output As #ffree
    If pheaders <> vbNullString Then bstr = Replace(pheaders, "\", vbTab) & vbCrLf
    Do Until rs.EOF
        For i = 0 To rs.Fields.Count - 1
            gfield = IIf(IsNull(rs.Fields(i).Value) = True, vbNullString, rs.Fields(i).Value)
            If gfield = "" Then gfield = " "
            bstr = bstr & gfield
            If i <> rs.Fields.Count - 1 Then bstr = bstr & vbTab
        Next i
        Print #ffree, bstr
        bstr = vbNullString
        rs.MoveNext
    Loop
    
    Close #ffree
    rs.Close
    Set rs = Nothing
    ExportToFile = True
End Function

Public Function Program_Config(ByVal param As String) As String
    Dim sql As String
    Dim rs  As ADODB.Recordset
    
    sql = "select pvalue from prg_config where param='" & param & "'"
    Set rs = de.conn.Execute(sql)
    
    If rs.EOF = True Then
        MsgBox "Parameter " & param & " not found!", vbOKOnly Or vbCritical, "Program Config Error"
        Exit Function
    End If
    Program_Config = IIf(IsNull(rs.Fields("pvalue").Value) = True, vbNullString, rs.Fields("pvalue").Value)
End Function

Public Sub Update_Program_Config(ByVal param As String, ByVal pvalue As String)
    Dim sql As String
    Dim rs  As ADODB.Recordset
    
    sql = "update prg_config set pvalue='" & Replace(pvalue, "'", "''") & "' where param='" & param & "'"
    de.conn.Execute sql
End Sub

Public Sub Export_Config(ByVal param As String, ByVal pfilename As String)
    Dim pvalue  As String
    Dim ffree   As Long
    
    pvalue = Program_Config(param)
    
    ffree = FreeFile
    Open pfilename For Output As #ffree
    Print #ffree, pvalue
    Close #ffree
End Sub

Public Function Date_Complete(ByVal dt As String) As String
    Dim gday    As Long
    Dim gmonth  As Long
    Dim gyear   As Long
    
    Dim ps      As Long
    Dim pn      As Long
    
    Dim mdate   As String
    
    dt = Replace(dt, "-", "/")
    
    ps = InStr(dt, "/")
    If ps Then
        gday = Val(Left(dt, ps - 1))
    End If
    
    pn = InStr(ps + 1, dt, "/")
    
    If pn Then
        gmonth = Val(Mid(dt, ps + 1, pn - ps))
        gyear = Val(Mid(dt, pn + 1))
    Else
        If pn < Len(dt) Then gmonth = Val(Mid(dt, ps + 1))
    End If
    
    If gmonth = 0 Then gmonth = Month(Date)
    If gyear = 0 Then gyear = Year(Date)
    
    mdate = Format(gday, "00") & "/" & Format(gmonth, "00") & "/" & Format(gyear, "00")
    
    If IsDate(mdate) = True Then
        Date_Complete = mdate
    End If
    
End Function

Public Sub ComboBox_Fill(ByRef cb As ComboBox, ByVal entries As String, Optional ByVal default_choice As Long = 0)
    Dim sp()    As String
    Dim i       As Long
    
    sp() = Split(entries, "|")
    cb.Clear
    For i = 0 To UBound(sp()) Step 2
        cb.AddItem sp(i)
        cb.ItemData(cb.NewIndex) = sp(i + 1)
        If (i / 2) = default_choice Then cb.ListIndex = (i / 2)
    Next i
End Sub

Public Function DBField_Value(ByRef pfield As ADODB.Field) As String
    If IsNull(pfield.Value) = True Then
        DBField_Value = vbNullString
    Else
        DBField_Value = pfield.Value
    End If
End Function



Public Sub log_it(ByVal vstr As String)

    Exit Sub
    
    Dim ffree As Long
    
    ffree = FreeFile
    
    Open App.Path & "\log.txt" For Append As #ffree
    Print #ffree, vstr
    Close #ffree
End Sub



Public Sub Swap(ByRef va As Variant, ByRef vb As Variant)
    Dim tmp As Variant
    tmp = va
    va = vb
    vb = tmp
End Sub


'Purpose     :  Allows the user to select a file name from a local or network directory.
'Inputs      :  sInitDir            The initial directory of the file dialog.
'               sFileFilters        A file filter string, with the following format:
'                                   eg. "Excel Files;*.xls|Text Files;*.txt|Word Files;*.doc"
'               [sTitle]            The dialog title
'               [lParentHwnd]       The handle to the parent dialog that is calling this function.
'Outputs     :  Returns the selected path and file name or a zero length string if the user pressed cancel


Public Function BrowseForFile(sInitDir As String, Optional ByVal sFileFilters As String, Optional sTitle As String = "Open File", Optional lParentHwnd As Long) As String
    Dim tFileBrowse As OpenFileName
    Const clMaxLen As Long = 254
    
    tFileBrowse.lStructSize = Len(tFileBrowse)
    
    'Replace friendly deliminators with nulls
    sFileFilters = Replace(sFileFilters, "|", vbNullChar)
    sFileFilters = Replace(sFileFilters, ";", vbNullChar)
    If Right$(sFileFilters, 1) <> vbNullChar Then
        'Add final delimiter
        sFileFilters = sFileFilters & vbNullChar
    End If
    
    'Select a filter
    tFileBrowse.lpstrFilter = sFileFilters & "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar
    'create a buffer for the file
    tFileBrowse.lpstrFile = String(clMaxLen, " ")
    'set the maximum length of a returned file
    tFileBrowse.nMaxFile = clMaxLen + 1
    'Create a buffer for the file title
    tFileBrowse.lpstrFileTitle = Space$(clMaxLen)
    'Set the maximum length of a returned file title
    tFileBrowse.nMaxFileTitle = clMaxLen + 1
    'Set the initial directory
    tFileBrowse.lpstrInitialDir = sInitDir
    'Set the parent handle
    tFileBrowse.hwndOwner = lParentHwnd
    'Set the title
    tFileBrowse.lpstrTitle = sTitle
    
    'No flags
    tFileBrowse.flags = 0

    'Show the dialog
    If GetOpenFileName(tFileBrowse) Then
        BrowseForFile = Trim$(tFileBrowse.lpstrFile)
        If Right$(BrowseForFile, 1) = vbNullChar Then
            'Remove trailing null
            BrowseForFile = Left$(BrowseForFile, Len(BrowseForFile) - 1)
        End If
    End If
End Function
