VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "-"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9660
   OleObjectBlob   =   "Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ADDRESS_HOTEL_Change()
Call Find_Value2("*" & ADDRESS_HOTEL.Value & "*", ThisWorkbook.Worksheets("Help").Range("HOTELS"))
    ADDRESS_HOTEL.BackColor = &H80000005
End Sub

Private Sub AGE_Change()
SpinButton1.Value = Val(AGE.Value)
End Sub

Private Sub AGE_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

SpinButton1.Value = Val(AGE.Value)
End Sub

Private Sub BYKE_COLOUR_Change()
BYKE_COLOUR.BackColor = &H80000005
PROVERKA_Click
End Sub

Private Sub BYKE_REG_Change()
List_of_COLOR BYKE_COLOUR
'PERIOD_HIRE_Change
BYKE_REG.BackColor = &H80000005
PROVERKA_Click
Dim data_base As Worksheet
Set data_base = ThisWorkbook.Worksheets("������ ����������")
TECH_PERIOD.Value = Application.WorksheetFunction.SumIfs(data_base.Range("d3:d300"), data_base.Range("A3:A300"), BYKE_TYPE, data_base.Range("b3:b300"), BYKE_REG)
BYKE_SPEED_FACKT = Application.WorksheetFunction.SumIfs(data_base.Range("e3:e300"), data_base.Range("A3:A300"), BYKE_TYPE, data_base.Range("b3:b300"), BYKE_REG)
End Sub

Private Sub BYKE_SPEED_FACKT_Change()
    BYKE_SPEED_FACKT.BackColor = &H80000005
    
'    If Val(TECH_PERIOD.Value) - Val(BYKE_SPEED_FACKT.Value) < ThisWorkbook.Worksheets("������ ����������").Cells(1, 5).Value Then
'        BYKE_SPEED_FACKT.BackColor = &HFFFF&     ' - ����� ����
'
'    Else
'        TECH_PERIOD.BackColor = &H80000005 ' - ����� ����
'    End If
    
    PROVERKA_Click
End Sub

Private Sub BYKE_SPEED_FACKT_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub Closer_Click()
SPIDOMETR.Show
End Sub

Private Sub Continie_Click()
Dim iLastRow As Long
Dim data_base, report As Worksheet

Set report = ThisWorkbook.Worksheets("�����")
Set data_base = ThisWorkbook.Worksheets("������ ����������")
iLastRow = data_base.Cells(Rows.Count, 1).End(xlUp).Row

NEW_DOGOVOR.Value = Application.WorksheetFunction.Max(data_base.Range("P4:P500"), ThisWorkbook.Worksheets("�����").Range("P1:P500000"))

If Closer.Enabled = False Then
A = MsgBox("��������� ������ �" & DOGOVOR & " - ���� �������� �� " & HIRE_COMING_TO _
& vbNewLine & "   � ������� ����� �" & NEW_DOGOVOR & " ������� �� " & DateValue(Now) & " (�������)", vbCritical, "��������!!! ������ ��������� ����� �����!")
Else
End If
  
  '===============================================================
'     A = DateSerial(Year(Now), Month(Now), Day(Now))
'     B = DateValue(HIRE_COMING_TO)
     
     ' - ���������� ����� �� ������
     
'     If DateValue(HIRE_COMING_TO) > DateSerial(Year(Now), Month(Now), Day(Now)) Then ' << ===============   ������ ������ ������ �������
'     n = DateDiff("d", HIRE_COMING_TO.Value, DateSerial(Year(Now), Month(Now), Day(Now)))
'     A = MsgBox("�������� ���� �������� ��������" & DateValue(HIRE_COMING_TO) & vbNewLine & "�������� ���� " & n, vbQuestion, "������ ������ ������ �������")
'         HIRE_COMING_FROM.Text = Format(DateSerial(Year(HIRE_COMING_TO), Month(HIRE_COMING_TO), Day(HIRE_COMING_TO)), "dd.mm.yyyy")

' ��������� ������� � �������� ��� � �����
        Set archiv = ThisWorkbook.Worksheets("�����")
        iLastRow_archiv = archiv.Cells(Rows.Count, 1).End(xlUp).Row
        iLastRow_archiv = iLastRow_archiv + 1
archiv.Unprotect
report.Unprotect
        If data_base.Cells(ActiveCell.Row, 1).Value = Form.BYKE_TYPE.Value And _
        (data_base.Cells(ActiveCell.Row, 2).Value = Form.BYKE_REG.Value Or data_base.Cells(ActiveCell.Row, 2).Value = Val(Form.BYKE_REG.Value)) And _
        data_base.Cells(ActiveCell.Row, 3).Value = Form.BYKE_COLOUR.Value And _
        data_base.Cells(ActiveCell.Row, 7).Value = "������" Then
        
            right_row = ActiveCell.Row
        Else
            For i = 4 To 200
                If data_base.Cells(i, 1).Value = BYKE_TYPE.Value And (data_base.Cells(i, 2).Value = BYKE_REG.Value Or data_base.Cells(i, 2).Value = Val(BYKE_REG.Value)) And _
                    data_base.Cells(i, 3).Value = BYKE_COLOUR.Value And data_base.Cells(i, 7).Value <> "������" Then
                right_row = i
                i = 200
                Else
                End If
            Next i
        End If

' ��������� ������� � �������� ��� � �����
            


            data_base.Rows(right_row).Copy
            archiv.Rows(iLastRow_archiv).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            archiv.Rows(iLastRow_archiv).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            
            archiv.Cells(iLastRow_archiv, 28).Value = Now

archiv.Cells.FormatConditions.Delete

    mes3 = MsgBox("������� ������� � �����", vbInformation, "������� � " & " " & Form.DOGOVOR.Value & " " & Form.BYKE_TYPE.Value & " " & Form.BYKE_REG.Value)

'�������� ���������� �� ����� �����
            gde = data_base.Cells(right_row, 14).Value
                Select Case gde
                    Case "BAHT": report.Cells(11, 3).Value = report.Cells(11, 3).Value - data_base.Cells(right_row, 13).Value
                    Case "EURO": report.Cells(12, 3).Value = report.Cells(12, 3).Value - data_base.Cells(right_row, 13).Value
                    Case "DOLLAR": report.Cells(13, 3).Value = report.Cells(13, 3).Value - data_base.Cells(right_row, 13).Value
                    Case "��������": report.Cells(14, 3).Value = report.Cells(14, 3).Value - 1
               End Select

            gde = data_base.Cells(right_row, 7).Value
            Select Case gde
                Case "������": report.Cells(5, 3).Value = report.Cells(5, 3).Value - 1
    '            Case "��������": report.Cells(6, 3).Value = report.Cells(6, 3).Value + 1
    '            Case "�����": report.Cells(7, 3).Value = report.Cells(7, 3).Value + 1
                Case "������": report.Cells(8, 3).Value = report.Cells(8, 3).Value - 1
    '            report.Cells(8, 5).Value = data_base.Cells(I, 1).Value & "-" & data_base.Cells(I, 2).Value & ";"
    '
                Case "������": report.Cells(9, 3).Value = report.Cells(9, 3).Value - 1
                Case "�����": report.Cells(9, 3).Value = report.Cells(9, 3).Value - 1
            End Select
        
   mes4 = MsgBox("������� ������� � ��ר�", vbInformation, "������� � " & " " & Form.DOGOVOR.Value & " " & Form.BYKE_TYPE.Value & " " & Form.BYKE_REG.Value)
        
'------ ������� ����� � ������ �� �����������
            Status.Caption = 2
            NEW_DOGOVOR.Visible = True
            DOGOVOR.Visible = False
            ScrollBar1.Visible = False
            If Application.WorksheetFunction.Max(data_base.Range("P4:P1000")) >= 10000 Then
            NEW_DOGOVOR.Value = Application.WorksheetFunction.Max(data_base.Range("P4:P500"), ThisWorkbook.Worksheets("�����").Range("P1:P500000")) + 1
            Else
            NEW_DOGOVOR.Value = 10001
            End If
            
            Frame1.Enabled = True
            Frame2.Enabled = True
            

            List_of_model BYKE_TYPE
            List_of_number BYKE_REG
            List_of_COLOR BYKE_COLOUR
            
            
            Form.Caption = "������� �" & NEW_DOGOVOR.Value & " �� " & HIRE_COMING_FROM
            
            For i = 7 To 30
            data_base.Cells(right_row, i).Value = ""
            Next i

            ��������.Enabled = True
            PE4AT.Enabled = False
            Continie.Enabled = False
        
        
    If DateValue(HIRE_COMING_TO) > DateSerial(Year(Now), Month(Now), Day(Now)) Then
'       report.Cells(I + 1, 5).Value = "����� ������ �������"
    Else
        HIRE_COMING_FROM.Value = HIRE_COMING_TO.Value
        HIRE_COMING_FROM.Enabled = False
        TIME_FROM.Enabled = False
        TIME_TO.Enabled = False
        n = DateDiff("d", HIRE_COMING_TO.Value, DateSerial(Year(Now), Month(Now), Day(Now)))
        HIRE_COMING_TO.Value = Format(DateSerial(Year(DateAdd("d", Val(n), HIRE_COMING_FROM)), Month(DateAdd("d", Val(n), HIRE_COMING_FROM)), Day(DateAdd("d", Val(n), HIRE_COMING_FROM))), "dd.mm.yyyy")
    
        QUANTITY.clear
        For i = n To 31
        QUANTITY.AddItem i
        Next i
        QUANTITY.AddItem "�����"
  
    
    End If
    
    If QUANTITY = "�����" Then
        HIRE_COMING_TO.Value = Format(DateSerial(Year(DateAdd("m", 1, HIRE_COMING_FROM)), Month(DateAdd("m", 1, HIRE_COMING_FROM)), Day(DateAdd("m", 1, HIRE_COMING_FROM))), "dd.mm.yyyy")
    Else
        HIRE_COMING_TO.Value = Format(DateSerial(Year(DateAdd("d", Val(QUANTITY), HIRE_COMING_FROM)), Month(DateAdd("d", Val(QUANTITY), HIRE_COMING_FROM)), Day(DateAdd("d", Val(QUANTITY), HIRE_COMING_FROM))), "dd.mm.yyyy")
    End If
    
    HIRE_COMING_FROM.Enabled = False


    If n <= 0 Then
    QUANTITY.Value = -n
    Else
    QUANTITY.Value = n
    End If
    
    BYKE_TYPE.Enabled = False
    BYKE_REG.Enabled = False
    BYKE_COLOUR.Enabled = False

'==========================================================
  archiv.Protect
  report.Protect
  


End Sub

Private Sub CUSTOMER_NAME_Change()
Call Find_Value("*" & CUSTOMER_NAME.Value & "*", ThisWorkbook.Worksheets("Help").Range("A1:A10000"))
    CUSTOMER_NAME.BackColor = &H80000005
    PROVERKA_Click
End Sub





Private Sub DL_ISSUED_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 46 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub DOCUMENT_Change()
    DOCUMENT.BackColor = &H80000005
    PROVERKA_Click
End Sub

Private Sub List_of_hotel_Click()
ADDRESS_HOTEL.Value = List_of_hotel.Value
End Sub

Private Sub List_of_name_Click()
Dim dannie As Worksheet
Set dannie = ThisWorkbook.Worksheets("Help")

CUSTOMER_NAME.Value = List_of_name.Value

        Set R = ThisWorkbook.Worksheets("Help").Range("a1:a200000")
        Set x = R.Find(What:=Form.CUSTOMER_NAME.Value, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not x Is Nothing Then
            Form.AGE.Value = x.Offset(, 1) ' - �������
            Form.NATIONALITY.Value = x.Offset(, 2) ' - ��������������
            Form.ADDRESS_HOTEL.Value = x.Offset(, 3) ' - �����
            Form.ROOM.Value = x.Offset(, 4) ' - �������
            Form.PASSPORT.Value = x.Offset(, 5) ' - �������
            Form.P_ISSUED.Value = x.Offset(, 6) ' - ���� �������� ��������
            Form.DRIVERS_LICENCE.Value = x.Offset(, 7) ' - �����
            Form.DL_ISSUED.Value = x.Offset(, 8) ' - ���� �������� ����
            Form.TELEFON.Value = x.Offset(, 9) ' - �������
            
            Form.Zametka.Caption = x.Offset(, 10) ' - �������
        End If
End Sub




Private Sub P_ISSUED_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'MsgBox (KeyAscii)
If KeyAscii < 46 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub PE4AT_Click()
Dim ������� As String

If Status.Caption = 3 Then
������� = DOGOVOR.Value
Else
������� = NEW_DOGOVOR.Value
End If

ThisWorkbook.Sheets("�������").Cells(1, 57) = �������

mes = MsgBox("����������� ������� �" & ������� & " �� " & HIRE_COMING_FROM & vbNewLine & _
vbNewLine & "   " & BYKE_TYPE & "   � " & BYKE_REG & "  ����  " & BYKE_COLOUR, vbQuestion + vbYesNo, "������ ��������")

Select Case mes
   Case vbYes: ThisWorkbook.Sheets("�������").PrintOut Copies:=2
   Case vbNo:
End Select



End Sub

'Private Sub PERIOD_RATE_Change()
'Dim data_base As Worksheet
'Set data_base = ThisWorkbook.Worksheets("������ ����������")
'
'If Status.Caption = 3 Then
'    If RATE_TCS > 1000 Then
'        PERIOD_RATE.Value = "�����"
'    Else
'        PERIOD_RATE.Value = "����"
'    End If
'Else
'    If PERIOD_RATE.Value = "�����" Then
'        k = Application.WorksheetFunction.Match(BYKE_TYPE.Value, ThisWorkbook.Worksheets("����").Range("C1:S1"), 0)
'        RATE_TCS.Value = Application.WorksheetFunction.VLookup(30, ThisWorkbook.Worksheets("����").Range("b3:S10"), k, True)
'        HIRE_COMING_TO.Value = DateSerial(Year(DateAdd("m", Val(QUANTITY), HIRE_COMING_FROM)), Month(DateAdd("m", Val(QUANTITY), HIRE_COMING_FROM)), Day(DateAdd("m", QUANTITY, HIRE_COMING_FROM)))
'        PERIOD_HIRE.Value = "�����"
'    Else
'        PERIOD_RATE.Value = "����"
'        PERIOD_HIRE.Value = "����"
'                If data_base.Cells(ActiveCell.Row, 1) = "" And BYKE_TYPE.Value = "" Then
'                    RATE_TCS.Value = 0
'                Else
'                    k = Application.WorksheetFunction.Match(BYKE_TYPE.Value, ThisWorkbook.Worksheets("����").Range("B1:M1"), 0)
'                    RATE_TCS.Value = Application.WorksheetFunction.VLookup(0, ThisWorkbook.Worksheets("����").Range("B3:M10"), k, False)
'                End If
'
'                If HIRE_COMING_FROM = "" Then
'                    HIRE_COMING_TO.Value = ""
'                Else
'                    HIRE_COMING_TO.Value = DateSerial(Year(DateAdd("d", QUANTITY, HIRE_COMING_FROM)), Month(DateAdd("d", QUANTITY, HIRE_COMING_FROM)), Day(DateAdd("d", QUANTITY, HIRE_COMING_FROM)))
'                End If
'    End If
'    PRICE.Value = RATE_TCS * PER
'    TOTAL = PRICE * Val(QUANTITY)
'End If
'
'End Sub



Private Sub PRICE_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub PROVERKA_Click()
i = 0
If CUSTOMER_NAME.Value = "" Then
    CUSTOMER_NAME.BackColor = &HC0C0FF
    i = i + 1
Else
    CUSTOMER_NAME.BackColor = &H80000005
End If

If OptionButton1.Value = True Then
    If VALUTA_FIELD = "" Then
        VALUTA_FIELD.BackColor = &HC0C0FF
        i = i + 1
    Else
        VALUTA_FIELD.BackColor = &H80000005
    End If
Else
    
    If DOCUMENT = "" Then
        DOCUMENT.BackColor = &HC0C0FF
        i = i + 1
    Else
        DOCUMENT.BackColor = &H80000005
    End If
End If


If BYKE_TYPE.Value = "" Then
    BYKE_TYPE.BackColor = &HC0C0FF
    RATE_TCS.BackColor = &HC0C0FF
    i = i + 1
Else
    BYKE_TYPE.BackColor = &H80000005
    RATE_TCS.BackColor = &H80000005
End If

If BYKE_REG.Value = "" Then
    BYKE_REG.BackColor = &HC0C0FF
    i = i + 1
Else
    BYKE_REG.BackColor = &H80000005
End If

If BYKE_COLOUR.Value = "" Then
    BYKE_COLOUR.BackColor = &HC0C0FF
    i = i + 1
Else
    BYKE_COLOUR.BackColor = &H80000005
End If

If TECH_PERIOD.Value = "" Then
    TECH_PERIOD.BackColor = &HC0C0FF
    i = i + 1
Else
    TECH_PERIOD.BackColor = &H80000005
End If

If BYKE_SPEED_FACKT.Value = "" Or BYKE_SPEED_FACKT.Value = 0 Then
    BYKE_SPEED_FACKT.BackColor = &HC0C0FF
    i = i + 1
Else
    BYKE_SPEED_FACKT.BackColor = &H80000005
    
    If Val(TECH_PERIOD.Value) - Val(BYKE_SPEED_FACKT.Value) < ThisWorkbook.Worksheets("������ ����������").Cells(1, 5).Value Then
        BYKE_SPEED_FACKT.BackColor = &HFFFF&     ' - ����� ����
        
    Else
        TECH_PERIOD.BackColor = &H80000005 ' - ����� ����
    End If
    
End If



If i > 0 Then
    ��������.Enabled = False
Else
    ��������.Enabled = True
    PROVERKA.Enabled = False
End If


If Status.Caption = 3 Then
    ��������.Enabled = False
    PE4AT.Enabled = True
    PROVERKA.Enabled = False
Else
    If i > 0 Then
        ��������.Enabled = False
        PE4AT.Enabled = False
    Else
        ��������.Enabled = True
        PE4AT.Enabled = True
        PROVERKA.Enabled = False
    End If
End If


End Sub

Private Sub QUANTITY_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub RATE_TCS_Change()
If RATE_TCS.Value = 0 Then
    RATE_TCS.BackColor = &HC0C0FF ' - ������� ���� �������
Else
    RATE_TCS.BackColor = &H80000005 ' - ����� ����
End If
End Sub

Private Sub SpinButton1_Change()
AGE.Value = SpinButton1.Value
End Sub

Private Sub TECH_PERIOD_Change()
TECH_PERIOD.BackColor = &H80000005 ' - ����� ����
End Sub

Private Sub TECH_PERIOD_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub TextBox1_Change()
Call Find_Value("*" & TextBox1.Value & "*", ThisWorkbook.Worksheets("Help").Range("A1:A300"))
End Sub

Private Sub Find_Value(sValue As String, rFindRange As Range)
 If sValue = "*" Then List_of_name.clear:  Exit Sub
 Dim rFndRng As Range
 Dim sAddress As String
 Set rFndRng = rFindRange.Find(What:=sValue, LookIn:=xlValues, LookAt:=xlWhole)
 If rFndRng Is Nothing Then Exit Sub
 List_of_name.clear
 sAddress = rFndRng.Address
 Do
 List_of_name.AddItem rFndRng
 Set rFndRng = rFindRange.FindNext(rFndRng)
 Loop While sAddress <> rFndRng.Address
End Sub

Private Sub Find_Value2(sValue As String, rFindRange As Range)
 If sValue = "*" Then List_of_hotel.clear:  Exit Sub
 Dim rFndRng As Range
 Dim sAddress As String
 Set rFndRng = rFindRange.Find(What:=sValue, LookIn:=xlValues, LookAt:=xlWhole)
 If rFndRng Is Nothing Then Exit Sub
 List_of_hotel.clear
 sAddress = rFndRng.Address
 Do
 List_of_hotel.AddItem rFndRng
 Set rFndRng = rFindRange.FindNext(rFndRng)
 Loop While sAddress <> rFndRng.Address
End Sub


Private Sub TOTAL_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If Status.Caption = "2" And CloseMode = vbFormControlMenu Then
A = MsgBox("������� ��� ����������� ����������", , "��������� �����")
Cancel = True
Else
End If

End Sub

Private Sub VALUTA_FIELD_Change()
    VALUTA_FIELD.BackColor = &H80000005 ' - ����� ����
    PROVERKA_Click
End Sub

Private Sub VALUTA_FIELD_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub vbHelp_Click()

If vbHelp.Value = True Then
vbHelp.Caption = "<<<"
Form.Width = 654
Else
vbHelp.Caption = ">>>"
Form.Width = 490
End If
End Sub

Private Sub ��������_Click()
Dim iLastRow, ifinans, iClient As Long
Dim data_base, report, finans As Worksheet

If P_ISSUED.Value = "" Then P_ISSUED.BackColor = &HC0C0FF   ' - red

If P_ISSUED.Value <> "" Then
    If DateValue(P_ISSUED.Value) < DateValue(Now) Then
            P_ISSUED.BackColor = &HC0C0FF
    Else
    P_ISSUED.BackColor = &H80000005
    End If
    
Else
    P_ISSUED.BackColor = &H80000005
End If


Set report = ThisWorkbook.Worksheets("�����")
Set data_base = ThisWorkbook.Worksheets("������ ����������")
Set finans = ThisWorkbook.Worksheets("�������")

iLastRow = data_base.Cells(Rows.Count, 1).End(xlUp).Row
ifinans = finans.Cells(Rows.Count, 1).End(xlUp).Row + 1
iClient = ThisWorkbook.Worksheets("Help").Cells(Rows.Count, 1).End(xlUp).Row + 1

Set R = ThisWorkbook.Worksheets("Help").Range("a1:a5000")
Set x = R.Find(What:=Form.CUSTOMER_NAME.Value, LookIn:=xlValues, LookAt:=xlWhole)

report.Unprotect
finans.Unprotect

If x Is Nothing Then
    ThisWorkbook.Worksheets("Help").Cells(iClient, 1) = Form.CUSTOMER_NAME.Value   ' - ���
    ThisWorkbook.Worksheets("Help").Cells(iClient, 2) = Form.AGE.Value   ' - �������
    ThisWorkbook.Worksheets("Help").Cells(iClient, 3) = Form.NATIONALITY.Value  ' - ��������������
    ThisWorkbook.Worksheets("Help").Cells(iClient, 4) = Form.ADDRESS_HOTEL.Value   ' - �����
    ThisWorkbook.Worksheets("Help").Cells(iClient, 5) = Form.ROOM.Value  ' - �������
    ThisWorkbook.Worksheets("Help").Cells(iClient, 6) = Form.PASSPORT.Value  ' - �������
    ThisWorkbook.Worksheets("Help").Cells(iClient, 7) = Form.P_ISSUED.Value ' - ���� �������� ��������
    ThisWorkbook.Worksheets("Help").Cells(iClient, 8) = Form.DRIVERS_LICENCE.Value  ' - �����
    ThisWorkbook.Worksheets("Help").Cells(iClient, 9) = Form.DL_ISSUED.Value  ' - ���� �������� ����
    ThisWorkbook.Worksheets("Help").Cells(iClient, 10) = Form.TELEFON.Value  ' - �������
    
'    Form.Zametka.Caption =  ' - �������
End If


n = DateDiff("d", HIRE_COMING_TO.Value, DateSerial(Year(Now), Month(Now), Day(Now)))


     gde = data_base.Cells(ActiveCell.Row, 7).Value
        Select Case gde
            Case "������": report.Cells(5, 3).Value = report.Cells(5, 3).Value + 1
            Case "��������": report.Cells(6, 3).Value = report.Cells(6, 3).Value - 1
            Case "�����": report.Cells(7, 3).Value = report.Cells(7, 3).Value - 1
            Case "������": report.Cells(8, 3).Value = report.Cells(8, 3).Value - 1
           
            Case "������": report.Cells(9, 3).Value = report.Cells(9, 3).Value + 1
            Case "�����": report.Cells(9, 3).Value = report.Cells(9, 3).Value + 1
        End Select

If data_base.Cells(ActiveCell.Row, 1).Value = BYKE_TYPE.Value And _
(data_base.Cells(ActiveCell.Row, 2).Value = BYKE_REG.Value Or data_base.Cells(ActiveCell.Row, 2).Value = Val(BYKE_REG.Value)) And _
data_base.Cells(ActiveCell.Row, 3).Value = BYKE_COLOUR.Value And _
data_base.Cells(ActiveCell.Row, 7).Value <> "������" Then

    right_row = ActiveCell.Row
Else

    For i = 4 To 200
        If data_base.Cells(i, 1).Value = BYKE_TYPE.Value And (data_base.Cells(i, 2).Value = BYKE_REG.Value Or data_base.Cells(i, 2).Value = Val(BYKE_REG.Value)) And _
            data_base.Cells(i, 3).Value = BYKE_COLOUR.Value And data_base.Cells(i, 7).Value <> "������" Then
        right_row = i
        i = 200
        Else
        End If
    Next i


'    right_row = iLastRow
End If


data_base.Cells(right_row, 1) = BYKE_TYPE.Value
data_base.Cells(right_row, 2) = BYKE_REG.Value
data_base.Cells(right_row, 3) = BYKE_COLOUR.Value

data_base.Cells(right_row, 4) = TECH_PERIOD.Value
data_base.Cells(right_row, 5) = BYKE_SPEED_FACKT.Value
'data_base.Cells(right_row, 6) = BYKE_REG.Value
data_base.Cells(right_row, 7) = "������"
'A = Left(HIRE_COMING_FROM, 2)
'B = Mid(HIRE_COMING_FROM, 4, 2)
'C = Right(HIRE_COMING_FROM, 2)

data_base.Cells(right_row, 8).Value = DateValue(HIRE_COMING_FROM) + TimeValue(TIME_FROM)

data_base.Cells(right_row, 9) = QUANTITY.Value
data_base.Cells(right_row, 10) = DateValue(HIRE_COMING_TO) + TimeValue(TIME_TO)

data_base.Cells(right_row, 11) = RATE_TCS.Value
data_base.Cells(right_row, 12) = TOTAL.Value

If OptionButton1 = True Then
    data_base.Cells(right_row, 13) = VALUTA_FIELD.Value
    data_base.Cells(right_row, 14) = VALUTA.Value
Else
    data_base.Cells(right_row, 13) = DOCUMENT.Value
    data_base.Cells(right_row, 14) = "��������"
End If


data_base.Cells(right_row, 15) = 1
data_base.Cells(right_row, 16) = NEW_DOGOVOR.Value

data_base.Cells(right_row, 17) = CUSTOMER_NAME.Value
data_base.Cells(right_row, 18) = AGE.Value
data_base.Cells(right_row, 19) = NATIONALITY.Value

data_base.Cells(right_row, 20) = ADDRESS_HOTEL.Value
data_base.Cells(right_row, 21) = ROOM.Value

data_base.Cells(right_row, 22) = PASSPORT.Value
data_base.Cells(right_row, 23) = P_ISSUED.Value
data_base.Cells(right_row, 24) = DRIVERS_LICENCE.Value
data_base.Cells(right_row, 25) = DL_ISSUED.Value
data_base.Cells(right_row, 26) = "'" & TELEFON.Value
data_base.Cells(right_row, 27) = Prime4anie.Value ' - ����������
data_base.Cells(right_row, 28) = Now ' - ������� ����

ThisWorkbook.Sheets("�������").Cells(1, 57) = NEW_DOGOVOR.Value

     gde = data_base.Cells(right_row, 14).Value
        Select Case gde
            Case "BAHT": report.Cells(11, 3).Value = report.Cells(11, 3).Value + data_base.Cells(right_row, 13).Value
            Case "EURO": report.Cells(12, 3).Value = report.Cells(12, 3).Value + data_base.Cells(right_row, 13).Value
            Case "DOLLAR": report.Cells(13, 3).Value = report.Cells(13, 3).Value + data_base.Cells(right_row, 13).Value
            Case "��������": report.Cells(14, 3).Value = report.Cells(14, 3).Value + 1
           
       End Select


     gde = data_base.Cells(right_row, 7).Value
        Select Case gde
            Case "������": report.Cells(5, 3).Value = report.Cells(5, 3).Value + 1
            Case "��������": report.Cells(6, 3).Value = report.Cells(6, 3).Value + 1
            Case "�����": report.Cells(7, 3).Value = report.Cells(7, 3).Value + 1
            Case "������": report.Cells(8, 3).Value = report.Cells(8, 3).Value + 1
            report.Cells(8, 5).Value = data_base.Cells(i, 1).Value & "-" & data_base.Cells(i, 2).Value & ";"
            
            Case "������": report.Cells(9, 3).Value = report.Cells(9, 3).Value + 1
            Case "�����": report.Cells(9, 3).Value = report.Cells(9, 3).Value + 1
        End Select
        
For i = 20 To 46

        If report.Cells(i, 2).Value <> "" And report.Cells(i + 1, 2).Value = "" Then
            report.Cells(i + 1, 2).Value = data_base.Cells(right_row, 16).Value '- ����� ��������
            report.Cells(i + 1, 3).Value = data_base.Cells(right_row, 12).Value ' - ����� ��������
            report.Cells(i + 1, 4).Value = "'" & data_base.Cells(right_row, 2).Value ' - ����� �����
            
            If Status.Caption = "2" Then
            A = DOGOVOR.Value
                If DateValue(HIRE_COMING_TO) > DateValue(Now) Then report.Cells(i + 1, 5).Value = "��������� ���. � " & DOGOVOR.Value
                If DateValue(HIRE_COMING_TO) = DateValue(Now) Then report.Cells(i + 1, 5).Value = "��������� ���. � " & DOGOVOR.Value & _
                " c " & Left(data_base.Cells(right_row, 8).Value, 5) & " �� " & Left(data_base.Cells(right_row, 10).Value, 10)
                
                If DateValue(HIRE_COMING_TO) < DateValue(Now) Then report.Cells(i + 1, 5).Value = "����� ������ �������"
            Exit For
            Else
            report.Cells(i + 1, 5).Value = "����:" & data_base.Cells(right_row, 9).Value & " � " & Left(data_base.Cells(right_row, 8).Value, 5) _
            & " �� " & Left(data_base.Cells(right_row, 10).Value, 10)
'            report.Cells(I + 1, 5).Value = DateValue(data_base.Cells(ActiveCell.Row, 8).Value) & "-" & DateValue(data_base.Cells(ActiveCell.Row, 10).Value) & " " & data_base.Cells(ActiveCell.Row, 17).Value & " " & data_base.Cells(ActiveCell.Row, 26).Value
'            report.Cells(Cells(I, 1), Cells(I, 5)).Locked = True
            Exit For
            End If
        Else
        End If

Next i


'        '======================
'            For I = 20 To 32
'                    If report.Cells(I, 2).Value <> "" And report.Cells(I + 1, 2).Value = "" Then
'                        report.Cells(I + 1, 2).Value = data_base.Cells(right_row, 16).Value
'                        report.Cells(I + 1, 3).Value = data_base.Cells(right_row, 12).Value
'                        report.Cells(I + 1, 4).Value = "'" & data_base.Cells(right_row, 2).Value
'
'
'                        I = 32
'            Else
'            End If
'            Next I
'        '======================


'Nomer.Value = Application.WorksheetFunction.Max(finans.Columns(12)) + 1

'finans.Cells(ifinans, 1).Value = 1

finans.Cells(ifinans, 1).Value = DateValue(Now) ' -����
finans.Cells(ifinans, 2).Value = "������" '- ��� ��������
finans.Cells(ifinans, 3).Value = Format(Application.WorksheetFunction.Max(finans.Columns(12)) + 1, "��-0000000")         ' -� ���������
finans.Cells(ifinans, 4).Value = "������"      ' - ������
finans.Cells(ifinans, 5).Value = Form.NEW_DOGOVOR.Value    ' - ����������
finans.Cells(ifinans, 6).Value = "BAHT"        ' - ������

A = data_base.Cells(right_row, 12).Value
finans.Cells(ifinans, 7).Value = data_base.Cells(right_row, 12).Value        ' - �����

finans.Cells(ifinans, 12).Value = Val(Right(finans.Cells(ifinans, 3).Value, 7))       ' - �����
finans.Cells(ifinans, 13).Value = Now        ' - ���� ������

report.Protect
finans.Protect
     
mes = MsgBox("����������� ������� �" & NEW_DOGOVOR & " �� " & HIRE_COMING_FROM & " ?", vbQuestion + vbYesNo, "������ ��������")

Select Case mes
   Case vbYes:
    If Status.Caption = 3 Then
        ������� = DOGOVOR.Value
    Else
        ������� = NEW_DOGOVOR.Value
    End If

    ThisWorkbook.Sheets("�������").Cells(1, 57) = �������
    ThisWorkbook.Sheets("�������").PrintOut Copies:=2
   
   Case vbNo: GoTo next_step
End Select


next_step: A = MsgBox("������� " & ������� & "� ����!", vbQuestion, "������� �������")

Dim ctrl As Control
    For Each ctrl In Me.Controls("Frame2").Controls
    A = TypeName(ctrl)
        If TypeName(ctrl) = "TextBox" Then
            ctrl.Text = ""
          
        End If
    Next
Form.Hide

        
End Sub

Private Sub OptionButton2_Click()
'        DOCUMENT.Value = data_base.Cells(ActiveCell.Row, 13).Value
        DOCUMENT.Visible = True
        VALUTA.Visible = False
        VALUTA_FIELD.Visible = False
        PROVERKA_Click
End Sub


Private Sub QUANTITY_Change()

    If BYKE_TYPE.Value = "" Then
    Else
        If Val(Status.Caption) < 3 Then
            If Val(QUANTITY.Value) > 31 Then QUANTITY.Value = "�����"
            
            If QUANTITY.Value = "�����" Then
                k = Application.WorksheetFunction.Match(BYKE_TYPE.Value, ThisWorkbook.Worksheets("����").Range("c1:BB1"), 0)
                RATE_TCS.Value = Application.WorksheetFunction.VLookup(30, ThisWorkbook.Worksheets("����").Range("B3:BB10"), k + 1, True)
                
                TOTAL = Val(RATE_TCS.Value)
                If HIRE_COMING_FROM = "" Then
                Else
                HIRE_COMING_TO.Value = Format(DateSerial(Year(DateAdd("m", 1, HIRE_COMING_FROM)), Month(DateAdd("m", 1, HIRE_COMING_FROM)), Day(DateAdd("m", 1, HIRE_COMING_FROM))), "dd.mm.yyyy")
                End If
            Else
            
            If HIRE_COMING_TO.Value = "" Then
            Else
                n = DateDiff("d", HIRE_COMING_TO.Value, DateSerial(Year(Now), Month(Now), Day(Now)))
                If Val(QUANTITY.Value) < n Then QUANTITY.Value = n
            
            End If
                k = Application.WorksheetFunction.Match(BYKE_TYPE.Value, ThisWorkbook.Worksheets("����").Range("c1:II1"), 0)
                RATE_TCS.Value = Application.WorksheetFunction.VLookup(Val(QUANTITY.Value), ThisWorkbook.Worksheets("����").Range("B3:II10"), k + 1, True)
                
                PRICE.Value = Val(RATE_TCS) * Val(PER)
                TOTAL = PRICE * Val(QUANTITY)
                
                If HIRE_COMING_FROM = "" Then
                Else
                HIRE_COMING_TO.Value = Format(DateSerial(Year(DateAdd("d", Val(QUANTITY), HIRE_COMING_FROM)), Month(DateAdd("d", Val(QUANTITY), HIRE_COMING_FROM)), Day(DateAdd("d", Val(QUANTITY), HIRE_COMING_FROM))), "dd.mm.yyyy")
                End If
            End If
        Else
        End If
    End If

End Sub


Private Sub UserForm_Activate()

Dim data_base As Worksheet
Set data_base = ThisWorkbook.Worksheets("������ ����������")

If data_base.Cells(ActiveCell.Row, 1).Value <> "" And Cells(ActiveCell.Row, 7).Value = "������" Then
    Status.Caption = "3"
    ��������_������
    Continie.Enabled = True
    PROVERKA.Enabled = False
    PE4AT.Enabled = True
    vbHelp.Enabled = False
Else

    mes = MsgBox("������� ������� ������ ?", vbInformation + vbOKCancel, "��������� ������� �����")
    Select Case mes
        Case vbOK:
            Status.Caption = "1"
            Continie.Enabled = False
            �����_�������
            PROVERKA_Click
            PROVERKA.Enabled = True
            PE4AT.Enabled = False
            Closer.Enabled = False
            vbHelp.Enabled = True
            vbHelp_Click

        Case vbCancel:
        WHERE_BIKE.Show
    End Select
    
End If

End Sub

Private Sub �����_�������()
Dim data_base As Worksheet
Set data_base = ThisWorkbook.Worksheets("������ ����������")
Continie.Enabled = False
NEW_DOGOVOR.Visible = True
NEW_DOGOVOR.Enabled = False

DOGOVOR.Visible = False
ScrollBar1.Visible = False
If Application.WorksheetFunction.Max(data_base.Range("P4:P300")) >= 10000 Then
NEW_DOGOVOR.Value = Application.WorksheetFunction.Max(data_base.Range("P4:P300"), ThisWorkbook.Worksheets("�����").Range("P1:P20000")) + 1
Else
NEW_DOGOVOR.Value = 10001
End If

Frame1.Enabled = True
Frame2.Enabled = True

For Each ctrl In Me.Controls("Frame1").Controls
    If TypeName(ctrl) = "TextBox" Then ctrl.Value = ""
Next

PERIOD_RATE.clear
PERIOD_RATE.AddItem "����"
PERIOD_RATE.AddItem "�����"

QUANTITY.clear
For i = 1 To 31
QUANTITY.AddItem i
Next i
QUANTITY.AddItem "�����"

TIME_FROM.clear
For i = 10 To 22
    TIME_FROM.AddItem i & ":00"
    TIME_FROM.AddItem i & ":30"
Next i

TIME_TO.clear
For i = 10 To 22
    TIME_TO.AddItem i & ":00"
    TIME_TO.AddItem i & ":30"
Next i

'TIME_FROM.Value = Hour(Now()) & ":00"


VALUTA.clear
VALUTA.AddItem "BAHT"
VALUTA.AddItem "DOLLAR"
VALUTA.AddItem "EURO"

VALUTA.Value = "BAHT"

PERIOD_HIRE.clear
PERIOD_HIRE.AddItem "����"
PERIOD_HIRE.AddItem "�����"


List_of_model BYKE_TYPE
BYKE_TYPE.Value = data_base.Cells(ActiveCell.Row, 1)
BYKE_REG.Value = data_base.Cells(ActiveCell.Row, 2)
BYKE_COLOUR.Value = data_base.Cells(ActiveCell.Row, 3)
TECH_PERIOD.Value = data_base.Cells(ActiveCell.Row, 4)
'TECH_PERIOD.Value =
BYKE_SPEED_FACKT.Value = data_base.Cells(ActiveCell.Row, 5)

RATE_TCS.Value = data_base.Cells(ActiveCell.Row, 11)

If data_base.Cells(ActiveCell.Row, 1) = "" Then
    RATE_TCS.Value = 0
Else
    k = Application.WorksheetFunction.Match(BYKE_TYPE.Value, ThisWorkbook.Worksheets("����").Range("C1:II1"), 0)
    RATE_TCS.Value = Application.WorksheetFunction.VLookup(0, ThisWorkbook.Worksheets("����").Range("B3:II10"), k, False)
End If



'RATE_TCS.Value = Application.WorksheetFunction.VLookup(0, list2, k, True)
PER.Value = 1

If RATE_TCS.Value > 1000 Then
PERIOD_RATE.Value = "�����"
Else
PERIOD_RATE.Value = "����"
End If

'PERIOD_RATE.Value = data_base.Cells(ActiveCell.Row, )
PRICE.Value = RATE_TCS * PER

HIRE_COMING_FROM.Text = Format(DateSerial(Year(Now()), Month(Now()), Day(Now())), "dd.mm.yyyy")
HIRE_COMING_TO.Value = Format(DateSerial(Year(DateAdd("d", Val(QUANTITY), HIRE_COMING_FROM)), Month(DateAdd("d", Val(QUANTITY), HIRE_COMING_FROM)), Day(DateAdd("d", Val(QUANTITY), HIRE_COMING_FROM))), "dd.mm.yyyy")
'HIRE_COMING_FROM.Value = DateSerial(Year(Now()), Month(Now()), Day(Now()))

If Minute(Now()) > 30 Then
    TIME_FROM.Value = Hour(Now()) & ":00"
Else
    TIME_FROM.Value = Hour(Now()) & ":30"
End If

'TIME_FROM.Value = Hour(Now()) & ":" & Minute(Now())
'TIME_FROM.Value = Format(TIME_FROM.Value, "Short Time")
'QUANTITY.Value = 1

'If RATE_TCS.Value > 1000 Then
'    PERIOD_HIRE.Value = "�����"
'    HIRE_COMING_TO.Value = DateSerial(Year(DateAdd("m", QUANTITY, HIRE_COMING_FROM)), Month(DateAdd("m", QUANTITY, HIRE_COMING_FROM)), Day(DateAdd("m", QUANTITY, HIRE_COMING_FROM)))
'Else
'    PERIOD_HIRE.Value = "����"
'    HIRE_COMING_TO.Value = DateSerial(Year(DateAdd("d", QUANTITY, HIRE_COMING_FROM)), Month(DateAdd("d", QUANTITY, HIRE_COMING_FROM)), Day(DateAdd("d", QUANTITY, HIRE_COMING_FROM)))
'End If

'HIRE_COMING_TO.Value = Format(DateSerial(Year(data_base.Cells(ActiveCell.Row, 10)), Month(data_base.Cells(ActiveCell.Row, 10)), Day(data_base.Cells(ActiveCell.Row, 10))), "dd.mm.yyyy")
TIME_TO.Value = TIME_FROM.Value
TIME_TO.Value = Format(TIME_TO.Value, "Short Time")
'TOTAL.Value = data_base.Cells(ActiveCell.Row, )

'If data_base.Cells(ActiveCell.Row, 14).Value <> "��������" Then
    OptionButton1.Value = True

'Else
'    OptionButton2.Value = True
'End If
QUANTITY.Value = 1
Form.Caption = "������� �" & NEW_DOGOVOR.Value & " �� " & HIRE_COMING_FROM

End Sub

Private Sub BYKE_TYPE_Change()
List_of_number BYKE_REG
'PERIOD_HIRE_Change
BYKE_TYPE.BackColor = &H80000005

QUANTITY_Change
'If BYKE_TYPE.Value = "" Then
'
'Else
'k = Application.WorksheetFunction.Match(BYKE_TYPE.Value, ThisWorkbook.Worksheets("����").Range("c1:S1"), 0)
'RATE_TCS.Value = Application.WorksheetFunction.VLookup(Val(QUANTITY.Value), ThisWorkbook.Worksheets("����").Range("B3:S10"), k + 1, True)
'
'End If
'
'If Status.Caption = 1 Then
'PRICE.Value = Val(RATE_TCS.Value) * Val(QUANTITY.Value)
'    PRICE.Value = Val(RATE_TCS) * Val(PER)
'    TOTAL = Val(PRICE) * Val(QUANTITY)
'Else
'End If

End Sub

Private Sub DOGOVOR_Change()
ScrollBar1.Max = DOGOVOR.ListCount - 1
ScrollBar1.Min = 0
ScrollBar1.Value = DOGOVOR.ListIndex
End Sub

Private Sub OptionButton1_Click()
        VALUTA.Value = "BAHT"
'        VALUTA_FIELD.Value = data_base.Cells(ActiveCell.Row, 13).Value
        DOCUMENT.Visible = False
        VALUTA.Visible = True
        VALUTA_FIELD.Visible = True
PROVERKA_Click
End Sub

Private Sub ScrollBar1_Change()
A = DOGOVOR.ListIndex
B = ScrollBar1.Value
C = DOGOVOR.ListCount

DOGOVOR.ListIndex = ScrollBar1.Value
���������_������
End Sub

Private Sub ��������_������()
Dim data_base As Worksheet
Set data_base = ThisWorkbook.Worksheets("������ ����������")

DOGOVOR.Visible = True
ScrollBar1.Visible = True

NEW_DOGOVOR.Visible = False

List_of_DOGOVOR DOGOVOR
DOGOVOR.Value = data_base.Cells(ActiveCell.Row, 16)

ScrollBar1.Max = DOGOVOR.ListCount - 1
ScrollBar1.Min = 0
ScrollBar1.Value = DOGOVOR.ListIndex

PERIOD_RATE.clear
PERIOD_RATE.AddItem "����"
PERIOD_RATE.AddItem "�����"

QUANTITY.clear
For i = 1 To 31
QUANTITY.AddItem i
Next i
QUANTITY.AddItem "�����"


VALUTA.clear
VALUTA.AddItem "BAHT"
VALUTA.AddItem "DOLLAR"
VALUTA.AddItem "EURO"


CUSTOMER_NAME.Value = data_base.Cells(ActiveCell.Row, 17)
AGE.Value = data_base.Cells(ActiveCell.Row, 18)
NATIONALITY.Value = data_base.Cells(ActiveCell.Row, 19)
PASSPORT.Value = data_base.Cells(ActiveCell.Row, 22)
P_ISSUED.Value = data_base.Cells(ActiveCell.Row, 23)
ADDRESS_HOTEL.Value = data_base.Cells(ActiveCell.Row, 20)
ROOM.Value = data_base.Cells(ActiveCell.Row, 21)
TELEFON.Value = data_base.Cells(ActiveCell.Row, 26)
DRIVERS_LICENCE.Value = data_base.Cells(ActiveCell.Row, 24)
DL_ISSUED.Value = data_base.Cells(ActiveCell.Row, 25)
Prime4anie.Value = data_base.Cells(ActiveCell.Row, 27)
    
BYKE_TYPE.Value = data_base.Cells(ActiveCell.Row, 1)
BYKE_REG.Value = data_base.Cells(ActiveCell.Row, 2)
BYKE_COLOUR.Value = data_base.Cells(ActiveCell.Row, 3)
TECH_PERIOD.Value = data_base.Cells(ActiveCell.Row, 4)
BYKE_SPEED_FACKT.Value = data_base.Cells(ActiveCell.Row, 5)

RATE_TCS.Value = data_base.Cells(ActiveCell.Row, 11)
PER.Value = 1

If RATE_TCS.Value > 1000 Then
PERIOD_RATE.Value = "�����"
Else
PERIOD_RATE.Value = "����"
End If

'PERIOD_RATE.Value = data_base.Cells(ActiveCell.Row, )
PRICE.Value = data_base.Cells(ActiveCell.Row, 12)
TOTAL.Value = data_base.Cells(ActiveCell.Row, 12)

HIRE_COMING_FROM.Value = Format(DateSerial(Year(data_base.Cells(ActiveCell.Row, 8)), Month(data_base.Cells(ActiveCell.Row, 8)), Day(data_base.Cells(ActiveCell.Row, 8))), "dd.mm.yyyy")
TIME_FROM.Value = Hour(data_base.Cells(ActiveCell.Row, 8)) & ":" & Minute(data_base.Cells(ActiveCell.Row, 8))
TIME_FROM.Value = Format(TIME_FROM.Value, "Short Time")
QUANTITY.Value = data_base.Cells(ActiveCell.Row, 9)

If RATE_TCS.Value > 1000 Then
PERIOD_HIRE.Value = "�����"
Else
PERIOD_HIRE.Value = "����"
End If

HIRE_COMING_TO.Value = Format(DateSerial(Year(data_base.Cells(ActiveCell.Row, 10)), Month(data_base.Cells(ActiveCell.Row, 10)), Day(data_base.Cells(ActiveCell.Row, 10))), "dd.mm.yyyy")
TIME_TO.Value = Hour(data_base.Cells(ActiveCell.Row, 10)) & ":" & Minute(data_base.Cells(ActiveCell.Row, 10))
TIME_TO.Value = Format(TIME_TO.Value, "Short Time")
'TOTAL.Value = data_base.Cells(ActiveCell.Row, )

If data_base.Cells(ActiveCell.Row, 14).Value <> "��������" Then
    OptionButton1.Value = True
        VALUTA.Value = data_base.Cells(ActiveCell.Row, 14).Value
        VALUTA_FIELD.Value = data_base.Cells(ActiveCell.Row, 13).Value
        DOCUMENT.Visible = False
        VALUTA.Visible = True
        VALUTA_FIELD.Visible = True
Else
    OptionButton2.Value = True
        DOCUMENT.Value = data_base.Cells(ActiveCell.Row, 13).Value
        DOCUMENT.Visible = True
        VALUTA.Visible = False
        VALUTA_FIELD.Visible = False
End If

Form.Caption = "������� �" & DOGOVOR.Value & " �� " & HIRE_COMING_FROM

     A = DateSerial(Year(Now), Month(Now), Day(Now))
     B = DateValue(HIRE_COMING_TO)
     
     ' - ���������� ����� �� ������
     
     If DateValue(HIRE_COMING_TO) >= DateSerial(Year(Now), Month(Now), Day(Now)) Then ' << ===============   ������ ������ ������ �������
        Closer.Enabled = True
        Continie.Enabled = True
     Else
        Closer.Enabled = False
        Continie.Enabled = True
     End If


End Sub

Private Sub ���������_������()
Dim data_base As Worksheet
Set data_base = ThisWorkbook.Worksheets("������ ����������")
Dim dResult As Double

'List_of_DOGOVOR DOGOVOR
'DOGOVOR.Value = data_base.Cells(ActiveCell.Row, 16)

data_base.Range("P:P").Find(What:=DOGOVOR.Value, LookAt:=xlWhole).Select
Label32.Caption = ActiveCell.Row
current_row = Val(Label32.Caption)

CUSTOMER_NAME.Value = data_base.Cells(current_row, 17)
AGE.Value = data_base.Cells(current_row, 18)
NATIONALITY.Value = data_base.Cells(current_row, 19)
PASSPORT.Value = data_base.Cells(current_row, 22)
P_ISSUED.Value = data_base.Cells(current_row, 23)
ADDRESS_HOTEL.Value = data_base.Cells(current_row, 20)
ROOM.Value = data_base.Cells(current_row, 21)
TELEFON.Value = data_base.Cells(current_row, 26)
DRIVERS_LICENCE.Value = data_base.Cells(current_row, 24)
DL_ISSUED.Value = data_base.Cells(current_row, 25)
Prime4anie.Value = data_base.Cells(current_row, 27)
    
BYKE_TYPE.Value = data_base.Cells(current_row, 1)

BYKE_REG.Value = data_base.Cells(current_row, 2)
BYKE_COLOUR.Value = data_base.Cells(current_row, 3)

TECH_PERIOD.Value = data_base.Cells(current_row, 4)
BYKE_SPEED_FACKT.Value = data_base.Cells(current_row, 5)

RATE_TCS.Value = data_base.Cells(current_row, 11)
PER.Value = 1

'PERIOD_HIRE.clear
'PERIOD_HIRE.AddItem "����"
'PERIOD_HIRE.AddItem "�����"

If RATE_TCS.Value > 1000 Then
PERIOD_RATE.Value = "�����"
Else
PERIOD_RATE.Value = "����"
End If

'PERIOD_RATE.Value = data_base.Cells(ActiveCell.Row, )
PRICE.Value = data_base.Cells(current_row, 12)

HIRE_COMING_FROM.Value = Format(DateSerial(Year(data_base.Cells(current_row, 8)), Month(data_base.Cells(current_row, 8)), Day(data_base.Cells(current_row, 8))), "dd.mm.yyyy")
TIME_FROM.Value = Hour(data_base.Cells(current_row, 8)) & ":" & Minute(data_base.Cells(current_row, 8))
TIME_FROM.Value = Format(TIME_FROM.Value, "Short Time")
QUANTITY.Value = data_base.Cells(current_row, 9)

'If RATE_TCS.Value > 1000 Then
'PERIOD_HIRE.Value = "�����"
'Else
'PERIOD_HIRE.Value = "����"
'End If

HIRE_COMING_TO.Value = Format(DateSerial(Year(data_base.Cells(current_row, 10)), Month(data_base.Cells(current_row, 10)), Day(data_base.Cells(current_row, 10))), "dd.mm.yyyy")
TIME_TO.Value = Hour(data_base.Cells(current_row, 10)) & ":" & Minute(data_base.Cells(current_row, 10))
TIME_TO.Value = Format(TIME_TO.Value, "Short Time")
TOTAL.Value = data_base.Cells(current_row, 12)

If data_base.Cells(current_row, 14).Value <> "��������" Then
    OptionButton1.Value = True
        VALUTA.Value = data_base.Cells(current_row, 14).Value
        VALUTA_FIELD.Value = data_base.Cells(current_row, 13).Value
        DOCUMENT.Visible = False
        VALUTA.Visible = True
        VALUTA_FIELD.Visible = True
Else
    OptionButton2.Value = True
        DOCUMENT.Value = data_base.Cells(current_row, 13).Value
        DOCUMENT.Visible = True
        VALUTA.Visible = False
        VALUTA_FIELD.Visible = False
End If

Form.Caption = "������� �" & DOGOVOR.Value & " �� " & HIRE_COMING_FROM
     A = DateSerial(Year(Now), Month(Now), Day(Now))
     B = DateValue(HIRE_COMING_TO)
     
     ' - ���������� ����� �� ������
     
     If DateValue(HIRE_COMING_TO) >= DateSerial(Year(Now), Month(Now), Day(Now)) Then ' << ===============   ������ ������ ������ �������
        Closer.Enabled = True
        Continie.Enabled = True
     Else
        Closer.Enabled = False
        Continie.Enabled = True
     End If

End Sub
Sub clear()
Dim ctrl As Control
    For Each ctrl In Me.Controls("Frame2").Controls
    A = TypeName(ctrl)
        If TypeName(ctrl) = "TextBox" Then
            ctrl.Enabled = False
'            Ctrl.BackColor = &H8000000F
            ctrl.SpecialEffect = effectETCHED
           
        End If
    Next
End Sub


Sub List_of_model(B As MSForms.ComboBox)
Dim i As Long, C As Long, S$, L As Long, R As Long
Dim data_base As Worksheet

Set data_base = ThisWorkbook.Worksheets("������ ����������")

BYKE_TYPE.clear

For i = 4 To data_base.Cells.Rows.Count '��������� � ��������� ������
  S = data_base.Cells(i, 1)             '������� �������
  If S = "" Then Exit For
  L = 0
  R = B.ListCount - 1
  Do While R >= L
    C = (R + L) \ 2
    Select Case StrComp(S, B.List(C))
      Case -1: R = C - 1
      Case 1: L = C + 1
      Case Else: GoTo NextI
    End Select
  Loop
    If data_base.Cells(i, 7).Value <> "������" Then B.AddItem S, L
    
  
NextI:
Next i
End Sub
 
Sub List_of_number(B As MSForms.ComboBox)
Dim i As Long, C As Long, S$, L As Long, R As Long
Dim data_base As Worksheet

Set data_base = ThisWorkbook.Worksheets("������ ����������")

BYKE_REG.clear

For i = 4 To data_base.Cells.Rows.Count '��������� � 4-�� ������
  S = data_base.Cells(i, 2)             '������� �������
  If S = "" Then Exit For
  L = 0
  R = B.ListCount - 1
  Do While R >= L
    C = (R + L) \ 2
    Select Case StrComp(S, B.List(C))
      Case -1: R = C - 1
      Case 1: L = C + 1
      Case Else: GoTo NextI
    End Select
  Loop
    If data_base.Cells(i, 1).Value = BYKE_TYPE.Value And data_base.Cells(i, 7).Value <> "������" Then
    B.AddItem S, L
    Else
    End If
  
NextI:
Next i

If BYKE_REG.ListCount = 1 Then BYKE_REG.ListIndex = 0

End Sub

Sub List_of_DOGOVOR(B As MSForms.ComboBox)
Dim i As Long, C As Long, S$, L As Long, R As Long
Dim data_base As Worksheet

Set data_base = ThisWorkbook.Worksheets("������ ����������")

DOGOVOR.clear

For i = 4 To data_base.Cells.Rows.Count '��������� �� ������ ������
  S = data_base.Cells(i, 16)             '16-��� �������
  If data_base.Cells(i, 1) = "" Then Exit For
  L = 0
  R = B.ListCount - 1
  Do While R >= L
    C = (R + L) \ 2
    Select Case StrComp(S, B.List(C))
      Case -1: R = C - 1
      Case 1: L = C + 1
      Case Else: GoTo NextI
    End Select
  Loop
    If data_base.Cells(i, 7).Value = "������" And data_base.Cells(i, 16).Value <> "��� ������" Then B.AddItem S, L
    
  
NextI:
Next i
End Sub

Sub List_of_COLOR(B As MSForms.ComboBox)
Dim i As Long, C As Long, S$, L As Long, R As Long
Dim data_base As Worksheet

Set data_base = ThisWorkbook.Worksheets("������ ����������")

BYKE_COLOUR.clear

For i = 4 To data_base.Cells.Rows.Count '��������� �� ������ ������
  S = data_base.Cells(i, 3)             '�������� �������
  If S = "" Then Exit For
  L = 0
  R = B.ListCount - 1
  Do While R >= L
    C = (R + L) \ 2
    Select Case StrComp(S, B.List(C))
      Case -1: R = C - 1
      Case 1: L = C + 1
      Case Else: GoTo NextI
    End Select
  Loop
    If data_base.Cells(i, 1).Value = BYKE_TYPE.Value And (data_base.Cells(i, 2).Value = Val(BYKE_REG.Value) Or data_base.Cells(i, 2).Value = BYKE_REG.Value) Then
    B.AddItem S, L
    Else
    End If
  
NextI:
Next i

If BYKE_COLOUR.ListCount = 1 Then BYKE_COLOUR.ListIndex = 0

End Sub


Private Sub ��������_Click()

End Sub




