Public Class clsJEPosting

    Public Const Formtype = "MIPLJEP"
    Dim objForm As SAPbouiCOM.Form
    Dim strSQL As String
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim QCHeader As SAPbouiCOM.DBDataSource
    Dim JELine As SAPbouiCOM.DBDataSource
    Dim objCombobox As SAPbouiCOM.ComboBox
    Dim objRecordSet As SAPbobsCOM.Recordset
    Dim objcheckbox As SAPbouiCOM.CheckBox

    Public Sub LoadScreen()
        objForm = objAddOn.objUIXml.LoadScreenXML("JEPosting.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, Formtype)
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "select distinct(U_ExpType) from [@MIPLPREP]"
        objRecordSet.DoQuery(strSQL)
        objForm.Items.Item("14").Enabled = False
        While Not objRecordSet.EoF
            objCombobox = objForm.Items.Item("4").Specific
            objCombobox.ValidValues.Add(objRecordSet.Fields.Item("U_ExpType").Value, objRecordSet.Fields.Item("U_ExpType").Value)
            objRecordSet.MoveNext()
        End While
        JELine = objForm.DataSources.DBDataSources.Item("@MIPLJEP1")
        objForm.Items.Item("6").Specific.Active = True
        objForm.Items.Item("6").Specific.String = "A"
        objForm.Items.Item("12").Specific.Active = True
        objForm.Items.Item("12").Specific.String = "A"
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent, ByVal BubbleEvent As Boolean)

        If pVal.BeforeAction = True Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
            End Select
        Else
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "1" And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then

                    End If
                    If pVal.ItemUID = "101" Then
                        JEPosting(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    LoadExisitingData(FormUID)
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objMatrix = objForm.Items.Item("7").Specific
                    If pVal.ItemUID = "7" And pVal.ColUID = "15" Then
                        validate(FormUID)
                        CalculateAmount(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "7" And pVal.ColUID = "1" And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        CalculateTotalAmount(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    If pVal.ItemUID = "7" And pVal.ColUID = "1" Then
                        CheckAllRows(FormUID)
                    End If
            End Select
        End If
    End Sub

    Private Sub CalculateTotalAmount(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("7").Specific
        Dim totalamount As Double = 0
        For i As Integer = 1 To objMatrix.RowCount
            If objMatrix.Columns.Item("1").Cells.Item(i).Specific.checked Then
                totalamount = totalamount + CDbl(objMatrix.Columns.Item("21").Cells.Item(i).Specific.String)
            End If
        Next
        objForm.Items.Item("14").Specific.String = totalamount
    End Sub
    Private Sub CheckAllRows(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("7").Specific
        For i As Integer = 1 To objMatrix.RowCount
            objcheckbox = objMatrix.Columns.Item("1").Cells.Item(i).Specific
            objcheckbox.Checked = True
        Next
    End Sub
    Private Function validate(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("7").Specific
        Dim x As Date, y As Date
        'Dim i As Integer
        'i = Date.Compare(y, x)
        'MsgBox(i)
        For i = 1 To objMatrix.RowCount
            x = objAddOn.objGenFunc.GetDateTimeValue(objMatrix.Columns.Item("15").Cells.Item(i).Specific.String)
            y = objAddOn.objGenFunc.GetDateTimeValue(objMatrix.Columns.Item("11").Cells.Item(i).Specific.String)
            ' If CDate(y).ToString("ddMMyyyy") < CDate(x).ToString("ddMMyyyy") Then
            If y < x Then
                objAddOn.objApplication.SetStatusBarMessage("Seems to be Negative Days")
                Return False
            End If
        Next i
        Return True
    End Function
    Private Function JEPosting(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("7").Specific
        Dim objJE As SAPbobsCOM.JournalEntries
        objJE = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
        For i = 1 To objMatrix.RowCount
            If objMatrix.Columns.Item("1").Cells.Item(i).Specific.checked And objMatrix.Columns.Item("22").Cells.Item(i).Specific.string = "O" Then
                objJE.ReferenceDate = objAddOn.objGenFunc.GetDateTimeValue(objForm.Items.Item("12").Specific.String)
                objJE.DueDate = objAddOn.objGenFunc.GetDateTimeValue(objForm.Items.Item("12").Specific.String)
                objJE.TaxDate = objAddOn.objGenFunc.GetDateTimeValue(objForm.Items.Item("12").Specific.String)
                objJE.Memo = objMatrix.Columns.Item("2").Cells.Item(i).Specific.String
                objJE.Lines.BPLID = 1
                objJE.Lines.AccountCode = objMatrix.Columns.Item("5").Cells.Item(i).Specific.string ' debit gl
                objJE.Lines.ContraAccount = objMatrix.Columns.Item("3").Cells.Item(i).Specific.string ' credit gl
                objJE.Lines.Credit = 0
                objJE.Lines.Debit = objMatrix.Columns.Item("21").Cells.Item(i).Specific.string


                objJE.Lines.ShortName = objMatrix.Columns.Item("5").Cells.Item(i).Specific.string ' debit gl
                objJE.Lines.Add()
                objJE.Lines.SetCurrentLine(1)
                objJE.Lines.BPLID = 1
                objJE.Lines.AccountCode = objMatrix.Columns.Item("3").Cells.Item(i).Specific.string ' credit gl
                objJE.Lines.ContraAccount = objMatrix.Columns.Item("5").Cells.Item(i).Specific.string ' credit gl
                objJE.Lines.Credit = objMatrix.Columns.Item("21").Cells.Item(i).Specific.string
                objJE.Lines.Debit = 0

                ''objJE.Lines.Line_ID = 1
                objJE.Lines.ShortName = objMatrix.Columns.Item("3").Cells.Item(i).Specific.string ' credit gl

                If (0 <> objJE.Add()) Then
                    objAddOn.objApplication.SetStatusBarMessage(objAddOn.objCompany.GetLastErrorDescription)
                Else
                    objMatrix.Columns.Item("21A").Cells.Item(i).Specific.string = CStr(objAddOn.objCompany.GetNewObjectKey())
                    ' objJE.SaveXML("c:\temp\JournalEntries" + Str(vJE.JdtNum) + ".xml")
                End If
            End If
        Next
        Return True
    End Function

    Private Sub CalculateAmount(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("7").Specific
        JELine = objForm.DataSources.DBDataSources.Item("@MIPLJEP1")
        Dim startdate As Date, fromdate As Date
        Dim enddate As Date, todate As Date
        Dim i As Integer, totaldays As Integer, balancedays As Integer, rentdays As Integer
        Dim prepaiddays As Integer
        Dim expensetodate As Double, outstandrent As Double, prepaidrent As Double, postoutamt As Double
        For i = 1 To objMatrix.RowCount
            startdate = objAddOn.objGenFunc.GetDateTimeValue(objMatrix.Columns.Item("10").Cells.Item(i).Specific.String)
            enddate = objAddOn.objGenFunc.GetDateTimeValue(objMatrix.Columns.Item("11").Cells.Item(i).Specific.String)
            fromdate = objAddOn.objGenFunc.GetDateTimeValue(objMatrix.Columns.Item("14").Cells.Item(i).Specific.String)
            todate = objAddOn.objGenFunc.GetDateTimeValue(objMatrix.Columns.Item("15").Cells.Item(i).Specific.String)
           
            If objMatrix.Columns.Item("15").Cells.Item(i).Specific.String <> "" Then
                totaldays = CInt(DateDiff(DateInterval.Day, startdate, fromdate))
                Dim TotalAmt As Double = CDbl(objMatrix.Columns.Item("9").Cells.Item(i).Specific.String)
                Dim Noofdays As Integer = CInt(objMatrix.Columns.Item("12").Cells.Item(i).Specific.String)

                expensetodate = CDbl(TotalAmt / Noofdays * totaldays)
                outstandrent = CDbl(TotalAmt - expensetodate)
                balancedays = CInt(DateDiff(DateInterval.Day, fromdate, enddate))
                rentdays = CInt(DateDiff(DateInterval.Day, fromdate, todate))
                prepaiddays = CInt(balancedays - rentdays)
                postoutamt = CDbl(outstandrent / balancedays * rentdays)
                prepaidrent = CDbl(outstandrent - postoutamt)
                objMatrix.GetLineData(i)
                JELine.SetValue("U_ExpAmnt", 0, expensetodate)
                JELine.SetValue("U_OutstdRnt", 0, outstandrent)
                JELine.SetValue("U_Blncdays", 0, balancedays)
                JELine.SetValue("U_ExpDays", 0, rentdays)
                JELine.SetValue("U_PrpdDays", 0, prepaiddays)
                JELine.SetValue("U_PrpdRent", 0, prepaidrent)
                JELine.SetValue("U_PostAmt", 0, postoutamt)
                objMatrix.SetLineData(i)

                'objMatrix.Columns.Item("13").Cells.Item(i).Specific.String = expensetodate
                'objMatrix.Columns.Item("16").Cells.Item(i).Specific.String = outstandrent
                'objMatrix.Columns.Item("17").Cells.Item(i).Specific.String = balancedays
                'objMatrix.Columns.Item("18").Cells.Item(i).Specific.String = rentdays
                'objMatrix.Columns.Item("19").Cells.Item(i).Specific.String = prepaiddays
                'objMatrix.Columns.Item("21").Cells.Item(i).Specific.String = postoutamt
                'objMatrix.Columns.Item("20").Cells.Item(i).Specific.String = prepaidrent                

            End If
        Next i

        ' objForm.Update()
    End Sub

    Public Sub LoadExisitingData(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("7").Specific
        Dim i As Integer

        strSQL = " select T0.U_CreditGL,T0.U_CreditGL1,T0.U_DebitGL,T0.U_DebitGL1,T0.DocEntry,T0.LineId,T0.U_Desc,max(T1.U_PrpdRent) PrpdRent,T0.U_RentAmt,T0.U_NoofDays," & _
                  "T0.U_FromDate,T0.U_ToDate,Max(T1.U_LstExDte) FromDate, sum(T1.U_PostAmt) ExpenseToDate, min(T1.U_Blncdays) Blncdays, min(T1.U_prpddays) prpddays, max(T1.U_expdays) expdays " & _
             " from [@MIPLPREP1] T0 join [@MIPLPREP] T2 on T0.DocEntry =T2.DocEntry  " & _
             " left outer join [@MIPLJEP1] T1 on T1.U_BaseEntry=T0.DocEntry  and T1.U_BaseLineNum =T0.LineId " & _
             " where T2.U_ExpType = '" & objForm.Items.Item("4").Specific.selected.value & "' " & _
             " group by T0.U_CreditGL,T0.U_CreditGL1,T0.U_DebitGL,T0.U_DebitGL1,T0.U_Desc,T0.U_RentAmt,T0.U_NoofDays,T0.U_FromDate,T0.U_ToDate,T0.DocEntry,T0.LineId"

        objRecordSet.DoQuery(strSQL)
        objMatrix.Clear()
        objForm.Items.Item("1000002").Specific.String = objRecordSet.Fields.Item("DocEntry").Value
        While Not objRecordSet.EoF
            objMatrix.AddRow()
            JELine.Clear()
            objMatrix.GetLineData(objMatrix.RowCount)
            JELine.SetValue("U_Desc", 0, objRecordSet.Fields.Item("U_Desc").Value)
            JELine.SetValue("U_CreditGL", 0, objRecordSet.Fields.Item("U_CreditGL").Value)
            JELine.SetValue("U_CreditGL1", 0, objRecordSet.Fields.Item("U_CreditGL1").Value)
            JELine.SetValue("U_DebitGL", 0, objRecordSet.Fields.Item("U_DebitGL").Value)
            JELine.SetValue("U_DebitGL1", 0, objRecordSet.Fields.Item("U_DebitGL1").Value)
            JELine.SetValue("U_BaseEntry", 0, objRecordSet.Fields.Item("DocEntry").Value)
            JELine.SetValue("U_BaseLineNum", 0, objRecordSet.Fields.Item("LineId").Value)
            JELine.SetValue("U_RentAmt", 0, objRecordSet.Fields.Item("U_RentAmt").Value)
            JELine.SetValue("U_FROMDate", 0, Format(CDate(objRecordSet.Fields.Item("U_FromDate").Value), "yyyyMMdd"))
            JELine.SetValue("U_ToDate", 0, Format(CDate(objRecordSet.Fields.Item("U_ToDate").Value), "yyyyMMdd"))
            JELine.SetValue("U_NoofDays", 0, objRecordSet.Fields.Item("U_NoofDays").Value)
            JELine.SetValue("U_ExpAmnt", 0, objRecordSet.Fields.Item("ExpenseToDate").Value)
            'objAddOn.WriteSMSLog(CStr(objRecordSet.Fields.Item("FromDate").Value))
            If CStr((objRecordSet.Fields.Item("FromDate").Value)) = "00:00:00" Then
                JELine.SetValue("U_FrmDate", 0, Format(CDate(objRecordSet.Fields.Item("U_FromDate").Value), "yyyyMMdd")) 'Format(CDate(objRecordSet.Fields.Item("U_lstExDte").Value), "yyyyMMdd"))
                JELine.SetValue("U_Status", 0, "O")
            Else
                JELine.SetValue("U_FrmDate", 0, Format(CDate(objRecordSet.Fields.Item("FromDate").Value), "yyyyMMdd"))
                JELine.SetValue("U_Status", 0, "O")
            End If

            If CDate(objRecordSet.Fields.Item("U_ToDate").Value) = (CDate(objRecordSet.Fields.Item("FromDate").Value)) Then
                JELine.SetValue("U_Status", 0, "C")
            End If
            'JELine.SetValue("U_OutstdRnt", 0, objRecordSet.Fields.Item("OutstdRnt").Value)  
            'JELine.SetValue("U_Blncdays", 0, objRecordSet.Fields.Item("Blncdays").Value)
            'JELine.SetValue("U_ExpDays", 0, objRecordSet.Fields.Item("expdays").Value)
            'JELine.SetValue("U_PrpdDays", 0, objRecordSet.Fields.Item("prpddays").Value)
            'JELine.SetValue("U_PrpdRent", 0, objRecordSet.Fields.Item("PrpdRent").Value)
            'JELine.SetValue("U_LstExDte", 0, objForm.Items.Item("6").Specific.String)
            objMatrix.SetLineData(objMatrix.RowCount)
            objRecordSet.MoveNext()
        End While
        For i = 1 To objMatrix.RowCount
            If objMatrix.Columns.Item("22").Cells.Item(i).Specific.String = "C" Then 'Status
                DisableRows(FormUID, i)
            End If
        Next i

    End Sub
    Private Sub DisableRows(ByVal FormUID As String, ByVal RowID As Integer)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("7").Specific
        objMatrix.CommonSetting.SetCellEditable(RowID, 1, False)
        objMatrix.CommonSetting.SetCellEditable(RowID, 15, False)
        objMatrix.CommonSetting.SetCellEditable(RowID, 16, False)
        objMatrix.CommonSetting.SetCellEditable(RowID, 17, False)
        objMatrix.CommonSetting.SetCellEditable(RowID, 18, False)
        objMatrix.CommonSetting.SetCellEditable(RowID, 19, False)
        objMatrix.CommonSetting.SetCellEditable(RowID, 20, False)
        objMatrix.CommonSetting.SetCellEditable(RowID, 21, False)
        objMatrix.CommonSetting.SetCellEditable(RowID, 22, False)
    End Sub
    Private Sub UpdateStatusClosed(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("7").Specific
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "update [@MIPLJEP1] set U_Status='C' where U_Select ='Y'"
        objRecordSet.DoQuery(strSQL)

    End Sub
End Class
