Public Class clsJEPostingPro

    Public Const Formtype = "MIPLJEPO"
    Dim objForm As SAPbouiCOM.Form
    Dim strSQL As String
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim QCHeader As SAPbouiCOM.DBDataSource
    Dim JELine As SAPbouiCOM.DBDataSource
    Dim objCombobox As SAPbouiCOM.ComboBox
    Dim objRecordSet As SAPbobsCOM.Recordset
    Dim objcheckbox As SAPbouiCOM.CheckBox
    Public Sub LoadScreen()
        objForm = objAddOn.objUIXml.LoadScreenXML("JEPostingPro.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, Formtype)
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "select distinct(U_ExpType) from [@MIPLPROV]"
        objRecordSet.DoQuery(strSQL)
        objForm.Items.Item("13").Enabled = False
        While Not objRecordSet.EoF
            objCombobox = objForm.Items.Item("4").Specific
            objCombobox.ValidValues.Add(objRecordSet.Fields.Item("U_ExpType").Value, objRecordSet.Fields.Item("U_ExpType").Value)
            objRecordSet.MoveNext()
        End While
        JELine = objForm.DataSources.DBDataSources.Item("@MIPLJEPO1")
        objForm.Items.Item("6").Specific.Active = True
        objForm.Items.Item("6").Specific.String = "A"
        objForm.Items.Item("8").Specific.Active = True
        objForm.Items.Item("8").Specific.String = "A"
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent, ByVal BubbleEvent As Boolean)
        If pVal.BeforeAction = True Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK

            End Select
        Else
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "15" Then
                        JEPosting(FormUID)
                    ElseIf pVal.ItemUID = "16" Then
                        LoadDate(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    LoadExisitingData(FormUID)

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objMatrix = objForm.Items.Item("11").Specific
                    If pVal.ItemUID = "11" And pVal.ColUID = "10" Then
                        CalculatePostout(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "11" And pVal.ColUID = "1" Then
                        CalculateTotalAmount(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    If pVal.ItemUID = "11" And pVal.ColUID = "1" Then
                        CheckAllRows(FormUID)
                    End If
            End Select
        End If
    End Sub
    Private Sub LoadDate(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("11").Specific
        Dim startdate As Date
        startdate = objAddOn.objGenFunc.GetDateTimeValue(objMatrix.Columns.Item("10").Cells.Item(1).Specific.String)
        For i = 1 To objMatrix.RowCount
            objMatrix.Columns.Item("10").Cells.Item(i).Specific.String = CDate(startdate).ToString("ddMMyyyy") 'DateTime.Parse(startdate)
        Next
    End Sub
    Private Sub CheckAllRows(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("11").Specific
        For i As Integer = 1 To objMatrix.RowCount
            objcheckbox = objMatrix.Columns.Item("1").Cells.Item(i).Specific
            objcheckbox.Checked = True
        Next
    End Sub
    Public Sub LoadExisitingData(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("11").Specific
        strSQL = "select T0.U_EMPCode, T0.U_EMPName,T0.U_CreditGL,T0.U_BalDate,max(T1.U_BalDate) LastpaidDate,T0.U_LastPaidTill,T0.U_CreditGL1,T0.U_DebitGL,T0.DocEntry,T0.LineId,T0.U_DebitGL1,T0.U_Basic " & _
                 "from [@MIPLPROV1] T0 join [@MIPLPROV] T2 on T0.DocEntry =T2.DocEntry left outer join [@MIPLJEPO1] T1 on T1.U_BaseEntry=T0.DocEntry  and T1.U_BaseLineNum =T0.LineId " & _
                 "where T2.U_ExpType='" & objForm.Items.Item("4").Specific.selected.value & "'" & _
                 "group by T0.U_EMPCode, T0.U_EMPName,T0.U_CreditGL,T0.U_CreditGL1,T0.U_DebitGL,T0.U_DebitGL1,T0.U_Basic,T0.U_LastPaidTill,T0.DocEntry,T0.LineId,T0.U_BalDate"
        objRecordSet.DoQuery(strSQL)
        objMatrix.Clear()
        objForm.Items.Item("10").Specific.String = objRecordSet.Fields.Item("DocEntry").Value
        While Not objRecordSet.EoF
            objMatrix.AddRow()
            JELine.Clear()
            objMatrix.GetLineData(objMatrix.RowCount)
            JELine.SetValue("U_EMPCode", 0, objRecordSet.Fields.Item("U_EMPCode").Value)
            JELine.SetValue("U_EMPName", 0, objRecordSet.Fields.Item("U_EMPName").Value)
            JELine.SetValue("U_CreditGL", 0, objRecordSet.Fields.Item("U_CreditGL").Value)
            JELine.SetValue("U_CreditGL1", 0, objRecordSet.Fields.Item("U_CreditGL1").Value)
            JELine.SetValue("U_DebitGL", 0, objRecordSet.Fields.Item("U_DebitGL").Value)
            JELine.SetValue("U_DebitGL1", 0, objRecordSet.Fields.Item("U_DebitGL1").Value)
            JELine.SetValue("U_BaseEntry", 0, objRecordSet.Fields.Item("DocEntry").Value)
            JELine.SetValue("U_BaseLineNum", 0, objRecordSet.Fields.Item("LineId").Value)
            JELine.SetValue("U_Basic", 0, objRecordSet.Fields.Item("U_Basic").Value)
            JELine.SetValue("U_FromDate", 0, Format(CDate(objRecordSet.Fields.Item("U_BalDate").Value), "yyyyMMdd"))
            'objAddOn.WriteSMSLog(CStr(objRecordSet.Fields.Item("LastpaidDate").Value))
            If CStr((objRecordSet.Fields.Item("LastpaidDate").Value)) = "00:00:00" Then
                JELine.SetValue("U_LastPaidTill", 0, Format(CDate(objRecordSet.Fields.Item("U_LastPaidTill").Value), "yyyyMMdd"))
            Else
                JELine.SetValue("U_LastPaidTill", 0, Format(CDate(objRecordSet.Fields.Item("LastpaidDate").Value), "yyyyMMdd"))
            End If

            objMatrix.SetLineData(objMatrix.RowCount)
            objRecordSet.MoveNext()
        End While

    End Sub
    Private Sub CalculateTotalAmount(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("11").Specific
        Dim totalamount As Double = 0
        For i As Integer = 1 To objMatrix.RowCount
            If objMatrix.Columns.Item("1").Cells.Item(i).Specific.checked Then
                totalamount = totalamount + CDbl(objMatrix.Columns.Item("12").Cells.Item(i).Specific.String)
            End If
        Next
        objForm.Items.Item("13").Specific.String = totalamount
    End Sub

    Private Function JEPosting(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("11").Specific
        Dim objJE As SAPbobsCOM.JournalEntries
        objJE = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
        For i = 1 To objMatrix.RowCount
            If objMatrix.Columns.Item("1").Cells.Item(i).Specific.checked And objMatrix.Columns.Item("13").Cells.Item(i).Specific.string = "O" Then
                objJE.ReferenceDate = objAddOn.objGenFunc.GetDateTimeValue(objForm.Items.Item("8").Specific.String)
                objJE.DueDate = objAddOn.objGenFunc.GetDateTimeValue(objForm.Items.Item("8").Specific.String)
                objJE.TaxDate = objAddOn.objGenFunc.GetDateTimeValue(objForm.Items.Item("8").Specific.String)
                objJE.Memo = objMatrix.Columns.Item("2").Cells.Item(i).Specific.String
                objJE.Lines.BPLID = 1
                objJE.Lines.AccountCode = objMatrix.Columns.Item("6").Cells.Item(i).Specific.string ' debit gl
                objJE.Lines.ContraAccount = objMatrix.Columns.Item("4").Cells.Item(i).Specific.string ' credit gl
                objJE.Lines.Credit = 0
                objJE.Lines.Debit = objMatrix.Columns.Item("12").Cells.Item(i).Specific.string


                objJE.Lines.ShortName = objMatrix.Columns.Item("6").Cells.Item(i).Specific.string ' debit gl
                objJE.Lines.Add()
                objJE.Lines.SetCurrentLine(1)
                objJE.Lines.BPLID = 1
                objJE.Lines.AccountCode = objMatrix.Columns.Item("4").Cells.Item(i).Specific.string ' credit gl
                objJE.Lines.ContraAccount = objMatrix.Columns.Item("6").Cells.Item(i).Specific.string ' debit gl
                objJE.Lines.Credit = objMatrix.Columns.Item("12").Cells.Item(i).Specific.string
                objJE.Lines.Debit = 0

                ''objJE.Lines.Line_ID = 1
                objJE.Lines.ShortName = objMatrix.Columns.Item("4").Cells.Item(i).Specific.string ' credit gl
                'objJE.SaveToFile("E:\ItemCreationAddonProject\je.xml")
                If (0 <> objJE.Add()) Then
                    objAddOn.objApplication.SetStatusBarMessage(objAddOn.objCompany.GetLastErrorDescription)
                Else
                    objMatrix.Columns.Item("12A").Cells.Item(i).Specific.string = CStr(objAddOn.objCompany.GetNewObjectKey())
                    ' objJE.SaveXML("c:\temp\JournalEntries" + Str(vJE.JdtNum) + ".xml")
                End If
            End If
        Next

        Return True
    End Function

    Private Sub CalculatePostout(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("11").Specific
        Dim postoutamt As Double
        Dim startdate As Date, enddate As Date
        Dim noofdays As Integer
        For i = 1 To objMatrix.RowCount
            startdate = objAddOn.objGenFunc.GetDateTimeValue(objMatrix.Columns.Item("9").Cells.Item(i).Specific.String)
            enddate = objAddOn.objGenFunc.GetDateTimeValue(objMatrix.Columns.Item("10").Cells.Item(i).Specific.String)

            If objMatrix.Columns.Item("10").Cells.Item(i).Specific.String <> "" Then
                noofdays = CInt(DateDiff(DateInterval.Day, startdate, enddate))
                objMatrix.Columns.Item("11").Cells.Item(i).Specific.String = noofdays
                'objMatrix.Columns.Item("11").Cells.Item(i).Specific.String = noofdays
                noofdays = CInt(objMatrix.Columns.Item("11").Cells.Item(i).Specific.String)
                Dim basic As Double = CDbl(objMatrix.Columns.Item("8").Cells.Item(i).Specific.String)
                postoutamt = CDbl((basic * noofdays) / 365)
                objMatrix.GetLineData(i)
                JELine.SetValue("U_Postout", 0, postoutamt)
                objMatrix.SetLineData(i)
                'objMatrix.Columns.Item("12").Cells.Item(i).Specific.String = postoutamt
            End If
        Next

    End Sub
End Class
