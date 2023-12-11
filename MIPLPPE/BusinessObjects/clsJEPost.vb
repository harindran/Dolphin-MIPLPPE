Public Class clsJEPost
    Public Const Formtype = "MIPLJEP"
    Dim objForm As SAPbouiCOM.Form
    Dim strSQL As String
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim QCHeader As SAPbouiCOM.DBDataSource
    Dim JELine As SAPbouiCOM.DBDataSource
    Dim objCombobox As SAPbouiCOM.ComboBox
    Dim objRecordSet As SAPbobsCOM.Recordset
   'sDate.Substring(2, 1)
    Public Sub LoadScreen()
        objForm = objAddOn.objUIXml.LoadScreenXML("JE Posting.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, Formtype)

        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "select Code,Name from [@MIPLEX]"
        objRecordSet.DoQuery(strSQL)
        While Not objRecordSet.EoF

            objCombobox = objForm.Items.Item("6").Specific
            objCombobox.ValidValues.Add(objRecordSet.Fields.Item("Code").Value, objRecordSet.Fields.Item("Name").Value)
            objRecordSet.MoveNext()
        End While
        JELine = objForm.DataSources.DBDataSources.Item("@MIPLJE1")
    End Sub
    Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        Try

      
        If pVal.BeforeAction Then
        Else
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "101" Then

                            JEPosting(FormUID)
                        ElseIf pVal.ItemUID = "9" Then
                            LoadMatrix(FormUID)
                        End If

            End Select
        End If
        'select T0.* from [@MIPLPPE1] T1 join [@MIPLOPPE] T0 on T1.DocEntry=T0.DocEntry where  T0.U_ExType ='Visa' and T1.U_DocDate <='20180531' 
        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(ex.ToString)
        End Try
    End Sub
    Private Function JEPosting(ByVal FormUID As String)
        'update status in input screen
        Try

       
        Dim objJE As SAPbobsCOM.JournalEntries
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("10").Specific
           

        For intloop As Integer = 1 To objMatrix.RowCount
                If objMatrix.Columns.Item("1").Cells.Item(intloop).Specific.checked And objMatrix.Columns.Item("9").Cells.Item(intloop).Specific.string = "O" Then
                    objJE = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                    objJE.ReferenceDate = objAddOn.objGenFunc.GetDateTimeValue(objMatrix.Columns.Item("3").Cells.Item(intloop).Specific.string)
                    objJE.DueDate = objAddOn.objGenFunc.GetDateTimeValue(objMatrix.Columns.Item("3").Cells.Item(intloop).Specific.string)
                    objJE.TaxDate = objAddOn.objGenFunc.GetDateTimeValue(objMatrix.Columns.Item("3").Cells.Item(intloop).Specific.string)
                    objJE.Lines.BPLID = 1
                    objJE.Lines.AccountCode = objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string ' debit gl
                    objJE.Lines.ContraAccount = objMatrix.Columns.Item("4").Cells.Item(intloop).Specific.string ' credit gl
                    objJE.Lines.Credit = 0
                    objJE.Lines.Debit = objMatrix.Columns.Item("8").Cells.Item(intloop).Specific.string

               
                    objJE.Lines.ShortName = objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string ' debit gl
                    objJE.Lines.Add()
                    objJE.Lines.SetCurrentLine(1)
                    objJE.Lines.BPLID = 1
                    objJE.Lines.AccountCode = objMatrix.Columns.Item("4").Cells.Item(intloop).Specific.string ' credit gl
                    objJE.Lines.ContraAccount = objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string ' credit gl
                    objJE.Lines.Credit = objMatrix.Columns.Item("8").Cells.Item(intloop).Specific.string
                    objJE.Lines.Debit = 0

                    ''objJE.Lines.Line_ID = 1
                  
                    objJE.Lines.ShortName = objMatrix.Columns.Item("4").Cells.Item(intloop).Specific.string ' credit gl

                    If (0 <> objJE.Add()) Then
                        objAddOn.objApplication.SetStatusBarMessage(objAddOn.objCompany.GetLastErrorDescription)
                    Else
                        objAddOn.objApplication.SetStatusBarMessage("Succeeded in adding a journal entry", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        strSQL = "update [@MIPLPPE1] set U_Status='C' , U_DocNum='" & objAddOn.objCompany.GetNewObjectKey() & "' where  DocEntry='" & objMatrix.Columns.Item("10").Cells.Item(intloop).Specific.string & "' and LineId=" & CInt(objMatrix.Columns.Item("11").Cells.Item(intloop).Specific.string) & " and U_Status='O'"

                        objRecordSet.DoQuery(strSQL)
                        objRecordSet = Nothing
                        objMatrix.Columns.Item("9").Cells.Item(intloop).Specific.string = "C"
                        ' objJE.SaveXML("c:\temp\JournalEntries" + Str(vJE.JdtNum) + ".xml")
                    End If
                End If
        Next


        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(ex.Message)
            Return False
        End Try
        Return True
    End Function
    Private Sub LoadMatrix(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("10").Specific
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim Status As String = objForm.Items.Item("8").Specific.string
        Try
            strSQL = "select T1.DocEntry,T1.LineId,T1.U_DocDate,T0.U_CreditGL,T0.U_CreditGL1,T1.U_DebitGL,T1.U_DebitGL1,T1.U_TotAmt,T1.U_Status from [@MIPLPPE1] T1 join [@MIPLOPPE] T0 on T1.DocEntry=T0.DocEntry where T1.U_Status='" & Status & "' and T0.U_ExType ='" & objForm.Items.Item("6").Specific.selected.value & "' and T1.U_DocDate <='" & objAddOn.objGenFunc.GetDateForDDMMYYYY(objForm.Items.Item("4").Specific.String) & "' "

        Catch ex As Exception
            strSQL = "select T1.DocEntry,T1.LineId,T1.U_DocDate,T0.U_CreditGL,T0.U_CreditGL1,T1.U_DebitGL,T1.U_DebitGL1,T1.U_TotAmt,T1.U_Status from [@MIPLPPE1] T1 join [@MIPLOPPE] T0 on T1.DocEntry=T0.DocEntry where T1.U_Status='" & Status & "' and T0.U_ExType ='" & objForm.Items.Item("6").Specific.selected.value & "' and T1.U_DocDate <='" & objAddOn.objGenFunc.GetDateTimeValue(objForm.Items.Item("4").Specific.String) & "' "
        End Try

        objRecordSet.DoQuery(strSQL)

        While Not objRecordSet.EoF
            If objMatrix.RowCount = 0 Then
                objMatrix.AddRow()
            ElseIf objMatrix.Columns.Item("2").Cells.Item(objMatrix.RowCount).Specific.String <> "" Then
                objMatrix.AddRow()

            End If
            JELine.Clear()

            objMatrix.GetLineData(objMatrix.RowCount)
            JELine.SetValue("U_Type", 0, objForm.Items.Item("6").Specific.selected.value)
            JELine.SetValue("U_PostingDate", 0, Format(CDate(objRecordSet.Fields.Item("U_DocDate").Value), "yyyyMMdd"))
            JELine.SetValue("U_CreditGL", 0, objRecordSet.Fields.Item("U_CreditGL").Value)
            JELine.SetValue("U_CreditGL1", 0, objRecordSet.Fields.Item("U_CreditGL1").Value)
            JELine.SetValue("U_DebitGL", 0, objRecordSet.Fields.Item("U_DebitGL").Value)
            JELine.SetValue("U_DebitGL1", 0, objRecordSet.Fields.Item("U_DebitGL1").Value)
            JELine.SetValue("U_Amount", 0, objRecordSet.Fields.Item("U_TotAmt").Value)
            JELine.SetValue("U_Status", 0, objRecordSet.Fields.Item("U_Status").Value)
            JELine.SetValue("U_BaseRef", 0, objRecordSet.Fields.Item("DocEntry").Value)
            JELine.SetValue("U_BaseLinNum", 0, objRecordSet.Fields.Item("LineId").Value)
            objMatrix.SetLineData(objMatrix.RowCount)
            objRecordSet.MoveNext()
        End While
    End Sub
End Class
