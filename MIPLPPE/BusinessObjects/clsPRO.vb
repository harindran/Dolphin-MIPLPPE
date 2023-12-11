Public Class clsPRO
    Public Const Formtype = "MIPLPRO"
    Dim objForm As SAPbouiCOM.Form
    Dim strSQL As String
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim QCHeader As SAPbouiCOM.DBDataSource
    Dim objCombobox As SAPbouiCOM.ComboBox
    Dim objRecordSet As SAPbobsCOM.Recordset
    Dim PROLine As SAPbouiCOM.DBDataSource
    Dim PROHeader As SAPbouiCOM.DBDataSource

    Public Sub LoadScreen()
        objForm = objAddOn.objUIXml.LoadScreenXML("Provision.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, Formtype)
        objForm.AutoManaged = True
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "select Code,Name,U_AppType from [@MIPLEX]  where U_PrePro='PRO'"
        objRecordSet.DoQuery(strSQL)
        While Not objRecordSet.EoF
            objCombobox = objForm.Items.Item("4").Specific
            objCombobox.ValidValues.Add(objRecordSet.Fields.Item("Code").Value, objRecordSet.Fields.Item("Name").Value)
            objCombobox = objForm.Items.Item("4B").Specific
            objCombobox.ValidValues.Add(objRecordSet.Fields.Item("U_AppType").Value, objRecordSet.Fields.Item("U_AppType").Value)
            objRecordSet.MoveNext()
        End While
        objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        objForm.Items.Item("8").Specific.Active = True
        objForm.Items.Item("8").Specific.String = "A"
        PROLine = objForm.DataSources.DBDataSources.Item("@MIPLPROV1")
        PROHeader = objForm.DataSources.DBDataSources.Item("@MIPLPROV")
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent, ByVal BubbleEvent As Boolean)
        If pVal.BeforeAction = True Then
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        RemoveEmptyRows(FormUID)
                    End If
                    If pVal.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

                    End If
            End Select
        Else
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "2A" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        AddNewRow(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    LoadLedger(FormUID)
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objMatrix = objForm.Items.Item("12").Specific
                    If pVal.ItemUID = "8" Then
                        objMatrix.AddRow(1, pVal.Row)

                    ElseIf pVal.ItemUID = "12" And pVal.ColUID = "9" Then
                        addRow(FormUID)
                        'objMatrix.AddRow(1, pVal.Row)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    If pVal.ItemUID = "12" And (pVal.ColUID = "1" Or pVal.ColUID = "3" Or pVal.ColUID = "5") Or pVal.ItemUID = "10" Then
                        CFL(FormUID, pVal)
                    End If
            End Select
        End If
    End Sub
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        objForm = objAddOn.objApplication.Forms.Item(BusinessObjectInfo.FormUID)
        If BusinessObjectInfo.BeforeAction = True Then
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If validate(BusinessObjectInfo.FormUID) = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        BubbleEvent = True
                    End If
            End Select
        Else
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.ActionSuccess = True Then
                        objMatrix = objForm.Items.Item("12").Specific
                        For i As Integer = 1 To objMatrix.RowCount
                            DisableRows(BusinessObjectInfo.FormUID, i) ' disabling old rows
                        Next i
                    End If
            End Select
        End If
    End Sub
     Public Sub DisableRows(ByVal FormUID As String, ByVal RowID As Integer)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("12").Specific
        For i = 1 To objMatrix.Columns.Count - 1
            objMatrix.CommonSetting.SetCellEditable(RowID, i, False)
        Next i
    End Sub
    
    Private Sub AddNewRow(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("12").Specific
        If objMatrix.RowCount > 1 Then
            If objMatrix.Columns.Item("1").Cells.Item(objMatrix.RowCount - 1).Specific.String <> "" Then
                PROLine.Clear()
                objMatrix.AddRow()
            End If
        ElseIf objMatrix.RowCount = 1 Then
            If objMatrix.Columns.Item("1").Cells.Item(1).Specific.String <> "" Then
                PROLine.Clear()
                objMatrix.AddRow()
            End If
        ElseIf objMatrix.RowCount = 0 Then
            objMatrix.AddRow()
        End If
    End Sub
    Private Sub addRow(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        PROLine = objForm.DataSources.DBDataSources.Item("@MIPLPROV1")
        objMatrix = objForm.Items.Item("12").Specific
        Dim startdate As Date, enddate As Date
        Dim noofdays As Integer, i As Integer
        objForm.Items.Item("6").Specific.String = objAddOn.objGenFunc.GetDocNum(Formtype)
        For i = 1 To objMatrix.RowCount
            startdate = objAddOn.objGenFunc.GetDateTimeValue(objMatrix.Columns.Item("8").Cells.Item(i).Specific.String)
            enddate = objAddOn.objGenFunc.GetDateTimeValue(objMatrix.Columns.Item("9").Cells.Item(i).Specific.String)
            noofdays = CInt(DateDiff(DateInterval.Day, startdate, enddate))
            objMatrix.Columns.Item("10").Cells.Item(i).Specific.String = noofdays
        Next i
    End Sub
    Private Function validate(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("12").Specific
        If objMatrix.RowCount = 0 Then
            objAddOn.objApplication.SetStatusBarMessage("Atleast one row should be added")
            Return False
        ElseIf objForm.Items.Item("4B").Specific.String = "" Or objMatrix.Columns.Item("1").Cells.Item(1).Specific.ToString = "" Or objMatrix.Columns.Item("2").Cells.Item(1).Specific.ToString = "" Or objMatrix.Columns.Item("7").Cells.Item(1).Specific.ToString = "" Or objMatrix.Columns.Item("8").Cells.Item(1).Specific.ToString = "" Or objMatrix.Columns.Item("9").Cells.Item(1).Specific.ToString = "" Then
            objAddOn.objApplication.SetStatusBarMessage("Fields Should Not Be Empty")
            Return False
        End If
        Return True
    End Function
    Private Sub LoadLedger(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("12").Specific
        strSQL = "select Code,Name,U_CreditGL,U_CreditGL1,U_DebitGL,U_DebitGL1 from [@MIPLEX]  where Code='" & objForm.Items.Item("4").Specific.selected.value & "'"
        objRecordSet.DoQuery(strSQL)
        objForm.Items.Item("10").Specific.String = objRecordSet.Fields.Item("U_CreditGL").Value
        objForm.Items.Item("11").Specific.String = objRecordSet.Fields.Item("U_CreditGL1").Value
        objMatrix.Clear()
        While Not objRecordSet.EoF
            objMatrix.AddRow()
            PROLine.Clear()
            objMatrix.GetLineData(objMatrix.RowCount)
            'JELine.SetValue("U_Desc", 0, objRecordSet.Fields.Item("U_Desc").Value)
            PROLine.SetValue("U_CreditGL", 0, objRecordSet.Fields.Item("U_CreditGL").Value)
            PROLine.SetValue("U_CreditGL1", 0, objRecordSet.Fields.Item("U_CreditGL1").Value)
            PROLine.SetValue("U_DebitGL", 0, objRecordSet.Fields.Item("U_DebitGL").Value)
            PROLine.SetValue("U_DebitGL1", 0, objRecordSet.Fields.Item("U_DebitGL1").Value)
            'objAddOn.WriteSMSLog(CStr(objRecordSet.Fields.Item("FromDate").Value))          
            objMatrix.SetLineData(objMatrix.RowCount)
            objRecordSet.MoveNext()
        End While
    End Sub
    Private Sub RemoveEmptyRows(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("12").Specific
        For i As Integer = objMatrix.RowCount To 1 Step -1
            If objMatrix.Columns.Item("1").Cells.Item(i).Specific.String <> "" Then
                Exit For
            Else
                objMatrix.DeleteRow(i)
            End If
        Next
    End Sub

    Private Sub CFL(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent)
        Try
            Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
            Dim objDataTable As SAPbouiCOM.DataTable
            objCFLEvent = pval
            objDataTable = objCFLEvent.SelectedObjects
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("12").Specific
            Select Case objCFLEvent.ChooseFromListUID
                Case "CFL_1"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objForm.Items.Item("10").Specific.String = objDataTable.GetValue("AcctCode", 0)
                            objForm.Items.Item("11").Specific.String = objDataTable.GetValue("AcctName", 0)
                        End If
                    Catch ex As Exception
                        objForm.Items.Item("10").Specific.String = objDataTable.GetValue("AcctCode", 0)
                        objForm.Items.Item("11").Specific.String = objDataTable.GetValue("AcctName", 0)
                    End Try
                Case "CFL_2"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objMatrix.Columns.Item("3").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("AcctCode", 0)
                            objMatrix.Columns.Item("4").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("AcctName", 0)
                        End If
                    Catch ex As Exception
                        objMatrix.Columns.Item("3").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("AcctCode", 0)
                        objMatrix.Columns.Item("4").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("AcctName", 0)
                    End Try
                Case "CFL_3"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objMatrix.Columns.Item("5").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("AcctCode", 0)
                            objMatrix.Columns.Item("6").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("AcctName", 0)
                        End If
                    Catch ex As Exception
                        objMatrix.Columns.Item("5").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("AcctCode", 0)
                        objMatrix.Columns.Item("6").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("AcctName", 0)
                    End Try
                Case "CFL_4"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objMatrix.Columns.Item("1").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("empID", 0)
                            objMatrix.Columns.Item("2").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("firstName", 0)
                        End If
                    Catch ex As Exception
                        objMatrix.Columns.Item("1").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("empID", 0)
                        objMatrix.Columns.Item("2").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("firstName", 0)
                    End Try
            End Select

        Catch ex As Exception
            'MsgBox("Error : " + ex.Message + vbCrLf + "Position : " + ex.StackTrace, MsgBoxStyle.Critical)
        End Try

    End Sub
End Class
