Public Class clsPPE
    Public Const Formtype = "MIPLPREP"
    Dim objForm As SAPbouiCOM.Form
    Dim strSQL As String
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim PPEHeader As SAPbouiCOM.DBDataSource
    Dim PPELine As SAPbouiCOM.DBDataSource
    Dim objComboBox As SAPbouiCOM.ComboBox

    Dim objRecordSet As SAPbobsCOM.Recordset
    Dim DebitGL As String
    Dim DebitGL1 As String
    Public Sub LoadScreen()
        objForm = objAddOn.objUIXml.LoadScreenXML("PPEInput.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, Formtype)
        'objForm.Items.Item("21").Specific.validvalues.loadseries(Formtype, SAPbouiCOM.BoSeriesMode.sf_Add)
        'objForm.Items.Item("21").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        ' PPEHeader.SetValue("DocNum", 0, objAddOn.objGenFunc.GetDocNum(Formtype))
    
        objForm.AutoManaged = True
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "select Code,Name from [@MIPLEX]"
        objRecordSet.DoQuery(strSQL)
        While Not objRecordSet.EoF

            objCombobox = objForm.Items.Item("10").Specific
            objCombobox.ValidValues.Add(objRecordSet.Fields.Item("Code").Value, objRecordSet.Fields.Item("Name").Value)
            objRecordSet.MoveNext()
        End While
        objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        objForm.Items.Item("8").Specific.Active = True
        objForm.Items.Item("8").Specific.String = "A"
        PPELine = objForm.DataSources.DBDataSources.Item("@MIPLPPE1")
        PPEHeader = objForm.DataSources.DBDataSources.Item("@MIPLOPPE")
    End Sub
    Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        If pVal.BeforeAction = True Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If pVal.ItemUID = "1" And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                If Not Validate(FormUID) Then
                                    BubbleEvent = False
                                End If
                                RemoveEmptyRows(FormUID)
                            End If

                    End Select
                  
            End Select
        Else
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "29" Then
                        LoadMatrix(FormUID)
                    ElseIf pVal.ItemUID = "29A" Then
                        addItems(FormUID, "30")
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "13" Then

                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "10" Then
                        TypeSelection(FormUID)
                    End If
                    'If pVal.ItemUID = "21" Then
                    '    objForm.Items.Item("4").Specific.String = objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("21").Specific.selected.value, Formtype)
                    'End If
            End Select
        End If

    End Sub
    Private Function validate(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("10").Specific
        If objMatrix.RowCount = 0 Then
            objAddOn.objApplication.SetStatusBarMessage("Atleast one row should be added")
            Return False
        ElseIf objMatrix.Columns.Item("1").Cells.Item(1).Specific.ToString = "" Then
            objAddOn.objApplication.SetStatusBarMessage("Empty Row Found")
            Return False
        End If
        Return True
    End Function
    Private Sub TypeSelection(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("30").Specific
        objComboBox = objForm.Items.Item("10").Specific

        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "Select * from [@MIPLEX] where Code='" & objComboBox.Selected.Value & "'"
        objRecordSet.DoQuery(strSQL)
        If Not objRecordSet.EoF Then
            objForm.Items.Item("12").Specific.String = CStr(objRecordSet.Fields.Item("U_Desc").Value)
            objForm.Items.Item("18").Specific.String = CStr(objRecordSet.Fields.Item("U_CreditGL").Value)
            objForm.Items.Item("18A").Specific.String = CStr(objRecordSet.Fields.Item("U_CreditGL1").Value)
            DebitGL = CStr(objRecordSet.Fields.Item("U_DebitGL").Value)
            DebitGL1 = CStr(objRecordSet.Fields.Item("U_DebitGL1").Value)
        End If
        objForm.Freeze(True)
        Select Case objComboBox.Selected.Value
            Case "Visa"
                For intloop As Integer = 1 To objMatrix.Columns.Count - 1
                    objMatrix.Columns.Item(intloop).Visible = False
                Next
                objMatrix.Columns.Item("1").Visible = True
                Dim secondcol As String = "2"
                objMatrix.Columns.Item(secondcol).Visible = True
                objMatrix.Columns.Item("3").Visible = True
                objMatrix.Columns.Item("4").Visible = True
                objMatrix.Columns.Item("14").Visible = True
                objMatrix.Columns.Item("14A").Visible = True
                objMatrix.Columns.Item("11").Visible = True
                objMatrix.Columns.Item("15").Visible = True

                objForm.Items.Item("29").Enabled = True '-----------Load
                objForm.Items.Item("29A").Enabled = False '------------Add Row
            Case "Rent"
                For intloop As Integer = 1 To objMatrix.Columns.Count - 1
                    objMatrix.Columns.Item(intloop).Visible = False

                Next
                objMatrix.Columns.Item("1").Visible = True
                objMatrix.Columns.Item("3").Visible = True
                objMatrix.Columns.Item("4").Visible = True
                objMatrix.Columns.Item("14").Visible = True
                objMatrix.Columns.Item("14A").Visible = True
                objMatrix.Columns.Item("11").Visible = True
                objMatrix.Columns.Item("15").Visible = True

                objForm.Items.Item("29").Enabled = True
                objForm.Items.Item("29A").Enabled = False
            Case Else
                For intloop As Integer = 1 To objMatrix.Columns.Count - 1
                    objMatrix.Columns.Item(intloop).Visible = True
                Next

                objForm.Items.Item("29").Enabled = False
                objForm.Items.Item("29A").Enabled = True
        End Select
        objForm.Freeze(False)

    End Sub
    Private Sub LoadMatrix(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("30").Specific
        PPELine.Clear()
        objMatrix.Clear()
        Dim totalamt = IIf(objForm.Items.Item("28").Specific.String = "", 0, CDbl(objForm.Items.Item("28").Specific.String))
        Dim NoofPeriods = IIf(objForm.Items.Item("26").Specific.String = "", 0, CInt(objForm.Items.Item("26").Specific.String))
        Dim divider As Integer = 0
        Select Case objForm.Items.Item("24").Specific.selected.value
            Case "M"
                divider = 12
            Case "Y"
                divider = 1
            Case "Q"
                divider = 4
        End Select
        '----------------------
        Dim startdate As Date
        startdate = objAddOn.objGenFunc.GetDateTimeValue(objForm.Items.Item("14").Specific.String)
        Dim noofdays As Integer = Date.DaysInMonth(startdate.Year, startdate.Month)
        startdate = startdate.AddDays(noofdays - startdate.Day)
        '----------------------
        objMatrix.Clear()
        If totalamt <> 0 And NoofPeriods <> 0 Then
            For intloop As Integer = 1 To NoofPeriods
                If objMatrix.RowCount = 0 Then
                    objMatrix.AddRow()
                ElseIf objMatrix.Columns.Item("1").Cells.Item(objMatrix.RowCount).Specific.String <> "" Then
                    objMatrix.AddRow()
                End If
                objMatrix.GetLineData(objMatrix.RowCount)
                PPELine.SetValue("U_Desc", 0, objForm.Items.Item("10").Specific.selected.value)
                PPELine.SetValue("U_Details", 0, objForm.Items.Item("12").Specific.string)
                PPELine.SetValue("U_DocNum", 0, "")
                PPELine.SetValue("U_DocDate", 0, startdate.ToString("yyyyMMdd"))

                'objAddOn.objApplication.MessageBox(CStr(startdate.Month) & "-------" & CStr(noofdays))
                If startdate.Month + 1 > 12 Then
                    noofdays = Date.DaysInMonth(startdate.Year + 1, 1)
                Else
                    noofdays = Date.DaysInMonth(startdate.Year, startdate.Month + 1)
                End If

                startdate = startdate.AddDays(noofdays)

                'PPELine.SetValue("U_FromDate", 0, "")
                'PPELine.SetValue("U_ToDate", 0, "")
                'PPELine.SetValue("U_ClDate", 0, "")
                PPELine.SetValue("U_DebitGL", 0, DebitGL)
                PPELine.SetValue("U_DebitGL1", 0, DebitGL1)
                'PPELine.SetValue("U_NoofDays", 0, "")
                'PPELine.SetValue("U_TotDays", 0, "")
                'PPELine.SetValue("U_PpDays", 0, "")
                PPELine.SetValue("U_TotAmt", 0, totalamt / NoofPeriods)
                'PPELine.SetValue("U_PpAmt", 0, "")
                'PPELine.SetValue("U_ExpAmt", 0, "")
                PPELine.SetValue("U_Status", 0, "O")
                objMatrix.SetLineData(objMatrix.RowCount)

                PPELine.Clear()
                objForm.Update()
            Next
        End If
        objForm.Refresh()
    End Sub
    Private Sub addItems(ByVal FormUID As String, ByVal MatrixID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item(MatrixID).Specific
        Dim AddRow As Boolean
        If objMatrix.RowCount <= 0 Then
            AddRow = True
        ElseIf objMatrix.Columns.Item("1").Cells.Item(objMatrix.RowCount).Specific.string <> "" Then
            AddRow = True
        End If
        If AddRow Then
            PPELine.Clear()
            objMatrix.AddRow()
            objMatrix.GetLineData(objMatrix.RowCount)
            PPELine.SetValue("U_Desc", 0, objForm.Items.Item("10").Specific.selected.value)
            PPELine.SetValue("U_Details", 0, objForm.Items.Item("12").Specific.string)
            PPELine.SetValue("U_DocNum", 0, "")
            PPELine.SetValue("U_DebitGL", 0, DebitGL)
            PPELine.SetValue("U_DebitGL1", 0, DebitGL1)
            PPELine.SetValue("U_Status", 0, "O")
            objMatrix.SetLineData(objMatrix.RowCount)
            PPELine.Clear()
            objForm.Update()

        End If

    End Sub
    Private Sub RemoveEmptyRows(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("30").Specific
        For i As Integer = objMatrix.RowCount To 1 Step -1
            If objMatrix.Columns.Item("1").Cells.Item(i).Specific.String <> "" Then
                Exit For
            Else
                objMatrix.DeleteRow(i)
            End If
        Next
    End Sub
End Class
