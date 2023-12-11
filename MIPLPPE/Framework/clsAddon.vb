Imports System.IO

Public Class clsAddOn
    Public WithEvents objApplication As SAPbouiCOM.Application
    Public objCompany As SAPbobsCOM.Company
    Dim oProgBarx As SAPbouiCOM.ProgressBar
    Public objGenFunc As Mukesh.SBOLib.GeneralFunctions
    Public objUIXml As Mukesh.SBOLib.UIXML
    Public ZB_row As Integer = 0
    Public SOMenuID As String = "0"
   
    Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
    Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
    Dim ret As Long
    Dim str As String
    Dim objForm As SAPbouiCOM.Form
    Dim MenuCount As Integer = 0
    ' Public objPPE As clsPPE
    Public objPRE As clsPRE
    Public objJEPost As clsJEPost
    Public objJEPosting As clsJEPosting
    Public objPRO As clsPRO
    Public objJEPostingPro As clsJEPostingPro
    Public HANA As Boolean = False
    ' Public HANA As Boolean = True
    Public HWKEY() As String = New String() {"H0922924113", "Q0198611247", "T0264302252", "V0913316776", "F0123559701", "L1552968038", "M0090876837", "Y1334940735", "A0061802481"}
    Private Sub CheckLicense()

    End Sub
    Function isValidLicense() As Boolean
        Try
            objApplication.Menus.Item("257").Activate()
            Dim CrrHWKEY As String = objApplication.Forms.ActiveForm.Items.Item("79").Specific.Value.ToString.Trim
            objApplication.Forms.ActiveForm.Close()

            For i As Integer = 0 To HWKEY.Length - 1
                If HWKEY(i).Trim = CrrHWKEY.Trim Then
                    Return True
                End If
            Next
            MsgBox("Installing Add-On failed due to License mismatch", MsgBoxStyle.OkOnly, "License Management")
            Return False
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return True
    End Function
    Public Sub Intialize()
        Dim objSBOConnector As New Mukesh.SBOLib.SBOConnector
        objApplication = objSBOConnector.GetApplication(System.Environment.GetCommandLineArgs.GetValue(1))
        objCompany = objSBOConnector.GetCompany(objApplication)
        Try
            createObjects()
            loadMenu()
            createTables()
            createUDOs()
            addJobCardReporttype()
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.ToString)
            End
        End Try
        If isValidLicense() Then
            objApplication.SetStatusBarMessage("Addon connected successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Else
            objApplication.SetStatusBarMessage("Failed To Connect, Please Check The License Configuration", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            objCompany.Disconnect()
            objApplication = Nothing
            objCompany = Nothing
            End
        End If
    End Sub
    Private Sub createUDOs()
        Dim objUDFEngine As New Mukesh.SBOLib.UDFEngine(objCompany)
        Dim ct1(1) As String
        ct1(0) = ""
        objUDFEngine.createUDO("@MIPLEX", "MIPLEX", "Expenses", ct1, SAPbobsCOM.BoUDOObjType.boud_MasterData)
        'ct1(0) = "@MIPLPPE1"
        'objUDFEngine.createUDO("@MIPLOPPE", "MIPLPPE", "Prepaid", ct1, SAPbobsCOM.BoUDOObjType.boud_Document, False, True)
        ct1(0) = "MIPLPREP1"
        objUDFEngine.createUDO("MIPLPREP", "MIPLPREP", "Prepaid", ct1, SAPbobsCOM.BoUDOObjType.boud_Document, False, False)
        ct1(0) = "MIPLPROV1"
        objUDFEngine.createUDO("MIPLPROV", "MIPLPROV", "Provision", ct1, SAPbobsCOM.BoUDOObjType.boud_Document, False, False)
        ct1(0) = "MIPLJEP1"
        objUDFEngine.createUDO("MIPLJEP", "MIPLJEP", "JE Posting", ct1, SAPbobsCOM.BoUDOObjType.boud_Document, False, False)
        ct1(0) = "MIPLJEPO1"
        objUDFEngine.createUDO("MIPLJEPO", "MIPLJEPO", "JE Posting Pro", ct1, SAPbobsCOM.BoUDOObjType.boud_Document, False, False)

        objAddOn.objApplication.SetStatusBarMessage("UDO Created", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    End Sub
    Private Sub createObjects()
        objGenFunc = New Mukesh.SBOLib.GeneralFunctions(objCompany)
        objUIXml = New Mukesh.SBOLib.UIXML(objApplication)
        'objPPE = New clsPPE
        objPRE = New clsPRE
        objJEPost = New clsJEPost
        objPRO = New clsPRO
        objJEPosting = New clsJEPosting
        objJEPostingPro = New clsJEPostingPro
    End Sub
    Private Sub objApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ItemEvent
        Try
            Select Case pVal.FormTypeEx
                Case clsPRE.Formtype
                    objPRE.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsJEPosting.Formtype
                    objJEPosting.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsPRO.Formtype
                    objPRO.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clsJEPostingPro.Formtype
                    objJEPostingPro.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Case "139" 'sales order
                    '    objSO.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Case "149" 'sales Quotation
                    '    objSQ.ItemEvent(FormUID, pVal, BubbleEvent)
            End Select
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
    End Sub
    Private Sub objApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles objApplication.FormDataEvent
        Try
            Select Case BusinessObjectInfo.FormTypeEx
                Case clsPRE.Formtype
                    objPRE.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case clsPRO.Formtype
                    objPRO.FormDataEvent(BusinessObjectInfo, BubbleEvent)

            End Select
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
    End Sub
    Public Shared Function ErrorHandler(ByVal p_ex As Exception, ByVal objApplication As SAPbouiCOM.Application)
        Dim sMsg As String = Nothing
        If p_ex.Message = "Form - already exists [66000-11]" Then
            Return True
            Exit Function  'ignore error
        End If
        Return False
    End Function
    Private Sub objApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objApplication.MenuEvent
        If pVal.BeforeAction Then
            Select Case pVal.MenuUID

            End Select
        Else
            Try
                Select Case pVal.MenuUID
                    Case clsPRE.Formtype
                        objPRE.LoadScreen()
                    Case clsJEPosting.Formtype
                        objJEPosting.LoadScreen()
                    Case clsPRO.Formtype
                        objPRO.LoadScreen()
                    Case clsJEPostingPro.Formtype
                        objJEPostingPro.LoadScreen()
                    Case "1282", "1290", "1289", "1291", "1281"
                        If objApplication.Forms.ActiveForm.UniqueID.Contains(clsPRE.Formtype) Or objApplication.Forms.ActiveForm.UniqueID.Contains(clsPRO.Formtype) Then
                            '        objJobCard.LoadSeries(objApplication.Forms.ActiveForm.UniqueID)
                            '        objJobCard.LoadSalesEmp(objApplication.Forms.ActiveForm.UniqueID)
                            '   objPRE.DisableRows(objApplication.Forms.ActiveForm.UniqueID)
                            'objPRO.DisableRows(objApplication.Forms.ActiveForm.UniqueID)
                        End If
                End Select
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If
    End Sub
    Private Sub loadMenu()
        If objApplication.Menus.Item("43520").SubMenus.Exists("MIPLPPE") Then Return
        MenuCount = objApplication.Menus.Item("43520").SubMenus.Count

        CreateMenu("", MenuCount + 1, "Prepaid & Provision", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLPPE", objApplication.Menus.Item("43520"))

        CreateMenu("", 1, "Prepaid", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLPRE", objApplication.Menus.Item("MIPLPPE"))
        CreateMenu("", 1, "Prepaid", SAPbouiCOM.BoMenuType.mt_STRING, clsPRE.Formtype, objApplication.Menus.Item("MIPLPRE"))
        CreateMenu("", 2, "JE Posting", SAPbouiCOM.BoMenuType.mt_STRING, clsJEPosting.Formtype, objApplication.Menus.Item("MIPLPRE"))

        CreateMenu("", 2, "Provision", SAPbouiCOM.BoMenuType.mt_POPUP, "MIPLPROV", objApplication.Menus.Item("MIPLPPE"))
        CreateMenu("", 1, "Provision", SAPbouiCOM.BoMenuType.mt_STRING, clsPRO.Formtype, objApplication.Menus.Item("MIPLPROV"))
        CreateMenu("", 2, "JE Posting", SAPbouiCOM.BoMenuType.mt_STRING, clsJEPostingPro.Formtype, objApplication.Menus.Item("MIPLPROV"))

        objApplication.SetStatusBarMessage("Menu Created!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    End Sub
    Private Function CreateMenu(ByVal ImagePath As String, ByVal Position As Int32, ByVal DisplayName As String, ByVal MenuType As SAPbouiCOM.BoMenuType, ByVal UniqueID As String, ByVal ParentMenu As SAPbouiCOM.MenuItem) As SAPbouiCOM.MenuItem
        Try
            Dim oMenuPackage As SAPbouiCOM.MenuCreationParams
            oMenuPackage = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oMenuPackage.Image = ImagePath
            oMenuPackage.Position = Position
            oMenuPackage.Type = MenuType
            oMenuPackage.UniqueID = UniqueID
            oMenuPackage.String = DisplayName
            ParentMenu.SubMenus.AddEx(oMenuPackage)
        Catch ex As Exception
            objApplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
        Return ParentMenu.SubMenus.Item(UniqueID)
    End Function
    Private Sub createTables()
        Dim objUDFEngine As New Mukesh.SBOLib.UDFEngine(objCompany)
        objAddOn.objApplication.SetStatusBarMessage("Creating Tables Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, False)
        ' WriteSMSLog("0")
        objUDFEngine.CreateTable("MIPLEX", "Expense Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("@MIPLEX", "Desc", "Description", 100)
        objUDFEngine.AddAlphaField("@MIPLEX", "PrePro", "Prepaid/Provision", 100, "PE,PRO", "Prepaid,Provision", "PE")
        objUDFEngine.AddAlphaField("@MIPLEX", "AppType", "Applicable Type", 40)
        objUDFEngine.AddAlphaField("@MIPLEX", "CreditGL", "CreditGL Code", 20)
        objUDFEngine.AddAlphaField("@MIPLEX", "CreditGL1", "CreditGL Name", 30)
        objUDFEngine.AddAlphaField("@MIPLEX", "DebitGL", "DebitGL", 20)
        objUDFEngine.AddAlphaField("@MIPLEX", "DebitGL1", "DebitGL Name", 30)

        objUDFEngine.CreateTable("MIPLPREP", "Prepaid Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objUDFEngine.AddDateField("@MIPLPREP", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("@MIPLPREP", "ExpType", "Expense Type", 20)
        objUDFEngine.AddAlphaField("@MIPLPREP", "AppType", "Applicable Type", 40)
        objUDFEngine.AddAlphaField("@MIPLPREP", "DocNum", "Doc Number", 50)
        objUDFEngine.AddAlphaField("@MIPLPREP", "CreditGL", "Credit GL", 20)
        objUDFEngine.AddAlphaField("@MIPLPREP", "CreditGLN", "Credit GLName", 100)

        objUDFEngine.CreateTable("MIPLPREP1", "PRE Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("@MIPLPREP1", "Desc", "Description", 50)
        objUDFEngine.AddDateField("@MIPLPREP1", "FromDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddDateField("@MIPLPREP1", "ToDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("@MIPLPREP1", "CreditGL", "Credit GL", 20)
        objUDFEngine.AddAlphaField("@MIPLPREP1", "CreditGL1", "Credit GLName", 100)
        objUDFEngine.AddAlphaField("@MIPLPREP1", "DebitGL", "Debit GL", 20)
        objUDFEngine.AddAlphaField("@MIPLPREP1", "DebitGL1", "Debit GLName", 100)
        objUDFEngine.AddNumericField("@MIPLPREP1", "NoofDays", "No of Days", 10)
        objUDFEngine.AddFloatField("@MIPLPREP1", "RentAmt", "Rent Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objUDFEngine.AddNumericField("@MIPLPREP1", "PendDays", "Pending Days", 10)
        objUDFEngine.AddFloatField("@MIPLPREP1", "PendAmt", "Pending Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objUDFEngine.AddDateField("@MIPLPREP1", "LstExDte", "last Expense Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objUDFEngine.CreateTable("MIPLPROV", "Provision Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objUDFEngine.AddDateField("@MIPLPROV", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("@MIPLPROV", "ExpType", "Expense Type", 20)
        objUDFEngine.AddAlphaField("@MIPLPROV", "AppType", "Applicable Type", 40)
        objUDFEngine.AddAlphaField("@MIPLPROV", "DocNum", "Doc Number", 50)
        objUDFEngine.AddAlphaField("@MIPLPROV", "CreditGL", "Credit GL", 20)
        objUDFEngine.AddAlphaField("@MIPLPROV", "CreditGL1", "Credit GL Name", 100)

        objUDFEngine.CreateTable("MIPLPROV1", "PRO Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("@MIPLPROV1", "EMPCode", "Employee Code", 30)
        objUDFEngine.AddAlphaField("@MIPLPROV1", "EMPName", "Employee Name", 100)
        objUDFEngine.AddAlphaField("@MIPLPROV1", "Desc", "Description", 50)
        objUDFEngine.AddAlphaField("@MIPLPROV1", "CreditGL", "Credit GL", 20)
        objUDFEngine.AddAlphaField("@MIPLPROV1", "CreditGL1", "Credit GLName", 100)
        objUDFEngine.AddAlphaField("@MIPLPROV1", "DebitGL", "Debit GL", 20)
        objUDFEngine.AddAlphaField("@MIPLPROV1", "DebitGL1", "Debit GLName", 100)
        objUDFEngine.AddNumericField("@MIPLPROV1", "NoofDays", "No of Days", 10)
        objUDFEngine.AddDateField("@MIPLPROV1", "BalDate", "Balance Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddDateField("@MIPLPROV1", "LastPaidTill", "LastPaidTill", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddFloatField("@MIPLPROV1", "Postout", "Post out", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objUDFEngine.AddFloatField("@MIPLPROV1", "Basic", "Basic", SAPbobsCOM.BoFldSubTypes.st_Sum)

        objUDFEngine.CreateTable("MIPLJEP", "JEPosting Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objUDFEngine.AddDateField("@MIPLJEP", "ExpDate", "Last Expense Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("@MIPLJEP", "ExpType", "Expense Type", 20)
        objUDFEngine.AddAlphaField("@MIPLJEP", "Status", "Status", 5, "O,C", "Open,Closed", "O")
        objUDFEngine.AddFloatField("@MIPLJEP", "TotAmt", "Total Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objUDFEngine.AddDateField("@MIPLJEP", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objUDFEngine.CreateTable("MIPLJEP1", "JEPosting Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("@MIPLJEP1", "Select", "Select", 10, "Y,N", "Yes,No", "N")
        objUDFEngine.AddAlphaField("@MIPLJEP1", "Desc", "Description", 50)
        objUDFEngine.AddAlphaField("@MIPLJEP1", "CreditGL", "Credit GL", 20)
        objUDFEngine.AddAlphaField("@MIPLJEP1", "CreditGL1", "Credit GLName", 100)
        objUDFEngine.AddAlphaField("@MIPLJEP1", "DebitGL", "Debit GL", 20)
        objUDFEngine.AddAlphaField("@MIPLJEP1", "DebitGL1", "Debit GLName", 100)
        objUDFEngine.AddAlphaField("@MIPLJEP1", "BaseEntry", "Base Entry", 20)
        objUDFEngine.AddAlphaField("@MIPLJEP1", "BaseLineNum", "Base Line Num", 20)
        objUDFEngine.AddFloatField("@MIPLJEP1", "RentAmt", "Rent Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objUDFEngine.AddDateField("@MIPLJEP1", "FROMDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddDateField("@MIPLJEP1", "ToDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddNumericField("@MIPLJEP1", "NoofDays", "No of Days", 10)
        objUDFEngine.AddFloatField("@MIPLJEP1", "ExpAmnt", "Last Expense Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objUDFEngine.AddDateField("@MIPLJEP1", "FrmDate", "From Date1", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddDateField("@MIPLJEP1", "LstExDte", "last Expense Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddFloatField("@MIPLJEP1", "OutstdRnt", "Outstanding Rent Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objUDFEngine.AddNumericField("@MIPLJEP1", "Blncdays", "Balance Days", 10)
        objUDFEngine.AddNumericField("@MIPLJEP1", "ExpDays", "Expense Days", 10)
        objUDFEngine.AddNumericField("@MIPLJEP1", "PrpdDays", "Prepaid Days", 10)
        objUDFEngine.AddFloatField("@MIPLJEP1", "PrpdRent", "Prepaid Rent Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objUDFEngine.AddFloatField("@MIPLJEP1", "PostAmt", "Post out Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objUDFEngine.AddAlphaField("@MIPLJEP1", "Status", "Status", 10, "O,C", "Open,Closed", "O")
        objUDFEngine.AddAlphaField("@MIPLJEP1", "JEEntry", "JE Entry", 30)

        objUDFEngine.CreateTable("MIPLJEPO", "JEPosting ProHeader", SAPbobsCOM.BoUTBTableType.bott_Document)
        objUDFEngine.AddDateField("@MIPLJEPO", "ExpDate", "Last Expense Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("@MIPLJEPO", "ExpType", "Expense Type", 20)
        objUDFEngine.AddDateField("@MIPLJEPO", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddAlphaField("@MIPLJEPO", "DocNum", "Doc Number", 50)
        objUDFEngine.AddFloatField("@MIPLJEPO", "TotAmt", "Total Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)

        objUDFEngine.CreateTable("MIPLJEPO1", "JEPosting ProLines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUDFEngine.AddAlphaField("@MIPLJEPO1", "Select", "Select", 10, "Y,N", "Yes,No", "N")
        objUDFEngine.AddAlphaField("@MIPLJEPO1", "CreditGL", "Credit GL", 20)
        objUDFEngine.AddAlphaField("@MIPLJEPO1", "CreditGL1", "Credit GLName", 100)
        objUDFEngine.AddAlphaField("@MIPLJEPO1", "DebitGL", "Debit GL", 20)
        objUDFEngine.AddAlphaField("@MIPLJEPO1", "DebitGL1", "Debit GLName", 100)
        objUDFEngine.AddAlphaField("@MIPLJEPO1", "BaseEntry", "Base Entry", 20)
        objUDFEngine.AddAlphaField("@MIPLJEPO1", "BaseLineNum", "Base Line Num", 20)
        objUDFEngine.AddAlphaField("@MIPLJEPO1", "EMPCode", "Employee Code", 30)
        objUDFEngine.AddAlphaField("@MIPLJEPO1", "EMPName", "Employee Name", 100)
        objUDFEngine.AddFloatField("@MIPLJEPO1", "Basic", "Basic", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objUDFEngine.AddDateField("@MIPLJEPO1", "LastPaidTill", "LastPaidTill", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddDateField("@MIPLJEPO1", "FromDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddDateField("@MIPLJEPO1", "BalDate", "Balance Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUDFEngine.AddNumericField("@MIPLJEPO1", "NoofDays", "No of Days", 10)
        objUDFEngine.AddFloatField("@MIPLJEPO1", "Postout", "Postout Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objUDFEngine.AddAlphaField("@MIPLJEPO1", "Status", "Status", 10, "O,C", "Open,Closed", "O")
        objUDFEngine.AddAlphaField("@MIPLJEPO1", "JEEntry", "JE Entry", 30)
        objUDFEngine.AddDateField("@MIPLJEPO1", "LstExDte", "last Expense Date", SAPbobsCOM.BoFldSubTypes.st_None)

        'objUDFEngine.CreateTable("MIPLJE", "JEPosting Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        'objUDFEngine.AddDateField("@MIPLJE", "ToDate", "ToDate", SAPbobsCOM.BoFldSubTypes.st_None)
        'objUDFEngine.AddAlphaField("@MIPLJE", "ExType", "Expense Type", 20)
        'objUDFEngine.AddAlphaField("@MIPLJE", "DocStatus", "Doc Status", 10)

        'objUDFEngine.CreateTable("MIPLJE1", "JEPosting Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        'objUDFEngine.AddDateField("@MIPLJE1", "PostingDate", "PostingDate", SAPbobsCOM.BoFldSubTypes.st_None)
        'objUDFEngine.AddAlphaField("@MIPLJE1", "Select", "Select", 10, "Y,N", "Yes,No", "N")
        'objUDFEngine.AddAlphaField("@MIPLJE1", "Type", "CreditGL Code", 10)
        'objUDFEngine.AddAlphaField("@MIPLJE1", "CreditGL", "CreditGL Code", 20)
        'objUDFEngine.AddAlphaField("@MIPLJE1", "CreditGL1", "CreditGL Name", 30)
        'objUDFEngine.AddAlphaField("@MIPLJE1", "DebitGL", "DebitGL", 20)
        'objUDFEngine.AddAlphaField("@MIPLJE1", "DebitGL1", "DebitGL Name", 30)
        'objUDFEngine.AddFloatField("@MIPLJE1", "Amount", "Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        'objUDFEngine.AddAlphaField("@MIPLJE1", "Status", "Status", 10)
        'objUDFEngine.AddAlphaField("@MIPLJE1", "BaseRef", "BaseRef", 10)
        'objUDFEngine.AddNumericField("@MIPLJE1", "BaseLinNum", "BaseLinNum", 10)



        '*******************  Table ******************* START********************************* END
    End Sub
    Private Sub objApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objApplication.RightClickEvent
        If eventInfo.BeforeAction Then
        Else
            If eventInfo.FormUID.Contains("QC") And (eventInfo.ItemUID = "20") And eventInfo.Row > 0 Then

                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try

                    If objAddOn.objApplication.Menus.Exists("ditem") Then
                        objAddOn.objApplication.Menus.RemoveEx("ditem")
                    End If
                Catch ex As Exception

                End Try
                Try

                    oMenuItem = objAddOn.objApplication.Menus.Item("1280").SubMenus.Item("ditem")
                    ZB_row = eventInfo.Row
                Catch ex As Exception
                    Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                    oCreationPackage = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                    oCreationPackage.UniqueID = "ditem"
                    oCreationPackage.String = "Delete Row"
                    oCreationPackage.Enabled = True

                    oMenuItem = objAddOn.objApplication.Menus.Item("1280") 'Data'
                    oMenus = oMenuItem.SubMenus
                    oMenus.AddEx(oCreationPackage)
                    ZB_row = eventInfo.Row
                End Try
                If eventInfo.ItemUID <> "45" Then
                    '   Dim oMenuItem As SAPbouiCOM.MenuItem
                    '  Dim oMenus As SAPbouiCOM.Menus
                    Try
                        objAddOn.objApplication.Menus.RemoveEx("ditem")
                    Catch ex As Exception
                        ' MessageBox.Show(ex.Message)
                    End Try
                End If
            End If
            End If
    End Sub

    Private Sub objApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles objApplication.AppEvent
        If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
            Try
                ' objUIXml.LoadMenuXML("RemoveMenu.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded)
                If objCompany.Connected Then objCompany.Disconnect()
                objCompany = Nothing
                objApplication = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objCompany)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objApplication)
                GC.Collect()
            Catch ex As Exception
            End Try
            End
        End If
    End Sub

    Private Sub applyFilter()
        Dim oFilters As SAPbouiCOM.EventFilters
        Dim oFilter As SAPbouiCOM.EventFilter
        oFilters = New SAPbouiCOM.EventFilters
        'Item Master Data 
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)




        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK)


    End Sub
    Public Sub WriteSMSLog(ByVal Str As String)
        Dim fs As FileStream
        Dim chatlog As String = Application.StartupPath & "\Log_" & Today.ToString("yyyyMMdd") & ".txt"
        If File.Exists(chatlog) Then
        Else
            fs = New FileStream(chatlog, FileMode.Create, FileAccess.Write)
            fs.Close()
        End If
        ' Dim objReader As New System.IO.StreamReader(chatlog)
        Dim sdate As String
        sdate = Now
        'objReader.Close()
        If System.IO.File.Exists(chatlog) = True Then
            Dim objWriter As New System.IO.StreamWriter(chatlog, True)
            objWriter.WriteLine(sdate & " : " & Str)
            objWriter.Close()
        Else
            Dim objWriter As New System.IO.StreamWriter(chatlog, False)
            ' MsgBox("Failed to send message!")
        End If
    End Sub
    Private Sub addJobCardReporttype()
        'Dim rptTypeService As SAPbobsCOM.ReportTypesService
        'Dim newType As SAPbobsCOM.ReportType
        'Dim newtypeParam As SAPbobsCOM.ReportTypeParams
        'Dim newReportParam As SAPbobsCOM.ReportLayoutParams
        'Dim ReportExists As Boolean = False
        'Try


        '    Dim newtypesParam As SAPbobsCOM.ReportTypesParams
        '    rptTypeService = objAddOn.objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
        '    newtypesParam = rptTypeService.GetReportTypeList

        '    Dim i As Integer
        '    For i = 0 To newtypesParam.Count - 1
        '        If newtypesParam.Item(i).TypeName = clsJobCard.FormType And newtypesParam.Item(i).MenuID = clsJobCard.FormType Then
        '            ReportExists = True
        '            Exit For
        '        End If
        '    Next i

        '    If Not ReportExists Then
        '        rptTypeService = objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
        '        newType = rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType)


        '        newType.TypeName = clsJobCard.FormType
        '        newType.AddonName = "JC2Addon"
        '        newType.AddonFormType = clsJobCard.FormType
        '        newType.MenuID = clsJobCard.FormType
        '        newtypeParam = rptTypeService.AddReportType(newType)

        '        Dim rptService As SAPbobsCOM.ReportLayoutsService
        '        Dim newReport As SAPbobsCOM.ReportLayout
        '        rptService = objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService)
        '        newReport = rptService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout)
        '        newReport.Author = objCompany.UserName
        '        newReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal
        '        newReport.Name = clsJobCard.FormType
        '        newReport.TypeCode = newtypeParam.TypeCode

        '        newReportParam = rptService.AddReportLayout(newReport)

        '        newType = rptTypeService.GetReportType(newtypeParam)
        '        newType.DefaultReportLayout = newReportParam.LayoutCode
        '        rptTypeService.UpdateReportType(newType)

        '        Dim oBlobParams As SAPbobsCOM.BlobParams
        '        oBlobParams = objCompany.GetCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams)
        '        oBlobParams.Table = "RDOC"
        '        oBlobParams.Field = "Template"
        '        Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment
        '        oKeySegment = oBlobParams.BlobTableKeySegments.Add
        '        oKeySegment.Name = "DocCode"
        '        oKeySegment.Value = newReportParam.LayoutCode

        '        Dim oFile As FileStream
        '        oFile = New FileStream(Application.StartupPath + "\JobCard.rpt", FileMode.Open)
        '        Dim fileSize As Integer
        '        fileSize = oFile.Length
        '        Dim buf(fileSize) As Byte
        '        oFile.Read(buf, 0, fileSize)
        '        oFile.Dispose()

        '        Dim oBlob As SAPbobsCOM.Blob
        '        oBlob = objCompany.GetCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob)
        '        oBlob.Content = Convert.ToBase64String(buf, 0, fileSize)
        '        objCompany.GetCompanyService.SetBlob(oBlobParams, oBlob)
        '    End If
        'Catch ex As Exception
        '    objApplication.MessageBox(ex.ToString)
        'End Try

    End Sub

    Private Sub objApplication_LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean) Handles objApplication.LayoutKeyEvent

        ''BubbleEvent = True
        'If eventInfo.BeforeAction = True Then
        '    If eventInfo.FormUID.Contains(clsJobCard.FormType) Then
        '        objJobCard.LayoutKeyEvent(eventInfo, BubbleEvent)
        '    End If
        'End If
    End Sub
End Class


