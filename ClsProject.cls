VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsProject
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 03 Oct 22
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete - Deletes record
'===============================================================
' Index - ProjectNo
'---------------------------------------------------------------
Option Explicit

Private pProjectNo As Integer
Private pProjectName As String
Private pCaseManager As ClsCBSUser
Private pLoanTerm As String
Private pProjectWorkflow As ClsWorkflow
Private pCBSCommission As Single
Private pCBSCommPC As Single
Private pExitFee As Single
Private pDebt As Single
Private pExitFeePC As Single
Private pWorkflows As ClsWorkflows
Private pContacts As ClsContacts
Private pParent As ClsSPV
Private pProjectWFNo As Integer
Private pClient As ClsClient
Private pSPV As ClsSPV
Private pFirstClientInt As Integer
Private pSecondClientRef As String
Private pDistribution As String
Private pCompleteDate As Date
Private pStartDate As Date
Private pConsComenceDte As Date
Private pMS As ClsContact
Private pValuer As ClsContact
Private WithEvents pProjectForm As FrmProjectForm
Attribute pProjectForm.VB_VarHelpID = -1
Private WithEvents pProjContacts As FrmProjConts
Attribute pProjContacts.VB_VarHelpID = -1

'---------------------------------------------------------------
Public Property Get ProjectNo() As Integer
    ProjectNo = pProjectNo
End Property

Public Property Let ProjectNo(ByVal vNewValue As Integer)
    pProjectNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ProjectName() As String
    ProjectName = pProjectName
End Property

Public Property Let ProjectName(ByVal ProjectName As String)
    pProjectName = ProjectName
End Property

'---------------------------------------------------------------
Public Property Get CaseManager() As ClsCBSUser
    Set CaseManager = pCaseManager
End Property

Public Property Let CaseManager(ByVal vNewValue As ClsCBSUser)
    Set pCaseManager = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get LoanTerm() As String
    LoanTerm = pLoanTerm
End Property

Public Property Let LoanTerm(ByVal vNewValue As String)
    pLoanTerm = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ProjectWorkflow() As ClsWorkflow
    If pProjectWorkflow Is Nothing Then
    Set pProjectWorkflow = New ClsWorkflow
    pProjectWorkflow.Initialize Me
    End If
    
    If pProjectWFNo > 0 Then pProjectWorkflow.DBGet CStr(pProjectWFNo)
    Set ProjectWorkflow = pProjectWorkflow
End Property

Public Property Let ProjectWorkflow(ByRef vNewValue As ClsWorkflow)
    vNewValue.Initialize Me
    pProjectWorkflow.Name = "Project"
    Set pProjectWorkflow = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get CBSCommission() As Single
    CBSCommission = pCBSCommission
End Property

Public Property Let CBSCommission(ByVal vNewValue As Single)
    pCBSCommission = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get CBSCommPC() As Single
    CBSCommPC = pCBSCommPC
End Property

Public Property Let CBSCommPC(ByVal vNewValue As Single)
    pCBSCommPC = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ExitFee() As Single
    ExitFee = pExitFee
End Property

Public Property Let ExitFee(ByVal vNewValue As Single)
    pExitFee = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Debt() As Single
    Debt = pDebt
End Property

Public Property Let Debt(ByVal vNewValue As Single)
    pDebt = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ExitFeePC() As Single
    ExitFeePC = pExitFeePC
End Property

Public Property Let ExitFeePC(ByVal vNewValue As Single)
    pExitFeePC = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Workflows() As ClsWorkflows
    Set Workflows = pWorkflows
End Property

Public Property Let Workflows(ByRef vNewValue As ClsWorkflows)
    Set pWorkflows = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Contacts() As ClsContacts
    Set Contacts = pContacts
End Property

Public Property Let Contacts(ByRef vNewValue As ClsContacts)
    Set pContacts = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Client() As ClsClient
    Set Client = pClient
End Property

Public Property Let Client(ByRef Client As ClsClient)
    Set pClient = Client
End Property

'---------------------------------------------------------------
Public Property Get SPV() As ClsSPV
    Set SPV = pSPV
End Property

Public Property Let SPV(ByRef SPV As ClsSPV)
    Set pSPV = SPV
End Property
    
'---------------------------------------------------------------
Public Property Get FirstClientInt() As Integer
    FirstClientInt = pFirstClientInt
End Property

Public Property Let FirstClientInt(ByRef FirstClientInt As Integer)
    pFirstClientInt = FirstClientInt
End Property
    
'---------------------------------------------------------------
Public Property Get SecondClientRef() As String
    SecondClientRef = pSecondClientRef
End Property

Public Property Let SecondClientRef(ByRef SecondClientRef As String)
    pSecondClientRef = SecondClientRef
End Property
    
'---------------------------------------------------------------
Public Property Get Distribution() As String
    Distribution = pDistribution
End Property

Public Property Let Distribution(ByRef Distribution As String)
    pDistribution = Distribution
End Property
    
'---------------------------------------------------------------
Public Property Get CompleteDate() As Date
    CompleteDate = pCompleteDate
End Property

Public Property Let CompleteDate(ByRef CompleteDate As Date)
    pCompleteDate = CompleteDate
End Property
    
'---------------------------------------------------------------
Public Property Get StartDate() As Date
    StartDate = pStartDate
End Property

Public Property Let StartDate(ByRef StartDate As Date)
    pStartDate = StartDate
End Property
    
'---------------------------------------------------------------
Public Property Get ConsComenceDte() As Date
    ConsComenceDte = pConsComenceDte
End Property

Public Property Let ConsComenceDte(ByRef ConsComenceDte As Date)
    pConsComenceDte = ConsComenceDte
End Property
    
'---------------------------------------------------------------
Public Property Get MS() As ClsContact
    Set MS = pMS
End Property

Public Property Let MS(ByRef MS As ClsContact)
    Set pMS = MS
End Property
    
'---------------------------------------------------------------
Public Property Get Valuer() As ClsContact
    Set Valuer = pValuer
End Property

Public Property Let Valuer(ByRef Valuer As ClsContact)
    Set pValuer = Valuer
End Property
    
'---------------------------------------------------------------
Public Property Get Parent() As ClsSPV
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsSPV)
    Set pParent = obj
End Property
'---------------------------------------------------------------

' ===============================================================
' Method DBNew
' displays new SPV form
' ---------------------------------------------------------------
Public Sub DBNew()
    
    ProjectWorkflow.Name = "Project"
    pStartDate = Now
    
    With pProjectForm
        .Tag = "New"
        .ClearForm
        .BtnUpdate.Caption = "Create"
        .TxtConsComenceDte = Format(Now, " dd mmm yy")
        .TxtCBSCommission = 0
        .TxtCBSCommPC = 0
        .TxtDebt = 0
        .TxtExitFee = 0
        .TxtExitFeePC = 0
        DBSave
        
        With pProjectWorkflow
            .Name = "Project"
            .DBSave
            .WorkflowType = enProject
            .ActiveStep.Start
            .DBSave
        End With
        
        .Show
    End With
End Sub

' ===============================================================
' Method DisplayContactFrm
' displays project contact form
' ---------------------------------------------------------------
Public Sub DisplayContactFrm()
    Dim Contact As ClsContact
    Dim i As Integer
    
    With pProjContacts
        .LstContacts.Clear
        For Each Contact In pContacts
            With .LstContacts
                .AddItem
                .List(i, 0) = Contact.ContactNo
                .List(i, 1) = Contact.ContactName
                .List(i, 2) = Contact.Organisation
                i = i + 1
            End With
        Next
        If .Visible = False Then .Show
    End With
End Sub

' ===============================================================
' Method DisplayForm
' displays client form
' ---------------------------------------------------------------
Public Sub DisplayForm()
    With pProjectForm
        .TxtCBSCommission = pCBSCommission
        .TxtCBSCommPC = pCBSCommPC
        .TxtLoanTerm = pLoanTerm
        .TxtProjectNo = pProjectNo
        .TxtProjName = pProjectName
        If pFirstClientInt > 0 Then .CmoFirstClientInt = pFirstClientInt
        If pSecondClientRef > "" Then .TxtSecondClientRef = pSecondClientRef
        If pDistribution > "" Then .CmoDistribution = pDistribution
        If pCaseManager.CBSUserNo > 0 Then .CmoCaseManager = pCaseManager.CBSUserNo
        If pClient.ClientNo > 0 Then .CmoClientNo = pClient.ClientNo
        If pSPV.SPVNo > 0 Then .CmoSPVNo = pSPV.SPVNo
        .TxtExitFee = pExitFee
        .TxtDebt = pDebt
        .TxtExitFeePC = pExitFeePC
        If pConsComenceDte > 0 Then .TxtConsComenceDte = pConsComenceDte Else .TxtConsComenceDte = ""
        
        .TxtMSName = pMS.ContactName
        .TxtMSPhone = pMS.Phone1
        .TxtMSEmail = pMS.EmailAddress
        
        .TxtValName = pValuer.ContactName
        .TxtValPhone = pValuer.Phone1
        .TxtValEmail = pValuer.EmailAddress
        .Show
    End With
End Sub

' ===============================================================
' Method DBGet(Key as Integer)
' Gets class from Database
' ---------------------------------------------------------------
Public Sub DBGet(ByVal ProjectNameNo As Variant)
    Dim RstProject As Recordset

    If IsNumeric(ProjectNameNo) Then
        Set RstProject = ModDatabase.SQLQuery("SELECT * FROM TblProject WHERE ProjectNo = " & ProjectNameNo)
    Else
        Set RstProject = ModDatabase.SQLQuery("SELECT * FROM TblProject WHERE ProjectName = '" & ProjectNameNo & "'")
    End If
    
    With RstProject
        If .RecordCount > 0 Then
            If Not IsNull(!ProjectNo) Then pProjectNo = !ProjectNo
            If Not IsNull(!LoanTerm) Then pLoanTerm = !LoanTerm
            If Not IsNull(!CBSCommission) Then pCBSCommission = !CBSCommission
            If Not IsNull(!CBSCommPC) Then pCBSCommPC = !CBSCommPC
            If Not IsNull(!ExitFee) Then pExitFee = !ExitFee
            If Not IsNull(!Debt) Then pDebt = !Debt
            If Not IsNull(!ExitFeePC) Then pExitFeePC = !ExitFeePC
            If Not IsNull(!ProjectNo) Then pProjectNo = !ProjectNo
            If Not IsNull(!ProjectWFNo) Then pProjectWFNo = !ProjectWFNo
            If Not IsNull(!ProjectName) Then pProjectName = !ProjectName
            If Not IsNull(!FirstClientInt) Then pFirstClientInt = !FirstClientInt
            If Not IsNull(!SecondClientRef) Then pSecondClientRef = !SecondClientRef
            If Not IsNull(!Distribution) Then pDistribution = !Distribution
            If Not IsNull(!CompleteDate) Then pCompleteDate = !CompleteDate
            If Not IsNull(!StartDate) Then pStartDate = !StartDate
            If Not IsNull(!ConsComenceDte) Then pConsComenceDte = !ConsComenceDte
            
            pSPV.DBGet !SPVNo
            pMS.DBGet !MS
            pValuer.DBGet !Valuer
            pClient.DBGet !ClientNo
        pCaseManager.DBGet !CaseManager
            pContacts.GetCollection "Project", pProjectNo
            ProjectWorkflow.DBGet !ProjectWFNo
        End If
    End With
    Set RstProject = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
' ---------------------------------------------------------------
Public Sub DBSave()
    Dim RstProject As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstProject = ModDatabase.SQLQuery("SELECT * FROM TblProject WHERE ProjectNo = " & pProjectNo)
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(ProjectNo) FROM TblProject ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstProject
        If .RecordCount = 0 Then
            .AddNew
            pProjectNo = LastNo + 1
        Else
            .Edit
        End If

        pSPV.DBSave
        pClient.DBSave
        pCaseManager.DBSave
        pMS.DBSave
        pValuer.DBSave
        
        pWorkflows.SaveCollection
        pContacts.SaveCollection
        pProjectWorkflow.DBSave
        
        !ProjectNo = CleanSQLText(pProjectNo, True)
        !LoanTerm = CleanSQLText(pLoanTerm, True)
        !CBSCommission = CleanSQLText(pCBSCommission, True)
        !CBSCommPC = CleanSQLText(pCBSCommPC, True)
        !ExitFee = CleanSQLText(pExitFee, True)
        !Debt = CleanSQLText(pDebt, True)
        !ExitFeePC = CleanSQLText(pExitFeePC, True)
        !ProjectNo = CleanSQLText(pProjectNo, True)
        !FirstClientInt = CleanSQLText(pFirstClientInt, True)
        !SecondClientRef = CleanSQLText(pSecondClientRef, True)
        !Distribution = CleanSQLText(pDistribution, True)
        !CompleteDate = CleanSQLText(pCompleteDate, True)
        !StartDate = CleanSQLText(pStartDate, True)
        !ProjectName = CleanSQLText(pProjectName, True)
        !ProjectWFNo = CleanSQLText(pProjectWorkflow.WorkflowNo, True)
        !CaseManager = CleanSQLText(pCaseManager.CBSUserNo, True)
        !MS = CleanSQLText(pMS.ContactNo, True)
        !Valuer = CleanSQLText(pValuer.ContactNo, True)
        !SPVNo = CleanSQLText(pSPV.SPVNo, True)
        !ClientNo = CleanSQLText(pClient.ClientNo, True)
        !ConsComenceDte = CleanSQLText(pConsComenceDte, True)

        pSPV.ClientNo = pClient.ClientNo

        .Update
    End With

    Set RstProject = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete
' Deletes Record
' ---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)

    DB.Execute "DELETE * FROM TblProject WHERE ProjectNo = " & pProjectNo

End Sub

' ===============================================================
' Event pProjContacts_Add
' Event from contact form to add new Contact
' ---------------------------------------------------------------
Private Sub pProjContacts_Add()
    Dim Contact As ClsContact
    
    Set Contact = New ClsContact
    With Contact
        .ContactName = pProjContacts.TxtContactName
        .ContactType = "Project"
        .ContactIndex = pProjectNo
        .Organisation = pProjContacts.CmoOrganisation
        .Phone1 = pProjContacts.TxtPhone1
        .Phone2 = pProjContacts.TxtPhone2
        .EmailAddress = pProjContacts.TxtEmailAddress
        .Position = pProjContacts.TxtPosition
        .Notes = pProjContacts.xTxtNotes
        .DBSave
    End With
    pContacts.Add Contact
    
    DisplayContactFrm
    
End Sub

' ===============================================================
' Event pProjContacts_Delete
' Event from contact form to delete Contact
' ---------------------------------------------------------------
Private Sub pProjContacts_Delete(ContactNo As Integer)
    
    pContacts.Remove ContactNo
    DB.Execute "DELETE * FROM TblContact WHERE ContactNo = " & ContactNo
    DisplayContactFrm
End Sub

' ===============================================================
' Event pProjContacts_Display
' Event from contact form to display Contact
' ---------------------------------------------------------------
Private Sub pProjContacts_Display(ContactNo As Integer)
    Dim Contact As ClsContact
    
    Set Contact = New ClsContact
    Set Contact = pContacts(CStr(ContactNo))
    
    If Contact Is Nothing Then Exit Sub
    
    With pProjContacts
        .EnableForm True
        .TxtContactName = Contact.ContactName
        .TxtEmailAddress = Contact.EmailAddress
        .CmoOrganisation = Contact.Organisation
        .TxtPhone1 = Contact.Phone1
        .TxtPhone2 = Contact.Phone2
        .TxtPosition = Contact.Position
        .xTxtNotes = Contact.Notes
    End With
    
    Set Contact = Nothing
End Sub

' ===============================================================
' Event pProjContacts_Edit
' Event from contact form to edit Contact
' ---------------------------------------------------------------
Private Sub pProjContacts_Edit(ContactNo As Integer)
    Dim Contact As ClsContact
    
    Set Contact = New ClsContact
    Set Contact = pContacts(CStr(ContactNo))
    
    With pProjContacts
        Contact.ContactName = .TxtContactName
        Contact.EmailAddress = .TxtEmailAddress
        Contact.Organisation = .CmoOrganisation
        Contact.Phone1 = .TxtPhone1
        Contact.Phone2 = .TxtPhone2
        Contact.Position = .TxtPosition
    End With
    
    Contact.DBSave
    
    DisplayContactFrm
    
    Set Contact = Nothing

End Sub

' ===============================================================
' Event pProjectForm_DisplayContactsFrm
' Event from project form to display Contact
' ---------------------------------------------------------------
Private Sub pProjectForm_DisplayContactsFrm()
    DisplayContactFrm
End Sub

' ===============================================================
' Event pProjectForm_OpenClientForm
' Event from project form to open client form
' ---------------------------------------------------------------
Private Sub pProjectForm_OpenClientForm()
    pClient.DisplayForm
End Sub

' ===============================================================
' Event pProjectForm_OpenSPVForm
' Event from project form to open SPV form
' ---------------------------------------------------------------
Private Sub pProjectForm_OpenSPVForm()
    pSPV.DisplayForm
End Sub

' ===============================================================
' Event pContactForm_Update
' Event from client form to update Contact
' ---------------------------------------------------------------
Private Sub pProjectForm_Update()
    Set pSPV = New ClsSPV
    Set pClient = New ClsClient
    Set pCaseManager = New ClsCBSUser
    
    With pProjectForm
       pCBSCommission = .TxtCBSCommission
        If .TxtCBSCommPC <> "" Then pCBSCommPC = .TxtCBSCommPC
       pLoanTerm = .TxtLoanTerm
       pClient.DBGet .CmoClientNo.List(.CmoClientNo.ListIndex, 0)
        
        With .CmoFirstClientInt
            If .ListIndex <> -1 Then pFirstClientInt = .List(.ListIndex, 0)
        End With
        
       pSecondClientRef = .TxtSecondClientRef
        pDistribution = .CmoDistribution
       pSPV.DBGet .CmoSPVNo.List(.CmoSPVNo - 1, 0)
       pCaseManager.DBGet .CmoCaseManager.List(.CmoCaseManager - 1, 0)
        If .TxtExitFee <> "" Then pExitFee = .TxtExitFee
       pConsComenceDte = .TxtConsComenceDte
        If .TxtDebt <> "" Then pDebt = .TxtDebt
        If .TxtExitFeePC <> "" Then pExitFeePC = .TxtExitFeePC
       pProjectName = .TxtProjName
        
        If .TxtMSName <> "" Then
            pMS.ContactType = "MS"
            pMS.ContactName = .TxtMSName
            pMS.Phone1 = .TxtMSPhone
            pMS.EmailAddress = .TxtMSEmail
        End If
        
        If .TxtValName <> "" Then
            pValuer.ContactType = "Valuer"
            pValuer.ContactName = .TxtValName
            pValuer.Phone1 = .TxtValPhone
            pValuer.EmailAddress = .TxtValEmail
        End If
    End With
    DBSave
End Sub

' ===============================================================
' Event pContactForm_CreateNew
' Event from SPV form to create new Contact
' ---------------------------------------------------------------
Private Sub pProjectForm_CreateNew()
    With pProjectForm
        .Hide
    End With
    DBNew
End Sub

' ===============================================================
' Event pContactForm_Delete
' Deletes Contact
' ---------------------------------------------------------------
Private Sub pProjectForm_Delete()
    DBDelete
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsSPV)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

Private Sub Class_Initialize()
    Set pWorkflows = New ClsWorkflows
    Set pContacts = New ClsContacts
    Set pCaseManager = New ClsCBSUser
    Set pProjectForm = New FrmProjectForm
    Set pSPV = New ClsSPV
    Set pClient = New ClsClient
    Set pMS = New ClsContact
    Set pValuer = New ClsContact
    Set pProjContacts = New FrmProjConts
   
    Set pParent = New ClsSPV

    pWorkflows.Initialize Me
End Sub
' ===============================================================
Public Sub Terminate()
    pWorkflows.Terminate


    Set Me.Parent = Nothing

    Set pWorkflows = Nothing
    Set pContacts = Nothing
    Set pCaseManager = Nothing
    Set pProjectForm = Nothing
    Set pSPV = Nothing
    Set pClient = Nothing
    Set pMS = Nothing
    Set pValuer = Nothing
    Set pProjContacts = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pWorkflows = Nothing
    Set pProjectWorkflow = Nothing
    Set pContacts = Nothing
    Set pCaseManager = Nothing
    Set pProjectForm = Nothing
    Set pMS = Nothing
    Set pValuer = Nothing
    Set pProjContacts = Nothing
End Sub
' ===============================================================
