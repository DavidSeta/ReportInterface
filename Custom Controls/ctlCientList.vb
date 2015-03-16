'======================================================================================================
'Class ctlCientList
'frmMain main panel for clients
'======================================================================================================
Public Class ctlCientList
    Public CS As String
    Public ReadClientsFromXML As Boolean = False

    Public Clients() As Clients = Nothing

    Private ReadOnly Property Client(ByVal Index As Integer) As Clients
        Get
            Dim bValid As Boolean = True
            If Index < Clients.GetLowerBound(0) Then bValid = False
            If Index > Clients.GetUpperBound(0) Then bValid = False
            If bValid Then Return Clients(Index)

            Return New Clients
        End Get
    End Property
    '=====================================================================================================
    'ctlCientList_Load() - Build ListView (LV) in ctlCientList & cmbFilter (Filter Client by:)
    '=====================================================================================================
    Private Sub ctlCientList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetFilters()
    End Sub
    '=====================================================================================================
    'cmbFilter_SelectedIndexChanged() - Apply selected filter from cmbFilter (Filter Client by:)
    '=====================================================================================================
    Private Sub cmbFilter_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFilter.SelectedIndexChanged
        If Me.cmbFilter.SelectedIndex >= 0 Then
            NewList(Me.cmbFilter.SelectedItem.value)
        End If
    End Sub

#Region " NewList Sub "
    '=====================================================================================================
    'SetNew() - Define filters for panel frmMain.ctlCientList invoked by frmMain: Me.ClientList.SetNew()
    '=====================================================================================================
    Public Sub SetNew()
        If Me.cmbFilter.Items.Count >= 0 Then
            Me.cmbFilter.SelectedIndex = 0
            NewList(Me.cmbFilter.SelectedItem.value)
        End If
    End Sub
    '=====================================================================================================
    'NewList(ByVal Filter As FilterList) - Clear and Build panel
    '=====================================================================================================
    Public Sub NewList(ByVal Filter As FilterList)
        Me.LV.Clear()
        'Get Connection information and fill frmMain.ctlCientList.LV
        Dim db As New clDatabase
        db.cs = Me.CS
        Dim oClients() As Clients = db.GetClientList

        'Setup Clients
        Select Case Filter
            Case FilterList.All_Clients : Me.Clients_All(oClients)
            Case FilterList.Google_Clients : Me.Clients_Google(oClients)
            Case FilterList.Google_5_Clients : Me.Clients_Google5(oClients)
            Case FilterList.MSN_Clients : Me.Clients_MSN(oClients)
            Case FilterList.Verizon_Clients : Me.Clients_Verizon(oClients)
            Case FilterList.Yahoo_Clients : Me.Clients_Yahoo(oClients)
            Case FilterList.ASK_Clients : Me.Clients_ASK(oClients)
            Case FilterList.GA_Clients : Me.Clients_GA(oClients)
            Case FilterList.Arb_Clients : Me.Clients_Arb(oClients)
            Case Else : Me.Clients_All(oClients)
        End Select

        UpdateList()
    End Sub
    '=====================================================================================================
    'UpdateList() - Update panel list
    '=====================================================================================================
    Public Sub UpdateList()

        'Setup Listview Box
        Me.LV.View = View.Details
        Me.LV.MultiSelect = True
        Me.LV.GridLines = True
        Me.LV.FullRowSelect = True

        'Add Headers
        Me.LV.Columns.Add("Ref", 29, HorizontalAlignment.Left)
        Me.LV.Columns.Add("Client", 145, HorizontalAlignment.Left)
        Me.LV.Columns.Add("Google Report", 95, HorizontalAlignment.Left)
        Me.LV.Columns.Add("MSN Report", 95, HorizontalAlignment.Left)
        Me.LV.Columns.Add("Yahoo Report", 95, HorizontalAlignment.Left)
        Me.LV.Columns.Add("Verizon Report", 95, HorizontalAlignment.Left)
        Me.LV.Columns.Add("ASK Report", 95, HorizontalAlignment.Left)
        Me.LV.Columns.Add("GA Report", 95, HorizontalAlignment.Left)
        Me.LV.Columns.Add("Arbitrage", 25, HorizontalAlignment.Left)

        'Fill Listview with Clients
        Dim nCount As Integer
        For nCount = 0 To Clients.GetUpperBound(0)
            Dim lvRef As New ListViewItem((nCount + 1).ToString("00"))

            Dim lvClient As New ListViewItem.ListViewSubItem(lvRef, Clients(nCount).Name)
            Dim lvGoogle As New ListViewItem.ListViewSubItem(lvRef, IIf(Clients(nCount).Google.Has_AccountID, "Google Waiting...", "None"))
            Dim lvMSN As New ListViewItem.ListViewSubItem(lvRef, IIf(Clients(nCount).MSN.Has_AccountID, "MSN Waiting...", "None"))
            Dim lvYahoo As New ListViewItem.ListViewSubItem(lvRef, IIf(Clients(nCount).Yahoo.Has_AccountID, "Yahoo Waiting...", "None"))
            Dim lvVerizon As New ListViewItem.ListViewSubItem(lvRef, IIf(Clients(nCount).Verizon.Has_AccountID, "Verizon Waiting...", "None"))
            Dim lvASK As New ListViewItem.ListViewSubItem(lvRef, IIf(Clients(nCount).ASK.Has_AccountID, "ASK Waiting...", "None"))
            Dim lvGA As New ListViewItem.ListViewSubItem(lvRef, IIf(Clients(nCount).GA.Has_AccountID, "GA Waiting...", "None"))
            Dim lvArbitrage As New ListViewItem.ListViewSubItem(lvRef, IIf(Clients(nCount).IsArbitrage = 1, "Y", "N"))

            lvRef.SubItems.Add(lvClient)
            lvRef.SubItems.Add(lvGoogle)
            lvRef.SubItems.Add(lvMSN)
            lvRef.SubItems.Add(lvYahoo)
            lvRef.SubItems.Add(lvVerizon)
            lvRef.SubItems.Add(lvASK)
            lvRef.SubItems.Add(lvGA)
            lvRef.SubItems.Add(lvArbitrage)

            Me.LV.Items.Add(lvRef)
            Clients(nCount).LV_ID = Me.LV.Items.Count - 1
        Next
    End Sub
#End Region 'NewList Sub

#Region " Create ClientList Subs "
    Private Sub Clients_All(ByVal oClients As Clients())
        Clients = Nothing

        Dim Index As Integer = 0
        If Me.ReadClientsFromXML = False Then
            For Each client As Clients In oClients
                If client.HasAnyAccountID Then
                    If IsNothing(Clients) Then
                        Index = 0
                        ReDim Clients(0)
                    Else
                        Index = Clients.GetUpperBound(0) + 1
                    End If
                    ReDim Preserve Clients(Index)
                    Clients(Index) = client
                    Clients(Index).Index = Index
                    Clients(Index).LV_ID = 0
                End If
            Next
            Exit Sub
        End If
        Index = 0

        Dim nCount As Integer
        Dim Settings As New clSettings
        For nCount = 1 To Settings.Clients.GetUpperBound(0)
            If IsNothing(Clients) Then
                Index = 0
                ReDim Clients(Index)
            Else
                Index = Clients.GetUpperBound(0) + 1
                ReDim Preserve Clients(Index)
            End If
            Clients(Index) = New Clients
            Clients(Index).Index = Index
            Clients(Index).Name = Settings.Clients(nCount).UserName
            Clients(Index).CustID = Settings.Clients(nCount).CustID
            Clients(Index).Google.AccountID = Settings.Clients(nCount).GoogleID
            Clients(Index).Google.AccountPWD = Settings.Clients(nCount).GooglePassword
            Clients(Index).MSN.AccountID = Settings.Clients(nCount).MSNID
            Clients(Index).Verizon.AccountID = Settings.Clients(nCount).VerizonID
            Clients(Index).Yahoo.AccountID = Settings.Clients(nCount).YahooID
            Clients(Index).Yahoo.SubAccountID = Settings.Clients(nCount).YahooSubID
            Clients(Index).ASK.AccountID = Settings.Clients(nCount).ASKID
        Next
    End Sub

    Private Sub Clients_Google(ByVal oclients As Clients())
        Clients = Nothing

        Dim Index As Integer = 0
        If Me.ReadClientsFromXML = False Then
            For Each client As Clients In oclients
                If client.Google.Has_AccountID Then
                    If IsNothing(Clients) Then
                        Index = 0
                        ReDim Clients(0)
                    Else
                        Index = Clients.GetUpperBound(0) + 1
                    End If
                    ReDim Preserve Clients(Index)
                    Clients(Index) = client
                    Clients(Index).Index = Index
                    Clients(Index).LV_ID = 0
                End If
            Next
            Exit Sub
        End If
        Index = 0

        Dim nCount As Integer
        Dim Settings As New clSettings
        For nCount = 1 To Settings.Clients.GetUpperBound(0)
            If Settings.Clients(nCount).GoogleID <> "" Then
                If IsNothing(Clients) Then
                    Index = 0
                    ReDim Clients(Index)
                Else
                    Index = Clients.GetUpperBound(0) + 1
                    ReDim Preserve Clients(Index)
                End If
                Clients(Index) = New Clients
                Clients(Index).Index = Index
                Clients(Index).Name = Settings.Clients(nCount).UserName
                Clients(Index).CustID = Settings.Clients(nCount).CustID
                Clients(Index).Google.AccountID = Settings.Clients(nCount).GoogleID
                Clients(Index).Google.AccountPWD = Settings.Clients(nCount).GooglePassword
                Clients(Index).MSN.AccountID = Settings.Clients(nCount).MSNID
                Clients(Index).Verizon.AccountID = Settings.Clients(nCount).VerizonID
                Clients(Index).Yahoo.AccountID = Settings.Clients(nCount).YahooID
                Clients(Index).Yahoo.SubAccountID = Settings.Clients(nCount).YahooSubID
                Clients(Index).ASK.AccountID = Settings.Clients(nCount).ASKID
            End If
        Next
    End Sub

    Private Sub Clients_Google5(ByVal oclients As Clients())
        Clients = Nothing

        Dim Index As Integer = 0
        If Me.ReadClientsFromXML = False Then
            For Each client As Clients In oclients
                If client.Google.Has_AccountID Then
                    If client.Google.Has_AcctPwd Then
                        If IsNothing(Clients) Then
                            Index = 0
                            ReDim Clients(0)
                        Else
                            Index = Clients.GetUpperBound(0) + 1
                        End If
                        ReDim Preserve Clients(Index)
                        Clients(Index) = client
                        Clients(Index).Index = Index
                        Clients(Index).LV_ID = 0
                    End If
                End If
            Next
            Exit Sub
        End If
        Index = 0

        Dim nCount As Integer
        Dim Settings As New clSettings
        For nCount = 1 To Settings.Clients.GetUpperBound(0)
            If Settings.Clients(nCount).GoogleID <> "" Then
                If Settings.Clients(nCount).GooglePassword = "search33432" Then
                    If IsNothing(Clients) Then
                        Index = 0
                        ReDim Clients(Index)
                    Else
                        Index = Clients.GetUpperBound(0) + 1
                        ReDim Preserve Clients(Index)
                    End If
                    Clients(Index) = New Clients
                    Clients(Index).Index = Index
                    Clients(Index).Name = Settings.Clients(nCount).UserName
                    Clients(Index).CustID = Settings.Clients(nCount).CustID
                    Clients(Index).Google.AccountID = Settings.Clients(nCount).GoogleID
                    Clients(Index).Google.AccountPWD = Settings.Clients(nCount).GooglePassword
                    Clients(Index).MSN.AccountID = Settings.Clients(nCount).MSNID
                    Clients(Index).Verizon.AccountID = Settings.Clients(nCount).VerizonID
                    Clients(Index).Yahoo.AccountID = Settings.Clients(nCount).YahooID
                    Clients(Index).Yahoo.SubAccountID = Settings.Clients(nCount).YahooSubID
                    Clients(Index).ASK.AccountID = Settings.Clients(nCount).ASKID
                End If
            End If
        Next
    End Sub

    Private Sub Clients_MSN(ByVal oClients As Clients())
        Clients = Nothing

        Dim Index As Integer = 0
        If Me.ReadClientsFromXML = False Then
            For Each client As Clients In oClients
                If client.MSN.Has_AccountID Then
                    If IsNothing(Clients) Then
                        Index = 0
                        ReDim Clients(0)
                    Else
                        Index = Clients.GetUpperBound(0) + 1
                    End If
                    ReDim Preserve Clients(Index)
                    Clients(Index) = client
                    Clients(Index).Index = Index
                    Clients(Index).LV_ID = 0
                End If
            Next
            Exit Sub
        End If
        Index = 0

        Dim nCount As Integer
        Dim Settings As New clSettings
        For nCount = 1 To Settings.Clients.GetUpperBound(0)
            If Settings.Clients(nCount).MSNID <> "" Then
                If IsNothing(Clients) Then
                    Index = 0
                    ReDim Clients(Index)
                Else
                    Index = Clients.GetUpperBound(0) + 1
                    ReDim Preserve Clients(Index)
                End If
                Clients(Index) = New Clients
                Clients(Index).Index = Index
                Clients(Index).Name = Settings.Clients(nCount).UserName
                Clients(Index).CustID = Settings.Clients(nCount).CustID
                Clients(Index).Google.AccountID = Settings.Clients(nCount).GoogleID
                Clients(Index).Google.AccountPWD = Settings.Clients(nCount).GooglePassword
                Clients(Index).MSN.AccountID = Settings.Clients(nCount).MSNID
                Clients(Index).Verizon.AccountID = Settings.Clients(nCount).VerizonID
                Clients(Index).Yahoo.AccountID = Settings.Clients(nCount).YahooID
                Clients(Index).Yahoo.SubAccountID = Settings.Clients(nCount).YahooSubID
                Clients(Index).ASK.AccountID = Settings.Clients(nCount).ASKID
            End If
        Next
    End Sub

    Private Sub Clients_Verizon(ByVal oClients As Clients())
        Clients = Nothing

        Dim Index As Integer = 0
        If Me.ReadClientsFromXML = False Then
            For Each client As Clients In oClients
                If client.Verizon.Has_AccountID Then
                    If IsNothing(Clients) Then
                        Index = 0
                        ReDim Clients(0)
                    Else
                        Index = Clients.GetUpperBound(0) + 1
                    End If
                    ReDim Preserve Clients(Index)
                    Clients(Index) = client
                    Clients(Index).Index = Index
                    Clients(Index).LV_ID = 0
                End If
            Next
            Exit Sub
        End If
        Index = 0

        Dim nCount As Integer
        Dim Settings As New clSettings
        For nCount = 1 To Settings.Clients.GetUpperBound(0)
            If Settings.Clients(nCount).VerizonID <> "" Then
                If IsNothing(Clients) Then
                    Index = 0
                    ReDim Clients(Index)
                Else
                    Index = Clients.GetUpperBound(0) + 1
                    ReDim Preserve Clients(Index)
                End If
                Clients(Index) = New Clients
                Clients(Index).Index = Index
                Clients(Index).Name = Settings.Clients(nCount).UserName
                Clients(Index).CustID = Settings.Clients(nCount).CustID
                Clients(Index).Google.AccountID = Settings.Clients(nCount).GoogleID
                Clients(Index).Google.AccountPWD = Settings.Clients(nCount).GooglePassword
                Clients(Index).MSN.AccountID = Settings.Clients(nCount).MSNID
                Clients(Index).Verizon.AccountID = Settings.Clients(nCount).VerizonID
                Clients(Index).Yahoo.AccountID = Settings.Clients(nCount).YahooID
                Clients(Index).Yahoo.SubAccountID = Settings.Clients(nCount).YahooSubID
                Clients(Index).ASK.AccountID = Settings.Clients(nCount).ASKID
            End If
        Next
    End Sub

    Private Sub Clients_Yahoo(ByVal oClients As Clients())
        Clients = Nothing

        Dim Index As Integer = 0
        If Me.ReadClientsFromXML = False Then
            For Each client As Clients In oClients
                If client.Yahoo.Has_AccountID Or client.Yahoo.Has_SubAccountID Then
                    If IsNothing(Clients) Then
                        Index = 0
                        ReDim Clients(0)
                    Else
                        Index = Clients.GetUpperBound(0) + 1
                    End If
                    ReDim Preserve Clients(Index)
                    Clients(Index) = client
                    Clients(Index).Index = Index
                    Clients(Index).LV_ID = 0
                End If
            Next
            Exit Sub
        End If
        Index = 0

        Dim nCount As Integer
        Dim Settings As New clSettings
        For nCount = 1 To Settings.Clients.GetUpperBound(0)
            If Settings.Clients(nCount).YahooID <> "" Then
                If IsNothing(Clients) Then
                    Index = 0
                    ReDim Clients(Index)
                Else
                    Index = Clients.GetUpperBound(0) + 1
                    ReDim Preserve Clients(Index)
                End If
                Clients(Index) = New Clients
                Clients(Index).Index = Index
                Clients(Index).Name = Settings.Clients(nCount).UserName
                Clients(Index).CustID = Settings.Clients(nCount).CustID
                Clients(Index).Google.AccountID = Settings.Clients(nCount).GoogleID
                Clients(Index).Google.AccountPWD = Settings.Clients(nCount).GooglePassword
                Clients(Index).MSN.AccountID = Settings.Clients(nCount).MSNID
                Clients(Index).Verizon.AccountID = Settings.Clients(nCount).VerizonID
                Clients(Index).Yahoo.AccountID = Settings.Clients(nCount).YahooID
                Clients(Index).Yahoo.SubAccountID = Settings.Clients(nCount).YahooSubID
                Clients(Index).ASK.AccountID = Settings.Clients(nCount).ASKID
            End If
        Next
    End Sub
    Private Sub Clients_ASK(ByVal oClients As Clients())
        Clients = Nothing

        Dim Index As Integer = 0
        If Me.ReadClientsFromXML = False Then
            For Each client As Clients In oClients
                If client.ASK.Has_AccountID Then
                    If IsNothing(Clients) Then
                        Index = 0
                        ReDim Clients(0)
                    Else
                        Index = Clients.GetUpperBound(0) + 1
                    End If
                    ReDim Preserve Clients(Index)
                    Clients(Index) = client
                    Clients(Index).Index = Index
                    Clients(Index).LV_ID = 0
                End If
            Next
            Exit Sub
        End If

        Index = 0
        Dim nCount As Integer
        Dim Settings As New clSettings
        For nCount = 1 To Settings.Clients.GetUpperBound(0)
            If Settings.Clients(nCount).ASKID <> "" Then
                If IsNothing(Clients) Then
                    Index = 0
                    ReDim Clients(Index)
                Else
                    Index = Clients.GetUpperBound(0) + 1
                    ReDim Preserve Clients(Index)
                End If
                Clients(Index) = New Clients
                Clients(Index).Index = Index
                Clients(Index).Name = Settings.Clients(nCount).UserName
                Clients(Index).CustID = Settings.Clients(nCount).CustID
                Clients(Index).Google.AccountID = Settings.Clients(nCount).GoogleID
                Clients(Index).Google.AccountPWD = Settings.Clients(nCount).GooglePassword
                Clients(Index).MSN.AccountID = Settings.Clients(nCount).MSNID
                Clients(Index).Verizon.AccountID = Settings.Clients(nCount).VerizonID
                Clients(Index).Yahoo.AccountID = Settings.Clients(nCount).YahooID
                Clients(Index).Yahoo.SubAccountID = Settings.Clients(nCount).YahooSubID
                Clients(Index).ASK.AccountID = Settings.Clients(nCount).ASKID
            End If
        Next
    End Sub
    Private Sub Clients_GA(ByVal oClients As Clients())
        Clients = Nothing

        Dim Index As Integer = 0
        If Me.ReadClientsFromXML = False Then
            For Each client As Clients In oClients
                If client.GA.Has_AccountID Then
                    If IsNothing(Clients) Then
                        Index = 0
                        ReDim Clients(0)
                    Else
                        Index = Clients.GetUpperBound(0) + 1
                    End If
                    ReDim Preserve Clients(Index)
                    Clients(Index) = client
                    Clients(Index).Index = Index
                    Clients(Index).LV_ID = 0
                End If
            Next
            Exit Sub
        End If

    End Sub
    Private Sub Clients_Arb(ByVal oClients As Clients())
        Clients = Nothing

        Dim Index As Integer = 0
        If Me.ReadClientsFromXML = False Then
            For Each client As Clients In oClients
                If client.IsArbitrage Then
                    If IsNothing(Clients) Then
                        Index = 0
                        ReDim Clients(0)
                    Else
                        Index = Clients.GetUpperBound(0) + 1
                    End If
                    ReDim Preserve Clients(Index)
                    Clients(Index) = client
                    Clients(Index).Index = Index
                    Clients(Index).LV_ID = 0
                End If
            Next
            Exit Sub
        End If

    End Sub
#End Region 'Create ClientList Functions

#Region " SetFilters Sub "
    Private Sub SetFilters()
        Me.cmbFilter.Items.Add(New clMyItem("None (All Clients)", FilterList.All_Clients))
        Me.cmbFilter.Items.Add(New clMyItem("Google Clients", FilterList.Google_Clients))
        Me.cmbFilter.Items.Add(New clMyItem("Google 5 Clients", FilterList.Google_5_Clients))
        Me.cmbFilter.Items.Add(New clMyItem("MSN Clients", FilterList.MSN_Clients))
        Me.cmbFilter.Items.Add(New clMyItem("Verizon Clients", FilterList.Verizon_Clients))
        Me.cmbFilter.Items.Add(New clMyItem("Yahoo Clients", FilterList.Yahoo_Clients))
        Me.cmbFilter.Items.Add(New clMyItem("ASK Clients", FilterList.ASK_Clients))
        Me.cmbFilter.Items.Add(New clMyItem("GA Clients", FilterList.GA_Clients))
        Me.cmbFilter.Items.Add(New clMyItem("Arb Clients", FilterList.Arb_Clients))
    End Sub
#End Region 'SetFilters Sub

#Region " Actions Done Subs "
    'STEP 1
    '=============================================================================================
    'Done_Request() called by frmMain.vb Start_Verizon() ONLY                                     
    'Sets label to "REQUESTED" in grid display                                                    
    '=============================================================================================
    Public Sub Done_Request(ByVal Index As Integer, ByVal ServiceType As WebService)
        Select Case ServiceType
            Case WebService.Google
                Me.LV.Items(Clients(Index).LV_ID).SubItems(2).Text = "REQUESTED"

            Case WebService.MSN
                Me.LV.Items(Clients(Index).LV_ID).SubItems(3).Text = "REQUESTED"

            Case WebService.Yahoo
                Me.LV.Items(Clients(Index).LV_ID).SubItems(4).Text = "REQUESTED"

            Case WebService.Verizon
                Me.LV.Items(Clients(Index).LV_ID).SubItems(5).Text = "REQUESTED"

            Case WebService.ASK
                Me.LV.Items(Clients(Index).LV_ID).SubItems(6).Text = "REQUESTED"

            Case WebService.GA
                Me.LV.Items(Clients(Index).LV_ID).SubItems(7).Text = "REQUESTED"
        End Select
    End Sub
    '=============================================================================================
    'Done_ReportID() called by frmMain.vb Start_Engine()                                          
    'Save returned ReportID in Clients(Index).Engine.ReportID                                     
    'Sets label to "GOT_ID" in grid display                                                       
    '=============================================================================================
    Public Sub Done_ReportID(ByVal Index As Integer, ByVal ServiceType As WebService, ByVal sReportID As String)
        Done_ReportID(Index, ServiceType, sReportID, False)
    End Sub

    Public Sub Done_ReportID(ByVal Index As Integer, ByVal ServiceType As WebService, ByVal sReportID As String, ByVal bYahooCM As Boolean)
        Select Case ServiceType
            Case WebService.Google
                Clients(Index).Google.ReportID = sReportID
                Me.LV.Items(Clients(Index).LV_ID).SubItems(2).Text = "GOT_ID"

            Case WebService.MSN
                Clients(Index).MSN.ReportID = sReportID
                Me.LV.Items(Clients(Index).LV_ID).SubItems(3).Text = "GOT_ID"

            Case WebService.Yahoo
                If bYahooCM = True Then
                    Clients(Index).Yahoo.CM_ReportID = sReportID
                Else
                    Clients(Index).Yahoo.Key_ReportID = sReportID
                End If
                Me.LV.Items(Clients(Index).LV_ID).SubItems(4).Text = "GOT_ID"

            Case WebService.Verizon
                Clients(Index).Verizon.ReportID = sReportID
                Me.LV.Items(Clients(Index).LV_ID).SubItems(5).Text = "GOT_ID"

            Case WebService.ASK
                Clients(Index).ASK.ReportID = sReportID
                Me.LV.Items(Clients(Index).LV_ID).SubItems(6).Text = "GOT_ID"

            Case WebService.GA
                Clients(Index).GA.ReportID = sReportID
                Me.LV.Items(Clients(Index).LV_ID).SubItems(7).Text = "GOT_ID"
        End Select
    End Sub
    'STEP 2
    '=============================================================================================
    'Done_Status() called by frmMain.vb GetStatus_Engine()                                        
    'Set Clients(Index).Engine.Status with status returned from engine: READY, FAILED, INPROGRESS,
    'PENDING                                                                                      
    'Sets label to Status in grid display and when READY via Done_Ready()                         
    'Set Clients(Index).Engine.DownloadReady = True                                               
    '=============================================================================================
    Public Sub Done_Status(ByVal Index As Integer, ByVal ServiceType As WebService, ByVal Status As String)
        Done_Status(Index, ServiceType, Status, False)
    End Sub

    Public Sub Done_Status(ByVal Index As Integer, ByVal ServiceType As WebService, ByVal Status As String, ByVal bYahooCM As Boolean)
        Select Case ServiceType
            Case WebService.Google
                If Status = "READY" Then
                    Me.Done_Ready(Index, ServiceType)
                Else
                    Clients(Index).Google.Status = Status
                    Me.LV.Items(Clients(Index).LV_ID).SubItems(2).Text = Status
                End If

            Case WebService.MSN
                If Status = "READY" Then
                    Me.Done_Ready(Index, ServiceType)
                Else
                    Clients(Index).MSN.Status = Status
                    Me.LV.Items(Clients(Index).LV_ID).SubItems(3).Text = Status
                End If

            Case WebService.Yahoo
                If Status = "READY" Then
                    Me.Done_Ready(Index, ServiceType, bYahooCM)
                Else
                    If bYahooCM = True Then
                        Clients(Index).Yahoo.CM_Status = Status
                    Else
                        Clients(Index).Yahoo.Key_Status = Status
                    End If
                    Me.LV.Items(Clients(Index).LV_ID).SubItems(4).Text = Status
                End If

            Case WebService.Verizon
                If Status = "READY" Then
                    Me.Done_Ready(Index, ServiceType)
                Else
                    Clients(Index).Verizon.Status = Status
                    Me.LV.Items(Clients(Index).LV_ID).SubItems(5).Text = Status
                End If

            Case WebService.ASK
                If Status = "READY" Then
                    Me.Done_Ready(Index, ServiceType)
                Else
                    Clients(Index).ASK.Status = Status
                    Me.LV.Items(Clients(Index).LV_ID).SubItems(6).Text = Status
                End If

            Case WebService.GA
                If Status = "READY" Then
                    Me.Done_Ready(Index, ServiceType)
                Else
                    Clients(Index).GA.Status = Status
                    Me.LV.Items(Clients(Index).LV_ID).SubItems(7).Text = Status
                End If
        End Select
    End Sub
    '=============================================================================================
    'Done_Ready() called by ctlCientList.vb Done_Status()                                         
    'Set Clients(Index).Engine.DownloadReady = True                                               
    'Sets label to "READY" in grid display                                                        
    '=============================================================================================
    Public Sub Done_Ready(ByVal Index As Integer, ByVal ServiceType As WebService)
        Done_Ready(Index, ServiceType, False)
    End Sub

    Public Sub Done_Ready(ByVal Index As Integer, ByVal ServiceType As WebService, ByVal bYahooCM As Boolean)
        Select Case ServiceType
            Case WebService.Google
                Clients(Index).Google.DownloadReady = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(2).Text = "READY"

            Case WebService.MSN
                Clients(Index).MSN.DownloadReady = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(3).Text = "READY"

            Case WebService.Yahoo
                If bYahooCM Then
                    Clients(Index).Yahoo.CM_DownloadReady = True
                Else
                    Clients(Index).Yahoo.Key_DownloadReady = True
                End If
                Me.LV.Items(Clients(Index).LV_ID).SubItems(4).Text = "READY"

            Case WebService.Verizon
                Clients(Index).Verizon.DownloadReady = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(5).Text = "READY"

            Case WebService.ASK
                Clients(Index).ASK.DownloadReady = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(6).Text = "READY"

            Case WebService.GA
                Clients(Index).GA.DownloadReady = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(7).Text = "READY"
        End Select
    End Sub
    'STEP 3
    '=============================================================================================
    'Done_Download() called by frmMain.vb Download_Engine_Once()                                  
    'Set Clients(Index).Engine.Downloaded = True                                                  
    'Sets label to "DOWNLOADED" in grid display                                                   
    '=============================================================================================
    Public Sub Done_Download(ByVal Index As Integer, ByVal ServiceType As WebService)
        Done_Download(Index, ServiceType, False)
    End Sub

    Public Sub Done_Download(ByVal Index As Integer, ByVal ServiceType As WebService, ByVal bYahooCM As Boolean)
        Select Case ServiceType
            Case WebService.Google
                Clients(Index).Google.Downloaded = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(2).Text = "DOWNLOADED"

            Case WebService.MSN
                Clients(Index).MSN.Downloaded = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(3).Text = "DOWNLOADED"

            Case WebService.Yahoo
                If bYahooCM Then
                    Clients(Index).Yahoo.CM_Downloaded = True
                Else
                    Clients(Index).Yahoo.Key_Downloaded = True
                End If
                Me.LV.Items(Clients(Index).LV_ID).SubItems(4).Text = "DOWNLOADED"

            Case WebService.Verizon
                Clients(Index).Verizon.Downloaded = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(5).Text = "DOWNLOADED"

            Case WebService.ASK
                Clients(Index).ASK.Downloaded = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(6).Text = "DOWNLOADED"

            Case WebService.GA
                Clients(Index).GA.Downloaded = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(7).Text = "DOWNLOADED"

        End Select
    End Sub
    'STEP 4
    '=============================================================================================
    'Done_URL() called by frmMain.vb GetStatus_MSN() AND GetURL_Verizon() only                    
    'Set Clients(Index).Engine.DownloadRead = True AND Clients(Index).Engine.URL = sURL           
    'Sets label to "READY" in grid display                                                        
    '=============================================================================================
    Public Sub Done_URL(ByVal Index As Integer, ByVal ServiceType As WebService, ByVal sURL As String)
        Done_URL(Index, ServiceType, sURL, False)
    End Sub

    Public Sub Done_URL(ByVal Index As Integer, ByVal ServiceType As WebService, ByVal sURL As String, ByVal bYahooCM As Boolean)
        Select Case ServiceType
            Case WebService.Google
                Clients(Index).Google.URL = sURL
                Clients(Index).Google.DownloadReady = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(2).Text = "READY"

            Case WebService.MSN
                Clients(Index).MSN.URL = sURL
                Clients(Index).MSN.DownloadReady = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(3).Text = "READY"

            Case WebService.Yahoo
                If bYahooCM Then
                    Clients(Index).Yahoo.CM_URL = sURL
                    Clients(Index).Yahoo.CM_DownloadReady = True
                Else
                    Clients(Index).Yahoo.Key_URL = sURL
                    Clients(Index).Yahoo.Key_DownloadReady = True
                End If
                Me.LV.Items(Clients(Index).LV_ID).SubItems(4).Text = "READY"

            Case WebService.Verizon
                Clients(Index).Verizon.URL = sURL
                Clients(Index).Verizon.DownloadReady = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(5).Text = "READY"

            Case WebService.ASK
                Clients(Index).ASK.URL = sURL
                Clients(Index).ASK.DownloadReady = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(6).Text = "READY"

            Case WebService.GA
                Clients(Index).GA.URL = sURL
                Clients(Index).GA.DownloadReady = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(7).Text = "READY"
        End Select
    End Sub

   

    Public Sub Done_MyProcess(ByVal Index As Integer, ByVal ServiceType As WebService)
        Select Case ServiceType
            Case WebService.Google
                Me.LV.Items(Clients(Index).LV_ID).SubItems(2).Text = "Done"

            Case WebService.MSN
                Me.LV.Items(Clients(Index).LV_ID).SubItems(3).Text = "Done"

            Case WebService.Yahoo
                Me.LV.Items(Clients(Index).LV_ID).SubItems(4).Text = "Done"

            Case WebService.Verizon
                Me.LV.Items(Clients(Index).LV_ID).SubItems(5).Text = "Done"

            Case WebService.ASK
                Me.LV.Items(Clients(Index).LV_ID).SubItems(6).Text = "Done"

            Case WebService.GA
                Me.LV.Items(Clients(Index).LV_ID).SubItems(7).Text = "Done"

        End Select
    End Sub
    '=============================================================================================
    '                                    EXCEPTIONS & FAILURES                                    
    '                                                                                             
    '=============================================================================================
    '=============================================================================================
    'Done_Failed() called by frmMain.vb Start_Engine()                                            
    'Sets label to "FAILED" in grid display                                                       
    'Sets Clients(Index).Engine.Downloaded = True in order to get out of looping                  
    '=============================================================================================
    Public Sub Done_Failed(ByVal Index As Integer, ByVal ServiceType As WebService)
        Done_Failed(Index, ServiceType, False)
    End Sub

    Public Sub Done_Failed(ByVal Index As Integer, ByVal ServiceType As WebService, ByVal bYahooCM As Boolean)
        Select Case ServiceType
            Case WebService.Google
                Clients(Index).Google.Downloaded = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(2).Text = "FAILED"

            Case WebService.MSN
                Clients(Index).MSN.Downloaded = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(3).Text = "FAILED"

            Case WebService.Yahoo
                If bYahooCM Then
                    Clients(Index).Yahoo.CM_Downloaded = True
                Else
                    Clients(Index).Yahoo.Key_Downloaded = True
                End If
                Me.LV.Items(Clients(Index).LV_ID).SubItems(4).Text = "FAILED"

            Case WebService.Verizon
                Clients(Index).Verizon.Downloaded = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(5).Text = "FAILED"

            Case WebService.ASK
                Clients(Index).ASK.Downloaded = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(6).Text = "FAILED"

            Case WebService.GA
                Clients(Index).GA.Downloaded = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(7).Text = "FAILED"
        End Select
    End Sub
    '=============================================================================================
    'Done_InvalidID() called by frmMain.vb Start_Verizon()                                        
    'Sets label to "ID ERROR" in grid display                                                     
    'Sets Clients(Index).Engine.Downloaded = True in order to get out of looping                  
    '=============================================================================================
    Public Sub Done_InvalidID(ByVal Index As Integer, ByVal ServiceType As WebService)
        Done_InvalidID(Index, ServiceType, False)
    End Sub

    Public Sub Done_InvalidID(ByVal Index As Integer, ByVal ServiceType As WebService, ByVal bYahooCM As Boolean)
        Select Case ServiceType
            Case WebService.Google
                Clients(Index).Google.Downloaded = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(2).Text = "ID ERROR"

            Case WebService.MSN
                Clients(Index).MSN.Downloaded = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(3).Text = "ID ERROR"

            Case WebService.Yahoo
                If bYahooCM Then
                    Clients(Index).Yahoo.CM_Downloaded = True
                Else
                    Clients(Index).Yahoo.Key_Downloaded = True
                End If
                Me.LV.Items(Clients(Index).LV_ID).SubItems(4).Text = "ID ERROR"

            Case WebService.Verizon
                Clients(Index).Verizon.Downloaded = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(5).Text = "ID ERROR"

            Case WebService.ASK
                Clients(Index).ASK.Downloaded = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(6).Text = "ID ERROR"

            Case WebService.GA
                Clients(Index).GA.Downloaded = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(7).Text = "ID ERROR"
        End Select
    End Sub
    'NO LONGER USED
    Public Sub Done_URLRequest(ByVal Index As Integer, ByVal ServiceTYpe As WebService)
        Done_URLRequest(Index, ServiceTYpe, False)
    End Sub

    Public Sub Done_URLRequest(ByVal Index As Integer, ByVal ServiceTYpe As WebService, ByVal bYahooCM As Boolean)
        Select Case ServiceTYpe
            Case WebService.Google
                Clients(Index).Google.UrlRequested = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(2).Text = "GETTING_URL"

            Case WebService.MSN
                Clients(Index).MSN.UrlRequested = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(3).Text = "GETTING_URL"

            Case WebService.Yahoo
                If bYahooCM Then
                    Clients(Index).Yahoo.CM_UrlRequested = True
                Else
                    Clients(Index).Yahoo.Key_UrlRequested = True
                End If
                Me.LV.Items(Clients(Index).LV_ID).SubItems(4).Text = "GETTING_URL"

            Case WebService.Verizon
                Clients(Index).Verizon.UrlRequested = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(5).Text = "GETTING_URL"

            Case WebService.ASK
                Clients(Index).ASK.UrlRequested = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(6).Text = "GETTING_URL"

            Case WebService.GA
                Clients(Index).GA.UrlRequested = True
                Me.LV.Items(Clients(Index).LV_ID).SubItems(7).Text = "GETTING_URL"
        End Select
    End Sub
#End Region 'Actions Done Subs

    Public Sub Freeze(ByVal bFreeze)
        Me.cmbFilter.Enabled = Not bFreeze
        Me.btnCustom.Enabled = Not bFreeze
    End Sub

    Public Function IsAllDownloaded_Yahoo() As Boolean
        Dim Result As Boolean = True

        For Each client As Clients In Clients
            If client.Yahoo.Has_AccountID Then
                If client.Yahoo.Key_Downloaded = False Then Result = False
                If client.Yahoo.CM_Downloaded = False Then Result = False
            End If
        Next

        Return Result
    End Function

    Public Function IsAllDownloaded_MSN() As Boolean
        Dim Result As Boolean = True

        For Each client As Clients In Clients
            If client.MSN.Has_AccountID Then
                If client.MSN.Downloaded = False Then Result = False
            End If
        Next

        Return Result
    End Function

    Public Function IsAllDownloaded_Google() As Boolean
        Dim Result As Boolean = True

        For Each client As Clients In Clients
            If client.Google.Has_AccountID Then
                If client.Google.Downloaded = False Then Result = False
            End If
        Next

        Return Result
    End Function

    Public Function IsAllDownloaded_Verizon() As Boolean
        Dim Result As Boolean = True

        For Each client As Clients In Clients
            If client.Verizon.Has_AccountID Then
                If client.Verizon.Downloaded = False Then Result = False
            End If
        Next

        Return Result
    End Function

    Public Function IsAnyReady_Verizon() As Boolean
        Dim Result As Boolean = False

        For Each client As Clients In Clients
            If client.Verizon.Has_AccountID Then
                If client.Verizon.DownloadReady Then
                    If client.Verizon.Downloaded = False Then Result = True
                End If
            End If
        Next

        Return Result
    End Function

    Public Function IsAllDownloaded_ASK() As Boolean
        Dim Result As Boolean = True

        For Each client As Clients In Clients
            If client.ASK.Has_AccountID Then
                If client.ASK.Downloaded = False Then Result = False
            End If
        Next

        Return Result
    End Function

    Public Function IsAllDownloaded_GA() As Boolean
        Dim Result As Boolean = True

        For Each client As Clients In Clients
            If client.GA.Has_AccountID Then
                If client.GA.Downloaded = False Then Result = False
            End If
        Next

        Return Result
    End Function


    '===============================================================================================================
    'My Custom Selection button pushed.  Rebuild the listview (LV) for individually selected clients
    '===============================================================================================================
    Private Sub btnCustom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCustom.Click
        Dim oClients As Clients() = Nothing
        Dim Index As Integer = 0

        'Loop each selected client and build oClients
        For Each client As Clients In Clients
            If Me.LV.Items(client.LV_ID).Selected Then
                If IsNothing(oClients) Then
                    Index = 0
                    ReDim oClients(Index)
                Else
                    Index = oClients.GetUpperBound(0) + 1
                End If
                ReDim Preserve oClients(Index)
                oClients(Index) = client
                oClients(Index).Index = Index
                oClients(Index).LV_ID = 0
            End If
        Next

        If IsNothing(oClients) Then Exit Sub

        Me.cmbFilter.SelectedIndex = -1
        Me.LV.Clear()
        Clients = oClients
        UpdateList()
    End Sub

End Class
