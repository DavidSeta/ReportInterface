Imports System.IO
Imports System.Data
Imports System.Net

Namespace VerizonWebServices
    Public Class VerizonReports
         
        Private WithEvents myService As VerizonWebServices.ReportingWebServiceClient
        'Private WithEvents myService As VerizonWebServices.ReportingServiceClient
        Private _BaseFilePath As String = ""
        Public DirectoryNamesAsRecID As Boolean = False

        'Private Const idsDefault_UserName As String = "CPROCTOR@MOREVISIBILITY.COM"
        'Private Const idsDefault_Password As String = "superpages"

        Public Username As String = ""
        Public Password As String = ""

        Public Property BaseFilePath() As String
            Get
                Return _BaseFilePath
            End Get
            Set(ByVal value As String)
                If value.Length > 0 Then
                    If Right(value, 1) <> "\" Then value = value & "\"
                End If
                _BaseFilePath = value
            End Set
        End Property

        Private ReadOnly Property SavePath() As String
            Get
                Return BaseFilePath & "Verizon\"
            End Get
        End Property

#Region " Sub New (Overloaded)"
        Sub New(ByVal EngAuth As ClEngineSecurity)
            If Username = "" Then Username = EngAuth.VerizonUserName
            If Password = "" Then Password = EngAuth.VerizonPassword

            NewService()
        End Sub

        Sub New(ByVal sUserName As String, ByVal sPassword As String)
            Username = sUserName
            Password = sPassword

            NewService()
        End Sub

        Private Sub NewService()
             
            myService = New VerizonWebServices.ReportingWebServiceClient
            myService.ClientCredentials.UserName.UserName = Username
            myService.ClientCredentials.UserName.Password = Password


        End Sub
#End Region 'Sub New  (Overloaded)

        '==================================================================================================================================================
        ' Invoke the WSDL methods for report generation.  Clients data structure passed in with all the params for report
        ' Invoked from frmMain.vb                                                                                        
        '==================================================================================================================================================
        Public Function ReqestDailyReportID(ByVal client As Clients) As String
            'Get parameters from the client DS passed 
            Dim sClientID As String = client.Verizon.AccountID

            'update myService.companyId
            'Removed July 2010 - will set here
            'SetClient(sClientID)

            'New July 2010
            Dim scheduleReportReq As VerizonWebServices.scheduleReport1
            scheduleReportReq = New VerizonWebServices.scheduleReport1

            scheduleReportReq.username = Username
            scheduleReportReq.password = Password
            scheduleReportReq.companyId = sClientID

            'Massage parameters to meet requirements

            Dim Param As New ReportParameters
            If client.Verizon.MultiDayReport Then
                Param.StartDate = client.Verizon.StartDate
                Param.EndDate = client.Verizon.EndDate
            Else
                Param.StartDate = client.Verizon.StartDate
                Param.EndDate = client.Verizon.StartDate
            End If

            Param.Username = Me.Username
            Param.Password = Me.Password
            Param.CompanyID = sClientID

            Dim sReportID As String = ""
            Dim oResult As String = ""
  
            Dim pReportType As String = Param.strReportType
            Dim pStartDate As String = Format(Param.StartDate, "MM/dd/yyyy")
            Dim pEndDate As String = Format(Param.EndDate, "MM/dd/yyyy")

            Dim oScheduleReportRequest As VerizonWebServices.scheduleReportScheduleReportRequest
            oScheduleReportRequest = New VerizonWebServices.scheduleReportScheduleReportRequest
            oScheduleReportRequest.reportType = pReportType
            oScheduleReportRequest.startDate = pStartDate
            oScheduleReportRequest.endDate = pEndDate

            '7/2010 - New Request object

            Dim oScheduleReport As VerizonWebServices.scheduleReport
            oScheduleReport = New VerizonWebServices.scheduleReport

            oScheduleReport.scheduleReportRequest = oScheduleReportRequest
            scheduleReportReq.scheduleReport = oScheduleReport

            Dim oReportResult As VerizonWebServices.scheduleReportResponse
            oReportResult = New VerizonWebServices.scheduleReportResponse

            Try
                'oReportResult = myService.scheduleReport(oReportReq)
                oReportResult = myService.scheduleReport(scheduleReportReq)

            Catch ex As Exception
                Debug.Print("Verizon Schedule Report Error on Client " & client.Name & " - " & ex.Message)
                Select Case Left(ex.Message.ToLower, 18)
                    Case "invalid company id"
                        sReportID = "ID ERROR"

                    Case "the underlying con" 'the underlying connection was closed:
                        sReportID = "RETRY"

                    Case "the operation has " 'the operation has timed out"
                        sReportID = "TIMEOUT"

                    Case Else
                        Debug.Print("Unknown Verizon Schedule Report Error on Client " & client.Name & " -- " & ex.Message)
                        sReportID = "ERROR"

                End Select
                Log("Verizon ReportID ERROR ON " & client.LogName(True) & ex.Message)

            End Try

            If sReportID = "" Then
                If Not IsNothing(oReportResult.return.reportId) Then
                    sReportID = oReportResult.return.reportId.ToString
                Else
                    sReportID = "ERROR"
                End If

            End If

            Debug.Print("Client: " & client.Name & " -- ReportID: " & sReportID)
            Return sReportID
        End Function

        '=======================================================================================================================
        ' Invoke the WSDL methods for retrieval of Report URL. Clients data structure passed in with all the params for reportID
        ' Invoked from frmMain.vb                                                                                               
        '=======================================================================================================================
        Public Function RequestReportURL(ByVal client As Clients) As String
            Dim sClientID As String = client.Verizon.AccountID
            Dim Request As String = ""
            Dim Result As String = ""
            Dim sURL As String = ""

            'Removed July 2010 - takes place here
            'update myService.companyId
            'SetClient(sClientID)

            'Dim getURLReq As VerizonWebServices.getReportURLRequest1
            'getURLReq = New VerizonWebServices.getReportURLRequest1

            Dim getURLReq As VerizonWebServices.getReportURL1
            getURLReq = New VerizonWebServices.getReportURL1
            getURLReq.username = Username
            getURLReq.password = Password
            getURLReq.companyId = sClientID

            '7/2010 - New Request object

            'Dim oGetReportURLReq As VerizonWebServices.GetReportURLRequest
            'oGetReportURLReq = New VerizonWebServices.GetReportURLRequest

            Dim oGetReportURLReq As VerizonWebServices.getReportURL
            oGetReportURLReq = New VerizonWebServices.getReportURL
            Dim oGetReportURL As VerizonWebServices.getReportURLGetReportURLRequest
            oGetReportURL = New VerizonWebServices.getReportURLGetReportURLRequest
            oGetReportURL.reportId = client.Verizon.ReportID

            oGetReportURLReq.getReportURLRequest = oGetReportURL
            getURLReq.getReportURL = oGetReportURLReq

            'July 2010 New Response object
            Dim oReportURLResult As VerizonWebServices.getReportURLResponse
            oReportURLResult = New VerizonWebServices.getReportURLResponse

            Try
                oReportURLResult = myService.getReportURL(getURLReq)
            Catch ex As Exception
                Select Case ex.Message.ToLower
                    Case "the report is not yet complete, please try again later"
                        sURL = "RETRY"

                    Case "the report you requested is empty"
                        sURL = "EMPTY"

                    Case Else
                        sURL = "UNKNOWN"
                        Debug.Print("Verizon URL Error ClientID " & client.Name & " -- " & ex.Message)
                End Select
                Log("Verizon URL ERROR ON " & client.LogName(True) & ex.Message)

            End Try

            If sURL = "" Then
                If Not IsNothing(oReportURLResult.return.reportURL) Then
                    sURL = oReportURLResult.return.reportURL
                Else
                    sURL = "ERROR"
                End If
            End If


            'Edit for report not found returned
            Select Case sURL.ToLower
                Case "report not found"
                    sURL = "RETRY"
                Case "report is in the queue"
                    sURL = "RETRY"
            End Select
            Return sURL
        End Function

        Public Function DownloadReport(ByVal Client As Clients) As Boolean
            Dim bResult As Boolean = True
            Dim sFileName As String = ""
            ''update myService.companyId
            'SetClient(Client.Verizon.AccountID)

            If Not Directory.Exists(SavePath) Then Directory.CreateDirectory(SavePath)
            Dim myDir As String = Me.GetMyDir(Client)

            If Not Directory.Exists(myDir) Then Directory.CreateDirectory(myDir)

            If Client.Verizon.MultiDayReport Then
                sFileName = myDir & "VM_" & Replace(Client.Name, " ", "") & Client.Verizon.EndDate.ToString("_yyyy_MM_dd") & ".xml"
            Else
                sFileName = myDir & "VD_" & Replace(Client.Name, " ", "") & Client.Verizon.EndDate.ToString("_yyyy_MM_dd") & ".xml"
            End If

            Try
                Dim sURL As String = Client.Verizon.URL
                Log(Client.LogName(True) & "URL: " & sURL)
                If sURL <> "" Then
                    bResult = Me.SaveFile(sURL, sFileName)
                    If bResult Then
                        Log("File Saved")
                    Else
                        Log("--- FILE NOT SAVED ---")
                    End If
                Else
                    Log("--- FILE NOT SAVED (NO URL) ---")
                End If

            Catch ex As Exception
                Log("VERIZON ERROR ON " & Client.LogName(True) & ex.Message)

            End Try

            Try
                'Clear Job on WebService
                Log("Clearing JOB on Verizon for " & Client.LogName)
                DeleteReport(Client.Verizon.AccountID, Client.Verizon.ReportID)
                Log("Verizon Job Cleared for " & Client.LogName)

            Catch ex As Exception
                Log("VERIZON ERROR ON " & Client.LogName(True) & ex.Message)

            End Try

            Return bResult
        End Function

        Public Function DoesFileExist(ByVal Client As Clients) As Boolean

            If Not Directory.Exists(SavePath) Then Directory.CreateDirectory(SavePath)
            Dim myDir As String = Me.GetMyDir(Client)
            If Not Directory.Exists(myDir) Then Directory.CreateDirectory(myDir)

            Dim sFileName As String = ""
            Dim MyFile As String = ""
            Dim sCM As String = ""

            If Client.Verizon.MultiDayReport Then
                MyFile = "VM_" & Replace(Client.Name, " ", "") & Client.Verizon.EndDate.ToString("_yyyy_MM_dd") & ".xml"
            Else
                MyFile = "VD_" & Replace(Client.Name, " ", "") & Client.Verizon.EndDate.ToString("_yyyy_MM_dd") & ".xml"
            End If

            sFileName = myDir & MyFile

            Return File.Exists(sFileName)
        End Function
        '===================================================================================================
        'GetMyDir - Build directory path
        '===================================================================================================

        Private Function GetMyDir(ByVal Client As Clients) As String
            '----------------------------------------------------------------------------------------------------------------
            ' Changed 3/14/07 so directories are now RecID's instead of names
            '----------------------------------------------------------------------------------------------------------------
            Dim sDir As String
            If DirectoryNamesAsRecID Then
                sDir = SavePath & Client.CustID & "\"
            Else
                sDir = SavePath & Client.Name & "\"
            End If
            Return sDir
        End Function

        '===================================================================================================
        'DeleteReport - Clean up on Verizon 
        '===================================================================================================
        Private Sub DeleteReport(ByVal sClientID As String, ByVal sReportID As String)

            Dim Request As String = ""
            Dim Result As String = ""
            Dim sURL As String = ""

            Dim dltReportReq As VerizonWebServices.deleteReport1
            dltReportReq = New VerizonWebServices.deleteReport1
            dltReportReq.username = Username
            dltReportReq.password = Password
            dltReportReq.companyId = sClientID

            '7/2010 - New Request object

            Dim odeleteReport As VerizonWebServices.deleteReportDeleteReportRequest
            odeleteReport = New VerizonWebServices.deleteReportDeleteReportRequest
            odeleteReport.reportId = sReportID

            Dim oDltReportReq As VerizonWebServices.deleteReport
            oDltReportReq = New VerizonWebServices.deleteReport
            oDltReportReq.deleteReportRequest = odeleteReport

            dltReportReq.deleteReport = oDltReportReq

            'July 2010 New Response object
            Dim oDltReportResult As VerizonWebServices.deleteReportResponse
            oDltReportResult = New VerizonWebServices.deleteReportResponse

            Try
                oDltReportResult = myService.deleteReport(dltReportReq)
            Catch ex As Exception               
                Log("Verizon Delete Report ERROR ON " & sClientID & " " & ex.Message)
            End Try
 
        End Sub

        'Public Sub GetReportStatus2_Async(ByVal sClientID As String)
        '    Dim p As New Test.GetReportListRequest
        '    p.companyId = sClientID
        '    p.username = Me.Username
        '    p.password = Me.Password

        '    myServiceX.getReportListAsync(p)
        'End Sub

        'Removed July 2010 now SetClient is performed in each function
        '=============================================================================
        ' SetClient - Update WSDL value for clientID we are running reports
        '=============================================================================
        'Private Sub SetClient(ByVal sClientID As String)
        '    Dim CompanyID = New VerizonWebServices.ReportParameters
        '    CompanyID.Text = New String() {sClientID}
        '    myService.companyIdValue.Text = CompanyID
        'End Sub
        '=============================================================================
        ' SaveFile - Using System.Net.WebClient.DownloadFile(sURI,sFileSave)
        '=============================================================================
        Private Function SaveFile(ByVal sURl As String, ByVal sFileSave As String) As Boolean
            If File.Exists(sFileSave) Then File.Delete(sFileSave)

            If sURl = "EMPTY" Then
                File.WriteAllText(sFileSave, "<Results></Results>")
            Else
                Dim client As New WebClient
                client.DownloadFile(sURl, sFileSave)
            End If

            Return File.Exists(sFileSave)
        End Function



        'Public Function CreateUrlRequest(ByVal sUsername As String, ByVal sPassword As String, ByVal sCompanyID As String, ByVal sReportID As String) As VerizonWebServices.GetReportURLRequest
        '    Dim req As New VerizonWebServices.GetReportURLRequest
        '    req.username = sUsername
        '    req.password = sPassword
        '    req.companyId = sCompanyID
        '    req.reportId = sReportID
        '    Return req
        'End Function

        'DK 11/08/2007 Removed since not used
        'Public Sub ReqestDailyReportID_Async(ByVal client As Clients)
        '    Dim sClientID As String = client.Verizon.AccountID
        '    'update myService.companyId
        '    SetClient(sClientID)


        '    Dim Request As New VerizonWebServices.ScheduleReportRequest
        '    If client.Verizon.MultiDayReport Then
        '        Request = Me.CreateKeywordReportRequest(Me.Username, Me.Password, sClientID, client.Verizon.StartDate, client.Verizon.EndDate)
        '    Else
        '        Request = Me.CreateKeywordReportRequest(Me.Username, Me.Password, sClientID, client.Verizon.EndDate)
        '    End If

        '    myService.scheduleReportAsync(Request, client.Index)
        'End Sub
        'DK 11/08/2007 Removed since not used
        'Public Sub RequestReportURL_Async(ByVal client As Clients)
        '    Dim sClientID As String = client.Verizon.AccountID
        '    'update myService.companyId
        '    SetClient(sClientID)

        '    Dim Request As New VerizonWebServices.GetReportURLRequest
        '    Request = Me.CreateUrlRequest(Me.Username, Me.Password, sClientID, client.Verizon.ReportID)
        '    myService.getReportURLAsync(Request, client.Index)
        'End Sub

        'Public Function GetReportStatus(ByVal sClientID As String, ByVal sReportID As String) As VerizonWebServices.GetReportListResponse
        '    'update myService.companyId
        '    SetClient(sClientID)

        '    Dim ListRequest As New VerizonWebServices.GetReportListRequest
        '    ListRequest.username = Me.Username
        '    ListRequest.password = Me.Password
        '    ListRequest.companyId = sClientID

        '    Dim ReportList As VerizonWebServices.GetReportListResponse = Nothing
        '    Try
        '        ReportList = myService.getReportList(ListRequest)

        '    Catch ex As Exception
        '        Debug.Print(ex.Message)
        '        'Log("Verizon STATUS ERROR ON " & Client.LogName(True) & ex.Message)
        '    End Try
        '    Return Nothing

        '    'Return ReportList
        'End Function

        'Private Function GetReportURL(ByVal sClientID As String, ByVal sReportID As String) As String
        '    'update myService.companyId
        '    SetClient(sClientID)

        '    Dim paramURL As New VerizonWebServices.GetReportURLRequest
        '    paramURL.companyId = sClientID
        '    paramURL.username = Me.Username
        '    paramURL.password = Me.Password
        '    paramURL.reportId = sReportID

        '    Dim Response As VerizonWebServices.GetReportURLResponse = Nothing
        '    Try
        '        Response = myService.getReportURL(paramURL)

        '    Catch ex As Exception
        '        'Log("Verizon URL ERROR ON " & Client.LogName(True) & ex.Message)

        '    End Try
        '    Return Response.URL
        'End Function
        'Private Sub myService_getReportURLCompleted(ByVal sender As Object, ByVal e As getReportURLCompletedEventArgs) Handles myService.getReportURLCompleted
        '    Try
        '        Dim Index As Integer = CInt(e.UserState)
        '        If IsNothing(e.Error) Then
        '            'Return URL
        '            RaiseEvent GotURL(Index, e.Result.URL)
        '        Else
        '            'Return Something else
        '            Select Case e.Error.Message.ToLower
        '                Case "the report is not yet complete, please try again later"
        '                    RaiseEvent GotURL(Index, "RETRY")

        '                Case "the report you requested is empty"
        '                    RaiseEvent GotURL(Index, "EMPTY")

        '                Case Else
        '                    RaiseEvent GotURL(Index, "UNKNOWN")
        '                    Debug.Print("URL Error (ClientID:" & Index.ToString("00") & ") " & e.Error.Message)
        '            End Select
        '        End If

        '    Catch ex As Exception
        '        Dim a As String = ""
        '    End Try
        'End Sub

        'Private Sub myService_scheduleReportCompleted(ByVal sender As Object, ByVal e As scheduleReportCompletedEventArgs) Handles myService.scheduleReportCompleted
        '    Try

        '        'invalid company id
        '        'Request Error (ClientID:00) invalid company id

        '        Dim Index As Integer = CInt(e.UserState)
        '        If IsNothing(e.Error) Then
        '            RaiseEvent Scheduled(Index, e.Result.reportId)
        '        Else
        '            Select Case e.Error.Message.ToLower
        '                Case "invalid company id"
        '                    RaiseEvent Scheduled(Index, "ID ERROR")

        '                Case Else
        '                    Debug.Print("Request Error (ClientID:" & Index.ToString("00") & ") " & e.Error.Message)
        '                    MsgBox("(" & Index.ToString("00") & ") " & e.Error.Message)

        '                    'Server closed connection
        '                    RaiseEvent Scheduled(Index, "ERROR")
        '            End Select
        '        End If
        '    Catch ex As Exception

        '    End Try
        'End Sub
#Region " CreateKeywordReportRequest (Overloaded)"
        '        Public Function CreateKeywordReportRequest(ByVal params As ReportParameters) As VerizonWebServices.ScheduleReportRequest
        '            Dim Request As New VerizonWebServices.ScheduleReportRequest
        '            Request.companyId = params.CompanyID
        '            Request.endDate = params.EndDate.ToString("MM/dd/yyyy")
        '            Request.password = params.Password
        '            Request.reportType = params.strReportType
        '            Request.startDate = params.StartDate.ToString("MM/dd/yyyy")
        '            Request.username = params.UserName
        '            Return Request
        '        End Function

        '        Public Function CreateKeywordReportRequest(ByVal sUserName As String, ByVal sPassword As String, ByVal sCompanyID As String, ByVal dSingleDate As Date) As VerizonWebServices.ScheduleReportRequest
        '            Dim Param As New ReportParameters
        '            Param.StartDate = dSingleDate
        '            Param.EndDate = dSingleDate
        '            Param.UserName = sUserName
        '            Param.Password = sPassword
        '            Param.CompanyID = sCompanyID
        '            Return Me.CreateKeywordReportRequest(Param)
        '        End Function

        '        Public Function CreateKeywordReportRequest(ByVal sUsername As String, ByVal sPassword As String, ByVal sCompanyID As String, ByVal dStartDate As Date, ByVal dEndDate As Date) As VerizonWebServices.ScheduleReportRequest
        '            Dim Param As New ReportParameters
        '            Param.StartDate = dStartDate
        '            Param.EndDate = dEndDate
        '            Param.UserName = sUsername
        '            Param.Password = sPassword
        '            Param.CompanyID = sCompanyID
        '            Return Me.CreateKeywordReportRequest(Param)
        '        End Function
#End Region 'CreateKeywordReportRequest (Overloaded)
        'Private Sub myServiceX_getReportListCompleted(ByVal sender As Object, ByVal e As Test.getReportListCompletedEventArgs) Handles myServiceX.getReportListCompleted

        'End Sub
    End Class


#Region " ReportParameters Class "
    Public Class ReportParameters
        Public Username As String = ""
        Public Password As String = ""
        Public CompanyID As String = ""
        Public ReportID As String = ""

        Private _ReportType As String = "KeywordActivity"
        Public StartDate As Date = Today
        Public EndDate As Date = Today

        Sub New()
            'Set the Default Values
            UserName = ""
            Password = ""
            ReportID = ""
            ReportType = ReportTypes.KeywordActivity
            StartDate = Today
            EndDate = Today
        End Sub

        Public ReadOnly Property strReportType()
            Get
                Return _ReportType
            End Get
        End Property

        Public Property ReportType() As ReportTypes
            Get
                Select Case _ReportType.ToLower
                    Case "keywordactivity"
                        Return ReportTypes.KeywordActivity

                    Case "adoverview"
                        Return ReportTypes.AdOverview

                    Case "activityreport"
                        Return ReportTypes.ActivityReport

                    Case Else
                        Return ReportTypes.KeywordActivity

                End Select
            End Get
            Set(ByVal value As ReportTypes)
                Select Case value
                    Case ReportTypes.KeywordActivity
                        _ReportType = "KeywordActivity"

                    Case ReportTypes.AdOverview
                        _ReportType = "AdOverview"

                    Case ReportTypes.ActivityReport
                        _ReportType = "ActivityReport"

                    Case Else
                        _ReportType = "KeywordActivity"

                End Select

            End Set
        End Property
    End Class
#End Region 'ReportParameters Class

    Public Enum ReportTypes
        KeywordActivity
        AdOverview
        ActivityReport
    End Enum

End Namespace
