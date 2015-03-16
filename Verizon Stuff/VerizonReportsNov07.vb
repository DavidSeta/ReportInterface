Imports System.IO
Imports System.Data
Imports System.Net

Namespace VerizonWebServices
    Public Class VerizonReports
        Private WithEvents myService As VerizonWebServices.ReportingService
        Private WithEvents myServiceX As Test.reporting

        Public Event Scheduled(ByVal ClientIndex As Integer, ByVal ReportID As String)
        Public Event GotURL(ByVal ClientIndex As Integer, ByVal URL As String)

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
            myServiceX = New Test.reporting

            myService = New VerizonWebServices.ReportingService

            Dim UserValue As New VerizonWebServices.string2
            UserValue.Text = New String() {Username}

            Dim Passwordvalue As New VerizonWebServices.[string]
            Passwordvalue.Text = New String() {Password}

            myService.usernameValue = UserValue
            myService.passwordValue = Passwordvalue

        End Sub
#End Region 'Sub New  (Overloaded)

#Region " CreateKeywordReportRequest (Overloaded)"
        Public Function CreateKeywordReportRequest(ByVal params As ReportParameters) As VerizonWebServices.ScheduleReportRequest
            Dim Request As New VerizonWebServices.ScheduleReportRequest
            Request.companyId = params.CompanyID
            Request.endDate = params.EndDate.ToString("MM/dd/yyyy")
            Request.password = params.Password
            Request.reportType = params.strReportType
            Request.startDate = params.StartDate.ToString("MM/dd/yyyy")
            Request.username = params.UserName
            Return Request
        End Function

        Public Function CreateKeywordReportRequest(ByVal sUserName As String, ByVal sPassword As String, ByVal sCompanyID As String, ByVal dSingleDate As Date) As VerizonWebServices.ScheduleReportRequest
            Dim Param As New ReportParameters
            Param.StartDate = dSingleDate
            Param.EndDate = dSingleDate
            Param.UserName = sUserName
            Param.Password = sPassword
            Param.CompanyID = sCompanyID
            Return Me.CreateKeywordReportRequest(Param)
        End Function

        Public Function CreateKeywordReportRequest(ByVal sUsername As String, ByVal sPassword As String, ByVal sCompanyID As String, ByVal dStartDate As Date, ByVal dEndDate As Date) As VerizonWebServices.ScheduleReportRequest
            Dim Param As New ReportParameters
            Param.StartDate = dStartDate
            Param.EndDate = dEndDate
            Param.UserName = sUsername
            Param.Password = sPassword
            Param.CompanyID = sCompanyID
            Return Me.CreateKeywordReportRequest(Param)
        End Function
#End Region 'CreateKeywordReportRequest (Overloaded)

        Public Function CreateUrlRequest(ByVal sUsername As String, ByVal sPassword As String, ByVal sCompanyID As String, ByVal sReportID As String) As VerizonWebServices.GetReportURLRequest
            Dim req As New VerizonWebServices.GetReportURLRequest
            req.username = sUsername
            req.password = sPassword
            req.companyId = sCompanyID
            req.reportId = sReportID
            Return req
        End Function

        '==================================================================================================================================================
        ' Use these for the Auto Stuff
        '==================================================================================================================================================
        Public Function ReqestDailyReportID(ByVal client As Clients) As String
            Dim sClientID As String = client.Verizon.AccountID
            SetClient(sClientID)

            Dim Request As New VerizonWebServices.ScheduleReportRequest
            If client.Verizon.MultiDayReport Then
                Request = Me.CreateKeywordReportRequest(Me.Username, Me.Password, sClientID, client.Verizon.StartDate, client.Verizon.EndDate)
            Else
                Request = Me.CreateKeywordReportRequest(Me.Username, Me.Password, sClientID, client.Verizon.EndDate)
            End If

            Dim sReportID As String = ""
            Dim oResult As VerizonWebServices.ScheduleReportResponse = Nothing
            Try
                oResult = myService.scheduleReport(Request)
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
                If Not IsNothing(oResult) Then sReportID = oResult.reportId
            End If

            Debug.Print("Client: " & client.Name & " -- ReportID: " & sReportID)
            Return sReportID
        End Function

        Public Sub ReqestDailyReportID_Async(ByVal client As Clients)
            Dim sClientID As String = client.Verizon.AccountID
            SetClient(sClientID)

            Dim Request As New VerizonWebServices.ScheduleReportRequest
            If client.Verizon.MultiDayReport Then
                Request = Me.CreateKeywordReportRequest(Me.Username, Me.Password, sClientID, client.Verizon.StartDate, client.Verizon.EndDate)
            Else
                Request = Me.CreateKeywordReportRequest(Me.Username, Me.Password, sClientID, client.Verizon.EndDate)
            End If

            myService.scheduleReportAsync(Request, client.Index)
        End Sub

        Public Function RequestReportURL(ByVal client As Clients) As String
            Dim sClientID As String = client.Verizon.AccountID
            SetClient(sClientID)

            Dim Request As New VerizonWebServices.GetReportURLRequest
            Request = Me.CreateUrlRequest(Me.Username, Me.Password, sClientID, client.Verizon.ReportID)

            Dim sURL As String = ""
            Dim Result As VerizonWebServices.GetReportURLResponse = Nothing
            Try
                Result = myService.getReportURL(Request)
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
                If Not IsNothing(Result) Then sURL = Result.URL
            End If

            Return sURL
        End Function

        Public Sub RequestReportURL_Async(ByVal client As Clients)
            Dim sClientID As String = client.Verizon.AccountID
            SetClient(sClientID)

            Dim Request As New VerizonWebServices.GetReportURLRequest
            Request = Me.CreateUrlRequest(Me.Username, Me.Password, sClientID, client.Verizon.ReportID)
            myService.getReportURLAsync(Request, client.Index)
        End Sub

        Public Function DownloadReport(ByVal Client As Clients) As Boolean
            Dim bResult As Boolean = True
            Dim sFileName As String = ""

            SetClient(Client.Verizon.AccountID)

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

            If Client.Google.MultiDayReport Then
                MyFile = "VM_" & Replace(Client.Name, " ", "") & Client.Verizon.EndDate.ToString("_yyyy_MM_dd") & ".xml"
            Else
                MyFile = "VD_" & Replace(Client.Name, " ", "") & Client.Verizon.EndDate.ToString("_yyyy_MM_dd") & ".xml"
            End If

            sFileName = myDir & MyFile

            Return File.Exists(sFileName)
        End Function

        '==================================================================================================================================================

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


        Private Function DeleteReport(ByVal sClientID As String, ByVal sReportID As String) As String
            SetClient(sClientID)

            Dim DeleteRequest As New VerizonWebServices.DeleteReportRequest
            DeleteRequest.companyId = sClientID
            DeleteRequest.username = Me.Username
            DeleteRequest.password = Me.Password
            DeleteRequest.reportId = sReportID

            Dim DeleteResponse As VerizonWebServices.DeleteReportResponse = myService.deleteReport(DeleteRequest)

            Return DeleteResponse.successCode
        End Function

        Public Sub GetReportStatus2_Async(ByVal sClientID As String)
            Dim p As New Test.GetReportListRequest
            p.companyId = sClientID
            p.username = Me.Username
            p.password = Me.Password

            myServiceX.getReportListAsync(p)
        End Sub

        Public Function GetReportStatus(ByVal sClientID As String, ByVal sReportID As String) As VerizonWebServices.GetReportListResponse
            SetClient(sClientID)

            Dim ListRequest As New VerizonWebServices.GetReportListRequest
            ListRequest.username = Me.Username
            ListRequest.password = Me.Password
            ListRequest.companyId = sClientID

            Dim ReportList As VerizonWebServices.GetReportListResponse = Nothing
            Try
                ReportList = myService.getReportList(ListRequest)

            Catch ex As Exception
                Debug.Print(ex.Message)
                'Log("Verizon STATUS ERROR ON " & Client.LogName(True) & ex.Message)
            End Try
            Return Nothing

            'Return ReportList
        End Function

        Private Function GetReportURL(ByVal sClientID As String, ByVal sReportID As String) As String
            SetClient(sClientID)

            Dim paramURL As New VerizonWebServices.GetReportURLRequest
            paramURL.companyId = sClientID
            paramURL.username = Me.Username
            paramURL.password = Me.Password
            paramURL.reportId = sReportID

            Dim Response As VerizonWebServices.GetReportURLResponse = Nothing
            Try
                Response = myService.getReportURL(paramURL)

            Catch ex As Exception
                'Log("Verizon URL ERROR ON " & Client.LogName(True) & ex.Message)

            End Try
            Return Response.URL
        End Function

        Private Sub SetClient(ByVal sClientID As String)
            Dim CompanyID As New VerizonWebServices.companyId
            CompanyID.Text = New String() {sClientID}
            myService.companyIdValue = CompanyID
        End Sub

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

        Private Sub myService_getReportListCompleted(ByVal sender As Object, ByVal e As getReportListCompletedEventArgs) Handles myService.getReportListCompleted

        End Sub

        Private Sub myService_getReportURLCompleted(ByVal sender As Object, ByVal e As getReportURLCompletedEventArgs) Handles myService.getReportURLCompleted
            Try
                Dim Index As Integer = CInt(e.UserState)
                If IsNothing(e.Error) Then
                    'Return URL
                    RaiseEvent GotURL(Index, e.Result.URL)
                Else
                    'Return Something else
                    Select Case e.Error.Message.ToLower
                        Case "the report is not yet complete, please try again later"
                            RaiseEvent GotURL(Index, "RETRY")

                        Case "the report you requested is empty"
                            RaiseEvent GotURL(Index, "EMPTY")

                        Case Else
                            RaiseEvent GotURL(Index, "UNKNOWN")
                            Debug.Print("URL Error (ClientID:" & Index.ToString("00") & ") " & e.Error.Message)
                    End Select
                End If

            Catch ex As Exception
                Dim a As String = ""
            End Try
        End Sub

        Private Sub myService_scheduleReportCompleted(ByVal sender As Object, ByVal e As scheduleReportCompletedEventArgs) Handles myService.scheduleReportCompleted
            Try

                'invalid company id
                'Request Error (ClientID:00) invalid company id

                Dim Index As Integer = CInt(e.UserState)
                If IsNothing(e.Error) Then
                    RaiseEvent Scheduled(Index, e.Result.reportId)
                Else
                    Select Case e.Error.Message.ToLower
                        Case "invalid company id"
                            RaiseEvent Scheduled(Index, "ID ERROR")

                        Case Else
                            Debug.Print("Request Error (ClientID:" & Index.ToString("00") & ") " & e.Error.Message)
                            MsgBox("(" & Index.ToString("00") & ") " & e.Error.Message)

                            'Server closed connection
                            RaiseEvent Scheduled(Index, "ERROR")
                    End Select
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub myServiceX_getReportListCompleted(ByVal sender As Object, ByVal e As Test.getReportListCompletedEventArgs) Handles myServiceX.getReportListCompleted

        End Sub
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
