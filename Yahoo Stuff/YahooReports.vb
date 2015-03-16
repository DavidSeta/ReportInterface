Imports System
Imports System.IO
Imports System.Data
Imports System.Xml
Imports System.Net
Imports System.Collections
Imports System.Text
Imports System.Collections.Generic
Imports System.Globalization

Namespace YahooWebServices
    '
    'Class to interface with the class YahooServices.vb, namespace also YahooWebServices, generated via the wsdl.exe.
    '
    Public Class YahooReports
        Private WithEvents myService As YahooWebServices.BasicReportServiceService

        Private _BaseFilePath As String = ""
        Public DirectoryNamesAsRecID As Boolean = False

        Private Const IDS_EWS_ACCESS_HTTP_PROTOCOL As String = "https"
        'Production = "global.marketing.ews.yahooapis.com"; Sandbox = "sandbox.marketing.ews.yahooapis.com"
        Private Const IDS_EWS_LOCATION_SERVICE_ENDPOINT As String = "global.marketing.ews.yahooapis.com"
        Private Const IDS_EWS_VERSION As String = "V7"

        'Private Const idsDefault_UserName As String = "dleitch"
        'Private Const idsDefault_Password As String = "search3343"
        'Private Const idsDefault_License As String = "77F9E34A7A22A0F9"

        Private UserName As String = ""
        Private Password As String = ""
        Private License As String = ""
        Private GroupAcct As String = "anmv-ysm"
        Private GroupPWD As String = "AmericaNd820"

        Private _locationCache As Hashtable = New Hashtable()
        Private _MasterAccountToken As YahooWebServices.masterAccountID
        Private _SecurityToken As YahooWebServices.Security
        Private _LicenseToken As YahooWebServices.license
        Private _GroupAcct As YahooWebServices.onBehalfOfUsername
        Private _GroupPWD As YahooWebServices.onBehalfOfPassword


#Region " Output Properties (ReadOnly) "
        Public ReadOnly Property CommandGroup() As String
            Get
                Return myService.commandGroupValue.Text(0)
            End Get
        End Property

        Public ReadOnly Property RemainingQuota() As String
            Get
                Return myService.remainingQuotaValue.Text(0)
            End Get
        End Property

        Public ReadOnly Property TimeTakenMillis() As String
            Get
                Return myService.timeTakenMillisValue.Text(0)
            End Get
        End Property
#End Region 'Output Properties (ReadOnly)

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
                Return BaseFilePath & "Yahoo\"
            End Get
        End Property

#Region " Sub New (Overloaded)"
        Sub New(ByVal EngAuth As ClEngineSecurity)
            If UserName = "" Then UserName = EngAuth.YahooUserName
            If Password = "" Then Password = EngAuth.YahooPassword
            If License = "" Then License = EngAuth.YahooLicense

            SetTokens()
        End Sub

        '' ''Sub New(ByVal sUserName As String, ByVal sPassword As String, ByVal sLicense As String)
        '' ''    UserName = sUserName
        '' ''    Password = sPassword
        '' ''    License = sLicense

        '' ''    SetTokens()
        '' ''End Sub

        Private Sub SetTokens()
            _MasterAccountToken = New YahooWebServices.masterAccountID

            _SecurityToken = New YahooWebServices.Security
            _SecurityToken.UsernameToken = New YahooWebServices.UsernameToken
            _SecurityToken.UsernameToken.Username = UserName
            _SecurityToken.UsernameToken.Password = Password

            _LicenseToken = New YahooWebServices.license
            _LicenseToken.Text = New String() {License}


        End Sub
#End Region 'Sub New  (Overloaded)

#Region " New Client Stuff "
        Private Sub NewClient(ByVal sMasterID, ByVal sGroupAcct, ByVal sGroupPWD)
            'New On Behalf of User and PWD required
            _GroupAcct = New YahooWebServices.onBehalfOfUsername
            _GroupAcct.Text = New String() {sGroupAcct}

            _GroupPWD = New YahooWebServices.onBehalfOfPassword
            _GroupPWD.Text = New String() {sGroupPWD}

            _MasterAccountToken.Text = New String() {sMasterID}

            myService = New YahooWebServices.BasicReportServiceService
            'Set timeout so that the Yahoo server will initiate any timeouts..
            myService.Timeout = Integer.MaxValue

            Dim endPointLocation As String = getEndPoint(sMasterID, sGroupAcct, sGroupPWD)
            'If endPointLocation.Length < 7 Then

            'End If
            If Right(endPointLocation, 1) <> "/" Then endPointLocation = endPointLocation & "/"

            myService.Url = endPointLocation & IDS_EWS_VERSION & "/BasicReportService"
            myService.masterAccountIDValue = _MasterAccountToken
            myService.SecurityValue = _SecurityToken
            myService.licenseValue = _LicenseToken

            'Yahoo Group Logic
            If sGroupAcct <> "" Then
                myService.onBehalfOfPasswordValue = _GroupPWD
                myService.onBehalfOfUsernameValue = _GroupAcct
            End If

            'myService.
        End Sub

        Private Function getEndPoint(ByVal masterAccountID As String, ByVal sGroupAcct As String, ByVal sGroupPWD As String) As String


            'Place Location Service in Hashtable for subsequent use
            Dim endPointLocation As String = CType(_locationCache(masterAccountID), String)
            If Not IsNothing(endPointLocation) Then Return endPointLocation

            'We do not have in our Hashtable - Get From Location Service
            Try
                Dim LS As New YahooWebServices.LocationServiceService

                LS.Url = CType((IDS_EWS_ACCESS_HTTP_PROTOCOL & "://" & IDS_EWS_LOCATION_SERVICE_ENDPOINT & "/services/" & IDS_EWS_VERSION & "/LocationService"), String)
                LS.masterAccountIDValue = _MasterAccountToken
                LS.SecurityValue = _SecurityToken
                LS.licenseValue = _LicenseToken
                'Yahoo Group Logic
                If sGroupAcct <> "" Then
                    LS.onBehalfOfUsernameValue = _GroupAcct
                    LS.onBehalfOfPasswordValue = _GroupPWD
                End If

                'use LocationService to get the address prefix for the rest of the services
                endPointLocation = LS.getMasterAccountLocation()
                LS.Dispose()

                'persist location address for the master account
                _locationCache.Add(masterAccountID, endPointLocation)

            Catch ex As Exception
                Debug.Print("Exeption: " + ex.Message)
            End Try

            Return endPointLocation
        End Function
#End Region 'New Client Stuff

#Region " My Service Calls "
        Public Function RequestReport(ByVal myAccountID As String, ByVal BasicReportRequest As YahooWebServices.BasicReportRequest) As Long
            Dim nResult As Long = 0
            Dim oFileOutputFormat As FileOutputFormat = New FileOutputFormat
            oFileOutputFormat.fileOutputType = FileOutputType.TSV
            oFileOutputFormat.fileOutputTypeSpecified = True
            oFileOutputFormat.zipped = False
            oFileOutputFormat.zippedSpecified = True

            Try
                nResult = myService.addReportRequest(myAccountID, BasicReportRequest, oFileOutputFormat)
                'nResult = myService.addReportRequestForAccountID(myAccountID, BasicReportRequest)
            Catch ex As Exception
                'MsgBox(ex.Message)
            End Try

            Return nResult
        End Function

        Public Function GetReportURL(ByVal ReportID As Integer, ByVal FileOutputFormat As YahooWebServices.FileOutputFormat) As String
            Dim sResult As String = ""

            'Try
            '    sResult = myService.getReportOutputUrl(ReportID, FileOutputFormat)
            'Catch ex As Exception
            '    'MsgBox(ex.Message)
            'End Try
            Dim oReportDownloadInfo As ReportDownloadInfo = New ReportDownloadInfo
            oReportDownloadInfo = myService.getReportDownloadUrl(ReportID)
            sResult = oReportDownloadInfo.downloadUrl
            Return sResult

            
        End Function

        Public Function GetReportList(ByVal bOnlyCompleted As Boolean) As YahooWebServices.ReportInfo()
            Dim oResult() As YahooWebServices.ReportInfo = Nothing

            Try
                oResult = myService.getReportList(bOnlyCompleted)
            Catch ex As Exception
                'MsgBox(ex.Message)
            End Try

            Return oResult
        End Function

        'Public Function RequestReportAggregate(ByVal Client.Yahoo.AccountID,ByVal BasicReportRequest As YahooWebServices.BasicReportRequest) As Long
        '    Dim nResult As Long = 0

        '    Try
        '        nResult = myService.addReportRequestWithAccountAggregation(BasicReportRequest)
        '    Catch ex As Exception
        '        'MsgBox(ex.Message)
        '    End Try

        '    Return nResult
        'End Function

        Public Sub DeleteReports(ByVal ReportIDs() As Long)
            Try
                myService.deleteReports(ReportIDs)
            Catch ex As Exception
                'MsgBox(ex.Message)
            End Try
        End Sub

        Public Sub DeleteReport(ByVal ReportID As Long)
            Try
                myService.deleteReport(ReportID)
            Catch ex As Exception
                'MsgBox(ex.Message)
            End Try
        End Sub
#End Region 'My Service Calls

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


        '==================================================================================================================================================
        ' Use these for the Auto Stuff
        '==================================================================================================================================================
        Public Function ReqestDailyReportID(ByVal Client As Clients, ByVal CM As Boolean) As String
            'Set to New User
            NewClient(Client.Yahoo.AccountID, Client.Yahoo.GroupAcct, Client.Yahoo.GroupPWD)


            'Request Report
            Dim myID As Long = 0
            Try
                myID = Me.RequestReport(Client.Yahoo.YSMV6ID, GetDailyReportRequest(Client.Yahoo.AccountID, Client.Yahoo.StartDate, Client.Yahoo.EndDate, CM))
            Catch ex As Exception
                myID = 0
                Log("YAHOO ERROR ON " & Client.LogName(True) & ex.Message)
            End Try

            If myID = 0 Then Return ""

            Return myID.ToString
        End Function

        Public Sub ClearAllReports(ByVal Client As Clients)
            'Set to New User
            NewClient(Client.Yahoo.AccountID, Client.Yahoo.GroupAcct, Client.Yahoo.GroupPWD)

            Dim currentJobs() As YahooWebServices.ReportInfo = myService.getReportList(False)
            If IsNothing(currentJobs) Then Exit Sub

            Dim nCount As Integer
            For nCount = 0 To currentJobs.GetUpperBound(0)
                myService.deleteReport(currentJobs(nCount).reportID)
            Next

        End Sub

        Public Function GetReportStatus(ByVal Client As Clients, ByVal bYahooCM As Boolean) As String
            'Set to New User
            NewClient(Client.Yahoo.AccountID, Client.Yahoo.GroupAcct, Client.Yahoo.GroupPWD)

            Dim myID As Long
            If bYahooCM Then
                myID = CLng(Client.Yahoo.CM_ReportID)
            Else
                myID = CLng(Client.Yahoo.Key_ReportID)
            End If

            Dim sResult As String = "UNKNOWN"

            Dim status As YahooWebServices.ReportInfo() = myService.getReportList(False)
            If Not IsNothing(status) Then
                Dim nCount As Integer
                For nCount = 0 To status.GetUpperBound(0)
                    Dim nTempID As Long = status(nCount).reportID
                    If nTempID = myID Then
                        If status(nCount).status.ToUpper = "COMPLETE" Then
                            sResult = "READY"
                        Else
                            sResult = status(nCount).status
                        End If
                    End If
                Next
            End If

            Return sResult
        End Function

        Public Function DownloadReport(ByVal Client As Clients, ByVal bYahooCM As Boolean) As Boolean
            'Set to New User
            NewClient(Client.Yahoo.AccountID, Client.Yahoo.GroupAcct, Client.Yahoo.GroupPWD)

            'Get URL
            Dim bResult As Boolean = False
            Dim myID As Long
            If bYahooCM Then
                myID = CLng(Client.Yahoo.CM_ReportID)
            Else
                myID = CLng(Client.Yahoo.Key_ReportID)
            End If

            If myID = 0 Then Return bResult


            'Download
            If Not Directory.Exists(SavePath) Then Directory.CreateDirectory(SavePath)
            Dim myDir As String = Me.GetMyDir(Client)

            If Not Directory.Exists(myDir) Then Directory.CreateDirectory(myDir)
            Dim oReportDownloadInfo As ReportDownloadInfo = New ReportDownloadInfo

            Try
                Dim Format As YahooWebServices.FileOutputFormat = Me.GetYahooOUtputFormat
                oReportDownloadInfo = myService.getReportDownloadUrl(myID)
                Dim sURL As String = oReportDownloadInfo.downloadUrl

                'Dim sURL As String = myService.getReportOutputUrl(myID, Format)
                Log(Client.LogName(True) & "URL: " & sURL)
                If sURL <> "" Then

                    Dim sStart As String = ""
                    If Client.Yahoo.MultiDayReport Then
                        If bYahooCM Then
                            sStart = "YMa_"
                        Else
                            sStart = "YM_"
                        End If
                    Else
                        If bYahooCM Then
                            sStart = "YDa_"
                        Else
                            sStart = "YD_"
                        End If
                    End If
                    bResult = Me.SaveFile(sURL, myDir & sStart & Replace(Client.Name, " ", "") & Client.Yahoo.EndDate.ToString("_yyyy_MM_dd") & ".csv")

                    If bResult Then
                        Log("File Saved")
                    Else
                        Log("-- FILE NOT SAVED --")
                    End If
                Else
                    Log("-- FILE NOT SAVED (NO URL) --")
                End If

            Catch ex As Exception
                Log("YAHOO ERROR on " & Client.LogName(True) & ex.Message)

            End Try

            Try
                'Clear Job
                myService.deleteReport(myID)

            Catch ex As Exception
                Log("Clearing JOB on YAHOO for " & Client.LogName)
                Log("YAHOO ERROR on " & Client.LogName(True) & ex.Message)
            End Try

            Return bResult
        End Function

        Public Function DoesFileExist(ByVal Client As Clients, ByVal bYahooCM As Boolean) As Boolean

            If Not Directory.Exists(SavePath) Then Directory.CreateDirectory(SavePath)
            Dim myDir As String = Me.GetMyDir(Client)
            If Not Directory.Exists(myDir) Then Directory.CreateDirectory(myDir)

            Dim sFileName As String = ""
            Dim MyFile As String = ""
            Dim sCM As String = ""
            If bYahooCM Then sCM = "a"

            If Client.Yahoo.MultiDayReport Then
                MyFile = "YM" & sCM & "_" & Replace(Client.Name, " ", "") & Client.Yahoo.EndDate.ToString("_yyyy_MM_dd") & ".csv"
            Else
                MyFile = "YD" & sCM & "_" & Replace(Client.Name, " ", "") & Client.Yahoo.EndDate.ToString("_yyyy_MM_dd") & ".csv"
            End If
            sFileName = myDir & MyFile

            Return File.Exists(sFileName)
        End Function

        Public Function SaveFile(ByVal sUrl As String, ByVal sFileSave As String) As Boolean
            If File.Exists(sFileSave) Then File.Delete(sFileSave)

            Dim client As New WebClient
            client.DownloadFile(sUrl, sFileSave)
            Return File.Exists(sFileSave)
        End Function

        Public Function GetYahooOUtputFormat() As YahooWebServices.FileOutputFormat
            Dim Format As New YahooWebServices.FileOutputFormat
            Format.fileOutputType = YahooWebServices.FileOutputType.CSV_EXCEL
            Format.fileOutputTypeSpecified = True
            Format.zipped = False
            Format.zippedSpecified = True

            Return Format
        End Function

        Public Function GetDailyReportRequest(ByVal sAccountID As String, ByVal dStartDate As Date, ByVal dEndDate As Date, ByVal CM As Boolean) As YahooWebServices.BasicReportRequest
            Dim Request As New YahooWebServices.BasicReportRequest

            If CM Then
                Request.reportType = YahooWebServices.BasicReportType.AdGroupSummary
            Else
                Request.reportType = YahooWebServices.BasicReportType.KeywordSummary
            End If

            Request.reportTypeSpecified = True
            Request.dateRangeSpecified = False
            Request.startDate = DateAdd(DateInterval.Day, 1, dStartDate)
            Request.startDateSpecified = True
            Request.endDate = DateAdd(DateInterval.Day, 1, dEndDate)
            Request.endDateSpecified = True
            Request.reportName = sAccountID & "_" & dEndDate.ToString("MM-dd-yy")

            Return Request
        End Function
        '==================================================================================================================================================

        'Public Function GetDaily(ByVal sClientID As String, ByVal dDate As Date) As Boolean
        '    Dim myID As Integer = 0
        '    Dim sURL As String = ""
        '    Dim bFailed As Boolean = False
        '    Dim nClick As Integer = 0

        '    Debug.Print("")
        '    Debug.Print("---------------------------------")
        '    Debug.Print("    Yahoo Web Services")
        '    Debug.Print("---------------------------------")

        '    'Set to New User
        '    NewClient(Client.Yahoo.AccountID, Client.Yahoo.GroupAcct, Client.Yahoo.GroupPWD)
        '    NewClient(sClientID)
        '    Debug.Print("ClientID: " & sClientID)

        '    'Get Current Job List
        '    Dim currentJobs() As YahooWebServices.ReportInfo = myService.getReportList(False)
        '    Debug.Print("Clearing Current JObs")

        '    'Clear All jobs
        '    If IsNothing(currentJobs) Then
        '        Debug.Print("No Current Jobs")
        '    Else
        '        Dim nCount As Integer
        '        Debug.Print("Current Jobs: = " & currentJobs.GetUpperBound(0))
        '        For nCount = 0 To currentJobs.GetUpperBound(0)
        '            myService.deleteReport(currentJobs(nCount).reportID)
        '            Debug.Print("Old Job: " & currentJobs(nCount).reportID.ToString & " cleared")
        '        Next
        '        Debug.Print("Current Jobs Cleared")
        '        Debug.Print("")
        '    End If

        '    'Request Report
        '    myID = Me.RequestReportAggregate(GetDailyReportRequest(sClientID, dDate, dDate, False))
        '    Debug.Print("myID: " & myID.ToString)
        '    If myID = 0 Then
        '        Debug.Print("ID Error --  Failed Report")
        '        Return False
        '    End If

        '    'Wait Until Done
        '    While True
        '        Dim status As YahooWebServices.ReportInfo() = myService.getReportList(False)
        '        Dim nCount As Integer
        '        For nCount = 0 To status.GetUpperBound(0)
        '            Dim nTempID As Integer = status(nCount).reportID
        '            If nTempID = myID Then
        '                If status(nCount).status = "COMPLETE" Then
        '                    Debug.Print("Job Completed")
        '                    Exit While
        '                End If
        '                If status(nCount).status = "FAILED" Then
        '                    bFailed = True
        '                    Debug.Print("Job Failed")
        '                    Exit While
        '                End If
        '            End If
        '        Next
        '        nClick = nClick + 1
        '        Debug.Print("Sleeping(" & nClick.ToString("000") & ") - Status: " & status(0).status.ToString)
        '        System.Threading.Thread.Sleep(3000)
        '    End While

        '    'Get URL
        '    If myID <> 0 Then
        '        If bFailed = False Then
        '            sURL = myService.getReportOutputUrl(myID, Me.GetYahooOUtputFormat)
        '            Debug.Print("myURL: " & sURL)
        '        End If
        '    End If

        '    'Download
        '    If sURL <> "" Then
        '        Dim bresult As Boolean = Me.SaveFile(sURL, "C:\TestTim\Y_" & myID.ToString & ".csv")
        '        Debug.Print("Download Success: " & bresult.ToString)
        '    End If

        '    'Clear Job
        '    If myID <> 0 Then myService.deleteReport(myID)
        '    Debug.Print("Job: " & myID.ToString & " is cleared")

        '    Debug.Print("All Done")
        '    Debug.Print("---------------------------------")
        '    Debug.Print("")

        '    Return True
        'End Function

    End Class
End Namespace

