Imports System.IO
Imports System.Data
Imports System.Xml
Imports System.Net
Imports System.Text


Namespace ASKWebServices
    Public Class ASKReports
        Private WithEvents myService As ASKWebServices.ReportsService
        Private WithEvents myCMService As ASKWebServices.CampaignsService
        Private WithEvents myAdService As ASKWebServices.AdsService

        Private _BaseFilePath As String = ""
        Public DirectoryNamesAsRecID As Boolean = False

        Private idsDefault_Password As String
        Public Useragent As String
        Public Email As String
        Public Password As String
        Public Token As String
        Public AppToken As String

        Public AccountName As String

        Public StartDate As Date = Today
        Public EndDate As Date = Today

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
                Return BaseFilePath & "ASK\"
            End Get
        End Property

#Region " Sub New "
        '========================================================================================
        'New(ByVal engineAuthentication As ClEngineSecurity)
        'Called by frmMain_Load()
        '========================================================================================
        Sub New(ByVal engineAuthentication As ClEngineSecurity)
            If Token = "" Then Token = engineAuthentication.ASKLicense
            'ASK mandate to turn off the HTTP Header "Expect: 100-Continue"
            'For explaination of this issue:
            'http://haacked.com/archive/2004/05/15/http-web-request-expect-100-continue.aspx
            System.Net.ServicePointManager.Expect100Continue = False
            NewService()
        End Sub

        Sub NewService()
            myService = New ASKWebServices.ReportsService
            myCMService = New ASKWebServices.CampaignsService
            myAdService = New ASKWebServices.AdsService
        End Sub
#End Region 'Sub New


#Region " My Functions invoked via frmMain Start_ASK()"
        '=============================================================================================
        'ReqestReportID(ByVal client As Clients) called by frmMain.vb Start_ASK()
        'Return String with ReportID or verbiage reflecting an error occurred while requesting: "ERROR"
        'Perform initialization 
        '=============================================================================================
        Public Function ReqestReportID(ByVal client As Clients) As String
            Dim myID As Long = -1
            'Build File Name
            Dim sFile As String = BldFileName(client)

            'Instantiate and populate properties for myService via NewService():
            SetClient(client.ASK.AccountID, client.ASK.AccountPWD)

            'Request report Dim oClients() As Clients = db.GetClientList
            Dim oRsp() As ReportInterface.ASKWebServices.KeywordReportStruct

            'Request report variable oCMRsp
            Dim oCMRsp As ReportInterface.ASKWebServices.CampaignStruct
            'Request report variable oAdRsp
            Dim oAdRsp As ReportInterface.ASKWebServices.AdStruct

            Try
                oRsp = myService.GetAllKeywordReports(Token, client.ASK.AccountID, client.ASK.StartDate, client.ASK.EndDate)
            Catch ex As Exception
                Log("ASK ERROR ON " & client.LogName(True) & ex.Message)
                myID = 0

            End Try
            'Had problems with API - return an error indicator
            If myID = 0 Then Return "ERROR"



            'Build a Hashtable object to hold CampaignID --> Description & AdID --> Description
            Dim htCMDesc As Hashtable = New Hashtable()
            Dim htAdDesc As Hashtable = New Hashtable()
            Dim nCMID As Integer
            Dim nAdID As Integer
            Dim sCMDesc As String = String.Empty
            Dim sAdDesc As String = String.Empty

            'Write to file
            Dim objWriter As StreamWriter
            objWriter = New StreamWriter(sFile)
            Dim sLine As New StringBuilder()
            Dim nCount As Integer

            'Build header line
            sLine.Append("Campaign,AdGroup,Keyword,Clicks,Costs,CTR,CPC,AvgPos")
            sLine.AppendLine()
            objWriter.Write(sLine)
            sLine.Length = 0

            Dim sVar As String = String.Empty

            'Loop all Keywords returned from the API
            For nCount = 0 To oRsp.GetUpperBound(0)
                'Do we have the Campaign Description already; if not get it
                nCMID = oRsp(nCount).campaign_id
                If htCMDesc.ContainsKey(nCMID) Then
                    sCMDesc = htCMDesc(nCMID).ToString
                Else
                    'Let's get the description for the Campaign
                    Try
                        oCMRsp = myCMService.GetCampaign(Token, nCMID)
                        sCMDesc = oCMRsp.name.ToString
                    Catch ex As Exception
                        Log("ASK ERROR ON Campaign Desc " & client.LogName(True) & ex.Message)
                        sCMDesc = nCMID.ToString
                    End Try
                    htCMDesc.Add(nCMID, sCMDesc)
                End If
                sLine.Append(sCMDesc)
                sLine.Append(",")

                'Do we have the AdGroup Description already; if not get it
                nAdID = oRsp(nCount).ad_id
                If htAdDesc.ContainsKey(nAdID) Then
                    sAdDesc = htAdDesc(nAdID).ToString
                Else
                    'Let's get the description for the AdGroup
                    Try
                        oAdRsp = myAdService.GetAd(Token, nAdID)
                        sAdDesc = oAdRsp.title.ToString
                    Catch ex As Exception
                        Log("ASK ERROR ON AdGroup Desc " & client.LogName(True) & ex.Message)
                        sAdDesc = nCMID.ToString
                    End Try
                    htAdDesc.Add(nAdID, sAdDesc)
                End If
                sLine.Append(sAdDesc)
                sLine.Append(",")

                sVar = oRsp(nCount).keyword.ToString
                sLine.Append(sVar)
                sLine.Append(",")

                sVar = oRsp(nCount).clicks.ToString
                sLine.Append(sVar)
                sLine.Append(",")

                sVar = oRsp(nCount).cost.ToString
                sLine.Append(sVar)
                sLine.Append(",")

                sVar = oRsp(nCount).ctr.ToString
                sLine.Append(sVar)
                sLine.Append(",")

                sVar = oRsp(nCount).average_cpc.ToString
                sLine.Append(sVar)
                sLine.Append(",")

                sVar = oRsp(nCount).average_position.ToString
                sLine.Append(sVar)

                sLine.AppendLine()
                objWriter.Write(sLine)
                sLine.Length = 0

            Next

            objWriter.Close()

            Return myID.ToString
        End Function
        '=============================================================================================
        'SetClient(ByVal sEMail As String, ByVal sPassword As String) called by ReqestReportID()
        'Instantiate and populate properties for myService via NewService():
        'myService = New AskWebServices.ReportService 
        'Used to invoke web services from ASK API
        '=============================================================================================
        Public Sub SetClient(ByVal sEMail As String, ByVal sPassword As String)
            Email = sEMail

            If Left(sPassword, 1) = "&" Then sPassword = ""
            If IsNothing(sPassword) Then sPassword = ""

            Password = idsDefault_Password
            If sPassword.Length > 0 Then Password = sPassword

            'Instantiate myService() = New ASKWebServices.ReportsService
            NewService()
            'Set timeout per documentation
            myService.Timeout = Integer.MaxValue
            myCMService.Timeout = Integer.MaxValue
            myAdService.Timeout = Integer.MaxValue

        End Sub
        '=============================================================================================
        'BldFileName(ByVal client As Clients) called by ASKReport.ReqestReportID 
        'Return String with full name of file.
        'Check for existing downloaded file in the local directory structure via naming conventions
        'AD_TheMetropolitanMuseumofArt_2008_03_03.csv for a daily report or
        'AM_TheMetropolitanMuseumofArt_2008_03_03.csv for a report spanning multiple days
        '=============================================================================================
        Public Function BldFileName(ByVal Client As Clients) As String

            'Perform Directory work
            If Not Directory.Exists(SavePath) Then Directory.CreateDirectory(SavePath)
            Dim myDir As String = Me.GetMyDir(Client)
            If Not Directory.Exists(myDir) Then Directory.CreateDirectory(myDir)
            'Perform File naming
            Dim sFileName As String = ""
            Dim MyFile As String = ""
            If Client.ASK.MultiDayReport Then
                MyFile = "AM_" & Replace(Client.Name, " ", "") & Client.ASK.EndDate.ToString("_yyyy_MM_dd") & ".csv"
            Else
                MyFile = "AD_" & Replace(Client.Name, " ", "") & Client.ASK.EndDate.ToString("_yyyy_MM_dd") & ".csv"
            End If
            sFileName = myDir & MyFile

            Return sFileName
        End Function
        '=============================================================================================
        'DoesFileExist(ByVal client As Clients) called by frmMain.vb Start_ASK()
        'Return Boolean
        'Check for existing downloaded file in the local directory structure via naming conventions
        'AD_TheMetropolitanMuseumofArt_2008_03_03.csv for a daily report or
        'AM_TheMetropolitanMuseumofArt_2008_03_03.csv for a report spanning multiple days
        '=============================================================================================
        Public Function DoesFileExist(ByVal Client As Clients) As Boolean

            If Not Directory.Exists(SavePath) Then Directory.CreateDirectory(SavePath)
            Dim myDir As String = Me.GetMyDir(Client)
            If Not Directory.Exists(myDir) Then Directory.CreateDirectory(myDir)

            Dim sFileName As String = ""
            Dim MyFile As String = ""
            If Client.ASK.MultiDayReport Then
                MyFile = "AM_" & Replace(Client.Name, " ", "") & Client.ASK.EndDate.ToString("_yyyy_MM_dd") & ".csv"
            Else
                MyFile = "AD_" & Replace(Client.Name, " ", "") & Client.ASK.EndDate.ToString("_yyyy_MM_dd") & ".csv"
            End If
            sFileName = myDir & MyFile

            Return File.Exists(sFileName)
        End Function
 

        '=============================================================================================
        'SaveFile(ByVal sUrl As String, ByVal sFileSave As String) called by DownloadReport()
        'Return Boolean
        'Using System.Net.WebClient DownloadFile(sUrl, sFileSave) method
        '=============================================================================================
        Public Function SaveFile(ByVal sUrl As String, ByVal sFileSave As String) As Boolean
            If File.Exists(sFileSave) Then File.Delete(sFileSave)

            Dim client As New WebClient
            client.DownloadFile(sUrl, sFileSave)

            Return File.Exists(sFileSave)
        End Function
 

        '======================================================================================
        ' Changed 3/14/07 so directories are now RecID's instead of names
        '======================================================================================
        Private Function GetMyDir(ByVal Client As Clients) As String

            Dim sDir As String
            If DirectoryNamesAsRecID Then
                sDir = SavePath & Client.CustID & "\"
            Else
                sDir = SavePath & Client.Name & "\"
            End If
            Return sDir
        End Function

#End Region

    End Class
End Namespace
