Imports System.IO
Imports System
Imports System.Net
'========================================================================================
'frmMain() Initialization for all variables/engines and data base access
'========================================================================================
Public Class frmMain
    Private Const IDS_Version As String = "PBJ: 1.201409.8 (03/04/15) YV7 GV201409 MSNV9 GAV3 SPV3 ASKV3 "

    Private Settings As New clSettings
    Private WithEvents EngineAuthentication As New ClEngineSecurity
    Private WithEvents EngineNames As New clEngineNames
    'GA
    Private WithEvents GAW As GAWebServices.GAReports
    Private WithEvents GAF As clGAFile
    'AdWords
    Private WithEvents GW As GoogleWebServices.GoogleReports
    Private WithEvents GF As clGoogleFile

    Private WithEvents MW As MSNWebServices.MSNReports
    Private WithEvents MF As clMSNFile

    'Private WithEvents VW As VerizonWebServices.VerizonReports
    'Private WithEvents VF As clVerizonFile

    'Private WithEvents YW As YahooWebServices.YahooReports
    'Private WithEvents YF As clYahooFile

    'Private WithEvents AW As ASKWebServices.ASKReports
    'Private WithEvents AF As clASKFile


    '-------------------------------------------------------------------
    ' nSteps Holder for Processing
    '-------------------------------------------------------------------
    Private nSteps As Integer = 99
    ' 00 = Idle
    ' 01 = Start 1
    ' 02 = Download 1
    ' 03 = Download Done 1
    ' 04 = Start 2
    ' 05 = Download 2
    ' 06 = Download Done 2
    ' 07 = Start 3
    ' 08 = Download 3
    ' 09 = Download Done 3
    ' 10 = Process
    ' 99 = All Done

    '-------------------------------------------------------------------
    'Booleans for What to Process
    '-------------------------------------------------------------------
    Private bDoMSNEOM As Boolean = False
    Private bDoMSNFlexEOM As Boolean = False
    Private bDoGoogleEOM As Boolean = False
    Private bDoGoogleFlexEOM As Boolean = False
    Private bDoGoogleFlexMM As Boolean = False
    Private bDoGoogle As Boolean = False
    Private bDoYahoo As Boolean = False
    Private bDoMSN As Boolean = False
    Private bDoVerizon As Boolean = False
    Private bDoASK As Boolean = False
    Private bDoGA As Boolean = False

    Private bDoDownload As Boolean = False
    Private bDoProcess As Boolean = False

    Private bFormLoading As Boolean = False
    Private nSleepSeconds As Integer = 5
    Private nMinuteCount As Integer = 0
    '-------------------------------------------------------------------

    Private bWorking As Boolean = True
    Private bAuto As Boolean = False
    Private sAutoCommand As String = ""

    Private bYahoo_CM As Boolean = False
    Private LastWS As WebService = WebService.All

    Private nSleepingCount As Integer = 0


#Region " Form Load and Auto Process Stuff "

    Private Sub frmMain_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click

    End Sub
    '========================================================================================
    'frmMain_Load() Initialization for all variables/engines and data base access
    '========================================================================================
    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Process settings.xml creating Settings Data Structure of the common variables 
        Settings = New clSettings
        SaveLoc = Settings.SaveBaseLocation

        Me.ClientList.CS = Settings.ConnectionString
        Me.ClientList.ReadClientsFromXML = Settings.ReadClientsFromXML

        'Set the Flags to Idle
        bAuto = False
        bFormLoading = False
        bDoProcess = False
        bDoDownload = False
        bDoGoogle = False
        bDoYahoo = False
        bDoMSN = False
        bDoVerizon = False
        bDoASK = False
        bDoGA = False
        bDoGoogleEOM = False
        bDoMSNEOM = False
        bDoMSNFlexEOM = False
        bDoGoogleFlexMM = False
        bDoGoogleFlexEOM = False
        bYahoo_CM = False
        nSteps = 99

        'Build Engines Combo Box 
        Me.SetWebServiceCombo()
        Me.Text = "Report Interface - " & IDS_Version

        'Condition Special Functions/Buttons via settings.xml
        Me.Button1.Visible = Settings.FixButton
        Me.Button2.Visible = Settings.FixButton
        Me.Button3.Visible = Settings.FixButton

        MyStatus("")
        Line("")

        'Instantiate the different Engines DataBase objects and Web Services API

        'GA Automation
        GAF = New clGAFile(Settings)
        GAW = New GAWebServices.GAReports(EngineAuthentication)
        GAW.BaseFilePath = Settings.SaveBaseLocation
        GAW.DirectoryNamesAsRecID = Settings.DirectoryNamesAsRecID

        'AdWords
        GF = New clGoogleFile(Settings)
        GW = New GoogleWebServices.GoogleReports(EngineAuthentication)
        GW.BaseFilePath = Settings.SaveBaseLocation
        GW.DirectoryNamesAsRecID = Settings.DirectoryNamesAsRecID

        MF = New clMSNFile(Settings)
        MW = New MSNWebServices.MSNReports(EngineAuthentication)
        MW.BaseFilePath = Settings.SaveBaseLocation
        MW.DirectoryNamesAsRecID = Settings.DirectoryNamesAsRecID

        'VF = New clVerizonFile(Settings)
        'VW = New VerizonWebServices.VerizonReports(EngineAuthentication)
        'VW.BaseFilePath = Settings.SaveBaseLocation
        'VW.DirectoryNamesAsRecID = Settings.DirectoryNamesAsRecID

        'YF = New clYahooFile(Settings)
        'YW = New YahooWebServices.YahooReports(EngineAuthentication)
        'YW.BaseFilePath = Settings.SaveBaseLocation
        'YW.DirectoryNamesAsRecID = Settings.DirectoryNamesAsRecID

        'AF = New clASKFile(Settings)
        'AW = New ASKWebServices.ASKReports(EngineAuthentication)
        'AW.BaseFilePath = Settings.SaveBaseLocation
        'AW.DirectoryNamesAsRecID = Settings.DirectoryNamesAsRecID

        'Build ListView (LV) 
        Me.ClientList.SetNew()

        'Set the Flags to Idle
        bAuto = False
        bFormLoading = False
        bDoProcess = False
        bDoDownload = False
        bDoGoogle = False
        bDoYahoo = False
        bDoMSN = False
        bDoVerizon = False
        bDoASK = False
        bYahoo_CM = False
        bDoGA = False
        bDoMSNEOM = False
        bDoMSNFlexEOM = False
        bDoGoogleEOM = False
        bDoGoogleFlexMM = False
        bDoGoogleFlexEOM = False
        nSteps = 99

        'Run from Scheduler will include parameter for the engine to run or "/ALL"
        If My.Application.CommandLineArgs.Count > 0 Then
            Dim nCount As Integer
            Dim arg As String
            For nCount = 0 To My.Application.CommandLineArgs.Count - 1

                arg = My.Application.CommandLineArgs(nCount)
                Select Case arg.ToUpper
                    Case "/ALL"
                        bDoGoogle = True
                        bDoMSN = True
                        bDoVerizon = True
                        bDoYahoo = True
                        bDoASK = True
                        bDoDownload = True
                        bDoProcess = True
                        bAuto = True
                        bFormLoading = True

                    Case "/GOOGLE"
                        bDoGoogle = True
                        bDoDownload = True
                        bDoProcess = True
                        bAuto = True
                        bFormLoading = True

                    Case "/YAHOO"
                        bDoYahoo = True
                        bDoDownload = True
                        bDoProcess = True
                        bAuto = True
                        bFormLoading = True

                    Case "/VERIZON"
                        bDoVerizon = True
                        bDoDownload = True
                        bDoProcess = True
                        bAuto = True
                        bFormLoading = True

                    Case "/MSN"
                        bDoMSN = True
                        bDoDownload = True
                        bDoProcess = True
                        bAuto = True
                        bFormLoading = True

                    Case "/ASK"
                        bDoASK = True
                        bDoDownload = True
                        bDoProcess = True
                        bAuto = True
                        bFormLoading = True

                    Case "/GA"
                        bDoGA = True
                        bDoDownload = True
                        bDoProcess = True
                        bAuto = True
                        bFormLoading = True

                    Case "/MSNEOM"
                        bDoMSNEOM = True
                        bDoMSN = True
                        bDoDownload = True
                        bDoProcess = True
                        bAuto = True
                        bFormLoading = True

                    Case "/GOOGLEEOM"
                        bDoGoogleEOM = True
                        bDoGoogle = True
                        bDoDownload = True
                        bDoProcess = True
                        bAuto = True
                        bFormLoading = True
                        bDoGA = True

                        'Case "/GOOGLEFLEXEOM"
                        '    bDoGoogleFlexEOM = True
                        '    bDoGoogle = False
                        '    bDoDownload = True
                        '    bDoProcess = False
                        '    bAuto = True
                        '    bFormLoading = True

                        'Case "/GOOGLEFLEXMM"
                        '    bDoGoogleFlexMM = True
                        '    bDoGoogle = False
                        '    bDoDownload = True
                        '    bDoProcess = True
                        '    bAuto = True
                        '    bFormLoading = True

                        'Case "/MSNFLEXEOM"
                        '    bDoMSNFlexEOM = True
                        '    bDoMSN = False
                        '    bDoDownload = True
                        '    bDoProcess = False
                        '    bAuto = True
                        '    bFormLoading = True
                End Select
            Next
        Else
            SetWebServices(Me.cmbWebService.SelectedItem)
        End If

        If bAuto = True Then
            Me.Enabled = False
            UpdateSleep(2)
            Me.Timer1.Interval = 1000
            Me.Timer1.Start()
        End If

    End Sub
    '========================================================================================
    'SetWebServiceCombo() Build Combo Box with all available WebServices                     
    '========================================================================================
    Private Sub SetWebServiceCombo()
        Me.cmbWebService.Items.Clear()
        Me.cmbWebService.Items.Add(WebService.All)
        Me.cmbWebService.Items.Add(WebService.Google)
        Me.cmbWebService.Items.Add(WebService.MSN)
        Me.cmbWebService.Items.Add(WebService.Verizon)
        Me.cmbWebService.Items.Add(WebService.Yahoo)
        Me.cmbWebService.Items.Add(WebService.ASK)
        Me.cmbWebService.Items.Add(WebService.GA)
        'Me.cmbFilter.Items.Add(New clMyItem("None (All Clients)", FilterList.All_Clients))
        Me.cmbWebService.SelectedIndex = 0
    End Sub
    '========================================================================================
    'RunAuto() Auto-run implies previous date and all engines; unless one of the following:  
    '  1. Google Analytics (Date - 2 data)                                                   
    '  2. Google EOM       (Previous months data)                                            
    '  3. Google Flex EOM  (Previous months data)                                            
    '  3. MSN Flex EOM  (Previous months data)                                            
    '========================================================================================
    Private Sub RunAuto()
        bFormLoading = False

        'Logic to get previous months start and end days
        Dim d, dteEOM, dteStartofMonth As Date
        d = Today.AddMonths(-1)
        dteEOM = d.AddDays(DateTime.DaysInMonth(d.Year, d.Month) - d.Day)

        dteStartofMonth = d.AddDays(-(d.Day - 1))
        If bDoGoogleEOM = True Then
            Me.dtStart.Value = dteStartofMonth
            Me.dtEnd.Value = dteEOM
        ElseIf bDoMSNEOM = True Then
            Me.dtStart.Value = dteStartofMonth
            Me.dtEnd.Value = dteEOM
        ElseIf bDoGoogleFlexMM = True Then
            Me.dtStart.Value = DateAdd(DateInterval.Day, -2, Today)
            Me.dtEnd.Value = DateAdd(DateInterval.Day, -2, Today)
        ElseIf bDoGoogleFlexEOM = True Then
            Me.dtStart.Value = dteStartofMonth
            Me.dtEnd.Value = dteEOM
        ElseIf bDoMSNFlexEOM = True Then
            Me.dtStart.Value = dteStartofMonth
            Me.dtEnd.Value = dteEOM
            'No Longer Look back 2 days, since at EOM repulling all 
            'ElseIf bDoGA = True Then
            '    Me.dtStart.Value = DateAdd(DateInterval.Day, -2, Today)
            '    Me.dtEnd.Value = DateAdd(DateInterval.Day, -2, Today)
        Else
            Me.dtStart.Value = DateAdd(DateInterval.Day, -1, Today)
            Me.dtEnd.Value = DateAdd(DateInterval.Day, -1, Today)
        End If

        Application.DoEvents()

        nSteps = 0
        Execute()
    End Sub
#End Region 'Form Load and Auto Process Stuff

#Region " Execution Control Stuff"
    '========================================================================================
    'Execute() Main line for processing clients and engines                                  
    'nSteps will correspond to the next phase to be run                                      
    'nSteps = 01 - Start_ALL()                                                               
    'nSteps = 02 - Download                                                                  
    '========================================================================================
    Private Sub Execute()

        '00 - Idle     
        If nSteps = 0 Then
            Running(True)
            StartLog(Me.StartDate, Me.EndDate, WebService.All, "Download")
            Me.nSleepSeconds = 0 'Start at 0 - we keep adding 10
            Me.nMinuteCount = 0

            If bDoDownload = True Then
                nSteps = 1      'Start the Download 
            ElseIf bDoProcess = True Then
                nSteps = 50     'Start Processing
            Else
                nSteps = 99     'We have nothong to do
            End If
        End If

        '--------------------------------------------------------------------------------------------
        ' First Download Attempt                                                                     
        '--------------------------------------------------------------------------------------------
        '01 - (01) Start
        If nSteps = 1 Then
            Start_ALL()
            nSteps = nSteps + 1 'We Started - Move On
        End If

        '02 - (02) Download
        If nSteps = 2 Then
            If Me.Loop_DownloadProcess = True Then
                nSteps = 7 ' If True All Download is Done
            Else
                nSteps = nSteps + 1
                Me.UpdateSleep(nSleepSeconds)
                Timer1.Start()
                Return
            End If

        End If

        '--------------------------------------------------------------------------------------------
        ' Second Download Attempt
        '--------------------------------------------------------------------------------------------
        '03 - (03) Start
        If nSteps = 3 Then
            Start_ALL()
            nSteps = nSteps + 1 'We Started - Move On
        End If

        '04 - (04) Download
        If nSteps = 4 Then
            If Me.Loop_DownloadProcess = True Then
                nSteps = 7 ' If True All Download is Done
            Else
                nSteps = nSteps + 1
                Me.UpdateSleep(nSleepSeconds)
                Timer1.Start()
                Return
            End If
        End If

        '--------------------------------------------------------------------------------------------
        ' Third Download Attempt
        '--------------------------------------------------------------------------------------------
        '05 - (05) Start
        If nSteps = 5 Then
            Start_ALL()
            nSteps = nSteps + 1 'We Started - Move On
        End If

        '06 - (06) Download
        If nSteps = 6 Then
            If Me.Loop_DownloadProcess = True Then
                nSteps = 7 ' If True All Download is Done
            Else
                'nSteps = nSteps + 1
                Me.UpdateSleep(nSleepSeconds)
                Timer1.Start()
                Return
            End If
        End If

        '--------------------------------------------------------------------------------------------
        ' Done with downloads - process if selected
        '--------------------------------------------------------------------------------------------
        '07 - Download done
        If nSteps = 7 Then
            If bDoProcess = True Then
                nSteps = 50     'Go Ahead and Process
            Else
                nSteps = 99     'We are Done Here
            End If
        End If

        '---------------------------------------------------------------------------------------------
        ' Process the files we have                                                                   
        ' Only perform UpdateDB_GA() for bDoGoogleEOM = True                                          
        ' Since this EOM job runs last in the series of EOM jobs                                      
        '---------------------------------------------------------------------------------------------
        '50 - Process
        If nSteps = 50 Then
            If bDoProcess Then
                If Me.bDoGoogle Then Me.UpdateDB_Google()
                If Me.bDoMSN Then Me.UpdateDB_MSN()
                'If Me.bDoVerizon Then Me.UpdateDB_Verizon()
                'If Me.bDoYahoo Then Me.UpdateDB_Yahoo()
                'If Me.bDoASK Then Me.UpdateDB_ASK()
                If Me.bDoGA Then Me.UpdateDB_GA()
                If Me.bDoGoogleEOM Then Me.UpdateDB_GA()
                If Me.bDoMSNEOM Then Me.UpdateDB_GA()
            End If
            nSteps = 99
        End If

        '--------------------------------------------------------------------------------------------
        ' We are done -- House cleaning
        '--------------------------------------------------------------------------------------------
        If nSteps = 99 Then
            'Send Emails if auto Process and We tried to do some downloads
            If Me.bAuto = True Then
                If bDoDownload = True Then SendEmails()
            End If
            'Close MSN
            'If bDoMSN = True Then
            '    Try
            '        Me.MW.CloseMSNService()
            '    Catch ex As Exception
            '    End Try
            'End If

            'Reset the Functrions
            bDoDownload = False
            bDoProcess = False

            If bYahoo_CM = True Then Me.cmbWebService.SelectedItem = Me.LastWS
            bYahoo_CM = False

            MyStatus("")
            EndLog()

            Running(False)
            If Me.bAuto = True Then
                'Close(MSN)
                If bDoMSN = True Then
                    'Try
                    '    Me.MW.CloseMSNService()
                    'Catch ex As Exception
                    'End Try
                End If
                Me.Close()

            End If
            'Reset the flags 
            Me.ClientList.SetNew()
            Exit Sub
        End If

        Me.UpdateSleep(nSleepSeconds)
        Timer1.Start()
    End Sub
    '========================================================================================
    'Loop_DownloadProcess() Get report url and download                                      
    '========================================================================================
    Private Function Loop_DownloadProcess() As Boolean
        Dim AllDone As Boolean = True
        Dim Downloaded As Integer = 0

        'Do the Download Stuff
        If Me.bDoDownload = True Then
            'Check and do all once
            If bDoGoogle = True Then
                Downloaded = Downloaded + Loop_Download_Google()
                If Me.ClientList.IsAllDownloaded_Google = False Then AllDone = False
            End If

            '    If bDoYahoo = True Then
            '        Downloaded = Downloaded + Loop_Download_Yahoo()
            '        If Me.ClientList.IsAllDownloaded_Yahoo = False Then AllDone = False
            '    End If

            If bDoMSN = True Then
                Downloaded = Downloaded + Loop_Download_MSN()
                If Me.ClientList.IsAllDownloaded_MSN = False Then AllDone = False
            End If

            '    If bDoVerizon = True Then
            '        Downloaded = Downloaded + Loop_Download_Verizon()
            '        If Me.ClientList.IsAllDownloaded_Verizon = False Then AllDone = False
            '    End If

            '    If bDoASK = True Then
            '        Downloaded = Downloaded + Loop_Download_ASK()
            '        If Me.ClientList.IsAllDownloaded_ASK = False Then AllDone = False
            '    End If
        End If

        'Adjust the Sleep time so we do not waste money (google)
        If AllDone = True Then
            'Put back to 0 seconds
            Me.nSleepSeconds = 0
            Me.nMinuteCount = 0
        Else
            'If Downloaded = 0 Then
            'Add 5 seconds to sleep time (max 60)
            Me.nSleepSeconds = Me.nSleepSeconds + 20
            If Me.nSleepSeconds > 60 Then
                nMinuteCount = nMinuteCount + 1
                Me.nSleepSeconds = 60

                'Halt Processes if we max out
                If nMinuteCount = 5 Then
                    '**************************** Put code here to hard error out what is not done ***********************************
                    Dim nCount As Integer
                    For nCount = 0 To Me.ClientList.Clients.GetUpperBound(0)
                        Dim Client As Clients = Me.ClientList.Clients(nCount)

                        If bDoGoogle Then
                            If Client.Google.Has_AccountID Then
                                Client.Google.Downloaded = True
                            End If
                        End If

                        If bDoMSN Then
                            If Client.MSN.Has_AccountID Then
                                Client.MSN.Downloaded = True
                            End If
                        End If

                        If bDoVerizon Then
                            If Client.Verizon.Has_AccountID Then
                                Client.Verizon.Downloaded = True
                            End If
                        End If

                        If bDoYahoo Then
                            If Client.Yahoo.Has_AccountID Then
                                Client.Yahoo.Key_Downloaded = True
                                Client.Yahoo.CM_Downloaded = True
                            End If
                        End If

                        If bDoASK Then
                            If Client.ASK.Has_AccountID Then
                                Client.ASK.Downloaded = True
                            End If
                        End If

                    Next
                End If
            End If
            'Else
            ''Put back to 0 seconds
            'Me.nSleepSeconds = 0
            'Me.nMinuteCount = 0
            'End If
        End If

        Return AllDone
    End Function
    '========================================================================================
    'Loop_Download_Google() Get Google report url and download                               
    '========================================================================================
    Private Function Loop_Download_Google() As Integer
        Dim Downloaded As Integer = 0

        If bDoGoogle = True Then
            If Me.ClientList.IsAllDownloaded_Google = False Then

                Downloaded = Me.Download_Google_Once()
            End If
        End If

        Return Downloaded
    End Function
    ''========================================================================================
    ''Loop_Download_Yahoo() Get Yahoo report url and download                                 
    ''========================================================================================
    'Private Function Loop_Download_Yahoo() As Integer
    '    Dim Downloaded As Integer = 0

    '    If bDoYahoo = True Then
    '        If Me.ClientList.IsAllDownloaded_Yahoo = False Then
    '            GetStatus_Yahoo()
    '            Downloaded = Me.Download_Yahoo_Once()
    '        End If
    '    End If

    '    Return Downloaded
    'End Function
    ''========================================================================================
    ''Loop_Download_MSN() Get MSN report url and download                                     
    ''========================================================================================
    Private Function Loop_Download_MSN() As Integer
        Dim Downloaded As Integer = 0

        If bDoMSN = True Then
            If Me.ClientList.IsAllDownloaded_MSN = False Then
                GetStatus_MSN()
                Downloaded = Me.Download_MSN_Once()
            End If
        End If

        Return Downloaded
    End Function
    ''========================================================================================
    ''Loop_Download_Verizon() Get Verizon report url and download                             
    ''========================================================================================
    'Private Function Loop_Download_Verizon() As Integer
    '    Dim Downloaded As Integer = 0

    '    If bDoVerizon = True Then
    '        If Me.ClientList.IsAllDownloaded_Verizon = False Then
    '            GetURL_Verizon()
    '            Downloaded = Me.Download_Verizon_Once()
    '        End If
    '    End If

    '    Return Downloaded
    'End Function
    ''========================================================================================
    ''Loop_Download_ASK() Get ASK report url and download                                     
    ''========================================================================================
    'Private Function Loop_Download_ASK() As Integer
    '    Dim Downloaded As Integer = 0

    '    If bDoASK = True Then
    '        If Me.ClientList.IsAllDownloaded_ASK = False Then
    '            Downloaded = 1
    '        End If
    '    End If

    '    Return Downloaded
    'End Function
    '========================================================================================
    'SendEmails() Report results of this run                                                 
    '========================================================================================
    Private Sub SendEmails()
        'Figure Email Data
        Dim nCount As Integer
        Dim bGoodDownload As Boolean = True
        Dim sMailBody As String = "Download Results" & vbCrLf & vbCrLf

        If bDoGoogle Then sMailBody = sMailBody & "GOOGLE was processed in this run." & vbCrLf
        If bDoMSN Then sMailBody = sMailBody & "MSN was processed in this run." & vbCrLf
        If bDoYahoo Then sMailBody = sMailBody & "YAHOO was processed in this run." & vbCrLf
        If bDoVerizon Then sMailBody = sMailBody & "VERIZON was processed in this run." & vbCrLf
        If bDoASK Then sMailBody = sMailBody & "ASK was processed in this run." & vbCrLf
        If bDoGA Then sMailBody = sMailBody & "GA Revenue was processed in this run." & vbCrLf
        If bDoGoogleEOM Then sMailBody = sMailBody & "GOOGLE EOM was processed in this run." & vbCrLf
        If bDoGoogleFlexEOM Then sMailBody = sMailBody & "GOOGLE FLEX EOM was processed in this run." & vbCrLf
        If bDoMSNEOM Then sMailBody = sMailBody & "MSN EOM was processed in this run." & vbCrLf

        If bDoGoogleFlexEOM Then
            Dim Msg() As String = Me.CheckForProcessed_Google(True)
            If Msg.GetUpperBound(0) = 0 Then
                sMailBody = sMailBody & vbCrLf & vbCrLf & "Google Flex Good - None Missing"
            Else
                bGoodDownload = False
                sMailBody = sMailBody & vbCrLf & vbCrLf
                For nCount = 0 To Msg.GetUpperBound(0)
                    sMailBody = sMailBody & Msg(nCount) & vbCrLf
                Next
            End If
        End If

        If bDoGoogle Then
            Dim Msg() As String = Me.CheckForProcessed_Google(bDoProcess)
            If Msg.GetUpperBound(0) = 0 Then
                sMailBody = sMailBody & vbCrLf & vbCrLf & "Google Good - None Missing"
            Else
                bGoodDownload = False
                sMailBody = sMailBody & vbCrLf & vbCrLf
                For nCount = 0 To Msg.GetUpperBound(0)
                    sMailBody = sMailBody & Msg(nCount) & vbCrLf
                Next
            End If
        End If

        If bDoMSN Then
            Dim Msg() As String = Me.CheckForProcessed_MSN(bDoProcess)
            If Msg.GetUpperBound(0) = 0 Then
                sMailBody = sMailBody & vbCrLf & vbCrLf & "MSN Good - None Missing"
            Else
                bGoodDownload = False
                sMailBody = sMailBody & vbCrLf & vbCrLf
                For nCount = 0 To Msg.GetUpperBound(0)
                    sMailBody = sMailBody & Msg(nCount) & vbCrLf
                Next
            End If
        End If

        'If bDoYahoo Then
        '    Dim Msg() As String = Me.CheckForProcessed_Yahoo(bDoProcess)
        '    If Msg.GetUpperBound(0) = 0 Then
        '        sMailBody = sMailBody & vbCrLf & vbCrLf & "Yahoo Good - None Missing"
        '    Else
        '        bGoodDownload = False
        '        sMailBody = sMailBody & vbCrLf & vbCrLf
        '        For nCount = 0 To Msg.GetUpperBound(0)
        '            sMailBody = sMailBody & Msg(nCount) & vbCrLf
        '        Next
        '    End If
        'End If

        'If bDoVerizon Then
        '    Dim Msg() As String = Me.CheckForProcessed_Verizon(bDoProcess)
        '    If Msg.GetUpperBound(0) = 0 Then
        '        sMailBody = sMailBody & vbCrLf & vbCrLf & "Verizon Good - None Missing"
        '    Else
        '        bGoodDownload = False
        '        sMailBody = sMailBody & vbCrLf & vbCrLf
        '        For nCount = 0 To Msg.GetUpperBound(0)
        '            sMailBody = sMailBody & Msg(nCount) & vbCrLf
        '        Next
        '    End If
        'End If

        'If bDoASK Then
        '    Dim Msg() As String = Me.CheckForProcessed_ASK(bDoProcess)
        '    If Msg.GetUpperBound(0) = 0 Then
        '        sMailBody = sMailBody & vbCrLf & vbCrLf & "ASK Good - None Missing"
        '    Else
        '        bGoodDownload = False
        '        sMailBody = sMailBody & vbCrLf & vbCrLf
        '        For nCount = 0 To Msg.GetUpperBound(0)
        '            sMailBody = sMailBody & Msg(nCount) & vbCrLf
        '        Next
        '    End If
        'End If

        'Verizon NO Data Stuff
        If bDoVerizon Then
            sMailBody = sMailBody & vbCrLf & vbCrLf
            For Each client As Clients In Me.ClientList.Clients
                If client.Verizon.Has_AccountID Then
                    If client.Verizon.NoData Then
                        Dim sMyMsg As String = "Verizon Note: " & client.LogName(True) & "Report is empty; check Verizon portal for client and if traffic found contact IT and request re-pull."
                        sMailBody = sMailBody & sMyMsg & vbCrLf
                    End If
                    If client.Verizon.ErrorReport Then
                        Dim sMyMsg As String = "Verizon Note: " & client.LogName(True) & "Report error; check Verizon portal for client and if traffic found contact IT and request re-pull."
                        sMailBody = sMailBody & sMyMsg & vbCrLf
                    End If
                End If
            Next
        End If

        Try
            If Settings.Mail_From_Email <> "" Then
                Dim msgFROM As New Mail.MailAddress(Settings.Mail_From_Email, Settings.Mail_From_Name)
                Dim sTO_Email As String = "davidk@setacorporation.com"
                'Dim sTO_Email As String = "tim@alboc.com"
                If Settings.Mail_To_Email.GetUpperBound(0) > 0 Then
                    If Settings.Mail_To_Email(1) <> "" Then sTO_Email = Settings.Mail_To_Email(1)
                End If
                Dim msgTO As New Mail.MailAddress(sTO_Email)

                Dim message As New Mail.MailMessage(msgFROM, msgTO)

                If Settings.Mail_To_Email.GetUpperBound(0) > 1 Then
                    For nCount = 2 To Settings.Mail_To_Email.GetUpperBound(0)
                        If Settings.Mail_To_Email(nCount) <> "" Then
                            message.To.Add(Settings.Mail_To_Email(nCount))
                        End If
                    Next
                End If

                If Settings.Mail_CC_Email.GetUpperBound(0) > 0 Then
                    For nCount = 1 To Settings.Mail_CC_Email.GetUpperBound(0)
                        If Settings.Mail_CC_Email(nCount) <> "" Then
                            message.CC.Add(Settings.Mail_CC_Email(nCount))
                        End If
                    Next
                End If

                'message.Bcc.Add("tjc_info@earthlink.net")
                message.Bcc.Add("okiebug1399@gmail.com")
                message.BodyEncoding = System.Text.Encoding.UTF8
                message.Body = sMailBody

                If bGoodDownload Then
                    message.Subject = Now.ToString("MM/dd/yyyy") & " - Good Download"
                Else
                    message.Subject = Now.ToString("MM/dd/yyyy") & " - Download ERRORS - See Message"
                    If File.Exists(MyFile) Then
                        File.Copy(MyFile, ErrLogFile, True)
                        Dim LogAttachment As New Mail.Attachment(ErrLogFile)
                        message.Attachments.Add(LogAttachment)
                    Else
                        message.Subject = Now.ToString("MM/dd/yyyy") & " - Download ERRORS - See Message (No Log File)"
                    End If
                End If

                ' SmtpClient is used to send the e-mail
                Dim mailClient As New Mail.SmtpClient(Settings.Mail_Server)
                mailClient.UseDefaultCredentials = False
                mailClient.Send(message)

            End If
        Catch ex As Exception
            Log("Email Error - " & ex.Message)
            If Not IsNothing(ex.InnerException) Then Log("InnerMsg: " & ex.InnerException.Message)
        End Try
    End Sub

    
#End Region 'Execution Control Stuff

#Region " All Functions "

    Private Sub Start_ALL()
        'Google Work
        If bDoGoogleEOM = True Then Me.Start_GoogleEOM()
        If bDoGoogle = True Then Me.Start_Google()

        'MSN Work
        If bDoMSNEOM = True Then Me.Start_MSNEOM()
        If bDoMSN = True Then Me.Start_MSN()

        'GA Work
        If bDoGA = True Then Me.Start_GA()

        'If bDoMSNFlexEOM = True Then Me.Start_FlexMSNEOM()
        'If bDoGoogleFlexEOM = True Then Me.Start_FlexGoogleEOM()
        'If bDoYahoo = True Then Me.Start_Yahoo()
        'If bDoVerizon = True Then Me.Start_Verizon()
        'If bDoASK = True Then Me.Start_ASK()
    End Sub

    Private Sub ResetDB_ALL()
        Me.ResetDB_Google()
        'Me.ResetDB_MSN()
        'Me.ResetDB_Yahoo()
        'Me.ResetDB_ASK()
        'Me.ResetDB_Verizon()
    End Sub
#End Region 'All Functions

#Region " Google Functions "
    Private Sub Start_GoogleEOM()
        Log("Start Google EOM: " & Now.ToString("hh:mm:ss"))
        Dim sRowsDlt As String
        Dim msg As String = "Starting Google EOM Delete"
        MyStatus(msg & "...")
        Log(msg & ": " & Now.ToString("hh:mm:ss"))

        'Step 1: Update IsArbitrage Clients object Client.Google.AccountID="" to remove Google Account and prevent pulling month of data
        'Bypass IsArbitrage clients
        For Each Client As Clients In Me.ClientList.Clients
            If Client.Google.Has_AccountID And Client.IsArbitrage = 1 Then
                Client.Google.AccountID = ""
            End If
        Next

        'Step 2: Delete past months Google Perfdata
        'Bypass IsArbitrage clients
        For Each Client As Clients In Me.ClientList.Clients
            If Client.Google.Has_AccountID And Client.IsArbitrage <> 1 And Client.Google.EOMDeleted = False Then

                'Client.Google.ClearDownloadInfo()   *NOTE: Removed May 5 2011
                Client.Google.GoogleEOM = True
                Client.Google.StartDate = StartDate()
                Client.Google.EndDate = EndDate()
                'Log Requesting action
                Log(Client.LogName & ": Deleting PerfData for " & Client.Google.StartDate.ToShortDateString & " to " & Client.Google.EndDate.ToShortDateString)
                MyStatus(msg & " (" & Client.Name & ")...")
                sRowsDlt = GF.ExecuteEOM(Client)
                'Log Requesting action
                Log(Client.LogName & " " & sRowsDlt & " rows deleted.")
                Client.Google.EOMDeleted = True
                Application.DoEvents()
            End If
        Next

        MyStatus(msg & "...Done")
        Log(msg & "...Done " & Now.ToString("hh:mm:ss"))

        '-- Start Feb 3, 2015 MODS
        'No longer reset old GA download files, we will grab fresh from GA in Start_GA
        'msg = "Reseting file from GA "
        'MyStatus(msg & "...")
        'Log(msg & Now.ToString("hh:mm:ss"))

        ''Step 3: Reset GA files moving out of "Processed" folder
        ''Bypass IsArbitrage clients
        'For Each Client As Clients In Me.ClientList.Clients
        '    If Client.GA.Has_AccountID And Client.IsArbitrage <> 1 And Client.GA.EOMDeleted = False Then
        '        Client.GA.StartDate = StartDate()
        '        Client.GA.EndDate = EndDate()
        '        'Log Requesting action
        '        Log(Client.LogName & ": Reseting File for " & Client.GA.StartDate.ToShortDateString & " to " & Client.GA.EndDate.ToShortDateString)
        '        MyStatus(msg & " (" & Client.Name & ")...")
        '        'Will repull month of data 12/2/2014
        '        'GAF.ResetFiles(Client, Client.GA.StartDate, Client.GA.EndDate)
        '        Client.GA.EOMDeleted = True

        '        Application.DoEvents()
        '    End If
        'Next

        'MyStatus(msg & "...Done")
        'Log(msg & "...Done " & Now.ToString("hh:mm:ss"))
        '-- END Feb 3, 2015 MODS

    End Sub
    Private Sub Start_FlexGoogleEOM()
        Dim intTotClients As Integer = 0
        Dim msgDownloadCnt As String = ""
        Dim msg As String = "Starting FlexEOM Google "
        MyStatus(msg & "...")
        Log(msg & Now.ToString("hh:mm:ss"))
        bDoGoogle = True
        Dim ReportID As String
        'SQL needed for building GoogleReportID
        Dim oDB As New clDatabase
        oDB.cs = Settings.ConnectionString

        '
        'Step 1: Update Not IsArbitrage Clients object Client.Google.AccountID="" to remove Google Account 
        'and prevent pulling month of data in Loop_Download_Google()
        '
        For Each Client As Clients In Me.ClientList.Clients
            If Client.Google.Has_AccountID And Client.IsArbitrage <> 1 Then
                Client.Google.AccountID = ""
            End If
        Next
        '
        'Step 2: Request Report-Only process IsArbitrage Clients 
        '
        For Each Client As Clients In Me.ClientList.Clients
            If Client.Google.Has_AccountID And Client.IsArbitrage = 1 Then
                Client.Google.GoogleEOM = True
                If Client.Google.Has_EOMReportID Then
                    Client.Google.ClearDownloadNotReport()
                    Client.Google.StartDate = StartDate()
                    Client.Google.EndDate = EndDate()
                    ReportID = Client.Google.EOMReportID
                    Me.ClientList.Done_ReportID(Client.Index, WebService.Google, ReportID)

                Else
                    'Client.Google.ClearDownloadInfo()    *NOTE: Removed May 5 2011
                    Client.Google.StartDate = StartDate()
                    Client.Google.EndDate = EndDate()
                    Client.Google.GoogleEOM = True
                    'Log Requesting action
                    Log(Client.LogName & ": Requesting FlexEOM Google Report")
                    MyStatus(msg & " (" & Client.Name & ")...")
                    'Perform Report Build and Request
                    ReportID = GW.ReqestReports(Client)
                    Log(Client.LogName & ": GW ID: " & ReportID)
                    'If error in request for report; perform clean up
                    If ReportID = "0" Or ReportID = "ERROR" Then
                        Me.ClientList.Done_Failed(Client.Index, WebService.Google)
                    Else
                        Me.ClientList.Done_ReportID(Client.Index, WebService.Google, ReportID)
                        Me.ClientList.Done_Download(Client.Index, WebService.Google)
                        nSteps = 6
                    End If

                End If
                intTotClients += 1
                Application.DoEvents()
            End If
        Next
        MyStatus(msg & "...Done")
        Log(msg & "...Done " & Now.ToString("hh:mm:ss"))

        ''
        ''Step 3: Download Report via Loop_Download_Google() until ClientList.IsAllDownloaded_Google = True 
        ''
        'msg = "Downloading Reports FlexEOM Google Status "
        'MyStatus(msg & "...")
        'Log(msg & Now.ToString("hh:mm:ss"))
        'Dim intDownloadedTot As Integer = 0
        'Dim intDownloaded As Integer = 0
        'Dim AllDone As Boolean = True

        'Do While Me.ClientList.IsAllDownloaded_Google = False
        '    intDownloaded = Loop_Download_Google()
        '    intDownloadedTot += intDownloaded
        '    If Me.ClientList.IsAllDownloaded_Google = False Then AllDone = False
        '    msgDownloadCnt = "Downloaded " & CStr(intDownloaded)
        '    MyStatus(msgDownloadCnt & "...")
        '    'Adjust the Sleep time so we do not waste money (google)
        '    If AllDone = True Then
        '        'Put back to 0 seconds
        '        Me.nSleepSeconds = 0
        '        Me.nMinuteCount = 0
        '    Else
        '        If intDownloaded = 0 Then
        '            'Add 5 seconds to sleep time (max 60)
        '            Me.nSleepSeconds = Me.nSleepSeconds + 5
        '            If Me.nSleepSeconds > 60 Then
        '                nMinuteCount = nMinuteCount + 1
        '                Me.nSleepSeconds = 60

        '                'Halt Processes if we max out
        '                If nMinuteCount = 7 Then

        '                    'Not responding so flag as all done
        '                    For Each Client As Clients In Me.ClientList.Clients
        '                        If Client.Google.Has_AccountID And Client.IsArbitrage = 1 Then
        '                            If Client.Google.Has_AccountID Then
        '                                Client.Google.Downloaded = True
        '                            End If
        '                        End If

        '                    Next
        '                End If
        '            End If
        '        Else
        '            'Put back to 0 seconds
        '            Me.nSleepSeconds = 0
        '            Me.nMinuteCount = 0
        '        End If
        '    End If

        'Loop
        'msgDownloadCnt = "Downloaded " & CStr(intDownloadedTot)
        'MyStatus(msgDownloadCnt & "...")
        'MyStatus(msg & "...Done")
        'Log(msgDownloadCnt)
        'Log(msg & "...Done " & Now.ToString("hh:mm:ss"))

        '
        'Step : Update Database 
        '
        msg = "Update DB from Google FlexEOM"
        MyStatus(msg & "...")

        StartLog(Me.StartDate, Me.EndDate, WebService.Google, "Updating")

        For Each Client As Clients In Me.ClientList.Clients
            If Client.Google.Has_AccountID And Client.IsArbitrage = 1 Then
                MyStatus(msg & " (" & Client.Name & ")...")

                Log(Client.LogName & ": Updating FLEX Google")

                GF.Execute(Client, True)

                Log(Client.LogName & ": FLEX Google Updated")

                Me.ClientList.Done_MyProcess(Client.Index, WebService.Google)
                Application.DoEvents()
            End If
        Next

        EndLog()

        bDoGoogle = False
        MyStatus(msg & "...Done")
        Log(msg & "...Done " & Now.ToString("hh:mm:ss"))
        nSteps = 98
    End Sub
   
    Private Sub Start_Google()
        Dim msg As String = "Starting Google Reports"
        MyStatus(msg & "...")

        'V201109 now uses GoogleReports.vb>RequestReports(Client)
        For Each Client As Clients In Me.ClientList.Clients
            If Client.Google.Has_AccountID And Client.Google.Downloaded = False Then
                If (Client.Google.Has_ReportID = False) Then
                    'Client.Google.ClearDownloadInfo()   *NOTE: Removed May 5 2011
                    Client.Google.StartDate = StartDate()
                    Client.Google.EndDate = EndDate()
                    If GW.DoesFileExist(Client) = False Then
                        'Log Requesting action
                        Log(Client.LogName & ": Requesting Google Report")
                        MyStatus(msg & " (" & Client.Name & ")...")
                        'Perform Report Build, Request and download
                        Dim ReportID As String = GW.ReqestReports(Client)
                        Log(Client.LogName & ": GW ID: " & ReportID)
                        'If error in request for report; perform clean up
                        If ReportID = "0" Or ReportID = "ERROR" Then
                            Me.ClientList.Done_Failed(Client.Index, WebService.Google)
                        Else
                            Me.ClientList.Done_ReportID(Client.Index, WebService.Google, ReportID)
                            Me.ClientList.Done_Download(Client.Index, WebService.Google)
                            nSteps = 6
                        End If
                    Else
                        Me.ClientList.Done_Download(Client.Index, WebService.Google)
                    End If
                    Application.DoEvents()
                End If

            End If
        Next


        MyStatus(msg & "...Done")
    End Sub
    'Private Sub Start_GoogleV13()
    '    Dim msg As String = "Starting Google Reports"
    '    MyStatus(msg & "...")

    '    'Reverting back to V13 - Remarketing shit
    '    For Each Client As Clients In Me.ClientList.Clients
    '        If Client.Google.Has_AccountID And Client.Google.Downloaded = False Then
    '            If (Client.Google.Has_ReportID = False) Then
    '                'Client.Google.ClearDownloadInfo()   *NOTE: Removed May 5 2011
    '                Client.Google.StartDate = StartDate()
    '                Client.Google.EndDate = EndDate()
    '                If GW.DoesFileExist(Client) = False Then
    '                    'Log Requesting action
    '                    Log(Client.LogName & ": Requesting Google Report")
    '                    MyStatus(msg & " (" & Client.Name & ")...")
    '                    'Perform Report Build and Request
    '                    Dim ReportID As String = GW.ReqestReportID(Client)
    '                    Log(Client.LogName & ": Google ID: " & ReportID)
    '                    'If error in request for report; perform clean up
    '                    If ReportID = "ERROR" Then
    '                        Me.ClientList.Done_Failed(Client.Index, WebService.Google)
    '                    Else
    '                        Me.ClientList.Done_ReportID(Client.Index, WebService.Google, ReportID)
    '                    End If
    '                Else
    '                    Me.ClientList.Done_Download(Client.Index, WebService.Google)
    '                End If
    '                Application.DoEvents()
    '            End If

    '        End If
    '    Next


    '    'Remove for V13
    '    'Dim ReportID As String
    '    ''SQL needed for building GoogleReportID
    '    'Dim oDB As New clDatabase
    '    'oDB.cs = Settings.ConnectionString


    '    'For Each Client As Clients In Me.ClientList.Clients
    '    '    If Client.Google.Has_AccountID And Client.Google.Downloaded = False Then
    '    '        'V2010 added for report defintion id
    '    '        'NEW: GoogleReportID file stores existing report definition for yesterday's data pulls
    '    '        If Client.Google.Has_ReportID And ((StartDate() = EndDate()) And (StartDate() = DateAdd(DateInterval.Day, -1, Today))) Then
    '    '            Client.Google.ClearDownloadNotReport()
    '    '            Client.Google.StartDate = StartDate()
    '    '            Client.Google.EndDate = EndDate()
    '    '            ReportID = Client.Google.ReportID
    '    '            Me.ClientList.Done_ReportID(Client.Index, WebService.Google, ReportID)

    '    '            'V2010 added for report defintion id
    '    '            'NEW: GoogleReportID file stores existing report definition for EOM data pulls
    '    '        ElseIf Client.Google.Has_EOMReportID And bDoGoogleEOM Then
    '    '            Client.Google.ClearDownloadNotReport()
    '    '            Client.Google.StartDate = StartDate()
    '    '            Client.Google.EndDate = EndDate()
    '    '            ReportID = Client.Google.EOMReportID
    '    '            Me.ClientList.Done_ReportID(Client.Index, WebService.Google, ReportID)

    '    '        Else
    '    '            If (Client.Google.Has_ReportID = False) Then

    '    '                Client.Google.ClearDownloadInfo()
    '    '                Client.Google.StartDate = StartDate()
    '    '                Client.Google.EndDate = EndDate()
    '    '                If GW.DoesFileExist(Client) = False Then
    '    '                    'Log Requesting action
    '    '                    Log(Client.LogName & ": Requesting Google Report")
    '    '                    MyStatus(msg & " (" & Client.Name & ")...")
    '    '                    'Perform ReportDefinition Build and Log Report ID
    '    '                    ReportID = GW.ReqestReportID(Client)
    '    '                    Log(Client.LogName & ": Google ID: " & ReportID)
    '    '                    'If error in request for report; perform clean up
    '    '                    If ReportID = "ERROR" Then
    '    '                        Me.ClientList.Done_Failed(Client.Index, WebService.Google)
    '    '                    Else
    '    '                        Me.ClientList.Done_ReportID(Client.Index, WebService.Google, ReportID)
    '    '                        'Save to database for reuse if yesterday only
    '    '                        If Client.Google.GoogleEOM Or ((StartDate() = EndDate()) And (StartDate() = DateAdd(DateInterval.Day, -1, Today))) Then
    '    '                            'Instantiate and set fields for use in SQL DELETE FROM PERFDATA
    '    '                            oDB.AddorUpdateGoogleReport(Client.CustID, ReportID, Client.Google.GoogleEOM)
    '    '                        End If
    '    '                    End If
    '    '                Else
    '    '                    Me.ClientList.Done_Download(Client.Index, WebService.Google)
    '    '                End If
    '    '            End If

    '    '        End If
    '    '        Application.DoEvents()
    '    '    End If
    '    'Next

    '    MyStatus(msg & "...Done")
    'End Sub
    ''=====================================================================
    '' GetStatus_Google() Called from Loop_DownloadProcess                 
    '' until Client.Google.Downloaded = True for all clinets               
    '' Me.ClientList.IsAllDownloaded_Google = True                         
    ''=====================================================================
    'Private Sub GetStatus_Google()
    '    Dim msg As String = "Update Google Status"
    '    MyStatus(msg & "...")

    '    For Each Client As Clients In Me.ClientList.Clients
    '        If Client.Google.Has_AccountID Then
    '            If Client.Google.Downloaded = False Then
    '                If Client.Google.DownloadReady = False Then
    '                    MyStatus(msg & " (" & Client.Name & ")...")

    '                    Log(Client.LogName & ": Getting Google Status")
    '                    Dim Status As String = GW.GetReportStatus(Client)  'V2010 returns "READY"

    '                    Log(Client.LogName & ": Google Status: " & Status)

    '                    'Update Clients(Index).Google.DownloadReady = True
    '                    If Status.Length > 0 Then
    '                        Me.ClientList.Done_Status(Client.Index, WebService.Google, Status)
    '                    End If

    '                End If
    '            End If
    '        End If
    '    Next

    '    MyStatus(msg & "...Done")
    'End Sub
    '=====================================================================
    ' Download_Google_Once() Called from Loop_DownloadProcess             
    ' until Client.Google.Downloaded = True for all clinets               
    ' Me.ClientList.IsAllDownloaded_Google = True                         
    '=====================================================================
    Private Function Download_Google_Once() As Integer
        Dim msg As String = "Downloading Google Reports"
        MyStatus(msg & "...")
        Dim Downloaded As Integer = 0

        For Each Client As Clients In Me.ClientList.Clients
            If Client.Google.Has_AccountID Then
                If Client.Google.Downloaded = False Then
                    If Client.Google.DownloadReady Then
                        MyStatus(msg & " (" & Client.Name & ")...")

                        Log(Client.LogName & ": Downloading Google")
                        'Perform file download from URL returned via ID; and save in directory/file name
                        'Me.GW.DownloadReport(Client)

                        Log(Client.LogName & ": Downloaded")
                        'Set Clients(Index).Google.Downloaded = True
                        Me.ClientList.Done_Download(Client.Index, WebService.Google)
                        Downloaded = Downloaded + 1
                    End If
                End If
            End If
        Next

        MyStatus(msg & "...Done")
        Return Downloaded
    End Function

    Private Sub UpdateDB_Google()
        Dim msg As String = "Update DB from Google"
        Dim sResult As String = ""
        Dim bFlexEOM As Boolean = False
        MyStatus(msg & "...")

        StartLog(Me.StartDate, Me.EndDate, WebService.Google, "Updating")

        For Each Client As Clients In Me.ClientList.Clients
            If Client.Google.Has_AccountID Then
                MyStatus(msg & " (" & Client.Name & ")...")

                Log(Client.LogName & ": Updating Google")

                GF.Execute(Client, bFlexEOM)

                Log(Client.LogName & ": Google Updated")

                Me.ClientList.Done_MyProcess(Client.Index, WebService.Google)
                Application.DoEvents()
            End If
        Next

        MyStatus(msg & "...Done")

        'Added for Remarketing and "boomuserlist" keyword conversion via sp on SQL server
        Log("Stored Procedure for Google ReMarketing Started: " & Now.ToString("hh:mm:ss"))
        Dim msgStoredProc As String = "Running Stored Procedure BoomUserList"
        MyStatus(msgStoredProc)
        sResult = GF.UpdateRemarketing_Keywords(StartDate(), EndDate())
        msgStoredProc = sResult & " Records Updated"
        MyStatus(msgStoredProc)
        Log("Stored Procedure for Google ReMarketing Ended: " & Now.ToString("hh:mm:ss") & " Records Updated: " & sResult)

        If bDoGoogleEOM Then
            Log("End Google EOM: " & Now.ToString("hh:mm:ss"))
        End If

        EndLog()

    End Sub
   

    Private Sub ResetDB_Google()
        Dim msg As String = "Reseting DB from Google"
        MyStatus(msg & "...")

        For Each Client As Clients In Me.ClientList.Clients
            If Client.Google.Has_AccountID Then
                MyStatus(msg & " (" & Client.Name & ")...")

                GF.ResetFiles(Client, StartDate, EndDate)
                Application.DoEvents()
            End If
        Next

        MyStatus(msg & "...Done")
    End Sub
    ''Arbitrage Reset for Nov
    'Private Sub ResetEOMDB_MSN()
    '    Dim msg As String = "Reseting EOM DB from Flex Bid"
    '    MyStatus(msg & "...")

    '    For Each Client As Clients In Me.ClientList.Clients
    '        If Client.Google.Has_AccountID And Client.IsArbitrage = 1 Then
    '            'Client.Google.ClearDownloadInfo()   *NOTE: Removed May 5 2011
    '            Client.Google.StartDate = StartDate()
    '            Client.Google.EndDate = EndDate()
    '            MyStatus(msg & " (" & Client.Name & ")...")

    '            MF.ResetFilesEOM(Client, StartDate, EndDate)
    '            Application.DoEvents()
    '        End If
    '    Next

    '    MyStatus(msg & "...Done")
    'End Sub
    'Arbitrage Reset for Nov
    Private Sub ResetEOMDB_Google()
        Dim msg As String = "Reseting EOM DB for PPC"
        MyStatus(msg & "...")

        For Each Client As Clients In Me.ClientList.Clients
            If Client.Google.Has_AccountID And Client.IsArbitrage <> 1 Then
                'Client.Google.ClearDownloadInfo()  *NOTE: Removed May 5 2011
                Client.Google.StartDate = StartDate()
                Client.Google.EndDate = EndDate()
                MyStatus(msg & " (" & Client.Name & ")...")

                GF.ResetFilesEOM(Client, StartDate, EndDate)
                Application.DoEvents()
            End If
        Next

        MyStatus(msg & "...Done")
    End Sub
    ''Arbitrage Reset for Nov
    'Private Sub ResetEOMDB_Yahoo()
    '    Dim msg As String = "Reseting EOM DB from Flex Bid"
    '    MyStatus(msg & "...")

    '    For Each Client As Clients In Me.ClientList.Clients
    '        If Client.Yahoo.Has_AccountID And Client.IsArbitrage = 1 Then
    '            Client.Yahoo.ClearDownloadInfo()
    '            Client.Yahoo.StartDate = StartDate()
    '            Client.Yahoo.EndDate = EndDate()
    '            MyStatus(msg & " (" & Client.Name & ")...")

    '            YF.ResetFilesEOM(Client, StartDate, EndDate)
    '            Application.DoEvents()
    '        End If
    '    Next

    '    MyStatus(msg & "...Done")
    'End Sub
#End Region 'Google Functions

    '#Region " ASK Functions "

    '    Private Sub Start_ASK()
    '        Dim msg As String = "Starting ASK Reports"
    '        MyStatus(msg & "...")

    '        For Each Client As Clients In Me.ClientList.Clients
    '            If Client.ASK.Has_AccountID And Client.ASK.Downloaded = False Then
    '                Client.ASK.ClearDownloadInfo()
    '                Client.ASK.StartDate = StartDate()
    '                Client.ASK.EndDate = EndDate()
    '                If AW.DoesFileExist(Client) = False Then
    '                    'Log Requesting action
    '                    Log(Client.LogName & ": Requesting ASK Report")
    '                    MyStatus(msg & " (" & Client.Name & ")...")
    '                    'Perform Report Build and Request
    '                    Dim ReportID As String = AW.ReqestReportID(Client)
    '                    Log(Client.LogName & ": ASK ID: " & ReportID)
    '                    'If error in request for report; perform clean up
    '                    If ReportID = "0" Then
    '                        Me.ClientList.Done_Failed(Client.Index, WebService.ASK)
    '                    Else
    '                        Me.ClientList.Done_ReportID(Client.Index, WebService.ASK, ReportID)
    '                        Me.ClientList.Done_Download(Client.Index, WebService.ASK)
    '                        nSteps = 6
    '                    End If
    '                Else
    '                    Me.ClientList.Done_Download(Client.Index, WebService.ASK)
    '                End If
    '                Application.DoEvents()
    '            End If
    '        Next

    '        MyStatus(msg & "...Done")
    '    End Sub




    '    Private Sub UpdateDB_ASK()
    '        Dim msg As String = "Update DB from ASK"
    '        MyStatus(msg & "...")

    '        StartLog(Me.StartDate, Me.EndDate, WebService.ASK, "Updating")

    '        For Each Client As Clients In Me.ClientList.Clients
    '            If Client.ASK.Has_AccountID Then
    '                MyStatus(msg & " (" & Client.Name & ")...")

    '                Log(Client.LogName & ": Updating ASK")

    '                AF.Execute(Client)

    '                Log(Client.LogName & ": ASK Updated")

    '                Me.ClientList.Done_MyProcess(Client.Index, WebService.ASK)
    '                Application.DoEvents()
    '            End If
    '        Next

    '        EndLog()

    '        MyStatus(msg & "...Done")
    '    End Sub

    '    Private Sub ResetDB_ASK()
    '        Dim msg As String = "Reseting DB from ASK"
    '        MyStatus(msg & "...")

    '        For Each Client As Clients In Me.ClientList.Clients
    '            If Client.ASK.Has_AccountID Then
    '                MyStatus(msg & " (" & Client.Name & ")...")

    '                AF.ResetFiles(Client, StartDate, EndDate)
    '                Application.DoEvents()
    '            End If
    '        Next

    '        MyStatus(msg & "...Done")
    '    End Sub
    '#End Region 'ASK Functions

#Region " GA Functions "

    Private Sub Start_GA()
        Dim msg As String = "Starting GA Reports"
        MyStatus(msg & "...")

        For Each Client As Clients In Me.ClientList.Clients
            If Client.GA.Has_AccountID And Client.GA.Downloaded = False Then
                Client.GA.ClearDownloadInfo()
                Client.GA.StartDate = StartDate()
                Client.GA.EndDate = EndDate()
                If GAW.DoesFileExist(Client) = False Then
                    'Log Requesting action
                    Log(Client.LogName & ": Requesting GA Report")
                    MyStatus(msg & " (" & Client.Name & ")...")
                    'Perform Report Build and Request
                    Dim ReportID As String = GAW.ReqestReports(Client)
                    Log(Client.LogName & ": GA ID: " & ReportID)
                    'If error in request for report; perform clean up
                    If ReportID = "0" Or ReportID = "ERROR" Then
                        Me.ClientList.Done_Failed(Client.Index, WebService.GA)
                    Else
                        Me.ClientList.Done_ReportID(Client.Index, WebService.GA, ReportID)
                        Me.ClientList.Done_Download(Client.Index, WebService.GA)
                        nSteps = 6
                    End If
                Else
                    Me.ClientList.Done_Download(Client.Index, WebService.GA)
                    nSteps = 6
                End If
                Application.DoEvents()
            End If
        Next

        '-- START Feb 3, 2015 Mods
        'Set the next step to 5 rather than 6, since MSN EOM, we must work through
        ' checking status of request and then downloading.
        'Google the request also downloads

        If bDoMSNEOM = True Then
            nSteps = 5
        End If
        '-- END Feb 3, 2015 Mods
    End Sub




    Private Sub UpdateDB_GA()
        Dim msg As String = "Update DB from GA Revenue"
        MyStatus(msg & "...")

        StartLog(Me.StartDate, Me.EndDate, WebService.GA, "Updating")

        For Each Client As Clients In Me.ClientList.Clients
            If Client.GA.Has_AccountID Then
                MyStatus(msg & " (" & Client.Name & ")...")

                Log(Client.LogName & ": Updating GA")

                GAF.Execute(Client)

                Log(Client.LogName & ": GA Updated")

                Me.ClientList.Done_MyProcess(Client.Index, WebService.GA)
                Application.DoEvents()
            End If
        Next

        EndLog()

        MyStatus(msg & "...Done")
    End Sub

#End Region 'GA Functions


#Region " MSN Functions "

    Private Sub Start_MSNEOM()
        Dim sRowsDlt As String
        Dim msg As String = "Starting MSN EOM Delete"
        MyStatus(msg & "...")
        Log(msg)

        'Step 1: Update IsArbitrage Clients object Client.MSN.AccountID="" to remove MSN Account and prevent pulling month of data
        'Bypass IsArbitrage clients
        For Each Client As Clients In Me.ClientList.Clients
            If Client.MSN.Has_AccountID And Client.IsArbitrage = 1 Then
                Client.MSN.AccountID = ""
            End If
        Next

        'Step 2: Delete past months MSN Perfdata
        'Bypass IsArbitrage clients
        For Each Client As Clients In Me.ClientList.Clients
            If Client.MSN.Has_AccountID And Client.IsArbitrage <> 1 And Client.MSN.EOMDeleted = False Then

                'Client.MSN.ClearDownloadInfo()   *NOTE: Removed May 5 2011
                Client.MSN.MSNEOM = True
                Client.MSN.StartDate = StartDate()
                Client.MSN.EndDate = EndDate()
                'Log Requesting action
                Log(Client.LogName & ": Deleting PerfData for " & Client.MSN.StartDate.ToShortDateString & " to " & Client.MSN.EndDate.ToShortDateString)
                MyStatus(msg & " (" & Client.Name & ")...")
                sRowsDlt = MF.ExecuteEOM(Client)
                'Log Requesting action
                Log(Client.LogName & " " & sRowsDlt & " rows deleted.")
                Client.MSN.EOMDeleted = True
                Application.DoEvents()
            End If
        Next

        MyStatus(msg & "...Done")
        Log(msg & "...Done")

        'msg = "Reseting file from GA"
        'MyStatus(msg & "...")
        'Log(msg)

        ''Step 3: Reset GA files moving out of "Processed" folder
        ''Bypass IsArbitrage clients
        'For Each Client As Clients In Me.ClientList.Clients
        '    If Client.GA.Has_AccountID And Client.IsArbitrage <> 1 And Client.GA.EOMDeleted = False Then
        '        Client.GA.StartDate = StartDate()
        '        Client.GA.EndDate = EndDate()
        '        'Log Requesting action
        '        Log(Client.LogName & ": Reseting File for " & Client.GA.StartDate.ToShortDateString & " to " & Client.GA.EndDate.ToShortDateString)
        '        MyStatus(msg & " (" & Client.Name & ")...")

        '        GAF.ResetFiles(Client, Client.GA.StartDate, Client.GA.EndDate)
        '        Client.GA.EOMDeleted = True

        '        Application.DoEvents()
        '    End If
        'Next

        'MyStatus(msg & "...Done")
        'Log(msg & "...Done")

    End Sub
  

    Private Sub Start_MSN()
        Dim msg As String = "Starting MSN Reports"
        MyStatus(msg & "...")

        'Clear the MSN Unknown Status Count
        MSN_Unknown_Count = 0

        For Each Client As Clients In Me.ClientList.Clients
            If Client.MSN.Has_AccountID And Client.MSN.Downloaded = False Then
                'Client.MSN.ClearDownloadInfo()  *NOTE: Removed May 5 2011
                Client.MSN.StartDate = StartDate()
                Client.MSN.EndDate = EndDate()

                If MW.DoesFileExist(Client) = False Then
                    MyStatus(msg & " (" & Client.Name & ")...")
                    Log(Client.LogName & ": Requesting MSN Report")

                    Dim ReportID As String = MW.ReqestReportID(Client)
                    Log(Client.LogName & ": MSN ID: " & ReportID)

                    Me.ClientList.Done_ReportID(Client.Index, WebService.MSN, ReportID)
                Else
                    Me.ClientList.Done_Download(Client.Index, WebService.MSN)
                End If

                Application.DoEvents()
            End If
        Next

        MyStatus(msg & "...Done")
    End Sub

    Private Sub Start_FlexMSNEOM()
        Dim intTotClients As Integer = 0
        Dim msgDownloadCnt As String = ""
        Dim msg As String = "Starting FlexEOM MSN"
        MyStatus(msg & "...")
        bDoMSN = True
        Dim ReportID As String

        '
        'Step 1: Update Not IsArbitrage Clients object Client.MSN.AccountID="" to remove MSN Account 
        'and prevent pulling month of data in Loop_Download_MSN()
        '
        For Each Client As Clients In Me.ClientList.Clients
            If Client.MSN.Has_AccountID And Client.IsArbitrage <> 1 Then
                Client.MSN.AccountID = ""
            End If
        Next
        '
        'Step 2: Request Report-Only process IsArbitrage Clients 
        '
        For Each Client As Clients In Me.ClientList.Clients
            If Client.MSN.Has_AccountID And Client.IsArbitrage = 1 Then
                'Client.MSN.ClearDownloadInfo()  *NOTE: Removed May 5 2011
                Client.MSN.StartDate = StartDate()
                Client.MSN.EndDate = EndDate()
                Client.MSN.MSNEOM = True

                If MW.DoesFileExist(Client) = False Then
                    'Log Requesting action
                    Log(Client.LogName & ": Requesting FlexEOM MSN Report")
                    MyStatus(msg & " (" & Client.Name & ")...")

                    ReportID = MW.ReqestReportID(Client)
                    Log(Client.LogName & ": MSN ID: " & ReportID)

                    Me.ClientList.Done_ReportID(Client.Index, WebService.MSN, ReportID)
                Else
                    Me.ClientList.Done_Download(Client.Index, WebService.MSN)
                End If

                intTotClients += 1
                Application.DoEvents()
            End If
        Next
        MyStatus(msg & "...Done")
        Log(msg & "...Done")

        '
        'Step 3: Download Report via Loop_Download_MSN() until ClientList.IsAllDownloaded_MSN = True 
        '
        msg = "Downloading Reports FlexEOM MSN Status"
        MyStatus(msg & "...")
        Log(msg)
        Dim intDownloadedTot As Integer = 0
        Dim intDownloaded As Integer = 0
        Dim AllDone As Boolean = True

        Do While Me.ClientList.IsAllDownloaded_MSN = False
            intDownloaded = Loop_Download_MSN()
            intDownloadedTot += intDownloaded
            If Me.ClientList.IsAllDownloaded_MSN = False Then AllDone = False
            msgDownloadCnt = "Downloaded " & CStr(intDownloaded)
            MyStatus(msgDownloadCnt & "...")
            'Adjust the Sleep time so we do not waste money (google)
            If AllDone = True Then
                'Put back to 0 seconds
                Me.nSleepSeconds = 0
                Me.nMinuteCount = 0
            Else
                If intDownloaded = 0 Then
                    'Add 5 seconds to sleep time (max 60)
                    Me.nSleepSeconds = Me.nSleepSeconds + 5
                    If Me.nSleepSeconds > 60 Then
                        nMinuteCount = nMinuteCount + 1
                        Me.nSleepSeconds = 60

                        'Halt Processes if we max out
                        If nMinuteCount = 7 Then

                            'Not responding so flag as all done
                            For Each Client As Clients In Me.ClientList.Clients
                                If Client.MSN.Has_AccountID And Client.IsArbitrage = 1 Then
                                    If Client.MSN.Has_AccountID Then
                                        Client.MSN.Downloaded = True
                                    End If
                                End If

                            Next
                        End If
                    End If
                Else
                    'Put back to 0 seconds
                    Me.nSleepSeconds = 0
                    Me.nMinuteCount = 0
                End If
            End If

        Loop
        msgDownloadCnt = "Downloaded " & CStr(intDownloadedTot)
        MyStatus(msgDownloadCnt & "...")
        MyStatus(msg & "...Done")
        Log(msgDownloadCnt)
        Log(msg & "...Done")

        '
        'Step : Update Database 
        '
        msg = "Update DB from MSN FlexEOM"
        MyStatus(msg & "...")

        StartLog(Me.StartDate, Me.EndDate, WebService.MSN, "Updating")

        For Each Client As Clients In Me.ClientList.Clients
            If Client.MSN.Has_AccountID And Client.IsArbitrage = 1 Then
                MyStatus(msg & " (" & Client.Name & ")...")

                Log(Client.LogName & ": Updating FLEX MSN")

                MF.Execute(Client, True)

                Log(Client.LogName & ": FLEX MSN Updated")

                Me.ClientList.Done_MyProcess(Client.Index, WebService.MSN)
                Application.DoEvents()
            End If
        Next

        EndLog()

        bDoMSN = False
        MyStatus(msg & "...Done")
        Log(msg & "...Done")
        nSteps = 98
    End Sub

    Private Sub GetStatus_MSN()
        Dim msg As String = "Update MSN Status"
        MyStatus(msg & "...")

        For Each Client As Clients In Me.ClientList.Clients
            If Client.MSN.Has_AccountID Then
                If Client.MSN.Downloaded = False Then
                    If Client.MSN.DownloadReady = False Then
                        MyStatus(msg & " (" & Client.Name & ")...")

                        Log(Client.LogName & ": Getting MSN Status")

                        Dim Status As String
                        Status = MW.GetReportStatus(Client)
                        Log(Client.LogName & ": MSN Status: " & Status)
                        If Status = "UNKNOWN" Then
                            MSN_Unknown_Count += 1
                        Else
                            MSN_Unknown_Count = 0
                        End If

                        If MSN_Unknown_Count = 10 Then
                            'Make this read done
                            Log("MSN ERROR on " & Client.LogName(True) & "GIVING UP")
                            Me.ClientList.Done_Download(Client.Index, WebService.MSN)
                        End If
                        'Status Update
                        If Status.Length > 0 Then
                            Me.ClientList.Done_Status(Client.Index, WebService.MSN, Status)
                        End If
                        'Ready - do URL update 
                        If Status = "READY" Then
                            Me.ClientList.Done_URL(Client.Index, WebService.MSN, Client.MSN.URL)
                        End If
                    End If
                End If
            End If
        Next

        MyStatus(msg & "...Done")
    End Sub

    Private Function Download_MSN_Once() As Integer
        Dim msg As String = "Downloading MSN Reports"
        MyStatus(msg & "...")
        Dim Downloaded As Integer = 0

        For Each Client As Clients In Me.ClientList.Clients
            If Client.MSN.Has_AccountID Then
                If Client.MSN.Downloaded = False Then
                    If Client.MSN.DownloadReady Then
                        MyStatus(msg & " (" & Client.Name & ")...")

                        Log(Client.LogName & ": Downloading MSN")

                        Me.MW.DownloadReport(Client)

                        Log(Client.LogName & ": Downloaded")

                        Me.ClientList.Done_Download(Client.Index, WebService.MSN)
                        Downloaded = Downloaded + 1
                    End If
                End If
            End If
        Next

        MyStatus(msg & "...Done")
        Return Downloaded
    End Function

    Private Sub UpdateDB_MSN()
        Dim bFlexEOM As Boolean = False
        Dim msg As String = "Update DB from MSN"
        Dim sResult As String = ""
        MyStatus(msg & "...")

        StartLog(Me.StartDate, Me.EndDate, WebService.MSN, "Updating")

        For Each Client As Clients In Me.ClientList.Clients
            If Client.MSN.Has_AccountID Then
                MyStatus(msg & " (" & Client.Name & ")...")


                Log(Client.LogName & ": Updating MSN")

                MF.Execute(Client, bFlexEOM)

                Log(Client.LogName & ": MSN Updated")

                Me.ClientList.Done_MyProcess(Client.Index, WebService.MSN)
                Application.DoEvents()
            End If
        Next

        'Added for Remarketing and "boomuserlist" keyword conversion via sp on SQL server
        Log("Stored Procedure for MSN Category Mapping Started: " & Now.ToString("hh:mm:ss"))
        Dim msgStoredProc As String = "Running Stored Procedure BoomUserList"
        MyStatus(msgStoredProc)
        sResult = MF.UpdateRemarketing_Keywords(StartDate(), EndDate())
        msgStoredProc = sResult & " Records Updated"
        MyStatus(msgStoredProc)
        Log("Stored Procedure for Category Mapping Started Ended: " & Now.ToString("hh:mm:ss"))

        EndLog()

        MyStatus(msg & "...Done")
    End Sub

    Private Sub ResetDB_MSN()
        Dim msg As String = "Reseting DB from MSN"
        MyStatus(msg & "...")

        For Each Client As Clients In Me.ClientList.Clients
            If Client.MSN.Has_AccountID Then
                MyStatus(msg & " (" & Client.Name & ")...")

                MF.ResetFiles(Client, StartDate, EndDate)
                Application.DoEvents()
            End If
        Next

        MyStatus(msg & "...Done")
    End Sub
#End Region 'MSN Functions

    '#Region " Verizon Functions "

    '    Private Sub Start_Verizon()
    '        Dim msg As String = "Starting Verizon Reports"
    '        MyStatus(msg & "...")

    '        For Each Client As Clients In Me.ClientList.Clients
    '            If Client.Verizon.Has_AccountID And Client.Verizon.Downloaded = False Then
    '                'Client.Verizon.ClearDownloadInfo()  Client.Verizon.ClearDownloadInfo()
    '                Client.Verizon.StartDate = StartDate()
    '                Client.Verizon.EndDate = EndDate()

    '                If VW.DoesFileExist(Client) = False Then
    '                    MyStatus(msg & " (" & Client.Name & ")...")
    '                    Log(Client.LogName & ": Requesting Verizon Report")

    '                    Dim ReportID As String = VW.ReqestDailyReportID(Client)
    '                    Log(Client.LogName & ": Verizon ID: " & ReportID)

    '                    Select Case ReportID
    '                        Case "ID ERROR"
    '                            Me.ClientList.Done_InvalidID(Client.Index, WebService.Verizon)

    '                        Case "ERROR"
    '                            Me.ClientList.Done_Failed(Client.Index, WebService.Verizon)

    '                        Case "RETRY"
    '                            Me.ClientList.Done_Failed(Client.Index, WebService.Verizon)

    '                        Case "TIMEOUT"
    '                            Me.ClientList.Done_Failed(Client.Index, WebService.Verizon)

    '                        Case Else
    '                            If IsNumeric(ReportID) Then
    '                                Me.ClientList.Done_ReportID(Client.Index, WebService.Verizon, ReportID)
    '                            End If
    '                    End Select
    '                    Me.ClientList.Done_Request(Client.Index, WebService.Verizon)
    '                Else
    '                    Me.ClientList.Done_Download(Client.Index, WebService.Verizon)
    '                End If
    '                Application.DoEvents()
    '            End If
    '        Next

    '        MyStatus(msg & "...Done")
    '    End Sub

    '    Private Sub GetURL_Verizon()
    '        Dim msg As String = "Getting Verizon URLs"
    '        MyStatus(msg & "...")

    '        For Each Client As Clients In Me.ClientList.Clients
    '            If Client.Verizon.Has_AccountID Then
    '                If Client.Verizon.Has_ReportID Then
    '                    If Client.Verizon.DownloadReady = False Then
    '                        If Client.Verizon.Has_URL = False Then
    '                            If Client.Verizon.UrlRequested = False Then
    '                                MyStatus(msg & " (" & Client.Name & ")...")
    '                                Dim sURL As String = Me.VW.RequestReportURL(Client)

    '                                If sURL.ToLower.IndexOf("empty") > -1 Then
    '                                    'This is the same as a good url
    '                                    Debug.Print("Empty Report URL for Client " & Client.Name)
    '                                    Me.ClientList.Done_URL(Client.Index, WebService.Verizon, sURL)
    '                                    Client.Verizon.NoData = True
    '                                Else

    '                                    Select Case sURL
    '                                        Case "EMPTY"
    '                                            'This is the same as a good url
    '                                            Debug.Print("Empty URL for Client " & Client.Name)
    '                                            Me.ClientList.Done_URL(Client.Index, WebService.Verizon, sURL)
    '                                            Client.Verizon.NoData = True

    '                                        Case "RETRY"
    '                                            Debug.Print("Report URL not found for Client " & Client.Name)

    '                                        Case "UNKNOWN"
    '                                            'Don't do anything at this time

    '                                        Case Else
    '                                            'This is a good url
    '                                            If sURL.Length > 0 Then
    '                                                Me.ClientList.Done_URL(Client.Index, WebService.Verizon, sURL)
    '                                            End If
    '                                    End Select
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Next

    '        MyStatus(msg & "...Done")
    '    End Sub

    '    Private Function Download_Verizon_Once() As Integer
    '        Dim msg As String = "Downloading Verizon Reports"
    '        MyStatus(msg & "...")
    '        Dim Downloaded As Integer = 0

    '        For Each Client As Clients In Me.ClientList.Clients
    '            If Client.Verizon.Has_AccountID Then
    '                If Client.Verizon.Downloaded = False Then
    '                    If Client.Verizon.DownloadReady Then
    '                        If Client.Verizon.NoData = True Then
    '                            Me.ClientList.Done_Download(Client.Index, WebService.Verizon)
    '                            Downloaded = Downloaded + 1
    '                        Else
    '                            MyStatus(msg & " (" & Client.Name & ")...")

    '                            Log(Client.LogName & ": Downloading Verizon")

    '                            Me.VW.DownloadReport(Client)

    '                            Log(Client.LogName & ": Downloaded")

    '                            Me.ClientList.Done_Download(Client.Index, WebService.Verizon)
    '                            Downloaded = Downloaded + 1
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Next

    '        MyStatus(msg & "...Done")
    '        Return Downloaded
    '    End Function

    '    Private Sub UpdateDB_Verizon()
    '        Dim msg As String = "Update DB from Verizon"
    '        MyStatus(msg & "...")

    '        StartLog(Me.StartDate, Me.EndDate, WebService.Verizon, "Updating")

    '        For Each Client As Clients In Me.ClientList.Clients
    '            If Client.Verizon.Has_AccountID Then
    '                MyStatus(msg & " (" & Client.Name & ")...")


    '                Log(Client.LogName & ": Updating Verizon")

    '                VF.Execute(Client)

    '                Log(Client.LogName & ": Verizon Updated")

    '                Me.ClientList.Done_MyProcess(Client.Index, WebService.Verizon)
    '                Application.DoEvents()
    '            End If
    '        Next

    '        EndLog()

    '        MyStatus(msg & "...Done")
    '    End Sub

    '    Private Sub ResetDB_Verizon()
    '        Dim msg As String = "Reseting DB from Verizon"
    '        MyStatus(msg & "...")

    '        For Each Client As Clients In Me.ClientList.Clients
    '            If Client.Verizon.Has_AccountID Then
    '                MyStatus(msg & " (" & Client.Name & ")...")

    '                VF.ResetFiles(Client, StartDate, EndDate)
    '                Application.DoEvents()
    '            End If
    '        Next

    '        MyStatus(msg & "...Done")
    '    End Sub
    '#End Region 'Verizon Functions

    '#Region " Yahoo Functions "
    '    '=============================================================================
    '    'Start_Yahoo()
    '    'Invoke the YahooReports.vb methods that wrap around Yahoo API 
    '    '=============================================================================
    '    Private Sub Start_Yahoo()
    '        Dim msg As String = "Starting Yahoo Reports"
    '        MyStatus(msg & "...")

    '        For Each Client As Clients In Me.ClientList.Clients
    '            If ((Client.Yahoo.Has_AccountID) And (Client.Yahoo.Key_Downloaded = False Or Client.Yahoo.CM_Downloaded = False)) Then
    '                Client.Yahoo.ClearDownloadInfo()

    '                'Clear all OLD report jobs
    '                YW.ClearAllReports(Client)

    '                If Settings.YahooOnlyProcessFromEndDate Then
    '                    Client.Yahoo.StartDate = EndDate()
    '                    Client.Yahoo.EndDate = EndDate()
    '                Else
    '                    Client.Yahoo.StartDate = StartDate()
    '                    Client.Yahoo.EndDate = EndDate()
    '                End If

    '                'Request the Keyword Report
    '                If YW.DoesFileExist(Client, False) = False Then
    '                    MyStatus(msg & " (" & Client.Name & ")...")
    '                    Log(Client.LogName & ": Requesting Yahoo(Key) Report")

    '                    Dim ReportID As String = YW.ReqestDailyReportID(Client, False)
    '                    Log(Client.LogName & ": Yahoo(Key) ID: " & ReportID)

    '                    Me.ClientList.Done_ReportID(Client.Index, WebService.Yahoo, ReportID, False)
    '                Else
    '                    Me.ClientList.Done_Download(Client.Index, WebService.Yahoo, False)
    '                End If
    '                Application.DoEvents()

    '                'Request the Content Match Report
    '                If YW.DoesFileExist(Client, True) = False Then
    '                    MyStatus(msg & " (" & Client.Name & ")...")
    '                    Log(Client.LogName & ": Requesting Yahoo(CM) Report")

    '                    Dim ReportID As String = YW.ReqestDailyReportID(Client, True)
    '                    Log(Client.LogName & ": Yahoo(CM) ID: " & ReportID)

    '                    Me.ClientList.Done_ReportID(Client.Index, WebService.Yahoo, ReportID, True)
    '                Else
    '                    Me.ClientList.Done_Download(Client.Index, WebService.Yahoo, True)
    '                End If
    '                Application.DoEvents()
    '            End If
    '        Next

    '        MyStatus(msg & "...Done")
    '    End Sub
    '    '=============================================================================
    '    'GetStatus_Yahoo() As Interger
    '    'Invoke the YahooReports.vb methods that wrap around Yahoo API to check status
    '    '=============================================================================
    '    Private Function GetStatus_Yahoo() As Integer
    '        Dim msg As String = "Update Yahoo Status"
    '        MyStatus(msg & "...")

    '        For Each Client As Clients In Me.ClientList.Clients
    '            If Client.Yahoo.Has_AccountID Then

    '                'Keyword report status check
    '                If Client.Yahoo.Key_Downloaded = False Then
    '                    If Client.Yahoo.Key_DownloadReady = False Then
    '                        MyStatus(msg & " (" & Client.Name & ")...")

    '                        Log(Client.LogName & ": Getting Yahoo(Key) Status")
    '                        Dim Status As String = YW.GetReportStatus(Client, False)
    '                        Log(Client.LogName & ": Yahoo(Key) Status: " & Status)

    '                        If Status.Length > 0 Then
    '                            Me.ClientList.Done_Status(Client.Index, WebService.Yahoo, Status, False)
    '                        End If
    '                    End If
    '                End If
    '                'Content Match or AdGroup report status check
    '                If Client.Yahoo.CM_Downloaded = False Then
    '                    If Client.Yahoo.CM_DownloadReady = False Then
    '                        MyStatus(msg & " (" & Client.Name & ")...")

    '                        Log(Client.LogName & ": Getting Yahoo(CM) Status")
    '                        Dim Status As String = YW.GetReportStatus(Client, True)
    '                        Log(Client.LogName & ": Yahoo(CM) Status: " & Status)

    '                        If Status.Length > 0 Then
    '                            Me.ClientList.Done_Status(Client.Index, WebService.Yahoo, Status, True)
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Next

    '        MyStatus(msg & "...Done")
    '    End Function
    '    '=============================================================================
    '    'Download_Yahoo_Once() As Interger
    '    'Invoke the YahooReports.vb methods that wrap around Yahoo API to get 
    '    '=============================================================================
    '    Private Function Download_Yahoo_Once() As Integer
    '        Dim msg As String = "Downloading Yahoo Reports"
    '        MyStatus(msg & "...")
    '        Dim Downloaded As Integer = 0

    '        For Each Client As Clients In Me.ClientList.Clients
    '            If Client.Yahoo.Has_AccountID Then

    '                'Keyword Report Download
    '                If Client.Yahoo.Key_Downloaded = False Then
    '                    If Client.Yahoo.Key_DownloadReady Then
    '                        MyStatus(msg & " (" & Client.Name & ")...")

    '                        Log(Client.LogName & ": Downloading Yahoo(Key)")

    '                        Me.YW.DownloadReport(Client, False)

    '                        Log(Client.LogName & ": Downloaded")

    '                        Me.ClientList.Done_Download(Client.Index, WebService.Yahoo, False)
    '                        Downloaded = Downloaded + 1
    '                    End If
    '                End If

    '                'CM Report Download
    '                If Client.Yahoo.CM_Downloaded = False Then
    '                    If Client.Yahoo.CM_DownloadReady Then
    '                        MyStatus(msg & " (" & Client.Name & ")...")

    '                        Log(Client.LogName & ": Downloading Yahoo(CM)")

    '                        Me.YW.DownloadReport(Client, True)

    '                        Log(Client.LogName & ": Downloaded")

    '                        Me.ClientList.Done_Download(Client.Index, WebService.Yahoo, True)
    '                        Downloaded = Downloaded + 1
    '                    End If
    '                End If
    '            End If
    '        Next

    '        MyStatus(msg & "...Done")
    '        Return Downloaded
    '    End Function

    '    Private Sub UpdateDB_Yahoo()
    '        Dim msg As String = "Update DB from Yahoo"
    '        MyStatus(msg & "...")

    '        StartLog(Me.StartDate, Me.EndDate, WebService.Yahoo, "Updating")

    '        For Each Client As Clients In Me.ClientList.Clients
    '            If Client.Yahoo.Has_AccountID Then
    '                MyStatus(msg & " (" & Client.Name & ")...")

    '                Log(Client.LogName & ": Updating Yahoo")

    '                YF.Execute(Client)

    '                Log(Client.LogName & ": Yahoo Updated")

    '                Me.ClientList.Done_MyProcess(Client.Index, WebService.Yahoo)
    '                Application.DoEvents()
    '            End If
    '        Next

    '        EndLog()

    '        MyStatus(msg & "...Done")
    '    End Sub

    '    Private Sub ResetDB_Yahoo()
    '        Dim msg As String = "Reseting DB from Yahoo"
    '        MyStatus(msg & "...")

    '        For Each Client As Clients In Me.ClientList.Clients
    '            If Client.Yahoo.Has_AccountID Then
    '                MyStatus(msg & " (" & Client.Name & ")...")

    '                YF.ResetFiles(Client, StartDate, EndDate)
    '                Application.DoEvents()
    '            End If
    '        Next

    '        MyStatus(msg & "...Done")
    '    End Sub
    '#End Region 'Yahoo Functions

#Region " Private Functions and Stuff "
    Private Sub MyStatus(ByVal sMessage As String)
        Me.lblStatus.Text = sMessage
        Application.DoEvents()
    End Sub

    Private Sub UpdateLine(ByVal sLine)
        Me.lblLine.Text = sLine
        Application.DoEvents()
    End Sub

    Private Function StartDate() As Date
        Return CDate(dtStart.Value.ToString("MM/dd/yyyy"))
    End Function

    Private Function EndDate() As Date
        Return CDate(dtEnd.Value.ToString("MM/dd/yyyy"))
    End Function

    Private Sub UpdateSleep(ByVal nNewSleep As Integer)
        Me.nSleepingCount = nNewSleep
        If nNewSleep <= 0 Then
            MyStatus("")
        Else
            MyStatus("Sleeping (" & nNewSleep.ToString & ")...")
        End If
    End Sub

    Private Sub Running(ByVal IsRunning As Boolean)
        Me.Panel3.Enabled = Not IsRunning
        Me.ClientList.Freeze(IsRunning)
        Application.DoEvents()
    End Sub

    Private Sub SetWebServices(ByVal myWebservices As WebService)
        'Set all To False
        Me.bDoGoogle = False
        Me.bDoYahoo = False
        Me.bDoMSN = False
        Me.bDoVerizon = False
        Me.bDoASK = False
        Me.bDoGA = False

        'Set the appropriate ones
        Select Case myWebservices
            Case WebService.All
                Me.bDoGoogle = True
                Me.bDoYahoo = True
                Me.bDoMSN = True
                Me.bDoVerizon = True
                Me.bDoASK = True
                Me.bDoGA = True

            Case WebService.Google
                Me.bDoGoogle = True

            Case WebService.MSN
                Me.bDoMSN = True

            Case WebService.Verizon
                Me.bDoVerizon = True

            Case WebService.Yahoo
                Me.bDoYahoo = True

            Case WebService.ASK
                Me.bDoASK = True

            Case WebService.GA
                Me.bDoGA = True
        End Select
    End Sub
#End Region 'Private Functions and Stuff

#Region " Control and Class Events "
    Private Sub dtStart_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtStart.ValueChanged
        dtEnd.Value = dtStart.Value
    End Sub

    Private Sub dtEnd_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtEnd.ValueChanged
        If dtEnd.Value < dtStart.Value Then dtStart.Value = dtEnd.Value
    End Sub

    'Private Sub VW_GotURL(ByVal ClientIndex As Integer, ByVal URL As String) Handles VW.GotURL
    '    Select Case URL
    '        Case "EMPTY"
    '            'This is the same as a good url
    '            Debug.Print("Empty URL (" & ClientIndex.ToString & ")")
    '            Me.ClientList.Done_URL(ClientIndex, WebService.Verizon, URL)

    '        Case "RETRY"
    '            'Retry to get the URL
    '            Debug.Print("Retrying URL (" & ClientIndex.ToString & ")")
    '            Me.VW.RequestReportURL_Async(Me.ClientList.Clients(ClientIndex))
    '            Me.ClientList.Done_URLRequest(ClientIndex, WebService.Verizon)

    '        Case "UNKNOWN"
    '            'Retry to get the URL (also)
    '            Debug.Print("Unknown URL (" & ClientIndex.ToString & ")")
    '            Me.VW.RequestReportURL_Async(Me.ClientList.Clients(ClientIndex))
    '            Me.ClientList.Done_URLRequest(ClientIndex, WebService.Verizon)

    '        Case Else
    '            'This is a good url
    '            Me.ClientList.Done_URL(ClientIndex, WebService.Verizon, URL)

    '    End Select
    'End Sub

    'Private Sub VW_Scheduled(ByVal ClientIndex As Integer, ByVal ReportID As String) Handles VW.Scheduled
    '    Select Case ReportID
    '        Case "ID ERROR"
    '            Me.ClientList.Done_InvalidID(ClientIndex, WebService.Verizon)

    '        Case "ERROR"
    '            VW.ReqestDailyReportID_Async(Me.ClientList.Clients(ClientIndex))
    '            Me.ClientList.Done_Request(ClientIndex, WebService.Verizon)

    '        Case Else
    '            Me.ClientList.Done_ReportID(ClientIndex, WebService.Verizon, ReportID)

    '    End Select
    'End Sub

    'Private Sub YF_NewLine(ByVal sLine As String) Handles YF.NewLine
    '    Me.Line(sLine)
    'End Sub

    'Private Sub MF_NewLine(ByVal sLine As String) Handles MF.NewLine
    '    Me.Line(sLine)
    'End Sub

    Private Sub GF_NewLine(ByVal sLine As String) Handles GF.NewLine
        Me.Line(sLine)
    End Sub

    'Private Sub VF_NewLine(ByVal sLine As String) Handles VF.NewLine
    '    Me.Line(sLine)
    'End Sub
#End Region 'Control and Class Events

#Region " Public Subs Status and Line "
    Public Sub Status(ByVal Text As String)
        Me.lblStatus.Text = Text
    End Sub

    Public Sub Line(ByVal Text As String)
        Me.lblLine.Text = Text
    End Sub
#End Region 'Public Subs Status and Line

#Region " CheckForProcessed Functions "
    Private Function CheckForProcessed_Google(ByVal bCheckProcessedDir As Boolean) As String()
        Dim Clients() As Clients = Me.ClientList.Clients
        Return GF.CheckNotProcessed(Clients, EndDate, bCheckProcessedDir)
    End Function

    Private Function CheckForProcessed_MSN(ByVal bCheckProcessedDir As Boolean) As String()
        Dim Clients() As Clients = Me.ClientList.Clients
        Return MF.CheckNotProcessed(Clients, EndDate, bCheckProcessedDir)
    End Function

    'Private Function CheckForProcessed_Verizon(ByVal bCheckProcessedDir As Boolean) As String()
    '    Dim Clients() As Clients = Me.ClientList.Clients
    '    Return VF.CheckNotProcessed(Clients, EndDate, bCheckProcessedDir)
    'End Function

    'Private Function CheckForProcessed_Yahoo(ByVal bCheckProcessedDir As Boolean) As String()
    '    Dim Clients() As Clients = Me.ClientList.Clients
    '    Return YF.CheckNotProcessed(Clients, EndDate, bCheckProcessedDir)
    'End Function
    'Private Function CheckForProcessed_ASK(ByVal bCheckProcessedDir As Boolean) As String()
    '    Dim Clients() As Clients = Me.ClientList.Clients
    '    Return AF.CheckNotProcessed(Clients, EndDate, bCheckProcessedDir)
    'End Function
    Private Function CheckForProcessed_GA(ByVal bCheckProcessedDir As Boolean) As String()
        Dim Clients() As Clients = Me.ClientList.Clients
        Return GAF.CheckNotProcessed(Clients, EndDate, bCheckProcessedDir)
    End Function

    Private Sub DisplayMessage(ByVal Msg() As String, ByVal Ws As WebService)
        Dim myMsg As String = ""
        Select Case Ws
            Case WebService.Google
                myMsg = "Google: None Missing"
            Case WebService.MSN
                myMsg = "MSN: None Missing"
            Case WebService.Verizon
                myMsg = "Verizon: None Missing"
            Case WebService.Yahoo
                myMsg = "Yahoo: None Missing"
            Case WebService.ASK
                myMsg = "ASK: None Missing"
            Case WebService.GA
                myMsg = "GA: None Missing"
        End Select

        If Not IsNothing(Msg) Then
            If Msg.GetUpperBound(0) > 0 Then
                myMsg = ""
                Dim nCount As Integer
                For nCount = 0 To Msg.GetUpperBound(0)
                    myMsg = myMsg & Msg(nCount) & vbCrLf
                Next
            End If
        End If
        MsgBox(myMsg)
    End Sub
#End Region 'CheckForProcessed Functions

#Region " (Controls on Form) Events "

    '================================================================================================
    'CheckForProcessed Button Click Handler:
    'Scan Client's directory for EndDate engine data
    '================================================================================================
    Private Sub btnCheckForProcessed_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckForProcessed.Click
        Running(True)

        If bDoGoogle = True Then
            Dim sGoogleMsg() As String = Me.CheckForProcessed_Google(True)
            Me.DisplayMessage(sGoogleMsg, WebService.Google)
        End If

        'If bDoMSN = True Then
        '    Dim MSG_MSN() As String = Me.CheckForProcessed_MSN(True)
        '    Me.DisplayMessage(MSG_MSN, WebService.MSN)
        'End If

        ''If bDoYahoo = True Then
        ''    Dim MSG_Yahoo() As String = Me.CheckForProcessed_Yahoo(True)
        ''    Me.DisplayMessage(MSG_Yahoo, WebService.Yahoo)
        ''End If

        'If bDoVerizon = True Then
        '    Dim MSG_Verizon() As String = Me.CheckForProcessed_Verizon(True)
        '    Me.DisplayMessage(MSG_Verizon, WebService.Verizon)
        'End If

        'If bDoASK = True Then
        '    Dim MSG_ASK() As String = Me.CheckForProcessed_ASK(True)
        '    Me.DisplayMessage(MSG_ASK, WebService.ASK)
        'End If

        Running(False)
    End Sub
    '================================================================================================
    'Download Button Click Handler:
    'Download dates selected from engine(s) for client(s)
    '================================================================================================
    Private Sub btnDownload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDownload.Click
        bDoDownload = True
        bDoProcess = False
        bYahoo_CM = False

        nSteps = 0
        Execute()
    End Sub
    '================================================================================================
    'DownloadProcess Button Click Handler:
    'Download and update DB dates selected from engine(s) for client(s)
    '================================================================================================
    Private Sub btnDownloadProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDownloadProcess.Click
        bDoDownload = True
        bDoProcess = True
        bYahoo_CM = False

        nSteps = 0
        Execute()
    End Sub
    '================================================================================================
    'Process Button Click Handler:
    'Update DB dates selected from engine(s) for client(s) - No API used, processes existing files
    '================================================================================================
    Private Sub btnProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProcess.Click
        bDoDownload = False
        bDoProcess = True
        bYahoo_CM = False

        nSteps = 0
        Execute()
    End Sub
    '================================================================================================
    'Process Reset Click Handler:
    'Moves existing downloaded files from "Processed" folder and 
    'Update DB dates selected from engine(s) for client(s) - No API used, processes existing files
    '================================================================================================
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Running(True)
        Select Case Me.cmbWebService.SelectedItem
            Case WebService.All
                Me.ResetDB_ALL()

            Case WebService.Google
                Me.ResetDB_Google()

                'Case WebService.MSN
                '    Me.ResetDB_MSN()

                'Case WebService.Verizon
                '    Me.ResetDB_Verizon()

                'Case WebService.Yahoo
                '    Me.ResetDB_Yahoo()

                'Case WebService.ASK
                '    Me.ResetDB_ASK()
        End Select
        Running(False)
    End Sub
    '================================================================================================
    'Process WebService_Selected Combo Box Selected Value Changed Event Handler:
    'Selected WebService boolean updated: bDoGoogle; bDoYahoo; bDoMSN; bDoVerizon; bDoASK  
    '================================================================================================
    Private Sub cmbWebService_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbWebService.SelectedValueChanged
        SetWebServices(Me.cmbWebService.SelectedItem)
    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.Timer1.Enabled = False
        Me.Timer1.Stop()

        UpdateSleep(nSleepingCount - 1)
        If nSleepingCount <= 0 Then
            If Me.bFormLoading = True Then
                RunAuto()
            Else
                Execute()
            End If
        Else
            Timer1.Interval = 1000
            Timer1.Start()
        End If
    End Sub
    'Special Functions Buttons; display conditioned by settings.xml FixButton
    '================================================================================================
    'Process Reset Click Handler:
    'Moves existing downloaded files from "Processed" folder and 
    'Update DB dates selected from engine(s) for client(s) - No API used, processes existing files
    '================================================================================================
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'RenameDirs()
        'Using to reset all files for Google May EOM
        Running(True)
        Select Case Me.cmbWebService.SelectedItem
            Case WebService.All
                Me.ResetDB_ALL()

            Case WebService.Google
                Me.ResetEOMDB_Google()

                'Case WebService.MSN
                '    Me.ResetEOMDB_MSN()

                'Case WebService.Verizon
                '    Me.ResetDB_Verizon()

                'Case WebService.Yahoo
                '    Me.ResetEOMDB_Yahoo()
                '    'Me.ResetDB_Yahoo()

                'Case WebService.ASK
                '    Me.ResetDB_ASK()
        End Select
        Running(False)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim sRowsDlt As String
        Dim msg As String = "Starting Google EOM Delete"
        MyStatus(msg & "...")
        Log(msg)

        ''Step 1: Update IsArbitrage Clients object Client.Google.AccountID="" to remove Google Account and prevent pulling month of data
        ''Only IsArbitrage clients
        'For Each Client As Clients In Me.ClientList.Clients
        '    If Client.Google.Has_AccountID And Client.IsArbitrage = 1 Then
        '        Client.Google.AccountID = ""
        '    End If
        'Next

        ''Step 2: Delete past months Google Perfdata
        ''Bypass IsArbitrage clients
        'For Each Client As Clients In Me.ClientList.Clients
        '    If Client.Google.Has_AccountID And Client.IsArbitrage <> 1 And Client.Google.EOMDeleted = False Then

        '        Client.Google.ClearDownloadInfo()
        '        Client.Google.GoogleEOM = True
        '        Client.Google.StartDate = StartDate()
        '        Client.Google.EndDate = EndDate()
        '        'Log Requesting action
        '        Log(Client.LogName & ": Deleting PerfData for " & Client.Google.StartDate.ToShortDateString & " to " & Client.Google.EndDate.ToShortDateString)
        '        MyStatus(msg & " (" & Client.Name & ")...")
        '        sRowsDlt = GF.ExecuteEOM(Client)
        '        'Log Requesting action
        '        Log(Client.LogName & " " & sRowsDlt & " rows deleted.")
        '        Client.Google.EOMDeleted = True
        '        Application.DoEvents()
        '    End If
        'Next

        'MyStatus(msg & "...Done")
        'Log(msg & "...Done")

        'msg = "Reseting file from GA"
        'MyStatus(msg & "...")
        'Log(msg)

        ''Step 3: Reset GA files moving out of "Processed" folder
        ''Bypass IsArbitrage clients
        'For Each Client As Clients In Me.ClientList.Clients
        '    If Client.GA.Has_AccountID And Client.IsArbitrage <> 1 And Client.GA.EOMDeleted = False Then
        '        Client.GA.StartDate = StartDate()
        '        Client.GA.EndDate = EndDate()
        '        'Log Requesting action
        '        Log(Client.LogName & ": Reseting File for " & Client.GA.StartDate.ToShortDateString & " to " & Client.GA.EndDate.ToShortDateString)
        '        MyStatus(msg & " (" & Client.Name & ")...")

        '        GAF.ResetFiles(Client, Client.GA.StartDate, Client.GA.EndDate)
        '        Client.GA.EOMDeleted = True

        '        Application.DoEvents()
        '    End If
        'Next

        MyStatus(msg & "...Done")
        Log(msg & "...Done")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim msg As String = "Reseting file from Flex Google"
        MyStatus(msg & "...")
        Log(msg)

        'Step 3: Reset FlexGoogle files moving out of "Processed" folder

        For Each Client As Clients In Me.ClientList.Clients
            If Client.Google.Has_AccountID And Client.IsArbitrage = 1 Then
                Client.Google.StartDate = StartDate()
                Client.Google.EndDate = EndDate()
                'Log Requesting action
                Log(Client.LogName & ": Reseting File for " & Client.Google.StartDate.ToShortDateString & " to " & Client.Google.EndDate.ToShortDateString)
                MyStatus(msg & " (" & Client.Name & ")...")

                GF.ResetFiles(Client, Client.Google.StartDate, Client.Google.EndDate)

                Application.DoEvents()
            End If
        Next

        MyStatus(msg & "...Done")
        Log(msg & "...Done")



        'Dim msg As String = "Reseting file from GA"
        'MyStatus(msg & "...")
        'Log(msg)

        ''Step 3: Reset GA files moving out of "Processed" folder
        ''Bypass IsArbitrage clients
        'For Each Client As Clients In Me.ClientList.Clients
        '    If Client.GA.Has_AccountID And Client.IsArbitrage <> 1 And Client.GA.EOMDeleted = False Then
        '        Client.GA.StartDate = StartDate()
        '        Client.GA.EndDate = EndDate()
        '        'Log Requesting action
        '        Log(Client.LogName & ": Reseting File for " & Client.GA.StartDate.ToShortDateString & " to " & Client.GA.EndDate.ToShortDateString)
        '        MyStatus(msg & " (" & Client.Name & ")...")

        '        GAF.ResetFiles(Client, Client.GA.StartDate, Client.GA.EndDate)
        '        Client.GA.EOMDeleted = True

        '        Application.DoEvents()
        '    End If
        'Next

        'MyStatus(msg & "...Done")
        'Log(msg & "...Done")

        'bDoDownload = True
        'bDoProcess = False
        'SendEmails()
        'bDoDownload = False
        'bDoProcess = False

    End Sub
#End Region '(Controls on Form) Events

#Region " Rename Directories ONE TIME UPDATE"
    Private Sub RenameDirs()
        Dim msg As String

        '*********** Google ****************
        msg = "Changing Dir for Google"
        MyStatus(msg & "...")

        For Each Client As Clients In Me.ClientList.Clients
            If Client.Google.Has_AccountID Then
                MyStatus(msg & " (" & Client.Name & ")...")

                GF.RenameDirs(Client)
                Application.DoEvents()
            End If
        Next

        MyStatus(msg & "...Done")

        ''*********** MSN ****************
        'msg = "Changing Dir for MSN"
        'MyStatus(msg & "...")

        'For Each Client As Clients In Me.ClientList.Clients
        '    If Client.MSN.Has_AccountID Then
        '        MyStatus(msg & " (" & Client.Name & ")...")

        '        MF.RenameDirs(Client)
        '        Application.DoEvents()
        '    End If
        'Next

        'MyStatus(msg & "...Done")

        ''*********** Verizon ****************
        'msg = "Changing Dir for Verizon"
        'MyStatus(msg & "...")

        'For Each Client As Clients In Me.ClientList.Clients
        '    If Client.Verizon.Has_AccountID Then
        '        MyStatus(msg & " (" & Client.Name & ")...")

        '        VF.RenameDirs(Client)
        '        Application.DoEvents()
        '    End If
        'Next

        'MyStatus(msg & "...Done")

        ''*********** Yahoo ****************
        'msg = "Changing Dir for Yahoo"
        'MyStatus(msg & "...")

        'For Each Client As Clients In Me.ClientList.Clients
        '    If Client.Yahoo.Has_AccountID Then
        '        MyStatus(msg & " (" & Client.Name & ")...")

        '        YF.RenameDirs(Client)
        '        Application.DoEvents()
        '    End If
        'Next

        'MyStatus(msg & "...Done")

    End Sub
#End Region 'Rename Directories

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class