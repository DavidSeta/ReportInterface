Imports System
Imports System.IO

#Region " Class modCommon "
'======================================================================================================
'Module modCommon - Common tasks used throughout application
'======================================================================================================
Module modCommon
    Public SaveLoc As String = ""
    Public MSN_Unknown_Count As Integer = 0
    '==================================================================================================================
    'EndLog()
    'Final Log entry, documenting dates and engines.
    '==================================================================================================================
    Public Sub EndLog()
        Dim Data As New ArrayList

        Data.Add("----------------------------------")
        Data.Add("Process Ended: " & Now.ToString("hh:mm:ss"))
        Data.Add("----------------------------------")
        Data.Add("")

        Log(Data)
    End Sub
    '==================================================================================================================
    'StartLog(ByVal StartDate As Date, ByVal EndDate As Date, ByVal Type As WebService, ByVal Command As String)
    'Initial Log entry, documenting dates and engines.
    '==================================================================================================================
    Public Sub StartLog(ByVal StartDate As Date, ByVal EndDate As Date, ByVal Type As WebService, ByVal Command As String)
        Dim Data As New ArrayList

        Data.Add("")
        Data.Add("----------------------------------")
        Data.Add(Command.Trim & " Process Started: " & Now.ToString("hh:mm:ss"))
        Data.Add(Command.Trim & " Start Date: " & StartDate.ToString("yyyy-MM-dd"))
        Data.Add(Command.Trim & " End Date: " & EndDate.ToString("yyyy-MM-dd"))
        Select Case Type
            Case WebService.All
                Data.Add(Command.Trim & " Type: ALL")

            Case WebService.Google
                Data.Add(Command.Trim & " Type: GOOGLE")

            Case WebService.MSN
                Data.Add(Command.Trim & " Type: MSN")

            Case WebService.Verizon
                Data.Add(Command.Trim & " Type: VERIZON")

            Case WebService.Yahoo
                Data.Add(Command.Trim & " Type: YAHOO")

            Case WebService.ASK
                Data.Add(Command.Trim & " Type: ASK")

            Case WebService.GA
                Data.Add(Command.Trim & " Type: GA")
        End Select
        Data.Add("----------------------------------")

        Log(Data)
    End Sub
    '==================================================================================================================
    'MyFile() As String - Create current log file name for writing transactions throughout the application.
    'Format is : Log_YYYY_MM_DD.txt
    '==================================================================================================================
    Public Function MyFile() As String
        Dim sFile As String = SaveLoc
        If Right(sFile, 1) <> "\" Then sFile = sFile & "\"
        sFile = sFile & "Log_" & Today.ToString("yyyy_MM_dd") & ".txt"
        Return sFile
    End Function
    '==================================================================================================================
    'ErrLogFile() As String - Create currrent log file name for error reporting in order to avoid locks.
    'Format is : Log_YYYY_MM_DD.txt
    '==================================================================================================================
    Public Function ErrLogFile() As String
        Dim sFile As String = SaveLoc
        If Right(sFile, 1) <> "\" Then sFile = sFile & "\"
        sFile = sFile & "Errors\"
        sFile = sFile & "Log_" & Today.ToString("yyyy_MM_dd") & ".txt"
        Return sFile
    End Function
    '==================================================================================================================
    'Log(Data As String) - Write Data passed to current log file.
    '==================================================================================================================
    Public Sub Log(ByVal Data As String)
        Using sw As StreamWriter = New StreamWriter(MyFile, True)
            sw.WriteLine(Data)
            Debug.Print(Data)
            sw.Close()
        End Using
    End Sub
    '==================================================================================================================
    'Log(Data As ArrayList) - Write Data ArrayList passed to current log file.
    '==================================================================================================================
    Public Sub Log(ByVal Data As ArrayList)
        Using sw As StreamWriter = New StreamWriter(MyFile, True)
            For Each sText As String In Data
                sw.WriteLine(sText)
                Debug.Print(sText)
            Next
            sw.Close()
        End Using
    End Sub

End Module
#End Region 'Class modCommon

#Region " Public Enums "
Public Enum FilterList
    All_Clients
    Google_Clients
    Google_5_Clients
    MSN_Clients
    Verizon_Clients
    Yahoo_Clients
    ASK_Clients
    GA_Clients
    Arb_Clients
End Enum

Public Enum WebService
    All
    Google
    MSN
    Verizon
    Yahoo
    ASK
    GA
End Enum
Public Enum Dimension
    browser = 0
    browserVersion = 1
    city = 2
    connectionSpeed = 3
    continent = 4
    countOfVisits = 5
    country = 6
    dateenum = 7
    day = 8
    daysSinceLastVisit = 9
    flashVersion = 10
    hostname = 11
    hour = 12
    javaEnabled = 13
    language = 14
    latitude = 15
    longitude = 16
    month = 17
    networkDomain = 18
    networkLocation = 19
    pageDepth = 20
    operatingSystem = 21
    operatingSystemVersion = 22
    region = 23
    screenColors = 24
    screenResolution = 25
    subContinent = 25
    userDefinedValue = 26
    visitorType = 26
    week = 27
    year = 28
    adContent = 29
    adGroup = 30
    adSlot = 31
    adSlotPosition = 32
    campaign = 33
    keyword = 34
    medium = 35
    referralPath = 36
    source = 37
    exitPagePath = 38
    landingPagePath = 39
    pagePath = 40
    pageTitle = 41
    affiliation = 42
    daysToTransaction = 43
    productCategory = 44
    productName = 45
    productSku = 46
    transactionId = 47
    searchCategory = 48
    searchDestinationPage = 49
    searchKeyword = 50
    searchKeywordRefinement = 51
    searchStartPage = 52
    searchUsed = 53
    customVarName1 = 54
    customVarName2 = 55
    customVarName3 = 56
    customVarName4 = 57
    customVarName5 = 58
    customVarValue1 = 59
    customVarValue2 = 60
    customVarValue3 = 61
    customVarValue4 = 62
    customVarValue5 = 63
    eventCategory = 64
    eventAction = 65
    eventLabel = 66

End Enum
Public Enum Metric
    bounces = 0
    entrances = 1
    exits = 2
    newVisits = 3
    pageviews = 4
    timeOnPage = 5
    timeOnSite = 6
    visitors = 7
    visits = 8
    adCost = 9
    adClicks = 10
    CPC = 11
    CPM = 12
    CTR = 13
    impressions = 14
    uniquePageviews = 15
    itemQuantity = 16
    itemRevenue = 17
    transactionRevenue = 18
    transactions = 19
    transactionShipping = 20
    transactionTax = 21
    uniquePurchases = 22
    searchDepth = 23
    searchDuration = 24
    searchExits = 25
    searchRefinements = 26
    searchUniques = 27
    searchVisits = 28
    goal1Completions = 29
    goal2Completions = 30
    goal3Completions = 31
    goal4Completions = 32
    goalCompletionsAll = 33
    goal1Starts = 34
    goal2Starts = 35
    goal3Starts = 36
    goal4Starts = 37
    goalStartsAll = 38
    goal1Value = 39
    goal2Value = 40
    goal3Value = 41
    goal4Value = 42
    goalValueAll = 43
    pageviewsPerVisit = 44
    entranceRate = 45
    entranceBounceRate = 46
    visitBounceRate = 47
    avgTimeOnPage = 48
    exitRate = 49
    goalValuePerVisit = 50
    goal1ConversionRate = 51
    goal2ConversionRate = 52
    goal3ConversionRate = 53
    goal4ConversionRate = 54
    goal5ConversionRate = 55
    goalConversionRateAll = 56
    goal1Abandons = 57
    goal2Abandons = 58
    goal3Abandons = 59
    goal4Abandons = 60
    goal5Abandons = 61
    goalAbandonsAll = 62
    goal1AbandonRate = 63
    goal2AbandonRate = 64
    goal3AbandonRate = 65
    goal4AbandonRate = 66
    goal5AbandonRate = 67
    goalAbandonRateAll = 68
    percentNewVisits = 69
    avgTimeOnSite = 70
    goal16Completions = 71
End Enum
#End Region 'Public Enums

#Region " Clients Class "
'===========================================================================================
'Class Clients
'DataStructure to hold client specifics.  Main parameter passed throughout this project
'===========================================================================================
'11/1/09 Sherv added is Arbitrage property to the clients
Public Class Clients
    Private _Name As String
    Public LV_ID As Integer
    Public CustID As String
    Public Index As Integer
    Private _isArbitrage As Integer

    Public Yahoo As Data_WS_YahooX
    Public Google As Data_WS
    Public MSN As Data_WS
    Public Verizon As Data_WS
    Public ASK As Data_WS_ASK
    Public GA As Data_WS_GA
    Public Property IsArbitrage() As Integer
        Get
            Return _isArbitrage
        End Get
        Set(ByVal value As Integer)
            _isArbitrage = value
        End Set
    End Property
    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = Replace(value, ".", "")
        End Set
    End Property

    Public Function LogName(ByVal AddDash As Boolean) As String
        Dim sName As String = "(" & CustID & ") " & Name
        If AddDash Then sName = sName & " -- "
        Return sName
    End Function

    Public Function LogName() As String
        Return LogName(False)
    End Function


    Public ReadOnly Property HasAnyAccountID() As Boolean
        Get
            Dim Result As Boolean = False

            If Yahoo.Has_AccountID Then Result = True
            If Google.Has_AccountID Then Result = True
            If MSN.Has_AccountID Then Result = True
            If Verizon.Has_AccountID Then Result = True
            If ASK.Has_AccountID Then Result = True
            If GA.Has_AccountID Then Result = True

            Return Result

        End Get
    End Property
    '===========================================================================================
    'New()- Instantiate Clients class.
    '===========================================================================================
    Sub New()
        Name = ""
        CustID = ""
        Yahoo = New Data_WS_YahooX
        MSN = New Data_WS
        Verizon = New Data_WS
        Google = New Data_WS
        ASK = New Data_WS_ASK
        GA = New Data_WS_GA

    End Sub

End Class
#End Region ' Class Clients

#Region " Class Data_WS "

'===========================================================================================
'Class Data_WS - Common DataStructure used by Google, MSN, Verizon to store parameters
'and results from API.
'===========================================================================================
Public Class Data_WS
    Public GGLClientID As String
    Public AccountID As String
    Public AccountPWD As String

    Public ReportID As String
    Public EOMReportID As String
    Public GoogleEOM As Boolean
    Public MSNEOM As Boolean
    Public URL As String
    Public Status As String

    Private _StartDate As Date
    Private _EndDate As Date

    Public DownloadReady As Boolean
    Public UrlRequested As Boolean
    Public Downloaded As Boolean
    Public Completed As Boolean
    Public NoData As Boolean
    Public ErrorReport As Boolean
    Public EOMDeleted As Boolean
    '===========================================================================================
    'New()- Instantiate Data_WS class invoked from Clients.New()
    '===========================================================================================
    Sub New()
        AccountID = ""
        AccountPWD = ""
        ClearDownloadInfo()
    End Sub
    '===========================================================================================
    'ClearDownloadNotReport()- Initialization for newly instantiated Data_WS class.  Invoked from New
    '===========================================================================================
    Public Sub ClearDownloadNotReport()

        URL = ""
        Status = ""

        StartDate = Today
        EndDate = Today

        DownloadReady = False
        UrlRequested = False
        Downloaded = False
        Completed = False
        NoData = False
        ErrorReport = False
        EOMDeleted = False

    End Sub
    '===========================================================================================
    'ClearDownloadInfo()- Initialization for newly instantiated Data_WS class.  Invoked from New
    '===========================================================================================
    Public Sub ClearDownloadInfo()
        ReportID = ""
        URL = ""
        Status = ""

        StartDate = Today
        EndDate = Today

        DownloadReady = False
        UrlRequested = False
        Downloaded = False
        Completed = False
        NoData = False
        ErrorReport = False
        EOMDeleted = False
        GoogleEOM = False
        MSNEOM = False
    End Sub


    Public ReadOnly Property Has_AccountID() As Boolean
        Get
            If IsNothing(AccountID) Then Return False
            If AccountID = "" Then Return False
            Return True
        End Get
    End Property

    Public ReadOnly Property Has_AcctPwd() As Boolean
        Get
            If IsNothing(AccountPWD) Then Return False
            Return IIf(AccountPWD.Length > 0, True, False)
        End Get
    End Property

    Public ReadOnly Property Has_ReportID() As Boolean
        Get
            If IsNothing(ReportID) Then Return False
            If ReportID = "" Then Return False

            Return True

        End Get
    End Property
    Public ReadOnly Property Has_EOMReportID() As Boolean
        Get
            If IsNothing(EOMReportID) Then Return False
            If ReportID = "" Then Return False
            'Temporary for V13
            Return False
        End Get
    End Property

    Public ReadOnly Property Has_URL() As Boolean
        Get
            If IsNothing(URL) Then Return False
            If URL = "" Then Return False
            Return True
        End Get
    End Property

    Public ReadOnly Property MultiDayReport() As Boolean
        Get
            If EndDate = StartDate Then Return False
            Return True
        End Get
    End Property
    Public ReadOnly Property IsYesterday() As Boolean
        Get
            If EndDate = StartDate And StartDate = DateAdd(DateInterval.Day, -1, Today) Then
                Return True
            Else
                Return False
            End If

        End Get
    End Property

    Public Property EndDate() As Date
        Get
            Return _EndDate
        End Get
        Set(ByVal value As Date)
            _EndDate = value
            If _StartDate > value Then _StartDate = value
        End Set
    End Property

    Public Property StartDate() As Date
        Get
            Return _StartDate
        End Get
        Set(ByVal value As Date)
            _StartDate = value
            If _EndDate < value Then _EndDate = value
        End Set
    End Property
End Class

#End Region ' Class Data_WS

#Region " Class Data_WS_YahooX"
'===========================================================================================
'Class Data_WS - Common DataStructure used by Yahoo to store parameter and results from API.
'===========================================================================================
Public Class Data_WS_YahooX
    Public YSMV6ID As String
    Public AccountID As String
    Public AccountPWD As String

    Public GroupAcct As String
    Public GroupPWD As String

    Public SubAccountID As String = ""
    Public ContentMatch As Boolean = False

    Private _StartDate As Date
    Private _EndDate As Date

    Public Key_ReportID As String
    Public Key_URL As String
    Public Key_Status As String

    Public Key_DownloadReady As Boolean
    Public Key_UrlRequested As Boolean
    Public Key_Downloaded As Boolean
    Public Key_Completed As Boolean

    Public CM_ReportID As String
    Public CM_URL As String
    Public CM_Status As String

    Public CM_DownloadReady As Boolean
    Public CM_UrlRequested As Boolean
    Public CM_Downloaded As Boolean
    Public CM_Completed As Boolean
    '===========================================================================================
    'New()- Instantiate Data_WS_YahooX class invoked from Clients.New()
    '===========================================================================================
    Sub New()
        AccountID = ""
        AccountPWD = ""
        ClearDownloadInfo()
    End Sub
    '====================================================================================================
    'ClearDownloadInfo()- Initialization for newly instantiated Data_WS_YahooX class.  Invoked from New()
    '====================================================================================================
    Public Sub ClearDownloadInfo()
        StartDate = Today
        EndDate = Today

        Key_ReportID = ""
        Key_URL = ""
        Key_Status = ""

        Key_DownloadReady = False
        Key_UrlRequested = False
        Key_Downloaded = False
        Key_Completed = False

        CM_ReportID = ""
        CM_URL = ""
        CM_Status = ""

        CM_DownloadReady = False
        CM_UrlRequested = False
        CM_Downloaded = False
        CM_Completed = False
    End Sub

    Public ReadOnly Property Has_AccountID() As Boolean
        Get
            If IsNothing(AccountID) Then Return False
            If AccountID = "" Then Return False
            Return True
        End Get
    End Property

    Public ReadOnly Property Has_SubAccountID() As Boolean
        Get
            Return IIf(SubAccountID.Length > 0, True, False)
        End Get
    End Property

    Public ReadOnly Property Has_AcctPwd() As Boolean
        Get
            If IsNothing(AccountPWD) Then Return False
            Return IIf(AccountPWD.Length > 0, True, False)
        End Get
    End Property
    Public ReadOnly Property Has_GroupAcct() As Boolean
        Get
            If IsNothing(GroupAcct) Then Return False
            If GroupAcct = "" Then Return False
            Return True
        End Get
    End Property
    Public ReadOnly Property Has_GroupPwd() As Boolean
        Get
            If IsNothing(GroupPWD) Then Return False
            Return IIf(GroupPWD.Length > 0, True, False)
        End Get
    End Property

    Public ReadOnly Property Has_Key_ReportID() As Boolean
        Get
            If IsNothing(Key_ReportID) Then Return False
            If Key_ReportID = "" Then Return False
            Return True
        End Get
    End Property

    Public ReadOnly Property Has_CM_ReportID() As Boolean
        Get
            If IsNothing(CM_ReportID) Then Return False
            If CM_ReportID = "" Then Return False
            Return True
        End Get
    End Property

    Public ReadOnly Property Has_Key_URL() As Boolean
        Get
            If IsNothing(Key_URL) Then Return False
            If Key_URL = "" Then Return False
            Return True
        End Get
    End Property

    Public ReadOnly Property Has_CM_URL() As Boolean
        Get
            If IsNothing(CM_URL) Then Return False
            If CM_URL = "" Then Return False
            Return True
        End Get
    End Property

    Public ReadOnly Property MultiDayReport() As Boolean
        Get
            If EndDate = StartDate Then Return False
            Return True
        End Get
    End Property

    Public Property EndDate() As Date
        Get
            Return _EndDate
        End Get
        Set(ByVal value As Date)
            _EndDate = value
            If _StartDate > value Then _StartDate = value
        End Set
    End Property

    Public Property StartDate() As Date
        Get
            Return _StartDate
        End Get
        Set(ByVal value As Date)
            _StartDate = value
            If _EndDate < value Then _EndDate = value
        End Set
    End Property
End Class
#End Region 'Clients Data_WS_YahooX

#Region " Class Data_WS_ASK"
'==============================================================================================
'Class Data_WS_ASK - Common DataStructure used by ASK to store parameters and results from API.
'==============================================================================================
Public Class Data_WS_ASK

    Public AccountID As String
    Public AccountPWD As String

    Public ReportID As String
    Public URL As String
    Public Status As String

    Private _StartDate As Date
    Private _EndDate As Date

    Public DownloadReady As Boolean
    Public UrlRequested As Boolean
    Public Downloaded As Boolean
    Public Completed As Boolean
    Public NoData As Boolean
    '===========================================================================================
    'New()- Instantiate Data_WS class invoked from Clients.New()
    '===========================================================================================
    Sub New()
        AccountID = ""
        AccountPWD = ""
        ClearDownloadInfo()
    End Sub
    '===========================================================================================
    'ClearDownloadInfo()- Initialization for newly instantiated Data_WS class.  Invoked from New
    '===========================================================================================
    Public Sub ClearDownloadInfo()
        ReportID = ""
        URL = ""
        Status = ""

        StartDate = Today
        EndDate = Today

        DownloadReady = False
        UrlRequested = False
        Downloaded = False
        Completed = False
        NoData = False
    End Sub


    Public ReadOnly Property Has_AccountID() As Boolean
        Get
            If IsNothing(AccountID) Then Return False
            If AccountID = "" Then Return False
            Return True
        End Get
    End Property

    Public ReadOnly Property Has_AcctPwd() As Boolean
        Get
            If IsNothing(AccountPWD) Then Return False
            Return IIf(AccountPWD.Length > 0, True, False)
        End Get
    End Property

    Public ReadOnly Property Has_ReportID() As Boolean
        Get
            If IsNothing(ReportID) Then Return False
            If ReportID = "" Then Return False
            Return True
        End Get
    End Property

    Public ReadOnly Property Has_URL() As Boolean
        Get
            If IsNothing(URL) Then Return False
            If URL = "" Then Return False
            Return True
        End Get
    End Property

    Public ReadOnly Property MultiDayReport() As Boolean
        Get
            If EndDate = StartDate Then Return False
            Return True
        End Get
    End Property

    Public Property EndDate() As Date
        Get
            Return _EndDate
        End Get
        Set(ByVal value As Date)
            _EndDate = value
            If _StartDate > value Then _StartDate = value
        End Set
    End Property

    Public Property StartDate() As Date
        Get
            Return _StartDate
        End Get
        Set(ByVal value As Date)
            _StartDate = value
            If _EndDate < value Then _EndDate = value
        End Set
    End Property
End Class

#End Region ' Class Data_WS_ASK

#Region " Class Data_WS_GA"
'===========================================================================================
'Class Data_WS_GA - Common DataStructure used by Google Analytic and results from API.      
'===========================================================================================
Public Class Data_WS_GA

    Public AccountID As String
    Public AccountPWD As String

    Public ReportID As String
    Public URL As String
    Public Status As String

    Private _StartDate As Date
    Private _EndDate As Date

    Public DownloadReady As Boolean
    Public UrlRequested As Boolean
    Public Downloaded As Boolean
    Public Completed As Boolean
    Public NoData As Boolean
    Public EOMDeleted As Boolean
    '===========================================================================================
    'New()- Instantiate Data_WS class invoked from Clients.New()
    '===========================================================================================
    Sub New()
        AccountID = ""
        AccountPWD = ""
        ClearDownloadInfo()
    End Sub
    '===========================================================================================
    'ClearDownloadInfo()- Initialization for newly instantiated Data_WS class.  Invoked from New
    '===========================================================================================
    Public Sub ClearDownloadInfo()
        ReportID = ""
        URL = ""
        Status = ""

        StartDate = Today
        EndDate = Today

        DownloadReady = False
        UrlRequested = False
        Downloaded = False
        Completed = False
        NoData = False
        EOMDeleted = False

    End Sub


    Public ReadOnly Property Has_AccountID() As Boolean
        Get
            If IsNothing(AccountID) Then Return False
            If AccountID = "" Then Return False
            Return True
        End Get
    End Property

    Public ReadOnly Property Has_AcctPwd() As Boolean
        Get
            If IsNothing(AccountPWD) Then Return False
            Return IIf(AccountPWD.Length > 0, True, False)
        End Get
    End Property

    Public ReadOnly Property Has_ReportID() As Boolean
        Get
            If IsNothing(ReportID) Then Return False
            If ReportID = "" Then Return False
            Return True
        End Get
    End Property

    Public ReadOnly Property Has_URL() As Boolean
        Get
            If IsNothing(URL) Then Return False
            If URL = "" Then Return False
            Return True
        End Get
    End Property

    Public ReadOnly Property MultiDayReport() As Boolean
        Get
            If EndDate = StartDate Then Return False
            Return True
        End Get
    End Property

    Public Property EndDate() As Date
        Get
            Return _EndDate
        End Get
        Set(ByVal value As Date)
            _EndDate = value
            If _StartDate > value Then _StartDate = value
        End Set
    End Property

    Public Property StartDate() As Date
        Get
            Return _StartDate
        End Get
        Set(ByVal value As Date)
            _StartDate = value
            If _EndDate < value Then _EndDate = value
        End Set
    End Property
End Class

#End Region ' Class Data_WS_GA

#Region " Class Data_WW_Yahoo "
'No Longer used
Public Class Data_WS_Yahoo
    Inherits Data_WS

    Public SubAccountID As String = ""
    Public ContentMatch As Boolean = False

    Public ReadOnly Property Has_SubAccountID() As Boolean
        Get
            Return IIf(SubAccountID.Length > 0, True, False)
        End Get
    End Property
End Class

#End Region ' Class Data_WS_Yahoo 

#Region " clMyItem Class "
Public Class clMyItem
    Private sName As String
    Private nValue As Long

    Public Sub New(ByVal Name As String, ByVal Value As Long)
        sName = Name
        nValue = Value
    End Sub

    Public Function Name() As String
        Return sName
    End Function

    Public Function Value() As Long
        Return nValue
    End Function

    Public Overrides Function ToString() As String
        Return sName
    End Function
End Class
#End Region 'clMyItem Class
