Imports System.Xml
'=================================================================================================
'Class clSettings will be common Data Structure used throughout this project to pass search engine
'authentication information.
'=================================================================================================
#Region " clSettings Class "
Public Class clSettings
    Private Const _FileName As String = "settings.xml"
    Private _SaveBaseLocation As String
    Private _ConnectionString As String
    Private _DirectoryNamesAsRecID As Boolean = False

    Private _GAUser As String
    Private _GAPwd As String
    Private _GAClientID As String
    Private _GAClientSecret As String
    Private _GARefreshToken As String

    'OAuth2 for AdWords
    Private _GoogleAdWordsClientID As String
    Private _GoogleAdWordsClientSecret As String
    Private _GoogleAdWordsRefreshToken As String

    Public _Mail_To_Email(0) As String
    Public _Mail_CC_Email(0) As String
    Public _Mail_From_Email As String = ""
    Public _Mail_From_Name As String = ""
    Public _Mail_Server As String = ""

    Public Yahoo As clYahoo
    Public Google As clGoogle
    Public MSN As clMSN
    Public Verizon As clVerizon
    Public ASK As clASK
    Public GA As clGA

    Public ReadClientsFromXML As Boolean = False
    Public FixButton As Boolean = False
    Public YahooOnlyProcessFromEndDate As Boolean = True

    Public Clients() As clClient

    Public YahooColumnsKeyword() As clColumns
    Public YahooColumnsAdGroup() As clColumns
    Public MSNColumns() As clColumns
    Public GoogleColumns() As clColumns
    Public VerizonColumns() As clColumns
    Public ASKColumns() As clColumns
    Public GAColumns() As clColumns
    'GA Visitor Data
    Public GAVColumns() As clColumns
    Public ColumnTypesV() As clColumns

    Public ColumnTypes() As clColumns

    Public ReadOnly Property Mail_Server() As String
        Get
            Return _Mail_Server
        End Get
    End Property

    Public ReadOnly Property Mail_From_Email() As String
        Get
            Return _Mail_From_Email
        End Get
    End Property

    Public ReadOnly Property Mail_From_Name() As String
        Get
            Return _Mail_From_Name
        End Get
    End Property

    Public ReadOnly Property Mail_To_Email() As String()
        Get
            Return _Mail_To_Email
        End Get
    End Property

    Public ReadOnly Property Mail_CC_Email() As String()
        Get
            Return _Mail_CC_Email
        End Get
    End Property

    Public ReadOnly Property DirectoryNamesAsRecID() As Boolean
        Get
            Return _DirectoryNamesAsRecID
        End Get
    End Property

    Public ReadOnly Property FileName() As String
        Get
            Dim sPath As String = Application.StartupPath
            If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
            Return sPath & _FileName
        End Get
    End Property

    Public ReadOnly Property SaveBaseLocation() As String
        Get
            Return _SaveBaseLocation
        End Get
    End Property
    Public ReadOnly Property GAPwd() As String
        Get
            Return _GAPwd
        End Get
    End Property
    Public ReadOnly Property GAUser() As String
        Get
            Return _GAUser
        End Get
    End Property
    Public ReadOnly Property GAClientID() As String
        Get
            Return _GAClientID
        End Get
    End Property
    Public ReadOnly Property GAClientSecret() As String
        Get
            Return _GAClientSecret
        End Get
    End Property
    Public ReadOnly Property GARefreshToken() As String
        Get
            Return _GARefreshToken
        End Get
    End Property
    Public ReadOnly Property GoogleAdWordsRefreshToken() As String
        Get
            Return _GoogleAdWordsRefreshToken
        End Get
    End Property

    Public ReadOnly Property GoogleAdWordsClientID() As String
        Get
            Return _GoogleAdWordsClientID
        End Get
    End Property
    Public ReadOnly Property GoogleAdWordsClientSecret() As String
        Get
            Return _GoogleAdWordsClientSecret
        End Get
    End Property


    Public ReadOnly Property ConnectionString() As String
        Get
            Return _ConnectionString
        End Get
    End Property
    '===============================================================================================
    'New() Invoked in frmMain.frmMain_Load()
    'Process FileName(settings.xml), using name/value pair processing to populate the variables
    'for this class clSettings.
    '===============================================================================================
    Sub New()
        Dim reader As New XmlTextReader(FileName)

        _Mail_To_Email(0) = ""
        _Mail_CC_Email(0) = ""
        With reader
            .WhitespaceHandling = WhitespaceHandling.None
            While .Read
                Select Case .Name.ToLower
                    Case "readclientsfromxml"
                        Dim sResult As String = .GetAttribute("value")
                        If IsNothing(sResult) Then
                            Me.ReadClientsFromXML = False
                        Else
                            Me.ReadClientsFromXML = IIf(sResult.ToLower = "true", True, False)
                        End If

                    Case "gauser"
                        _GAUser = .GetAttribute("value")

                    Case "gapwd"
                        _GAPwd = .GetAttribute("value")

                    Case "garefreshtoken"
                        _GARefreshToken = .GetAttribute("value")
                    Case "gaclientsecret"
                        _GAClientSecret = .GetAttribute("value")
                    Case "gaclientid"
                        _GAClientID = .GetAttribute("value")

                        'OAuth2 for Adwords
                    Case "googleadwordsrefreshtoken"
                        _GoogleAdWordsRefreshToken = .GetAttribute("value")
                    Case "googleadwordsclientsecret"
                        _GoogleAdWordsClientSecret = .GetAttribute("value")
                    Case "googleadwordsclientid"
                        _GoogleAdWordsClientID = .GetAttribute("value")

                    Case "mail_from_email"
                        _Mail_From_Email = .GetAttribute("value")

                    Case "mail_from_name"
                        _Mail_From_Name = .GetAttribute("value")

                    Case "mail_to_email"
                        Dim Index As Integer = _Mail_To_Email.GetUpperBound(0) + 1
                        ReDim Preserve _Mail_To_Email(Index)
                        _Mail_To_Email(Index) = .GetAttribute("value")

                    Case "mail_cc_email"
                        Dim Index As Integer = _Mail_CC_Email.GetUpperBound(0) + 1
                        ReDim Preserve _Mail_CC_Email(Index)
                        _Mail_CC_Email(Index) = .GetAttribute("value")

                    Case "mail_server"
                        _Mail_Server = .GetAttribute("value")

                    Case "directorynamesasrecid"
                        Dim sResult As String = .GetAttribute("value")
                        If IsNothing(sResult) Then
                            Me._DirectoryNamesAsRecID = False
                        Else
                            Me._DirectoryNamesAsRecID = IIf(sResult.ToLower = "true", True, False)
                        End If

                    Case "fixbutton"
                        Dim sResult As String = .GetAttribute("value")
                        If IsNothing(sResult) Then
                            Me.FixButton = False
                        Else
                            Me.FixButton = IIf(sResult.ToLower = "true", True, False)
                        End If

                    Case "yahooonlyprocessenddate"
                        Dim sResult As String = .GetAttribute("value")
                        If IsNothing(sResult) Then
                            Me.YahooOnlyProcessFromEndDate = False
                        Else
                            Me.YahooOnlyProcessFromEndDate = IIf(sResult.ToLower = "true", True, False)
                        End If

                    Case "verizonsecurity"
                        Verizon = New clVerizon(.GetAttribute("UserName"), _
                                                .GetAttribute("Password"))

                    Case "yahoosecurity"
                        Yahoo = New clYahoo(.GetAttribute("UserName"), _
                                            .GetAttribute("Password"), _
                                            .GetAttribute("License"))

                    Case "googlesecurity"
                        Google = New clGoogle(.GetAttribute("UserAgent"), _
                                              .GetAttribute("Email"), _
                                              .GetAttribute("Password"), _
                                              .GetAttribute("Token"), _
                                              .GetAttribute("AppToken"))

                    Case "msnsecurity"
                        MSN = New clMSN(.GetAttribute("UserName"), _
                                        .GetAttribute("Password"), _
                                        .GetAttribute("UserAccessKey"))

                    Case "clients"
                        If .NodeType = XmlNodeType.Element Then ReDim Clients(0)

                    Case "client"
                        Dim Client As New clClient(.GetAttribute("name"), _
                                                   .GetAttribute("CustID"), _
                                                   .GetAttribute("YahooAccountID"), _
                                                   .GetAttribute("YahooSubAccount"), _
                                                   .GetAttribute("GoogleLogin"), _
                                                   .GetAttribute("GooglePassword"), _
                                                   .GetAttribute("MSNAccountID"), _
                                                   .GetAttribute("VerizonAccountID"), _
                                                   .GetAttribute("ASKAccountID"), _
                                                   .GetAttribute("GAAccountID"))

                        Dim index = Clients.GetUpperBound(0) + 1
                        ReDim Preserve Clients(index)
                        Clients(index) = Client

                    Case "savebaselocation"
                        _SaveBaseLocation = .GetAttribute("value")
                        If _SaveBaseLocation.Length > 0 Then
                            If Right(_SaveBaseLocation, 1) <> "\" Then _SaveBaseLocation = _SaveBaseLocation & "\"
                        End If

                    Case "yahoocolumnskeyword"
                        If .NodeType = XmlNodeType.Element Then ReDim Me.YahooColumnsKeyword(0)

                    Case "yahoocolumnkeyword"
                        Dim Column As New clColumns(.GetAttribute("name"), .GetAttribute("value"))
                        Dim index = YahooColumnsKeyword.GetUpperBound(0) + 1
                        ReDim Preserve YahooColumnsKeyword(index)
                        YahooColumnsKeyword(index) = Column

                    Case "yahoocolumnsadgroup"
                        If .NodeType = XmlNodeType.Element Then ReDim Me.YahooColumnsAdGroup(0)

                    Case "yahoocolumnadgroup"
                        Dim Column As New clColumns(.GetAttribute("name"), .GetAttribute("value"))
                        Dim index = YahooColumnsAdGroup.GetUpperBound(0) + 1
                        ReDim Preserve YahooColumnsAdGroup(index)
                        YahooColumnsAdGroup(index) = Column

                    Case "verizoncolumns"
                        If .NodeType = XmlNodeType.Element Then ReDim Me.VerizonColumns(0)

                    Case "verizoncolumn"
                        Dim Column As New clColumns(.GetAttribute("name"), .GetAttribute("value"))
                        Dim index = VerizonColumns.GetUpperBound(0) + 1
                        ReDim Preserve VerizonColumns(index)
                        VerizonColumns(index) = Column

                    Case "msncolumns"
                        If .NodeType = XmlNodeType.Element Then ReDim Me.MSNColumns(0)

                    Case "msncolumn"
                        Dim Column As New clColumns(.GetAttribute("name"), .GetAttribute("value"))
                        Dim index = MSNColumns.GetUpperBound(0) + 1
                        ReDim Preserve MSNColumns(index)
                        MSNColumns(index) = Column

                    Case "googlecolumns"
                        If .NodeType = XmlNodeType.Element Then ReDim Me.GoogleColumns(0)

                    Case "googlecolumn"
                        Dim Column As New clColumns(.GetAttribute("name"), .GetAttribute("value"))
                        Dim index = GoogleColumns.GetUpperBound(0) + 1
                        ReDim Preserve GoogleColumns(index)
                        GoogleColumns(index) = Column

                    Case "askcolumns"
                        If .NodeType = XmlNodeType.Element Then ReDim Me.ASKColumns(0)

                    Case "askcolumn"
                        Dim Column As New clColumns(.GetAttribute("name"), .GetAttribute("value"))
                        Dim index = ASKColumns.GetUpperBound(0) + 1
                        ReDim Preserve ASKColumns(index)
                        ASKColumns(index) = Column

                    Case "gacolumns"
                        If .NodeType = XmlNodeType.Element Then ReDim Me.GAColumns(0)

                    Case "gacolumn"
                        Dim Column As New clColumns(.GetAttribute("name"), .GetAttribute("value"))
                        Dim index = GAColumns.GetUpperBound(0) + 1
                        ReDim Preserve GAColumns(index)
                        GAColumns(index) = Column

                        'Tire Choice Specific
                    Case "gavcolumns"
                        If .NodeType = XmlNodeType.Element Then ReDim Me.GAVColumns(0)

                    Case "gavcolumn"
                        Dim Column As New clColumns(.GetAttribute("name"), .GetAttribute("value"))
                        Dim index = GAVColumns.GetUpperBound(0) + 1
                        ReDim Preserve GAVColumns(index)
                        GAVColumns(index) = Column

                    Case "columntypes"
                        If .NodeType = XmlNodeType.Element Then ReDim Me.ColumnTypes(0)

                    Case "columntype"
                        Dim Column As New clColumns(.GetAttribute("key").ToLower, .GetAttribute("value"))
                        Dim index = ColumnTypes.GetUpperBound(0) + 1
                        ReDim Preserve ColumnTypes(index)
                        ColumnTypes(index) = Column

                        'GA Visitor Data
                    Case "columntypesv"
                        If .NodeType = XmlNodeType.Element Then ReDim Me.ColumnTypesV(0)

                    Case "columntypev"
                        Dim Column As New clColumns(.GetAttribute("key").ToLower, .GetAttribute("value"))
                        Dim index = ColumnTypesV.GetUpperBound(0) + 1
                        ReDim Preserve ColumnTypesV(index)
                        ColumnTypesV(index) = Column

                    Case "connectionstring"
                        _ConnectionString = .GetAttribute("value")

                    Case Else
                End Select
            End While
            .Close()
        End With
    End Sub

End Class
#End Region 'clSettings Class

#Region " Verizon Class "
'===============================================================================================
'Class clVerizon will be common Data Structure used throughout this project to pass Verizon
'search engine authentication information.
'===============================================================================================
Public Class clVerizon
    Private _UserName As String
    Private _Password As String

    Public ReadOnly Property Username() As String
        Get
            Return _UserName
        End Get
    End Property

    Public ReadOnly Property Password() As String
        Get
            Return _Password
        End Get
    End Property

    Sub New(ByVal sUserName As String, ByVal sPassword As String)
        _UserName = sUserName
        _Password = sPassword
    End Sub
End Class
#End Region 'Verizon Class

#Region " Yahoo Class "
'===============================================================================================
'Class clYahoo will be common Data Structure used throughout this project to pass Yahoo
'search engine authentication information.
'===============================================================================================
Public Class clYahoo
    Private _UserName As String
    Private _Password As String
    Private _License As String

    Public ReadOnly Property Username() As String
        Get
            Return _UserName
        End Get
    End Property
    Public ReadOnly Property Password() As String
        Get
            Return _Password
        End Get
    End Property

    Public ReadOnly Property License() As String
        Get
            Return _License
        End Get
    End Property

    Sub New(ByVal sUserName As String, ByVal sPassword As String, ByVal sLicense As String)
        _UserName = sUserName
        _Password = sPassword
        _License = sLicense
    End Sub
End Class
#End Region 'Yahoo Class

#Region " MSN Class "
'===============================================================================================
'Class clMSN will be common Data Structure used throughout this project to pass MSN
'search engine authentication information.
'===============================================================================================
Public Class clMSN
    Private _UserName As String
    Private _Password As String
    Private _UserAccessKey As String

    Public ReadOnly Property UserName()
        Get
            Return _UserName
        End Get
    End Property

    Public ReadOnly Property Password()
        Get
            Return _Password
        End Get
    End Property

    Public ReadOnly Property UserAccessKey() As String
        Get
            Return _UserAccessKey
        End Get
    End Property

    Sub New(ByVal sUserName As String, ByVal sPassword As String, ByVal sUserAccessKey As String)
        _UserName = sUserName
        _Password = sPassword
        _UserAccessKey = sUserAccessKey
    End Sub
End Class
#End Region 'MSN Class

#Region " GA Class "
'===============================================================================================
'Class clGA will be common Data Structure used throughout this project to pass GA
'search engine authentication information.
'===============================================================================================
Public Class clGA
    Private _UserName As String
    Private _Password As String

    Public ReadOnly Property Username() As String
        Get
            Return _UserName
        End Get
    End Property

    Public ReadOnly Property Password() As String
        Get
            Return _Password
        End Get
    End Property

    Sub New(ByVal sUserName As String, ByVal sPassword As String)
        _UserName = sUserName
        _Password = sPassword
    End Sub
End Class
#End Region 'GA Class

#Region " ASK Class "
'===============================================================================================
'Class clASK will be common Data Structure used throughout this project to pass ASK
'search engine authentication information.
'===============================================================================================
Public Class clASK
    Private _UserName As String
    Private _Password As String

    Public ReadOnly Property Username() As String
        Get
            Return _UserName
        End Get
    End Property

    Public ReadOnly Property Password() As String
        Get
            Return _Password
        End Get
    End Property

    Sub New(ByVal sUserName As String, ByVal sPassword As String)
        _UserName = sUserName
        _Password = sPassword
    End Sub
End Class
#End Region 'ASK Class

#Region " Google Class "
'===============================================================================================
'Class clGoogle will be common Data Structure used throughout this project to pass Google
'search engine authentication information.
'===============================================================================================
Public Class clGoogle
    Private _UserAgent As String
    Private _Email As String
    Private _Password As String
    Private _Token As String
    Private _AppToken As String

    Public ReadOnly Property UserAgent()
        Get
            Return _UserAgent
        End Get
    End Property

    Public ReadOnly Property Email()
        Get
            Return _Email
        End Get
    End Property

    Public ReadOnly Property Password()
        Get
            Return _Password
        End Get
    End Property

    Public ReadOnly Property Token()
        Get
            Return _Token
        End Get
    End Property

    Public ReadOnly Property AppToken()
        Get
            Return _AppToken
        End Get
    End Property

    Sub New(ByVal sUserAgent As String, ByVal sEmail As String, ByVal sPassword As String, ByVal sToken As String, ByVal sAppToken As String)
        _UserAgent = sUserAgent
        _Email = sEmail
        _Password = sPassword
        _Token = sToken
        _AppToken = sAppToken
    End Sub
End Class
#End Region 'Google Class

#Region " Client Class "
'===============================================================================================
'Class clClient will be common Data Structure used throughout this project to pass search engine
'authentication information.
'===============================================================================================
Public Class clClient
    Private _Username As String
    Private _GoogleID As String
    Private _MSNID As String
    Private _YahooID As String
    Private _YahooSubID As String
    Private _CustID As String
    Private _GooglePassword As String
    Private _VerizonID As String
    Private _ASKID As String
    Private _GAID As String

    Public ReadOnly Property GooglePassword() As String
        Get
            Return _GooglePassword
        End Get
    End Property

    Public ReadOnly Property CustID() As String
        Get
            Return _CustID
        End Get
    End Property

    Public ReadOnly Property UserName() As String
        Get
            Return _Username
        End Get
    End Property

    Public ReadOnly Property GoogleID() As String
        Get
            Return _GoogleID
        End Get
    End Property

    Public ReadOnly Property MSNID() As String
        Get
            Return _MSNID
        End Get
    End Property

    Public ReadOnly Property VerizonID() As String
        Get
            Return _VerizonID
        End Get
    End Property

    Public ReadOnly Property ASKID() As String
        Get
            Return _ASKID
        End Get
    End Property
    Public ReadOnly Property GAID() As String
        Get
            Return _GAID
        End Get
    End Property

    Public ReadOnly Property YahooID() As String
        Get
            Return _YahooID
        End Get
    End Property

    Public ReadOnly Property YahooSubID() As String
        Get
            Return _YahooSubID
        End Get
    End Property
    '===============================================================================================
    'New() Invoked in clSettings.New() when processing clients element.
    '===============================================================================================
    Sub New(ByVal sUserName As String, ByVal sCustID As String, ByVal sYahooID As String, ByVal sYahooSubID As String, ByVal sGoogleID As String, ByVal sGooglePwd As String, ByVal sMSNID As String, ByVal sVerizonID As String, ByVal sASKID As String, ByVal sGAID As String)
        _Username = sUserName
        _GoogleID = sGoogleID
        _MSNID = sMSNID
        _VerizonID = sVerizonID
        _ASKID = sASKID
        _GAID = sGAID
        _YahooID = sYahooID
        _YahooSubID = sYahooSubID
        _CustID = sCustID
        _GooglePassword = sGooglePwd
    End Sub
End Class
#End Region 'Client Class

#Region " ColumnNames Class "
Public Class clColumns
    Private _Name As String
    Private _Value As String

    Public ReadOnly Property Name() As String
        Get
            Return _Name
        End Get
    End Property

    Public ReadOnly Property Value() As String
        Get
            Return _Value
        End Get
    End Property

    Sub New(ByVal sName As String, ByVal sValue As String)
        _Name = sName
        _Value = sValue
    End Sub
End Class
#End Region 'ColumnNames Class