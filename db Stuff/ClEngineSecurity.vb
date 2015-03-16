Imports System.Data.SqlClient
Public Class ClEngineSecurity
    Dim osettings As clSettings
    Dim sCnstr As String
    'GA OAuth 2.0
    Dim _GAOAuthID As String
    Dim _GASecret As String
    Dim _GARefreshToken As String

    Dim _GoogleUserAgent As String
    Dim _GoogleEmail As String
    Dim _GooglePassword As String
    Dim _GoogleToken As String
    Dim _GoogleAppToken As String
    'OAuth 2.0
    Dim _GoogleOAuthID As String
    Dim _GoogleSecret As String
    Dim _GoogleRefreshToken As String

    Dim _MSNUsername As String
    Dim _MSNPassword As String
    Dim _MSNAccessKey As String
    'OAuth 2.0
    Dim _MSNOAuthID As String
    Dim _MSNSecret As String
    Dim _MSNRefreshToken As String

    Dim _VerizonUserName As String
    Dim _VerizonPassword As String

    Dim _ASKLicense As String
    Dim _GAProfile As String

    Dim _YahooUserName As String
    Dim _YahooPassword As String
    Dim _YahooLicense As String
    '===================================================================================================================
    'New() - Instantiate Object. Engine authentication work.  Resturn the contents of the file SearchEngines and find   
    ' engines: Google, Bing and Google Analytics via searching SearchEngines.enginename for the strings "GOOGLE", "BING"
    ' and GOOGLE ANALYTICS.                                                                                             
    '===================================================================================================================
    Sub New()
        Dim cnn As SqlConnection
        Dim da As SqlDataAdapter
        Dim ds As New DataSet
        'Obtain settings.xml and grab connection string for use with SQL connections
        osettings = New clSettings()
        sCnstr = osettings.ConnectionString

        'Build Query with fields from SearchEngines table we require
        Dim ssql As String = " SELECT OAuthID, Secret, RefreshToken, UserID, Password, License, SecurityEmail, AppToken, EngineID, Enginename"
        ssql &= " FROM SearchEngines WHERE EngineID <= 5 OR EngineName = 'Google Analytics' ORDER BY Seq"

        'Query and return results via SqlDataAdapter and place into DataSet
        cnn = New SqlConnection(sCnstr)
        da = New SqlDataAdapter(ssql, cnn)
        da.Fill(ds)
        'Loop through returned table from Query and assign engine authentication information to properties
        Dim strEngine As String
        Dim drow As DataRow
        For Each drow In ds.Tables(0).Rows
            'Extract enginename and force to UpperCase
            strEngine = drow("EngineName").ToString().ToUpper
            'Based on enginename - assign to properties
            If strEngine.IndexOf("ANALYTICS") > -1 Then
                'GA Information for OAuth Authentication
                _GAOAuthID = drow("OAuthID")
                _GASecret = drow("Secret")
                _GARefreshToken = drow("RefreshToken")

            ElseIf strEngine.IndexOf("GOOGLE") > -1 Then
                'Google  Information for OAuth Authentication
                _GoogleUserAgent = drow("UserID")
                _GooglePassword = drow("Password")
                _GoogleToken = drow("License")
                _GoogleAppToken = drow("AppToken")
                _GoogleOAuthID = drow("OAuthID")
                _GoogleSecret = drow("Secret")
                _GoogleRefreshToken = drow("RefreshToken")

            ElseIf strEngine.IndexOf("BING") > -1 Then
                'BING Information for OAuth Authentication
                _MSNUsername = drow("UserID")
                _MSNPassword = drow("Password")
                _MSNAccessKey = drow("License")
                _MSNOAuthID = drow("OAuthID")
                _MSNSecret = drow("Secret")
                _MSNRefreshToken = drow("RefreshToken")

                'ElseIf strEngine.IndexOf("YAHOO") > -1 Then
                '    'YAHOO Information for Authentication
                '    _YahooUserName = drow("UserID")
                '    _YahooPassword = drow("Password")
                '    _YahooLicense = drow("License")

                'ElseIf strEngine.IndexOf("SUPERPAGES") > -1 Then
                '    'VERIZON Information for Authentication
                '    _VerizonUserName = drow("UserID")
                '    _VerizonPassword = drow("Password")

                'ElseIf strEngine.IndexOf("ASK") > -1 Then
                '    'ASK Information for Authentication
                '    _ASKLicense = drow("License")
            End If

        Next

        ''Verizon
        'drow = ds.Tables(0).Rows(3)
        '_VerizonUserName = drow("UserID")
        '_VerizonPassword = drow("Password")

        ''ASK
        'drow = ds.Tables(0).Rows(4)
        '_ASKLicense = drow("License")
        'Dim drow As DataRow = ds.Tables(0).Rows(0)
        ''Google
        '_GoogleUserAgent = drow("UserID")
        '_GooglePassword = drow("Password")
        '_GoogleEmail = drow("SecurityEmail")
        '_GoogleToken = drow("License")
        '_GoogleAppToken = drow("AppToken")

        ''Yahoo
        'drow = ds.Tables(0).Rows(1)
        '_YahooUserName = drow("UserID")
        '_YahooPassword = drow("Password")
        '_YahooLicense = drow("License")


        ''MSN
        'drow = ds.Tables(0).Rows(2)
        '_MSNUsername = drow("UserID")
        '_MSNPassword = drow("Password")
        '_MSNAccessKey = drow("License")

    End Sub
    'GA Authentication Properties
    Public Property GAOAuthID() As String
        Get
            Return _GAOAuthID
        End Get
        Set(ByVal value As String)
            _GAOAuthID = value
        End Set
    End Property
    Public Property GASecret() As String
        Get
            Return _GASecret
        End Get
        Set(ByVal value As String)
            _GASecret = value
        End Set
    End Property
    Public Property GARefreshToken() As String
        Get
            Return _GARefreshToken
        End Get
        Set(ByVal value As String)
            _GARefreshToken = value
        End Set
    End Property
    'GA Authentication Properties
    Public Property GoogleOAuthID() As String
        Get
            Return _GoogleOAuthID
        End Get
        Set(ByVal value As String)
            _GoogleOAuthID = value
        End Set
    End Property
    Public Property GoogleSecret() As String
        Get
            Return _GoogleSecret
        End Get
        Set(ByVal value As String)
            _GoogleSecret = value
        End Set
    End Property
    Public Property GoogleRefreshToken() As String
        Get
            Return _GoogleRefreshToken
        End Get
        Set(ByVal value As String)
            _GoogleRefreshToken = value
        End Set
    End Property

    Public Property GoogleUserAgent() As String
        Get
            Return _GoogleUserAgent
        End Get
        Set(ByVal value As String)
            _GoogleUserAgent = value
        End Set
    End Property
    Public Property GoogleEmail() As String
        Get
            Return _GoogleEmail
        End Get
        Set(ByVal value As String)
            _GoogleEmail = value
        End Set
    End Property
    Public Property GooglePassword() As String
        Get
            Return _GooglePassword
        End Get
        Set(ByVal value As String)
            _GooglePassword = value
        End Set
    End Property
    Public Property GoogleToken() As String
        Get
            Return _GoogleToken
        End Get
        Set(ByVal value As String)
            _GoogleToken = value
        End Set
    End Property
    Public Property GoogleAppToken() As String
        Get
            Return _GoogleAppToken
        End Get
        Set(ByVal value As String)
            _GoogleAppToken = value
        End Set
    End Property
    'MSN Authentication Properties
    Public Property MSNOAuthID() As String
        Get
            Return _MSNOAuthID
        End Get
        Set(ByVal value As String)
            _MSNOAuthID = value
        End Set
    End Property
    Public Property MSNSecret() As String
        Get
            Return _MSNSecret
        End Get
        Set(ByVal value As String)
            _MSNSecret = value
        End Set
    End Property
    Public Property MSNRefreshToken() As String
        Get
            Return _MSNRefreshToken
        End Get
        Set(ByVal value As String)
            _MSNRefreshToken = value
        End Set
    End Property
    Public Property MSNUserName() As String
        Get
            Return _MSNUserName
        End Get
        Set(ByVal value As String)
            _MSNUserName = value
        End Set
    End Property
    Public Property MSNPassword() As String
        Get
            Return _MSNPassword
        End Get
        Set(ByVal value As String)
            _MSNPassword = value
        End Set
    End Property
    Public Property MSNAccessKey() As String
        Get
            Return _MSNAccessKey
        End Get
        Set(ByVal value As String)
            _MSNAccessKey = value
        End Set
    End Property
    'YAHOO Authentication Properties
    Public Property YahooUserName() As String
        Get
            Return _YahooUserName
        End Get
        Set(ByVal value As String)
            _YahooUserName = value
        End Set
    End Property
    Public Property YahooPassword() As String
        Get
            Return _YahooPassword
        End Get
        Set(ByVal value As String)
            _YahooUserName = value
        End Set
    End Property
    Public Property YahooLicense() As String
        Get
            Return _YahooLicense
        End Get
        Set(ByVal value As String)
            _YahooLicense = value
        End Set
    End Property
    'VERIZON Authentication Properties
    Public Property VerizonUserName() As String
        Get
            Return _VerizonUserName
        End Get
        Set(ByVal value As String)
            _VerizonUserName = value
        End Set
    End Property
    Public Property VerizonPassword() As String
        Get
            Return _VerizonPassword
        End Get
        Set(ByVal value As String)
            _VerizonPassword = value
        End Set
    End Property
    'ASK Authentication Properties
    Public Property ASKLicense() As String
        Get
            Return _ASKLicense
        End Get
        Set(ByVal value As String)
            _ASKLicense = value
        End Set
    End Property

End Class
