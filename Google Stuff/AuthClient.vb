Imports System
Imports System.Collections
Imports System.IO
Imports System.Net
Imports System.Reflection
Imports System.Text
Imports System.Web
Imports System.Web.Script.Serialization


Public Class AuthClient
    '
    ' OAuth2 URI
    '
    Private Const authOAFormat As String = "https://accounts.google.com/o/oauth2/token"
    Private Const authOARefreshToken As String = "refresh_token={0}&client_id={1}&client_secret={2}&grant_type=refresh_token"
    'Private Const authOABodyFormat As String = "code=4%2FcbH7EDcOLJbPRPAIuk8q_jotqVbu.4rWYve_iTXIYshQV0ieZDArw3U-AdAI&redirect_uri=urn:ietf:wg:oauth:2.0:oob&scope=https%3A%2F%2Fadwords.google.com%2Fapi%2Fadwords%2F&client_secret=vd8y42vwITiIisRmFidk7_XE&grant_type=authorization_code&client_id=151428027265.apps.googleusercontent.com"
    Public ClientID As String
    Public ClientSecret As String
    Public RefreshToken As String
    Public AccessToken As String
    Public tokenType As String
    Public GoogleMCC As String
    Private Settings As New clSettings
    ' 
    'Email to be used for authentication.
    ' 
    Dim email As String

    ' 
    'Password to be used for authentication.
    ' 
    Dim password As String

    ' 
    'Url endpoint for ClientLogin API.
    ' 
    Const URL As String = "https://www.google.com/accounts/ClientLogin"

    ' 
    'Account type to be used with ClientLogin API.
    ' 
    Const ACCOUNT_TYPE As String = "GOOGLE"

    ' 
    'Service type to be used with ClientLogin API.
    ' 
    Const SERVICE As String = "adwords"

    ' 
    'The source identifier string to be used with ClientLogin API.
    ' 
    Const SOURCE As String = "MV-AWAPI-DotNetLib-V201402"
    ''' <summary>
    ''' The time at which access token was updated.
    ''' </summary>
    Private m_updatedOn As DateTime = DateTime.MinValue

    ''' <summary>
    ''' The remaining lifetime on the access token.
    ''' </summary>
    Private m_expiresIn As Integer

    ''' <summary>
    ''' The OAuth2 tokens will be refreshed automatically if the time left for
    ''' access token expiry is less than this value in seconds.
    ''' </summary>
    Private oAuth2RefreshCutoffLimit As Integer = 60

    ''' <summary>
    ''' Gets or sets the time at which access token was retrieved.
    ''' </summary>
    Public Property UpdatedOn() As DateTime
        Get
            Return m_updatedOn
        End Get
        Set(value As DateTime)
            m_updatedOn = value
        End Set
    End Property
    ''' <summary>
    ''' Gets the remaining lifetime on the access token.
    ''' </summary>
    Public Property ExpiresIn() As Integer
        Get
            Return m_expiresIn
        End Get
        Set(value As Integer)
            m_expiresIn = value
        End Set
    End Property

    Public Sub New(ByVal engineAuthentication As ClEngineSecurity, ByVal bGA As Boolean)
        ClientID = engineAuthentication.GAOAuthID
        ClientSecret = engineAuthentication.GASecret
        RefreshToken = engineAuthentication.GARefreshToken
    End Sub
    Public Sub New(ByVal engineAuthentication As ClEngineSecurity)
        ClientID = engineAuthentication.GoogleOAuthID
        ClientSecret = engineAuthentication.GoogleSecret
        RefreshToken = engineAuthentication.GoogleRefreshToken
    End Sub
    Public Sub New(ByVal oSettings As clSettings)
        ClientID = oSettings.GAClientID
        ClientSecret = oSettings.GAClientSecret
        RefreshToken = oSettings.GARefreshToken
    End Sub
    Public Function GetAuthHeader() As String
        RefreshAccessTokenIfExpiring()
        Return Convert.ToString(Me.AccessToken)
    End Function
    Public Sub RefreshAccessTokenIfExpiring()
        If IsAccessTokenExpiring() Then
            Me.AccessToken = GetAccessToken()
        End If
    End Sub

    Private Function IsAccessTokenExpiring() As Boolean
        If Me.UpdatedOn = DateTime.MinValue Then
            Return True
        End If
        Return Me.UpdatedOn + New TimeSpan(0, 0, Me.ExpiresIn - oAuth2RefreshCutoffLimit) < DateTime.Now
    End Function
   
    '===================================================================================================================================
    'GetAccessToken() - Google and GA using to perform OAuth refresh token work                                                         
    '                                                                                                                                   
    '===================================================================================================================================
    Public Function GetAccessToken() As String
        Dim strJSON As String
        Dim Tokens() As String
        Dim item As String
        Dim tbOutput As StringBuilder
        tbOutput = New StringBuilder()

        Dim indentLevel As Integer = 0
        'Once we complete initial authorization we use authOARefreshToken and not authOABodyFormat
        Dim authBody As String = String.Format(authOARefreshToken, RefreshToken, ClientID, ClientSecret)

        'Dim authBody As String = authOABodyFormat
        Dim req As HttpWebRequest
        Dim response As HttpWebResponse

        Dim ser As JavaScriptSerializer

        Dim stream As Stream
        Dim sw As StreamWriter
        Dim sr As StreamReader

        'OAuth2.0
        'End point for all Token work
        'Via Refresh Token, we get and Access Token and use for one hour.
        req = HttpWebRequest.Create(authOAFormat)

        req.Method = "POST"
        req.ContentType = "application/x-www-form-urlencoded"
        req.UserAgent = "GoogleAdWordsReportInterface"

        stream = req.GetRequestStream()
        sw = New StreamWriter(stream)
        sw.Write(authBody)
        sw.Close()
        sw.Dispose()

        response = req.GetResponse()
        sr = New StreamReader(response.GetResponseStream())
        strJSON = sr.ReadToEnd()
        Dim values As Dictionary(Of String, String) = ParseJsonObjectResponse(strJSON)
        If values.ContainsKey("access_token") Then
            Me.AccessToken = values("access_token")
        End If
        If values.ContainsKey("refresh_token") Then
            Me.RefreshToken = values("refresh_token")
        End If
        If values.ContainsKey("token_type") Then
            Me.tokenType = values("token_type")
        End If
        If values.ContainsKey("expires_in") Then
            Me.m_expiresIn = Integer.Parse(values("expires_in"))
        End If

        Me.m_updatedOn = DateTime.Now
        Return Me.AccessToken

    End Function
    '==================================================================================================================================
    'ParseJsonObjectResponse() - New to V201309 - Request parsing                                                                               
    '                  Returns Hashtable - with value pairs from request                                                              
    '=================================================================================================================================
    ''' <summary>
    ''' Parses a json object response.
    ''' </summary>
    ''' <param name="contents">The json contents.</param>
    ''' <returns>A dictionary of key-value pairs.</returns>
    ''' <remarks>This is not a full-blown json parser, it can handle only cases
    ''' where the response is a json object without nested objects or arrays.
    ''' </remarks>
    Protected Function ParseJsonObjectResponse(contents As String) As Dictionary(Of String, String)
        Dim retval As New Dictionary(Of String, String)()
        Dim splits As String() = contents.Trim(New Char() {" "c, "{"c, "}"c}).Split(","c)
        For Each split As String In splits
            Dim subSplits As String() = split.Trim(New Char() {" "c, ControlChars.Cr, ControlChars.Lf}).Split(":"c)
            If subSplits.Length = 2 Then
                retval.Add(subSplits(0).Trim(New Char() {""""c, " "c}), subSplits(1).Trim(New Char() {""""c, " "c}))
            End If
        Next
        Return retval
    End Function

    '==================================================================================================================================
    'ParseResponse() - New to V200909 - Request parsing                                                                               
    '                  Returns Hashtable - with value pairs from request                                                              
    '=================================================================================================================================

    Function ParseResponse(ByVal response As WebResponse) As Hashtable
        Dim retVal As New Hashtable()
        Dim responseStream As Stream = response.GetResponseStream()
        Dim reader As New StreamReader(responseStream)

        Dim sResponse As String = reader.ReadToEnd()
        Dim splits As String() = sResponse.Split(ControlChars.Lf)
        For Each split As String In splits
            Dim subsplits As String() = split.Split("="c)
            If subsplits.Length >= 2 Then
                If Not String.IsNullOrEmpty(subsplits(0)) Then
                    If Not retVal.ContainsKey(subsplits(0)) Then
                        retVal.Add(subsplits(0), split.Substring(subsplits(0).Length + 1))
                    End If
                End If
            End If
        Next
        Return retVal
    End Function

    '=================================================================================================================================
    ''' <summary>
    ''' Extracts a ClientLogin failure and wraps it into
    ''' an AuthTokenException.
    ''' </summary>
    ''' <param name="ex">The exception originally thrown by webrequest
    ''' to ClientLogin endpoint.</param>
    ''' <returns></returns>
    '''   '=================================================================================================================================
    Private Function ExtractException(ByVal ex As WebException) As AuthTokenException
        Dim tblResponse As Hashtable = Nothing
        Dim response As WebResponse = Nothing

        response = ex.Response
        tblResponse = ParseResponse(response)

        Dim url As String = ""
        Dim [error] As String = ""
        Dim captchaToken As String = ""
        Dim captchaUrl As String = ""

        For Each key As String In tblResponse.Keys
            Select Case key
                Case "Url"
                    url = DirectCast(tblResponse(key), String)
                    Exit Select

                Case "Error"
                    [error] = DirectCast(tblResponse(key), String)
                    Exit Select

                Case "CaptchaToken"
                    captchaToken = DirectCast(tblResponse(key), String)
                    Exit Select

                Case "CaptchaUrl"
                    captchaUrl = DirectCast(tblResponse(key), String)
                    Exit Select
            End Select
        Next

        Dim errCode As AuthTokenErrorCode = AuthTokenErrorCode.Unknown

        Try
            errCode = DirectCast([Enum].Parse(GetType(AuthTokenErrorCode), [error], True), AuthTokenErrorCode)
            ' Enum does not have a tryParse.
        Catch generatedExceptionName As ArgumentException
        End Try

        Return New AuthTokenException(errCode, url, captchaToken, captchaUrl, "Login failed", ex)
    End Function
    '=============================================================================================================================
    'GetToken() - New to V200909 - RequestHeader.authtoken which must be generated via https://www.google.com/accounts/ClientLogin
    '             and parameters of email and password                                                                            
    '             Returns String - authToken                                                                                      
    '=============================================================================================================================
    'Public Function GetToken() As String
    '    Dim webRequest As WebRequest = HttpWebRequest.Create(URL)
    '    webRequest.Method = "POST"
    '    webRequest.ContentType = "application/x-www-form-urlencoded"

    '    Dim postParams As String = "accountType=" + HttpUtility.UrlEncode(ACCOUNT_TYPE) + "&Email=" + HttpUtility.UrlEncode(email) + "&Passwd=" + HttpUtility.UrlEncode(password) + "&service=" + HttpUtility.UrlEncode(SERVICE) + "&source=" + HttpUtility.UrlEncode(SOURCE)

    '    Dim postBytes As Byte() = Encoding.UTF8.GetBytes(postParams)
    '    webRequest.ContentLength = postBytes.Length

    '    Dim strmReq As Stream = webRequest.GetRequestStream()
    '    strmReq.Write(postBytes, 0, postBytes.Length)
    '    strmReq.Close()

    '    Dim tblResponse As Hashtable = Nothing
    '    Dim response As WebResponse = Nothing
    '    'Attempt to get authToken 
    '    Try
    '        response = webRequest.GetResponse()
    '    Catch ex As WebException
    '        Dim authException As AuthTokenException = ExtractException(ex)
    '        Throw authException
    '    End Try
    '    'Have response, now to parse out the authToken and return
    '    tblResponse = ParseResponse(response)

    '    If tblResponse.ContainsKey("Auth") Then
    '        Return DirectCast(tblResponse("Auth"), String)
    '    Else
    '        Throw New AuthTokenException(AuthTokenErrorCode.Unknown, "", "", "", "Login failed - Could not find Auth key in response", Nothing)
    '    End If
    'End Function

End Class
