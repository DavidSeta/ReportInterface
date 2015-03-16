Imports System
Imports System.Runtime.Serialization

''' <summary>
''' An enumeration for all the possible exceptions you can
''' get from ClientLogin API.
''' </summary>
Public Enum AuthTokenErrorCode
    ''' <summary>
    ''' The login request used a username or password that is not recognized.
    ''' </summary>
    BadAuthentication

    ''' <summary>
    ''' The account email address has not been verified. The user will need to
    ''' access their Google account directly to resolve the issue before logging
    ''' in using a non-Google application.
    ''' </summary>
    NotVerified

    ''' <summary>
    ''' The user has not agreed to terms. The user will need to access their Google
    ''' account directly to resolve the issue before logging in using a non-Google application.
    ''' </summary>
    TermsNotAgreed

    ''' <summary>
    ''' A CAPTCHA is required. (A response with this error code will also contain an image URL
    ''' and a CAPTCHA token.)
    ''' </summary>
    CaptchaRequired

    ''' <summary>
    ''' The error is unknown or unspecified; the request contained invalid input or was malformed.
    ''' </summary>
    Unknown

    ''' <summary>
    ''' The user account has been deleted.
    ''' </summary>
    AccountDeleted

    ''' <summary>
    ''' The user account has been disabled.
    ''' </summary>
    AccountDisabled

    ''' <summary>
    ''' The user's access to the specified service has been disabled. (The user account may
    ''' still be valid.)
    ''' </summary>
    ServiceDisabled

    ''' <summary>
    ''' The service is not available; try again later.
    ''' </summary>
    ServiceUnavailable
End Enum

''' <summary>
''' An exception class that represents an exception thrown by ClientLogin API.
''' </summary>
<Serializable()> _
Public Class AuthTokenException
    Inherits System.ApplicationException
    ''' <summary>
    ''' The error code associated with this Auth Exception.
    ''' </summary>
    Private m_errorCode As AuthTokenErrorCode

    ''' <summary>
    ''' The url that describes this error.
    ''' </summary>
    Private m_errorUrl As String = ""

    ''' <summary>
    ''' The token associated with the captcha, if a Require Captcha error
    ''' is triggered.
    ''' </summary>
    Private m_captchaToken As String = ""

    ''' <summary>
    ''' The url for the captcha. If Require Captcha error is triggered, then
    ''' this url should be presented to the users to unlock their accounts.
    ''' </summary>
    Private m_captchaUrl As String = ""

    ''' <summary>
    ''' A url that describes the error for this exception.
    ''' </summary>
    Public Property ErrorUrl() As String
        Get
            Return m_errorUrl
        End Get
        Set(ByVal value As String)
            m_errorUrl = value
        End Set
    End Property

    ''' <summary>
    ''' A token to identify the captcha if it gets triggered.
    ''' </summary>
    Public Property CaptchaToken() As String
        Get
            Return m_captchaToken
        End Get
        Set(ByVal value As String)
            m_captchaToken = value
        End Set
    End Property

    ''' <summary>
    ''' The url for the captcha page page. Append this url to
    ''' http://www.google.com/accounts.
    ''' </summary>
    Public Property CaptchaUrl() As String
        Get
            Return m_captchaUrl
        End Get
        Set(ByVal value As String)
            m_captchaUrl = value
        End Set
    End Property

    ''' <summary>
    ''' The error code that caused this exception.
    ''' </summary>
    Public Property ErrorCode() As AuthTokenErrorCode
        Get
            Return m_errorCode
        End Get
        Set(ByVal value As AuthTokenErrorCode)
            m_errorCode = value
        End Set
    End Property

    ''' <summary>
    ''' Public constructor.
    ''' </summary>
    ''' <param name="standardErrorCode">The error code for this exception.
    ''' </param>
    Public Sub New(ByVal standardErrorCode As AuthTokenErrorCode)
        Me.New(standardErrorCode, "", "", "", "", Nothing)
    End Sub

    ''' <summary>
    ''' Public constructor.
    ''' </summary>
    ''' <param name="standardErrorCode">The error code for this exception.
    ''' </param>
    ''' <param name="errorUrl">The error url for this exception.</param>
    Public Sub New(ByVal standardErrorCode As AuthTokenErrorCode, ByVal errorUrl As String)
        Me.New(standardErrorCode, errorUrl, "", "", "", Nothing)
    End Sub

    ''' <summary>
    ''' Public constructor.
    ''' </summary>
    ''' <param name="standardErrorCode">The error code for this exception.
    ''' </param>
    ''' <param name="errorUrl">The error url for this exception.</param>
    ''' <param name="captchaToken">The captcha token, if applicable.</param>
    ''' <param name="captchaUrl">The captcha url, if applicable.</param>
    Public Sub New(ByVal standardErrorCode As AuthTokenErrorCode, ByVal errorUrl As String, ByVal captchaToken As String, ByVal captchaUrl As String)
        Me.New(standardErrorCode, errorUrl, captchaToken, captchaUrl, "", Nothing)
    End Sub

    ''' <summary>
    ''' Public constructor.
    ''' </summary>
    ''' <param name="standardErrorCode">The error code for this exception.
    ''' </param>
    ''' <param name="errorUrl">The error url for this exception.</param>
    ''' <param name="captchaToken">The captcha token, if applicable.</param>
    ''' <param name="captchaUrl">The captcha url, if applicable.</param>
    ''' <param name="message">An additional error message for this exception,
    ''' added by the code throwing this exception.</param>
    Public Sub New(ByVal standardErrorCode As AuthTokenErrorCode, ByVal errorUrl As String, ByVal captchaToken As String, ByVal captchaUrl As String, ByVal message As String)
        Me.New(standardErrorCode, errorUrl, captchaToken, captchaUrl, message, Nothing)
    End Sub

    ''' <summary>
    ''' Public constructor.
    ''' </summary>
    ''' <param name="standardErrorCode">The error code for this exception.
    ''' </param>
    ''' <param name="errorUrl">The error url for this exception.</param>
    ''' <param name="captchaToken">The captcha token, if applicable.</param>
    ''' <param name="captchaUrl">The captcha url, if applicable.</param>
    ''' <param name="message">An additional error message for this exception,
    ''' added by the code throwing this exception.</param>
    ''' <param name="innerException">An inner exception that this exception
    ''' will wrap around.</param>
    Public Sub New(ByVal standardErrorCode As AuthTokenErrorCode, ByVal errorUrl As String, ByVal captchaToken As String, ByVal captchaUrl As String, ByVal message As String, ByVal innerException As Exception)
        MyBase.New(message, innerException)
        Me.m_errorCode = standardErrorCode
        Me.m_errorUrl = errorUrl
        Me.m_captchaToken = captchaToken
        Me.m_captchaUrl = captchaUrl
    End Sub

    ''' <summary>
    ''' Protected constructor, used by serialization frameworks while
    ''' deserializing an exception object.
    ''' </summary>
    ''' <param name="info">Info about the serialization context.</param>
    ''' <param name="context">A streaming context that represents the
    ''' serialization stream.</param>
    Protected Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
        MyBase.New(info, context)
        Me.m_errorCode = DirectCast(info.GetValue("ErrorCode", GetType(AuthTokenErrorCode)), AuthTokenErrorCode)
        Me.m_errorUrl = DirectCast(info.GetValue("ErrorUrl", GetType(String)), String)
        Me.m_captchaToken = DirectCast(info.GetValue("CaptchaToken", GetType(String)), String)
        Me.m_captchaUrl = DirectCast(info.GetValue("CaptchaUrl", GetType(String)), String)
    End Sub

    ''' <summary>
    ''' This method is called by serialization frameworks while serializing
    ''' an exception object.
    ''' </summary>
    ''' <param name="info">Info about the serialization context.</param>
    ''' <param name="context">A streaming context that represents the
    ''' serialization stream.</param>
    Public Overloads Overrides Sub GetObjectData(ByVal info As SerializationInfo, ByVal context As StreamingContext)
        MyBase.GetObjectData(info, context)
        info.AddValue("ErrorCode", m_errorCode, GetType(AuthTokenErrorCode))
        info.AddValue("ErrorUrl", m_errorCode, GetType(String))
        info.AddValue("CaptchaToken", m_errorCode, GetType(String))
        info.AddValue("CaptchaUrl", m_errorCode, GetType(String))
    End Sub
End Class
