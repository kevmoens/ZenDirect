Imports System.Net
Imports System.IO
Public Class oAuth2Prompt
    Private _AuthCode As String
    Public Property AuthCode() As String
        Get
            Return _AuthCode
        End Get
        Set(ByVal value As String)
            _AuthCode = value
        End Set
    End Property
    Private Sub oAuth2Prompt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        WebBrowser1.Navigate("http://api.zendirect.com/accounts/login/default.aspx?client_id=9GdS4YvjVB1bp2GK9XEuhw4f8zSKuVh5&redirect_uri=http://zendirect.lakeco.com")
    End Sub

    Private Sub WebBrowser1_Navigated(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserNavigatedEventArgs) Handles WebBrowser1.Navigated
        If e.Url.ToString.ToUpper.StartsWith("http://zendirect.lakeco.com/?code=".ToUpper) Then
            _AuthCode = e.Url.ToString.Substring("http://zendirect.lakeco.com/?code=".Length)
            DialogResult = DialogResult.OK
            Me.Close()
        End If
    End Sub

End Class

