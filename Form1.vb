Imports System.Collections.Specialized
Imports System.Net
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        g_AuthToken = GetAuthToken()
        DataGridView1.DataSource = Contacts.GetContacts()
    End Sub

    Public Function GetAuthToken() As String
        Dim oAuthPrompt As New oAuth2Prompt()
        If Not oAuthPrompt.ShowDialog() = DialogResult.OK Then End

        'set uri
        Dim request_uri As String = " http://api.zendirect.com/v2/oauth2/access_token"
        'set data to post
        Dim data As String

        data = "client_id=" & ReadSetting("client_id") & "&client_secret=" & ReadSetting("client_secret") & "&redirect_uri=" & System.Uri.EscapeDataString(ReadSetting("redirect_uri")) & "&grant_type=authorization_code&code=" & oAuthPrompt.AuthCode
        Dim dataBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(data)
        'initialize request
        Dim req As WebRequest = WebRequest.Create(request_uri)
        'set request properties
        req.Method = "POST"
        req.ContentLength = dataBytes.Length
        req.ContentType = "application/x-www-form-urlencoded"
        req.Headers.Add("Accept-Encoding: gzip, deflate")
        req.Headers.Add("Accept-Language: en-US,en;q=0.8")
        System.Net.ServicePointManager.Expect100Continue = False
        'open the request stream
        Dim dataStream As IO.Stream = req.GetRequestStream()
        'write to the stream
        dataStream.Write(dataBytes, 0, dataBytes.Length)
        'close the stream
        dataStream.Close()
        'obtain response
        Dim res As WebResponse = req.GetResponse()
        dataStream = res.GetResponseStream()
        'read response
        Dim reader As New IO.StreamReader(dataStream)
        Dim oAuth = JObject.Parse(reader.ReadToEnd())
        Dim strAuthToken = oAuth.Property("access_token").Value.ToString
        Return strAuthToken


    End Function

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        DataGridView1.DataSource = Nothing
        DataGridView1.DataSource = Contacts.GetContacts()
    End Sub

    Private Sub btnDeleteOne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteOne.Click
        If DataGridView1.CurrentRow IsNot Nothing Then
            If MessageBox.Show("Delete [" & CType(DataGridView1.CurrentRow.DataBoundItem, Contact).name & "] ?", "Remove Contact", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                CType(DataGridView1.CurrentRow.DataBoundItem, Contact).Delete()
                DataGridView1.DataSource = Nothing
                DataGridView1.DataSource = Contacts.GetContacts()
            End If
        End If
    End Sub

    Private Sub btnDeleteAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteAll.Click
        Dim oContacts = CType(DataGridView1.DataSource, Contacts)
        Dim i As Integer = 0
        Dim iCount As Integer = oContacts.Count.ToString
        For Each oContact In oContacts
            lblStatus.Text = i.ToString & " of " & iCount.ToString
            i += 1
            Application.DoEvents()
            oContact.Delete(False)
        Next
        lblStatus.Text = ""
        DataGridView1.DataSource = Nothing
        DataGridView1.DataSource = Contacts.GetContacts()

    End Sub

    Private Sub btnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImport.Click

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then

            Dim book = New LinqToExcel.ExcelQueryFactory(OpenFileDialog1.FileName)

            Dim query = (From row In book.Worksheet(0) Select New Contact() With {.firstname = row("First Name").Cast(Of String)(), _
                                                                                  .lastname = row("Last Name").Cast(Of String)(), _
                                                                                  .company = row("Company").Cast(Of String)(), _
                                                                                  .street = row("Street").Cast(Of String)(), _
                                                                                  .city = row("City").Cast(Of String)(), _
                                                                                  .state = row("State").Cast(Of String)(), _
                                                                                  .postalcode = row("Postal Code").Cast(Of String)(), _
                                                                                  .email = row("E-mail").Cast(Of String)(), _
                                                                                  .phone = row("Phone").Cast(Of String)() _
                                                                                  }).ToList
            Dim i As Integer = 0
            Dim iCount As Integer = query.Count.ToString
            For Each oNewContact In query
                lblStatus.Text = i.ToString & " of " & iCount.ToString
                i += 1
                Application.DoEvents()
                oNewContact.Save()
            Next
            lblStatus.Text = ""
            DataGridView1.DataSource = Nothing
            DataGridView1.DataSource = Contacts.GetContacts()
        End If
    End Sub

End Class
