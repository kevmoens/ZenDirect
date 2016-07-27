Imports System.Net
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Collections.Specialized

Public Class Contact

    Private _name As String
    Public Property name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
            If _name.Contains(" ") Then
                _firstname = _name.Substring(0, _name.IndexOf(" ")).TrimEnd({" "c})
                _lastname = _name.Substring(_name.IndexOf(" "))
            End If
        End Set
    End Property

    Private _firstname As String
    Public Property firstname() As String
        Get
            Return _firstname
        End Get
        Set(ByVal value As String)
            _firstname = value
        End Set
    End Property

    Private _lastname As String
    Public Property lastname() As String
        Get
            Return _lastname
        End Get
        Set(ByVal value As String)
            _lastname = value
        End Set
    End Property

    Private _company As String
    Public Property company() As String
        Get
            Return _company
        End Get
        Set(ByVal value As String)
            _company = value
        End Set
    End Property

    Private _title As String
    Public Property title() As String
        Get
            Return _title
        End Get
        Set(ByVal value As String)
            _title = value
        End Set
    End Property

    Private _street As String
    Public Property street() As String
        Get
            Return _street
        End Get
        Set(ByVal value As String)
            _street = value
        End Set
    End Property

    Private _street2 As String
    Public Property street2() As String
        Get
            Return _street2
        End Get
        Set(ByVal value As String)
            _street2 = value
        End Set
    End Property

    Private _city As String
    Public Property city() As String
        Get
            Return _city
        End Get
        Set(ByVal value As String)
            _city = value
        End Set
    End Property

    Private _state As String
    Public Property state() As String
        Get
            Return _state
        End Get
        Set(ByVal value As String)
            _state = value
        End Set
    End Property

    Private _postalcode As String
    Public Property postalcode() As String
        Get
            Return _postalcode
        End Get
        Set(ByVal value As String)
            _postalcode = value
        End Set
    End Property

    Private _country As String
    Public Property country() As String
        Get
            Return _country
        End Get
        Set(ByVal value As String)
            _country = value
        End Set
    End Property

    Private _email As String
    Public Property email() As String
        Get
            Return _email
        End Get
        Set(ByVal value As String)
            _email = value
        End Set
    End Property


    Private _phone As String
    Public Property phone() As String
        Get
            Return _phone
        End Get
        Set(ByVal value As String)
            _phone = value
        End Set
    End Property

    Private _birthdate As String
    Public Property birthdate() As String
        Get
            Return _birthdate
        End Get
        Set(ByVal value As String)
            _birthdate = value
        End Set
    End Property

    Private _id As String
    Public Property id() As String
        Get
            Return _id
        End Get
        Set(ByVal value As String)
            _id = value
        End Set
    End Property

    Private _contactId As String
    Public Property contactId() As String
        Get
            Return _contactId
        End Get
        Set(ByVal value As String)
            _contactId = value
        End Set
    End Property


    Private _contactgroup As String
    Public Property contactgroup() As String
        Get
            Return _contactgroup
        End Get
        Set(ByVal value As String)
            _contactgroup = value
        End Set
    End Property

    Public Sub Delete(Optional ByVal ShowMessage As Boolean = True)
        Dim resourceUrl = String.Format("http://api.zendirect.com/v2/contact/deletecontact/{0}/", _id)
        Dim delContact = New NameValueCollection()

        Dim client = New WebClient()
        client.Headers("Authorization") = "Bearer " + g_AuthToken
        Dim result = client.UploadValues(resourceUrl, "DELETE", delContact)
        If ShowMessage Then
            Dim response = System.Text.Encoding.Default.GetString(result)
            MsgBox(response)
        End If
        'Dim contactResponse = JsonConvert.DeserializeObject(response)
    End Sub
    Public Sub Save()
        Dim resourceUrl = "http://api.zendirect.com/v2/contact/savecontact/"
        Dim newContact = New NameValueCollection()
        newContact.Add("firstname", If(_firstname Is Nothing, "", _firstname))
        newContact.Add("lastname", If(_lastname Is Nothing, "", _lastname))
        newContact.Add("company", If(_company Is Nothing, "", _company))
        newContact.Add("street", If(_street Is Nothing, "", _street))
        newContact.Add("street2", If(_street2 Is Nothing, "", _street2))
        newContact.Add("city", If(_city Is Nothing, "", _city))
        newContact.Add("state", If(_state Is Nothing, "", _state))
        newContact.Add("postalcode", If(_postalcode Is Nothing, "", _postalcode))
        newContact.Add("country", If(_country Is Nothing, "United States", _country))
        newContact.Add("email", If(_email Is Nothing, "", _email))
        newContact.Add("phone", If(_phone Is Nothing, "", _phone))
        newContact.Add("birthdate", "")
        newContact.Add("contactgroup", If(_contactgroup Is Nothing, "", _contactgroup))
        Dim client = New WebClient()

        client.Headers("Authorization") = "Bearer " + g_AuthToken
        Dim result = client.UploadValues(resourceUrl, "POST", newContact)

        Dim response = System.Text.Encoding.Default.GetString(result)

        Dim contactResponse = JsonConvert.DeserializeObject(response)
    End Sub
End Class
Public Class Contacts
    Inherits List(Of Contact)
    Public Shared Function GetContacts() As Contacts
        Dim resourceUrl = "http://api.zendirect.com/v2/contact/getContacts/id (optional)/"

        Dim req = WebRequest.Create(resourceUrl)
        req.Method = "GET"
        req.ContentType = "application/json"
        req.Headers("Authorization") = "Bearer " + g_AuthToken
        Try
            Dim resp = req.GetResponse()


            Dim reader = New IO.StreamReader(resp.GetResponseStream())
            Dim resultData = reader.ReadToEnd()
            Dim oData = JObject.Parse(resultData)
            Dim strContacts = oData.Property("contacts").Value.ToString
            Dim jsonResult = JsonConvert.DeserializeObject(Of Contacts)(strContacts)
            Return jsonResult
        Catch ex As Exception
            'MsgBox("Error Loading Records")
            Return New Contacts()
        End Try
    End Function
End Class