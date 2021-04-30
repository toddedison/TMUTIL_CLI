Imports RestSharp
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Web.HttpUtility
Public Class tmWrapper

End Class

Public Class genJSONobj
    Public isSuccess$
    Public Message$
    Public Data$
End Class
Public Class TM_Client
    Public lastError$ = ""
    Public tmFQDN$ 'ie http://localhost https://myserver.myzone.com
    Public isConnected As Boolean
    Private slToken$

    Public Sub New(fqdN$, uN$, pw$)

        tmFQDN = fqdN

        Me.isConnected = False

        'Console.WriteLine("New TM_Client activated: " + Replace(tmFQDN, "https://", ""))

        Dim client = New RestClient(fqdN + "/token")
        Dim request = New RestRequest(Method.POST)
        Dim response As IRestResponse

        client.Timeout = -1

        request.RequestFormat = DataFormat.None

        request.AddHeader("Content-Type", "application/x-www-form-urlencoded")
        request.AddParameter("username", uN)
        request.AddParameter("password", pw)
        request.AddParameter("grant_type", "password")

        response = client.Execute(request)
        'Console.WriteLine(response.Content)

        If response.IsSuccessful = True Then
            Me.isConnected = True
            Dim O As JObject = JObject.Parse(response.Content)
            slToken = O.SelectToken("access_token")
            Exit Sub
        End If

        If IsNothing(response.ErrorMessage) = False Then
            lastError = response.ErrorMessage.ToString
        Else
            lastError = response.Content
        End If

        Console.WriteLine(lastError)
        'Console.WriteLine("Retrieved access token")
    End Sub

    Private Function getAPIData(ByVal urI$, Optional ByVal usePOST As Boolean = False, Optional ByVal addJSONbody$ = "") As String
        getAPIData = ""
        If isConnected = False Then Exit Function

        Dim client = New RestClient(tmFQDN + urI)
        Dim request As RestRequest
        If usePOST = False Then request = New RestRequest(Method.GET) Else request = New RestRequest(Method.POST)

        Dim response As IRestResponse

        request.AddHeader("Authorization", "Bearer " + slToken)
        request.AddHeader("Accept", "application/json")

        If Len(addJSONbody) Then
            request.AddHeader("Content-Type", "application/json")
            request.AddHeader("Accept-Encoding", "gzip, deflate, br")
            request.AddHeader("Connection", "keep-alive")
            request.AddParameter("application/json", addJSONbody, ParameterType.RequestBody)
        End If

        response = client.Execute(request)

        Dim O As JObject = JObject.Parse(response.Content)

        If CBool(O.SelectToken("IsSuccess")) = False Then
            getAPIData = "ERROR:Could not retrieve " + urI
            Exit Function
        End If

        Return O.SelectToken("Data").ToString
    End Function
    Public Function getGroups() As List(Of tmGroups)
        getGroups = New List(Of tmGroups)
        Dim jsoN$ = getAPIData("/api/groups")

        getGroups = JsonConvert.DeserializeObject(Of List(Of tmGroups))(jsoN)

    End Function

    Public Function getProjectsOfGroup(G As tmGroups) As List(Of tmProjInfo)
        getProjectsOfGroup = New List(Of tmProjInfo)

        Dim jBody$ = ""
        jBody = JsonConvert.SerializeObject(G)
        Dim jsoN$ = getAPIData("/api/group/projects", True, jBody$)

        getProjectsOfGroup = JsonConvert.DeserializeObject(Of List(Of tmProjInfo))(jsoN)
    End Function

    Public Function getProject(projID As Long) As tmModel
        getProject = New tmModel

        Dim jsoN$ = getAPIData("/api/diagram/" + projID.ToString + "/components")

        Dim nD As JObject = JObject.Parse(jsoN)
        Dim nodeJSON$ = nD.SelectToken("Model").ToString

        Dim nodeS As List(Of tmNodeData)
        nodeJSON = cleanJSON(nodeJSON)
        nodeJSON = cleanJSONright(nodeJSON)

        Dim SS As New JsonSerializerSettings
        SS.NullValueHandling = NullValueHandling.Ignore

        nodeS = JsonConvert.DeserializeObject(Of List(Of tmNodeData))(nodeJSON, SS)

        getProject = JsonConvert.DeserializeObject(Of tmModel)(jsoN)
        getProject.Nodes = nodeS
    End Function

End Class

Public Class tmGroups
        Public Id As Long
        Public Name$
        Public isSystem As Boolean
        Public ParentGroupId As Integer
        Public ParentGroupName$
        Public Department$
        Public DepartmentId$
        Public Guid As System.Guid
        Public Projects$
        Public GroupUsers$
        Public AllProjInfo As List(Of tmProjInfo)
    End Class
Public Class tmProjInfo
    Public Id As Long
    Public Name$
    Public Version$
    'Public Guid As System.Guid
    Public isInternal As Boolean
    Public RiskId As Integer
    'Public RiskName As Integer
    Public [Type] As String
    Public CreatedByUserEmail$
    Public Model As tmModel
End Class

Public Class tmProjThreat
        Public Id As Long
        'Public Guid as system.guid 'nulls
        Public Name$
        Public Description$
        Public RiskId As Integer
        Public RiskName$
        Public LibraryId As Integer
        Public Labels$
        Public Reference$
        Public Automated As Boolean
        'public ThreatAttributes$ 'nulls
        Public StatusName$
        Public IsHidden As Boolean
        Public CompanyId As Integer
        'Public SharingType$
        'Public ReadOnly As Boolean
        Public isDefault As Boolean
        Public DateCreated$
        Public LastUpdated$
        Public DepartmentId As Integer
        Public DepartmentName$
        'public Version$ 
        Public IsSystemDepartment As Boolean

    End Class

    Public Class tmProjSecReq
        Public Id As Long
        Public Name$
        Public Description$
        Public RiskId As Integer
        Public IsCompensatingControl As Boolean
        Public RiskName$
        Public Labels$
        Public LibraryId As Integer
        'Public Guid as System.Guid
        Public StatusName$
        Public SourceName$
        Public IsHidden As Boolean
    End Class

    Public Class tmNodeData
        Public key As Single
        Public category$
        Public Name$
        Public [type] As String
        Public ImageURI$
        Public NodeId As Integer
        Public ComponentId As Integer
        Public group As Single
        Public Location$
        Public color$
        Public BorderColor$
        Public BackgroundColor$
        Public TitleColor$
        Public TitleBackgroundColor$
        Public isGroup As Boolean
        Public isGraphExpanded As Boolean
        Public FullName$
        Public cpeid$ ' As Integer
        Public networkComponentId As Integer
        Public BorderThickness As Decimal
        Public AllowMove As Boolean
        Public LayoutWidth As Single
        Public LayoutHeight As Single
        Public componentProperties$
        Public listThreats As List(Of tmProjThreat)
        Public listSecurityRequirements As List(Of tmProjSecReq)
    Public ComponentTypeName$
    Public ComponentName$
    Public Id As Long
    Public threatCount As Integer
End Class

    Public Class tmModel
        Public Id As Long
        Public Name$
        Public IsAppliedForApproval As Boolean
        Public Guid As System.Guid
        Public Nodes As List(Of tmNodeData)
    End Class

