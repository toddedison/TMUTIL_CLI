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
    Public Function getAWSaccounts() As List(Of tmAWSacc)
        getAWSaccounts = New List(Of tmAWSacc)
        Dim jsoN$ = getAPIData("/api/thirdparty/list/AWS")
        getAWSaccounts = JsonConvert.DeserializeObject(Of List(Of tmAWSacc))(jsoN)

    End Function
    Public Function getVPCs(awsID As Long) As List(Of tmVPC)
        On Error GoTo errorcatch

        getVPCs = New List(Of tmVPC)
        Dim jsoN$ = getAPIData("/api/thirdparty/account/" + awsID.ToString + "/AWS")
        Dim O As JObject = JObject.Parse(jsoN)
        jsoN = O.SelectToken("VPCs").ToString
        getVPCs = JsonConvert.DeserializeObject(Of List(Of tmVPC))(jsoN)

errorcatch:
        'empty JSON returns empty list
    End Function


    Public Function getAllProjects() As List(Of tmProjInfo)
        getAllProjects = New List(Of tmProjInfo)
        Dim jsoN$ = getAPIData("/api/projects/GetAllProjects")

        getAllProjects = JsonConvert.DeserializeObject(Of List(Of tmProjInfo))(jsoN)

    End Function


    Public Function getLabels(Optional ByVal isSystem As Boolean = False) As List(Of tmLabels)
        getLabels = New List(Of tmLabels)
        Dim jsoN$ = getAPIData("/api/labels?isSystem=" + LCase(CStr(isSystem))) 'true") '?isSystem=false")

        getLabels = JsonConvert.DeserializeObject(Of List(Of tmLabels))(jsoN)
    End Function

    Public Function getProjectsOfGroup(G As tmGroups) As List(Of tmProjInfo)
        getProjectsOfGroup = New List(Of tmProjInfo)

        Dim jBody$ = ""
        jBody = JsonConvert.SerializeObject(G)
        Dim jsoN$ = getAPIData("/api/group/projects", True, jBody$)

        getProjectsOfGroup = JsonConvert.DeserializeObject(Of List(Of tmProjInfo))(jsoN)
    End Function

    Public Function prepVPCmodel(G As tmVPC) As String
        prepVPCmodel = ""

        Dim jBody$ = ""
        jBody = JsonConvert.SerializeObject(G)
        Return "[" + jBody + "]"
    End Function

    Public Async Sub createVPCmodel(G As tmVPC)
        Call getAPIData("/api/thirdparty/createthreatmodelfromvpc", True, prepVPCmodel(G))
    End Sub

    Public Function loadAllGroups(ByRef T As TM_Client, Optional ByVal louD As Boolean = False) As List(Of tmGroups)

        Dim GRP As List(Of tmGroups) = T.getGroups

        For Each G In GRP
            G.AllProjInfo = T.getProjectsOfGroup(G)
            If louD Then Console.WriteLine(G.Name + ": " + G.AllProjInfo.Count.ToString + " models")
            For Each P In G.AllProjInfo
                P.Model = T.getProject(P.Id)
                If louD Then Console.WriteLine(vbCrLf)
                If louD Then Console.WriteLine(col5CLI("Proj/Comp Name", "#/TYPE COMP", "# THR", "# SR"))
                If louD Then Console.WriteLine(col5CLI("--------------", "-----------", "-----", "----"))
                With P.Model
                    T.addThreats(P.Model, P.Id) 'add in details of threats not inside Diagram API (eg Protocols)
                    If louD Then Console.WriteLine(col5CLI(P.Name + " [Id " + P.Id.ToString + "]", .Nodes.Count.ToString, P.Model.modelNumThreats.ToString, P.Model.modelNumThreats(True).ToString))
                    '     Console.WriteLine(col5CLI("   ", "TYPE", "# THR", "# SR"))
                    For Each N In .Nodes
                        If N.category = "Collection" Then
                            If louD Then Console.WriteLine(col5CLI(" >" + N.FullName, "Collection", "", ""))
                            GoTo doneHere
                        End If
                        If N.category = "Project" Then
                            N.ComponentName = N.FullName
                            N.ComponentTypeName = "Nested Model"
                            '                            Console.WriteLine(col5CLI(" >" + N.FullName, "Nested Model", "", ""))
                            'GoTo doneHere
                        End If
                        If louD Then Console.WriteLine(col5CLI(" >" + N.ComponentName, N.ComponentTypeName, N.listThreats.Count.ToString, N.listSecurityRequirements.Count.ToString))
                        '                        Console.WriteLine("------------- " + N.Name + "-Threats:" + N.listThreats.Count.ToString + "-SecRqrmnts:" + N.listSecurityRequirements.Count.ToString)
doneHere:
                    Next
                End With
                If P.Model.Nodes.Count Then
                    If louD Then Console.WriteLine(col5CLI("----------------------------------", "------------------------", "------", "------"))
                End If
            Next

            If louD Then Console.WriteLine(vbCrLf)
        Next

        Return GRP
    End Function

    Public Sub addThreats(ByRef tmMOD As tmModel, ByVal projID As Long)
        ' not all threats are included in COMPONENTS API
        ' to capture all of tmModel and Nodes, pull in Threats API for Project
        ' this routine bridges the gap between info provided in 2 APIs and puts 
        ' as much into tmModel as it can

        Dim jsoN$ = getAPIData("/api/project/" + projID.ToString + "/threats?openStatusOnly=false")
        Dim tNodes As New List(Of tmTThreat)

        tNodes = JsonConvert.DeserializeObject(Of List(Of tmTThreat))(jsoN)

        For Each T In tNodes
            If T.SourceType = "Link" Then
                Dim mNode As Integer
                ' first check parent node.. if parent is a threat model (category=Project)
                ' then do not add the link
                mNode = tmMOD.ndxOFnode(T.SourceId)
                If mNode <> -1 Then
                    'unless id > 300000 and subsequent code creates artificial node
                    'this should be a project/nested model.. do not add these links
                    If tmMOD.Nodes(mNode).category = "Project" Then GoTo skipThose
                End If

                mNode = tmMOD.ndxOFnode(300000 + T.SourceId)
                If mNode = -1 Then
                    Call addNode(tmMOD, T)
                End If
                mNode = tmMOD.ndxOFnode(300000 + T.SourceId)

                With tmMOD.Nodes(mNode)
                    Dim newT As New tmProjThreat
                    newT.Name = T.ThreatName
                    newT.StatusName = T.StatusName
                    newT.Id = T.ThreatId
                    newT.Description = T.Description
                    ' notes is missing!
                    newT.RiskName = T.ActualRiskName
                    .listThreats.Add(newT)
                    .threatCount += 1
                End With
            End If
            'If Mid(T.ThreatName, 1, 4) = "11111111CVE-" Then
            ' have to manually add in CPE-related threats to node collection
            ' correction
            ' CVEs are added as Threats to nodes in the DIAGRAM API, but the "numThreats" property does not
            ' include CVEs as threats. Use node.listThreats.nodecount in lieu of node.numThreats API property
            '
            ' protocols of Nested Models are included in the counts of threats as components


skipThose:
        Next

        'getGroups = JsonConvert.DeserializeObject(Of List(Of tmGroups))(jsoN)
    End Sub

    Public Function ndxVPC(ByVal vpcID As Long, ByRef VPCs As List(Of tmVPC)) As Integer
        'finds object (for now, the node of a model) matching idOFnode
        Dim K As Long = 0

        For Each N In VPCs
            If N.Id = vpcID Then
                Return K
            End If
            K += 1
        Next

        Return -1
    End Function

    Public Sub addNode(ByRef tmMOD As tmModel, ByRef T As tmTThreat)
        ' create node inside tmMOD based on threat T (from Threats API)
        'Dim tmMOD As tmModel = tmPROJ.Model

        Dim newNode As tmNodeData = New tmNodeData

        If T.SourceType = "Link" Then
            newNode.Id = T.SourceId + 300000
            newNode.category = "Component"
            newNode.Name = T.SourceDisplayName
            newNode.type = "Link"
            newNode.ComponentName = T.SourceName
            newNode.ComponentTypeName = "Protocol"

            tmMOD.Nodes.Add(newNode)
            Exit Sub
        End If


    End Sub

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

Public Class tmAWSacc
    Public Id As Long
    Public Name$

End Class
Public Class tmVPC
    Public Id As Long
    Public VPCId$
    Public ThirdpartyInfoId As Long
    Public ProjectId? As Integer
    Public ProjectName$
    Public Version$
    Public VPCTags$
    Public IsSuccess As Boolean
    Public Notes$
    Public Labels$
    Public RiskId As Integer
    Public IsInternal As Boolean
    Public LastSync$
    Public Region$
    Public GroupPermissions() As String
    Public UserPermissions() As String
    Public [Type] As String

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

Public Class tmLabels
    Public Id As Long
    Public Name$
    Public IsSystem As Boolean
    Public Attributes$
End Class

Public Class tmProjInfo
    Public Id As Long
    Public Name$
    Public Version$
    'Public Guid As System.Guid
    Public isInternal As Boolean
    Public RiskId As Integer
    Public RiskName$
    Public CreatedByName$
    Public LastModifiedByName$
    'Public RiskName As Integer
    Public [Type] As String
    Public CreatedByUserEmail$
    Public Model As tmModel
    'Public nodesFromThreats As Integer 'beginning at 300k, nodes from threats like HTTPS
End Class

Public Class tmProjThreat
    ' these threats come from the DIAGRAM API as part of nodeArray
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
Public Class tmTThreat
    Public Id As Long

    Public ThreatName$
    Public SourceName$
    Public SourceDisplayName$
    Public ThreatId As Integer
    Public ActualRiskName$
    Public StatusName$
    Public Description$
    Public Notes$
    Public SourceId As Long
    Public SourceType$

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
    Public Notes$
    Public Id As Long
    Public threatCount As Integer

    Public Sub New()
        Me.listSecurityRequirements = New List(Of tmProjSecReq)
        Me.listThreats = New List(Of tmProjThreat)

    End Sub
End Class

Public Class TF_Threat
    '    {\"ImagePath\":\"/ComponentImage/DefaultComponent.jpg\",\"ComponentTypeId\":85,\"RiskId\":1,\"CodeTypeId\":1,\"DataClassificationId\":1,\"Name\":\"SampleThreat444444\",\"Labels\":\"AppSec and InfraSec,DevOps\",\"LibrayId\":66,\"Description\":\"<p>Example123</p>\",\"IsCopy\":false}
    Public ImagePath$
    Public ComponentTypeId As Integer
    Public RiskId As Integer
    Public CodeTypeId As Integer
    Public DataClassificationId As Integer
    Public Name$
    Public Labels$
    Public LibrayId As Integer
    Public Description$
    Public IsCopy As Boolean
End Class
Public Class tmModel
    ' model is built from DIAGRAM API
    ' includes all threats and SRs except
    '       protocol threats
    '       CPE CVEs
    '       Nested Models
    ' these require lookup into Threats API
    ' Can correlate CPE CVEs to Component of SourceId
    ' Pull Nested Model Threats in and also store to SourceId

    Public Id As Long
    Public Name$
    Public IsAppliedForApproval As Boolean
    Public Guid As System.Guid
    Public Nodes As List(Of tmNodeData)

    Public Function modelNumThreats(Optional ByVal SecReqInstead As Boolean = False) As Integer
        Dim nuM As Integer = 0
        For Each N In Me.Nodes
            If SecReqInstead Then nuM += N.listSecurityRequirements.Count Else nuM += N.listThreats.Count
        Next
        Return nuM
    End Function

    Public Function ndxOFnode(ByVal idOFnode As Long) As Integer
        'finds object (for now, the node of a model) matching idOFnode
        Dim K As Long = 0

        For Each N In Me.Nodes
            If N.Id = idOFnode Then
                Return K
            End If
            K += 1
        Next

        Return -1
    End Function

End Class

