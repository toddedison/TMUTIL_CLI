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
    Public Event resultsProcessed(projID As Long, projModelData As tmModel)
    Public Event componentBuilt(comP As tmComponent)
    Public Event threatBuilt(threaT As tmProjThreat)

    Public lib_Comps As List(Of tmComponent)
    Public lib_TH As List(Of tmProjThreat)
    Public lib_SR As List(Of tmProjSecReq)

    Public sysLabels As List(Of tmLabels)
    Public labelS As List(Of tmLabels)
    Public groupS As List(Of tmGroups)
    Public librarieS As List(Of tmLibrary)


    Public Sub New(fqdN$, uN$, pw$)
        On Error GoTo errorcatch

        lib_Comps = New List(Of tmComponent)
        lib_TH = New List(Of tmProjThreat)
        lib_SR = New List(Of tmProjSecReq)

        tmFQDN = fqdN

        Me.isConnected = False

        'addlog("New TM_Client activated: " + Replace(tmFQDN, "https://", ""))

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
        'addlog(response.Content)

        If response.IsSuccessful = True Then
            Me.isConnected = True
            Dim O As JObject = JObject.Parse(response.Content)
            slToken = O.SelectToken("access_token")
            Form1.addLOG("Retrieved Token")
            Exit Sub
        End If

        If IsNothing(response.ErrorMessage) = False Then
            lastError = response.ErrorMessage.ToString
        Else
            lastError = response.Content
        End If

        Form1.addLOG("ERROR: Could not get token - " + lastError)
        'addlog("Retrieved access token")
        Exit Sub

errorcatch:
        lastError = "Could not login - check URL"
    End Sub

    Private Function getAPIData(ByVal urI$, Optional ByVal usePOST As Boolean = False, Optional ByVal addJSONbody$ = "") As String
        getAPIData = ""
        If isConnected = False Then Exit Function

        Dim client = New RestClient(tmFQDN + urI)
        Dim request As RestRequest
        If usePOST = False Then request = New RestRequest(Method.GET) Else request = New RestRequest(Method.POST)

        Form1.addLOG("Accessing " + tmFQDN + urI + " POST: " + usePOST.ToString)

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

        On Error GoTo errorcatch

        If CBool(O.SelectToken("IsSuccess")) = False Then
            getAPIData = "ERROR:Could not retrieve " + urI
            Exit Function
        End If

        Return O.SelectToken("Data").ToString

        Exit Function

errorcatch:
        Return ""
    End Function

    Public Function getGroups() As List(Of tmGroups)
        getGroups = New List(Of tmGroups)
        Dim jsoN$ = getAPIData("/api/groups")

        If Len(jsoN) <> 0 And Mid(jsoN, 1, 6) <> "ERROR:" Then getGroups = JsonConvert.DeserializeObject(Of List(Of tmGroups))(jsoN)

    End Function

    Public Function getCompFrameworks() As List(Of tmCompFramework)
        getCompFrameworks = New List(Of tmCompFramework)
        Dim jsoN$ = getAPIData("/api/scf/complianceframeworks/true")

        getCompFrameworks = JsonConvert.DeserializeObject(Of List(Of tmCompFramework))(jsoN)

    End Function

    Public Function getLibraries() As List(Of tmLibrary)
        getLibraries = New List(Of tmLibrary)
        Dim jsoN$ = getAPIData("/api/libraries")

        getLibraries = JsonConvert.DeserializeObject(Of List(Of tmLibrary))(jsoN)

    End Function
    Public Function getTemplates() As List(Of tmTemplate)
        getTemplates = New List(Of tmTemplate)
        Dim jsoN$ = getAPIData("/api/templates") '/" + tempID.ToString)

        getTemplates = JsonConvert.DeserializeObject(Of List(Of tmTemplate))(jsoN)

    End Function
    Public Function getTemplate(tempID As Integer) As tmTemplate
        getTemplate = New tmTemplate
        Dim jsoN$ = getAPIData("/api/template/" + tempID.ToString)

        getTemplate = JsonConvert.DeserializeObject(Of tmTemplate)(jsoN)

    End Function

    Public Function getTmNodeArray(ByRef templatE As tmTemplate) As tmNodeArray
        getTmNodeArray = New tmNodeArray
        Dim jsoN$ = templatE.Json

        getTmNodeArray = JsonConvert.DeserializeObject(Of tmNodeArray)(jsoN)
    End Function


    Public Function getAllProjects() As List(Of tmProjInfo)
        getAllProjects = New List(Of tmProjInfo)
        '        Dim jsoN$ = getAPIData("/api/projects/GetAllProjects")
        Dim c$ = Chr(34)
        Dim bodY$ = "{" + c + "virtualGridStateRequestModel" + c + ":{" + c + "state" + c + ":{" + c + "skip" + c + ":0," + c + "take" + c + ":100}}," + c + "projectFilterModel" + c + ":{" + c + "ActiveProjects" + c + ":true}}"

        Dim jsoN$ = getAPIData("/api/projects/smartfilter", True, bodY)

        Dim O As JObject = JObject.Parse(jsoN)
        jsoN = O.SelectToken("Data").ToString


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


    Public Function getTFComponents(T As tfRequest) As List(Of tmComponent)
        getTFComponents = New List(Of tmComponent)

        Dim jBody$ = ""
        jBody = JsonConvert.SerializeObject(T)

        Dim jsoN$ = getAPIData("/api/threatframework", True, jBody$)

        getTFComponents = JsonConvert.DeserializeObject(Of List(Of tmComponent))(jsoN)
    End Function


    Public Function getTFSecReqs(T As tfRequest) As List(Of tmProjSecReq)
        getTFSecReqs = New List(Of tmProjSecReq)

        Dim jBody$ = ""
        jBody = JsonConvert.SerializeObject(T)

        Dim jsoN$ = getAPIData("/api/threatframework", True, jBody$)

        getTFSecReqs = JsonConvert.DeserializeObject(Of List(Of tmProjSecReq))(jsoN)
    End Function


    Public Function getTFThreats(T As tfRequest) As List(Of tmProjThreat)
        Dim TF As New List(Of tmProjThreat)

        Dim jBody$ = ""
        jBody = JsonConvert.SerializeObject(T)

        Dim jsoN$ = getAPIData("/api/threatframework", True, jBody$)

        TF = JsonConvert.DeserializeObject(Of List(Of tmProjThreat))(jsoN)

        Return TF
    End Function

    Public Sub buildCompObj(ByVal C As tmComponent) ' As tmComponent 'C As tmComponent, ByRef TH As List(Of tmProjThreat), ByRef SR As List(Of tmProjSecReq)) As tmComponent
        ' Assumes you are passing completed lists of TH and SR using methods above (/api/framework)
        ' This func uses an API QUERY to identify attached TH/DirectSR but only ID and NAME is returned
        ' Full list used to "attach" more complete TH/SR from lists passed ByRef
        ' Will query for Transitive SRs
        ' Recommend multi-threading this func
        ' Multi-threading full pull of SRs attached to Threats might be overkill, especially if they arent used :)

        Dim cResp As New tmCompQueryResp

        Dim jBody$ = ""
        Dim modeL$ = JsonConvert.SerializeObject(C) 'submits serialized model with escaped quotes

        Dim cReq As New tmTFQueryRequest
        With cReq
            .Model = modeL
            .LibraryId = C.LibraryId
            .EntityType = "Components"
        End With

        jBody = JsonConvert.SerializeObject(cReq)

        Dim json$ = getAPIData("/api/threatframework/query", True, jBody)

        'cResp = JsonConvert.DeserializeObject(Of tmCompQueryResp)(json)
        'would have been nice if this had worked.. cresp is a class made up of 3 lists
        'instead have to break up json/sloppy
        Dim thStr$ = Mid(json, 2, InStr(json, "],") - 1)
        json = Mid(json, InStr(json, "],") + 1)
        Dim srStr$ = Mid(json, InStr(json, ",") + 1, InStr(json, "]") - 1)
        'json = Mid(json, InStr(json, "],") + 1)
        'Dim attStr$ = Mid(json, InStr(json, ",") + 1, InStr(json, "]") - 1)

        Dim ndX As Integer

        With cResp
            .listThreats = JsonConvert.DeserializeObject(Of List(Of tmProjThreat))(thStr)
            .listDirectSRs = JsonConvert.DeserializeObject(Of List(Of tmProjSecReq))(srStr)
            'If Len(attStr) > 3 Then
            ' .listAttr = JsonConvert.DeserializeObject(Of List(Of tmAttribute))(attStr)
            ' If .listAttr.Count > 1 Then
            ' saveJSONtoFile(attStr, "attrib_" + C.Id + ".json")
            ' End If
            ' End If

            For Each T In .listThreats
                ndX = ndxTHlib(T.Id) ', lib_TH)
                C.listThreats.Add(lib_TH(ndX))
            Next
            For Each T In .listDirectSRs
                ndX = ndxSRlib(T.Id) ', lib_SR)
                If ndX <> -1 Then C.listDirectSRs.Add(lib_SR(ndX))
            Next

        End With

        Application.DoEvents()

        If C.ComponentTypeName = "Protocols" Or C.ComponentTypeName = "Security Control" Then GoTo skipForProtocolsAndControls

        Dim possibleTransSRs As New List(Of tmProjSecReq)
        possibleTransSRs = addTransitiveSRs(C)

        ' If C.Name = "Azure Blob Storage" Then
        ' Dim K As Integer
        ' K = 3
        ' End If

        For Each TSR In possibleTransSRs
            'C.listTransSRs.Add(TSR)
            If C.numLabels > 0 Then
                If numMatchingLabels(C.Labels, TSR.Labels) / C.numLabels > 0.9 Then C.listTransSRs.Add(TSR)
            End If
        Next

skipForProtocolsAndControls:

        C.isBuilt = True

        RaiseEvent componentBuilt(C)
    End Sub

    Private Function addTransitiveSRs(ByVal C As tmComponent) As List(Of tmProjSecReq)
        addTransitiveSRs = New List(Of tmProjSecReq)
        If C.listThreats.Count = 0 Then Exit Function

        Dim K As Integer
        Dim L As Integer

        Dim ndX As Integer

        Dim jBody$
        Dim jSon$
        Dim tcStr$
        Dim tsrStr$
        Dim transSRs As List(Of tmProjSecReq)
        Dim T As tmProjThreat
        Dim ndxTH As Integer
        Dim cReq As tmTFQueryRequest
        Dim modeL$

        '        Dim pauseVAL As Boolean

        For K = 0 To C.listThreats.Count - 1
            T = C.listThreats(K)
            'here we get the transitive SRs
            ndxTH = ndxTHlib(T.Id) ', lib_TH)


            If lib_TH(ndxTH).isBuilt = True Then
                Application.DoEvents()
                ' added doevents as this seems to trip// suspect lowlevel probs with isBuilt prop
                'this seems to happen a lot
                'if we already know SRs, just add them from coll and skip web request
                For Each TSRID In lib_TH(ndxTH).listLinkedSRs
                    If ndxSR(TSRID, addTransitiveSRs) <> -1 Then GoTo duphere
                    ndX = ndxSRlib(TSRID) ', lib_SR)
                    If ndX <> -1 Then addTransitiveSRs.Add(lib_SR(ndX))
duphere:
                Next
                GoTo skipTheLoad
                ' Else
            End If

            lib_TH(ndxTH).listLinkedSRs = New Collection
                    ' End If

                    cReq = New tmTFQueryRequest
                    Application.DoEvents()
                    modeL$ = JsonConvert.SerializeObject(T) 'submits serialized model with escaped quotes
                    Application.DoEvents()

                    With cReq
                        .Model = modeL
                        .LibraryId = T.LibraryId
                        .EntityType = "Threats"
                    End With

                    jBody$ = JsonConvert.SerializeObject(cReq)
                    jSon$ = getAPIData("/api/threatframework/query", True, jBody)
                    ' should come back with TESTCASES,SECREQ lists

                    tcStr$ = Mid(jSon, 2, InStr(jSon, "],") - 1) 'test cases here if you want 'em
                    jSon = Mid(jSon, InStr(jSon, "],") + 1)
                    tsrStr$ = Mid(jSon, InStr(jSon, ",") + 1, InStr(jSon, "]") - 1)

                    transSRs = New List(Of tmProjSecReq)
                    transSRs = JsonConvert.DeserializeObject(Of List(Of tmProjSecReq))(tsrStr)

            With lib_TH(ndxTH).listLinkedSRs
                For L = 0 To transSRs.Count - 1
                    Dim TSR As tmProjSecReq = transSRs(L)
                    Application.DoEvents()
                    .Add(TSR.Id)
                    Application.DoEvents()

                    If ndxSR(TSR.Id, addTransitiveSRs) <> -1 Then GoTo duplicate
                    ndX = ndxSRlib(TSR.Id) ', lib_SR)
                    If ndX <> -1 Then addTransitiveSRs.Add(lib_SR(ndX))
duplicate:
                Next
            End With
skipTheLoad:
            lib_TH(ndxTH).isBuilt = True

        Next


    End Function

    Public Sub defineTransSRs(ByVal threatID As Long)
        Dim jBody$
        Dim jSon$
        Dim tcStr$
        Dim tsrStr$
        Dim transSRs As List(Of tmProjSecReq)
        Dim T As tmProjThreat
        Dim ndxTH As Integer
        Dim cReq As tmTFQueryRequest
        Dim modeL$

        '        Dim pauseVAL As Boolean

        'here we get the transitive SRs
        ndxTH = ndxTHlib(threatID)
        T = lib_TH(ndxTH)

        cReq = New tmTFQueryRequest
        modeL$ = JsonConvert.SerializeObject(T) 'submits serialized model with escaped quotes

        With cReq
            .Model = modeL
            .LibraryId = T.LibraryId
            .EntityType = "Threats"
        End With

        jBody$ = JsonConvert.SerializeObject(cReq)
        jSon$ = getAPIData("/api/threatframework/query", True, jBody)
        ' should come back with TESTCASES,SECREQ lists

        tcStr$ = Mid(jSon, 2, InStr(jSon, "],") - 1) 'test cases here if you want 'em
        jSon = Mid(jSon, InStr(jSon, "],") + 1)
        tsrStr$ = Mid(jSon, InStr(jSon, ",") + 1, InStr(jSon, "]") - 1)

        transSRs = New List(Of tmProjSecReq)
        transSRs = JsonConvert.DeserializeObject(Of List(Of tmProjSecReq))(tsrStr)

        T.listLinkedSRs = New Collection

        For Each TSR In transSRs
            If grpNDX(T.listLinkedSRs, TSR.Id) <> -1 Then GoTo duplicate
            T.listLinkedSRs.Add(TSR.Id)
duplicate:
        Next
skipTheLoad:
        T.isBuilt = True

        RaiseEvent threatBuilt(T)
    End Sub

    Public Function numMatchingLabels(ByVal LofParent$, ByVal LofChild$) As Integer
        numMatchingLabels = 0
        Dim PL As New Collection
        PL = CSVtoCOLL(LofParent)
        Dim CL As New Collection
        CL = CSVtoCOLL(LofChild)

        For Each C In CL
            If grpNDX(PL, C) Then numMatchingLabels += 1
        Next
    End Function

    Public Function getMatchingLabels(ByVal LofParent$, ByVal LofChild$, Optional ByVal nonMatching As Boolean = False) As Collection
        getMatchingLabels = New Collection '  = 0
        Dim PL As New Collection
        PL = CSVtoCOLL(LofParent)
        Dim CL As New Collection
        CL = CSVtoCOLL(LofChild)

        For Each C In CL
            If grpNDX(PL, C) Then
                If nonMatching = False Then getMatchingLabels.Add(C)
            End If
        Next

        If nonMatching = True Then
            Dim unSelected As New Collection
            For Each PLBL In PL
                If grpNDX(getMatchingLabels, PLBL) = 0 Then
                    unSelected.Add(PLBL)
                End If
            Next
            Return unSelected
        End If

    End Function

    Public Class tmTFQueryRequest
        Public Model$
        Public LibraryId As Integer
        Public EntityType$
    End Class
    Public Function returnSRLZcomponent(C As tmComponent) As String
        Return JsonConvert.SerializeObject(C).ToString

    End Function

    Public Function loadAllGroups(ByRef T As TM_Client, Optional ByVal louD As Boolean = False) As List(Of tmGroups)

        Dim GRP As List(Of tmGroups) = T.getGroups

        For Each G In GRP
            G.AllProjInfo = T.getProjectsOfGroup(G)
            If louD Then Form1.addLOG(G.Name + ": " + G.AllProjInfo.Count.ToString + " models")
            For Each P In G.AllProjInfo
                P.Model = T.getProject(P.Id)
                If louD Then Form1.addLOG(vbCrLf)
                If louD Then Form1.addLOG(col5CLI("Proj/Comp Name", "#/TYPE COMP", "# THR", "# SR"))
                If louD Then Form1.addLOG(col5CLI("--------------", "-----------", "-----", "----"))
                With P.Model
                    T.addThreats(P.Model, P.Id) 'add in details of threats not inside Diagram API (eg Protocols)
                    If louD Then Form1.addLOG(col5CLI(P.Name + " [Id " + P.Id.ToString + "]", .Nodes.Count.ToString, P.Model.modelNumThreats.ToString, P.Model.modelNumThreats(True).ToString))
                    '     form1.addlog(col5CLI("   ", "TYPE", "# THR", "# SR"))
                    For Each N In .Nodes
                        If N.category = "Collection" Then
                            If louD Then Form1.addLOG(col5CLI(" >" + N.FullName, "Collection", "", ""))
                            GoTo doneHere
                        End If
                        If N.category = "Project" Then
                            N.ComponentName = N.FullName
                            N.ComponentTypeName = "Nested Model"
                            '                            form1.addlog(col5CLI(" >" + N.FullName, "Nested Model", "", ""))
                            'GoTo doneHere
                        End If
                        If louD Then Form1.addLOG(col5CLI(" >" + N.ComponentName, N.ComponentTypeName, N.listThreats.Count.ToString, N.listSecurityRequirements.Count.ToString))
                        '                        form1.addlog("------------- " + N.Name + "-Threats:" + N.listThreats.Count.ToString + "-SecRqrmnts:" + N.listSecurityRequirements.Count.ToString)
doneHere:
                    Next
                End With
                If P.Model.Nodes.Count Then
                    If louD Then Form1.addLOG(col5CLI("----------------------------------", "------------------------", "------", "------"))
                End If
            Next

            If louD Then Form1.addLOG(vbCrLf)
        Next

        Return GRP
    End Function
    Public Function pctMatchingLabels(C As tmComponent, S As tmProjSecReq) As Decimal


        Dim nL As Decimal
        nL = 0
        If IsNothing(C.Labels) = True Then
            'addLOG(C.Name + " (" + C.Id.ToString + ") has NULL labels")
            nL = 1
        Else
            If C.numLabels = 0 Then
                ' addLOG(C.Name + " (" + C.Id.ToString + ") has no labels")
                nL = 1
            Else
                nL = Math.Round(numMatchingLabels(C.Labels, S.Labels) / C.numLabels, 2)
            End If
        End If

        Return nL
    End Function
    Public Function ndxProject(projID As Long, ByRef allProj As List(Of tmProjInfo)) As Integer
        ndxProject = -1

        Dim ndX As Integer = 0
        For Each P In allProj
            If P.Id = projID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function
    Public Function ndxThreat(ID As Long, ByRef alL As List(Of tmProjThreat)) As Integer
        ndxThreat = -1

        Dim ndX As Integer = 0
        For Each P In alL
            If P.Id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function

    Public Function ndxTHlib(ID As Long) As Integer
        ndxTHlib = -1

        Dim ndX As Integer = 0
        For Each P In lib_TH
            If P.Id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function

    Public Function ndxSRlib(ID As Long) As Integer
        ndxSRlib = -1

        Dim ndX As Integer = 0
        For Each P In lib_SR
            If P.Id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function

    Public Function ndxTemplate(ID As Long, ByRef alL As List(Of tmTemplate)) As Integer
        ndxTemplate = -1

        Dim ndX As Integer = 0
        For Each P In alL
            If P.Id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function

    Public Function ndxLib(ID As Long) As Integer
        ndxLib = -1

        Dim ndX As Integer = 0
        For Each P In librarieS
            If P.Id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function
    Public Function ndxComp(ID As Long) As Integer ', ByRef alL As List(Of tmComponent)) As Integer
        ndxComp = -1

        'had to change this as passing byref from multi-threaded main causing issues

        Dim ndX As Integer = 0
        For Each P In lib_Comps
            If P.Id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function

    Public Function ndxCompbyName(name$) As Integer ', ByRef alL As List(Of tmComponent)) As Integer
        ndxCompbyName = -1

        'had to change this as passing byref from multi-threaded main causing issues

        Dim ndX As Integer = 0
        For Each P In lib_Comps
            If LCase(P.Name) = LCase(name) Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function

    Public Function ndxSRbyName(name$) As Integer ', ByRef alL As List(Of tmComponent)) As Integer
        ndxSRbyName = -1

        'had to change this as passing byref from multi-threaded main causing issues

        Dim ndX As Integer = 0
        For Each P In lib_SR
            If LCase(P.Name) = LCase(name) Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function

    Public Function ndxSR(ID As Long, ByRef alL As List(Of tmProjSecReq)) As Integer
        ndxSR = -1

        Dim ndX As Integer = 0
        For Each P In alL
            If P.Id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function


    Public Sub addThreats(ByRef tmMOD As tmModel, ByVal projID As Long)
        ' not all threats are included in COMPONENTS API
        ' to capture all of tmModel and Nodes, pull in Threats API for Project
        ' this routine bridges the gap between info provided in 2 APIs and puts 
        ' as much into tmModel as it can

        Dim jsoN$ = getAPIData("/api/project/" + projID.ToString + "/threats?openStatusOnly=false")
        Dim tNodes As New List(Of tmTThreat)

        tNodes = JsonConvert.DeserializeObject(Of List(Of tmTThreat))(jsoN)
        'be careful as this object tmtthreat is diff from projthreat
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

    Public Function getTemplateNodesJSON(ByRef temPnodes As tmNodeArray) As String
        getTemplateNodesJSON = ""

        Dim listOfActions As New List(Of tmNodeAction)
        For Each N In temPnodes.nodeDataArray
            Dim L As New tmNodeAction
            L.action = "Insert"
            L.type = "Node"
            L.changedNode = N
            listOfActions.Add(L)
        Next

        getTemplateNodesJSON = JsonConvert.SerializeObject(listOfActions)
    End Function

    Public Function getProject(projID As Long) As tmModel
        getProject = New tmModel

        Dim jsoN$ = getAPIData("/api/diagram/" + projID.ToString + "/components")

        Dim nD As JObject = JObject.Parse(jsoN)
        Dim nodeJSON$ = nD.SelectToken("Model").ToString

        Dim nodeS As List(Of tmNodeData)
        nodeJSON = cleanJSON(nodeJSON)
        nodeJSON = cleanJSONright(nodeJSON)

        'Dim SS As New JsonSerializerSettings With {.FloatParseHandling = FloatParseHandling.Decimal}

        'component key is a string with this api response

        nodeS = JsonConvert.DeserializeObject(Of List(Of tmNodeData))(nodeJSON) ', SS) ', SS)

        getProject = JsonConvert.DeserializeObject(Of tmModel)(jsoN)
        getProject.Nodes = nodeS
        'Form1.addLOG("Retrieved " + getProject.Name + ": " + getProject.Nodes.Count.ToString + " nodes")

        Call addThreats(getProject, projID) 'add in details of threats not inside Diagram API (eg Protocols)
        For Each N In getProject.Nodes
            If N.category = "Collection" Then
                GoTo doneHere
            End If
            If N.category = "Project" Then
                N.ComponentName = N.FullName
                N.ComponentTypeName = "Nested Model"
            End If
doneHere:
            'dumb.. while Threats in Diagram/Components are filled the SRs look like this:
            '    "Id" 0,
            '  "Name": "Ensure that the DBMS login used by the application has the lowest possible level of privileges in the DBMS",
            '  "Description": 0,
            '  "RiskId": 0,
            '  "IsCompensatingControl": false,
            '  "RiskName": 0,
            '  "Labels": 0,
            '  "LibraryId": 0,
            '  "Guid": 0,
            '  "StatusName": "Open",
            '  "SourceName": 0,
            '  "IsHidden": false
            '},
            '{

            For Each S In N.listSecurityRequirements
                Dim libNDX As Integer = ndxSRbyName(S.Name)
                If libNDX <> -1 Then
                    With lib_SR(ndxSRbyName(S.Name))
                        S.Id = .Id
                        S.IsHidden = .IsHidden
                        S.Labels = .Labels
                        S.RiskName = .RiskName
                        S.SourceName = N.FullName
                    End With
                End If
            Next
        Next

        RaiseEvent resultsProcessed(projID, getProject)

    End Function

End Class

Public Class tfRequest
    'obj used to request Threats, SRs, Components
    'unpublished API
    '    "EntityType""SecurityRequirements","LibraryId":"0","ShowHidden":false
    Public EntityType$
    Public LibraryId As Integer '0 = ALL
    Public ShowHidden As Boolean
End Class

Public Class tmCompFramework
    Public Id As Integer
    Public Name$
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
    Public LastModifiedDate As DateTime
    Public Model As tmModel
    'Public nodesFromThreats As Integer 'beginning at 300k, nodes from threats like HTTPS
    Public ReadOnly Property ProjID() As Long
        Get
            Return Id
        End Get
    End Property
    Public ReadOnly Property ProjName() As String
        Get
            Return Name
        End Get
    End Property
    Public ReadOnly Property Vers() As String
        Get
            Return Version
        End Get
    End Property
    Public ReadOnly Property Nodes() As Long
        Get
            If IsNothing(Model) = True Then Return 0 Else Return Me.Model.Nodes.Count
            Return Id
        End Get
    End Property
    Public ReadOnly Property NumTH() As Long
        Get
            If IsNothing(Model) = True Then Return 0 Else Return Me.Model.modelNumThreats
        End Get
    End Property
    Public ReadOnly Property NumSR() As Long
        Get
            If IsNothing(Model) = True Then Return 0 Else Return Me.Model.modelNumThreats(True)
        End Get
    End Property



End Class
Public Class buildComp
    Public C As tmComponent
    Public TH As List(Of tmProjThreat)
    Public SR As List(Of tmProjSecReq)
End Class

Public Class tmCompQueryResp
    Public listThreats As List(Of tmProjThreat)
    Public listDirectSRs As List(Of tmProjSecReq)
    'Public listTransSRs As List(Of tmProjSecReq)
    Public listAttr As List(Of tmAttribute)
End Class

Public Class tmAttribute
    Public Id As Long
    Public Name$
    Public Description$
End Class

Public Class tmComponent
    Public Id As Long
    Public Name$
    Public Description$
    Public ImagePath$
    Public Labels$
    Public Version$
    Public LibraryId As Integer
    Public ComponentTypeId As Integer
    Public ComponentTypeName$
    Public Color$
    Public Guid As System.Guid
    Public IsHidden As Boolean
    Public DiagralElementId As Integer
    Public ResourceTypeValue$
    'Public Attributes$

    Public listThreats As List(Of tmProjThreat)
    Public listDirectSRs As List(Of tmProjSecReq)
    Public listTransSRs As List(Of tmProjSecReq)
    Public isBuilt As Boolean
    Public Function numLabels() As Integer
        If Len(Labels) = 0 Then
            Return 0
        Else
            Return numCHR(Labels, ",") + 1
        End If
    End Function

    Public ReadOnly Property CompID() As Long
        Get
            Return Id
        End Get
    End Property
    Public ReadOnly Property CompName() As String
        Get
            Return Name
        End Get
    End Property
    Public ReadOnly Property TypeName() As String
        Get
            Return ComponentTypeName
        End Get
    End Property
    Public ReadOnly Property NumTH() As Integer
        Get
            Return listThreats.Count
        End Get
    End Property

    Public ReadOnly Property NumSR() As Integer
        Get
            Return listDirectSRs.Count + listTransSRs.Count
        End Get
    End Property


    Public Sub New()
        listThreats = New List(Of tmProjThreat)
        listDirectSRs = New List(Of tmProjSecReq)
        listTransSRs = New List(Of tmProjSecReq)
        isBuilt = False
    End Sub

End Class

Public Class tmLibrary
    Public Id As Long
    Public Name$
    Public Description$
    Public CompanyId As Integer
    Public Guid As System.Guid
    Public SharingType$
    Public [Readonly] As Boolean
    Public isDefault As Boolean
    Public DateCreated$
    Public LastUpdated$
    Public DepartmentId As Integer
    Public DepartmentName$
    Public Labels$
    Public Version$
    Public IsSystemDepartment As Boolean

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

    Public isBuilt As Boolean
    Public listLinkedSRs As Collection

    Public ReadOnly Property ThrID
        Get
            Return Id
        End Get
    End Property
    Public ReadOnly Property ThrName
        Get
            Return Name
        End Get
    End Property
    Public ReadOnly Property ThRisk
        Get
            Return RiskName
        End Get
    End Property
    Public Sub New()
        listLinkedSRs = New Collection
    End Sub
End Class

Public Class tmProjThreatShow
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

        Public isBuilt As Boolean
        Public listLinkedSRs As Collection

        Public ReadOnly Property ThrID
            Get
                Return Id
            End Get
        End Property
        Public ReadOnly Property ThrName
            Get
                Return Name
            End Get
        End Property
        Public ReadOnly Property ThRisk
            Get
                Return RiskName
            End Get
        End Property
        Public ReadOnly Property Status
            Get
                Return StatusName
            End Get
        End Property



        Public Sub New()
            listLinkedSRs = New Collection
        End Sub
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

Public Class tmModelNode
    Public nodE As tmNodeData


    Public ReadOnly Property compID As Long
        Get
            If IsNothing(nodE.NodeId) Then Return 0 Else Return nodE.NodeId
        End Get
    End Property
    Public ReadOnly Property compKey$
        Get
            Return nodE.compKey
        End Get
    End Property

    Public ReadOnly Property compName As String
        Get
            Return nodE.name
        End Get
    End Property
    Public ReadOnly Property compTypeName As String
        Get
            Return nodE.ComponentTypeName
        End Get
    End Property

    Public ReadOnly Property compFullName As String
        Get
            Return nodE.FullName
        End Get
    End Property
    Public ReadOnly Property compNotes As String
        Get
            Return nodE.Notes
        End Get
    End Property
    Public ReadOnly Property numTH As Integer
        Get
            Return nodE.listThreats.Count
        End Get
    End Property
    Public ReadOnly Property numSR As Integer
        Get
            Return nodE.listSecurityRequirements.Count
        End Get
    End Property

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

    Public lblMatchPct As Decimal
    Public Function numLabels() As Integer
        If Len(Labels) = 0 Then
            Return 0
        Else
            Return numCHR(Labels, ",") + 1
        End If
    End Function
    Public ReadOnly Property SRID
        Get
            Return Id
        End Get
    End Property
    Public ReadOnly Property SRName
        Get
            Return Name
        End Get
    End Property


End Class

    Public Class tmProjSecReqShow
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

        Public lblMatchPct As Decimal
        Public Function numLabels() As Integer
            If Len(Labels) = 0 Then
                Return 0
            Else
                Return numCHR(Labels, ",") + 1
            End If
        End Function
        Public ReadOnly Property SRID
            Get
                Return Id
            End Get
        End Property
        Public ReadOnly Property SRName
            Get
                Return Name
            End Get
        End Property

        Public ReadOnly Property Status
            Get
                Return StatusName
            End Get
        End Property

    End Class


    Public Class tmProps
    Public Id As Integer
    Public Name$
    Public Backend$
    'Public Guid? As Guid
    Public [Type] As String
End Class
Public Class tmTemplate
    Public Id As Long
    Public Name$
    Public LibraryId As Integer
    Public [Type] As String
    Public Json$
    '    "Description": null,
    '    "Labels": "",
    '    "Image": null,
    '    "DiagramData": "data:image/png;base64,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",
    '    "LibraryName": "Mike Horty",
    '    "isUpdate": false
    Public Description$
    Public Labels$
    Public Image$
    Public DiagramData$
    Public LibraryName$
    Public isUpdate As Boolean
    Public ReadOnly Property TempId
        Get
            Return Id
        End Get
    End Property
    Public ReadOnly Property TemplateName
        Get
            Return Name
        End Get
    End Property

End Class

Public Class tmNodeArray
    Public class$
    Public nodeDataArray() As tmNodeInsertData
    '    "Data": {
    '    "Id": 442,
    '    "Name": "SimpleDB_Test",
    '    "LibraryId": 66,
    '    "Type": null,
    '    "Json": "{ \"class\": \"Json\",\n  \"nodeDataArray\": '

    '	[ 
    '  \"key\":-8, \"category\":\"Component\", \"Name\":\"Database\", \"type\":\"Node\"
    '	, \"ImageURI\":\"/ComponentImage/database.png\", \"NodeId\":10, \"ComponentId\":10'
    '	, \"Location\"\"269.8867467026331 126.2433610730271\", \"color\":null, \"BorderCo'lor\":null
    '	, \"BackgroundColor\":null, \"TitleColor\":\"black\", \"TitleBackgroundColor\":\"transparent\"
    '	, \"isGroup\":false, \"isGraphExpanded\":null, \"FullName\":\"Database\", \"cpeid\":\"cpe:/a:oracle:mysql:3.22.26\"
    '	, \"networkComponentId\":null, \"BorderThickness\":0, \"AllowMove\":true, \"LayoutWidth\":null, \"LayoutHeight\":null
    '	, \"ComponentProperties\":null, \"listThreats\":null, \"listSecurityRequirements\":null, \"listDataElements\":null
    '	, \"listProperties\":null, \"listTestCases\":null, \"guid\":\"49d7a54c-3521-4c56-9081-6dc6166a0990\"
    '	, \"AWS_Id\":null, \"ExternalResourceId\":null, \"ExternalResourceARN\":null, \"ResourceType\":null, \"IsSystem\":false
    '	, \"SecurityGroups\":null, \"ThreatsCount\":0, \"SecurityRequirementsCount\":0, \"LibraryId\":1, \"ComponentTypeId\":21
    '	, \"ComponentTypeName\":\"Database\", \"Labels\":\"ThreatModeler,AppSec and InfraSec,Database\", \"ComponentName\":null
    '	, \"BucketPolicy\":null, \"EncryptionEnabled\":false, \"PubliclyAccessible\":false, \"LogsEnabled\":false
    '	, \"VersionEnabled\":false, \"MfaDeleteEnabled\":false, \"ObjectLockEnabled\":false, \"RoleName\":null, \"Policies\":null
    '	, \"InstanceStopped\":false, \"Tags\":null, \"Id\":10, \"Key\":0, \"isProtectedResource\":true, \"isHighValueTarget\":true
    '	, \"isImportThreats\":false, \"Notes\":\"Random notes\", \"threatCount\":0, \"hasCustomProperties\":false
    '	, \"DataElements\":[ 
    '				{\"Id\":8, \"Name\":\"Address\", \"Backend\":null, \"Guid\":\"7edd0871-7d1b-485a-9db4-94ef3202d8ba\", \"Type\":\"DataElements\"} 
    '				]
    '	, \"Roles\":[' {\"Id\":6, \"Name\":\"Admin\", \"Backend\":null, \"Guid\":\"06b9fff9-3451-4f8a-bcd2-0aa213f40cbe\", \"Type\":\"Roles\"} 
    '				]'
    '	, \"Components\":null
    ''	, \"Widgets\":[ {\"Id\":1, \"Name\":\"Form\", \"Backend\":\"XML\", \"Guid\":\"66960410-1056-49c2-94bd-e11eef2bb77f\", \"Type\":\"Widgets\"} '
    ''			']
    '	,' '\"CustomProperties\":null} ''
    '	],''
    '	\n  \"linkDataArray\": []}",
    '    "Description": null,
    '    "Labels": "",
    '    "Image": null,
    '    "DiagramData": "data:image/png;base64,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",
    '    "LibraryName": "Mike Horty",
    '    "isUpdate": false
    '  }

End Class
Public Class tmNodeAction
    Public action$
    Public [type] As String
    Public changedNode As tmNodeInsertData
End Class

Public Class tmNodeInsertData
    '      "SecurityGroups": [],
    '      "Widgets": [],
    '      "DataElements": [],
    '      "Roles": [],
    '      "Components": [],
    '      "CustomProperties": []

    ' perhaps not necessary in the end, but had to create a very strict class (all members, no additions eg listThreats, same sequence) for transaction POSTs

    Public key As Long
    Public name$
    Public FullName$
    Public guid? As Guid
    Public Location$
    Public NodeId? As Long
    Public [type] As String
    Public category$
    Public ImageURI$
    Public ComponentId? As Integer
    Public group? As Long
    Public TitleColor$
    Public TitleBackgroundColor$
    Public BackgroundColor$
    Public BorderThickness? As Decimal
    Public cpeid$ ' As String '$ ' As Integer
    Public SecurityGroups() As tmProps
    Public isHighValueTarget? As Boolean
    Public isImportThreats? As Boolean
    Public isProtectedResource? As Boolean
    Public BorderColor$
    Public ComponentTypeName$
    Public ComponentName$
    Public isGraphExpanded? As Boolean
    Public threatCount As Integer
    Public Notes$
    Public Widgets() As tmProps
    Public DataElements() As tmProps
    Public Roles() As tmProps
    Public Components() As tmProps
    Public CustomProperties() As tmProps
    '    Public networkComponentId? As Integer
    '    Public AllowMove? As Boolean
    '    Public LayoutWidth? As Single
    '    Public LayoutHeight? As Single
    '    Public color$
    '    Public isGroup As Boolean
    '    Public componentProperties$
    '    Public Id As Long
    '    Public Widgets() As tmProps
    '   Public DataElements() As tmProps
    '   Public Roles() As tmProps
    '   Public Components() As tmProps
    '   Public CustomProperties() As tmProps
End Class

Public Class tmNodeInModel
    'node info
    Public compInfo As tmNodeData
    Public modelInfo As tmProjInfo
    Public ReadOnly Property modID As Long
        Get
            Return modelInfo.Id
        End Get
    End Property
    Public ReadOnly Property modName As String
        Get
            Return modelInfo.Name
        End Get
    End Property
    Public ReadOnly Property modNodes As Integer
        Get
            Return modelInfo.Model.Nodes.Count
        End Get
    End Property

    Public ReadOnly Property compKey As Long
        Get
            '     Return compInfo.key
            Return 0
        End Get
    End Property
    Public ReadOnly Property compNodeID As Long
        Get
            Return compInfo.NodeId
        End Get
    End Property
    Public ReadOnly Property compID As Integer
        Get
            Return compInfo.ComponentId
        End Get
    End Property
    ' Public ReadOnly Property compName
    '     Get
    '         Return compInfo.name
    '     End Get
    ' End Property
    Public ReadOnly Property compFullName As String
        Get
            Return compInfo.FullName
        End Get
    End Property
    '  Public ReadOnly Property compType
    '      Get
    '          Return compInfo.type
    '      End Get
    '  End Property
    Public ReadOnly Property compCategory As String
        Get
            Return compInfo.category
        End Get
    End Property
    Public ReadOnly Property compGroup As Long
        Get
            Return compInfo.group
        End Get
    End Property
    Public ReadOnly Property compNotes As String
        Get
            Return compInfo.Notes
        End Get
    End Property

    'likely interesting/ non-int removed
    '    Public key As Long
    '    Public name$
    '    Public FullName$
    '    Public NodeId? As Long
    '    Public [type] As String
    '    Public category$
    '    Public ComponentId? As Integer
    '    Public group? As Long
    '    Public cpeid$ ' As String '$ ' As Integer
    '    Public SecurityGroups() As tmProps
    '    Public isHighValueTarget? As Boolean
    '    Public isImportThreats? As Boolean
    '    Public isProtectedResource? As Boolean
    '    Public ComponentTypeName$
    '    Public ComponentName$
    '    Public Notes$
    '    Public Widgets() As tmProps
    '    Public DataElements() As tmProps
    '    Public Roles() As tmProps
    '    Public Components() As tmProps
    '    Public CustomProperties() As tmProps

    'model info
    '                        T.addThreats(P.Model, P.Id) 'add in details of threats not inside Diagram API (eg Protocols)
    '                    If louD Then Form1.addLOG(col5CLI(P.Name + " [Id " + P.Id.ToString + "]", .Nodes.Count.ToString, P.Model.modelNumThreats.ToString, P.Model.modelNumThreats(True).ToString))
    '                    '     form1.addlog(col5CLI("   ", "TYPE", "# THR", "# SR"))
    '                    For Each N In .Nodes
    '    If N.category = "Collection" Then
    '    If louD Then Form1.addLOG(col5CLI(" >" + N.FullName, "Collection", "", ""))
    '         '                   GoTo doneHere
    '    End I'f
    '    If N.category = "Project" Then '
    '                           N.ComponentName = N.FullName
    '                            N.ComponentTypeName = "Nested Model"



End Class

Public Class tmNodeData
    '      "SecurityGroups": [],
    '      "Widgets": [],
    '      "DataElements": [],
    '      "Roles": [],
    '      "Components": [],
    '      "CustomProperties": []

    Public Property [key] As System.Int64
        Get
            Return nodeKey
        End Get
        Set(value As System.Int64)
            nodeKey = value
            keySet.Add(value.ToString)
        End Set
    End Property
    Public nodeKey As System.Int64
    Public category$
    Public name$
    Public [type] As String
    'Public guid? As Guid
    Public ImageURI$
    Public NodeId? As Long
    Public ComponentId? As Integer
    Public group? As Long
    Public Location$
    Public color$
    Public isHighValueTarget? As Boolean
    Public isImportThreats? As Boolean
    Public isProtectedResource? As Boolean
    Public BorderColor$
    Public BackgroundColor$
    Public TitleColor$
    Public TitleBackgroundColor$
    Public isGroup As Boolean
    Public isGraphExpanded? As Boolean
    Public FullName$
    Public cpeid$ ' As String '$ ' As Integer
    Public networkComponentId? As Integer
    Public BorderThickness? As Decimal
    Public AllowMove? As Boolean
    Public LayoutWidth? As Single
    Public LayoutHeight? As Single
    Public componentProperties$
    Public ComponentTypeName$
    Public ComponentName$
    Public Notes$
    Public Id As Long
    Public threatCount As Integer
    Public Widgets() As tmProps
    Public DataElements() As tmProps
    Public Roles() As tmProps
    Public Components() As tmProps
    Public CustomProperties() As tmProps

    Public keySet As Collection
    Public listThreats As List(Of tmProjThreat)
    Public listSecurityRequirements As List(Of tmProjSecReq)

    Public ReadOnly Property compKey$
        Get
            If keySet.Count Then Return keySet(1) Else Return 0
        End Get
    End Property
    Public Sub New()
        Me.listSecurityRequirements = New List(Of tmProjSecReq)
        Me.listThreats = New List(Of tmProjThreat)
        Me.keySet = New Collection
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

