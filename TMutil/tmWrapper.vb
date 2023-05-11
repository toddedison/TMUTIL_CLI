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

    Public lib_Comps As List(Of tmComponent)
    Public lib_TH As List(Of tmProjThreat)
    Public lib_SR As List(Of tmProjSecReq)
    Public lib_AT As List(Of tmAttribute)

    Public lib_Comps6 As List(Of tm6Component)
    Public lib_TH6 As List(Of tm6Threat)
    Public lib_SR6 As List(Of tm6SecReq)
    Public lib_AT6 As List(Of tm6Attribute)
    Public lib_CompStructures As List(Of tm6CompStructure)

    Public sysLabels As List(Of tmLabels)
    Public labelS As List(Of tmLabels)
    Public groupS As List(Of tmGroups)
    Public librarieS As List(Of tmLibrary)
    Public componentTypes As List(Of tmMiscTrigger)
    Public roleS As List(Of tmMiscTrigger)
    Public dataElements As List(Of tmMiscTrigger)

    Public isTMsix As Boolean

    Public Sub New(fqdN$, uN$, pw$)
        isTMsix = False

        tmFQDN = fqdN

        Me.isConnected = False


        If Len(uN) > 0 And Len(pw) = 0 Then
            isTMsix = True
            Console.WriteLine("6.0 credentials provided")
            slToken = uN ' the api key 6.0 is now associated with 5.5 bearer token
            Dim uCheck As currUser = New currUser
            Me.isConnected = True
            uCheck = getCurrentUser()

            Dim probError As Boolean = False

            If IsNothing(uCheck) = True Then
                probError = True
            Else
                If uCheck.email = "" Then probError = True
            End If

            If probError = True Then
                Me.isConnected = False
                Console.WriteLine("Unable to access with API KEY - are you sure " + fqdN + " is running 6.0?")
                Exit Sub
            Else
                Me.isConnected = True
                Console.WriteLine("Connecting as " + uCheck.userName + " to " + fqdN)
                Exit Sub
            End If
        End If
        'Console.WriteLine("New TM_Client activated: " + Replace(tmFQDN, "https://", ""))

        Console.WriteLine("Connecting to " + fqdN)
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
        Console.WriteLine("Login non-SSO method (demo2): " + response.ResponseStatus.ToString + "/ " + response.StatusCode.ToString + "/ Success: " + response.IsSuccessful.ToString)

        If response.IsSuccessful = False Then
            client = New RestClient(fqdN + "/idsvr/connect/token")
            request = New RestRequest(Method.POST)
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded")
            request.AddHeader("Accept", "application/json")

            request.AddParameter("username", uN)
            request.AddParameter("password", pw)
            request.AddParameter("grant_type", "password")
            request.AddParameter("client_id", "demo-resource-owner")
            request.AddParameter("client_secret", "geheim")
            request.AddParameter("scope", "openid profile email api")

            response = client.Execute(request)
            Console.WriteLine("Login SSO method (/idsvr/connect): " + response.ResponseStatus.ToString + "/ " + response.StatusCode.ToString + "/ Success: " + response.IsSuccessful.ToString)

        End If

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

    Private Function tfAdd_GetAPIData(ByVal urI$, addJSONbody$)

        ' DO NOT NEED THIS FUNCTION...

        tfAdd_GetAPIData = ""
        Dim client = New RestClient(tmFQDN + urI)
        Dim request As RestRequest
        request = New RestRequest(Method.POST)

        Dim response As IRestResponse
        client.Timeout = -1

        Dim c$ = Chr(34)

        'var request = New RestRequest(Method.POST);
        request.AddHeader("sec-ch-ua", c + "Google Chrome" + c + ";v=" + c + "95" + c + ", " + c + "Chromium" + c + ";v=" + c + "95" + c + ", " + c + ";Not A Brand" + c + ";v=" + c + "99" + c)
        request.AddHeader("sec-ch-ua-mobile", "?0")
        request.AddHeader("Authorization", "Bearer " + slToken)
        request.AddHeader("Content-Type", "multipart/form-data") '; boundary=----WebKitFormBoundary1bCpiyr5KkmiBIAJ")
        request.AddHeader("Accept", "application/json, text/plain, */*")
        client.UserAgent = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36"
        request.AddHeader("sec-ch-ua-platform", c + "macOS" + c)
        request.AlwaysMultipartFormData = True
        request.AddParameter("data", addJSONbody)

        response = client.Execute(request)


    End Function
    Private Function getAPIData(ByVal urI$, Optional ByVal usePOST As Boolean = False, Optional ByVal addJSONbody$ = "", Optional ByVal addingComp As Boolean = False, Optional ByVal addFileN$ = "", Optional ByVal thisIsPATCH As Boolean = False, Optional ByVal thisIsPUT As Boolean = False) As String
        getAPIData = ""
        If isConnected = False Then Exit Function

        Dim debugInfo As Boolean = False

        Dim client = New RestClient(tmFQDN + urI)
        Dim request As RestRequest

        If thisIsPUT = False And thisIsPATCH = False And usePOST = False Then
            request = New RestRequest(Method.GET)
        Else
            If usePOST = True Then request = New RestRequest(Method.POST)
            If thisIsPATCH = True Then request = New RestRequest(Method.PATCH)
            If thisIsPUT = True Then request = New RestRequest(Method.PUT)
        End If

        If debugInfo Then
            Dim gOrP$ = "GET"
            If usePOST Then gOrP = "POST"
            If thisIsPATCH Then gOrP = "PATCH"
            If thisIsPUT Then gOrP = "PUT"
            Console.WriteLine(vbCrLf + "------API CALL------" + vbCrLf + gOrP + ": " + tmFQDN + urI)
            'If usePOST Then Console.WriteLine("JSON Body: " + addJSONbody)
        End If

        Dim response As IRestResponse

        If isTMsix = False Then
            request.AddHeader("Authorization", "Bearer " + slToken)
        Else
            request.AddHeader("X-ThreatModeler-ApiKey", slToken)
        End If
        request.AddHeader("Accept", "application/json")

        If Len(addJSONbody) Then
            If addingComp = False Then
                request.AddHeader("Content-Type", "application/json")
            Else
                request.AddHeader("Content-Type", "multipart/form-data") '; boundary=----WebKitFormBoundary1bCpiyr5KkmiBIAJ")
            End If
            request.AddHeader("Accept-Encoding", "gzip, deflate, br")

            'request.AddHeader("Accept-Language", "zh")
            'adding this (Lenovo Chinese) messes up API calls for COmponentTypes and Attributes (prob more)

            request.AddHeader("Connection", "keep-alive")
            'If isTMsix = True Then
            '    request.AddJsonBody(addJSONbody)
            ' Else
            If addingComp = False Then
                request.AddParameter("application/json", addJSONbody, ParameterType.RequestBody)
            Else
                request.AddParameter("data", addJSONbody, ParameterType.RequestBody)
                request.AlwaysMultipartFormData = True
            End If
            'End If
        End If

        If debugInfo Then
            For Each hh In request.Parameters
                Dim hVal$ = hh.Value
                If Mid(hh.Value, 1, 7) = "Bearer " Then hVal = "Bearer redacted" Else hVal = hh.Value
                Console.WriteLine(hh.Name + ":" + hVal)
            Next
        End If

        If Len(addFileN) Then
            request.AddFile("File", addFileN)
            If debugInfo Then Console.WriteLine("File added: " + addFileN)
        End If

        response = client.Execute(request)

        If addingComp Or thisIsPUT Then
            Return response.Content
        End If

        If Len(response.Content) = 0 Then
            If debugInfo Then Console.WriteLine("No response provided")
            getAPIData = ""
            Exit Function
        End If

        Dim O As JObject = JObject.Parse(response.Content)

        If debugInfo Then
            Console.WriteLine(vbCrLf + "RESPONSE:" + vbCrLf + response.Content)
        End If

        Dim dStr = "Data"
        Dim isSuccessStr$ = "IsSuccess"

        If isTMsix = True Then
            isSuccessStr = "isSuccess"
            dStr = "data"
        End If

        If IsNothing(O.SelectToken(isSuccessStr)) = True Then
            ' here check for version
            If isTMsix = True Then
                If IsNothing(O.SelectToken("IsSuccess")) = False Then
                    ' this is a call with unexpected capitalization
                    isSuccessStr = "IsSuccess"
                    dStr = "Data"
                    GoTo keepGoing
                End If
            End If
            Console.WriteLine("API Request Rejected:  " + urI)
            getAPIData = ""
            Exit Function
        End If

keepGoing:

        '        If Len(addFileN) Then
        '            If O.SelectToken("IsSuccess").ToString = "true" Then
        '                Return O.SelectToken("IsSuccess").ToString
        '            Else
        '                Return response.Content
        '            End If
        '        End If

        If CBool(O.SelectToken(isSuccessStr)) = False Then
            getAPIData = "ERROR:Could not retrieve " + urI
            Exit Function
        End If

        If IsNothing(O.SelectToken(dStr)) = True Then
            If IsNothing(O.SelectToken(isSuccessStr)) = True Then
                Return ""
            Else
                Return O.SelectToken(isSuccessStr).ToString
            End If
        Else
            Return O.SelectToken(dStr).ToString
        End If
    End Function

    Public Function getMatildaDiscover(fileN$) As matildaApp
        getMatildaDiscover = New matildaApp

        Dim jsoN$ = streamReaderTxt(fileN)
        If jsoN = "" Then Exit Function

        getMatildaDiscover = JsonConvert.DeserializeObject(Of matildaApp)(jsoN)
    End Function

    Public Function getRonnieIssues(fileN$) As List(Of ronnieIssue)
        getRonnieIssues = New List(Of ronnieIssue)
        Dim jsoN$ = streamReaderTxt(fileN)
        If jsoN = "" Then Exit Function

        getRonnieIssues = JsonConvert.DeserializeObject(Of List(Of ronnieIssue))(jsoN)
    End Function
    Public Function getLibraries() As List(Of tmLibrary)
        getLibraries = New List(Of tmLibrary)
        Dim jsoN$ = getAPIData("/api/libraries")

        If jsoN = "" Then
            Return getLibraries 'returning an empty object
            Exit Function
        End If

        getLibraries = JsonConvert.DeserializeObject(Of List(Of tmLibrary))(jsoN)

    End Function

    Public Function getEntityMisc(T As tfRequest) As List(Of tmMiscTrigger)
        Dim TF As New List(Of tmMiscTrigger)

        Dim jBody$ = ""
        jBody = JsonConvert.SerializeObject(T)

        Dim jsoN$ = getAPIData("/api/threatframework", True, jBody$)

        TF = JsonConvert.DeserializeObject(Of List(Of tmMiscTrigger))(jsoN)

        Return TF
    End Function

    Public Function updateThreatStat(T As threatStatusUpdate) As String
        updateThreatStat = ""

        Dim jBody$ = ""
        jBody = JsonConvert.SerializeObject(T)

        Return jBody

        Exit Function

        Dim jsoN$ = getAPIData("/api/threatframework", True, jBody$)

        ' TF = JsonConvert.DeserializeObject(Of List(Of tmMiscTrigger))(jsoN)

        ' Return TF
    End Function

    Public Function setNewSRstatus(SR As tmupdateSR) As Boolean
        setNewSRstatus = False
        Dim jBody$ = ""
        jBody = "[" + JsonConvert.SerializeObject(SR) + "]"


        Dim json$ = getAPIData("/api/securityrequirement/status", True, jBody)
        ' Console.WriteLine(json)
        setNewSRstatus = True
    End Function

    Public Function getGroups() As List(Of tmGroups)
        getGroups = New List(Of tmGroups)
        Dim jsoN$ = getAPIData("/api/groups")

        getGroups = JsonConvert.DeserializeObject(Of List(Of tmGroups))(jsoN)

    End Function
    Public Function getGroups6() As List(Of tmGroups6)
        getGroups6 = New List(Of tmGroups6)
        Dim jsoN$ = getAPIData("/api/groups")

        getGroups6 = JsonConvert.DeserializeObject(Of List(Of tmGroups6))(jsoN)

    End Function

    Public Function getUsersSIX() As List(Of tmUser)
        getUsersSIX = New List(Of tmUser)
        Dim jsoN$ = getAPIData("/api/users")

        getUsersSIX = JsonConvert.DeserializeObject(Of List(Of tmUser))(jsoN)

    End Function

    '  /api/library/libraries
    Public Function getLibsSIX() As List(Of tmLibrary)
        getLibsSIX = New List(Of tmLibrary)
        Dim jsoN$ = getAPIData("/api/library/libraries")

        getLibsSIX = JsonConvert.DeserializeObject(Of List(Of tmLibrary))(jsoN)

    End Function

    Public Function getCurrentUser() As currUser
        getCurrentUser = New currUser
        Dim jsoN$ = getAPIData("/api/user/loginuserdetails")

        getCurrentUser = JsonConvert.DeserializeObject(Of currUser)(jsoN)

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
        '        Dim jsoN$ = getAPIData("/api/projects/GetAllProjects")
        Dim c$ = Chr(34)
        Dim bodY$ = "{" + c + "virtualGridStateRequestModel" + c + ":{" + c + "state" + c + ":{" + c + "skip" + c + ":0," + c + "take" + c + ":5000}}," + c + "projectFilterModel" + c + ":{" + c + "ActiveProjects" + c + ":true}}"

        Dim jsoN$ = getAPIData("/api/projects/smartfilter", True, bodY)

        Dim O As JObject = JObject.Parse(jsoN)
        jsoN = O.SelectToken("Data").ToString


        getAllProjects = JsonConvert.DeserializeObject(Of List(Of tmProjInfo))(jsoN)

    End Function

    Public Function getSRsOfProject6(guiD$) As List(Of tm6SRsOfModel)
        getSRsOfProject6 = New List(Of tm6SRsOfModel)

        Dim bodY$ = "{" + qT("guid") + ": " + qT(guiD) + "," + qT("pageNumber") + ": 1," + qT("pageLimit") + ": 10000}"
        Dim jsoN$ = getAPIData("/api/project/securityrequirements", True, bodY) '/smartfilter", True, bodY)

        Dim O As JObject = JObject.Parse(jsoN)
        jsoN = O.SelectToken("data").ToString

        getSRsOfProject6 = JsonConvert.DeserializeObject(Of List(Of tm6SRsOfModel))(jsoN)

    End Function

    Public Function getSRsOfThreat6(guiD$, threatID As Long) As List(Of tm6SRsOfModel)

        getSRsOfThreat6 = New List(Of tm6SRsOfModel)

        Dim jsoN$ = getAPIData("/api/project/" + guiD + "/allsecurityrequirements?ProjectThreatIds=" + threatID.ToString)

        getSRsOfThreat6 = JsonConvert.DeserializeObject(Of List(Of tm6SRsOfModel))(jsoN)

    End Function

    Public Function getThreatsOfProject6(guiD$) As List(Of tm6ThreatsOfModel)
        getThreatsOfProject6 = New List(Of tm6ThreatsOfModel)

        Dim bodY$ = "{" + qT("pageNumber") + ": 1," + qT("pageLimit") + ": 10000}"
        Dim jsoN$ = getAPIData("/api/project/" + guiD + "/threats", True, bodY)

        Dim O As JObject = JObject.Parse(jsoN)
        jsoN = O.SelectToken("data").ToString

        getThreatsOfProject6 = JsonConvert.DeserializeObject(Of List(Of tm6ThreatsOfModel))(jsoN)

    End Function

    Public Function getNodesOfProject6(guiD$) As List(Of tm6NodesOfModel)
        getNodesOfProject6 = New List(Of tm6NodesOfModel)

        Dim jsoN$ = getAPIData("/api/diagram/" + guiD) '/smartfilter", True, bodY)

        Dim O As JObject = JObject.Parse(jsoN)
        jsoN = O.SelectToken("Model").SelectToken("nodeDataArray").ToString

        getNodesOfProject6 = JsonConvert.DeserializeObject(Of List(Of tm6NodesOfModel))(jsoN)

    End Function


    Public Function getAllProjects6() As List(Of tmProjInfoShort6)
        getAllProjects6 = New List(Of tmProjInfoShort6)

        Dim bodY$ = "{" + qT("pageNumber") + ": 1," + qT("pageLimit") + ": 1000}"
        Dim jsoN$ = getAPIData("/api/project/projects", True, bodY) '/smartfilter", True, bodY)

        'Dim O As JObject = JObject.Parse(jsoN)
        'jsoN = O.SelectToken("data").ToString


        getAllProjects6 = JsonConvert.DeserializeObject(Of List(Of tmProjInfoShort6))(jsoN)

    End Function

    Public Function getProjActivity(projGuid$, Optional ByVal fromDate As DateTime = #1/1/2022#) As List(Of tmProjActivity6)
        getProjActivity = New List(Of tmProjActivity6)

        Dim datelimit$ = ""
        If fromDate <> #1/1/2022# Then datelimit = "?actionDateFrom=" + fromDate.ToString
        Dim jsoN$ = getAPIData("/api/project/" + projGuid + "/getalltransactionlogevent/1/0" + datelimit) '?actionDateFrom=2022-01-01")

        getProjActivity = JsonConvert.DeserializeObject(Of List(Of tmProjActivity6))(jsoN)
    End Function
    Public Function getPermissionsOfModel(projID As Long) As tm_userGroupPermissions
        getPermissionsOfModel = New tm_userGroupPermissions

        Dim jsoN$ = getAPIData("/api/project/" + projID.ToString + "/sharedgroupsusers")
        '        Dim c$ = Chr(34)

        'Dim grpStr$ = Mid(jsoN, 2, InStr(jsoN, "],") - 1)
        'jsoN = Mid(jsoN, InStr(jsoN, "],") + 1)
        'Dim usrStr$ = Mid(jsoN, InStr(jsoN, ",") + 1, InStr(jsoN, "]") - 1)

        getPermissionsOfModel = JsonConvert.DeserializeObject(Of tm_userGroupPermissions)(jsoN)

    End Function


    Public Function getProjThreats(projId As Integer) As List(Of tmTThreat)
        Dim tNodes As List(Of tmTThreat) = New List(Of tmTThreat)
        Dim jsoN$ = getAPIData("/api/project/" + projId.ToString + "/threats?openStatusOnly=false")
        tNodes = JsonConvert.DeserializeObject(Of List(Of tmTThreat))(jsoN)
        Return tNodes
    End Function
    Public Function getProjSecReqs(projId As Integer) As List(Of tmSimpleThreatSR)
        Dim tNodes As List(Of tmSimpleThreatSR) = New List(Of tmSimpleThreatSR)
        Dim jsoN$ = getAPIData("/api/project/" + projId.ToString + "/securityrequirements?openStatusOnly=false")
        tNodes = JsonConvert.DeserializeObject(Of List(Of tmSimpleThreatSR))(jsoN)
        Return tNodes
    End Function

    Public Function getLabels(Optional ByVal isSystem As Boolean = False) As List(Of tmLabels)
        getLabels = New List(Of tmLabels)
        Dim jsoN$ = ""
        If isTMsix = False Then
            jsoN = getAPIData("/api/labels?isSystem=" + LCase(CStr(isSystem))) 'true") '?isSystem=false")
        Else
            jsoN = getAPIData("/api/library/labels") 'true") '?isSystem=false")
        End If

        getLabels = JsonConvert.DeserializeObject(Of List(Of tmLabels))(jsoN)
    End Function
    Public Function getTemplates() As List(Of tmTemplate)
        getTemplates = New List(Of tmTemplate)
        Dim jsoN$ = getAPIData("/api/templates") '/" + tempID.ToString)

        getTemplates = JsonConvert.DeserializeObject(Of List(Of tmTemplate))(jsoN)

    End Function

    Public Function getTemplateSIX(tempID As Integer) As List(Of tmTemplate6)
        getTemplateSIX = New List(Of tmTemplate6)
        Dim jsoN$ = getAPIData("/api/template/detail", True, "[" + tempID.ToString + "]") '/" + tempID.ToString)

        getTemplateSIX = JsonConvert.DeserializeObject(Of List(Of tmTemplate6))(jsoN)
    End Function

    Public Function getTemplateList() As List(Of tmTemplate)
        getTemplateList = New List(Of tmTemplate)
        Dim jsoN$ = getAPIData("/api/template/templates") '/" + tempID.ToString)

        getTemplateList = JsonConvert.DeserializeObject(Of List(Of tmTemplate))(jsoN)


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

    Public Sub createVPCmodel(G As tmVPC)
        On Error Resume Next
        'this command times out - should be async

        Call getAPIData("/api/thirdparty/createthreatmodelfromvpc", True, prepVPCmodel(G))
    End Sub


    Public Function getTFComponents(T As tfRequest) As List(Of tmComponent)
        getTFComponents = New List(Of tmComponent)

        Dim jBody$ = ""
        Dim urL$ = "/api/threatframework"

        If isTMsix Then
            Dim tF6 As tfRequest6 = New tfRequest6
            With tF6
                .entityTypeName = "Component"
                .showHiddenOnly = False
                .libraryId = 0
            End With
            jBody = JsonConvert.SerializeObject(tF6)
            urL = "/api/library/getrecords"
        Else
            jBody = JsonConvert.SerializeObject(T)
        End If

        Dim jsoN$ = getAPIData(urL, True, jBody$)

        getTFComponents = JsonConvert.DeserializeObject(Of List(Of tmComponent))(jsoN)
    End Function

    Public Function getTF6Components() As List(Of tm6Component)
        getTF6Components = New List(Of tm6Component)

        Dim jBody$ = ""
        Dim urL$ = "/api/library/getrecords"

        Dim tF6 As tfRequest6 = New tfRequest6
        With tF6
            .entityTypeName = "Component"
            .showHiddenOnly = False
            .libraryId = 0
        End With
        jBody = JsonConvert.SerializeObject(tF6)

        Dim jsoN$ = getAPIData(urL, True, jBody$)

        getTF6Components = JsonConvert.DeserializeObject(Of List(Of tm6Component))(jsoN)
    End Function
    Public Function getTF6SingleComp(rID as integer) As tm6Component
        getTF6SingleComp = New tm6Component
        Dim c$ = Chr(34)
        Dim jBody$ = "{id: " + rID.ToString + ",entityTypeName: " + c + "component" + c + "}"
        Dim urL$ = "/api/library/getrecordbyid"

        Dim jsoN$ = getAPIData(urL, True, jBody$)

        getTF6SingleComp = JsonConvert.DeserializeObject(Of tm6Component)(jsoN)

    End Function
    Public Function getTF6Threats() As List(Of tm6Threat)
        getTF6Threats = New List(Of tm6Threat)

        Dim jBody$ = ""
        Dim urL$ = "/api/library/getrecords"

        Dim tF6 As tfRequest6 = New tfRequest6
        With tF6
            .entityTypeName = "Threat"
            .showHiddenOnly = False
            .libraryId = 0
        End With
        jBody = JsonConvert.SerializeObject(tF6)

        Dim jsoN$ = getAPIData(urL, True, jBody$)

        getTF6Threats = JsonConvert.DeserializeObject(Of List(Of tm6Threat))(jsoN)
    End Function

    Public Function getTF6Requirements() As List(Of tm6SecReq)
        getTF6Requirements = New List(Of tm6SecReq)

        Dim jBody$ = ""
        Dim urL$ = "/api/library/getrecords"

        Dim tF6 As tfRequest6 = New tfRequest6
        With tF6
            .entityTypeName = "SecurityRequirement"
            .showHiddenOnly = False
            .libraryId = 0
        End With
        jBody = JsonConvert.SerializeObject(tF6)

        Dim jsoN$ = getAPIData(urL, True, jBody$)

        getTF6Requirements = JsonConvert.DeserializeObject(Of List(Of tm6SecReq))(jsoN)
    End Function

    Public Function getTF6Attributes() As List(Of tm6Attribute)
        getTF6Attributes = New List(Of tm6Attribute)

        Dim jBody$ = ""
        Dim urL$ = "/api/library/getrecords"

        Dim tF6 As tfRequest6 = New tfRequest6
        With tF6
            .entityTypeName = "Property"
            .showHiddenOnly = False
            .libraryId = 0
        End With
        jBody = JsonConvert.SerializeObject(tF6)

        Dim jsoN$ = getAPIData(urL, True, jBody$)

        getTF6Attributes = JsonConvert.DeserializeObject(Of List(Of tm6Attribute))(jsoN)
    End Function

    Public Function getTF6CompDef(ByRef COMP As tm6Component) As tm6Component
        getTF6CompDef = COMP
        Dim compDEF As tm6CompStructure = New tm6CompStructure

        Dim c$ = Chr(34)
        Dim jBody$ = ""

        jBody$ = "{" + c + "entityTypeName" + c + ":" + c + "Component" + c + "," + c + "id" + c + ":" + COMP.id.ToString + "}"

        Dim urL$ = "/api/library/getentityreltionshipsrecords"

        Dim jsoN$ = getAPIData(urL, True, jBody$)

        If Mid(jsoN, 1, 6) = "ERROR:" Then
            getTF6CompDef = Nothing
            Console.WriteLine("ERROR: " + jsoN)
            Exit Function
        End If

        compdef = JsonConvert.DeserializeObject(Of tm6CompStructure)(jsoN)

        COMP.classDef = compDEF
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


    Public Function getCompFrameworks() As List(Of complianceDetails)
        Dim TF As New List(Of complianceDetails)

        Dim jsoN$ = getAPIData("/api/scf/complianceframeworks/false")

        TF = JsonConvert.DeserializeObject(Of List(Of complianceDetails))(jsoN)

        Return TF
    End Function


    Public Function getTFAttr(T As tfRequest) As List(Of tmAttribute)
        Dim TF As New List(Of tmAttribute)

        Dim jBody$ = ""
        jBody = JsonConvert.SerializeObject(T)

        Dim jsoN$ = getAPIData("/api/threatframework", True, jBody$)

        'Console.WriteLine(jsoN)

        TF = JsonConvert.DeserializeObject(Of List(Of tmAttribute))(jsoN)

        Return TF

        '        Dim nD As JObject = JObject.Parse(jsoN)
        '        Dim nodeJSON$ = nD.SelectToken("Model").ToString
        '
        ''       Dim nodeS As List(Of tmNodeData)
        '       nodeJSON = cleanJSON(nodeJSON)
        '        nodeJSON = cleanJSONright(nodeJSON)
        '
        '        Dim SS As New JsonSerializerSettings
        '        SS.NullValueHandling = NullValueHandling.Ignore
        '
        '        nodeS = JsonConvert.DeserializeObject(Of List(Of tmNodeData))(nodeJSON, SS)
        '
        ' getProject = JsonConvert.DeserializeObject(Of tmModel)(jsoN)
        ' getProject.Nodes = nodeS


    End Function

    Public Sub removeThreatFromComponent(C As tmComponent, threatID As Integer)
        '{"propertyIds":[101],"propertyEntityType":"Threats","sourceEntityIds":[449],"LibraryId":1,"EntityType":"Components","IsAddition":false}
        Dim P As New entityUpdate
        With P
            .propertyIds.Add(threatID)
            .propertyEntityType = "Threats"
            .sourceEntityIds.Add(C.Id)
            .LibraryId = C.LibraryId
            .EntityType = "Components"
            .IsAddition = False
        End With

        Dim jBody$ = JsonConvert.SerializeObject(P)
        Dim json$ = getAPIData("/api/threatframework/updateproperty", True, jBody)


    End Sub
    Public Sub removeAttributeFromComponent(C As tmComponent, attrID As Integer)
        '{"propertyIds":[1667],"propertyEntityType":"Attributes","sourceEntityIds":[3301],"LibraryId":66,"EntityType":"Components","IsAddition":false}
        Dim P As New entityUpdate
        With P
            .propertyIds.Add(attrID)
            .propertyEntityType = "Attributes"
            .sourceEntityIds.Add(C.Id)
            .LibraryId = C.LibraryId
            .EntityType = "Components"
            .IsAddition = False
        End With

        Dim jBody$ = JsonConvert.SerializeObject(P)
        Dim json$ = getAPIData("/api/threatframework/updateproperty", True, jBody)

    End Sub

    Public Sub addSRtoThreat(T As tmProjThreat, srID As Integer, Optional ByVal removeInstead As Boolean = False)
        '{"propertyIds"[116],"propertyEntityType":"SecurityRequirements","sourceEntityIds":[993],"LibraryId":1,"EntityType":"Threats","IsAddition":false}

        Dim P As New entityUpdate
        With P
            .propertyIds.Add(srID)
            .propertyEntityType = "SecurityRequirements"
            .sourceEntityIds.Add(T.Id)
            .LibraryId = T.LibraryId
            .EntityType = "Threats"
            If removeInstead Then .IsAddition = False Else .IsAddition = True
        End With

        Dim jBody$ = JsonConvert.SerializeObject(P)
        Dim json$ = getAPIData("/api/threatframework/updateproperty", True, jBody)

    End Sub

    Public Sub addAttributeToComponent(C As tmComponent, attrID As Integer)
        '{"propertyIds":[1667],"propertyEntityType":"Attributes","sourceEntityIds":[3301],"LibraryId":66,"EntityType":"Components","IsAddition":false}
        Dim P As New entityUpdateBID
        With P
            .propertyIds.Add(attrID)
            .propertyEntityType = "Attributes"
            .sourceEntityIds.Add(C.Id)
            .LibraryId = C.LibraryId
            .EntityType = "Components"
            .IsAddition = True
            .BackendId = 1
        End With

        Dim jBody$ = JsonConvert.SerializeObject(P)
        Dim json$ = getAPIData("/api/threatframework/updateproperty", True, jBody)

    End Sub

    Public Sub addThreatToComponent(C As tmComponent, threatID As Integer)
        '{"propertyIds":[1503],"propertyEntityType":"Threats","sourceEntityIds":[3301],"LibraryId":66,"EntityType":"Components","IsAddition":true,"BackendId":1}
        Dim P As New entityUpdateBID
        With P
            .propertyIds.Add(threatID)
            .propertyEntityType = "Threats"
            .sourceEntityIds.Add(C.Id)
            .LibraryId = C.LibraryId
            .EntityType = "Components"
            .IsAddition = True
            .BackendId = 1
        End With

        Dim jBody$ = JsonConvert.SerializeObject(P)
        Dim json$ = getAPIData("/api/threatframework/updateproperty", True, jBody)


    End Sub

    Public Function getUsers(deptID As Integer) As List(Of tmUser)
        getUsers = New List(Of tmUser)
        Dim jbody$ = "[" + deptID.ToString + "]"
        Dim json$ = getAPIData("/api/department/users", True, jbody)
        getUsers = JsonConvert.DeserializeObject(Of List(Of tmUser))(json)
    End Function

    Public Function getUsersOfGroup(grp As tmGroups) As List(Of tmUserOfGroup)
        getUsersOfGroup = New List(Of tmUserOfGroup)
        Dim jBody$ = JsonConvert.SerializeObject(grp)
        Dim json$ = getAPIData("/api/group/users", True, jBody)
        getUsersOfGroup = JsonConvert.DeserializeObject(Of List(Of tmUserOfGroup))(json)
    End Function

    Public Function removeUserFromGroup(removeReq As removeUserFromGroupReq) As Boolean
        removeUserFromGroup = False
        Dim jBody$ = JsonConvert.SerializeObject(removeReq)
        Dim json$ = getAPIData("/api/group/removeusers", True, jbody)

        If InStr(json, "ERROR") Then removeUserFromGroup = False Else removeUserFromGroup = True

    End Function

    Public Function transferUser(moveUser As moveUserToDeptReq) As Boolean
        transferUser = False
        Dim jBody$ = JsonConvert.SerializeObject(moveUser)
        Dim json$ = getAPIData("/api/user/update", True, jBody,,, True)

        If InStr(json, "ERROR") Then transferUser = False Else transferUser = True

    End Function

    Public Function deactivateUser(cancelUser As deactivateReq) As Boolean
        deactivateUser = False
        Dim jBody$ = JsonConvert.SerializeObject(cancelUser)
        Dim json$ = getAPIData("/api/user/deactivate", True, jBody,,, True)

        If InStr(json, "ERROR") Then deactivateUser = False Else deactivateUser = True

    End Function

    Public Function getDepartments() As List(Of tmDept)
        getDepartments = New List(Of tmDept)
        Dim json$ = getAPIData("/api/departments")
        getDepartments = JsonConvert.DeserializeObject(Of List(Of tmDept))(json)
    End Function

    Public Function isUserInList(userID As Integer, listOfUsers As List(Of tmUserOfGroup)) As Integer
        isUserInList = -1

        Dim ndX = 0

        For Each L In listOfUsers
            If L.UserId = userID Then
                isUserInList = ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function
    Public Function ndxUser(userID As Integer, listOfUsers As List(Of tmUser)) As Integer
        ndxUser = -1

        Dim ndX = 0

        For Each L In listOfUsers
            If L.Id = userID Then
                ndxUser = ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function



    Public Sub addThreatToAttribute(A As tmAttribute, threatID As Integer)
        '{"propertyIds":[1503],"propertyEntityType":"Threats","sourceEntityIds":[3301],"LibraryId":66,"EntityType":"Components","IsAddition":true,"BackendId":1}
        Dim P As New entityUpdateBID
        With P
            .propertyIds.Add(threatID)
            .propertyEntityType = "Threats"
            .sourceEntityIds.Add(A.Id)
            .LibraryId = A.LibraryId
            .EntityType = "Attributes"
            .IsAddition = True
            '.BackendId = 1
        End With

        Dim jBody$ = JsonConvert.SerializeObject(P)
        '        Console.WriteLine(jBody)
        Dim json$ = getAPIData("/api/threatframework/updateproperty", True, jBody)

    End Sub
    Public Sub createUser(Username$, Optional ByVal Name$ = "", Optional ByVal Email$ = "", Optional ByVal sendEmail As Boolean = True, Optional ByVal DeptID As Integer = 1, Optional ByVal RoleID As Integer = -1)
        'createUser = False
        '/api/user/create
        Dim newU As newUserJSON = New newUserJSON

        With newU
            .Username = Username
            If Len(Name) Then .Name = Name Else .Name = Username
            If Len(Email) Then .Email = Email Else .Email = Username
            .ReceiveLicenseEmailNotification = sendEmail
            .UserRoleId = RoleID
            .DepartmentId = DeptID
        End With

        Dim jBody$ = JsonConvert.SerializeObject(newU)
        Dim json$ = getAPIData("/api/user/create", True, jBody)

        'Username = "121313"
    End Sub
    Public Function createKISSmodelForImport(modelName$, Optional ByVal riskId As Integer = 1, Optional ByVal versioN As Integer = 1, Optional ByVal labelS$ = "", Optional ByVal modelType$ = "Others", Optional ByVal importType$ = "Kis") As String
        '{"Id":0,
        '"Name""test_meth_inst2",
        '"RiskId":1,"Labels":"",
        '"Version":"1","IsValidFile":false,
        '"Type":"AWSCloudApplication",
        '"CreatedThrough":"Blank","UserPermissions":[],
        '"GroupPermissions":[]}

        If importType = "Kis" Then
            If isTMsix = True Then importType = "kis"
        End If


        If isTMsix Then GoTo is6

        Dim newM As kisCreateModel_setProject = New kisCreateModel_setProject

        createKISSmodelForImport = ""


        With newM
            .Id = 0
            .Name = modelName
            .RiskId = riskId
            .Version = versioN
            .Labels = labelS
            .Type = modelType
            .CreatedThrough = "Blank"
            .UserPermissions = New List(Of String)
            .GroupPermissions = New List(Of String)
        End With

        Dim jBody$ = JsonConvert.SerializeObject(newM)
        Dim jSon$ = getAPIData("/api/project/create", True, jBody)
        Return jSon
        Exit Function

is6:

        Dim newM6 As kisCreateModel60_setProject = New kisCreateModel60_setProject

        createKISSmodelForImport = ""

        With newM6
            .Id = 0
            .Name = modelName
            .RiskId = riskId
            .Version = versioN.ToString
            .Labels = labelS
            .Type = modelType
            '.CreatedThrough = "Blank"
            .UserPermissions = New List(Of String)
            .GroupPermissions = New List(Of String)
            .IsValidFile = False
            .isImportant = False
            .IsInternal = True
        End With

        Dim jBody6$ = JsonConvert.SerializeObject(newM6)
        Dim jSon6$ = getAPIData("/api/project/create", True, jBody6)



        Return jSon6
    End Function

    Public Function importKISSmodel(fileN$, ByVal projNum$, Optional ByVal integrationType$ = "Kis") As String
        importKISSmodel = ""

        If integrationType = "Kis" Then
            If isTMsix = True Then integrationType = "kis"
        End If

        Dim jSon$ = getAPIData("/api/import/" + projNum + "/" + integrationType, True,,, fileN)

        '    Dim stopHere As Integer = 1

        If Mid(jSon, 1, 5) = "ERROR" Then
            Return jSon
        End If

        Return "Action Completed"
    End Function

    Public Function addEditSR(SR As tmProjSecReq) As String
        addEditSR = ""
        Dim P As New updateEntityObject
        With P
            .LibraryId = SR.LibraryId
            .EntityType = "SecurityRequirements"
            .Model = JsonConvert.SerializeObject(SR)
        End With

        Dim jBody$ = JsonConvert.SerializeObject(P)
        Dim json$ = getAPIData("/api/threatframework/master/addedit", True, jBody)
        Return json
    End Function

    Public Function grantModelAccess(modelID As Integer, groupList As List(Of groupPermsJSON)) As Boolean
        Dim newAccess As grantAccessReq = New grantAccessReq
        newAccess.Id = modelID
        newAccess.GroupPermissions = groupList

        Dim jBody$ = JsonConvert.SerializeObject(newAccess)
        Dim json$ = getAPIData("/api/project/grantaccess", True, jBody)
        If InStr(json, "ERROR") Then grantModelAccess = False Else grantModelAccess = True

    End Function

    '{"Id"2089,"GroupPermissions":[],"UserPermissions":[{"Id":287,"Email":"somenewuser@tm.com","Name":"SomeNewUser","Permission":0,"Type":"Users","UserId":287}]}

    Public Function revokeModelAccessFromUser(modelID As Integer, useRS As List(Of permUsersJSON)) As Boolean
        revokeModelAccessFromUser = False
        Dim removeAccess As grantAccessReq = New grantAccessReq
        removeAccess.Id = modelID
        removeAccess.UserPermissions = useRS

        Dim jBody$ = JsonConvert.SerializeObject(removeAccess)

        Dim json$ = getAPIData("/api/project/revokeaccess", True, jBody)
        If InStr(json, "ERROR") Then revokeModelAccessFromUser = False Else revokeModelAccessFromUser = True

    End Function

    Public Function addEditTH(TH As tmProjThreat) As String
        addEditTH = ""
        Dim P As New updateEntityObject
        With P
            .LibraryId = TH.LibraryId
            .EntityType = "Threats"
            .Model = JsonConvert.SerializeObject(TH)
        End With

        Dim jBody$ = JsonConvert.SerializeObject(P)
        Dim json$ = getAPIData("/api/threatframework/master/addedit", True, jBody)
        Return json
    End Function
    Public Function addEditATTR(AT As tmAttribute) As String
        addEditATTR = ""
        Dim P As New updateEntityObject

        Dim A As New tmAttrCreate
        '{"LibraryId":0,"EntityType":"Attributes","Model":"{\"ImagePath\":\"/ComponentImage/DefaultComponent.jpg\",\"ComponentTypeId\":3,\"RiskId\":1,\"CodeTypeId\":1,\"DataClassificationId\":1,\"Name\":\"New ATTR\",\"Labels\":\"ThreatModeler\",\"LibrayId\":0,\"Description\":\"Some ATTR Desc\",\"IsCopy\":false}"}
        With A
            .ImagePath = "/ComponentImage/DefaultComponent.jpg"
            .ComponentTypeId = 3
            .RiskId = 1
            .CodeTypeId = 1
            .DataClassificationId = 1
            .Name = AT.Name
            .Labels = "ThreatModeler"
            .LibrayId = AT.LibraryId
            .Description = AT.Name
            .IsCopy = False
        End With

        With P
            .LibraryId = AT.LibraryId
            .EntityType = "Attributes"
            .Model = JsonConvert.SerializeObject(A)
        End With

        Dim jBody$ = JsonConvert.SerializeObject(P)
        Dim json$ = getAPIData("/api/threatframework/master/addedit", True, jBody)
        Return json
    End Function




    ' 6.0 TF edits

    Public Function addComps60(newComps As List(Of entityModelComp)) As String
        Dim newObj As entityAddPayload = New entityAddPayload
        newObj.entityTypeName = "Component"
        newObj.model = JsonConvert.SerializeObject(newComps)

        Return addEntityObject(newObj)
    End Function

    Public Function addThreats60(newThreats As List(Of entityModelThreat)) As String
        Dim newObj As entityAddPayload = New entityAddPayload
        newObj.entityTypeName = "Threat"
        newObj.model = JsonConvert.SerializeObject(newThreats)

        Return addEntityObject(newObj)
    End Function

    Public Function addSRs60(newSRs As List(Of entityModelSR)) As String
        Dim newObj As entityAddPayload = New entityAddPayload
        newObj.entityTypeName = "SecurityRequirement"
        newObj.model = JsonConvert.SerializeObject(newSRs)

        Return addEntityObject(newObj)
    End Function

    Public Function addEntityObject(newEntity As entityAddPayload) As String

        Dim jBody$ = JsonConvert.SerializeObject(newEntity)
        Dim json$ = getAPIData("/api/library/addrecords", True, jBody)
        Return json
    End Function

    Public Function addComponentStructure(compStructure As compStructurePayload) As String
        Dim jBody$ = JsonConvert.SerializeObject(compStructure)
        Dim json$ = getAPIData("/api/library/SaveComponentRelationshipDetails", True, jbody)
        Return json
    End Function









    Public Function addSR(SR As tmProjSecReq) As Boolean
        addSR = False

        Dim P As New updateEntityObject

        '"{\"ImagePath\":\"/ComponentImage/DefaultComponent.jpg\",\"ComponentTypeId\":85,\"RiskId\":1,\"CodeTypeId\":1,\"DataClassificationId\":1,\"Name\":\"NewSR\",\"Labels\":\"ThreatModeler\",\"LibrayId\":0,\"Description\":\"SR Desc\",\"IsCopy\":false}"}

        Dim S As New addSRClass

        With S
            .ImagePath = "/ComponentImage/DefaultComponent.jpg"
            .ComponentTypeId = 85
            .RiskId = 1
            .CodeTypeId = 1
            .DataClassificationId = 1
            .Name = SR.Name
            .Labels = SR.Labels
            .LibrayId = SR.LibraryId
            .Description = SR.Description
            .IsCopy = False
        End With
        With P
            .LibraryId = SR.LibraryId
            .EntityType = "SecurityRequirements"
            .Model = JsonConvert.SerializeObject(S)
        End With

        Dim jBody$ = JsonConvert.SerializeObject(P)
        Dim json$ = getAPIData("/api/threatframework/master/addedit", True, jBody)

        Dim K As Integer
        K = 1
        If Val(json) Then addSR = True

    End Function

    Public Function addTH(SR As tmProjThreat) As Boolean
        addTH = False

        Dim P As New updateEntityObject

        '"{\"ImagePath\":\"/ComponentImage/DefaultComponent.jpg\",\"ComponentTypeId\":85,\"RiskId\":1,\"CodeTypeId\":1,\"DataClassificationId\":1,\"Name\":\"NewSR\",\"Labels\":\"ThreatModeler\",\"LibrayId\":0,\"Description\":\"SR Desc\",\"IsCopy\":false}"}

        Dim S As New addTHClass

        With S
            .ImagePath = "/ComponentImage/DefaultComponent.jpg"
            .ComponentTypeId = 85
            .RiskId = 1
            .CodeTypeId = 1
            .DataClassificationId = 1
            .Name = SR.Name
            .Labels = SR.Labels
            .LibrayId = SR.LibraryId
            .Description = SR.Description
            .IsCopy = False
        End With
        With P
            .LibraryId = SR.LibraryId
            .EntityType = "Threats"
            .Model = JsonConvert.SerializeObject(S)
        End With

        Dim jBody$ = JsonConvert.SerializeObject(P)
        Dim json$ = getAPIData("/api/threatframework/master/addedit", True, jBody)

        Dim K As Integer
        K = 1
        If Val(json) Then addTH = True

    End Function
    Public Sub addEditThreat(T As tmProjThreat)
        Dim P As New updateEntityObject
        With P
            .LibraryId = T.LibraryId
            .EntityType = "SecurityRequirements"
            .Model = JsonConvert.SerializeObject(T)
        End With

        Dim jBody$ = JsonConvert.SerializeObject(P)
        Dim json$ = getAPIData("/api/threatframework/master/addedit", True, jBody)

    End Sub

    Public Function hideItem(C As Object, Optional unHide As Boolean = False) As Boolean
        hideItem = False

        Dim jBody$ = ""
        Dim jSon$ = ""
        Dim P As New addCompClass
        With P
            .LibraryId = C.LibraryId
            '            .EntityType = "Components"
            '            .Model = JsonConvert.SerializeObject(COMP)
            .ActionType = "TF_ENTITY_HIDDEN"
            If unHide Then C.ishidden = False Else C.ishidden = True
        End With

        Dim TH As editTHClass = New editTHClass 'move this if rename works without deleting SRs

        Dim editJBODY$ = ""

        Select Case entityType(C)
            Case "Components"
                Dim COMP As New tmComponent
                COMP = C
                P.EntityType = "Components"

                With P
                    .EntityType = "Components"
                    .Model = JsonConvert.SerializeObject(COMP)
                End With

            Case "SecurityRequirements"
                Dim COMP As New tmProjSecReq
                COMP = C
                P.EntityType = "SecurityRequirements"

                With P
                    .EntityType = "SecurityRequirements"
                    .Model = JsonConvert.SerializeObject(COMP)
                End With

            Case "Threats"

                With C
                    TH.Name = .Name
                    TH.Description = .Description
                    TH.RiskId = .RiskId
                    TH.RiskName = .RiskName
                    TH.LibraryId = .LibraryId
                    TH.Labels = .Labels
                    TH.Reference = .Reference
                    TH.Automated = .Automated
                    TH.StatusName = .StatusName
                    TH.IsHidden = C.ishidden
                    TH.CompanyId = .CompanyId
                    TH.IsDefault = .isDefault
                    TH.DateCreated = .DateCreated
                    TH.LastUpdated = .LastUpdated
                    TH.DepartmentId = .DepartmentId
                    TH.DepartmentName = .DepartmentName
                    TH.IsSystemDepartment = .IsSystemDepartment
                    TH.LibrayId = .LibraryId
                End With

                ' have to do this as original method will erase SRs from Threat

                With P
                    .EntityType = "Threats"
                    .Model = JsonConvert.SerializeObject(TH)
                End With



            Case "Attributes"
                Dim COMP As New tmAttribute
                COMP = C
                P.EntityType = "Threats"

                With P
                    .EntityType = "Attributes"
                    .Model = JsonConvert.SerializeObject(COMP)
                End With

        End Select

        jBody$ = JsonConvert.SerializeObject(P)
        jSon = getAPIData("/api/threatframework/master/addedit", True, jBody)
        If Val(jSon) Then hideItem = True

        P.ActionType = "TF_COMPONENT_UPDATED"

        'Console.WriteLine("PAYLOAD:" + vbCrLf + jBody + vbCrLf + "RESPONSE:" + vbCrLf + jSon)
        'now change the name to include "_TM"
        '{"LibraryId"10,"EntityType":"Components","Model":"{\"Id\":5761,\"Name\":\"Third Party Web Service_hidden\",\"Description\":\"<p>This component represents a standardized way of integrating Web-based applications using the XML, SOAP, WSDL And UDDI open standards over an Internet protocol backbone. </p><p>It Is located out of your network And managed/controlled by a 3rd party. These web applications operate in a client-server model.</p>\",\"ImagePath\":\"/ComponentImage/imageedit_5_6025455498202111181857525963.png\",\"Labels\":\"ThreatModeler,AppSec And InfraSec,Third Party Web Service\",\"Version\":\"\",\"LibraryId\":10,\"ComponentTypeId\":3,\"ComponentTypeName\":\"Component\",\"Color\":\"/ComponentImage/imageedit_5_6025455498202111181857525963.png\",\"Guid\":\"ac207d5c-d579-4dee-8a08-2e95227befff\",\"IsHidden\":false,\"DiagralElementId\":0,\"ResourceTypeValue\":null,\"Attributes\":null,\"LibrayId\":10}","ActionType":"TF_COMPONENT_UPDATED","IsCopy":false}

        If P.EntityType = "Components" Then
            If unHide = False Then C.Name += "_hidden" Else C.Name = Replace(C.Name, "_hidden", "")
            P.Model = JsonConvert.SerializeObject(C)
            jBody$ = JsonConvert.SerializeObject(P)
            jSon = getAPIData("/api/threatframework/componentmaster/addedit", True, jBody, True)

            'Console.WriteLine("PAYLOAD:" + vbCrLf + jBody + vbCrLf + "RESPONSE:" + vbCrLf + jSon)
        Else
            If unHide = False Then C.Name += "_hidden" Else C.Name = Replace(C.Name, "_hidden", "")
            Select Case P.EntityType
                Case "Threats"
                    Call addEditTH(C)
                Case "SecurityRequirements"
                    Call addEditSR(C)
            End Select
        End If

    End Function
    Public Function addEditCOMP(C As tmComponent, Optional actioNtypE$ = "TF_COMPONENT_ADDED", Optional overrideLIB As Integer = 0, Optional overrideCTYPE As Integer = 0, Optional overrideCTname$ = "") As Boolean
        addEditCOMP = False
        Dim P As New addCompClass

        Dim newC As New addCompModel
        With newC
            .ImagePath = C.ImagePath '"/ComponentImage/DefaultComponent.jpg"
            If overrideCTYPE Then .ComponentTypeId = overrideCTYPE.ToString Else .ComponentTypeId = C.ComponentTypeId.ToString
            .RiskId = 1
            .CodeTypeId = 1
            .DataClassificationId = 1
            .Name = C.Name
            If overrideCTYPE Then .ComponentTypeName = overrideCTname Else .ComponentTypeName = C.ComponentTypeName
            .Labels = C.Labels
            If overrideLIB Then .LibrayId = overrideLIB Else .LibrayId = C.LibraryId
            .Description = C.Description
        End With

        With P
            If overrideLIB Then .LibraryId = overrideLIB Else .LibraryId = C.LibraryId
            .EntityType = "Components"
            .Model = JsonConvert.SerializeObject(newC)
            .ActionType = actioNtypE
            .IsCopy = False
        End With

        Dim jBody$ = JsonConvert.SerializeObject(P)
        Dim json$ = getAPIData("/api/threatframework/componentmaster/addedit", True, jBody, True)

        If InStr(json, "true") Then addEditCOMP = True


        '       "{""LibraryId"":94,""EntityType"":""Components"",""Model"":""{\""Id\"":3105,\""Name\"":\""Log Manager\"",\""Description\"":\""<p>A scalable solution for collecting, analyzing, storing and reporting on large volumes of network and security event logs. 
        'Log Manager collects, analyzes, stores And reports On Network security log events To help organizations protect themselves against threats, attacks And Security breaches. </p><p>Log Manager helps organizations meet compliance monitoring And reporting requirements And it can be 
        'seamlessly upgraded For a higher level Of threat protection .</p>\"",\ ""ImagePath\"":\""/ComponentImage/462642202006180335030840.png\"",\""Labels\"":null,\""Version\"":\""\"",\""LibraryId\"":94,\""ComponentTypeId\"":66,
        '\""ComponentTypeName\"":\""Security Control\"",\""Color\"":\""/ComponentImage/462642202006180335030840.png\"",\""Guid\"":\""1b6b1c50-0268-4109-a23d-63297ce72194\"",\""IsHidden\"":false,\""DiagralElementId\"":0,\""ResourceTypeValue\"":null,\""listThreats\"":[],\""listDirectSRs\
        '"":[],\""listTransSRs\"":[],\""listAttr\"":[],\""isBuilt\"":false,\""duplicateSRs\"":[],\""CompID\"":3105,\""CompName\"":\""Log Manager\"",\""TypeName\"":\""Security Control\"",\""NumTH\"":0,\""NumSR\"":0}""}"


        '{"LibraryId":10,"EntityType":"Components","Model":"{\"ImagePath\":\"/ComponentImage/DefaultComponent.jpg\",\"ComponentTypeId\":\"3\",\"RiskId\":1,
        '\"CodeTypeId\":1,\"DataClassificationId\":1,\"Name\":\"TestComp\",\"ComponentTypeName\":\"Component\",\"Labels\":\"ThreatModeler\",\"LibrayId\":10,\"Description\":\"Description here\"}","ActionType":"TF_COMPONENT_ADDED","IsCopy":false}



    End Function


    Public Function editThreat6(threaT As List(Of tm6Threat)) As Boolean
        editThreat6 = False

        Dim P As editEntityTH6 = New editEntityTH6

        P.entityTypeName = "Threat"

        Dim newTlist As List(Of editModelTH6) = New List(Of editModelTH6)

        Dim newT As editModelTH6

        For Each tH In threaT
            newT = New editModelTH6
            newT.id = tH.id
            newT.description = tH.description
            newT.guid = tH.guid
            newT.labels = tH.labels
            newT.isHidden = tH.isHidden
            newT.libraryId = tH.libraryId
            newT.name = tH.name
            newT.riskId = tH.riskId
            newTlist.Add(newT)
        Next

        Dim model$ = JsonConvert.SerializeObject(newTlist)
        P.Model = model
        Dim jBody$ = JsonConvert.SerializeObject(P)
        Dim json$ = getAPIData("/api/library/updaterecords",, jBody,,,, True)

        If InStr(json, "error") > 0 Then
            Console.WriteLine(json)
            Return False
        End If

        editThreat6 = True

    End Function



    Public Function editCOMP(C As tmComponent, Optional actioNtypE$ = "TF_COMPONENT_UPDATED", Optional overrideLIB As Integer = 0, Optional overrideCTYPE As Integer = 0, Optional overrideCTname$ = "") As Boolean
        editCOMP = False
        Dim P As New addCompClass

        With C
            '            .ImagePath = C.ImagePath '"/ComponentImage/DefaultComponent.jpg"
            If overrideCTYPE Then .ComponentTypeId = overrideCTYPE.ToString Else .ComponentTypeId = C.ComponentTypeId.ToString
            '           .RiskId = 1
            '          .CodeTypeId = 1
            '         .DataClassificationId = 1
            .Name = C.Name
            If overrideCTYPE Then .ComponentTypeName = overrideCTname Else .ComponentTypeName = C.ComponentTypeName
            .Labels = C.Labels
            '        If overrideLIB Then .LibrayId = overrideLIB Else .LibrayId = C.LibraryId
            .Description = C.Description
        End With

        With P
            If overrideLIB Then .LibraryId = overrideLIB Else .LibraryId = C.LibraryId
            .EntityType = "Components"
            .Model = JsonConvert.SerializeObject(C)
            .ActionType = actioNtypE
            .IsCopy = False
        End With

        Dim jBody$ = JsonConvert.SerializeObject(P)
        Dim json$ = getAPIData("/api/threatframework/componentmaster/addedit", True, jBody, True)

        If InStr(json, "true") Then editCOMP = True

    End Function
    Public Function threatsOfEntity(C As Object, ntyType$) As List(Of tmProjThreat)
        threatsOfEntity = New List(Of tmProjThreat)
        Dim cResp As New tmCompQueryResp

        Dim jBody$ = ""
        Dim modeL$ = JsonConvert.SerializeObject(C) 'submits serialized model with escaped quotes

        Dim cReq As New tmTFQueryRequest
        With cReq
            .Model = modeL
            .LibraryId = C.LibraryId
            .EntityType = ntyType
        End With

        jBody = JsonConvert.SerializeObject(cReq)

        Dim json$ = getAPIData("/api/threatframework/query", True, jBody)

        json = Mid(json, 2, InStr(json, "],") - 1)
        json = Mid(json, InStr(json, "],") + 1)

        threatsOfEntity = JsonConvert.DeserializeObject(Of List(Of tmProjThreat))(json)


    End Function

    Public Function entityType(C As Object) As String
        entityType$ = ""
        Dim ntyType$ = ""

        Select Case Replace(C.GetType.ToString, "TMutil.", "")
            Case "tmComponent" : ntyType = "Components"
            Case "tmProjThreat" : ntyType = "Threats"
            Case "tmProjSecReq" : ntyType = "SecurityRequirements"
            Case "tmAttribute" : ntyType = "Attributes"
        End Select
        Return ntyType
    End Function

    Public Function bestMatch(C As Object, Optional ByVal showMatch As Boolean = True) As Object
        bestMatch = New tmComponent With {.Id = 0}

        Dim ndX As Integer = 0
        Dim namE$ = C.name

        Dim ntyType$ = entityType(C)

        'Console.WriteLine("Finding Best Match of " + ntyType)

        Select Case ntyType$
            Case "Threats"
                Dim destTH As New tmProjThreat
                Dim gString$ = ""
                If IsNothing(C.guid) = False Then
                    gString = C.guid.ToString
                End If
                destTH = guidTHREAT(gString)
                If destTH.Id = 0 Then
                    ndX = ndxTHbyName(C.name)
                    If ndX <> -1 Then destTH = lib_TH(ndX)
                    If showMatch Then Console.WriteLine("Best Match NAME '" + C.name + " [" + C.id.ToString + "/" + C.guid.ToString + "] at " + tmFQDN + " [" + destTH.Id.ToString + "/" + gString + "]")
                Else
                    If showMatch Then Console.WriteLine("Best Match GUID '" + C.name + " [" + C.id.ToString + "/" + C.guid.ToString + "] at " + tmFQDN + " [" + destTH.Id.ToString + "/" + destTH.Guid.ToString + "]")
                End If
                Return destTH
            Case "Components"
                Dim destC As New tmComponent
                destC = guidCOMP(C.guid.ToString)
                If destC.Id = 0 Then
                    ndX = ndxCompbyName(C.name)
                    If ndX <> -1 Then destC = lib_Comps(ndX)
                    If showMatch Then Console.WriteLine("Best Match NAME '" + C.name + " [" + C.id.ToString + "/" + C.guid.ToString + "] at " + tmFQDN + " [" + destC.Id.ToString + "/" + destC.Guid.ToString + "]")
                Else
                    If showMatch Then Console.WriteLine("Best Match GUID '" + C.name + " [" + C.id.ToString + "/" + C.guid.ToString + "] at " + tmFQDN + " [" + destC.Id.ToString + "/" + destC.Guid.ToString + "]")
                End If
                Return destC

            Case "Attributes"
                Dim destAT As New tmAttribute
                destAT = guidATTR(C.guid.ToString)
                If destAT.Id = 0 Then
                    ndX = ndxATTRbyName(C.name)
                    If ndX <> -1 Then destAT = lib_AT(ndX)
                    If showMatch Then Console.WriteLine("Best Match NAME '" + C.name + " [" + C.id.ToString + "/" + C.guid.ToString + "] at " + tmFQDN + " [" + destAT.Id.ToString + "/" + destAT.Guid.ToString + "]")
                Else
                    If showMatch Then Console.WriteLine("Best Match GUID '" + C.name + " [" + C.id.ToString + "/" + C.guid.ToString + "] at " + tmFQDN + " [" + destAT.Id.ToString + "/" + destAT.Guid.ToString + "]")
                End If
                Return destAT

            Case "SecurityRequirements"
                Dim destSR As New tmProjSecReq
                destSR = guidSR(C.guid.ToString)
                If destSR.Id = 0 Then
                    ndX = ndxSRbyName(C.name)
                    If ndX <> -1 Then destSR = lib_SR(ndX)
                    If showMatch Then Console.WriteLine("Best Match NAME '" + C.name + " [" + C.id.ToString + "/" + C.guid.ToString + "] at " + tmFQDN + " [" + destSR.Id.ToString + "/" + destSR.Guid.ToString + "]")
                Else
                    If showMatch Then Console.WriteLine("Best Match GUID '" + C.name + " [" + C.id.ToString + "/" + C.guid.ToString + "] at " + tmFQDN + " [" + destSR.Id.ToString + "/" + destSR.Guid.ToString + "]")

                End If
                Return destSR

        End Select
    End Function

    Public Function comp_ATTRmapping(idAT As Integer, Optional ByVal filen$ = "comp_template_mappings.txt") As String
        comp_ATTRmapping = ""
        Dim FF As Integer = FreeFile()
        FileOpen(FF, filen, OpenMode.Input)

        Dim attrS As New Collection 'the IDs
        Dim a$ = ""
        Do Until EOF(FF) = True

            a$ = LineInput(FF)
            attrS = CSVtoCOLL(Mid(a, InStr(a, "ATTR:") + 1))
            If grpNDX(attrS, idAT.ToString) Then
                comp_ATTRmapping += Replace(Mid(a, 1, InStr(a, " ATTR:") - 1), "COMP:", "") + ","
            End If
        Loop

        If Len(comp_ATTRmapping) Then comp_ATTRmapping = Mid(comp_ATTRmapping, 1, Len(comp_ATTRmapping) - 1)

        FileClose(FF)
    End Function

    Public Function attNumThreats(AT As tmAttribute, Optional ByVal optionName As String = "") As Integer
        attNumThreats = 0
        For Each O In AT.Options
            If Len(optionName) Then
                If LCase(optionName) <> LCase(O.Name) Then GoTo skipThis
            End If
            attNumThreats += O.Threats.Count
skipThis:
        Next

    End Function

    Public Function findDUPS(C As Object, TM As TM_Client, Optional ByVal findGUIDSonly As Boolean = False) As List(Of tmComponent)
        findDUPS = New List(Of tmComponent)

        Dim ndX As Integer = 0
        Dim namE$ = C.name
        Dim guiD$ = ""
        If IsNothing(C.Guid) = False Then
            guiD = C.guid.ToString
        End If

        Dim ntyType$ = entityType(C)

        'Console.WriteLine("Finding Best Match of " + ntyType)

        Select Case ntyType$
            Case "Threats"
                For Each tH In TM.lib_TH
                    If IsNothing(tH.Guid) = False Then
                        If LCase(tH.Guid.ToString) = LCase(guiD) Then
                            Dim newT As New tmComponent
                            With newT
                                .Id = tH.Id
                                .Name = tH.Name
                                .Guid = tH.Guid
                                .Description = tH.Description
                                findDUPS.Add(newT)
                            End With
                            GoTo nextOneT
                        End If
                    End If
                    If findGUIDSonly Then GoTo nextOneT
                    If LCase(tH.Name) = LCase(namE) Then
                        Dim newT As New tmComponent
                        With newT
                            .Id = tH.Id
                            .Name = tH.Name
                            .Guid = tH.Guid
                            findDUPS.Add(newT)
                        End With

                    End If
nextOneT:
                Next

            Case "Components"
                For Each tH In TM.lib_Comps
                    If IsNothing(tH.Guid) = False Then
                        If LCase(tH.Guid.ToString) = LCase(guiD) Then
                            Dim newT As New tmComponent
                            With newT
                                .Id = tH.Id
                                .Name = tH.Name
                                .Guid = tH.Guid
                                .Description = tH.Description
                                findDUPS.Add(newT)
                            End With
                            GoTo nextOneC
                        End If
                    End If
                    If findGUIDSonly Then GoTo nextOneC
                    If LCase(tH.Name) = LCase(namE) Then
                        Dim newT As New tmComponent
                        With newT
                            .Id = tH.Id
                            .Name = tH.Name
                            .Guid = tH.Guid
                            findDUPS.Add(newT)
                        End With

                    End If
nextOneC:
                Next

            Case "Attributes"
                For Each tH In TM.lib_AT
                    If IsNothing(tH.Guid) = False Then
                        If LCase(tH.Guid.ToString) = LCase(guiD) Then
                            Dim newT As New tmComponent
                            With newT
                                .Id = tH.Id
                                .Name = tH.Name
                                .Guid = tH.Guid
                                '.Description = tH.Description
                                findDUPS.Add(newT)
                            End With
                            GoTo nextOneA
                        End If
                    End If
                    If findGUIDSonly Then GoTo nextOneA
                    If LCase(tH.Name) = LCase(namE) Then
                        Dim newT As New tmComponent
                        With newT
                            .Id = tH.Id
                            .Name = tH.Name
                            .Guid = tH.Guid
                            findDUPS.Add(newT)
                        End With

                    End If
nextOneA:
                Next

            Case "SecurityRequirements"
                For Each tH In TM.lib_SR
                    If IsNothing(tH.Guid) = False Then
                        If LCase(tH.Guid.ToString) = LCase(guiD) Then
                            Dim newT As New tmComponent
                            With newT
                                .Id = tH.Id
                                .Name = tH.Name
                                .Guid = tH.Guid
                                .Description = tH.Description
                                findDUPS.Add(newT)
                            End With
                            GoTo nextOneS
                        End If
                    End If
                    If findGUIDSonly Then GoTo nextOneS
                    If LCase(tH.Name) = LCase(namE) Then
                        Dim newT As New tmComponent
                        With newT
                            .Id = tH.Id
                            .Name = tH.Name
                            .Guid = tH.Guid
                            findDUPS.Add(newT)
                        End With

                    End If
nextOneS:
                Next
        End Select
    End Function


    Public Function threatsByBackend(C As Object) As List(Of tmBackEndThreats)
        threatsByBackend = New List(Of tmBackendThreats)
        Dim cResp As New tmCompQueryResp

        Dim jBody$ = ""
        Dim modeL$ = JsonConvert.SerializeObject(C) 'submits serialized model with escaped quotes

        Dim cReq As New tmTFQueryRequest
        With cReq
            .Model = modeL
            .LibraryId = C.LibraryId
            .EntityType = "Widgets"
        End With

        jBody = JsonConvert.SerializeObject(cReq)

        Dim json$ = getAPIData("/api/threatframework/query", True, jBody)

        json = Mid(json, InStr(json, "[") + 1)
        json = Mid(json, 1, InStr(json, "]") - 1) + "]"

        threatsByBackend = JsonConvert.DeserializeObject(Of List(Of tmBackendThreats))(json)


    End Function

    Public Sub buildCompObj(ByVal C As tmComponent, Optional ByVal quickLookup As Boolean = False, Optional ByVal quieT As Boolean = False, Optional ByVal noSRs As Boolean = False) ' As tmComponent 'C As tmComponent, ByRef TH As List(Of tmProjThreat), ByRef SR As List(Of tmProjSecReq)) As tmComponent
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
        json = Mid(json, InStr(json, "],") + 1)
        Dim attStr$ = Mid(json, InStr(json, ",") + 1, InStr(json, "]") - 1)

        Dim ndX As Integer

        With cResp
            .listThreats = JsonConvert.DeserializeObject(Of List(Of tmProjThreat))(thStr)
            .listDirectSRs = JsonConvert.DeserializeObject(Of List(Of tmProjSecReq))(srStr)
            If Len(attStr) > 3 Then
                .listAttr = JsonConvert.DeserializeObject(Of List(Of tmAttribute))(attStr)
                ' If .listAttr.Count > 1 Then
                ' saveJSONtoFile(attStr, "attrib_" + C.Id + ".json")
                'End If
            End If

            For Each T In .listThreats
                ndX = ndxTHlib(T.Id) ', lib_TH)
                C.listThreats.Add(lib_TH(ndX))
            Next

            If quickLookup Then GoTo skipForProtocolsAndControls

            For Each T In .listAttr
                Dim ndxAT = ndxATTR(T.Id)
                If ndxAT <> -1 Then C.listAttr.Add(lib_AT(ndxAT))
            Next

            For Each T In .listDirectSRs
                ndX = ndxSRlib(T.Id) ', lib_SR)
                If ndX <> -1 Then C.listDirectSRs.Add(lib_SR(ndX))
            Next

        End With


        If C.ComponentTypeName = "Protocols" Or C.ComponentTypeName = "Security Control" Or noSRs = True Then GoTo skipForProtocolsAndControls

        Dim possibleTransSRs As New List(Of tmProjSecReq)

        possibleTransSRs = addTransitiveSRs(C)


        For Each TSR In possibleTransSRs
            '    'C.listTransSRs.Add(TSR)
            If C.numLabels > 0 Then
                If numMatchingLabels(C.Labels, TSR.Labels) / C.numLabels > 0.9 Then
                    C.listTransSRs.Add(TSR)
                End If
                '        'Console.WriteLine(TSR.Id.ToString + TSR.Name)
            End If
        Next


        'GoTo skipForProtocolsAndControls

quickLKP1:
        'Console.WriteLine(vbCrLf)

        For Each AT In C.listAttr
            For Each O In AT.Options
                For Each TH In O.Threats
                    Dim possibleAttrSRs As New List(Of tmProjSecReq)
                    If quickLookup = False Then possibleAttrSRs = addTransitiveSRsOFattr(C, O)

                    Dim ndxT As Integer = ndxTHlib(TH.Id)

                    If ndxT <> -1 Then
                        For Each TSR In lib_TH(ndxT).listLinkedSRs
                            'C.listTransSRs.Add(TSR)
                            Dim ndxS As Integer = ndxSRlib(TSR)

                            'dont do label matching here but rather keep all SRs linked to threat and match labels when wanted
                            'If C.numLabels > 0 Then
                            'If numMatchingLabels(C.Labels, lib_SR(ndxS).Labels) / C.numLabels > 0.9 Then
                            TH.listLinkedSRs.Add(TSR)
                            'End If
                        Next
                    End If

                    '                    If possibleAttrSRs.Count Then Console.WriteLine("Label matching " + possibleAttrSRs.Count.ToString + " down to " + TH.listLinkedSRs.Count.ToString + " SRs")
                Next
            Next
        Next


skipForProtocolsAndControls:

        C.isBuilt = True

    End Sub

    Public Function returnSRsWithLabelMatch(T As tmProjThreat, C As tmComponent) As Collection
        returnSRsWithLabelMatch = New Collection
        Dim ndxS As Integer = 0

        With returnSRsWithLabelMatch
            For Each TSRID In T.listLinkedSRs
                Dim TSR As New tmProjSecReq
                ndxS = ndxSRlib(TSRID)
                If ndxS = -1 Then GoTo nextSR
                TSR = lib_SR(ndxS)

                If C.numLabels > 0 Then
                    If numMatchingLabels(C.Labels, TSR.Labels) / C.numLabels > 0.9 Then
                        .Add(TSR.Id)
                    End If
                    '        'Console.WriteLine(TSR.Id.ToString + TSR.Name)
                End If
nextSR:
            Next

        End With

    End Function

    Public Function getWorkflowEvents(projID As Integer) As List(Of tmWorkflowEvent)
        Dim jSon$ = ""
        jSon$ = getAPIData("/api/workflowhistory/" + projID.ToString)

        getWorkflowEvents = JsonConvert.DeserializeObject(Of List(Of tmWorkflowEvent))(jSon)

    End Function

    Public Function getTasks(projID As Integer) As List(Of tmTask)
        Dim jSon$ = ""
        jSon$ = getAPIData("/api/project/checklist/" + projID.ToString)

        getTasks = JsonConvert.DeserializeObject(Of List(Of tmTask))(jSon)

    End Function

    Public Function getNotesOfThreat(theThreat As tmTThreat, projID As Integer) As List(Of tmNote)
        getNotesOfThreat = New List(Of tmNote)
        Dim jSon$ = ""
        Dim jBody$ = ""

        Dim gnR As New getNoteReq
        With gnR
            .Id = theThreat.Id
            .ThreatId = theThreat.ThreatId
            .ProjectId = projID
        End With

        jBody = JsonConvert.SerializeObject(gnR)
        jSon$ = getAPIData("/api/threats/notes/all", True, jBody)

        getNotesOfThreat = JsonConvert.DeserializeObject(Of List(Of tmNote))(jSon)
    End Function


    Public Function getSRsOfThreat(threatID As Integer) As List(Of tmProjSecReq)
        If IsNothing(Me.lib_TH) = True Then
            Dim R As New tfRequest

            With R
                .EntityType = "Threats"
                .LibraryId = 0
                .ShowHidden = False
            End With

            Me.lib_TH = Me.getTFThreats(R)
            Console.WriteLine("Loaded threats from TF: " + Me.lib_TH.Count.ToString)
            R.EntityType = "SecurityRequirements"
            Me.lib_SR = Me.getTFSecReqs(R)
            Console.WriteLine("Loaded SRs from TF: " + Me.lib_SR.Count.ToString)
        End If

        getSRsOfThreat = New List(Of tmProjSecReq)


        Dim jBody$
        Dim jSon$
        Dim tcStr$
        Dim tsrStr$
        Dim T As tmProjThreat
        Dim ndxTH As Integer
        Dim cReq As tmTFQueryRequest
        Dim modeL$

        'Console.WriteLine(vbCrLf)

        ndxTH = ndxTHlib(threatID)
        T = lib_TH(ndxTH)

        'removed check for ISBUILT to save time - need to refine this

        T.listLinkedSRs = New Collection
        ' End If

        cReq = New tmTFQueryRequest
            modeL$ = JsonConvert.SerializeObject(T) 'submits serialized model with escaped quotes

        'Console.WriteLine("Compiling details for ThreatID " + T.Id.ToString)
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

        getSRsOfThreat = New List(Of tmProjSecReq)
        getSRsOfThreat = JsonConvert.DeserializeObject(Of List(Of tmProjSecReq))(tsrStr)

        For Each srOf In getSRsOfThreat
            T.listLinkedSRs.Add(srOf.Id)
        Next
        T.isBuilt = True
    End Function

    Private Function addTransitiveSRs(ByRef C As tmComponent) As List(Of tmProjSecReq)
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

        'Console.WriteLine(vbCrLf)

        For K = 0 To C.listThreats.Count - 1
            T = C.listThreats(K)
            'here we get the transitive SRs
            ndxTH = ndxTHlib(T.Id) ', lib_TH)

            'removed check for ISBUILT to save time - need to refine this

            lib_TH(ndxTH).listLinkedSRs = New Collection
            ' End If

            cReq = New tmTFQueryRequest
            modeL$ = JsonConvert.SerializeObject(T) 'submits serialized model with escaped quotes

            Console.SetCursorPosition(0, Console.CursorTop - 1)
            Console.WriteLine(spaces(80))
            Console.SetCursorPosition(0, Console.CursorTop - 1)

            Console.WriteLine("Compiling details for ThreatID " + T.Id.ToString)
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


            For L = 0 To transSRs.Count - 1
                    'add all transitive SRs to lib_TH Threat
                    Dim TSR As New tmProjSecReq
                    TSR = transSRs(L)
                If grpNDX(lib_TH(ndxTH).listLinkedSRs, TSR.Id) = 0 Then lib_TH(ndxTH).listLinkedSRs.Add(TSR.Id) 'add to Threat linkedSRs collection if unique 

                Dim ndXLIB As Integer = ndxSRlib(TSR.Id) 'here ndx represents index of SR in SR Library
                    If numMatchingLabels(C.Labels, lib_SR(ndXLIB).Labels) / C.numLabels > 0.9 Then

                    ndX = ndxSR(TSR.Id, addTransitiveSRs)
                    If ndX = -1 Then
                        addTransitiveSRs.Add(lib_SR(ndXLIB))
                    Else
                        C.duplicateSRs.Add(TSR.Id)
                    End If

duplicateHere:
                End If
                Next
skipTheLoad:
            lib_TH(ndxTH).isBuilt = True

        Next

        GC.Collect()

    End Function

    Private Function addTransitiveSRsOFattr(ByVal C As tmComponent, ByRef O As tmOptions) As List(Of tmProjSecReq)
        addTransitiveSRsOFattr = New List(Of tmProjSecReq)
        If O.Threats.Count = 0 Then Exit Function

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


        For K = 0 To O.Threats.Count - 1
            T = O.Threats(K)
            'here we get the transitive SRs
            ndxTH = ndxTHlib(T.Id) ', lib_TH)

            If ndxTH = -1 Then
                Console.WriteLine("WARNING - DB ERROR: Threat has been assigned that no longer exists: " + O.Threats(K).Name + " [" + O.Threats(K).Id.ToString + "]")
                GoTo skipTheLoad
            End If

            If lib_TH(ndxTH).isBuilt = True Then
                ' added doevents as this seems to trip// suspect lowlevel probs with isBuilt prop
                'this seems to happen a lot
                'if we already know SRs, just add them from coll and skip web request
                For Each TSRID In lib_TH(ndxTH).listLinkedSRs
                    If grpNDX(C.duplicateSRs, TSRID) Then C.duplicateSRs.Add(TSRID)

                    If ndxSR(TSRID, addTransitiveSRsOFattr) <> -1 Then
                        C.duplicateSRs.Add(TSRID)
                        GoTo duphere
                    End If
                    ndX = ndxSRlib(TSRID) ', lib_SR)
                    If ndX <> -1 Then addTransitiveSRsOFattr.Add(lib_SR(ndX))
duphere:
                Next
                GoTo skipTheLoad
                        ' Else
                    End If

                    lib_TH(ndxTH).listLinkedSRs = New Collection
                    ' End If

                    cReq = New tmTFQueryRequest
                    modeL$ = JsonConvert.SerializeObject(T) 'submits serialized model with escaped quotes

            Console.SetCursorPosition(0, Console.CursorTop - 1)
            Console.WriteLine(spaces(80))
            Console.SetCursorPosition(0, Console.CursorTop - 1)

            Console.WriteLine("Compiling details for Attribute " + O.Name + " Threat " + T.Id.ToString)
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
                            .Add(TSR.Id)

                    If grpNDX(C.duplicateSRs, TSR.Id) Then C.duplicateSRs.Add(TSR.Id)
                    If ndxSR(TSR.Id, addTransitiveSRsOFattr) <> -1 Then
                        C.duplicateSRs.Add(TSR.Id)
                        GoTo duplicate
                    End If
                    'ndX = ndxSRlib(TSR.Id) ', lib_SR)
                    addTransitiveSRsOFattr.Add(TSR)
duplicate:
                        Next
                    End With
skipTheLoad:
            If ndxTH <> -1 Then lib_TH(ndxTH).isBuilt = True

        Next



    End Function

    Public Sub defineTransSRs(ByVal threatID As Long, Optional ByVal louDloG As Boolean = False)
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

        'Console.WriteLine("Getting SRs of threat")

        cReq = New tmTFQueryRequest
        modeL$ = JsonConvert.SerializeObject(T) 'submits serialized model with escaped quotes

        With cReq
            .Model = modeL
            .LibraryId = T.LibraryId
            .EntityType = "Threats"
        End With


        If louDloG Then
            Console.WriteLine("Pulling ThreatID " + T.Id.ToString + "        ")
            Console.SetCursorPosition(0, Console.CursorTop - 1)
        End If

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
            If grpNDX(T.listLinkedSRs, TSR.Id) <> 0 Then GoTo duplicate
            T.listLinkedSRs.Add(TSR.Id)
duplicate:
        Next
skipTheLoad:

        'RaiseEvent threatBuilt(T)
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

    Public Function removeLabelFromSR(ByRef SR As tmProjSecReq, label$) As Boolean
        removeLabelFromSR = False

        Dim PL As New Collection
        PL = CSVtoCOLL(SR.Labels)

        Dim K As Integer = 0
        For K = PL.Count To 1 Step -1
            If LCase(PL(K)) = LCase(label) Then
                PL.Remove(K)
                removeLabelFromSR = True
            End If
        Next

        Dim lbL$ = ""

        If removeLabelFromSR Then
            For K = 1 To PL.Count
                lbL += PL(K) + ","
            Next
            lbL = Mid(lbL, 1, Len(lbL) - 1)
            SR.Labels = lbL
        End If

    End Function

    Public Function matchLabelsOnSR(ByRef SR As tmProjSecReq, labels$) As Boolean
        matchLabelsOnSR = False

        Dim Clabels As New Collection
        Dim SRlabels As New Collection
        Clabels = CSVtoCOLL(labels)
        SRlabels = CSVtoCOLL(SR.Labels)

        For Each C In Clabels
            If grpNDX(SRlabels, C) = 0 Then
                SRlabels.Add(C)
            End If
        Next
        matchLabelsOnSR = True

        Dim lbL$ = ""

        Dim K As Integer = 0

        For K = 1 To SRlabels.Count
            lbL += SRlabels(K) + ","
        Next
        lbL = Mid(lbL, 1, Len(lbL) - 1)
        SR.Labels = lbL

    End Function

    Public Class kisCreateModel60_setProject
        '      {
        '"Id": 0,
        '"Name": "Pratibha KisImport via API",
        '"RiskId": 3,
        '"Labels": null,
        '"Version": "1",
        '"IsValidFile": false,
        '"Type": "AWS Cloud Application",
        '"IsInternal": true,
        '"UserPermissions": [
        '  
        '],
        '"GroupPermissions": [
        '  
        '],
        '"Description": "",
        '"isImportant": false
        '}


        Public Id As Integer
        Public Name As String
        Public RiskId As Integer
        Public Labels As String
        Public Version As String
        Public IsValidFile As Boolean
        Public [Type] As String
        Public IsInternal As Boolean
        Public UserPermissions As List(Of String)
        Public GroupPermissions As List(Of String)
        Public Description As String
        Public isImportant As Boolean
        Public Sub New()
            UserPermissions = New List(Of String)
            GroupPermissions = New List(Of String)
        End Sub
    End Class
    Public Class kisCreateModel_setProject
        '{"Id":0,
        '"Name""test_meth_inst2",
        '"RiskId":1,"Labels":"",
        '"Version":"1","IsValidFile":false,
        '"Type":"AWSCloudApplication",
        '"CreatedThrough":"Blank","UserPermissions":[],
        '"GroupPermissions":[]}

        Public Id As Integer
        Public Name As String
        Public RiskId As Integer
        Public Labels As String
        Public Version As Integer
        Public isValidFile As Boolean
        Public [Type] As String
        Public CreatedThrough As String
        Public UserPermissions As List(Of String)
        Public GroupPermissions As List(Of String)
    End Class


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
                    Dim newT As New tmSimpleThreatSR
                    newT.Name = T.ThreatName
                    newT.StatusName = T.StatusName
                    newT.Id = T.ThreatId
                    'newT.Description = T.Description
                    ' notes is missing!
                    'newT.RiskName = T.ActualRiskName
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

    Public Function ndxGroupByName(ByVal grpName$, ByRef G As List(Of tmGroups)) As Integer
        'finds object (for now, the node of a model) matching idOFnode
        Dim K As Long = 0

        For Each N In G
            If LCase(N.Name) = LCase(grpName) Then
                Return K
            End If
            K += 1
        Next

        Return -1
    End Function

    Public Function projOfNDX(ByVal ID As Integer, ByRef P As List(Of tmProjInfo)) As tmProjInfo
        projOfNDX = New tmProjInfo

        For Each N In P
            If N.Id = ID Then
                Return N
            End If
        Next

    End Function


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
    Public Function ndxDEPT(ByVal deptID As Long, ByRef DEPTs As List(Of tmDept)) As Integer
        'finds object (for now, the node of a model) matching idOFnode
        Dim K As Long = 0

        For Each N In DEPTs
            If N.Id = deptID Then
                Return K
            End If
            K += 1
        Next

        Return -1
    End Function

    Public Function ndxATTR(ID As Long) As Integer
        ndxATTR = -1

        Dim ndX As Integer = 0

        If isTMsix = True Then GoTo tm6

        For Each P In lib_AT
            If P.Id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

        Return ndxATTR
        Exit Function

tm6:
        For Each P In lib_AT6
            If P.id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

        Return ndxATTR
        Exit Function

    End Function

    Public Function ndxMiscByName(name$, libSearch As List(Of tmMiscTrigger)) As Integer
        ndxMiscByName = -1
        name = LCase(name)
        Dim ndX As Integer = 0
        For Each P In libSearch
            If LCase(P.Name) = name Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function

    Public Function ndxATTRbyName(namE$) As Integer
        ndxATTRbyName = -1
        namE = LCase(namE)
        Dim ndX As Integer = 0
        For Each P In lib_AT
            If LCase(P.Name) = namE Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function

    Public Function ndxATTRofList(ID As Long, L As List(Of tmAttribute)) As Integer
        ndxATTRofList = -1

        Dim ndX As Integer = 0
        For Each P In L
            If P.Id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function


    Public Function ndxTHofList(ID As Long, L As List(Of tmProjThreat)) As Integer
        ndxTHofList = -1

        Dim ndX As Integer = 0
        For Each P In L
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

        If isTMsix = True Then GoTo tm6

        For Each P In lib_TH
            If P.Id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

        Return ndxTHlib
        Exit Function

tm6:

        For Each P In lib_TH6
            If P.id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

        Return ndxTHlib
        Exit Function


    End Function

    Public Function ndxSRlib(ID As Long) As Integer
        ndxSRlib = -1

        Dim ndX As Integer = 0

        If isTMsix = True Then GoTo tm6

        For Each P In lib_SR
            If P.Id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

        Return ndxSRlib
        Exit Function

tm6:
        For Each P In lib_SR6
            If P.id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

        Return ndxSRlib
        Exit Function

    End Function

    Public Function ndxSRlibByName(namE$) As Integer
        ndxSRlibByName = -1

        Dim ndX As Integer = 0
        For Each P In lib_SR
            If P.Name = namE Then
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

    Public Function ndxLibByName(namE$, libS As List(Of tmLibrary)) As Integer
        ndxLibByName = -1

        Dim ndX As Integer = 0
        For Each P In libS
            If LCase(P.Name) = LCase(namE) Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next
    End Function

    Public Function ndxCompType(ID As Long) As Integer
        ndxCompType = -1

        Dim ndX As Integer = 0
        For Each P In componentTypes
            If P.Id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function

    Public Function ndxCompTypeByName(namE$) As Integer
        ndxCompTypeByName = -1

        Dim ndX As Integer = 0
        For Each P In componentTypes
            If LCase(P.Name) = LCase(namE) Then
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

        If isTMsix = True Then GoTo tm6

        For Each P In lib_Comps
            If P.Id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next
        Exit Function

tm6:
        For Each P In lib_Comps6
            If P.id = ID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function
    Public Function ndxCompbyGUID(GUID$) As Integer ', ByRef alL As List(Of tmComponent)) As Integer
        ndxCompbyGUID = -1

        'had to change this as passing byref from multi-threaded main causing issues

        Dim ndX As Integer = 0
        For Each P In lib_Comps
            If P.Guid.ToString = GUID Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function

    Public Function guidTHREAT(GUID$) As tmProjThreat ', ByRef alL As List(Of tmComponent)) As Integer
        'ndxCompbyGUID = -1
        guidTHREAT = New tmProjThreat
        'had to change this as passing byref from multi-threaded main causing issues

        For Each P In lib_TH
            If P.Guid.ToString = GUID Then
                Return P
                Exit Function
            End If
        Next

    End Function
    Public Function guidCOMP(GUID$) As tmComponent ', ByRef alL As List(Of tmComponent)) As Integer
        'ndxCompbyGUID = -1
        guidCOMP = New tmComponent
        'had to change this as passing byref from multi-threaded main causing issues

        For Each P In lib_Comps
            If P.Guid.ToString = GUID Then
                Return P
                Exit Function
            End If
        Next

    End Function
    Public Function guidSR(GUID$) As tmProjSecReq ', ByRef alL As List(Of tmComponent)) As Integer
        'ndxCompbyGUID = -1
        guidSR = New tmProjSecReq
        'had to change this as passing byref from multi-threaded main causing issues

        For Each P In lib_SR
            If P.Guid.ToString = GUID Then
                Return P
                Exit Function
            End If
        Next

    End Function

    Public Function guidATTR(GUID$) As tmAttribute ', ByRef alL As List(Of tmComponent)) As Integer
        'ndxCompbyGUID = -1
        guidATTR = New tmAttribute
        'had to change this as passing byref from multi-threaded main causing issues

        For Each P In lib_AT
            If P.Guid.ToString = GUID Then
                Return P
                Exit Function
            End If
        Next

    End Function

    ' VERSION 6 NDXs
    '
    '
    '

    Public Function guidTHREAT6(Optional ByVal GUID$ = "", Optional ByVal Id As Integer = -1) As tm6Threat ', ByRef alL As List(Of tmComponent)) As Integer
        'ndxCompbyGUID = -1
        guidTHREAT6 = New tm6Threat
        'had to change this as passing byref from multi-threaded main causing issues

        If Id = -1 Then
            For Each P In lib_TH6
                If P.guid.ToString = GUID Then
                    Return P
                    Exit Function
                End If
            Next
        Else
            For Each P In lib_TH6
                If P.id = Id Then
                    Return P
                    Exit Function
                End If
            Next
        End If

    End Function
    Public Function guidCOMP6(Optional ByVal GUID$ = "", Optional ByVal Id As Integer = -1) As tm6Component ', ByRef alL As List(Of tmComponent)) As Integer
        'ndxCompbyGUID = -1
        guidCOMP6 = New tm6Component
        'had to change this as passing byref from multi-threaded main causing issues
        If Id = -1 Then
            For Each P In lib_Comps6
                If P.guid.ToString = GUID Then
                    Return P
                    Exit Function
                End If
            Next
        Else
            For Each P In lib_Comps6
                If P.id = Id Then
                    Return P
                    Exit Function
                End If
            Next
        End If

    End Function


    Public Function guidSR6(Optional ByVal GUID$ = "", Optional ByVal Id As Integer = -1) As tm6SecReq ', ByRef alL As List(Of tmComponent)) As Integer
        'ndxCompbyGUID = -1
        guidSR6 = New tm6SecReq
        'had to change this as passing byref from multi-threaded main causing issues

        If Id = -1 Then
            For Each P In lib_SR6
                If P.guid.ToString = GUID Then
                    Return P
                    Exit Function
                End If
            Next
        Else
            For Each P In lib_SR6
                If P.id = Id Then
                    Return P
                    Exit Function
                End If
            Next
        End If


    End Function

    Public Function guidATTR6(Optional ByVal GUID$ = "", Optional ByVal Id As Integer = -1) As tm6Attribute ', ByRef alL As List(Of tmComponent)) As Integer
        'ndxCompbyGUID = -1
        guidATTR6 = New tm6Attribute
        'had to change this as passing byref from multi-threaded main causing issues

        If Id = -1 Then
            For Each P In lib_AT6
                If P.guid.ToString = GUID Then
                    Return P
                    Exit Function
                End If
            Next
        Else
            For Each P In lib_AT6
                If P.id = Id Then
                    Return P
                    Exit Function
                End If
            Next
        End If

    End Function



    Public Function ndxCompbyName(name$, Optional ByVal typeOnly$ = "") As Integer ', ByRef alL As List(Of tmComponent)) As Integer
        ndxCompbyName = -1

        'had to change this as passing byref from multi-threaded main causing issues

        Dim ndX As Integer = 0

        If isTMsix Then GoTo tm6

        For Each P In lib_Comps

            If LCase(P.Name) = LCase(name) Then
                If Len(typeOnly) Then
                    If LCase(P.ComponentTypeName) <> LCase(typeOnly) Then GoTo nextInList
                End If
                Return ndX
                Exit Function
            End If
nextInList:
            ndX += 1
        Next

        Exit Function

tm6:
        For Each P In lib_Comps6

            If LCase(P.name) = LCase(name) Then
                If Len(typeOnly) Then
                    If LCase(P.componentTypeName) <> LCase(typeOnly) Then GoTo nextInList2
                End If
                Return ndX
                Exit Function
            End If
nextInList2:
            ndX += 1
        Next

    End Function

    Public Function ndxTHbyName(name$) As Integer ', ByRef alL As List(Of tmComponent)) As Integer
        ndxTHbyName = -1
        Dim ndX As Integer = 0

        'had to change this as passing byref from multi-threaded main causing issues
        If isTMsix Then GoTo tm6

        For Each P In lib_TH
            If LCase(P.Name) = LCase(name) Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

        Exit Function

tm6:
        For Each P In lib_TH6
            If LCase(P.name) = LCase(name) Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

    End Function

    Public Function ndxSRbyName(name$) As Integer ', ByRef alL As List(Of tmComponent)) As Integer
        ndxSRbyName = -1
        Dim ndX As Integer = 0

        'had to change this as passing byref from multi-threaded main causing issues
        If isTMsix Then GoTo tm6

        For Each P In lib_SR
            If LCase(P.Name) = LCase(name) Then
                Return ndX
                Exit Function
            End If
            ndX += 1
        Next

        Exit Function

tm6:

        For Each P In lib_SR6
            If LCase(P.name) = LCase(name) Then
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

        nodeS = JsonConvert.DeserializeObject(Of List(Of tmNodeData))(nodeJSON) ', SS) ', SS)

        getProject = JsonConvert.DeserializeObject(Of tmModel)(jsoN)
        getProject.Nodes = nodeS

        '        Call addThreats(getProject, projID) 'add in details of threats not inside Diagram API (eg Protocols)
    End Function

End Class

Public Class tm_userGroupPermissions
    Public Groups As List(Of permInfo)
    Public Users As List(Of permInfo)
End Class

Public Class permInfo
    Public Id As Integer
    Public Email As String
    Public Name As String
    Public Permission As Integer

    Public Function permString() As String
        Return getPermString(Me.Permission)
    End Function
End Class

Public Class entityUpdate
    '{"propertyIds":[101],"propertyEntityType":"Threats","sourceEntityIds":[449],"LibraryId":1,"EntityType":"Components","IsAddition":false}
    Public propertyIds As List(Of Integer)
    Public propertyEntityType$
    Public sourceEntityIds As List(Of Integer)
    Public LibraryId As Integer
    Public EntityType$
    Public IsAddition As Boolean

    Public Sub New()
        propertyIds = New List(Of Integer)
        sourceEntityIds = New List(Of Integer)
    End Sub
End Class

Public Class entityUpdateBID
    '{"propertyIds":[101],"propertyEntityType":"Threats","sourceEntityIds":[449],"LibraryId":1,"EntityType":"Components","IsAddition":false}
    Public propertyIds As List(Of Integer)
    Public propertyEntityType$
    Public sourceEntityIds As List(Of Integer)
    Public LibraryId As Integer
    Public EntityType$
    Public IsAddition As Boolean
    Public BackendId As Integer

    Public Sub New()
        propertyIds = New List(Of Integer)
        sourceEntityIds = New List(Of Integer)
    End Sub
End Class


Public Class updateEntityObject
    Public LibraryId As Integer
    Public EntityType$
    Public Model$ ' As tmProjSecReq

End Class

Public Class addCompClass
    '{"LibraryId":10,"EntityType":"Components","Model":"{\"ImagePath\":\"/ComponentImage/DefaultComponent.jpg\",\"ComponentTypeId\":\"3\",\"RiskId\":1,\"CodeTypeId\":1,\"DataClassificationId\":1,\"Name\":\"TestComp\",\"ComponentTypeName\":\"Component\",\"Labels\":\"ThreatModeler\",\"LibrayId\":10,\"Description\":\"Description here\"}","ActionType":"TF_COMPONENT_ADDED","IsCopy":false}
    Public LibraryId As Integer
    Public EntityType$
    Public Model$
    Public ActionType$
    Public IsCopy As Boolean

End Class


Public Class editEntityTH6
    Public EntityTypeName As String
    Public Model As String ' List(Of editModelTH6)
    Public Sub New()
        EntityTypeName = "Threat"
        Model = ""
    End Sub

End Class

Public Class editModelTH6
    '[{\"id\":2438,\"name\":\"Custom Threat\",\"description\":\"\",\"riskId\":1,\"libraryId\":10,\"guid\":\"9fac0c84-7d4e-4e15-a989-426901042460\",\"labels\":\"New Tag\",\"isHidden\":false}]
    Public id As Integer
    Public name As String
    Public description As String
    Public riskId As Integer
    Public libraryId As Integer
    Public guid As String
    Public labels As String
    Public isHidden As Boolean
End Class

Public Class editEntityTH
    Public LibraryId As Integer
    Public EntityType$
    Public ActionType$
    Public Model$
End Class
Public Class addSRClass
    '"{\"ImagePath\":\"/ComponentImage/DefaultComponent.jpg\",\"ComponentTypeId\":85,\"RiskId\":1,\"CodeTypeId\":1,\"DataClassificationId\":1,\"Name\":\"NewSR\",\"Labels\":\"ThreatModeler\",\"LibrayId\":0,\"Description\":\"SR Desc\",\"IsCopy\":false}"}
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

'{"LibraryId"0,"EntityType":"Threats","Model":"{\"ImagePath\":\"/ComponentImage/DefaultComponent.jpg\",\"ComponentTypeId\":85,\"RiskId\":1,\"CodeTypeId\":1,\"DataClassificationId\":1,\"Name\":\"New Threat\",\"Labels\":\"ThreatModeler\",\"LibrayId\":0,\"Description\":\"Descript\",\"IsCopy\":false}"}
Public Class addTHClass
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

Public Class editTHClass
    '{"LibraryId":0,"EntityType":"Threats","ActionType":"TF_ENTITY_HIDDEN",
    '"Model":"{\"Id\":24,\"Guid\":\"1d68189d-90aa-4cbd-8899-8a980d95c312\",
    '\"Name\":\"Attack through Shared Data\",\"Description\":\"<p>An attacker exploits a data structure shared between multiple applications Or an application pool to affect application behavior. Data may be shared between multiple applications Or between multiple threads of a single application. </p><p>Data sharing Is usually accomplished through mutual access to a single memory location. If an attacker can manipulate this shared data (usually by co-opting one of the applications Or threads) the other applications Or threads using the shared data will often continue to trust the validity of the compromised shared data And use it in their calculations. </p><p>This can result in invalid trust assumptions, corruption of additional data through the normal operations of the other users of the shared data, Or even cause a crash Or compromise of the sharing applications.\\n</p><p>Reference <a href =\\\ "https://capec.mitre.org/data/definitions/124.html\\\" > https : //capec.mitre.org/data/definitions/124.html</a> </p>\",
    '\"RiskId\":3,\"RiskName\":\"Medium\",\"LibraryId\":1,
    '\"Labels\":\"CAPEC-124,CAPEC-116,CAPEC-117,CAPEC-153,CAPEC-167,CAPEC-168,CAPEC-194,CAPEC-200,CAPEC-204,CAPEC-215,CAPEC-257,CAPEC-261,CAPEC-277,CAPEC-37,CAPEC-39,CAPEC-255,CAPEC-536,CAPEC-545,CAPEC-567,CAPEC-569,CAPEC-610,OWASP-A6,OWASP-A5-Security-Misconfiguration\",
    '\"Reference\":null,\"Automated\":false,\"ThreatAttributes\":null,\"StatusName\":null,\"IsHidden\":true,\"CompanyId\":0,
    '\"SharingType\":null,
    '\"Readonly\":false,\"IsDefault\":false,\"DateCreated\":\"0001-01-01T0000:00\",\"LastUpdated\":\"0001-01-01T00:00:00\",
    '\"DepartmentId\":0,
    '\"DepartmentName\":null,\"Version\":null,\"IsSystemDepartment\":false,\"LibrayId\":0}"}
    Public Name$
    Public Description$
    Public RiskId As Integer
    Public RiskName$
    Public LibraryId As Integer
    Public Labels$
    Public Reference$
    Public Automated As Boolean
    Public ThreatAttributes$
    Public StatusName$
    Public IsHidden As Boolean
    Public CompanyId As Integer
    Public SharingType As Integer
    Public [Readonly] As Boolean
    Public IsDefault As Boolean
    Public DateCreated$
    Public LastUpdated$
    Public DepartmentId As Integer
    Public DepartmentName$
    Public Version$
    Public IsSystemDepartment As Boolean
    Public LibrayId As Integer
End Class


Public Class addCompModel
    Public ImagePath$
    Public ComponentTypeId$ ' As Integer
    Public RiskId As Integer
    Public CodeTypeId As Integer
    Public DataClassificationId As Integer
    Public Name$
    Public ComponentTypeName$
    Public Labels$
    Public LibrayId As Integer
    Public Description$
End Class
Public Class tfRequest
    'obj used to request Threats, SRs, Components
    'unpublished API
    '    "EntityType""SecurityRequirements","LibraryId":"0","ShowHidden":false
    Public EntityType$
    Public LibraryId As Integer '0 = ALL
    Public ShowHidden As Boolean
End Class

Public Class tfRequest6
    Public libraryId As Integer
    Public entityTypeName As String
    Public showHiddenOnly As Boolean
End Class

Public Class groupPermsJSON

    '{
    ' "Id": 2074,
    ' "GroupPermissions": [
    '   {
    '     "Id": 43,
    '     "Email": null,
    '     "Name":  "R&D",
    ''     "Permission": 1,
    '     "Type": "Groups",
    '     "GroupId": 43
    '   },
    '    {
    '      "Id": 50,
    '      "Email": null,
    '      "Name":  "Patrick Group",
    '      "Permission": 2,
    '      "Type": "Groups",
    '      "GroupId": 50
    '    }
    '  ],
    '  "UserPermissions": []
    '
    ' {"Id":2089,"GroupPermissions":[],
    '"UserPermissions":
    '[{"Id":287,"Email":"somenewuser@tm.com","Name":"SomeNewUser","Permission":0,"Type":"Users","UserId":287}]}


    Public Id As Integer
    Public Email As String
    Public Name As String
    Public Permission As Integer
    Public [Type] As String
    Public GroupId As Integer
End Class

Public Class permUsersJSON
    Public Id As Integer
    Public Email As String
    Public Name As String
    Public Permission As Integer
    Public [Type] As String
    Public UserId As Integer
End Class


' group permissions
'
' {"Id":3,"Name":"Content Management","isSystem":true,"ParentGroupId":0,"ParentGroupName":null,"Department":"Corporate","DepartmentId":1,
' "Guid":"a1bfe705-7939-4f81-96df-7abeff0aa57e","Projects":null,
' "GroupUsers":[
'{"Id":260,"GroupId":0,"GroupName":null,"UserId":287,"UserName":"SomeNewUser","isReadOnly":false,"isAdmin":true,"Guid":null,"Activated":true,
' "UserEmail":"somenewuser@tm.com"}]}
' 
' /api/group/removeusers
' 

' move user
' {"Id"287,"Name":"SomeNewUser","Email":"somenewuser@tm.com","Username":"somenewuser@tm.com","UserRoleId":1,"UserRoleName":"Admin","Activated":true,
' "LastLogin":null, "DepartmentId": 121,"UserDepartmentName":"Corporate",
' "AspNetUserId":"daf42027-55a5-4bcb-bccb-72d7e1591c85","GroupUsers":null, "TransferOwnershipToUserId": null, "TransferOwnership": false,
' "ReceiveLicenseEmailNotification":false,"LicenseRenewalEmailNotificationSent":false,"Company":null, "ApiKey": null}


' deactivate user
' {"transferOwnership"false,"UserIdToDeactivates":["daf42027-55a5-4bcb-bccb-72d7e1591c85"]}

Public Class deactivateReq
    Public transferOwnership As Boolean
    Public UserIdToDeactivates As List(Of String) 'these are aspNetUserIds

End Class
Public Class moveUserToDeptReq
    Public Id As Integer
    Public Name As String
    Public Email As String
    Public Username As String
    Public UserRoleId As Integer
    Public UserRoleName As String
    Public Activated As Boolean
    Public LastLogin As String 'leave null
    Public DepartmentId As Integer
    Public UserDepartmentName As String
    Public AspNetUserId As String
    Public GroupUsers As String 'leave null
    Public TransferOwnershipToUserId As Integer 'leave null
    Public TransferOwnership As Boolean 'make it false until test transfer
    Public ReceiveLicenseEmailNotification As Boolean
    Public LicenseRenewalEmailNotificationSent As Boolean
    Public Company As String 'null
    Public ApiKey As String 'null

End Class
Public Class removeUserFromGroupReq
    Public Id As Integer
    Public Name As String
    Public isSystem As Boolean
    Public ParentGroupId As Integer
    Public ParentGroupName As String
    Public Department As String
    Public DepartmentId As Integer
    Public Guid? As System.Guid
    Public Projects As String
    Public GroupUsers As List(Of removeUserClass)
End Class
Public Class removeUserClass
    Public Id As Integer
    Public GroupId As Integer
    Public GroupName As String
    Public UserId As Integer
    Public UserName As String
    Public isReadOnly As Boolean
    Public isAdmin As Boolean
    Public Guid? As System.Guid
    Public Activated As Boolean
    Public UserEmail As String
End Class
Public Class grantAccessReq
    Public Id As Integer
    Public GroupPermissions As List(Of groupPermsJSON)
    Public UserPermissions As List(Of permUsersJSON)

    Public Sub New()
        GroupPermissions = New List(Of groupPermsJSON)
        UserPermissions = New List(Of permUsersJSON)
    End Sub
End Class
Public Class newUserJSON
    '{"DepartmentId":1,"UserRoleId":-1,"Name":"testname","Username":"testusername","Email":"testusername@email.com","ReceiveLicenseEmailNotification":true}
    Public DepartmentId As Integer
    Public UserRoleId As Integer
    Public Name As String
    Public Username As String
    Public Email As String
    Public ReceiveLicenseEmailNotification As Boolean
End Class
Public Class threatStatusUpdate
    Public Id As Long
    Public ThreatId As Long
    Public ProjectId As Long
    Public StatusId As Integer
End Class
Public Class complianceDetails
    Public Id As Integer
    Public Name As String
    Public Sections As List(Of compSection)

    Public Sub New()
        Sections = New List(Of compSection)
    End Sub
End Class

Public Class compSection
    Public Id As Integer
    Public Name As String
    Public Title As String
    Public Domain As String
    Public SecurityRequirements As List(Of compSRs)
    Public Sub New()
        SecurityRequirements = New List(Of compSRs)
    End Sub
End Class

Public Class compSRs
    Public Id As Integer
    Public Name As String
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
    Public GroupPermissions As List(Of String)
    Public UserPermissions As List(Of String)
    Public [Type] As String


End Class
Public Class tmGroups6
    Public id As Long
    Public name$
    Public groupUsers As List(Of groupUsers6)
End Class

Public Class groupUsers6
    Public userId As Long
    Public userName As String
    Public userEmail As String
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
    'Public Projects$
    'Public GroupUsers$
    Public AllProjInfo As List(Of tmProjInfo)
End Class

Public Class tmTemplate6
    Public id As String
    Public name As String
    Public json As String
    Public diagramData As String
    Public guid As String
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
    '    "DiagramData": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADoAAAA4CAYAAACsc+sjAAAJGElEQVRoQ+2ae4xcVR3HP+fce+ex291tF1qKtCW1YiGihArCIIhpi0ALgjFFIGp5KJAIakUi/NF0pyVFCALiE1AohsRaAyrBiqHUojA0gBLRNIDaQOkSCi3L7nZ2Zuc+jvndOdvOTnfq7AzONmlvdnPncc75/b6/7+91zh3FQXKpgwQnh4A2wvRV97zgdZQG56S6D5uWaGtr9wuFeeAcY0w4S2tnhlIcjnY8QBFGRWOiHVEUblNav4HW//KcxD+H8gP54cLwtjuvOfV1pZRpRI+x5rwvjC7pWZeY1d19c/ec404zEXOiIJhmVKRNGAFWV1O+G3tXyooeuQNau2KDSHvOG2Bef/e1l9d27nzmnp6eHlmoqaspoNetXj/1qBM/Pq+YH3g8OamLUn6AMAxQEhEVAMajYdkQBsdLkEx3MNT/joD/7Fvv5R+/9+qT/PGsVTm2KaDL1255IZGeNC/0h5UxYvSmltsXgzEo7eAkU1Fp9+Czqy457vSWA73hvtyNk488+pZSYff7D7AKjRixrfMw+nq3Lr31ylN+0QjYhim4ac2Lz6UPm3JyWCw1Ind8c4zB7Wij1Ne/ctWlx68Y3+Ty6IaBXrvst5unLTjxlGg4wIRhEyv9D7UN6KSLdh3eeeql7N23nd/TUqDXr9yQM4HJTPrITLzJ7YT5YYxvATeYiPYAMOXs7KQSOG0ehe27yL/yFjqts3esOKvFQLNP5JTjZUp9u0kc0UV65lTcjhSEBsKwXFQkg8ZKx2/GJkJJjjblLB2/VuBq0Aq/b5Ditl34/QUSYkx/eOKAxuoHIZEANFEM1u1sw+1Koz0P5TlozylHSfmvXE9H7pHBBCHGDwiLPn5/nrC/SJgvoFwH5TjgyESNCUoTCNQ2AGU2IBKlSyEmDMARRRU64YIorTU6Zs4QimGiMvvhcIgKI0wUoVwXlXTQWu9pLmKbyPuJALosuyGnHTdj/XL/+cG68GguK/KgricnakwUZO9YsaDFMbpqU05rnYmisJEkOM45Bq09YTx7+/IzWwv0gouX5U7IfDUz2NeLdqRH/f9dSmm6Js/ixc33ZR95aHVrgV54ybdz7Z3TM7OPPQe/MEAY+RhhN/bCelxxf4YpZ2tJRI6TiNl85e8PE4SF7KNr72ot0CWX9+SiwM/4fp4jjjqR7qlzSSQn4Q8X8IN8GbTtScpltRZ4M6r8aO2QSEzC8VIUh/rYtWMLO99+hWSqA4XOrluzovVAMSojihWGdqJw6Oj6AJ2TZ9DWOR3PTcd10ZgQIyUkbvor8NrSKQaQxl0pJx7rlwoM9veyu7+XwYFetHJJtU1GcoGC7Lo1PRMBlMxe7Q1RFMRMRuLGBrxEGtdrx3WTuF4qBqQFkJQhE8SlxQ+KBH6RoDREEBRsCfLK9/L+dE/VPQCAVsTbSHcTqxcRhWXwYRSgjHxS3m9qdNwEOMKm41aAkk5q7H32hAC96IpVOROFltHxZNzKWK3/pER6X+242XX3L2+t637mgqtzU6fPyZSK+YZPE+o3jyHd1sXOt7Zm1z/yw9YC/dTCi3Od3dMz3YfPolgYJJL2rdldSxXyeAcjWTjVxo43X2Uo/1726Q3rWgt0wXlX5krFfCaV7mDK4TNxHY8gCuKtSq0Yq4fB2FZKg1G4rofvF9n19mv4wTDJRDq74bH7Wwt0/uLLcxgyUhLkQEwKe3vHFFLpTjwvGZ9qCmBp1CXDSonZp9dVoKW8OC7S/YhHhKEfgyvk+yns7ovLiqwlWRhjshvXr5kIoCZORqKgZNOgVIx3JHKCl0ikcOKy4uF6aaTexrtNYQtTrq2Sl00Qz/ODEmFpGN8vEIQBWulySVJ7j0iBiQW6r0uK++49z5XX9lR370bUdksjXdNIfJcNUfNqPdAFiy5/3mBOqifu4u6nRge439OHqsWVMiuf/P2DrT0cm7/oMhHYULzUZ5yxRqkrNq5/4IFG5je1zZi/6LKXgbmNCG5gTm7j+jWfbGDeqBa7oflnnP3FIz3HXQg0dKhcr1BjzNKBmclf/vXeeyfmkcSIoqcvvnRKwiS/h+JYjJkDTKsXxNjjzHZQW5UyT4b52as3beoJmluv+R3yKPnnnntdssTuOaEx0xxNyii1EMNchZlhMN2gujDKQRkJGTni7wfeQantwBZlzNNhRBETbN/0xw9theafoo0o2FSMNmvlVs4/BLSV1m6FrEOMtsLKrZQxkYxeC7wHPNQKwNVAJwN9FYJfAh4FbgMG61DoHKAX+EcdY6XPfx74RB1jmx5SDXQK8C7wNvAM8DkrYSOwoEqaPCKrfB4hP6vZBawEbq8aK1uS6hOvSqB7j/pGT6yWUVkWxzpwqrXOPqfKI0B/B1wIdANPAPOATwNPAV8BrgROtYzcDPwB+JU1jDAvjcCZwFHAd4DFgDQFPwZusdqKoq/az+cDvwG+DMiPIsaSIZ4lMsWQZwHibT8DfgCcDPzI3mUd+exPlTarxegIUBl7g3XdbwF3Wmalud4C3AokgGPsTkZA/Rp4DJA1JgHXA88CXwLOt2P/XfGIVNb8sDXGxdZg4j3VMmYD4llnANcBwrYYSoz8HHAkIDrKd7LRECPvaR3rAXoN8BNgOSDsCbvC0Ecta9LXzgI+CGyyhhlx3aOB84DTgOOBjwGXAGstUAmRI+zc/wD3W2+pJeNr1kNk3l3A3bav3grI/2bgWKujeNSfa7WA1a4rhtgAiGuJwsKMxKFYUlxwqXXpsYBKzIrwGcA3LZivVwEVN+8EhC0Z+6BlpZaMN4ElwE3WaJKxfwo8bcNIDDhyiauL58RXLUYFyMPAF6yCf7Fg5URBwP4N+C7wc6DDMtoGyP5UYlq2bTLnNZutrwK+YWOsklHRQTK6uL4kPvlO5tSS8XlAmB+yBAiDZ9s4L1gj7QDkh1ejDgXqKS8Sc3fYxSV7PgJcYAFIMlhmgb5h4+sia8QTgEUVyUcMc2MFowPWtSSOZD3J8hISwnItGd+vqARSmiT+xaACTIwusS6XEDXqQKDRhqHdAh8rxQvD+YpykrJJYX97SvEGYanyqiWjC5ANePV4mSvJT3QS+aOuRoFWr3PAvz8E9ICnaJwKHmJ0nAY74IcfNIz+Fzd/jmZVvtBTAAAAAElFTkSuQmCC",
    '    "LibraryName": "Mike Horty",
    '    "isUpdate": false
    Public Description$
    Public Labels$
    Public Image$
    Public DiagramData$
    Public LibraryName$
    Public isUpdate As Boolean
    Public guid As String
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

'Public Class ronnieJSON
'    {"Collection": {"Engagement_Name": {"Metadata": {"Engagement_Type": "TM"},
'                  "Issues": [
'                        {     "Threat_Agent": "     TA-UNAUTHZ-EXT    TA-UNAUTHZ-INT   ",
'                              "Threat_Object": "         TO-ADFS-SERVICE      TO-Credentials-Admin    ",
'                              "Attack_Scenario": "   Description ",
'                              "Threat_type": "Spoofing,Non-Repudiatio,Elevation of Privilege,Lateral movement    ",
'                              "Attack_Surface": "   xxx  API   ",
'                              "Attack_Vector": "       Mitre Tactic:    Initial Access  Discovery  Collection    Technique: Phishing  Credential Harvesting  Account Discovery  Input Capture: Key Logging  Input Capture: Web Portal Capture  Man in the browser   ",
'                              "Attack_Goal": "   Description  ",
'                              "Attack_Impact": "High",
'                              "Threat_Rating": "Very High",
'                              "Control_Section": "       Control (Protect/Detect):        C-Security-Awareness   All employees should undergo continuous learning around security best practices.   C-AuthN-MFA   Console based access to critical systems should enforce use of Multi-factor Authentication.   C-AuthN-JIT   Just in time access to critical systems should be based on a check In/Check out process leveraging a credential vault.   C-Bastion-Host-Access ",
'                              "Recommendation": "   <description>     ",
'                              "ID": 1
'                        }]}}
' EXPANDED: "Control_Section": "       Control (Protect/Detect):        
'                                      C-Security-Awareness   All employees should undergo continuous learning around security best practices.   
'                                      C-AuthN-MFA   Console based access to critical systems should enforce use of Multi-factor Authentication.   
'                                      C-AuthN-JIT   Just in time access to critical systems should be based on a check In/Check out process leveraging a credential vault.   
'                                      C-Bastion-Host-Access ",
'                              
Public Class ronnieIssue
    Public Threat_Agent As String
    Public Threat_Object As String
    Public Attack_Scenario As String
    Public Threat_type As String
    Public Attack_Surface As String
    Public Attack_Vector As String
    Public Attack_Goal As String
    Public Attack_Impact As String
    Public Threat_Rating As String
    Public Control_Section As String
    Public Recommendation As String
    Public ID As Integer

    Public Function getAttackVectorSingle() As String
        Return Replace(LTrim(Attack_Vector), "  ", "<br>")
    End Function
    Public Function getAttackVectorList() As Collection
        Dim aV As Collection = New Collection
        Dim avStr$ = Replace(Attack_Vector, "  ", ",")
        Dim cList As Collection = CSVtoCOLL(avStr)

        For Each V In cList
            Dim a$ = LTrim(RTrim(V))
            If a = "Mitre Tactic:" Or a = "Initial Access" Or a = "Discovery" Or a = "Collection" Then GoTo skipThis
            aV.Add(a)
skipThis:
        Next

        Return aV
    End Function
    Public Function getThreatAgents() As Collection
        Dim aV As Collection = New Collection
        Dim avStr$ = Replace(Threat_Agent, "  ", ",")
        Dim cList As Collection = CSVtoCOLL(avStr)

        For Each V In cList
            Dim a$ = LTrim(RTrim(V))
            If Len(a) Then aV.Add(a)
        Next

        Return aV
    End Function

    Public Function getAgentModels() As List(Of entityModelComp)
        getAgentModels = New List(Of entityModelComp)

        For Each agenT In Me.getThreatAgents
            Dim newAgent As entityModelComp = New entityModelComp
            newAgent.name = agenT
            newAgent.labels = COLLtoCSV(Me.getThreatObjects)
            newAgent.imagepath = "/ComponentImage/devil.png"
            newAgent.description = "Threat Agent " + agenT + " related to objects: " + newAgent.labels
            getAgentModels.Add(newAgent)
        Next agenT

    End Function

    Public Function getThreatModels() As List(Of entityModelThreat)
        getThreatModels = New List(Of entityModelThreat)

        Dim lF$ = "<br>"
        Dim newThreat As entityModelThreat = New entityModelThreat

        newThreat.name = "Issue ID " + Me.ID.ToString + " " + Replace(COLLtoCSV(Me.getThreatObjects), ",", " ")
        newThreat.labels = LTrim(RTrim(Me.Threat_type))
        newThreat.setRiskIdByString(Me.Threat_Rating)
        newThreat.description = "Scenario: " + Me.Attack_Scenario + lF + "Impact    : " + Me.Attack_Impact + lF + "Rating    : " + Me.Threat_Rating + lF + "Goal     : " + Me.Attack_Goal + lF + lF + Me.getAttackVectorSingle
        getThreatModels.Add(newThreat)

    End Function

    Public Function getSReqModels(Optional splitControls As Boolean = False) As List(Of entityModelSR)

        getSReqModels = New List(Of entityModelSR)

        Select Case splitControls
            Case False
                Dim newSR As entityModelSR = New entityModelSR
                newSR.name = "Issue " + Me.ID.ToString + " Controls"
                newSR.description = Me.getControlsSingle
                getSReqModels.Add(newSR)

            Case True
                For Each sR In Me.getControls
                    Dim newSR As entityModelSR = New entityModelSR

                    newSR.name = sR.controlName
                    newSR.description = sR.controlDescription
                    getSReqModels.Add(newSR)
                Next sR
        End Select

    End Function


    Public Function getThreatObjects() As Collection
        Dim aV As Collection = New Collection
        Dim avStr$ = Replace(Threat_Object, "  ", ",")
        Dim cList As Collection = CSVtoCOLL(avStr)

        For Each V In cList
            Dim a$ = LTrim(RTrim(V))
            If Len(a) Then aV.Add(a)
        Next

        Return aV
    End Function

    Public Function getControls() As List(Of controlObj)
        ' "Control_Section": "       Control (Protect/Detect):        C-Security-Awareness   All employees should undergo continuous learning around security best practices.   C-AuthN-MFA   Console based access to critical systems should enforce use of Multi-factor Authentication.   C-AuthN-JIT   Just in time access to critical systems should be based on a check In/Check out process leveraging a credential vault.   C-Bastion-Host-Access ",
        Dim cS As List(Of controlObj) = New List(Of controlObj)
        Dim csStr$ = Replace(Control_Section, "  ", ",")
        Dim csList As Collection = CSVtoCOLL(csStr)

        Dim cLoop As Integer = 0
        For cLoop = 1 To csList.Count
            Dim a$ = LTrim(RTrim(csList(cLoop)))
            If Mid(a, 1, 2) = "C-" Then
                Dim newCS As New controlObj
                newCS.controlName = a
                If cLoop + 1 <= csList.Count Then newCS.controlDescription = LTrim(RTrim(csList(cLoop + 1))) Else newCS.controlDescription = ""
                cS.Add(newCS)
            End If
        Next

        Return cS
    End Function

    Public Function getControlsSingle() As String
        Return Replace(LTrim(Control_Section), "  ", "<br>") + "<br>" + Recommendation
    End Function

End Class

Public Class controlObj
    Public controlName$
    Public controlDescription$
End Class








' ------------- 6.0 Threat Framework modification classes -----------------
Public Class compStructurePayload
    Public [Type] As String
    Public Id As Integer
    Public threats As List(Of threatStructurePayload)
    Public securityRequirements As List(Of srStructurePayload)
    Public properties As List(Of attrStructurePayload)

    Public Sub New()
        Me.Type = "Component"
        threats = New List(Of threatStructurePayload)
        securityRequirements = New List(Of srStructurePayload)
        properties = New List(Of attrStructurePayload)
    End Sub
End Class
Public Class threatStructurePayload
    Public Id As Integer
    Public [Type] As String
    Public isDefault As Boolean
    Public securityRequirements As List(Of srStructurePayload)

    Public Sub New()
        Me.Type = "Threat"
        Me.isDefault = False
        securityRequirements = New List(Of srStructurePayload)
    End Sub
End Class

Public Class srStructurePayload
    Public Id As Integer
    Public [Type] As String

    Public Sub New()
        Me.Type = "SecurityRequirement"
    End Sub
End Class
Public Class entityAddPayload
    Public entityTypeName As String
    Public model As String
End Class

Public Class attrStructurePayload
    Public Id As Integer
    Public [Type] As String
    Public isOptional As Boolean
    Public options As List(Of optionStructurePayload)
    Public Sub New()
        Me.Type = "Property"
        options = New List(Of optionStructurePayload)
    End Sub
End Class

Public Class optionStructurePayload
    Public Id As Integer
    Public [Type] As String
    Public isDefault As Boolean
    Public threats As List(Of threatStructurePayload)

    Public Sub New()
        Me.Type = "Option"
        Me.isDefault = False
        threats = New List(Of threatStructurePayload)
    End Sub
End Class

Public Class entityModelComp
    ' COMP MODEL
    '
    '"[{\"id\":0,\"name\":\"Another new component\",\"description\":\"<p>Here is my description2222</p>\",\"labels\":\"tag1\",\"libraryId\":0,\"guid\":null,\"imagepath\":\"/ComponentImage/icon-developer-icon.png\",
    ' \"componentTypeId\":3,\"componentTypeName\":\"Component\"}]"}
    '
    ' THREAT MODEL
    ' {"EntityTypeName":"Threat","Model":"[{\"id\":0,\"name\":\"New threat\",\"description\":\"\",\"riskId\":1,\"libraryId\":0,\"guid\":null,\"labels\":\"\",\"isHidden\":false}]"}
    '
    ' SR
    ' {"EntityTypeName":"SecurityRequirement","Model":"[{\"id\":0,\"name\":\"New SR\",\"libraryId\":0,\"guid\":null,\"description\":\"\",\"labels\":\"\",\"isHidden\":false}]"}

    Public id As Integer
    Public name As String
    Public description As String
    Public labels As String
    Public libraryId As Integer
    Public guid As String
    Public imagepath As String
    Public componentTypeId As Integer
    Public componentTypeName As String

    Public Sub New()
        id = 0
        libraryId = 0
        imagepath = "/ComponentImage/icon-developer-icon.png"
        componentTypeName = "Component"
        componentTypeId = 3
    End Sub
End Class

Public Class entityModelThreat
    Public id As Integer
    Public name As String
    Public description As String
    Public riskId As Integer ' 1=very high, 5=very low
    Public libraryId As Integer
    Public guid As String
    Public labels As String
    Public isHidden As Boolean

    Public Sub setRiskIdByString(riskDescription)
        Select Case LCase(riskDescription)
            Case "very high"
                riskId = 1
            Case "high"
                riskId = 2
            Case "medium"
                riskId = 3
            Case "low"
                riskId = 4
            Case "very low"
                riskId = 5
        End Select
    End Sub

    Public Sub New()
        libraryId = 0
        isHidden = False
        riskId = 1
    End Sub
End Class

Public Class entityModelSR
        Public id As Integer
        Public name As String
    Public description As String
    Public libraryId As Integer
    Public guid As String
        Public labels As String
        Public isHidden As Boolean

    Public Sub New()
        libraryId = 0
        isHidden = False
    End Sub
End Class










    Public Class matildaApp
    ' 	"app_id": "7",
    '"app_name": "MJT",
    '"app_owner": "Rajesh",
    '"availability": "Medium",
    '"business_group": "Sales",
    '"cloud_provider": "AWS",
    '"cloud_suitability": "Yes",
    '"compliance": "SOCX",
    '"createdDate": "2022-08-15T17:32:59.049+00:00",

    Public [_id] As String
    Public app_code As String
    Public app_owner As String
    Public availability As String
    Public business_group As String
    Public cloud_provider As String
    Public cloud_suitability As String
    Public compliance As String
    Public createdDate As String
    'Public dependency As List(Of matildaDependency)   'dependency is a list inside of a list entitled 'REACTIVEWEB' - ignoring for now
    Public discovered_name As String
    Public elasticity As String
    Public environment As String
    Public hosts As List(Of matildaHost)
    Public services As List(Of matildaServiceOuter)

    Public Function getServices(host$) As List(Of matildaService)
        getServices = New List(Of matildaService)
        For Each sOut In Me.services
            For Each S In sOut.services
                If arrNDX(S.host, host) <> 0 Then
                    getServices.Add(S)
                End If
            Next
        Next
    End Function
End Class

Public Class matildaHost
    Public host_id As Integer
    Public ip As String
    Public name As String
    Public instance_type As String
    Public recommended_instance_type As String
    Public total_memory_gb As Long
    Public logical_processors As Integer
    Public serverType As String
    Public total_storage_gb As Long
    Public operating_system As String
    Public security As List(Of matildaPorts)

    '    "ip": "52.24.3.106",
    '			"name": "mongodb2.discovery.com",
    '			"instance_type": "c6g.large",
    '			"recommended_instance_type": "c6g.xlarge",
    '			"total_memory_gb": 3.83,
    '			"logical_processors": 2,
    '			"total_storage_gb": 49',
    '			"operating_sys'tem": "Ubuntu 20.04.3 LTS (Focal Fossa)",'
    '			"security": [{'
    '					"tcp": 27017
    '				},
    '  "os_recommendation" "Retain OS as is Ubuntu 20.04",
    '			"migration_strategy": "Replatform",
    '			"eol_status": "Expiring Soon",
    '			"eol_date": "04-30-2025",
    '			"discoveryMap": {
    '				"ip": "52.24.3.106",
    '				"host_name": "mongodb2.discovery.com"
    '			}

    Public os_recommendation As String
    Public migration_strategy As String
    Public eol_status As String
    Public eol_date As String
    Public discoveryMap As matildaDiscovery


End Class
Public Class matildaPorts
    Public tcp As Long
End Class
Public Class matildaDiscovery
    Public ip As String
    Public host_name As String
End Class
Public Class matildaServiceOuter
    Public name As String
    Public services As List(Of matildaService)
End Class
Public Class matildaService
    '    "name": "tomcat",
    '				"version": "9.0.16",
    '				"host": [
    '					"129.146.23.28"
    '				],
    '				"ports": [
    '					8080
    '				],
    '				"path": "/usr/share/tomcat9/etc/server.xml",
    '				"folders": [
    '					"/usr/share/tomcat9/webapps"
    '				],
    '				"config": "/var/lib/tomcat9",
    '				"discoveryMap": {
    '					"name": "tomcat",
    '					"version": "9.0.16"
    '				},
    '				"machineUser": "mdiscover",
    '				"serviceUser": null,
    '				"hostName": [
    '				'	"TomcatWebServer"
    '				]',
    '				"id": "62fa834ac485e7009a754921",'
    '				"recommended_service": "tomcat 9.0.16"
    '			}]
    Public name As String
    Public version As String
    Public host() As String
    Public ports() As Long
    Public recommended_service As String
End Class
Public Class tmLabels
    Public Id As Long
    Public Name$
    Public IsSystem As Boolean
    Public Attributes$
End Class

Public Class tmProjActivity6
    '  "projectId": 2090,
    '   "projectName": "WebAppAWS",
    '   "userId": 273,
    '   "userName": "Mike",
    '   "transactionName": "ProjectThreatStatusUpdate",
    '   "description": "Accessing Functionality Not Properly Constrained by ACLs Threat on API 2 Status Change : Open to Partially Mitigated",
    '   "notes": "ProjectThreat",
    '   "actionTime": "2022-12-01T00:03:38.473",
    '   "diagramObjectType": "Node",
    '   "source": "API 2",
    '   "sourceId": 23086,
    '   "innerModels": [
    ''  name
    '

    Public projectId As Integer
    Public projectName$
    Public userId As Integer
    Public userName$
    Public transactionName$
    Public description$
    Public notes$
    Public actionTime As DateTime
    Public diagramObjectType$
    Public source$
    Public sourceId$
    'Public innerModels As List(Of tmInnerModel6)

    Public Sub New()
        '    innerModels = New List(Of tmInnerModel6)
    End Sub
End Class



Public Class tm6ThreatsOfModel
    '{
    '    "id": 170004,
    '    "threatName": "Accessing Functionality Not Properly Constrained by ACLs",
    '    "sourceName": "Web Service",
    '    "sourceDisplayName": "Web Service",
    '    "actualRiskId": 2,
    '    "threatId": 1435,
    '    "projectId": 2220,
    '    "projectGuid": "1527b361-0f8f-4bad-a8c5-817eae2a04ad",
    '    "actualRisk": {
    '        "id": 2,
    '        "name": "High",
    '        "color": "#e32c2c",
    '        "score": 4
    '    },
    '    "actualRiskName": "High",
    '    "dateCreated": "0001-01-01T00:00:00",
    '    "statusName": "Mitigated by Control",
    '    "type": 0,
    '    "sourceId": 25431,
    '    "sourceType": "Node",
    '    "componentId": 248,
    '    "componentName": "Web Service",
    '    "componentTypeId": 3,
    '    "componentTypeName": "Component",
    '    "imagePath": "/ComponentImage/webservice.png",
    '    "threatGuid": "ce4b961f-c949-46c2-8e5d-3cfbc8b125bf",
    '    "elementId": 25431,
    '    "statusId": 3,
    '    "isMitigatedBySecurityControl": false,
    '    "canBeMitigatedBySecurityControlId": 0
    '}

    Public id As Long
    Public threatName$
    Public sourceName$ 'the diagram element responsible for this threat
    Public sourceDisplayName$
    Public actualRiskId As Integer
    Public threatId As Integer
    Public projectId As Integer
    Public projectGuid$
    ' actualrisk object
    Public actualRiskName$
    Public statusName$
    Public externalSourceId As String
    ' type is always 0
    Public sourceType$ 'usually/always Node
    Public componentId As Integer
    Public componentName$
    Public imagePath$
    Public threatGuid$
    Public elementId As Long ' copy of sourceId
    Public statusId As Integer ' coincides with statusName 
    Public isMitigatedBySecurityControl As Boolean
    Public canBeMitigatedBySecurityControlId As Integer


End Class

Public Class tm6SRsOfModel
    '      {
    '                "id": 610032,
    '                "projectId": 2220,
    '                "projectGuid": "1527b361-0f8f-4bad-a8c5-817eae2a04ad",
    '                "projectName": "API Testing",
    '                "securityRequirementId": 2257,
    '                "securityRequirementGuid": "4FF9A3CF-F2F7-4E51-B200-5698D08217F6",
    '                "securityRequirementName": "Applications that handle highly sensitive information should consider disabling copy and paste functionality",
    '                "projectSecurityRequirementId": 271931,
    '                "elementId": 25431,
    '                "elementName": "Web Service",
    '                "elementDisplayName": "Web Service",
    '                "isImplemented": false,
    '                "isOptional": false,
    '                "sourceId": 52,
    '                "sourceType": "Threats",
    '                "sourceName": "Collect Data from Common Resource Locations",
    '                "statusId": 4,
    '                "statusName": "Open",
    '                "projectSecurityRequirementStatusId": 4,
    '                "projectSecurityRequirementStatusName": "Open",
    '                "componentId": 248,
    '                "componentName": "Web Service",
    ''                "componentTypeId": 3,
    '                "componentTypeName": "Component",
    '                "imagePath": "/ComponentImage/webservice.png",
    '                "diagramNodeThreatId": 164436,
    '                "type": 0,
    '                "canBeMitigatedBySecurityControlId": 0,
    '                "diagramNodeSecurityRequirementStatusId": 4,
    '                "diagramNodeSecurityRequirementStatusName": "Open"
    '            }
    Public id As Long
    Public projectId As Integer
    Public projectGuid$
    Public projectName$
    Public securityRequirementId As Integer
    Public securityRequirementGuid$
    Public securityRequirementName$
    Public projectSecurityRequirementId As Long
    Public elementId As Long
    Public elementName$
    Public elementDisplayName$
    Public isImplemented As Boolean
    Public isOptional As Boolean
    Public externalSourceId As String
    Public sourceId As Integer ' the library ID of the source of this
    Public sourceType$   ' could be Properties, Threats, Node
    Public sourceName$
    Public statusId As Integer
    Public statusName$
    Public projectSecurityRequirementStatusId As Integer
    Public projectSecurityRequirementStatusName$
    Public componentId As Integer
    Public componentName$
    Public componentTypeId As Integer
    Public componentTypeName$
    Public imagePath$
    Public diagramNodeThreatId As Long

    ' skipping last 4 - duplicates or useless?


End Class

Public Class tm6NodesOfModel
    ' category types include component,container,trustboundary,project(nested),collection
    ' Field highlights how object is represented.. Protocols displayed as "linkDataArray" collection
    ' most of this unusable
    '    "nodeDataArray" [
    '                {
    '                    "Id": 25383,
    '                    "key": -18,
    '                    "category": "Component",
    '                    "Name": "Home",
    '                    "type": "Group",
    '                    "ImageURI": "/ComponentImage/home.png",
    '                    "NodeId": 34,
    '                    "ComponentId": 34,
    '                    "group": null,
    '                    "Location": "-266 -106",
    '                    "color": null,
    '                    "BorderColor": "black",
    '                    "BackgroundColor": "white",
    '                    "TitleColor": "black",
    '                    "TitleBackgroundColor": null,
    '                    "isGroup": true,
    '                    "isGraphExpanded": false,
    '                    "FullName": "Home",
    '                    "LocationX": null,
    '                    "LocationY": null,
    '                    "cpeid": null,
    '                    "networkComponentId": null,
    '                    "BorderThickness": 1.0,
    '                    "AllowMove": true,
    '                    "LayoutWidth": null,
    '                    "LayoutHeight": null,
    '                    "guid": "",
    '                    "AWS_Id": null,
    '                    "ExternalResourceId": null,
    '                    "ExternalResourceARN": null,
    '                    "ResourceType": null,
    '                    "IsSystem": false,
    '                    "SecurityGroups": null,
    '                    "Components": null,
    '                    "DataElements": null,
    '                    "Roles": null,
    '                    "Widgets": null,
    '                    "Properties": null,
    '                    "ThreatsCount": 0,
    '                    "SecurityRequirementsCount": 0,
    '                    "LibraryId": 0,
    '                    "ComponentTypeId": 0,
    '                    "ComponentTypeName": "Component",
    '                    "Labels": null,
    '                    "ComponentName": "Home",
    '                    "ComponentGuid": "4c7f9883-638e-4e1d-b2f3-63cd74620e2a",
    '                    "BucketPolicy": null,
    '                    "EncryptionEnabled": false,
    '                    "PubliclyAccessible": false,
    '                    "LogsEnabled": false,
    '                    "VersionEnabled": false,
    '                    "MfaDeleteEnabled": false,
    '                    "ObjectLockEnabled": false,
    '                    "RoleName": null,
    '                    "Policies": null,
    '                    "InstanceStopped": false,
    '                    "Tags": null,
    '                    "Notes": null,
    '                    "isImportThreats": false,
    '                    "isHighValueTarget": false,
    '                    "isProtectedResource": false,
    '                    "IsNode": true,
    '                    "ProjectId": null
    '                }

    ' usable fields
    '    {
    '                    "Id": 25430,
    '               '     "key": -878,
    '                    "category": "Component",
    '                    "Name": "My Web Service",
    '                    "ImageURI": "/ComponentImage/webservice.png",
    '                    "NodeId": 3759,
    '                    "ComponentId": 3759,                 
    '                    "FullName": "Custom Name",
    '                    "ComponentTypeName": "Component",
    '                    "ComponentName": "My Web Service",
    '                    "ComponentGuid": "61f9d2ac-c7a7-4118-b27d-d7f71323f2ac",
    '                 }

    Public Id As Long ' ID is unique across all threat models  
    Public category$ ' component,container,trustboundary,project(nested),collection
    Public Name$
    Public ImageURI$
    Public NodeId? As Integer
    Public FullName$
    Public ComponentTypeName$ ' component, protocol, securitycontrol, trustboundary
    Public ComponentGuid$
End Class

Public Class tmProjInfoShort6
    '       "name": "Test-Yash-1",
    '            "id" 2069,
    '            "isImportant": false,
    '            "description": "",
    '            "version": "1",
    '            "archived": false,
    '            "createDate": "2022-11-17T18:42:39.027",
    '            "createdById": 1,
    '            "riskId": 3,
    '            "riskName": "Medium",
    ''            "riskColor": "#fbbc05",
    '            "lastModifiedById": 1,
    '            "lastModifiedDate": "2022-11-17T18:42:53.557",
    '            "lastModifiedByName": "Corporate Admin",
    '            "createdByName": "Corporate Admin",
    '            "departmentName": "Corporate",
    '            "departmentId": 1,
    '            "guid": "19599980-52fe-4eee-9adf-33f03723febf",
    '            "type": "AWSCloudApplication",
    '            "permission": 2
    Public id As Long
    Public name$
    Public version$
    Public guid$ ' As System.Guid
    Public description$
    Public isImportant As Boolean
    Public archived As Boolean
    Public labels As String
    Public createDate As DateTime
    Public createdById As Integer
    Public riskId As Integer
    Public riskName$
    Public lastModifiedByName$
    Public lastModifiedById As Integer 'Name$
    Public LastModifiedDate As DateTime
    'Public RiskName As Integer
    Public [type] As String
    Public createdByName$
    Public departmentName$
    Public departmentId As Integer
    Public permission As Integer

End Class
Public Class tmProjInfo
    Public Id As Long
    Public Name$
    Public Version$
    'Public Guid As System.Guid
    Public isInternal As Boolean
    Public Labels As String
    Public CreateDate As DateTime
    Public RiskId As Integer
    Public RiskName$
    Public CreatedByName$
    Public LastModifiedByName$
    Public LastModifiedDate As DateTime
    'Public RiskName As Integer
    Public [Type] As String
    Public CreatedByUserEmail$
    Public Model As tmModel
    'Public nodesFromThreats As Integer 'beginning at 300k, nodes from threats like HTTPS
End Class

Public Class currUser
    'this is a 6.0 class
    Public email As String
    Public id As String
    Public userName As String
    Public userInfo As tmUserSix
End Class
Public Class tmUserSix
    Public Id As Integer
    Public Name As String
    Public Email As String
    Public Username As String
End Class
Public Class tmDept
    Public Id As Integer
    Public Name$
    Public Guid As System.Guid
    Public DateCreated As DateTime
    Public IsSystem As Boolean
    Public TotalLicenses As Integer
    Public UsedLicenses As Integer
    Public LibraryId As Integer
    Public SharingType$
    Public [Readonly] As Boolean
End Class
Public Class tmUser
    ' {
    '"Id" 271,
    '"Name": "Ahsan",
    '"Email": "ahsan.rehman@threatmodeler.com",
    '"Username": "ahsan.rehman@threatmodeler.com",
    '"UserRoleId": 1,
    '"UserRoleName": "Admin",
    '"Activated": false,
    '"LastLogin": "2022-02-02T20:25:25.263Z",
    '"DepartmentId": 1,
    '"UserDepartmentName": "Corporate",
    '"AspNetUserId": "81285fde-a096-4220-8b3d-7ba9b1e4838f",
    '"GroupUsers": null,
    '"TransferOwnershipToUserId": null,
    '"TransferOwnership": false,
    '"ReceiveLicenseEmailNotification": true,
    '"LicenseRenewalEmailNotificationSent": false,
    '"Company": null,
    '"ApiKey": null

    Public Id As Integer
    Public Name As String
    Public Email As String
    Public Username As String
    Public UserRoleId As Integer
    Public UserRoleName As String
    Public AspNetUserId As String
    Public Activated As Boolean
    Public LastLogin? As DateTime
    Public DepartmentId As Integer
    Public ApiKey As String

End Class

Public Class tmUserOfGroup
    Public Id As Integer
    Public UserId As Integer
    Public UserEmail$
    Public isAdmin As Boolean
    Public isReadOnly As Boolean
End Class
Public Class tmProjThreat
    ' these threats come from the DIAGRAM API as part of nodeArray
    Public Id As Long
    Public Guid? As System.Guid 'nulls
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

    '            "id" 99,
    '           "name": "Automobile",
    '            "companyId": 0,
    '            "guid": "0359a988-f6f7-421a-8a87-e68ab918f79f",
    '            "sharingType": "Public",
    '            "readonly": false,
    '            "isDefault": false,
    '            "dateCreated": "2022-11-02T17:45:49.123",
    '            "lastUpdated": "2022-11-02T17:54:00.903",
    '            "departmentId": 118,
    '            "departmentName": "Automobile",
    '            "labels": "",
    '            "version": "6.0",
    '            "isSystemDepartment": true
    '        },


End Class

Public Class tmBackendThreats
    '{
    '             "Id": 341,
    '             "BackendId": 1,
    '             "ThreatId": 0,
    '             "ThreatName": "Cross Site Request Forgery",
    '             "BackendName": "None",
    '             "WidgetId": 0,
    '             "WidgetName": null,
    '             "LibraryId": 1
    '         },

    Public Id As Long
    Public BackendId As Integer
    Public ThreatId As Integer
    Public ThreatName$
    Public BackendName$
    Public WidgetId As Integer


End Class

Public Class tmMiscTrigger
    '"Id": 1,
    '"Name": "Executive",
    '"Labels": "",
    '"LibraryId": 1,
    '"Guid": "c27efce9-788f-4534-9e95-285bd37f8c93",
    '"Description": null,
    '"IsHidden": false
    Public Id As Long
    Public Name$
    Public Labels$
    Public Description$
    Public LibraryId As Integer
    Public Guid? As System.Guid
    Public isHidden As Boolean

End Class
Public Class tmCompQueryResp
    Public listThreats As List(Of tmProjThreat)
    Public listDirectSRs As List(Of tmProjSecReq)
    'Public listTransSRs As List(Of tmProjSecReq)
    Public listAttr As List(Of tmAttribute)
End Class

Public Class tmAttrCreate
    '{"LibraryId":0,"EntityType":"Attributes","Model":"{\"ImagePath\":\"/ComponentImage/DefaultComponent.jpg\",\"ComponentTypeId\":3,\"RiskId\":1,\"CodeTypeId\":1,\"DataClassificationId\":1,\"Name\":\"New ATTR\",\"Labels\":\"ThreatModeler\",\"LibrayId\":0,\"Description\":\"Some ATTR Desc\",\"IsCopy\":false}"}
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

Public Class tmAttribute
    Public Id As Long
    Public Guid? As Guid
    Public Name$
    Public LibraryId As Integer
    Public Options() As tmOptions ' As List(Of tmOptions)

    Public Sub New()
        '        Options = New List(Of tmOptions)
    End Sub
    '    "Id" 1671,
    '    "Guid": "88404d54-fe0a-4a43-a2f8-a0238de44f4e",
    '    "Name": "Are there inputs with potential alternate encoding that could bypass input sanitization?",
    '    "Labels": null,
    '    "Description": "",
    '    "LibraryId": 66,
    '    "isSelected": false,
    '    "IsHidden": false,
    '    "IsOptional": true,
    '    "AttributeType": "SingleSelect",
    '    "PropertyTypeId": 0,
    '    "AttributeAnswer": null,
    '    "Options": [
    '      {
    '        "Id": 1827,
    '        "Name": "Yes",
    '        "IsDefault": false,
    '        "Threats": [
    '          {
    '            "Id": 343,
    '            "Guid": null,
    '            "Name": "Using Slashes and URL Encoding Combined to Bypass Validation Logic",
    '            "Description": null,

End Class
Public Class tmOptions
    Public Id As Integer
    Public Name$
    Public isDefault As Boolean
    Public Threats() As tmProjThreat

    Public Sub New()
        '        Threats = New List(Of tmProjThreat)
    End Sub
End Class
Public Class tm6TF
    Public componentS As List(Of tm6Component)
    Public securityRequirements As List(Of tm6SecReq)
    Public threatS As List(Of tm6Threat)
    ' public testCases as list(of tm6t)
    Public attributeS As List(Of tm6Attribute)
    Public optionS As List(Of tm6Options)
End Class
Public Class tm6Component
    Public id As Long
    Public name As String
    Public description As String
    Public labels As String
    Public version As String
    Public libraryId As Integer
    Public componentTypeId As Integer
    Public componentTypeName As String
    Public guid As String
    Public isHidden As Boolean
    Public isReadOnlyLibraryEntity As Boolean
    Public classDef As tm6CompStructure

    Public Function numDirectSRs() As Integer
        Return classDef.securityRequirements.Count
    End Function
    Public Function numThreats() As Integer
        Return classDef.threats.Count
    End Function
    Public Function numATTR(Optional ByVal localOnly As Boolean = True) As Integer
        If localOnly = False Then
            Return classDef.properties.Count
            Exit Function
        End If
        numATTR = 0
        For Each A In classDef.properties
            If A.isGlobal = False Then numATTR += 1
        Next
    End Function
    Public Function numTL_SR() As Integer
        ' includes numDirect and those attached to threats, but not ATTR
        numTL_SR = 0
        For Each S In classDef.threats
            numTL_SR += S.securityRequirements.Count
        Next
    End Function
    Public Function unqTL_SR() As Integer
        unqTL_SR = 0
        Dim srS As Collection = New Collection
        For Each S In classDef.threats
            For Each SR In S.securityRequirements
                If grpNDX(srS, SR.id) = 0 Then
                    srS.Add(SR.id)
                    unqTL_SR += 1
                    'Else
                    '    Dim K As Integer = 0
                    '    K = 1
                End If
            Next
        Next
    End Function
    Public Function maxTH() As Integer
        maxTH = numThreats()
        For Each A In classDef.properties
            If A.isGlobal = True Then GoTo skipProp
            Dim maxOption As Integer = 0
            For Each O In A.options
                If O.threats.Count > maxOption Then maxOption = O.threats.Count
            Next
            maxTH += maxOption
skipProp:
        Next
    End Function

    Public Function numAttrTH(Optional ByVal localOnly As Boolean = True) As Integer
        numAttrTH = 0
        For Each A In classDef.properties
            If A.isGlobal = True And localOnly = True Then GoTo skipProp
            For Each O In A.options
                numAttrTH += O.threats.Count
            Next
skipProp:
        Next
    End Function

    Public Function numAttrSR(Optional ByVal localOnly As Boolean = True) As Integer
        numAttrSR = 0
        For Each A In classDef.properties
            If A.isGlobal = True And localOnly = True Then GoTo skipProp
            For Each O In A.options
                For Each T In O.threats
                    numAttrSR += T.securityRequirements.Count
                Next
            Next
skipProp:
        Next
    End Function


    Public Function maxSR() As Integer
        maxSR = numTL_SR() + numDirectSRs()

        For Each A In classDef.properties
            If A.isGlobal = True Then GoTo skipProp
            Dim maxOption As Integer = 0
            For Each O In A.options
                For Each T In O.threats
                    If T.securityRequirements.Count > maxOption Then maxOption = T.securityRequirements.Count
                Next
            Next
            maxSR += maxOption
skipProp:
        Next

    End Function
    Public Sub New()
        Me.classDef = New tm6CompStructure
        With Me.classDef
            .threats = New List(Of tm6ThreatShort)
            .securityRequirements = New List(Of tm6SecReqShort)
            .properties = New List(Of tm6AttributesShort)
        End With
    End Sub
End Class
Public Class tm6SecReq
    Public id As Long
    Public name As String
    Public description As String
    Public riskId As Integer
    Public isCompensatingControl As Boolean
    Public riskName As String
    Public labels As String
    Public libraryId As Integer
    Public guid As String
    Public isHidden As Boolean
    Public isReadOnlyLibraryEntity As Boolean
End Class
Public Class tm6SecReqShort
    Public isCompensatingControl As Boolean
    Public applicableOnComponents As List(Of String)
    Public id As Long
    Public name As String
    Public libraryId As Integer
    Public isHidden As Boolean
End Class
Public Class tm6Threat
    Public id As Long
    Public guid As String
    Public name As String
    Public description As String
    Public riskId As Integer
    Public riskName As String
    Public libraryId As Integer
    Public labels As String
    Public automated As Boolean
    Public isHidden As Boolean
    Public isReadOnlyLibraryEntity As Boolean
End Class
Public Class tm6ThreatShort
    Public id As Long
    Public guid As String
    Public name As String
    Public riskId As Integer
    Public riskName As String
    Public riskDisplayName As String
    Public libraryId As Integer
    Public isHidden As Boolean
    Public securityRequirements As List(Of tm6SecReqShort)
End Class

Public Class tm6Attribute
    Public id As Long
    Public guid As String
    Public name As String
    Public labels As String
    Public description As String
    Public libraryId As Integer
    Public isSelected As Boolean
    Public isHidden As Boolean
    Public isOptional As Boolean
    Public attributeType As String
    Public propertyTypeId As Integer
    Public propertyTypeGuid As String
    Public options As List(Of tm6Options)
    Public isReadOnlyLibraryEntity As Boolean
End Class

Public Class tm6Options
    Public id As Integer
    Public name As String
    Public isDefault As Boolean
    Public guid As String
    Public propertyId As Integer
End Class
Public Class tm6OptionsShort
    Public id As Integer
    Public name As String
    Public isDefault As Boolean
    Public threats As List(Of tm6ThreatShort)
End Class
Public Class tm6AttributesShort
    Public id As Integer
    Public guid As String
    Public name As String
    Public libraryId As Integer
    Public isHidden As Boolean
    Public isOptional As Boolean
    Public isGlobal As Boolean
    Public attributeType As String
    Public propertyTypeId As Integer
    Public propertyTypeGuid As String
    Public options As List(Of tm6OptionsShort)
End Class
Public Class tm6CompStructure
    '        "id": 3081,
    '        "guid": "17626326-f596-4bfd-8744-3aa5ef404c37",
    '        "name": "5G security firewall",
    '        "imagePath": "/ComponentImage/images2202006102252443317.png",
    '        "isHidden": false,
    '        "libraryId": 94,
    '        "componentTypeId": 66,
    '        "componentTypeName": "Security Control"
    Public threats As List(Of tm6ThreatShort)
    Public securityRequirements As List(Of tm6SecReqShort)
    Public properties As List(Of tm6AttributesShort)

    Public Sub New()
        threats = New List(Of tm6ThreatShort)
        securityRequirements = New List(Of tm6SecReqShort)
        properties = New List(Of tm6AttributesShort)
    End Sub
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
    Public Guid? As System.Guid
    Public IsHidden As Boolean
    Public DiagralElementId As Integer
    Public ResourceTypeValue$
    'Public Attributes$

    Public listThreats As List(Of tmProjThreat)
    Public listDirectSRs As List(Of tmProjSecReq)
    Public listTransSRs As List(Of tmProjSecReq)
    Public listAttr As List(Of tmAttribute)
    Public isBuilt As Boolean

    Public modelsPresent As List(Of Integer)
    Public numInstancesPerModel As List(Of Integer)

    Public duplicateSRs As Collection

    Public Function numLabels() As Integer
        If Len(Labels) = 0 Then
            Return 0
        Else
            Return numCHR(Labels, ",") + 1
        End If
    End Function

    Public Function getModelNDX(ID As Integer) As Integer
        getModelNDX = -1

        Dim ctR As Integer = 0

        For Each M In modelsPresent
            If M = ID Then
                Return ctR
                Exit Function
            End If
            ctR += 1
        Next

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
        listAttr = New List(Of tmAttribute)
        isBuilt = False
        duplicateSRs = New Collection
        numInstancesPerModel = New List(Of Integer)
        modelsPresent = New List(Of Integer)
    End Sub

End Class
Public Class tmNote
    Public CreatedByName As String
    Public CreatedDate As DateTime
    Public Notes As String
End Class
Public Class tmTask
    'Id: 17157
    'IsCompleted: false
    'IsDeleted: false
    'IsExecuteButtonVisible: false
    'Name: "[@Amir] "
    'Notes: null
    'PriorityColor: "#c80e0e"
    'PriorityId: null
    'ProjectId: 2687
    'ProjectName: "webapp08251234"
    'ResourceTypeValueRelationshipId: null

    Public Id As Long
    Public IsCompleted As Boolean
    Public Name As String
    Public ProjectId As Integer
    Public ProjectName As String
    Public ResourceTypeValueRelationshipId? As Integer


    Public Function isByUser() As Boolean
        If IsNothing(Me.ResourceTypeValueRelationshipId) = True Then Return False Else Return True
    End Function
End Class
Public Class tmWorkflowEvent
    '    "Id": 0,
    '            "Name": "webapp083022",
    '            "Version": "2.0",
    '            "UserId": 0,
    '            "UserName": "amir.soltanzadeh@threatmodeler.com",
    '            "ProjectId": 2693,
    ''            "StatusId": 1,
    '            "StatusName": "Submit For Approval",
    '            "Notes": "Version 2.0 is submitted for approval.",
    '            "DiagramJson": null,
    '            "DateTime": "2022-09-01T13:14:36.75",
    Public Id As Integer
    Public Name As String
    Public UserName As String
    Public ProjectId As Integer
    Public StatusId As Integer
    Public StatusName As String
    Public Notes As String
    Public [DateTime] As DateTime
End Class
Public Class controlAffect
    Public ActionByName As String
    Public ActionDate As DateTime
    Public Action As String
    Public Control As String

    Public Sub New(Notes As List(Of tmNote))
        For Each N In Notes
            Me.ActionByName = N.CreatedByName
            Me.ActionDate = N.CreatedDate

            Dim mitG$ = "Changed Status to Mitigated by Security Control"
            Dim opeN$ = "Changed Status to Open due to removal of Security Control"

            If InStr(N.Notes, mitG) Then
                Me.Action = "Mitigated"
                Control = Mid(N.Notes, Len(mitG) + 2)
            End If
            If InStr(N.Notes, opeN) Then
                Me.Action = "Open"
                Control = Mid(N.Notes, Len(opeN) + 2)
            End If

        Next
    End Sub
End Class

Public Class tmupdateSR
    '[{“Id”:28012,”SecurityRequirementId”:245274,”ProjectId”:2431,”StatusId”:2}]
    'Public Id As Long
    Public Id As Long
    Public SecurityRequirementId As Long
    Public ProjectId As Long
    Public StatusId As Integer
    '    Public ExternalSourceId As String
    Public Sub setStatus(statusName$)
        Select Case statusName
            Case "Open"
                StatusId = 4
            Case "Closed"
                StatusId = 5
            Case "Exception"
                StatusId = 12
            Case "ToDo"
                StatusId = 9
            Case "Recommended"
                StatusId = 13
            Case "Unable To Test"
                StatusId = 14
        End Select
    End Sub
End Class
Public Class getNoteReq
    '{"Id":153193,"ThreatId":81,"ProjectId":2114}
    Public Id As Long
    Public ThreatId As Integer
    Public ProjectId As Integer
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
    Public Guid? As System.Guid
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
    Public listThreats As List(Of tmSimpleThreatSR)
    Public listSecurityRequirements As List(Of tmSimpleThreatSR)
    Public ComponentTypeName$
    Public ComponentName$
    Public Notes$
    Public Id As Long
    Public threatCount As Integer

    Public Sub New()
        Me.listSecurityRequirements = New List(Of tmSimpleThreatSR)
        Me.listThreats = New List(Of tmSimpleThreatSR)

    End Sub
End Class

Public Class tmSimpleThreatSR
    Public Id As Integer
    Public Name$
    Public StatusName$
    Public StatusId As Integer
    Public ElementId As Long
    Public SecurityRequirementId As Long
    Public SecurityRequirementName$

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

