Imports System
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Web.HttpUtility
Module Program
    Private T As TM_Client
    Private tf_components As List(Of tmComponent)
    Sub Main(args As String())

        Dim actionWord$ = ""

        actionWord$ = args(0)

        If args.Count.ToString < 1 Or argExist("help", args) Then
            Call giveHelp()
            End
        End If

        If actionWord = "makelogin" Then
            Dim fqD$ = argValue("fqdn", args)
            Dim usR$ = argValue("un", args)
            Dim pwD$ = argValue("pw", args)

            If fqD = "" Or usR = "" Or pwD = "" Then
                Console.WriteLine("Must provide all 3 arguments: FQDN, UN and PW")
            End If

            Console.WriteLine(usR + "|" + pwD)
            Call addLoginCreds(fqD, usR, pwD)
            Console.WriteLine("Credentials added to login file - note only the first set of credentials is used")
            End
        End If

        Dim lgnS = getLogins()
        If lgnS.Count = 0 Then
            Console.WriteLine("Make sure you have a login file - the first entry will be used to log in")
            End
        End If

        Dim uN$ = loginItem(lgnS(1), 1)
        Dim pW$ = loginItem(lgnS(1), 2)
        Dim fqdN$ = "https://" + loginItem(lgnS(1), 0) ', "")

        '''''''' Console.WriteLine(uN + "|" + pW)

        T = New TM_Client(fqdN, uN, pW)

        If T.isConnected = True Then
            '            addLoginCreds(fqdN, uN, pW)
            Console.WriteLine("Bearer Token Obtained/ Client connection established")
        Else
            End
        End If

        Console.WriteLine(vbCrLf + "ACTION: " + actionWord + vbCrLf)

        Dim GRP As List(Of tmGroups)

        Select Case LCase(actionWord)
            Case "summary"
                GRP = T.loadAllGroups(T, True)
                End


            Case "get_notes"
                GRP = T.loadAllGroups(T)
                For Each G In GRP
                    For Each PI In G.AllProjInfo
                        Dim NN As New List(Of tmNodeData)

                        For Each nodeD In PI.Model.Nodes
                            If Len(nodeD.Notes) And nodeD.Notes <> "0" Then
                                NN.Add(nodeD)
                            End If
                        Next
                        If NN.Count Then
                            Console.WriteLine("---------------------------------------------")
                            Console.WriteLine(">>>>> Project: " + PI.Name + " [" + PI.Id.ToString + "]")
                            For Each N In NN
                                Console.WriteLine(N.ComponentName + " (" + N.ComponentTypeName + "): " + N.Notes)
                            Next
                            Console.WriteLine("---------------------------------------------")
                        End If
                    Next
                Next
                End

            Case "get_labels"
                Dim LL As List(Of tmLabels)
                Dim isS As Boolean = False
                Console.WriteLine("ARG:" + argValue("issystem", args))
                If argValue("issystem", args) = "" Then
                    Console.WriteLine("ERROR: Must include IsSystem boolean.. eg get_labels --isSystem false")
                    End
                End If
                If LCase(argValue("issystem", args)) = "true" Then isS = True Else isS = False
                LL = T.getLabels(isS)
                For Each L In LL
                    Dim a$ = "System:"
                    If L.IsSystem = True Then a += "True" Else a += "False"
                    Console.WriteLine(fLine("Id:" + L.Id.ToString + " " + a, L.Name))
                Next
                End

            Case "get_threats", "get_threat"
                Dim R As New tfRequest

                With R
                    .EntityType = "Threats"
                    .LibraryId = 0
                    .ShowHidden = False
                End With

                T.lib_TH = T.getTFThreats(R)
                Console.WriteLine("Loaded threats: " + T.lib_TH.Count.ToString)

                R.EntityType = "Components"
                T.lib_Comps = T.getTFComponents(R)
                Console.WriteLine("Loaded comps: " + T.lib_Comps.Count.ToString)

                T.librarieS = T.getLibraries
                Console.WriteLine("Loaded libraries: " + T.librarieS.Count.ToString)

                Dim srchS$ = argValue("search", args)
                If Len(srchS) Then srchS = LCase(srchS)

                Dim numItems As Integer = 0

                If Val(argValue("id", args)) = 0 Then
                    For Each P In T.lib_TH
                        If Len(srchS) Then
                            If InStr(LCase(P.Labels), srchS) Then GoTo showTH
                            If InStr(LCase(P.Name), srchS) Then GoTo showTH
                            'If InStr(LCase(P.Description), srchS) Then GoTo showTH
                            GoTo skipTH
                        End If
showTH:
                        Dim libNDX As Integer = T.ndxLib(P.LibraryId)
                        Dim lName$ = ""
                        If libNDX <> -1 Then lName$ = T.librarieS(libNDX).Name

                        lName += " [" + P.LibraryId.ToString + "]"
                        lName += spaces(30 - Len(lName))

                        Dim idSt$ = "[" + P.Id.ToString + "]"

                        Console.WriteLine(lName + spaces(30 - Len(lName)) + idSt + spaces(10 - Len(idSt)) + P.Name) ' P.CreatedByName + Space(30 - Len(P.CreatedByName)), " ", "Vers " + P.Version))
                        numItems += 1
skipTH:
                    Next
                    Console.WriteLine("# of items in this list: " + numItems.ToString)
                    End
                Else
                    Dim tNDX As Integer = T.ndxTHlib(Val(argValue("id", args)))
                    If tNDX = -1 Then
                        Console.WriteLine("NOT FOUND: Threat ID " + argValue("id", args).ToString)
                        End
                    End If

                    R.EntityType = "SecurityRequirements"
                    T.lib_SR = T.getTFSecReqs(R)
                    Console.WriteLine("Loaded SRs: " + T.lib_SR.Count.ToString)

                    Call T.defineTransSRs(T.lib_TH(tNDX).Id) '.Id)

                    With T.lib_TH(tNDX)
                        Console.WriteLine(vbCrLf + "NAME       :" + .Name + spaces(80 - Len(.Name)) + "[" + .Id.ToString + "] # SR: " + .listLinkedSRs.Count.ToString)
                        Console.WriteLine("LIBRARY    :" + T.librarieS(T.ndxLib(.LibraryId)).Name) '.ToString')
                        Console.WriteLine("GUID       :" + .Guid.ToString)     ')

                        Console.WriteLine("LABELS     :" + .Labels)
                        Console.WriteLine("DESCRIPTION:" + .Description)

                        Console.WriteLine(vbCrLf)
                        Console.WriteLine("REQUIREMENTS: " + .listLinkedSRs.Count.ToString)

                        For Each S In .listLinkedSRs
                            Dim sNdx As Integer = T.ndxSRlib(S)
                            Dim idSt$ = "[" + T.lib_SR(sNdx).Id.ToString + "]"

                            Console.WriteLine(idSt + spaces(10 - Len(idSt)) + T.lib_SR(sNdx).Name) ' P.CreatedByName + Space(30 - Len(P.CreatedByName)), " ", "Vers " + P.Version))

                        Next
                        Console.WriteLine("=================================")

                    End With

                    Console.WriteLine("Components that use this Threat (this is slow - ^C to abort)")
                    Console.WriteLine("=================================")

                    Dim numComps As Integer = 0
                    For Each cc In T.lib_Comps
                        Call T.buildCompObj(cc, True)
                        If T.ndxTHofList(T.lib_TH(tNDX).Id, cc.listThreats) <> -1 Then
                            Console.WriteLine(cc.Name + " [" + cc.Id.ToString + "]")
                            numComps += 1
                        End If
                    Next
                    Console.WriteLine("# of Components with this Threat: " + numComps.ToString)
                End If

            Case "get_sr"
                Dim R As New tfRequest

                With R
                    .EntityType = "SecurityRequirements"
                    .LibraryId = 0
                    .ShowHidden = False
                End With
                T.lib_SR = T.getTFSecReqs(R)

                T.librarieS = T.getLibraries
                Console.WriteLine("Loaded libraries: " + T.librarieS.Count.ToString)

                Dim srchS$ = argValue("search", args)
                If Len(srchS) Then srchS = LCase(srchS)

                Dim numItems As Integer = 0

                If Val(argValue("id", args)) = 0 Then
                    For Each P In T.lib_SR
                        If Len(srchS) Then
                            If InStr(LCase(P.Labels), srchS) Then GoTo showSR
                            If InStr(LCase(P.Name), srchS) Then GoTo showSR
                            GoTo skipSR
                        End If
showSR:
                        Dim libNDX As Integer = T.ndxLib(P.LibraryId)
                        Dim lName$ = ""
                        If libNDX <> -1 Then lName$ = T.librarieS(libNDX).Name

                        lName += " [" + P.LibraryId.ToString + "]"
                        lName += spaces(30 - Len(lName))

                        Dim idSt$ = "[" + P.Id.ToString + "]"

                        Console.WriteLine(lName + spaces(30 - Len(lName)) + idSt + spaces(10 - Len(idSt)) + P.Name) ' P.CreatedByName + Space(30 - Len(P.CreatedByName)), " ", "Vers " + P.Version))
                        numItems += 1
skipSR:
                    Next
                    Console.WriteLine("# of items in this list: " + numItems.ToString)


                    End
                Else

                    Dim sNDX As Integer = T.ndxSRlib(Val(argValue("id", args)))
                    If sNDX = -1 Then
                        Console.WriteLine("NOT FOUND: Security Requirement ID " + argValue("id", args).ToString)
                        End
                    End If
                    With T.lib_SR(sNDX)
                        Console.WriteLine("REQUIREMENT : [" + .Id.ToString + "]" + spaces(10 - Len(.Id.ToString)) + " " + .Name)
                        Console.WriteLine("LIBRARY    :" + T.librarieS(T.ndxLib(.LibraryId)).Name) '.ToString')
                        Console.WriteLine("GUID       :" + .Guid.ToString)     ')

                        Console.WriteLine("DESCRIPTION :" + vbCrLf + .Description)
                        Console.WriteLine("LABELS      :" + .Labels)

                    End With

                    R.EntityType = "Threats"
                    T.lib_TH = T.getTFThreats(R)
                    'Console.WriteLine("Loaded threats: " + T.lib_TH.Count.ToString)

                    Console.WriteLine("Threats that use this SR (this is slow - ^C to abort)")
                    Console.WriteLine("=================================")

                    For Each tT In T.lib_TH
                        Call T.defineTransSRs(tT.Id, True)
                        If grpNDX(tT.listLinkedSRs, T.lib_SR(sNDX).Id) Then
                            Console.WriteLine("        THREAT : [" + tT.Id.ToString + "] " + spaces(10 - Len(tT.Id.ToString)) + tT.Name + " [" + tT.listLinkedSRs.Count.ToString + " SRs Total]")
                        End If
                    Next
                    '                    Console.WriteLine("# of Components with this Threat: " + numComps.ToString)

                End If
                End

            Case "get_groups"
                Dim GG As List(Of tmGroups)
                GG = T.getGroups()
                For Each G In GG
                    Console.WriteLine(fLine("Id:" + G.Id.ToString, G.Name))
                Next
                End


            Case "get_projects"
                Dim PP As List(Of tmProjInfo)
                PP = T.getAllProjects()
                For Each P In PP
                    Console.WriteLine(col5CLI(P.Name + " [" + P.Id.ToString + "]", P.Type, P.CreatedByName + Space(30 - Len(P.CreatedByName)), " ", "Vers " + P.Version))
                Next
                End


            Case "show_aws_iam"
                Dim awsACC As List(Of tmAWSacc)
                awsACC = T.getAWSaccounts()
                For Each A In awsACC
                    Console.WriteLine(A.Id.ToString + spaces(10 - Len(A.Id.ToString)) + A.Name)
                Next

            Case "create_vpc_model"
                Dim vpcID As Long = Val(argValue("vpcid", args))
                If Len(argValue("vpcid", args)) = 0 Then
                    Console.WriteLine("Must provide valid VPC ID value using arg --vpcid")
                    End
                End If
                Dim awsACC As List(Of tmAWSacc)
                awsACC = T.getAWSaccounts()

                For Each A In awsACC
                    Dim vpcS As List(Of tmVPC)
                    'Console.WriteLine("Searching ID " + A.Id.ToString + " " + A.Name)
                    vpcS = T.getVPCs(A.Id)

                    'Console.WriteLine("Looking for VPC ID " + vpcID.ToString + " inside " + A.Name + " (" + vpcS.Count.ToString + " VPCs)")

                    Dim ndX As Integer = T.ndxVPC(vpcID, vpcS)
                    If ndX = -1 Then GoTo nextOne

                    'here is VPC
                    With vpcS(ndX)
                        Console.WriteLine("Found " + vpcID.ToString + " [" + .VPCId + "] inside account " + A.Id.ToString + "-" + A.Name)

                        If IsNothing(.Version) = False Then
                            Console.WriteLine("A model already exists for this VPC (" + fqdN + "/diagram/" + .ProjectId.ToString + ")")
                            End
                        End If

                        .Version = Date.Now.ToString("MdyyyyHHmm")
                        .ProjectName = .VPCId
                        .Type = "AWSCloudApplication"
                        .Labels = "AWS"
                        .GroupPermissions = New List(Of String)
                        .UserPermissions = New List(Of String)
                    End With

                    Console.WriteLine("Sending command to create threat model from:" + vbCrLf + T.prepVPCmodel(vpcS(ndX)))
                    Call T.createVPCmodel(vpcS(ndX))
                    End

nextOne:
                Next
                Console.WriteLine("No VPC Id " + vpcID.ToString + " found")
                End


            Case "show_vpc"
                Dim awsID As Long = Val(argValue("awsid", args))
                If awsID = 0 Then
                    Console.WriteLine("Must provide valid AWS account ID value using arg --awsid")
                    End
                End If
                Dim vpcS As List(Of tmVPC)
                vpcS = T.getVPCs(awsID)

                If IsNothing(vpcS) Or vpcS.Count = 0 Then
                    Console.WriteLine("VPC not found")
                End If

                Console.WriteLine(col5CLI("VPC Name/ID", "Version Datestamp", "DRIFT/VPC", "DRIFT/TM")) ', vVers, vLast, vDriftVPC, vDriftTM))
                Console.WriteLine(col5CLI("===========", "=================", "=========", "========")) ', vVers, vLast, vDriftVPC, vDriftTM))

                For Each V In vpcS
                    Dim vVers$
                    Dim vLast$
                    Dim vDriftVPC$ = ""
                    Dim vDriftTM$ = ""
                    If IsNothing(V.Notes) = False Then
                        If InStr(V.Notes, "Following Nodes are available in AWS VPC") Then vDriftVPC = "TRUE" Else vDriftVPC = "FALSE"
                        If InStr(V.Notes, "Following Nodes are available in Threat Model") Then vDriftTM = "TRUE" Else vDriftTM = "FALSE"
                    End If
                    If IsNothing(V.Version) = False Then
                        vVers = V.Version
                    Else
                        vVers = "Not Created"
                        vDriftVPC = ""
                        vDriftTM = ""
                    End If

                    If IsNothing(V.LastSync) = False Then vLast = V.LastSync.ToString Else vLast = ""
                    Console.WriteLine(col5CLI(V.VPCId + " [" + V.Id.ToString + "]", vVers, vDriftVPC, vDriftTM))
                Next

            Case "attr_report"
                ' shows all entity types and # of threats

                Dim fileN$ = ""
                fileN = argValue("file", args)

                Dim FF As Integer = 0

                If Len(fileN) Then
                    safeKILL(fileN)
                    FF = FreeFile()
                    FileOpen(FF, fileN, OpenMode.Output)
                End If

                T.librarieS = T.getLibraries

                Dim R As tfRequest
                R = New tfRequest

                Console.WriteLine("Test")
                With R
                    .EntityType = "Attributes"
                    .LibraryId = 0
                    .ShowHidden = False
                End With
                T.lib_AT = T.getTFAttr(R)

                Dim c$ = Chr(34)

                R.EntityType = "DataElements"
                Dim TFdataElements As New List(Of tmMiscTrigger)
                TFdataElements = T.getEntityMisc(R)
                R.EntityType = "Roles"
                Dim TFroles As New List(Of tmMiscTrigger)
                TFroles = T.getEntityMisc(R)
                R.EntityType = "Widgets"
                Dim TFwidgets As New List(Of tmMiscTrigger)
                TFwidgets = T.getEntityMisc(R)

                Dim a$ = ""

                For Each D In TFdataElements
                    Dim TH As List(Of tmProjThreat) = T.threatsOfEntity(D, "DataElements")
                    a$ = "DataElement,DataType," + D.Name + "," + TH.Count.ToString + ",1"
                    Console.WriteLine(a$)
                    If FF Then PrintLine(FF, a)
                Next
                For Each tfR In TFroles
                    Dim TH As List(Of tmProjThreat) = T.threatsOfEntity(tfR, "Roles")
                    a$ = "Roles,Role," + tfR.Name + "," + TH.Count.ToString + ",1"
                    Console.WriteLine(a$)
                    If FF Then PrintLine(FF, a)
                Next
                For Each W In TFwidgets
                    Dim widgetThreats As List(Of tmBackendThreats) = T.threatsByBackend(W)

                    Dim backendScounteD As New Collection

                    For Each B In widgetThreats
                        If grpNDX(backendScounteD, B.BackendId) = 0 Then
                            a$ = "Widget," + W.Name + "," + B.BackendName + "," + numWidgetThreats(widgetThreats, B.BackendId).ToString + ",1"
                            Console.WriteLine(a$)
                            If FF Then PrintLine(FF, a)
                            backendScounteD.Add(B.BackendId)
                        End If
                    Next
                Next
                For Each AT In T.lib_AT
                    For Each O In AT.Options
                        Dim aName$ = c + "[" + AT.Id.ToString + "] " + spaces(4 - Len(AT.Id.ToString)) '+ Replace(AT.Name, vbCrLf, "")
                        If Len(AT.Name) > 80 Then aName += Mid(AT.Name, 1, 78) + ".." + c Else aName += AT.Name + c
                        Dim oName$ = c + O.Name + c
                        a$ = "Attribute," + aName + "," + oName + "," + O.Threats.Count.ToString + ",1"
                        Console.WriteLine(a$)
                        If FF Then PrintLine(FF, a)
                    Next
                Next

                If FF Then FileClose(FF)
                End


            Case "addattr_threat"
                Dim cID As Integer = Val(argValue("id", args))
                Dim thrID As Integer = Val(argValue("threatid", args))

                If cID = 0 Then
                    Console.WriteLine("Must provide either an --ID Or --NAME argument to identify component of interest")
                    End
                End If

                If thrID = 0 Then
                    Console.WriteLine("Must provide the --THREATID of the Threat you'd like to add")
                    End
                End If

                Dim R As tfRequest
                R = New tfRequest
                With R
                    .EntityType = "Attributes"
                    .LibraryId = 0
                    .ShowHidden = False
                End With
                T.lib_AT = T.getTFAttr(R)

                Dim ndxC As Integer

                Dim C As tmAttribute

                ndxC = T.ndxATTR(cID)
                If ndxC = -1 Then
                        Console.WriteLine("Attribute does not exist")
                        End
                    End If
                    Console.WriteLine("Found Attribute " + cID.ToString + ": " + T.lib_AT(ndxC).Name)
                    C = T.lib_AT(ndxC)

                R.EntityType = "Threats"
                T.lib_TH = T.getTFThreats(R)

                'Console.WriteLine("Loaded threats: " + T.lib_TH.Count.ToString)
                Dim ndxT As Integer = 0
                ndxT = T.ndxTHlib(thrID)
                If ndxT = -1 Then
                    Console.WriteLine("Threat does not exist")
                    End
                End If
                Call T.defineTransSRs(T.lib_TH(ndxT).Id)

                Console.WriteLine("      THREAT: [" + T.lib_TH(ndxT).Id.ToString + "] " + T.lib_TH(ndxT).Name + " - # Linked SRs: " + T.lib_TH(ndxT).listLinkedSRs.Count.ToString)

                Call T.addThreatToAttribute(C, thrID)
                Console.WriteLine("Added Threat [" + T.lib_TH(ndxT).Id.ToString + "] to Attribute [" + C.Id.ToString + "] ")
                End


            Case "addcomp_threat"
                Dim cID As Integer = Val(argValue("id", args))
                Dim cNAME = argValue("name", args)
                Dim thrID As Integer = Val(argValue("threatid", args))

                If cID = 0 And cNAME = "" Then
                    Console.WriteLine("Must provide either an --ID Or --NAME argument to identify component of interest")
                    End
                End If

                If thrID = 0 Then
                    Console.WriteLine("Must provide the --THREATID of the Threat you'd like to add")
                End
                End If

                Dim R As tfRequest
                R = New tfRequest
                With R
                    .EntityType = "Components"
                    .LibraryId = 0
                    .ShowHidden = False
                End With
                T.lib_Comps = T.getTFComponents(R)

                Dim ndxC As Integer

                Dim C As tmComponent

                If cID Then
                    ndxC = T.ndxComp(cID)
                    If ndxC = -1 Then
                        Console.WriteLine("Component does not exist")
                        End
                    End If
                    Console.WriteLine("Found Component " + cID.ToString + ": " + T.lib_Comps(ndxC).Name + " [" + T.lib_Comps(ndxC).CompID.ToString + "]")
                    C = T.lib_Comps(ndxC)
                Else
                    ndxC = T.ndxCompbyName(cNAME)
                    If ndxC = -1 Then
                        Console.WriteLine("Component does not exist")
                        End
                    End If
                    Console.WriteLine("Found Component: " + T.lib_Comps(ndxC).Name + " [" + T.lib_Comps(ndxC).CompID.ToString + "]")
                    C = T.lib_Comps(ndxC)
                End If

                R.EntityType = "Threats"
                T.lib_TH = T.getTFThreats(R)

                'Console.WriteLine("Loaded threats: " + T.lib_TH.Count.ToString)
                Dim ndxT As Integer = 0
                ndxT = T.ndxTHlib(thrID)
                If ndxT = -1 Then
                    Console.WriteLine("Threat does not exist")
                    End
                End If
                Call T.defineTransSRs(T.lib_TH(ndxT).Id)

                Console.WriteLine("      THREAT: [" + thrID.ToString + "] " + T.lib_TH(ndxT).Name + " - # Linked SRs: " + T.lib_TH(ndxT).listLinkedSRs.Count.ToString)

                R.EntityType = "SecurityRequirements"
                T.lib_SR = T.getTFSecReqs(R)



                Call addCompThreat(C, T.lib_TH(ndxT))
                Console.WriteLine("Threat Added to Component")
                End


            Case "addthreat_sr"
                Dim cID As Integer = Val(argValue("id", args))
                Dim srID As Integer = Val(argValue("srid", args))

                If cID = 0 Then
                    Console.WriteLine("Must provide an --ID to identify Threat of interest")
                    End
                End If

                If srID = 0 Then
                    Console.WriteLine("Must provide the --SRID (Requirement ID) of the SR you'd like to add")
                    End
                End If

                Dim R As tfRequest
                R = New tfRequest
                With R
                    .EntityType = "Threats"
                    .LibraryId = 0
                    .ShowHidden = False
                End With
                T.lib_TH = T.getTFThreats(R)

                Dim ndxC As Integer

                Dim C As tmProjThreat

                ndxC = T.ndxTHlib(cID)
                If ndxC = -1 Then
                        Console.WriteLine("Threat does not exist")
                        End
                    End If

                C = T.lib_TH(ndxC)
                Console.WriteLine("Found Threat " + cID.ToString + ": " + C.Name)

                R.EntityType = "SecurityRequirements"
                T.lib_SR = T.getTFSecReqs(R)

                'Console.WriteLine("Loaded threats: " + T.lib_TH.Count.ToString)
                Dim ndxT As Integer = 0
                ndxT = T.ndxSRlib(srID)
                If ndxT = -1 Then
                    Console.WriteLine("Security Requirement does not exist")
                    End
                End If

                Call T.addSRtoThreat(C, srID)
                Console.WriteLine("SR [" + srID.ToString + "] Added to Threat [" + C.Id.ToString + "]: " + C.Name)
                End

            Case "subthreat_sr"

                Dim cID As Integer = Val(argValue("id", args))
                Dim srID As Integer = Val(argValue("srid", args))

                If cID = 0 Then
                    Console.WriteLine("Must provide an --ID to identify Threat of interest")
                    End
                End If

                If srID = 0 Then
                    Console.WriteLine("Must provide the --SRID (Requirement ID) of the SR you'd like to delete")
                    End
                End If

                Dim R As tfRequest
                R = New tfRequest
                With R
                    .EntityType = "Threats"
                    .LibraryId = 0
                    .ShowHidden = False
                End With
                T.lib_TH = T.getTFThreats(R)

                Dim ndxC As Integer

                Dim C As tmProjThreat

                ndxC = T.ndxTHlib(cID)
                If ndxC = -1 Then
                    Console.WriteLine("Threat does not exist")
                    End
                End If

                C = T.lib_TH(ndxC)
                Console.WriteLine("Found Threat " + cID.ToString + ": " + C.Name)

                R.EntityType = "SecurityRequirements"
                T.lib_SR = T.getTFSecReqs(R)

                'Console.WriteLine("Loaded threats: " + T.lib_TH.Count.ToString)
                Dim ndxT As Integer = 0
                ndxT = T.ndxSRlib(srID)
                If ndxT = -1 Then
                    Console.WriteLine("Security Requirement does not exist")
                    End
                End If

                Call T.addSRtoThreat(C, srID, True)
                Console.WriteLine("SR [" + srID.ToString + "] pulled off Threat [" + C.Id.ToString + "]: " + C.Name)
                End


            Case "addcomp_attr"
                Dim cID As Integer = Val(argValue("id", args))
                Dim cNAME = argValue("name", args)
                Dim attID As Integer = Val(argValue("attrid", args))

                If cID = 0 And cNAME = "" Then
                    Console.WriteLine("Must provide either an --ID Or --NAME argument to identify component of interest")
                    End
                End If

                If attID = 0 Then
                    Console.WriteLine("Must provide the --ATTRID of the Threat you'd like to add")
                    End
                End If

                Dim R As tfRequest
                R = New tfRequest
                With R
                    .EntityType = "Components"
                    .LibraryId = 0
                    .ShowHidden = False
                End With
                T.lib_Comps = T.getTFComponents(R)

                Dim ndxC As Integer

                Dim C As tmComponent

                If cID Then
                    ndxC = T.ndxComp(cID)
                    If ndxC = -1 Then
                        Console.WriteLine("Component does not exist")
                        End
                    End If
                    Console.WriteLine("Found Component " + cID.ToString + ": " + T.lib_Comps(ndxC).Name + " [" + T.lib_Comps(ndxC).CompID.ToString + "]")
                    C = T.lib_Comps(ndxC)
                Else
                    ndxC = T.ndxCompbyName(cNAME)
                    If ndxC = -1 Then
                        Console.WriteLine("Component does not exist")
                        End
                    End If
                    Console.WriteLine("Found Component: " + T.lib_Comps(ndxC).Name + " [" + T.lib_Comps(ndxC).CompID.ToString + "]")
                    C = T.lib_Comps(ndxC)
                End If

                R.EntityType = "Threats"
                T.lib_TH = T.getTFThreats(R)
                R.EntityType = "Attributes"
                T.lib_AT = T.getTFAttr(R)

                'Console.WriteLine("Loaded threats: " + T.lib_TH.Count.ToString)
                Dim ndxAT As Integer = 0
                ndxAT = T.ndxATTR(attID)
                If ndxAT = -1 Then
                    Console.WriteLine("Attribute does not exist")
                    End
                End If

                Dim idSt$ = "[" + T.lib_AT(ndxAT).Id.ToString + "]"
                Console.WriteLine("ATTRIBUTE :" + idSt + spaces(10 - Len(idSt)) + T.lib_AT(ndxAT).Name) ' P.CreatedByName + Space(30 - Len(P.CreatedByName)), " ", "Vers " + P.Version))
                Call T.addAttributeToComponent(C, T.lib_AT(ndxAT).Id)
                Console.WriteLine("========================")
                R.EntityType = "SecurityRequirements"
                T.lib_SR = T.getTFSecReqs(R)

                With T.lib_AT(ndxAT)
                    For Each O In .Options
                        Console.WriteLine("OPTION  : " + O.Name + "     #THR: " + O.Threats.Count.ToString)
                        For Each tt In O.Threats
                            idSt$ = "[" + tt.Id.ToString + "]"
                            Console.WriteLine("========================")
                            Console.WriteLine("THREAT    : " + idSt + spaces(10 - Len(idSt)) + tt.Name) ' P.CreatedByName + Space(30 - Len(P.CreatedByName)), " ", "Vers " + P.Version))
                            Call T.defineTransSRs(tt.Id)
                            For Each tsR In T.lib_TH(T.ndxTHlib(tt.Id)).listLinkedSRs
                                Dim srNDX As Integer = T.ndxSRlib(tsR)
                                Dim SR As tmProjSecReq = T.lib_SR(srNDX)
                                Dim lName$ = T.lib_SR(srNDX).Name
                                Dim isAlreadySR As Boolean = False

                                If T.numMatchingLabels(C.Labels, T.lib_SR(srNDX).Labels) / C.numLabels > 0.9 Then
                                    lName += " [ALREADY APPLIED]"
                                    isAlreadySR = True
                                End If
                                Console.WriteLine("       SR : [" + T.lib_SR(srNDX).Id.ToString + "]" + spaces(10 - Len(T.lib_SR(srNDX).Id.ToString)) + " " + lName)

                                If isAlreadySR Then GoTo escapeHere
showEditSRAgain2:
                                Console.WriteLine("Keep this security requirement on this component? (y/n/desc)")
                                Dim keepThreat As Boolean = True
                                Dim result = Console.ReadKey()

                                Console.SetCursorPosition(0, Console.CursorTop - 1)
                                Console.WriteLine(spaces(80))
                                Console.SetCursorPosition(0, Console.CursorTop - 1)

                                If LCase(result.KeyChar.ToString) = "d" Then
                                    Console.WriteLine(T.lib_SR(srNDX).Description)
                                    GoTo showEditSRAgain2
                                End If

                                If LCase(result.KeyChar.ToString) = "y" Then
                                    If T.matchLabelsOnSR(SR, C.Name) Then
                                        Call T.addEditSR(SR)
                                    Else
                                        Console.WriteLine("Could not match labels on " + SR.Id.ToString + " to component " + C.Id.ToString)
                                    End If
                                End If

                                If LCase(result.KeyChar.ToString) = "n" Then
                                    Dim isDUP As Boolean = False
                                    If isDUP = True Then
                                        Console.WriteLine("** This SR is linked to multiple threats of this Component **")
                                        Console.WriteLine("Press 'O' to OVERRIDE & remove this SR from all of this component's threats..")
                                        result = Console.ReadKey()
                                        Console.SetCursorPosition(0, Console.CursorTop - 2)
                                        Console.WriteLine(spaces(120))
                                        Console.WriteLine(spaces(120))
                                        Console.SetCursorPosition(0, Console.CursorTop - 2)
                                        If result.KeyChar.ToString <> "O" Then
                                            GoTo escapeHere
                                        End If
                                    End If
                                    If T.removeLabelFromSR(SR, C.Name) Then
                                        Call T.addEditSR(SR)
                                    Else
                                        Console.WriteLine("Could not find label '" + C.Name + "' on SR")
                                    End If
                                Else
                                    'Console.WriteLine(vbCrLf + "Keeping SR on component's threat")
                                End If
escapeHere:




                            Next
                        Next
                        Console.WriteLine("========================")
                    Next
                End With

                End


            Case "get_comp"
                Dim cID As Integer = Val(argValue("id", args))
                Dim cNAME = argValue("name", args)

                Dim editSTR$ = argValue("edit", args)
                Dim doEDIT As Boolean = False
                If LCase(editSTR) = "true" Then doEDIT = True

                If cID = 0 And cNAME = "" Then
                    Console.WriteLine("Must provide either an ID or NAME argument to identify component of interest")
                    End
                End If

                Dim R As tfRequest
                R = New tfRequest
                With R
                    .EntityType = "Components"
                    .LibraryId = 0
                    .ShowHidden = False
                End With
                T.lib_Comps = T.getTFComponents(R)

                Dim ndxC As Integer

                Dim C As tmComponent

                If cID Then
                    ndxC = T.ndxComp(cID)
                    If ndxC = -1 Then
                        Console.WriteLine("Component does not exist")
                        End
                    End If
                    Console.WriteLine("Found Component " + cID.ToString + ": " + T.lib_Comps(ndxC).Name + " [" + T.lib_Comps(ndxC).CompID.ToString + "]")
                    C = T.lib_Comps(ndxC)
                Else
                    ndxC = T.ndxCompbyName(cNAME)
                    If ndxC = -1 Then
                        Console.WriteLine("Component does not exist")
                        End
                    End If
                    Console.WriteLine("Found Component: " + T.lib_Comps(ndxC).Name + " [" + T.lib_Comps(ndxC).CompID.ToString + "]")
                    C = T.lib_Comps(ndxC)
                End If
                Call getBuiltComponent(C, doEDIT)
                End

            Case "get_attr"
                Dim R As New tfRequest
                Dim cID As Integer = Val(argValue("id", args))
                Dim searcH$ = argValue("search", args)

                If Len(searcH) Then searcH = LCase(searcH)
                With R
                    .EntityType = "Attributes"
                    .LibraryId = 0
                    .ShowHidden = False
                End With

                T.lib_AT = T.getTFAttr(R)

                Dim numItems As Integer = 0
                Dim IDofAT As Integer = 0

                For Each P In T.lib_AT
                    If cID And P.Id <> cID Then GoTo skipATT
                    IDofAT = P.Id
                    If Len(searcH) And InStr(LCase(P.Name), searcH) = 0 Then GoTo skipATT
                    Dim idSt$ = "[" + P.Id.ToString + "]"
                    Console.WriteLine("ATTRIBUTE :" + idSt + spaces(10 - Len(idSt)) + P.Name) ' P.CreatedByName + Space(30 - Len(P.CreatedByName)), " ", "Vers " + P.Version))
                    numItems += 1
skipATT:
                Next
                If cID = 0 Then Console.WriteLine("# of items in this list: " + numItems.ToString)
                If cID And numItems = 1 Then
                    R.EntityType = "SecurityRequirements"
                    T.lib_SR = T.getTFSecReqs(R)
                    R.EntityType = "Threats"
                    T.lib_TH = T.getTFThreats(R)
                    'Console.WriteLine("Loaded SR,TH object data")

                    Console.WriteLine("========================")
                    With T.lib_AT(T.ndxATTR(IDofAT))
                        For Each O In .Options
                            Console.WriteLine("OPTION  : " + O.Name + "     #THR: " + O.Threats.Count.ToString)
                            For Each tt In O.Threats
                                Dim idSt$ = "[" + tt.Id.ToString + "]"
                                Console.WriteLine("========================")
                                Console.WriteLine("THREAT    : " + idSt + spaces(10 - Len(idSt)) + tt.Name) ' P.CreatedByName + Space(30 - Len(P.CreatedByName)), " ", "Vers " + P.Version))
                                Call T.defineTransSRs(tt.Id)
                                For Each tsR In T.lib_TH(T.ndxTHlib(tt.Id)).listLinkedSRs
                                    Dim srNDX As Integer = T.ndxSRlib(tsR)
                                    Console.WriteLine("       SR : [" + T.lib_SR(srNDX).Id.ToString + "]" + spaces(10 - Len(T.lib_SR(srNDX).Id.ToString)) + " " + T.lib_SR(srNDX).Name)

                                Next
                            Next
                            Console.WriteLine("========================")
                        Next
                    End With

                End If
                End

            Case "show_comp"
                Dim libName$ = argValue("lib", args)
                Dim typeName$ = argValue("type", args)
                Dim searchName$ = argValue("search", args)

                Dim R As tfRequest
                R = New tfRequest
                With R
                    .EntityType = "Components"
                    .LibraryId = 0
                    .ShowHidden = False
                End With
                T.lib_Comps = T.getTFComponents(R)
                T.librarieS = T.getLibraries

                Dim numLISTED As Integer = 0

                Console.WriteLine(col5CLI("NAME [ID]", "TYPE", "LIBRARY", "", spaces(20) + "LABELS"))
                Console.WriteLine(col5CLI("---------", "----", "-------", "", spaces(20) + "------"))

                For Each C In T.lib_Comps
                    Dim libNDX As Integer = T.ndxLib(C.LibraryId)
                    Dim lName$ = ""
                    If libNDX <> -1 Then lName$ = T.librarieS(libNDX).Name
                    If libName = "" And typeName = "" And searchName = "" Then GoTo doIT

                    If Len(libName) Then
                        If LCase(lName) <> LCase(libName) Then GoTo dontDOit
                    End If

                    If Len(typeName) Then
                        If LCase(typeName) <> LCase(C.ComponentTypeName) Then GoTo dontDOit
                    End If

                    If Len(searchName) Then
                        If InStr(LCase(C.CompName), searchName) = 0 Then GoTo dontDOit
                    End If

doIT:
                    lName += " [" + C.LibraryId.ToString + "]"
                    lName += spaces(30 - Len(lName))
                    numLISTED += 1
                    Console.WriteLine(col5CLI(C.CompName + " [" + C.CompID.ToString + "]", C.ComponentTypeName, lName, "", C.Labels))
dontDOit:
                Next
                Console.WriteLine("# in LIST: " + numLISTED.ToString)
                End

            Case "comp_compare"
                Dim cID1 As Integer = Val(argValue("id1", args))
                Dim cID2 As Integer = Val(argValue("id2", args))
                Dim cNAME1 = argValue("name1", args)
                Dim cNAME2 = argValue("name2", args)
                Dim showSR As Boolean = True
                Dim showDIFF As Boolean = False
                Dim shareONLY As Boolean = False

                If LCase(argValue("showsr", args)) = "false" Then showSR = False
                If LCase(argValue("diffonly", args)) = "true" Then showDIFF = True
                If LCase(argValue("shareonly", args)) = "true" Then shareONLY = True

                Dim R As tfRequest
                R = New tfRequest
                With R
                    .EntityType = "Components"
                    .LibraryId = 0
                    .ShowHidden = False
                End With
                T.lib_Comps = T.getTFComponents(R)


                Dim C1 As tmComponent = New tmComponent
                Dim C2 As tmComponent = New tmComponent

                Dim ndxC1 As Integer = 0
                Dim ndxC2 As Integer = 0

                Console.WriteLine("Loading components to compare..")
                If Len(cNAME1) Then
                    ndxC1 = T.ndxCompbyName(cNAME1)
                    If ndxC1 = -1 Then
                        Console.WriteLine("Component " + cNAME1 + " does not exist")
                        End
                    Else
                        C1 = T.lib_Comps(ndxC1)
                    End If
                End If
                If Len(cNAME2) Then
                    ndxC2 = T.ndxCompbyName(cNAME2)
                    If ndxC2 = -1 Then
                        Console.WriteLine("Component " + cNAME2 + " does not exist")
                        End
                    Else
                        C2 = T.lib_Comps(ndxC2)
                    End If
                End If

                If cID1 Then
                    ndxC1 = T.ndxComp(cID1)
                    If ndxC1 = -1 Then
                        Console.WriteLine("Component " + cID1.ToString + " does not exist")
                        End
                    Else
                        C1 = T.lib_Comps(ndxC1)
                    End If
                End If

                If cID2 Then
                    ndxC2 = T.ndxComp(cID1)
                    If ndxC2 = -1 Then
                        Console.WriteLine("Component " + cID2.ToString + " does not exist")
                        End
                    Else
                        C2 = T.lib_Comps(ndxC2)
                    End If
                End If

                If C1.Name = "" Or C2.Name = "" Then
                    Console.WriteLine("Unable to find both components to compare.. Try again.")
                    End
                End If

                Console.WriteLine("Found Component1: " + C1.Name + " [" + C1.CompID.ToString + "]")
                Console.WriteLine("Found Component2: " + C2.Name + " [" + C2.CompID.ToString + "]")

                Call compCompare(C1, C2, showSR, showDIFF, shareONLY)
                '                Call getBuiltComponent(C, False)
                End

            Case "get_libs"

                T.librarieS = T.getLibraries
                Console.WriteLine("Loaded libraries: " + T.librarieS.Count.ToString)

                For Each L In T.librarieS
                    Console.WriteLine(L.Name + " [" + L.Id.ToString + "]")
                Next

            Case "base64"
                Dim fN$ = argValue("file", args)

                If Dir(argValue("file", args)) = "" Or fN = "" Then
                    Console.WriteLine("File not found")
                    End
                End If

                Dim bigStr$ = ""
                Dim FF As Integer = FreeFile()

                FileOpen(FF, fN, OpenMode.Input)

                Do Until EOF(FF) = True
                    bigStr += LineInput(FF)
                Loop

                FileClose(FF)

                bigStr = Replace(bigStr, vbCrLf, "")

                Kill(fN)
                FF = FreeFile()
                FileOpen(FF, fN, OpenMode.Output)
                PrintLine(FF, bigStr)
                FileClose(FF)

                End

            Case "worddesc"
                Dim fN$ = argValue("file", args)

                If Dir(argValue("file", args)) = "" Or fN = "" Then
                    Console.WriteLine("File not found")
                    End
                End If

                Dim bigStr$ = ""
                Dim FF As Integer = FreeFile()

                FileOpen(FF, fN, OpenMode.Input)

                Do Until EOF(FF) = True
                    bigStr += LineInput(FF)
                Loop

                FileClose(FF)

                fN += "_word.htm"

                Dim TFT As TF_Threat = New TF_Threat
                With TFT
                    .CodeTypeId = 1
                    .LibrayId = 66
                    .ComponentTypeId = 85
                    .ImagePath = "/ComponentImage/DefaultComponent.jpg"
                    .IsCopy = False
                    .Name = "SampleName7236"
                    .Description = bigStr
                    .Labels = "DevOps"
                    .RiskId = 1
                    .DataClassificationId = 1
                    'name, labels and description
                End With

                bigStr = JsonConvert.SerializeObject(TFT)

                FF = FreeFile()
                FileOpen(FF, fN, OpenMode.Output)
                PrintLine(FF, bigStr)
                FileClose(FF)

                End


        End Select
        Dim K As Integer
        K = 1
        Console.WriteLine("Your command is not recognized..You may use HELP for a list of commands.")

    End Sub


    Private Sub compCompare(C1 As tmComponent, C2 As tmComponent, Optional ByVal showSR As Boolean = True, Optional ByVal showDIFF As Boolean = False, Optional ByVal shareONLY As Boolean = False)
        tf_components = New List(Of tmComponent)
        For Each C In T.lib_Comps
            tf_components.Add(C)
        Next

        If showSR = False Then Console.WriteLine("Suppressing Security Requirements")
        If showDIFF = True Then Console.WriteLine("Only showing DIFFERENCES")
        If shareONLY = True Then Console.WriteLine("Only showing SHARED properties")

        Dim R As tfRequest = New tfRequest
        With R
            .EntityType = "SecurityRequirements"
            .LibraryId = 0
            .ShowHidden = False
        End With

        T.lib_SR = T.getTFSecReqs(R)
        Console.WriteLine("Loaded security req: " + T.lib_SR.Count.ToString)

        R.EntityType = "Threats"
        T.lib_TH = T.getTFThreats(R)
        Console.WriteLine("Loaded threats: " + T.lib_TH.Count.ToString)

        R.EntityType = "Attributes"
        T.lib_AT = T.getTFAttr(R)
        Console.WriteLine("Loaded attributes: " + T.lib_AT.Count.ToString)
        Call T.buildCompObj(C1)
        Call T.buildCompObj(C2)


        Console.WriteLine(vbCrLf + vbCrLf + "===============================")
        Console.WriteLine(vbCrLf + vbCrLf + "BUILDING " + C1.CompName + " [" + C1.CompID.ToString + "]      Library: " + C1.LibraryId.ToString)
        Console.WriteLine("LABELS: " + C1.Labels)
        Console.WriteLine("DESC  : " + C1.Description)

        With C1
            Console.WriteLine("#1    : " + .Name + spaces(50 - Len(.CompName)) + .listThreats.Count.ToString + " Threats, " + .listDirectSRs.Count.ToString + " Direct SRs, " + .listTransSRs.Count.ToString + " Transitive SRs" + ", " + .listAttr.Count.ToString + " Attr")
        End With
        Console.WriteLine(vbCrLf + vbCrLf + "===============================")
        Console.WriteLine(vbCrLf + vbCrLf + "BUILDING " + C2.CompName + " [" + C2.CompID.ToString + "]      Library: " + C2.LibraryId.ToString)
        Console.WriteLine("LABELS: " + C2.Labels)
        Console.WriteLine("DESC  : " + C2.Description)
        With C2
            Console.WriteLine("#2    : " + .Name + spaces(50 - Len(.CompName)) + .listThreats.Count.ToString + " Threats, " + .listDirectSRs.Count.ToString + " Direct SRs, " + .listTransSRs.Count.ToString + " Transitive SRs" + ", " + .listAttr.Count.ToString + " Attr")
        End With
        Console.WriteLine(vbCrLf + vbCrLf + "===============================")


        With C1

            Dim compSTR$ = ""

            If showSR = False Then GoTo skipDirectSRs
            Console.WriteLine("----DIRECT SECURITY REQS---")

            If .listDirectSRs.Count Then
                For Each DSR In .listDirectSRs
                    If T.ndxSR(DSR.Id, C2.listDirectSRs) = -1 Then
                        compSTR = "[1 ONLY]"
                    Else
                        compSTR = "[SHARED]"
                    End If
                    If showDIFF = True And compSTR = "[SHARED]" Then GoTo skip1

                    If shareONLY = True And compSTR <> "[SHARED]" Then GoTo skip1

                    Console.WriteLine("     " + compSTR + " : " + DSR.Name)
skip1:
                Next
            End If

            If shareONLY Then GoTo skipDirectSRs

            If C2.listDirectSRs.Count Then
                For Each DSR In C2.listDirectSRs
                    If T.ndxSR(DSR.Id, .listDirectSRs) = -1 Then
                        compSTR = "[2 ONLY]"
                        Console.WriteLine("     " + compSTR + " : " + DSR.Name)
                    End If
                Next
            End If

skipDirectSRs:

            Dim allC1Threats As List(Of tmProjThreat)
            Dim allC2Threats As List(Of tmProjThreat)

            allC1Threats = New List(Of tmProjThreat)
            allC2Threats = New List(Of tmProjThreat)

            Console.WriteLine("------THREATS/SRs-------")

            With allC1Threats
                For Each th In C1.listThreats
                    .Add(th)
                Next
                For Each A In C1.listAttr
                    For Each O In A.Options
                        For Each Th In O.Threats
                            .Add(Th)
                        Next
                    Next
                Next
            End With

            With allC2Threats
                For Each th In C2.listThreats
                    .Add(th)
                Next
                For Each A In C2.listAttr
                    For Each O In A.Options
                        For Each Th In O.Threats
                            .Add(Th)
                        Next
                    Next
                Next
            End With

            Console.WriteLine("-------------------------------")
            For Each thR In allC1Threats
                Dim matchedSRs As Collection = T.returnSRsWithLabelMatch(thR, C1)
                Dim c2matchedSRs As New Collection

                compSTR = ""
                If T.ndxTHofList(thR.Id, allC2Threats) = -1 Then
                    compSTR = "[1 ONLY]"
                Else
                    compSTR = "[SHARED]"
                    c2matchedSRs = T.returnSRsWithLabelMatch(thR, C2)
                End If

                If showDIFF = True And compSTR = "[SHARED]" Then GoTo skipSRsFromThreat
                If shareONLY = True And compSTR <> "[SHARED]" Then GoTo skipSRsFromThreat


                Console.WriteLine("THREAT NAME    : [" + thR.Id.ToString + "]" + spaces(10 - Len(thR.Id.ToString)) + " " + thR.Name + " [" + matchedSRs.Count.ToString + " of " + thR.listLinkedSRs.Count.ToString + " SRs] " + compSTR)

                If showSR = False Then GoTo skipSRsFromThreat

                For Each sS In matchedSRs 'these have been label matched as they are threats of the collection
                    'SR represents index
                    Dim isDUP As Boolean = False

                    If grpNDX(c2matchedSRs, sS) = 0 Then
                        compSTR = "[1 ONLY]"
                    Else
                        compSTR = "[SHARED]"
                    End If

                    Dim sNdx As Integer = T.ndxSRlib(sS) ' this is NDX of TMCLIENT Library
                    Dim SR As tmProjSecReq = T.lib_SR(sNdx)
                    ' If T.numMatchingLabels(.Labels, T.lib_SR(sNdx).Labels) / .numLabels > 0.9 Then
                    Dim showName$ = "[" + SR.Id.ToString + "] " + spaces(10 - Len(SR.Id.ToString)) + SR.Name

                    Console.WriteLine("            SR : " + showName + " " + compSTR)
                Next
skipSRsFromThreat:

            Next

            For Each thR In allC2Threats

                compSTR = ""
                If T.ndxTHofList(thR.Id, allC1Threats) = -1 Then
                    Dim matchedSRs As Collection = T.returnSRsWithLabelMatch(thR, C2)

                    compSTR = "[2 ONLY]"

                    If showDIFF = True And compSTR = "[SHARED]" Then GoTo skipSRsFromThreat2
                    If shareONLY = True And compSTR <> "[SHARED]" Then GoTo skipSRsFromThreat2

                    Console.WriteLine("THREAT NAME    : [" + thR.Id.ToString + "]" + spaces(10 - Len(thR.Id.ToString)) + " " + thR.Name + " [" + matchedSRs.Count.ToString + " of " + thR.listLinkedSRs.Count.ToString + " SRs] " + compSTR)

                    If showSR = False Then GoTo skipSRsFromThreat2

                    For Each sS In matchedSRs 'these have been label matched as they are threats of the collection
                        'SR represents index
                        Dim isDUP As Boolean = False

                        Dim sNdx As Integer = T.ndxSRlib(sS) ' this is NDX of TMCLIENT Library
                        Dim SR As tmProjSecReq = T.lib_SR(sNdx)
                        ' If T.numMatchingLabels(.Labels, T.lib_SR(sNdx).Labels) / .numLabels > 0.9 Then
                        Dim showName$ = "[" + SR.Id.ToString + "] " + spaces(10 - Len(SR.Id.ToString)) + SR.Name

                        Console.WriteLine("            SR : " + showName + " " + compSTR)
                    Next
skipSRsFromThreat2:
                End If
            Next

            Console.WriteLine("-------------------------------")

            If .listAttr.Count Then
                For Each AT In .listAttr

                    If T.ndxATTRofList(AT.Id, C2.listAttr) = -1 Then
                        compSTR = "[1 ONLY]"
                    Else
                        compSTR = "[SHARED]"
                    End If
                    If showDIFF = True And compSTR = "[SHARED]" Then GoTo skip3
                    If shareONLY = True And compSTR <> "[SHARED]" Then GoTo skip3

                    Console.WriteLine("ATTRIBUTE      : [" + AT.Id.ToString + "] " + spaces(10 - Len(AT.Id.ToString)) + AT.Name + spaces(100 - Len(AT.Name)) + "    " + compSTR)
skip3:
                Next
            End If

            If shareONLY = True Then GoTo skipthat2

            If C2.listAttr.Count Then
                For Each AT In C2.listAttr

                    If T.ndxATTRofList(AT.Id, C1.listAttr) = -1 Then
                        compSTR = "[2 ONLY]"
                        Console.WriteLine("ATTRIBUTE      : [" + AT.Id.ToString + "] " + spaces(10 - Len(AT.Id.ToString)) + AT.Name + spaces(100 - Len(AT.Name)) + compSTR)
                    End If
                Next
            End If

skipThat2:

        End With
        Console.WriteLine("==============================")



    End Sub

    Public Function numWidgetThreats(ByRef W As List(Of tmBackendThreats), ndxBackEnd As Integer) As Integer
        numWidgetThreats = 0

        For Each backenD In W
            If backenD.BackendId = ndxBackEnd Then numWidgetThreats += 1
        Next

    End Function

    Private Sub addCompThreat(COMP As tmComponent, thR As tmProjThreat)
        Call T.addThreatToComponent(COMP, thR.Id)

        With COMP

            For Each SR In thR.listLinkedSRs
                'SR represents index
                Dim sNdx As Integer = T.ndxSRlib(SR)
                Dim lName$ = T.lib_SR(sNdx).Name
                Dim isAlreadySR As Boolean = False
                If T.numMatchingLabels(COMP.Labels, T.lib_SR(sNdx).Labels) / COMP.numLabels > 0.9 Then
                    lName += " [ALREADY APPLIED]"
                    isAlreadySR = True
                End If

                Console.WriteLine("            SR : [" + T.lib_SR(sNdx).Id.ToString + "] " + lName)

                If isAlreadySR Then GoTo escapehere

                Console.WriteLine("Add this security requirement to the component? (y/n)")
                Dim keepThreat As Boolean = True
                    Dim result = Console.ReadKey()

                    Console.SetCursorPosition(0, Console.CursorTop - 1)
                    Console.WriteLine(spaces(80))
                    Console.SetCursorPosition(0, Console.CursorTop - 1)

                If LCase(result.KeyChar.ToString) = "y" Then
                    'Console.WriteLine(vbCrLf + "Adding SR to component's threat")
                    If T.matchLabelsOnSR(T.lib_SR(sNdx), COMP.Name) Then
                        Call T.addEditSR(T.lib_SR(sNdx))

                    End If
                Else
                    If T.removeLabelFromSR(T.lib_SR(sNdx), COMP.Name) Then                    '
                        Call T.addEditSR(T.lib_SR(sNdx))
                        'make sure label is not there
                    End If
                End If
escapehere:

            Next

        End With
    End Sub

    Private Sub getBuiltComponent(ByRef COMP As tmComponent, Optional ByVal doEDITS As Boolean = False)
        '

        'T.labelS = T.getLabels()
        'Console.WriteLine("Loaded labels: " + T.labelS.Count.ToString)
        'T.sysLabels = T.getLabels(True)
        'Console.WriteLine("Loaded system labels: " + T.sysLabels.Count.ToString)
        'T.groupS = T.getGroups
        'Console.WriteLine("Loaded groups: " + T.groupS.Count.ToString)
        T.librarieS = T.getLibraries
        Console.WriteLine("Loaded libraries: " + T.librarieS.Count.ToString)

        Dim R As tfRequest
        R = New tfRequest
        With R
            .EntityType = "Components"
            .LibraryId = 0
            .ShowHidden = False
        End With

        tf_components = New List(Of tmComponent)
        For Each C In T.lib_Comps
            tf_components.Add(C)
        Next

        With R
            .EntityType = "SecurityRequirements"
            .LibraryId = 0
            .ShowHidden = False
        End With

        T.lib_SR = T.getTFSecReqs(R)
        Console.WriteLine("Loaded security req: " + T.lib_SR.Count.ToString)

        R.EntityType = "Threats"
        T.lib_TH = T.getTFThreats(R)
        Console.WriteLine("Loaded threats: " + T.lib_TH.Count.ToString)

        R.EntityType = "Attributes"
        T.lib_AT = T.getTFAttr(R)
        Console.WriteLine("Loaded attributes: " + T.lib_AT.Count.ToString)

        Call T.buildCompObj(COMP)

        Dim changeS As New Collection

        Console.WriteLine(vbCrLf + vbCrLf + "===============================")
        With COMP

            Dim lName$ = ""
            If T.ndxLib(.LibraryId) <> -1 Then
                lName = T.librarieS(T.ndxLib(.LibraryId)).Name + spaces(25 - Len(.CompName)) + "   LIB ID: " + .LibraryId.ToString
            Else
                lName = "[" + .LibraryId.ToString + "]"
            End If


            Console.WriteLine("DETAIL OF COMPONENT: " + .CompName + spaces(30 - Len(.CompName)) + "      ID: " + .CompID.ToString)
            Console.WriteLine("LIBRARY            : " + lName)
            Console.WriteLine("GUID               : " + .Guid.ToString)

            Console.WriteLine("LABELS             : " + .Labels)
            Console.WriteLine("IMAGE              : " + T.tmFQDN + .ImagePath)
            Console.WriteLine("DESCRIPTION        : " + .Description)
            Console.WriteLine("-------------------------------")
            Console.WriteLine("TRC CONSTRUCT      : " + .listThreats.Count.ToString + " Threats, " + .listDirectSRs.Count.ToString + " Direct SRs, " + .listTransSRs.Count.ToString + " Transitive SRs" + ", " + .listAttr.Count.ToString + " Attr")

            If .listDirectSRs.Count Then
                Console.WriteLine("-------------------------------")
                For Each DSR In .listDirectSRs
                    Console.WriteLine("     DIRECT SR : " + DSR.Name)
                Next
            End If

            Console.WriteLine("-------------------------------")
            For Each thR In .listThreats
                Dim matchedSRs As Collection = T.returnSRsWithLabelMatch(thR, COMP)

                Console.WriteLine("THREAT NAME    : [" + thR.Id.ToString + "]" + spaces(10 - Len(thR.Id.ToString)) + " " + thR.Name + " [" + matchedSRs.Count.ToString + " of " + thR.listLinkedSRs.Count.ToString + " SRs]")


                If doEDITS Then
showEditThreatAgain:
                    Console.WriteLine("Keep this threat? (y/n/desc)")
                    Dim keepThreat As Boolean = True
                    Dim result = Console.ReadKey()
                    Console.SetCursorPosition(0, Console.CursorTop - 1)
                    Console.WriteLine(spaces(80))
                    Console.SetCursorPosition(0, Console.CursorTop - 1)

                    If LCase(result.KeyChar.ToString) = "d" Then
                        Console.WriteLine(thR.Description)
                        GoTo showEditThreatAgain
                    End If

                    If LCase(result.KeyChar.ToString) = "n" Then
                        Console.WriteLine(vbCrLf + "Removing threat from component")
                        Call T.removeThreatFromComponent(COMP, thR.Id)
                        changeS.Add("Removed threat from Component mapping: " + thR.Name + " [" + thR.Id.ToString + "]")
                        GoTo skipSRsFromThreat
                    Else
                        'Console.WriteLine(vbCrLf + "Keeping threat on component")
                    End If
                End If

                For Each sS In matchedSRs 'these have been label matched as they are threats of the collection
                    'SR represents index
                    Dim isDUP As Boolean = False
                    Dim sNdx As Integer = T.ndxSRlib(sS) ' this is NDX of TMCLIENT Library
                    Dim SR As tmProjSecReq = T.lib_SR(sNdx)
                    ' If T.numMatchingLabels(.Labels, T.lib_SR(sNdx).Labels) / .numLabels > 0.9 Then
                    Dim showName$ = "[" + SR.Id.ToString + "] " + spaces(10 - Len(SR.Id.ToString)) + SR.Name
                    If grpNDX(COMP.duplicateSRs, SR.Id) Then
                        isDUP = True
                        showName += " [DUPLICATE]"
                    End If
                    Console.WriteLine("            SR : " + showName)
                    If doEDITS Then
showEditThreatAgain2:
                        Console.WriteLine("Keep this security requirement on this component? (y/n/desc)")
                        Dim keepThreat As Boolean = True
                        Dim result = Console.ReadKey()

                        Console.SetCursorPosition(0, Console.CursorTop - 1)
                        Console.WriteLine(spaces(80))
                        Console.SetCursorPosition(0, Console.CursorTop - 1)

                        If LCase(result.KeyChar.ToString) = "d" Then
                            Console.WriteLine(SR.Description)
                            GoTo showEditThreatAgain2
                        End If

                        If LCase(result.KeyChar.ToString) = "n" Then
                            If isDUP = True Then
                                Console.WriteLine("** This SR is linked to multiple threats of this Component **")
                                Console.WriteLine("Press 'O' to OVERRIDE & remove this SR from all of this component's threats..")
                                result = Console.ReadKey()
                                Console.SetCursorPosition(0, Console.CursorTop - 2)
                                Console.WriteLine(spaces(120))
                                Console.WriteLine(spaces(120))
                                Console.SetCursorPosition(0, Console.CursorTop - 2)
                                If result.KeyChar.ToString <> "O" Then
                                    GoTo escapeHere
                                End If
                            End If
                            Console.WriteLine(vbCrLf + "Removing SR from component's threat")
                            If T.removeLabelFromSR(SR, COMP.Name) Then
                                Call T.addEditSR(SR)
                                changeS.Add("Removed SR from Threat " + thR.Name + " [" + thR.Id.ToString + "] - '" + COMP.Name + "' label removed from SR: " + SR.Name + " [" + SR.Id.ToString + "]")

                            Else
                                Console.WriteLine("Could not find label '" + COMP.Name + "' on SR")
                            End If
                        Else
                            'Console.WriteLine(vbCrLf + "Keeping SR on component's threat")
                        End If
escapeHere:
                    End If

                    'End If
                Next
skipSRsFromThreat:

            Next


            If .listAttr.Count Then
                For Each AT In .listAttr
                    For Each O In AT.Options
                        If O.Threats.Count Then
                            Console.WriteLine("-------------------------------")
                            Console.WriteLine("ATTRIBUTE      : [" + AT.Id.ToString + "] " + spaces(10 - Len(AT.Id.ToString)) + AT.Name + spaces(100 - Len(AT.Name)) + "OPTION: " + O.Name + " DEFAULT: " + O.isDefault.ToString + " #TH: " + O.Threats.Count.ToString)

                            If doEDITS Then
                                Console.WriteLine("Keep this attribute on the component? (y/n)")
                                Dim keepThreat As Boolean = True
                                Dim result = Console.ReadKey()
                                Console.SetCursorPosition(0, Console.CursorTop - 1)
                                Console.WriteLine(spaces(80))
                                Console.SetCursorPosition(0, Console.CursorTop - 1)

                                If LCase(result.KeyChar.ToString) = "n" Then
                                    Console.WriteLine("Removing Attribute:" + AT.Name)
                                    Call T.removeAttributeFromComponent(COMP, AT.Id)
                                    changeS.Add("Removed attribute from Component '" + COMP.Name + "': " + AT.Name + " [" + AT.Id.ToString + "]")

                                    GoTo skipSRsFromThreatATTR
                                Else
                                    'Console.WriteLine(vbCrLf + "Keeping attribute")
                                End If
                            End If

                            For Each thR In O.Threats

                                Dim atThreatMatched As Collection = T.returnSRsWithLabelMatch(thR, COMP)

                                Console.WriteLine("        THREAT : [" + thR.Id.ToString + "] " + spaces(10 - Len(thR.Id.ToString)) + thR.Name + " [" + atThreatMatched.Count.ToString + " of " + thR.listLinkedSRs.Count.ToString + " SRs]")

                                ' Do we really want to modify ATTRIBUTE/THREAT mapping? 
                                '                                If doEDITS Then
                                '                                    Console.WriteLine("Keep this threat on the attribute? (y/n)")
                                '                                    Dim keepThreat As Boolean = True
                                '                                    Dim result = Console.ReadKey()
                                '                                    If LCase(result.KeyChar.ToString) = "n" Then
                                '                                        Console.WriteLine(vbCrLf + "Removing threat from attribute")
                                '                                        changeS.Add("Removed threat from Attribute mapping: " + thR.Name + " [" + thR.Id.ToString + "]")
                                '
                                '           GoTo skipSRsFromThreatATTR
                                '       Else
                                '           Console.WriteLine(vbCrLf + "Keeping threat on attribute")
                                '       End If
                                '   End If

                                For Each S In thR.listLinkedSRs
                                    Dim isDUP As Boolean = False
                                    Dim showName$ = ""
                                    Dim sNdx As Integer = T.ndxSRlib(S)
                                    If T.numMatchingLabels(.Labels, T.lib_SR(sNdx).Labels) / .numLabels > 0.9 Then
                                        showName = "[" + T.lib_SR(sNdx).Id.ToString + "] " + spaces(10 - Len(T.lib_SR(sNdx).Id.ToString)) + T.lib_SR(sNdx).Name
                                        If grpNDX(COMP.duplicateSRs, T.lib_SR(sNdx).Id) Then
                                            isDUP = True
                                            showName += " [DUPLICATE]"
                                        End If
                                        Console.WriteLine("            SR : " + showName)

                                        If doEDITS Then
showEditThreatAgain3:
                                            Console.WriteLine("Keep this security requirement on this component? (y/n/desc)")
                                            Dim keepThreat As Boolean = True
                                            Dim result = Console.ReadKey()
                                            Console.SetCursorPosition(0, Console.CursorTop - 1)
                                            Console.WriteLine(spaces(80))
                                            Console.SetCursorPosition(0, Console.CursorTop - 1)

                                            If LCase(result.KeyChar.ToString) = "d" Then
                                                Console.WriteLine(T.lib_SR(sNdx).Description)
                                                GoTo showEditThreatAgain3
                                            End If

                                            If LCase(result.KeyChar.ToString) = "n" Then

                                                If isDUP = True Then
                                                    Console.WriteLine("** This SR is linked to multiple threats of this Component **")
                                                    Console.WriteLine("Press 'O' to OVERRIDE & remove this SR from all of this component's threats..")
                                                    result = Console.ReadKey()
                                                    Console.SetCursorPosition(0, Console.CursorTop - 2)
                                                    Console.WriteLine(spaces(120))
                                                    Console.WriteLine(spaces(120))
                                                    Console.SetCursorPosition(0, Console.CursorTop - 2)
                                                    If result.KeyChar.ToString <> "O" Then
                                                        GoTo escapeToHere
                                                    End If
                                                End If

                                                'Console.WriteLine(vbCrLf + "Removing SR from component")
                                                If T.removeLabelFromSR(T.lib_SR(sNdx), COMP.Name) Then
                                                    '
                                                    Call T.addEditSR(T.lib_SR(sNdx))
                                                    changeS.Add("Removed SR from Threat " + thR.Name + " [" + thR.Id.ToString + "] - '" + COMP.Name + "' label removed from SR: " + T.lib_SR(sNdx).Name + " [" + T.lib_SR(sNdx).Id.ToString + "]")

                                                Else
                                                    Console.WriteLine("Could not find label '" + COMP.Name + "' on SR")
                                                End If
                                            Else
                                                'Console.WriteLine(vbCrLf + "Keeping SR on component")
                                            End If
escapeToHere:
                                        End If

                                    End If
                                Next

                            Next
                        End If

                    Next
skipSRsFromThreatATTR:

                Next
nope:

            End If
        End With
        Console.WriteLine("==============================")

        If doEDITS = True And changeS.Count > 0 Then
            Console.WriteLine("Changes made: " + changeS.Count.ToString + " total" + vbCrLf)
            For Each cc In changeS
                Console.WriteLine(cc)
            Next
        End If

    End Sub



    Private Function loginItem(ByRef L As String, ByVal itemNum As Integer) As String
        loginItem = ""

        Dim lDetails() As String = Split(L, "|")
        If UBound(lDetails) < 2 Then Exit Function

        If itemNum = 0 Then
            Return lDetails(0)
        Else
            Dim S3 As New Simple3Des("7w6e87twryut24876wuyeg")
            loginItem = S3.Decode(lDetails(itemNum))
        End If

        GC.Collect()
    End Function



    Private Sub giveHelp()
        Console.WriteLine("USAGE: TMCLI action --param1 param1_value --param2 param2_value" + vbCrLf)
        Console.WriteLine("ACTIONS:")
        Console.WriteLine("--------")
        Console.WriteLine(fLine("help", "Produces this list of actions and parameters"))

        Console.WriteLine(fLine("makelogin", "Create Login File, args: --FQDN (demo2.thr...), --UN, --PW"))

        Console.WriteLine(fLine("get_notes", "Returns notes associated with all components"))
        Console.WriteLine(fLine("get_groups", "Returns Groups"))
        Console.WriteLine(fLine("get_projects", "Returns Projects"))
        Console.WriteLine(fLine("get_libs", "Returns Libraries"))

        Console.WriteLine(fLine("get_labels", "Returns Labels, arg: --ISSYSTEM (True/False)"))
        Console.WriteLine(fLine("summary", "Returns a summary of all Threat Models"))

        Console.WriteLine(fLine("show_comp", "Returns list of Components, optional arg: --LIB (library name)"))
        Console.WriteLine(fLine("get_comp", "Returns details of Component, use arg: --ID (component ID) or --NAME (component name)) OPT: --EDIT true"))
        Console.WriteLine(fLine("get_threats", "Returns threat list or single threat, OPT arg: --ID (component ID), OPT --SEARCH text"))
        Console.WriteLine(fLine("get_attr", "Returns list of attributes or single attribute, OPT --ID Id OPT arg: --SEARCH text"))
        Console.WriteLine(fLine("get_sr", "Returns list of SRs or single Security Requirement, OPT --ID Id OPT arg: --SEARCH text"))

        Console.WriteLine(fLine("addcomp_threat", "Add a Threat and choose SRs for a Component, use arg: --ID (component ID) or --NAME (component name)) --THREATID (id)"))
        Console.WriteLine(fLine("addcomp_attr", "Add an Attribute, its Threats and choose SRs for a Component, use arg: --ID (component ID) or --NAME (component name)) --THREATID (id)"))
        Console.WriteLine(fLine("addattr_threat", "Add a Threat and choose SRs for an Attribute, use arg: --ID (attribute ID) --THREATID (id)"))
        Console.WriteLine(fLine("addthreat_sr", "Add a Security Requirement to a Threat, use arg: --ID (threat ID) --SRID (SR ID)"))
        Console.WriteLine(fLine("subthreat_sr", "Remove a Security Requirement from a Threat, use arg: --ID (threat ID) --SRID (SR ID)"))

        Console.WriteLine(fLine("show_aws_iam", "Show all available AWS IAM accounts"))
        Console.WriteLine(fLine("show_vpc", "Show VPCs of an account, arg: --AWSID (Id)"))
        Console.WriteLine(fLine("create_vpc_model", "Create a model from a VPC, arg: --VPCID (Id)"))

    End Sub

End Module
