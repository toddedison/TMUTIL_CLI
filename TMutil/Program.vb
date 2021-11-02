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

            Case "get_threats"
                Dim R As New tfRequest

                With R
                    .EntityType = "Threats"
                    .LibraryId = 0
                    .ShowHidden = False
                End With

                T.lib_TH = T.getTFThreats(R)
                Console.WriteLine("Loaded threats: " + T.lib_TH.Count.ToString)
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
                        Console.WriteLine("LABELS     :" + .Labels)
                        Console.WriteLine("DESCRIPTION:" + .Description)

                        Console.WriteLine(vbCrLf)
                        Console.WriteLine("REQUIREMENTS: " + .listLinkedSRs.Count.ToString)

                        For Each S In .listLinkedSRs
                            Dim sNdx As Integer = T.ndxSRlib(S)
                            Dim idSt$ = "[" + T.lib_SR(sNdx).Id.ToString + "]"

                            Console.WriteLine(idSt + spaces(10 - Len(idSt)) + T.lib_SR(sNdx).Name) ' P.CreatedByName + Space(30 - Len(P.CreatedByName)), " ", "Vers " + P.Version))

                        Next
                        Console.WriteLine("==========================")

                    End With
                End If

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

            Case "addcomp_threat"
                Dim cID As Integer = Val(argValue("id", args))
                Dim cNAME = argValue("name", args)
                Dim thrID As Integer = Val(argValue("threatid", args))

                If cID = 0 And cNAME = "" Then
                    Console.WriteLine("Must provide either an --ID or --NAME argument to identify component of interest")
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

                Console.WriteLine("      THREAT: [" + T.lib_Comps(ndxT).CompID.ToString + "] " + T.lib_TH(ndxT).Name + " - # Linked SRs: " + T.lib_TH(ndxT).listLinkedSRs.Count.ToString)

                R.EntityType = "SecurityRequirements"
                T.lib_SR = T.getTFSecReqs(R)



                Call addCompThreat(C, T.lib_TH(ndxT))
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

            Case "show_attr"
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
                Console.WriteLine("Loaded attributes: " + T.lib_AT.Count.ToString)
                Dim numItems As Integer = 0

                For Each P In T.lib_AT
                    If cID And P.Id <> cID Then GoTo skipATT
                    If Len(searcH) And InStr(LCase(P.Name), searcH) = 0 Then GoTo skipATT
                    Dim idSt$ = "[" + P.Id.ToString + "]"
                    Console.WriteLine(idSt + spaces(10 - Len(idSt)) + P.Name) ' P.CreatedByName + Space(30 - Len(P.CreatedByName)), " ", "Vers " + P.Version))
                    numItems += 1
skipATT:
                Next
                Console.WriteLine("# of items in this list: " + numItems.ToString)
                End

            Case "show_comp"
                Dim libName$ = argValue("lib", args)

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

                Console.WriteLine(col5CLI("NAME [ID]", "TYPE", "LIBRARY" + spaces(18), "LABELS"))
                Console.WriteLine(col5CLI("---------", "----", "-------" + spaces(18), "------"))

                For Each C In T.lib_Comps
                    Dim libNDX As Integer = T.ndxLib(C.LibraryId)
                    Dim lName$ = ""
                    If libNDX <> -1 Then lName$ = T.librarieS(libNDX).Name
                    If libName = "" Then GoTo doIT

                    If LCase(lName) <> LCase(libName) Then GoTo dontDOit

doIT:
                    lName += " [" + C.LibraryId.ToString + "]"
                    lName += spaces(30 - Len(lName))
                    numLISTED += 1
                    Console.WriteLine(col5CLI(C.CompName + " [" + C.CompID.ToString + "]", C.ComponentTypeName, lName, C.Labels))
dontDOit:
                Next
                Console.WriteLine("# in LIST: " + numLISTED.ToString)
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

    End Sub


    Private Sub addCompThreat(COMP As tmComponent, thR As tmProjThreat)
        Call T.addThreatToComponent(COMP, thR.Id)

        With COMP

            For Each SR In thR.listLinkedSRs
                'SR represents index
                Dim sNdx As Integer = T.ndxSRlib(SR)
                Console.WriteLine("            SR : " + T.lib_SR(sNdx).Name)

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

                End If

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
                Console.WriteLine("THREAT NAME    : " + thR.Name) ' + " [" + thR.listLinkedSRs.Count.ToString + " SRs]")

                If doEDITS Then
                    Console.WriteLine("Keep this threat? (y/n)")
                    Dim keepThreat As Boolean = True
                    Dim result = Console.ReadKey()
                    Console.SetCursorPosition(0, Console.CursorTop - 1)
                    Console.WriteLine(spaces(80))
                    Console.SetCursorPosition(0, Console.CursorTop - 1)

                    If LCase(result.KeyChar.ToString) = "n" Then
                        Console.WriteLine(vbCrLf + "Removing threat from component")
                        Call T.removeThreatFromComponent(COMP, thR.Id)
                        changeS.Add("Removed threat from Component mapping: " + thR.Name + " [" + thR.Id.ToString + "]")
                        GoTo skipSRsFromThreat
                    Else
                        'Console.WriteLine(vbCrLf + "Keeping threat on component")
                    End If
                End If


                For Each SR In thR.listLinkedSRs
                    'SR represents index
                    Dim sNdx As Integer = T.ndxSRlib(SR)
                    If T.numMatchingLabels(.Labels, T.lib_SR(sNdx).Labels) / .numLabels > 0.9 Then
                        Console.WriteLine("            SR : " + T.lib_SR(sNdx).Name)
                        If doEDITS Then
                            Console.WriteLine("Keep this security requirement on this component? (y/n)")
                            Dim keepThreat As Boolean = True
                            Dim result = Console.ReadKey()

                            Console.SetCursorPosition(0, Console.CursorTop - 1)
                            Console.WriteLine(spaces(80))
                            Console.SetCursorPosition(0, Console.CursorTop - 1)

                            If LCase(result.KeyChar.ToString) = "n" Then
                                Console.WriteLine(vbCrLf + "Removing SR from component's threat")
                                If T.removeLabelFromSR(T.lib_SR(sNdx), COMP.Name) Then
                                    Call T.addEditSR(T.lib_SR(sNdx))
                                    changeS.Add("Removed SR from Threat " + thR.Name + " [" + thR.Id.ToString + "] - '" + COMP.Name + "' label removed from SR: " + T.lib_SR(sNdx).Name + " [" + T.lib_SR(sNdx).Id.ToString + "]")

                                Else
                                    Console.WriteLine("Could not find label '" + COMP.Name + "' on SR")
                                End If
                            Else
                                'Console.WriteLine(vbCrLf + "Keeping SR on component's threat")
                            End If
                        End If

                    End If
                Next
skipSRsFromThreat:

            Next


            If .listAttr.Count Then
                For Each AT In .listAttr
                    For Each O In AT.Options
                        If O.Threats.Count Then
                            Console.WriteLine("-------------------------------")
                            Console.WriteLine("ATTRIBUTE      : " + AT.Name + spaces(100 - Len(AT.Name)) + "OPTION: " + O.Name + " DEFAULT: " + O.isDefault.ToString + " #TH: " + O.Threats.Count.ToString)

                            If doEDITS Then
                                Console.WriteLine("Keep this attribute on the component? (y/n)")
                                Dim keepThreat As Boolean = True
                                Dim result = Console.ReadKey()
                                Console.SetCursorPosition(0, Console.CursorTop - 1)
                                Console.WriteLine(spaces(80))
                                Console.SetCursorPosition(0, Console.CursorTop - 1)

                                If LCase(result.KeyChar.ToString) = "n" Then
                                    Console.WriteLine(vbCrLf + "Removing attribute from the component")
                                    Call T.removeAttributeFromComponent(COMP, AT.Id)
                                    changeS.Add("Removed attribute from Component '" + COMP.Name + "': " + AT.Name + " [" + AT.Id.ToString + "]")

                                    GoTo skipSRsFromThreatATTR
                                Else
                                    Console.WriteLine(vbCrLf + "Keeping attribute")
                                End If
                            End If

                            For Each thR In O.Threats
                                Console.WriteLine("        THREAT : " + thR.Name + " [" + thR.listLinkedSRs.Count.ToString + " SRs]")

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
                                    Dim sNdx As Integer = T.ndxSRlib(S)
                                    If T.numMatchingLabels(.Labels, T.lib_SR(sNdx).Labels) / .numLabels > 0.9 Then
                                        Console.WriteLine("            SR : " + T.lib_SR(sNdx).Name)
                                        If doEDITS Then
                                            Console.WriteLine("Keep this security requirement on this component? (y/n)")
                                            Dim keepThreat As Boolean = True
                                            Dim result = Console.ReadKey()
                                            Console.SetCursorPosition(0, Console.CursorTop - 1)
                                            Console.WriteLine(spaces(80))
                                            Console.SetCursorPosition(0, Console.CursorTop - 1)

                                            If LCase(result.KeyChar.ToString) = "n" Then
                                                Console.WriteLine(vbCrLf + "Removing SR from component")
                                                If T.removeLabelFromSR(T.lib_SR(sNdx), COMP.Name) Then
                                                    '
                                                    Call T.addEditSR(T.lib_SR(sNdx))
                                                    changeS.Add("Removed SR from Threat " + thR.Name + " [" + thR.Id.ToString + "] - '" + COMP.Name + "' label removed from SR: " + T.lib_SR(sNdx).Name + " [" + T.lib_SR(sNdx).Id.ToString + "]")

                                                Else
                                                    Console.WriteLine("Could not find label '" + COMP.Name + "' on SR")
                                                End If
                                            Else
                                                Console.WriteLine(vbCrLf + "Keeping SR on component")
                                            End If
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
            Console.WriteLine("Changes made:" + vbCrLf)
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
        Console.WriteLine(fLine("addcomp_threat", "Add a Threat and choose SRs for a Component, use arg: --ID (component ID) or --NAME (component name)) --THREATID (id)"))
        Console.WriteLine(fLine("get_threats", "Returns threat list or single threat, OPT arg: --ID (component ID), OPT --SEARCH text"))


        Console.WriteLine(fLine("show_aws_iam", "Show all available AWS IAM accounts"))
        Console.WriteLine(fLine("show_vpc", "Show VPCs of an account, arg: --AWSID (Id)"))
        Console.WriteLine(fLine("create_vpc_model", "Create a model from a VPC, arg: --VPCID (Id)"))

    End Sub

End Module
