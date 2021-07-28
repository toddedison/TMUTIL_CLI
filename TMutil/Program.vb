Imports System
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Web.HttpUtility
Module Program
    Private T As TM_Client
    Sub Main(args As String())
        Dim uN$ = "mike.horty@threatmodeler.com"
        Dim pW$ = "3642717$Yd!"
        Dim fqdN$ = "https://demo2.threatmodeler.us"

        T = New TM_Client(fqdN, uN, pW)

        If T.isConnected = True Then
            addLoginCreds(T, uN, pW)
            Console.WriteLine("Client connection established")
        Else
            End
        End If

        Dim actionWord$ = ""

        If args.Count.ToString < 1 Or argExist("help", args) Then
            Call giveHelp()
            End
        Else
            actionWord$ = args(0)
            Console.WriteLine("ACTION: " + actionWord + vbCrLf + vbCrLf)
        End If

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
                    Console.WriteLine(col5CLI(P.Name + " [" + P.Id.ToString + "]", P.Type, P.CreatedByName, " ", "Vers " + P.Version))
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
                If vpcID = 0 Then
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

                    End With

                    Console.WriteLine("Sending command to create threat model from:" + vbCrLf + T.prepVPCmodel(vpcS(ndX)))
                    Call T.createVPCmodel(vpcS(ndX))
                    End

nextOne:
                Next


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
                    Dim vDriftVPC$
                    Dim vDriftTM$
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
    Private Sub giveHelp()
        Console.WriteLine("USAGE: TMCLI action --param1 param1_value --param2 param2_value" + vbCrLf)
        Console.WriteLine("ACTIONS:")
        Console.WriteLine("--------")
        Console.WriteLine(fLine("help", "Produces this list of actions and parameters"))
        Console.WriteLine(fLine("get_notes", "Returns notes associated with all components"))
        Console.WriteLine(fLine("get_groups", "Returns groups"))
        Console.WriteLine(fLine("get_projects", "Returns groups"))
        Console.WriteLine(fLine("get_labels", "Returns labels, arg: --ISSYSTEM (True/False)"))
        Console.WriteLine(fLine("summary", "Returns a summary of all Threat Models"))
    End Sub

End Module
