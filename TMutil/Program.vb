Imports System

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

        Dim GRP As List(Of tmGroups) = T.getGroups

        For Each G In GRP
            G.AllProjInfo = T.getProjectsOfGroup(G)
            Console.WriteLine(G.Name + ": " + G.AllProjInfo.Count.ToString + " models")
            For Each P In G.AllProjInfo
                P.Model = T.getProject(P.Id)
                Console.WriteLine(vbCrLf)
                Console.WriteLine(col5CLI("Proj/Comp Name", "#/TYPE COMP", "# THR", "# SR"))
                Console.WriteLine(col5CLI("--------------", "-----------", "-----", "----"))
                With P.Model
                    Console.WriteLine(col5CLI(P.Name + " [Id " + P.Id.ToString + "]", .Nodes.Count.ToString, totalNumThreats(P.Model).ToString, totalNumThreats(P.Model, True)))
                    '     Console.WriteLine(col5CLI("   ", "TYPE", "# THR", "# SR"))

                    For Each N In .Nodes
                        If N.category = "Collection" Then
                            Console.WriteLine(col5CLI(" >" + N.FullName, "Collection", "", ""))
                            GoTo donehere
                        End If
                        If N.category = "Project" Then
                            Console.WriteLine(col5CLI(" >" + N.FullName, "Nested Model", "", ""))
                            GoTo doneHere
                        End If
                        Console.WriteLine(col5CLI(" >" + N.ComponentName, N.ComponentTypeName, N.threatCount.ToString, N.listSecurityRequirements.Count.ToString))
                        '                        Console.WriteLine("------------- " + N.Name + "-Threats:" + N.listThreats.Count.ToString + "-SecRqrmnts:" + N.listSecurityRequirements.Count.ToString)
doneHere:
                    Next
                End With
                If P.Model.Nodes.Count Then Console.WriteLine(col5CLI("----------------------------------", "------------------------", "------", "------"))
            Next

            Console.WriteLine(vbCrLf)
        Next

        Dim K As Integer
        K = 1

    End Sub
    Public Function totalNumThreats(ByVal M As tmModel, Optional ByVal SecReqInstead As Boolean = False) As Integer
        Dim nuM As Integer = 0
        For Each N In M.Nodes
            If SecReqInstead Then nuM += N.listSecurityRequirements.Count Else nuM += N.listThreats.Count
        Next
        Return nuM
    End Function
    Private Sub giveHelp()
        Console.WriteLine("USAGE: VCCLI action --param1 param1_value --param2 param2_value" + vbCrLf)
        Console.WriteLine("ACTIONS:")
        Console.WriteLine("--------")
        Console.WriteLine(fLine("help", "Produces this list of actions and parameters"))
        Console.WriteLine(fLine("get_seclabs_summary", "Returns a summary of Security Labs student & lesson activity"))
        Console.WriteLine(fLine("get_dast", "Returns a summary of DAST analyses and scans"))
    End Sub
    Private Function fLine(arg1$, arg2$, Optional ByVal numSpaces As Integer = 25) As String
        Return arg1 + spaces(numSpaces - Len(arg1)) + arg2
    End Function

    Private Function col5CLI(arg1$, arg2$, arg3$, arg4$, Optional ByVal arg5$ = "") As String
        col5CLI = ""
        arg1 = Mid(arg1, 1, 34)
        arg2 = Mid(arg2, 1, 24)
        col5CLI = arg1 + spaces(35 - Len(arg1)) + arg2 + spaces(25 - Len(arg2)) + arg3 + spaces(10 - Len(arg3))
        col5CLI += arg4 + spaces(10 - Len(arg4)) + arg5 + spaces(10 - Len(arg5))


    End Function
End Module
