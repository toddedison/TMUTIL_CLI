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

        Dim GRP As List(Of tmGroups) = T.loadAllGroups(T)

        Dim K As Integer
        K = 1

    End Sub
    Private Sub giveHelp()
        Console.WriteLine("USAGE: VCCLI action --param1 param1_value --param2 param2_value" + vbCrLf)
        Console.WriteLine("ACTIONS:")
        Console.WriteLine("--------")
        Console.WriteLine(fLine("help", "Produces this list of actions and parameters"))
        Console.WriteLine(fLine("get_seclabs_summary", "Returns a summary of Security Labs student & lesson activity"))
        Console.WriteLine(fLine("get_dast", "Returns a summary of DAST analyses and scans"))
    End Sub

End Module
