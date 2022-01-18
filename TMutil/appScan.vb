Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq


Public Class appScan
    Private sourceDirectory$
    Private fileBlockFilter$
    Private methodCounter As Counter
    Private fileCounter As Counter
    Private containerCounter As Counter
    Private graphCounter As Counter

    Private depthChangers As Collection
    Private methodDescriptors As Collection

    Private sourceFiles As List(Of SourceFile)
    Private codeContainers As List(Of CodeCollection)
    Private allMethods As List(Of Method)
    Private allClasses As List(Of ClassS)


    Private Class Counter
        Private currCount As Integer

        Public Sub New()
            currCount = 0
        End Sub

        Public Function addOne() As Integer
            currCount += 1
            Return currCount
        End Function
        Public Function getCount() As Integer
            Return currCount
        End Function
    End Class

    Private Class CodeCollection
        Public Id As Integer
        Public sourceFilename$
        Public ccType$ ' eg 'Module', 'Class'. 'File'
        Public Name$
        Public deptH As Integer
        Public patH As String
        Public ofParents As String
        Public graphId As Integer
        Public sourceLines As List(Of String)
        Public allMethods As List(Of Method)
        Public allClasses As List(Of ClassS)
        Public classRefs As List(Of classInst)
        Public Sub New()
            allMethods = New List(Of Method)
            allClasses = New List(Of ClassS)
            sourceLines = New List(Of String)
            classRefs = New List(Of classInst)
        End Sub

        Public Function totalLOC() As Long
            totalLOC = sourceLines.Count
            For Each M In allMethods
                totalLOC += M.sourceLines.Count
            Next
        End Function
    End Class

    Private Class Method
        Public Id As Integer
        Public Name As String
        Public sourceFilename$
        Public isPublic As Boolean
        Public doesReturn As Boolean
        Public returnType As String
        Public graphId As Integer
        Public patH As String
        Public ofParent As String 'what parents (csv) does this method belong to?
        Public parametersExpected As String
        Public sourceLines As List(Of String)
        Public classRefs As List(Of classInst)
        Public Sub New()
            sourceLines = New List(Of String)
            classRefs = New List(Of classInst)
        End Sub
    End Class

    Private Class classInst
        Public varName As String
        Public classPath As String
        Public isPublic As Boolean
    End Class
    Private Class Parameter
        Public Name As String
        Public pType As String
        Public isRequired As Boolean
    End Class

    Private Class SourceFile
        Public Id As Integer
        Public allMethods As List(Of Method)
        Public allClasses As List(Of ClassS)
        Public graphId As Integer
        Public fileName$
        Public linesOfCode$
        Public LOC As Long 'lines of code

        Public Sub New(fileN$, unqID As Integer, graphNdx As Integer)
            allMethods = New List(Of Method)
            allClasses = New List(Of ClassS)
            fileName = fileN
            Id = unqID
            graphId = graphNdx
        End Sub
    End Class

    Private Class ClassS
        Public Name As String
        Public sourceFilename$
        Public isPublic As Boolean
        Public path As String
        Public ofParent As String 'could be blank if depth=1
        Public classObjects As List(Of ClassObj)
        Public usedBy As List(Of Method)
        Public graphId As Integer
        Public classRefs As List(Of classInst)
        Public Sub New()
            classRefs = New List(Of classInst)
        End Sub

    End Class

    Private Class ClassObj
        Public Name$
        Public oType As String '
    End Class



    Public Sub doScan(dirToScan$, files2Block$)
        methodCounter = New Counter
        fileCounter = New Counter
        containerCounter = New Counter
        graphCounter = New Counter

        codeContainers = New List(Of CodeCollection)
        sourceFiles = New List(Of SourceFile)
        allMethods = New List(Of Method)
        allClasses=new list(of classs)

        Call getSourceFiles(dirToScan, files2Block)


        For Each S In sourceFiles
            Call loadCode(S)
            'Console.WriteLine("[ID " + S.Id.ToString + "/LOC " + S.LOC.ToString + "/CHR " + Len(S.linesOfCode).ToString + "]: " + S.fileName)

            Dim CC As New CodeCollection
            With CC
                .Name = stripToFilename(S.fileName)
                .deptH = 0
                .patH = "|FILE|" + S.fileName
                .ccType = "File"
                .Id = containerCounter.addOne
                .graphId = graphCounter.addOne
                .sourceFilename = S.fileName
                .ofParents = ""
            End With
            codeContainers.Add(CC)
        Next

        Dim K As Integer = 0
        For K = 0 To sourceFiles.Count - 1
            Call scanCodeForStructure(K)
            sourceFiles(K).linesOfCode = ""
            GC.Collect()
        Next K

        Call scanCodeForMapping()

        '     For Each cc In codeContainers
        '         If cc.allMethods.Count Then Console.WriteLine("[" + cc.graphId.ToString + "/" + cc.totalLOC.ToString + " LOC/" + cc.allMethods.Count.ToString + " Methods]: " + cc.patH)
        '     Next


        Dim newTM As makeTMJson = New makeTMJson
        Dim jSon$ = newTM.getJSON(Me)
        Call safeKILL("jsonfile.txt")
        saveJSONtoFile(jSon, "jsonfile.txt")
        'Console.WriteLine(jSon)
    End Sub

    Private Function ndxCodeCollection(pathOfCC) As Integer
        ndxCodeCollection = -1
        Dim ndX As Integer = 0
        For Each C In codeContainers
            If C.patH = pathOfCC Then
                Return ndX
            End If
            ndX += 1
        Next
    End Function

    Private Function ofCodeCollection(pathOfCC) As CodeCollection
        If pathOfCC = "" Then
            Return New CodeCollection
        End If
        pathOfCC = noComma(pathOfCC)

        If Mid(pathOfCC, Len(pathOfCC), 1) = "," Then pathOfCC = Mid(pathOfCC, 1, Len(pathOfCC) - 1)
        ofCodeCollection = New CodeCollection

        For Each C In codeContainers
            If C.patH = pathOfCC Then
                Return C
            End If
        Next
    End Function

    Private Function findGraphId(pathOfCC) As Integer
        pathOfCC = noComma(pathOfCC)
        findGraphId = 0
        findGraphId = ofCodeCollection(pathOfCC).graphId

        If findGraphId = 0 Then findGraphId = ofMethod(pathOfCC).graphId
    End Function

    Private Function ofMethod(pathOfCC) As Method
        pathOfCC = noComma(pathOfCC)

        ofMethod = New Method

        For Each C In allMethods
            If C.patH = pathOfCC Then
                Return C
            End If
        Next
    End Function

    Private Sub getSourceFiles(dirToScan$, files2Block$)
        Dim Directories As New IO.DirectoryInfo(dirToScan)
        Dim fileInfo = Directories.GetFiles("*.vb", IO.SearchOption.AllDirectories)

        Dim bC As New Collection

        bC = CSVtoCOLL(files2Block)

        For Each file In fileInfo
            Dim blockThis As Boolean = False
            For Each blk In bC
                If InStr(LCase(file.FullName), LCase(blk)) Then blockThis = True
            Next
            If blockThis = False Then
                ' These become our source files to scan
                '                Console.WriteLine("File: " + file.FullName)
                fileCounter.addOne()
                sourceFiles.Add(New SourceFile(file.FullName, fileCounter.getCount, graphCounter.addOne))
            End If
        Next
    End Sub
    Private Sub loadCode(S As SourceFile)
        S.linesOfCode = streamReaderTxt(S.fileName)
        S.LOC = countChars(S.linesOfCode, Chr(13))
        GC.Collect()
    End Sub

    Public Sub New()
        'setting up AppScan class

        'define depthChangers - these are structural changes that determine depth
        '
        ' file
        ' class XYZ    <depth 0
        ' sub         <depth 1
        ' function     < still depth 1
        ' class ABC    <depth 2
        ' sub          < depth 2 - sub of a class

        'VB stuff
        depthChangers = New Collection
        depthChangers.Add("Public Class ")
        depthChangers.Add("Private Class ")
        depthChangers.Add("Module ")
        depthChangers.Add("Class ")
        depthChangers.Add("End Class")
        depthChangers.Add("End Module")
        depthChangers.Add("Public NotInheritable Class ")
        depthChangers.Add("Private NotInheritable Class ")

        methodDescriptors = New Collection
        methodDescriptors.Add("Public Sub ")
        methodDescriptors.Add("Private Sub ")
        methodDescriptors.Add("Public Function ")
        methodDescriptors.Add("Private Function ")
        methodDescriptors.Add("Function ")
        methodDescriptors.Add("Sub ")
        methodDescriptors.Add("End Function")
        methodDescriptors.Add("End Sub")

        graphCounter = New Counter

    End Sub

    Private Function removeScope(ByVal scopeStr$) As String
        removeScope = ""
        ' scope is |A|Upper Lvl Obj,|B|Secondary,|C|Tertiary and so on
        'removing pulls off outermost object

        If InStr(scopeStr, ",") = 0 Then Return ""

        If Mid(scopeStr, Len(scopeStr), 1) = "," Then scopeStr = Mid(scopeStr, 1, Len(scopeStr) - 1)

        Dim C As New Collection
        C = CSVtoCOLL(scopeStr)

        Dim newScope$ = ""
        Dim K As Integer = 0

        For K = 1 To C.Count - 1
            newScope += C(K) + ","
        Next

        Return newScope
    End Function
    Private Sub scanCodeForStructure(ByVal sourceNum As Integer)

        With sourceFiles(sourceNum)

            Dim chrNum As Long = 1
            Dim LOC$ = ""
            Dim currScope$ = ""
            Dim deptH As Integer = 0

            Dim foundSomething As Boolean
            Dim nameOfObj$ = ""
            Dim methParams$ = ""

            Dim currMethod As Integer = 0

            Dim currPath As String = ""
            Dim inMethod As Boolean = False

            Dim currPathCC As Integer = 0
            currPathCC = ndxCodeCollection("|FILE|" + sourceFiles(sourceNum).fileName)

            currScope = codeContainers(currPathCC).patH + ","

            If InStr(.linesOfCode, Chr(13)) = 0 Then
                Exit Sub
            End If


            ' depth at file level is 0
            ' some languages (python) put executable code at depth 0, though first many lines of code are methods
            ' others (.net) put their dependency imports at depth 0
            ' |FILE| |MOD| |CLASS| |METHOD|
            ' currScope string will follow depth integer
            ' increases any time parser goes deeper into a MOD CLASS METHOD
            ' Path string navigation will determine grouping of model components
            ' source stored to either codecollection or method

            ' also build classRefs list
            ' represents variables declared as classes

            LOC = Mid(.linesOfCode, chrNum, InStr(Mid(.linesOfCode, chrNum), Chr(13)))
            chrNum += Len(LOC)

            Do Until InStr(LOC, Chr(13)) = 0
                'Console.WriteLine(Mid(LOC, 1, Len(LOC) - 1)) ' -1 to avoid extra LF

                foundSomething = False
                ' here 'FoundSomething' references anything that affects current object path
                ' currscope manipulated appropriately
                For Each D In depthChangers
                    If Mid(LTrim(LOC), 1, Len(D)) = D Then
                        foundSomething = True

                        nameOfObj$ = Replace(Trim(Mid(LTrim(LOC), Len(D))), vbCr, "")

                        If Mid(LTrim(LOC), 1, 4) = "End " Then
                            deptH -= 1
                            currScope = removeScope(currScope)
                        Else
                            deptH += 1
                            Dim pFx$ = "MOD"
                            If InStr(LOC, " Class ") Then pFx = "CLS"

                            Dim CC As New CodeCollection
                            CC.sourceFilename = .fileName
                            CC.Id = containerCounter.addOne
                            CC.ofParents = noComma(currScope)
                            CC.patH = noComma(currScope) + ",|" + pFx + "|" + nameOfObj
                            CC.Name = nameOfObj
                            If pFx = "MOD" Then
                                CC.ccType = "Module"
                            Else
                                CC.ccType = "Class"

                                Dim newClass As ClassS = New ClassS
                                With newClass
                                    .ofParent = noComma(currScope)
                                    .path = CC.patH
                                    .Name = nameOfObj
                                    .sourceFilename = CC.sourceFilename
                                    If InStr(LTrim(LOC), "Public ") Then .isPublic = True
                                    ' no graph id because class graphed as codecollection obj
                                End With
                                allClasses.Add(newClass)
                                ' adds discovered class -
                            End If
                            CC.deptH = deptH
                            CC.graphId = graphCounter.addOne

                            codeContainers.Add(CC)

                            currScope = CC.patH + ","

                        End If

                        GoTo skipMethLook
                    End If
                Next
                For Each M In methodDescriptors
                    If Mid(LTrim(LOC), 1, Len(M)) = M Then
                        foundSomething = True

                        If Mid(LTrim(LOC), 1, 3) = "End" Then
                            inMethod = False
                            currScope = removeScope(currScope)
                            GoTo skipMethLook
                        End If

                        inMethod = True

                        nameOfObj$ = Replace(Trim(Mid(LTrim(LOC), Len(M))), vbCr, "")

                        methParams = Mid(nameOfObj, InStr(nameOfObj, "("))
                        nameOfObj = Mid(nameOfObj, 1, InStr(nameOfObj, "(") - 1)

                        ' block some here
                        'If nameOfObj = "New" Then GoTo skipMethLook

                        'this is a new method
                        Dim newM As New Method
                        newM.Name = nameOfObj
                        newM.patH = noComma(currScope) + ",|METH|" + nameOfObj
                        If InStr(LTrim(LOC), "Function ") Then
                            newM.doesReturn = True
                        End If
                        newM.Id = methodCounter.addOne()
                        newM.ofParent = noComma(currScope)
                        newM.sourceFilename = .fileName
                        If InStr(LTrim(LOC), "Public ") Then
                            newM.isPublic = True
                        End If
                        newM.graphId = graphCounter.addOne
                        allMethods.Add(newM)

                        currScope = newM.patH

                        GoTo skipMethLook
                    End If
                Next
skipMethLook:

                If foundSomething Then
                    Dim a$ = "FOUND SOMETHING: "
                    a += currScope
                    'Console.WriteLine(a)
                    Dim somethiung As Integer = 0

                End If


                ' Here the code is applied to either the codecollection or method
                LOC = Replace(LOC, vbCr, "")

                If inMethod = True Then
                    Dim M As Method = ofMethod(currScope)
                    M.sourceLines.Add(LOC)
                Else
                    Dim M As CodeCollection = ofCodeCollection(currScope)
                    M.sourceLines.Add(LOC)
                End If

                LOC = Mid(.linesOfCode, chrNum, InStr(Mid(.linesOfCode, chrNum), Chr(13)))
                chrNum += Len(LOC) + 1

            Loop

        End With

        GoTo skipDebug

        For Each cc In codeContainers
            ' Console.WriteLine("[" + cc.Id.ToString + "/" + cc.sourceLines.Count.ToString + " LOC]: " + cc.patH)

            For Each mm In allMethods
                If InStr(mm.patH, cc.patH) Then
                    'Console.WriteLine("    [" + mm.Id.ToString + "/" + mm.sourceLines.Count.ToString + " LOC]: " + mm.Name)
                End If
            Next

        Next

skipDebug:

        'fix this.. duplicate method into every parent objecT?
        'seems to be adding methods more than once
        For Each mm In allMethods
            With ofCodeCollection(mm.ofParent)
                .allMethods.Add(mm)
            End With
        Next



    End Sub
    Private Sub scanCodeForMapping()

        For Each CC In codeContainers
            For Each L In CC.sourceLines
                'Console.WriteLine(L)
                Dim cR As classInst = detectClassRef(L)
                If Len(cR.classPath) Then
                    Console.WriteLine("Found classref: " + CC.Name + " uses var " + cR.varName + " for class " + ofCodeCollection(cR.classPath).Name)
                    CC.classRefs.Add(cR)
                    ' probably need to add this to classs too 
                End If
            Next
        Next

        For Each MM In allMethods
            For Each L In MM.sourceLines
                'Console.WriteLine(L)
                Dim cR As classInst = detectClassRef(L)
                If Len(cR.classPath) Then
                    Console.WriteLine("Found classref: " + MM.Name + " uses var " + cR.varName + " for class " + ofCodeCollection(cR.classPath).Name)
                    MM.classRefs.Add(cR)
                End If
            Next
        Next

    End Sub

    Private Function detectClassRef(L$) As classInst
        detectClassRef = New classInst
        'if classref, returns class path
        Dim maybeClass$ = ""
        Dim vName$ = ""

        If InStr(L, "(Of ") Then
            maybeClass = trimVal(Mid(L, InStr(L, "(Of ") + 4), ")")
            GoTo checkMaybe
        End If

        If InStr(L, " As ") Then ' And InStr(L, " New ") = 0 And InStr(L, "(Of ") = 0 Then
            maybeClass = Mid(L, InStr(L, " As") + 4)
            '            vName = Mid(L, 1, InStr(L, InStr(L, " As") - 1))
            '            vName = previousWord(vName)
            GoTo checkMaybe
        End If

        If InStr(L, " New ") Then
            maybeClass = Mid(L, InStr(L, " New ") + 7)
            '           vName = Replace(Replace(Replace(vName, "Public ", ""), "Private ", ""), "Dim ", "")
            GoTo checkMaybe
        End If


checkMaybe:
        If maybeClass = "" Then Exit Function

        maybeClass = Replace(Trim(maybeClass), vbCr, "")
        vName = Replace(Replace(Replace(Trim(L), "Public ", ""), "Private ", ""), "Dim ", "")
        vName = trimVal(vName, " (")

        For Each CC In codeContainers
            If CC.ccType = "Class" Then
                If maybeClass = CC.Name Then
                    With detectClassRef
                        .classPath = CC.patH
                        .varName = vName
                        If InStr(L, "Public ") Then .isPublic = True
                    End With
                End If
            End If
        Next

foundIt:
    End Function

    Public Class makeTMJson
        '    "ObjectId":9,
        '        "ObjectType":"component",
        '        "ParentGroupId":2,
        '        "ComponentGuid":"F47FCE3A-C854-4A90-88C1-0D2C005BC511",
        '        "Name":"Akamai CDN",
        '        "OutBoundId":[10],
        '        "OutBoundGuids":["31D863C7-8E0B-49F2-9363-539ED8E97E72"]

        Public Class tmObject
            Public ObjectId As Integer
            Public ObjectType As String
            Public ParentGroupId As Integer
            Public ComponentGuid As String
            Public Name As String
            Public OutBoundId As List(Of Integer)
            Public OutBoundGuids As List(Of guidList)

            Public Sub New()
                OutBoundId = New List(Of Integer)
                OutBoundGuids = New List(Of guidList)
            End Sub

        End Class
        Public Class guidList
            Public ListOfGuids As List(Of String)
            Public Sub New()
                ListOfGuids = New List(Of String)
            End Sub
        End Class

        Private Function getTmObject(objId As Integer, objType$, objName$, parentId As Integer) As tmObject
            getTmObject = New tmObject
            With getTmObject
                .Name = objName
                .ObjectId = objId
                .ParentGroupId = parentId
                If LCase(objType) = "method" Then .ObjectType = "Component" Else .ObjectType = "Container"

                Select Case LCase(objType)
                    Case "method"
                        .ComponentGuid = "deba4981-1f92-4b3b-94c3-e603a660aff4"
                    Case "class"
                        .ComponentGuid = "cb374e7d-fdc5-4f0a-a164-a82e52a1a6ca"
                    Case "module"
                        .ComponentGuid = "9aec34ec-1715-431b-bc5e-e1ff40aa2b88"
                    Case "file"
                        .ComponentGuid = "1c912bfb-cc5a-495d-84f3-79fdd2441132"
                        .Name = stripToFilename(.Name)
                End Select
            End With
        End Function

        Public Function getJSON(ByRef A As appScan) As String
            Dim allObjects As List(Of tmObject) = New List(Of tmObject)

            Dim parentObjects As Collection = New Collection

            'For Each CC In A.sourceFiles
            '    allObjects.Add(getTmObject(CC.graphId, "file", stripToFilename(CC.fileName), 0))
            'Next
            Dim unqIDs As New Collection

            For Each CC2 In A.codeContainers
                If grpNDX(unqIDs, CC2.graphId) Then
                    Console.WriteLine("Duplicate at CC " + CC2.graphId.ToString)
                Else
                    unqIDs.Add(CC2.graphId)
                    allObjects.Add(getTmObject(CC2.graphId, CC2.ccType, CC2.Name, A.findGraphId(CC2.ofParents)))
                End If
                For Each MM In CC2.allMethods

                    If grpNDX(unqIDs, MM.graphId) Then
                        'Console.WriteLine("Duplicate at " + MM.graphId.ToString)
                    Else
                        unqIDs.Add(MM.graphId)
                        allObjects.Add(getTmObject(MM.graphId, "method", MM.Name, A.findGraphId(MM.ofParent)))
                    End If
                Next
            Next

            Dim jBody$ = ""
            jBody = JsonConvert.SerializeObject(allObjects)

            Return jbody
        End Function
    End Class


End Class

