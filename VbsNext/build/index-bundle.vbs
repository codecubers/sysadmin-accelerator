

Option Explicit

Dim debug: debug = (WScript.Arguments.Named("debug") = "true")
if (debug) Then WScript.Echo "Debug is enabled"
Dim VBSPM_TEST_INDEX: VBSPM_TEST_INDEX = 1
Dim vbspmDir: vbspmDir=Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
Dim baseDir
With CreateObject("WScript.Shell")
    baseDir=.CurrentDirectory
End With

Public Function startsWith(str, prefix)
    startsWith = Left(str, Len(prefix)) = prefix
End Function

Public Function endsWith(str, suffix)
    endsWith = Right(str, Len(suffix)) = suffix
End Function

Public Function contains(str, char)
    contains = (Instr(1, str, char) > 0)
End Function

Public Function argsArray()
    Dim i
    ReDim arr(WScript.Arguments.Count-1)
    For i = 0 To WScript.Arguments.Count-1
        arr(i) = """"+WScript.Arguments(i)+""""
    Next
    argsArray = arr
End Function

Public Function argsDict()
    Dim i, param, dict
    set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    ReDim arr(WScript.Arguments.Count-1)
    For i = 1 To WScript.Arguments.Count-1
        param = WScript.Arguments(i)
        If startsWith(param, "/") And contains(param, ":") Then
            param = mid(param, 2)
            WScript.Echo "param to be split: " & param
            dict.Add Lcase(split(param, ":")(0)), split(param, ":")(1)
        Else
            dict.Add i, param
        End If
    Next
    set argsDict = dict
End Function	

Class Console

	Public Function fmt( str, args )
		Dim res
		res = ""

		Dim pos
		pos = 0

		Dim i
		For i = 1 To Len(str)

			If Mid(str,i,1)="%" Then
				If i<Len(str) Then

					If Mid(str,i+1,1)="%" Then
						res = res & "%"
						i = i + 1

					ElseIf Mid(str,i+1,1)="x" Then
						res = res & CStr(args(pos))
						pos = pos+1
						i = i + 1
					End If
				End If

			Else
				res = res & Mid(str,i,1)
			End If
		Next

		fmt = res
	End Function

End Class



Dim oConsole                         
set oConsole = new Console
PUblic Sub printf(str, args)

    str = Replace(str, "%s", "%x")
    str = Replace(str, "%i", "%x")
    str = Replace(str, "%f", "%x")
    str = Replace(str, "%d", "%x")
    WScript.Echo oConsole.fmt(str, args)
End Sub

Public Sub debugf(str, args)
    if (debug) Then printf str, args
End Sub

Public Sub EchoX(str, args)
    If Not IsNull(args) Then
        If IsArray(args) Then

            WScript.Echo oConsole.fmt(str, args)
        Else

            WScript.Echo oConsole.fmt(str, Array(args))
        End if
    Else
        WScript.Echo str
    End If
End Sub

Public Sub Echo(str) 
    EchoX str, NULL
End Sub

Public Sub EchoDX(str, args)
    if (debug) Then EchoX str, args
End Sub

Public Sub EchoD(str) 
    EchoDX str, NULL
End Sub	

Class Collection

    Private dict
    Private oThis
    Private m_Name

    Private Sub Class_Initialize()
        set dict = CreateObject("Scripting.Dictionary")
        set oThis = Me
        m_Name = "Undefined"
    End Sub

    Public Default Property Get Obj
        set Obj = dict
    End Property 
    Public Property Set Obj(d)
        set dict = d
    End Property

    Public Property Get Name
        Name = m_Name
    End Property
    Public Property Let Name(Value)
        m_Name = Value
    End Property

    Public Sub Add(Key, Value)
        dict.Add key, value
    End Sub

    Public Sub Remove(Key)
        If KeyExists(Key) Then
            dict.Remove(Key)
        Else
            RaiseErr "Key [" & Key & "] does not exists in collection."
        End If
    End Sub

    Public Sub RemoveAll()
        dict.RemoveAll()
    End Sub

    Public Property Get Count
        Count = dict.Count
    End Property

    Public Function GetItem(Key)
        If KeyExists(Key) Then
            GetItem = dict.Item(Key)
        Else

            RaiseErr "Key [" & Key & "] does not exists in collection."
        End If
    End Function

    Public Function GetItemAtIndex(Index)

        GetItemAtIndex = dict.Item(Index)
    End Function

    Public Function IndexOf(Key)
        IndexOf = dict.IndexOf(Key, 0)
    End Function

    Public Function KeyExists(Key)
        KeyExists = dict.Exists(Key)
    End Function

    Public Function toCSV
        toCSV = join(toArray(), ", ")
    End Function

    Public Function toArray
        toArray = dict.Items
    End Function

    Public Function isEmpty
        isEmpty = (dict.Count = 0)        
    End Function

    Private Sub RaiseErr(desc)
        Err.Clear
        Err.Raise 1000, "Collection Class Error", desc
    End Sub

    Private Sub Class_Terminate()
        set dict = Nothing
        set oThis = Nothing
    End Sub

End Class



	Class DictUtil

    Function SortDictionary(objDict, intSort)

        Const dictKey  = 1
        Const dictItem = 2

        Dim strDict()
        Dim objKey
        Dim strKey,strItem
        Dim X,Y,Z

        Z = objDict.Count

        If Z > 1 Then

            ReDim strDict(Z,2)
            X = 0

            For Each objKey In objDict
                strDict(X,dictKey)  = CStr(objKey)
                strDict(X,dictItem) = CStr(objDict(objKey))
                X = X + 1
            Next

            For X = 0 To (Z - 2)
            For Y = X To (Z - 1)
                If StrComp(strDict(X,intSort),strDict(Y,intSort),vbTextCompare) > 0 Then
                    strKey  = strDict(X,dictKey)
                    strItem = strDict(X,dictItem)
                    strDict(X,dictKey)  = strDict(Y,dictKey)
                    strDict(X,dictItem) = strDict(Y,dictItem)
                    strDict(Y,dictKey)  = strKey
                    strDict(Y,dictItem) = strItem
                End If
            Next
            Next

            objDict.RemoveAll

            For X = 0 To (Z - 1)
            objDict.Add strDict(X,dictKey), strDict(X,dictItem)
            Next

        End If
    End Function
End Class



	Class ArrayUtil

	Public Function toString(arr)
		If Not isArray(arr) Then
			toString = "Supplied parameter is not an array."
			Exit Function
		End If

		Dim s, i
		s = "Array{" & UBound(arr) & "} [" & vbCrLf
		For i = 0  To UBound(arr)
			s = s & vbTab & "[" & i & "] => [" & arr(i) & "]"
			If i < UBound(arr) Then s = s & ", "
			s = s &  vbCrLf
		Next
		s = s & "]"
		toString = s

	End Function

	Public Function contains(arr, s) 
		If Not isArray(arr) Then
			toString = "Supplied parameter is not an array."
			Exit Function
		End If

		Dim i, bFlag
		bFlag = False
		For i = 0  To UBound(arr)
			If arr(i) = s Then
				bFlag = True
				Exit For
			End If
		Next
		contains = bFlag
	End Function

End Class



Dim arrUtil
set arrUtil = new ArrayUtil	

Class PathUtil

	Private Property Get DOT
	DOT = "."
	End Property
	Private Property Get DOTDOT
	DOTDOT = ".."
	End Property

	Private oFSO
	Private m_base
	Private m_script
	Private m_temp

	Private Sub Class_Initialize()
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		m_script = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\")-1)
		m_base = m_script
		m_temp = Array()
		ReDim Preserve m_temp(0)
		m_temp(0) = m_script
	End Sub

	Public Property Get ScriptPath
	ScriptPath = m_script
	End Property

	Public Property Get BasePath
	BasePath = m_base
	End Property

	Public Property Let BasePath(path)
	Do While endsWith(path, "\")
		path = Left(Path, Len(path)-1)
	Loop
	m_base = Resolve(path)
	EchoDX "New Base Path: %x", m_base
	End Property

	Public Property Get TempBasePath
	TempBasePath = m_temp(UBound(m_temp))
	End Property

	Public Property Let TempBasePath(path)
	Do While endsWith(path, "\")
		path = Left(Path, Len(path)-1)
	Loop
	If arrUtil.contains(m_temp, path) Then
		EchoDX "Temp Path %x already exists; skipped", path
	Else
		ReDim Preserve m_temp(Ubound(m_temp)+1)
		m_temp(Ubound(m_temp)) = Resolve(path)
		EchoDX "New Temp Base Path: %x", m_temp(Ubound(m_temp))
	End If
	End Property

	Function Resolve(path)
		Dim pathBase, lPath, final
		EchoDX "path: %x", path
		If path = DOT Or path = DOTDOT Then
			path = path & "\"
		End If
		EchoDX "path: %x", path

		If oFSO.FolderExists(path) Then
			EchoD "FolderExists"
			Resolve = oFSO.GetFolder(path).path
			Exit Function
		End If

		If oFSO.FileExists(path) Then
			EchoD "FileExists"
			Resolve = oFSO.GetFile(path).path
			Exit Function
		End If

		pathBase = oFSO.BuildPath(m_base, path)
		EchoDX "Adding base %x to path %x. New Path: %x", Array(m_base, path, pathBase)

		If endsWith(pathBase, "\") Then
			If isObject(oFSO.GetFolder(pathBase)) Then
				EchoD "EndsWith '\' -> FolderExists"
				Resolve = oFSO.GetFolder(pathBase).Path
				Exit Function
			End If
		Else

			If oFSO.FolderExists(pathBase) Then
				EchoD "FolderExists"
				Resolve = oFSO.GetFolder(pathBase).path
				Exit Function
			End If

			If oFSO.FileExists(pathBase) Then
				EchoD "FileExists"
				Resolve = oFSO.GetFile(pathBase).path
				Exit Function
			End If

			Dim i
			i = Ubound(m_temp)
			Do
				lPath = oFSO.BuildPath(m_temp(i), path)
				EchoDX "Adding Temp Base path (%x) %x to path %x. New Path: %x", Array(i, m_temp(i), path, lPath)
				If oFSO.FileExists(lPath) Then
					final = oFSO.GetFile(lPath).path
					EchoDX "File Resolved with Temp Base %x", final
					Resolve = final
					Exit Function
				End If
				If oFSO.FolderExists(lPath) Then
					final = oFSO.GetFolder(lPath)
					EchoDX "Folder Resolved with Temp Base %x", final
					Resolve = final
					Exit Function
				End If
				i = i - 1
			Loop While i >= 0

			lPath = oFSO.BuildPath(m_script, path)
			EchoDX "Adding script path %x to path %x. New Path: %x", Array(m_script, path, lPath)
			If oFSO.FileExists(lPath) Then
				final = oFSO.GetFile(lPath).path
				EchoDX "File Resolved with Temp Base %x", final
				Resolve = final
				Exit Function
			End If
			If oFSO.FolderExists(lPath) Then
				final = oFSO.GetFolder(lPath)
				EchoDX "Folder Resolved with Temp Base %x", final
				Resolve = final
				Exit Function
			End If
		End If

		EchoD "Unable to Resolve"
		Resolve = path
	End Function

	Private Sub Class_Terminate()
		Set oFSO = Nothing
	End Sub

End Class



Dim putil
set putil = new PathUtil
putil.BasePath = baseDir
EchoX "Project location: %x", putil.BasePath	

Class FSO
	Private dir
	Private objFSO

	Private Sub Class_Initialize
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		dir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
	End Sub

	Public Sub setDir(s)
		dir = s
	End Sub

	Public Function getDir
		getDir = dir
	End Function

	Public Function GetFSO
		Set GetFSO = objFSO
	End Function

	Public Function FolderExists(fol)
		FolderExists = objFSO.FolderExists(fol)
	End Function

	Public Function CreateFolder(fol)
		CreateFolder = False
		If FolderExists(fol) Then
			CreateFolder = True
		Else
			objFSO.CreateFolder(fol)
			CreateFolder = FolderExists(fol)
		End If
	End Function

	Public Sub WriteFile(strFileName, strMessage, overwrite)
		Const ForReading = 1
		Const ForWriting = 2
		Const ForAppending = 8
		Dim mode
		Dim oFile

		mode = ForWriting
		If Not overwrite Then
			mode = ForAppending
		End If

		If objFSO.FileExists(strFileName) Then
			Set oFile = objFSO.OpenTextFile(strFileName, mode)
		Else
			Set oFile = objFSO.CreateTextFile(strFileName)
		End If
		oFile.WriteLine strMessage

		oFile.Close

		Set oFile = Nothing
	End Sub

	Public Function GetFileDir(ByVal file)
		EchoDX "GetFileDir( %x )", Array(file)
		Dim objFile
		Set objFile = objFSO.GetFile(file)
		GetFileDir = objFSO.GetParentFolderName(objFile) 
	End Function

	Public Function GetFilePath(ByVal file)
		EchoDX "GetFilePath( %x )", Array(file)
		Dim objFile
		On Error Resume Next
		Set objFile = objFSO.GetFile(file)
		On Error GoTo 0
		If IsObject(objFile) Then
			GetFilePath = objFile.Path 
		Else
			EchoDX "File %x not found; searching in directory %x", Array(file,dir)
			On Error Resume Next
			Set objFile = objFile.GetFile(objFSO.BuildPath(dir, file))
			On Error GoTo 0
			If IsObject(objFile) Then
				GetFilePath = objFile.Path 
			Else
				GetFilePath = "File [" & file & "] Not found"
			End If
		End If
	End Function

	Public Function GetFileName(ByVal file)
		GetFileName = objFSO.GetFile(file).Name
	End Function

	Public Function GetFileExtn(file)
		GetFileExtn = ""
		On Error Resume Next
		GetFileExtn = LCASE(objFSO.GetExtensionName(file))
		On Error GoTo 0
	End Function

	Public Function GetBaseName(ByVal file)
		GetBaseName = Replace(GetFileName(file), "." & GetFileExtn(file), "")
	End Function

	Public Function ReadFile(file)
		file = putil.Resolve(file)
		EchoDX "---> File resolved to: %x", Array(file)
		If Not FileExists(file) Then 
			Wscript.Echo "---> File " & file & " does not exists."
			ReadFile = ""
			Exit Function
		End If
		Dim objFile: Set objFile = objFSO.OpenTextFile(file)
		ReadFile = objFile.ReadAll()
		objFile.Close
	End Function

	Public Function FileExists(file)
		FileExists = objFSO.FileExists(file)
	End Function

	Public Sub DeleteFile(file)
		On Error Resume Next
		objFSO.DeleteFile(file)
		On Error GoTo 0
	End Sub

End Class



Dim cFS
set cFS = new FSO

cFS.setDir(baseDir)

Public Function log(msg)
cFS.WriteFile "build.log", msg, false
End Function

log "VBSPM Directory: " & vbspmDir	

Class ClassA
    public default sub CallMe
        WScript.Echo "I'm in ClassA"
    End Sub
End Class



	Class ClassB

    Private m_CLASSA

    Private Sub Class_Initialize
        set m_CLASSA = new CLASSA
    End Sub

    public default sub CallMe
        call m_CLASSA.CallMe
    End Sub
End Class



Dim ccb 
set ccb = new ClassB
ccb.CallMe

Public Sub Include(file)

End Sub
Public Sub Import(file)

End Sub


'================= File: C:\Users\nanda\Github\sysadmin-accelerator\VBNext\src\utils\Version.vbs =================
Class Version
    
    Private lngPrivateMajor
    Private lngPrivateMinor
    Private lngPrivateBuild
    Private lngPrivateRevision
    
    Private Sub Class_Initialize()
        lngPrivateMajor = CLng(0)
        lngPrivateMinor = CLng(0)
        lngPrivateBuild = CLng(-1)
        lngPrivateRevision = CLng(-1)
    End Sub

    Private Function TestObjectForData(ByVal objToCheck)
        
        Dim boolTestResult
        Dim boolFunctionReturn
        Dim intArrayUBound
    
        Err.Clear
    
        boolFunctionReturn = True
    
        'Check VarType(objToCheck) = 0
        On Error Resume Next
        boolTestResult = (VarType(objToCheck) = 0)
        If Err Then
            'Error occurred
            Err.Clear
            On Error Goto 0
        Else
            'No Error
            On Error Goto 0
            If boolTestResult = True Then
                'vbEmpty
                boolFunctionReturn = False
            End If
        End If
    
        'Check VarType(objToCheck) = 1
        On Error Resume Next
        boolTestResult = (VarType(objToCheck) = 1)
        If Err Then
            'Error occurred
            Err.Clear
            On Error Goto 0
        Else
            'No Error
            On Error Goto 0
            If boolTestResult = True Then
                'vbNull
                boolFunctionReturn = False
            End If
        End If
    
        'Check to see if objToCheck Is Nothing
        If boolFunctionReturn = True Then
            On Error Resume Next
            boolTestResult = (objToCheck Is Nothing)
            If Err Then
                'Error occurred
                Err.Clear
                On Error Goto 0
            Else
                'No Error
                On Error Goto 0
                If boolTestResult = True Then
                    'No data
                    boolFunctionReturn = False
                End If
            End If
        End If
    
        'Check IsEmpty(objToCheck)
        If boolFunctionReturn = True Then
            On Error Resume Next
            boolTestResult = IsEmpty(objToCheck)
            If Err Then
                'Error occurred
                Err.Clear
                On Error Goto 0
            Else
                'No Error
                On Error Goto 0
                If boolTestResult = True Then
                    'No data
                    boolFunctionReturn = False
                End If
            End If
        End If
    
        'Check IsNull(objToCheck)
        If boolFunctionReturn = True Then
            On Error Resume Next
            boolTestResult = IsNull(objToCheck)
            If Err Then
                'Error occurred
                Err.Clear
                On Error Goto 0
            Else
                'No Error
                On Error Goto 0
                If boolTestResult = True Then
                    'No data
                    boolFunctionReturn = False
                End If
            End If
        End If
        
        'Check objToCheck = vbNullString
        If boolFunctionReturn = True Then
            On Error Resume Next
            boolTestResult = (objToCheck = vbNullString)
            If Err Then
                'Error occurred
                Err.Clear
                On Error Goto 0
            Else
                'No Error
                On Error Goto 0
                If boolTestResult = True Then
                    'No data
                    boolFunctionReturn = False
                End If
            End If
        End If
    
        If boolFunctionReturn = True Then
            On Error Resume Next
            boolTestResult = IsArray(objToCheck)
            If Err Then
                'Error occurred
                Err.Clear
                On Error Goto 0
                boolTestResult = False
            Else
                'No Error
                On Error Goto 0
            End If
            If boolTestResult = True Then
                ' objToCheck is an array
                On Error Resume Next
                intArrayUBound = UBound(objToCheck)
                If Err Then
                    'Undimensioned array
                    Err.Clear
                    On Error Goto 0
                    intArrayUBound = -1
                Else
                    On Error Goto 0
                End If
                If intArrayUBound < 0 Then
                    boolFunctionReturn = False
                End If
            End If
        End If
    
        TestObjectForData = boolFunctionReturn
    End Function        

    Public Function Clone(ByRef objTargetVersionObject)
        ' Creates a copy of the current version object and stores it in the first (and only)
        ' argument supplied to this function

        ' Returns 0 if successful; non-zero otherwise

        Dim intReturnCode
        intReturnCode = 0
        Set objTargetVersionObject = New Version
        If lngPrivateRevision = CLng(-1) Then
            If lngPrivateBuild = CLng(-1) Then
                ' Initialize with Major/Minor only
                intReturnCode = objTargetVersionObject.InitFromMajorMinor(lngPrivateMajor, lngPrivateMinor)
            Else
                ' Initialize with Major/Minor/Build only
                intReturnCode = objTargetVersionObject.InitFromMajorMinorBuild(lngPrivateMajor, lngPrivateMinor, lngPrivateBuild)
            End If
        Else
            ' Initialize with Major/Minor/Build/Revision
            intReturnCode = objTargetVersionObject.InitFromMajorMinorBuildRevision(lngPrivateMajor, lngPrivateMinor, lngPrivateBuild, lngPrivateRevision)
        End If

        Clone = intReturnCode
    End Function

    Public Function CompareTo(ByVal objOtherVersionObject)
        ' Compares this version object to the version object supplied as an argument.
        ' Returns 1 if this version object is subsequent/later than the version supplied in
        '   the argument. Also returns 1 if the object supplied as an argument was null/
        '   nothing, or if the object supplied as an argument was not a valid version object
        ' Returns 0 if this version object is equal to the version supplied in the argument
        ' Returns -1 if this version object is before/earlier than the version supplied in the
        '   argument.
        Dim intResult
        Dim lngComparedMajor
        Dim lngComparedMinor
        Dim lngComparedBuild
        Dim lngComparedRevision

        Err.Clear

        intResult = 0
        If TestObjectForData(objOtherVersionObject) = False Then
            intResult = 1
        Else
            On Error Resume Next
            lngComparedMajor = CLng(objOtherVersionObject.Major)
            If Err Then
                Err.Clear
                On Error Goto 0
                intResult = 1
            Else
                lngComparedMinor = CLng(objOtherVersionObject.Minor)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    intResult = 1
                Else
                    lngComparedBuild = CLng(objOtherVersionObject.Build)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        intResult = 1
                    Else
                        lngComparedRevision = CLng(objOtherVersionObject.Revision)
                        If Err Then
                            Err.Clear
                            On Error Goto 0
                            intResult = 1
                        Else
                            On Error Goto 0
                        End If
                    End If
                End If
            End If
        End If

        If intResult = 0 Then
            If lngPrivateMajor <> lngComparedMajor Then
                If lngPrivateMajor < lngComparedMajor Then
                    intResult = -1
                Else
                    intResult = 1
                End If
            ElseIf lngPrivateMinor <> lngComparedMinor Then
                If lngPrivateMinor < lngComparedMinor Then
                    intResult = -1
                Else
                    intResult = 1
                End If
            ElseIf lngPrivateBuild <> lngComparedBuild Then
                If lngPrivateBuild < lngComparedBuild Then
                    intResult = -1
                Else
                    intResult = 1
                End If
            ElseIf lngPrivateRevision <> lngComparedRevision Then
                If lngPrivateRevision < lngComparedRevision Then
                    intResult = -1
                Else
                    intResult = 1
                End If
            End If
        End If
        CompareTo = intResult
    End Function

    Public Function CompareToString(ByVal strOtherVersion)
        ' Compares this version object to the string representation of a version number
        ' supplied as an argument.
        ' Returns 1 if this version object is subsequent/later than the version supplied in
        '   the argument. Also returns 1 if the string supplied as an argument was null/
        '   nothing/empty string, or if the string supplied as an argument was not a valid
        '   version object
        ' Returns 0 if this version object is equal to the version supplied in the argument
        ' Returns -1 if this version object is before/earlier than the version supplied in the
        '   argument.
        Dim objOtherVersionObject
        Dim intReturnCode
        Dim intResult
        Dim lngComparedMajor
        Dim lngComparedMinor
        Dim lngComparedBuild
        Dim lngComparedRevision

        Err.Clear

        intResult = 0
        If TestObjectForData(strOtherVersion) = False Then
            intResult = 1
        Else
            Set objOtherVersionObject = New Version
            intReturnCode = objOtherVersionObject.InitFromString(strOtherVersion)
            If intReturnCode <> 0 Then
                intResult = 1
            End If
        End If

        If intResult = 0 Then
            If TestObjectForData(objOtherVersionObject) = False Then
                intResult = 1
            Else
                On Error Resume Next
                lngComparedMajor = CLng(objOtherVersionObject.Major)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    intResult = 1
                Else
                    lngComparedMinor = CLng(objOtherVersionObject.Minor)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        intResult = 1
                    Else
                        lngComparedBuild = CLng(objOtherVersionObject.Build)
                        If Err Then
                            Err.Clear
                            On Error Goto 0
                            intResult = 1
                        Else
                            lngComparedRevision = CLng(objOtherVersionObject.Revision)
                            If Err Then
                                Err.Clear
                                On Error Goto 0
                                intResult = 1
                            Else
                                On Error Goto 0
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If intResult = 0 Then
            If lngPrivateMajor <> lngComparedMajor Then
                If lngPrivateMajor < lngComparedMajor Then
                    intResult = -1
                Else
                    intResult = 1
                End If
            ElseIf lngPrivateMinor <> lngComparedMinor Then
                If lngPrivateMinor < lngComparedMinor Then
                    intResult = -1
                Else
                    intResult = 1
                End If
            ElseIf lngPrivateBuild <> lngComparedBuild Then
                If lngPrivateBuild < lngComparedBuild Then
                    intResult = -1
                Else
                    intResult = 1
                End If
            ElseIf lngPrivateRevision <> lngComparedRevision Then
                If lngPrivateRevision < lngComparedRevision Then
                    intResult = -1
                Else
                    intResult = 1
                End If
            End If
        End If
        CompareToString = intResult
    End Function

    Public Function Equals(ByVal objOtherVersionObject)
        ' Compares the current version to the version object supplied as an argument.
        ' Returns True if the two versions are equal
        ' Returns False otherwise. Also returns False if the object supplied as an argument is
        '   not a valid Version object
        Dim boolResult
        Dim lngComparedMajor
        Dim lngComparedMinor
        Dim lngComparedBuild
        Dim lngComparedRevision

        Err.Clear

        boolResult = True
        If TestObjectForData(objOtherVersionObject) = False Then
            boolResult = False
        Else
            On Error Resume Next
            lngComparedMajor = CLng(objOtherVersionObject.Major)
            If Err Then
                Err.Clear
                On Error Goto 0
                boolResult = False
            Else
                lngComparedMinor = CLng(objOtherVersionObject.Minor)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    boolResult = False
                Else
                    lngComparedBuild = CLng(objOtherVersionObject.Build)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        boolResult = False
                    Else
                        lngComparedRevision = CLng(objOtherVersionObject.Revision)
                        If Err Then
                            Err.Clear
                            On Error Goto 0
                            boolResult = False
                        Else
                            On Error Goto 0
                        End If
                    End If
                End If
            End If
        End If

        If boolResult = True Then
            If lngPrivateMajor <> lngComparedMajor Then
                boolResult = False
            ElseIf lngPrivateMinor <> lngComparedMinor Then
                boolResult = False
            ElseIf lngPrivateBuild <> lngComparedBuild Then
                boolResult = False
            ElseIf lngPrivateRevision <> lngComparedRevision Then
                boolResult = False
            End If
        End If
        Equals = boolResult
    End Function

    Public Function GreaterThan(ByVal objOtherVersionObject)
        ' Compares the current version to the version object supplied as an argument.
        ' Returns True if this version object is greater than the version supplied as an
        '   argument
        ' Returns False otherwise. Also returns False if the object supplied as an argument is
        '   not a valid Version object
        Dim boolResult
        Dim lngComparedMajor
        Dim lngComparedMinor
        Dim lngComparedBuild
        Dim lngComparedRevision

        Err.Clear

        boolResult = True
        If TestObjectForData(objOtherVersionObject) = False Then
            boolResult = False
        Else
            On Error Resume Next
            lngComparedMajor = CLng(objOtherVersionObject.Major)
            If Err Then
                Err.Clear
                On Error Goto 0
                boolResult = False
            Else
                lngComparedMinor = CLng(objOtherVersionObject.Minor)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    boolResult = False
                Else
                    lngComparedBuild = CLng(objOtherVersionObject.Build)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        boolResult = False
                    Else
                        lngComparedRevision = CLng(objOtherVersionObject.Revision)
                        If Err Then
                            Err.Clear
                            On Error Goto 0
                            boolResult = False
                        Else
                            On Error Goto 0
                        End If
                    End If
                End If
            End If
        End If

        If boolResult = True Then
            If lngPrivateMajor < lngComparedMajor Then
                boolResult = False
            ElseIf lngPrivateMajor = lngComparedMajor Then
                If lngPrivateMinor < lngComparedMinor Then
                    boolResult = False
                ElseIf lngPrivateMinor = lngComparedMinor Then
                    If lngPrivateBuild < lngComparedBuild Then
                        boolResult = False
                    ElseIf lngPrivateBuild = lngComparedBuild Then
                        If lngPrivateRevision <= lngComparedRevision Then
                            boolResult = False
                        End If
                    End If
                End If
            End If
        End If
        GreaterThan = boolResult
    End Function

    Public Function GreaterThanOrEqual(ByVal objOtherVersionObject)
        ' Compares the current version to the version object supplied as an argument.
        ' Returns True if this version object is greater than or equal to the version supplied
        '   as an argument
        ' Returns False otherwise. Also returns False if the object supplied as an argument is
        '   not a valid Version object
        Dim boolResult
        Dim lngComparedMajor
        Dim lngComparedMinor
        Dim lngComparedBuild
        Dim lngComparedRevision

        Err.Clear

        boolResult = True
        If TestObjectForData(objOtherVersionObject) = False Then
            boolResult = False
        Else
            On Error Resume Next
            lngComparedMajor = CLng(objOtherVersionObject.Major)
            If Err Then
                Err.Clear
                On Error Goto 0
                boolResult = False
            Else
                lngComparedMinor = CLng(objOtherVersionObject.Minor)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    boolResult = False
                Else
                    lngComparedBuild = CLng(objOtherVersionObject.Build)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        boolResult = False
                    Else
                        lngComparedRevision = CLng(objOtherVersionObject.Revision)
                        If Err Then
                            Err.Clear
                            On Error Goto 0
                            boolResult = False
                        Else
                            On Error Goto 0
                        End If
                    End If
                End If
            End If
        End If

        If boolResult = True Then
            If lngPrivateMajor < lngComparedMajor Then
                boolResult = False
            ElseIf lngPrivateMinor < lngComparedMinor Then
                boolResult = False
            ElseIf lngPrivateBuild < lngComparedBuild Then
                boolResult = False
            ElseIf lngPrivateRevision < lngComparedRevision Then
                boolResult = False
            End If
        End If
        GreaterThanOrEqual = boolResult
    End Function

    Public Function InitFromMajorMinor(ByVal lngMajor, ByVal lngMinor)
        ' Initalizes a Version object from a pair of long integers supplied by two arguments.
        ' The first argument is the major version number
        ' The second argument is the minor version number
        ' For example: major.minor
        ' This method returns 0 if successful; non-zero otherwise.
        Dim intFunctionReturn

        intFunctionReturn = InitFromMajorMinorBuildRevision(lngMajor, lngMinor, 0, 0)

        If intFunctionReturn = 0 Then
            lngPrivateBuild = CLng(-1)
            lngPrivateRevision = CLng(-1)
        End If

        InitFromMajorMinor = intFunctionReturn
    End Function

    Public Function InitFromMajorMinorBuild(ByVal lngMajor, ByVal lngMinor, ByVal lngBuild)
        ' Initalizes a Version object from three long integers supplied by three arguments.
        ' The first argument is the major version number
        ' The second argument is the minor version number
        ' The third argument is the build number
        ' For example: major.minor.build
        ' This method returns 0 if successful; non-zero otherwise.
        Dim intFunctionReturn

        intFunctionReturn = InitFromMajorMinorBuildRevision(lngMajor, lngMinor, lngBuild, 0)

        If intFunctionReturn = 0 Then
            lngPrivateRevision = CLng(-1)
        End If

        InitFromMajorMinorBuild = intFunctionReturn
    End Function

    Public Function InitFromMajorMinorBuildRevision(ByVal lngMajor, ByVal lngMinor, ByVal lngBuild, ByVal lngRevision)
        ' Initalizes a Version object from four long integers supplied by four arguments.
        ' The first argument is the major version number
        ' The second argument is the minor version number
        ' The third argument is the build number
        ' The fourth argument is the revision number
        ' For example: major.minor.build.revision
        ' This method returns 0 if successful; non-zero otherwise.
        Dim intFunctionReturn
        Dim lngTempMajor
        Dim lngTempMinor
        Dim lngTempBuild
        Dim lngTempRevision

        Err.Clear

        intFunctionReturn = 0

        If TestObjectForData(lngMajor) = False Then
            ' Blank sections of the version number are allowed here
            lngTempMajor = CLng(0)
        Else
            On Error Resume Next
            lngTempMajor = CLng(lngMajor)
            If Err Then
                Err.Clear
                On Error Goto 0
                ' The "major" portion of the version number was not a valid long integer
                intFunctionReturn = -1
            Else
                On Error Goto 0
                If lngTempMajor < CLng(0) Then
                    ' Cannot have negative version numbers
                    intFunctionReturn = -2
                Else
                    lngTempMinor = CLng(0)
                    lngTempBuild = CLng(0)
                    lngTempRevision = CLng(0)
                End If
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            If TestObjectForData(lngMinor) = False Then
                ' Blank sections of the version number are allowed here
                ' Already set; nothing more to do
            Else
                On Error Resume Next
                lngTempMinor = CLng(lngMinor)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    ' The "minor" portion of the version number was not a valid long integer
                    intFunctionReturn = -3
                Else
                    On Error Goto 0
                    If lngTempMinor < CLng(0) Then
                        ' Cannot have negative version numbers
                        intFunctionReturn = -4
                    End If
                End If
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            If TestObjectForData(lngBuild) = False Then
                ' Blank sections of the version number are allowed here
                ' Already set; nothing more to do
            Else
                On Error Resume Next
                lngTempBuild = CLng(lngBuild)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    ' The "build" portion of the version number was not a valid long integer
                    intFunctionReturn = -5
                Else
                    On Error Goto 0
                    If lngTempBuild < CLng(0) Then
                        ' Cannot have negative version numbers
                        intFunctionReturn = -6
                    End If
                End If
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            If TestObjectForData(lngRevision) = False Then
                ' Blank sections of the version number are allowed here
                ' Already set; nothing more to do
            Else
                On Error Resume Next
                lngTempRevision = CLng(lngRevision)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    ' The "Revision" portion of the version number was not a valid long integer
                    intFunctionReturn = -7
                Else
                    On Error Goto 0
                    If lngTempRevision < CLng(0) Then
                        ' Cannot have negative version numbers
                        intFunctionReturn = -8
                    End If
                End If
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            lngPrivateMajor = lngTempMajor
            lngPrivateMinor = lngTempMinor
            lngPrivateBuild = lngTempBuild
            lngPrivateRevision = lngTempRevision
        End If

        InitFromMajorMinorBuildRevision = intFunctionReturn
    End Function

    Public Default Function InitFromString(ByVal strVersion)
        ' Initalizes a Version object from a version-formatted string supplied as an argument.
        ' Valid strings look like the following:
        ' major.minor
        ' major.minor.build
        ' or
        ' major.minor.build.revision
        ' Each part of the version string must be in decimal and convertable to a long integer
        ' This method returns 0 if successful; non-zero otherwise.
        Dim intFunctionReturn
        Dim arrVersion
        Dim intCountOfVersionSections
        Dim boolVersionSectionCountTest
        Dim lngTempMajor
        Dim lngTempMinor
        Dim lngTempBuild
        Dim lngTempRevision

        Err.Clear

        intFunctionReturn = 0

        If TestObjectForData(strVersion) = False Then
            ' No data was passed to function
            intFunctionReturn = -1
        Else
            On Error Resume Next
            arrVersion = Split(strVersion, ".")
            If Err Then
                Err.Clear
                On Error Goto 0
                ' Object passed to function was not a string, or an error occurred splitting
                ' the string
                intFunctionReturn = -2
            Else
                intCountOfVersionSections = UBound(arrVersion)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    ' Something went wrong reading the upper boundary of the array resulting
                    ' from the Split() function
                    intFunctionReturn = -3
                Else
                    boolVersionSectionCountTest = (intCountOfVersionSections > 3) Or (intCountOfVersionSections < 1)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        ' Something went wrong comparing the upper boundary to an interger
                        intFunctionReturn = -4
                    Else
                        On Error Goto 0
                    End If
                End If
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            If boolVersionSectionCountTest = True Then
                ' Less than two parts of the version string were passed (e.g., "1")
                ' or
                ' More than four parts of the version string were passed (e.g., "1.2.3.4.5")
                ' Neither is allowed here, nor the System.Version .NET analog
                intFunctionReturn = -5
            Else
                ' String appears valid so far and has 2-4 parts, e.g.:
                ' 1.2
                ' 1.2.3
                ' 1.2.3.4
                If TestObjectForData(arrVersion(0)) = False Then
                    ' Blank sections of the version number are not allowed during conversion
                    ' from string
                    intFunctionReturn = -6
                Else
                    On Error Resume Next
                    lngTempMajor = CLng(arrVersion(0))
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        ' The "major" portion of the version number was not a valid long
                        ' integer
                        intFunctionReturn = -7
                    Else
                        On Error Goto 0
                        If lngTempMajor < CLng(0) Then
                            ' Cannot have negative version numbers
                            intFunctionReturn = -8
                        Else
                            lngTempMinor = CLng(0)
                            lngTempBuild = CLng(0)
                            lngTempRevision = CLng(0)
                        End If
                    End If
                End If
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            If TestObjectForData(arrVersion(1)) = False Then
                ' Blank sections of the version number are not allowed during conversion
                ' from string
                intFunctionReturn = -9
            Else
                On Error Resume Next
                lngTempMinor = CLng(arrVersion(1))
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    ' The "minor" portion of the version number was not a valid long integer
                    intFunctionReturn = -10
                Else
                    On Error Goto 0
                    If lngTempMinor < CLng(0) Then
                        ' Cannot have negative version numbers
                        intFunctionReturn = -11
                    End If
                End If
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            If intCountOfVersionSections >= 2 Then
                ' Build portion of version should be present
                If TestObjectForData(arrVersion(2)) = False Then
                    ' Blank sections of the version number are not allowed during conversion
                    ' from string
                    intFunctionReturn = -12
                Else
                    On Error Resume Next
                    lngTempBuild = CLng(arrVersion(2))
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        ' The "build" portion of the version number was not a valid long integer
                        intFunctionReturn = -13
                    Else
                        On Error Goto 0
                        If lngTempBuild < CLng(0) Then
                            ' Cannot have negative version numbers
                            intFunctionReturn = -14
                        End If
                    End If
                End If
            Else
                lngTempBuild = CLng(-1)
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            If intCountOfVersionSections = 3 Then
                ' Revision portion of version should be present
                If TestObjectForData(arrVersion(3)) = False Then
                    ' Blank sections of the version number are not allowed during conversion
                    ' from string
                    intFunctionReturn = -15
                Else
                    On Error Resume Next
                    lngTempRevision = CLng(arrVersion(3))
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        ' The "revision" portion of the version number was not a valid long integer
                        intFunctionReturn = -16
                    Else
                        On Error Goto 0
                        If lngTempRevision < CLng(0) Then
                            ' Cannot have negative version numbers
                            intFunctionReturn = -17
                        End If
                    End If
                End If
            Else
                lngTempRevision = CLng(-1)
            End If
        End If

        If intFunctionReturn = 0 Then
            ' No error occurred
            lngPrivateMajor = lngTempMajor
            lngPrivateMinor = lngTempMinor
            lngPrivateBuild = lngTempBuild
            lngPrivateRevision = lngTempRevision
        End If

        InitFromString = intFunctionReturn
    End Function

    Public Function LessThan(ByVal objOtherVersionObject)
        ' Compares the current version to the version object supplied as an argument.
        ' Returns True if this version object is less than the version supplied as an argument
        ' Returns False otherwise. Also returns False if the object supplied as an argument is
        '   not a valid Version object
        Dim boolResult
        Dim lngComparedMajor
        Dim lngComparedMinor
        Dim lngComparedBuild
        Dim lngComparedRevision

        Err.Clear

        boolResult = True
        If TestObjectForData(objOtherVersionObject) = False Then
            boolResult = False
        Else
            On Error Resume Next
            lngComparedMajor = CLng(objOtherVersionObject.Major)
            If Err Then
                Err.Clear
                On Error Goto 0
                boolResult = False
            Else
                lngComparedMinor = CLng(objOtherVersionObject.Minor)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    boolResult = False
                Else
                    lngComparedBuild = CLng(objOtherVersionObject.Build)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        boolResult = False
                    Else
                        lngComparedRevision = CLng(objOtherVersionObject.Revision)
                        If Err Then
                            Err.Clear
                            On Error Goto 0
                            boolResult = False
                        Else
                            On Error Goto 0
                        End If
                    End If
                End If
            End If
        End If

        If boolResult = True Then
            If lngPrivateMajor > lngComparedMajor Then
                boolResult = False
            ElseIf lngPrivateMajor = lngComparedMajor Then
                If lngPrivateMinor > lngComparedMinor Then
                    boolResult = False
                ElseIf lngPrivateMinor = lngComparedMinor Then
                    If lngPrivateBuild > lngComparedBuild Then
                        boolResult = False
                    ElseIf lngPrivateBuild = lngComparedBuild Then
                        If lngPrivateRevision >= lngComparedRevision Then
                            boolResult = False
                        End If
                    End If
                End If
            End If
        End If
        LessThan = boolResult
    End Function

    Public Function LessThanOrEqual(ByVal objOtherVersionObject)
        ' Compares the current version to the version object supplied as an argument.
        ' Returns True if this version object is less than or equal to the version supplied as
        '   an argument
        ' Returns False otherwise. Also returns False if the object supplied as an argument is
        '   not a valid Version object
        Dim boolResult
        Dim lngComparedMajor
        Dim lngComparedMinor
        Dim lngComparedBuild
        Dim lngComparedRevision

        Err.Clear

        boolResult = True
        If TestObjectForData(objOtherVersionObject) = False Then
            boolResult = False
        Else
            On Error Resume Next
            lngComparedMajor = CLng(objOtherVersionObject.Major)
            If Err Then
                Err.Clear
                On Error Goto 0
                boolResult = False
            Else
                lngComparedMinor = CLng(objOtherVersionObject.Minor)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    boolResult = False
                Else
                    lngComparedBuild = CLng(objOtherVersionObject.Build)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        boolResult = False
                    Else
                        lngComparedRevision = CLng(objOtherVersionObject.Revision)
                        If Err Then
                            Err.Clear
                            On Error Goto 0
                            boolResult = False
                        Else
                            On Error Goto 0
                        End If
                    End If
                End If
            End If
        End If

        If boolResult = True Then
            If lngPrivateMajor > lngComparedMajor Then
                boolResult = False
            ElseIf lngPrivateMinor > lngComparedMinor Then
                boolResult = False
            ElseIf lngPrivateBuild > lngComparedBuild Then
                boolResult = False
            ElseIf lngPrivateRevision > lngComparedRevision Then
                boolResult = False
            End If
        End If
        LessThanOrEqual = boolResult
    End Function

    Public Function NotEquals(ByVal objOtherVersionObject)
        ' Compares the current version to the version object supplied as an argument.
        ' Returns True if the two versions are not equal. Also returns True if the object
        '   supplied as an argument is not a valid Version object
        ' Returns False otherwise. 
        Dim boolResult
        Dim lngComparedMajor
        Dim lngComparedMinor
        Dim lngComparedBuild
        Dim lngComparedRevision

        Err.Clear

        boolResult = False
        If TestObjectForData(objOtherVersionObject) = False Then
            boolResult = True
        Else
            On Error Resume Next
            lngComparedMajor = CLng(objOtherVersionObject.Major)
            If Err Then
                Err.Clear
                On Error Goto 0
                boolResult = True
            Else
                lngComparedMinor = CLng(objOtherVersionObject.Minor)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    boolResult = True
                Else
                    lngComparedBuild = CLng(objOtherVersionObject.Build)
                    If Err Then
                        Err.Clear
                        On Error Goto 0
                        boolResult = True
                    Else
                        lngComparedRevision = CLng(objOtherVersionObject.Revision)
                        If Err Then
                            Err.Clear
                            On Error Goto 0
                            boolResult = True
                        Else
                            On Error Goto 0
                        End If
                    End If
                End If
            End If
        End If

        If boolResult = False Then
            If lngPrivateMajor <> lngComparedMajor Then
                boolResult = True
            ElseIf lngPrivateMinor <> lngComparedMinor Then
                boolResult = True
            ElseIf lngPrivateBuild <> lngComparedBuild Then
                boolResult = True
            ElseIf lngPrivateRevision <> lngComparedRevision Then
                boolResult = True
            End If
        End If
        NotEquals = boolResult
    End Function

    Public Function ToString()
        ' Returns a dot-separated representation of the version number as a string.
        ' Valid strings look like the following:
        ' major.minor
        ' major.minor.build
        ' or
        ' major.minor.build.revision
        ' (where each part of the string is a long integer converted to string format)
        Dim strToReturn
        If lngPrivateRevision = CLng(-1) Then
            If lngPrivateBuild = CLng(-1) Then
                ' Output Major/Minor only
                strToReturn = CStr(lngPrivateMajor) & "." & CStr(lngPrivateMinor)
            Else
                ' Output Major/Minor/Build only
                strToReturn = CStr(lngPrivateMajor) & "." & CStr(lngPrivateMinor) & "." & CStr(lngPrivateBuild)
            End If
        Else
            ' Output Major/Minor/Build/Revision
            strToReturn = CStr(lngPrivateMajor) & "." & CStr(lngPrivateMinor) & "." & CStr(lngPrivateBuild) & "." & CStr(lngPrivateRevision)
        End If
        ToString = strToReturn
    End Function

    Public Property Get Major()
        Major = lngPrivateMajor
    End Property

    Public Property Get Minor()
        Minor = lngPrivateMinor
    End Property

    Public Property Get Build()
        Build = lngPrivateBuild
    End Property

    Public Property Get Revision()
        Revision = lngPrivateRevision
    End Property

    Public Property Get MajorRevision()
        ' Returns the "upper" 16-bits of the revision number. The upper 16-bits are down-
        ' shifted by 16 bits and returned as a 16-bit integer.
        ' If the revision was uninitialized (-1), then -1 is returned.
        Dim lngBitMask
        Dim lngShiftRightDivisor
        If lngPrivateRevision = CLng(-1) Then
            MajorRevision = CInt(-1)
        Else
            lngBitMask = &H7FFF0000
            lngShiftRightDivisor = &H10000
            MajorRevision = CInt((lngPrivateRevision And lngBitMask) / lngShiftRightDivisor)
        End If
    End Property

    Public Property Get MinorRevision()
        ' Returns the "lower" 16-bits of the revision number, returned as a 16-bit integer.
        ' If the revision was uninitialized (-1), then -1 is returned.
        Dim lngBitMask
        If lngPrivateRevision = CLng(-1) Then
            MinorRevision = CInt(-1)
        Else
            ' 65535 is FFFF in hex; can't use hex because it's interpreted as -1
            lngBitMask = CLng(65535)
            MinorRevision = CInt(lngPrivateRevision And lngBitMask)
        End If
    End Property
End Class



'================= File: C:\Users\nanda\Github\sysadmin-accelerator\VBNext\index.vbs =================
Include(".\src\utils\Version.vbs")

Dim oVersion
Set oVersion = new Version
WScript.Echo oVersion.toString()
