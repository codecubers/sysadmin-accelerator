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
