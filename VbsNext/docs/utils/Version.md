## Class Version

```VB
'region VersionClassMetadata ####################################################
    ' Implements a VBScript version of the .NET System.Version class. Useful because .NET
    ' objects are not readily accessible in VBScript, and version-processing/comparison is a
    ' common systems administration activity.
    '
    ' Version: 1.1.20210613.1
    '
    ' Public Methods:
    '   Clone(ByRef objTargetVersionObject)
    '   CompareTo(ByVal objOtherVersionObject)
    '   CompareToString(ByVal strOtherVersion)
    '   Equals(ByVal objOtherVersionObject)
    '   GreaterThan(ByVal objOtherVersionObject)
    '   GreaterThanOrEqual(ByVal objOtherVersionObject)
    '   InitFromMajorMinor(ByVal lngMajor, ByVal lngMinor)
    '   InitFromMajorMinorBuild(ByVal lngMajor, ByVal lngMinor, ByVal lngBuild)
    '   InitFromMajorMinorBuildRevision(ByVal lngMajor, ByVal lngMinor, ByVal lngBuild,
    '       ByVal lngRevision)
    '   InitFromString(ByVal strVersion)
    '   LessThan(ByVal objOtherVersionObject)
    '   LessThanOrEqual(ByVal objOtherVersionObject)
    '   NotEquals(ByVal objOtherVersionObject)
    '   ToString()
    '
    ' Public Properties:
    '   Major (get)
    '   Minor (get)
    '   Build (get)
    '   Revision (get)
    '   MajorRevision (get)
    '   MinorRevision (get)
    '
    ' Not implemented:
    '   GetHashCode
    '   Parse (see InitFromString method)
    '   TryFormat (see ToString method)
    '   TryParse (see InitFromString method)
    '
    ' Note: the creation of a class such as this one requires VBScript 5.0, which is included
    ' in Internet Explorer 5.0 and was made available as a standalone download. One can also
    ' install Windows Scripting Host 2.0, which includes VBScript 5.1 and is compatible.
    ' Previous versions of VBScript (e.g., VBScript 3.0, included in Internet Explorer 4, IIS
    ' 4, Outlook 98, and Windows Scripting Host 1.0) are not compatible.
    '
    ' Example 1:
    ' Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    ' Set colItems = objWMI.ExecQuery("Select Version from Win32_OperatingSystem")
    ' For Each objItem in colItems
    '   strOSString = objItem.Version
    ' Next
    ' Set versionOperatingSystem = New Version
    ' intReturnCode = versionOperatingSystem.InitFromString(strOSString)
    ' If intReturnCode = 0 Then
    '   ' Success
    '   If versionOperatingSystem.CompareToString("10.0") >= 0 Then
    '       WScript.Echo("Windows 10, Windows Server 2016, or newer!")
    '   Else
    '       WScript.Echo("Windows 8.1, Windows Server 2012 R2, or older!")
    '   End If
    ' Else
    '   WScript.Echo("An error occurred reading the OS version.")
    ' End If
    '
    ' Example 2:
    ' Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    ' Set colItems = objWMI.ExecQuery("Select Version from Win32_OperatingSystem")
    ' For Each objItem in colItems
    '   strOSString = objItem.Version
    ' Next
    ' Set versionCurrentOperatingSystem = New Version
    ' intReturnCode = versionCurrentOperatingSystem.InitFromString(strOSString)
    ' If intReturnCode <> 0 Then
    '   WScript.Echo("Failed to get the current operating system version!")
    ' End If
    ' Set versionWindows98 = New Version
    ' intReturnCode = versionWindows98.InitFromMajorMinorBuild(4,10,1998)
    ' Set versionWindows98SE = New Version
    ' intReturnCode = versionWindows98SE.InitFromMajorMinorBuild(4,10,2222)
    ' Set versionWindowsME = New Version
    ' intReturnCode = versionWindowsME.InitFromMajorMinor(4,90)
    ' bool9x = False
    ' If versionCurrentOperatingSystem.GreaterThanOrEqual(versionWindows98) And versionCurrentOperatingSystem.LessThanOrEqual(versionWindows98SE) Then
    '   bool9x = True
    ' ElseIf (versionCurrentOperatingSystem.Major = versionWindowsME.Major) And (versionCurrentOperatingSystem.Minor = versionWindowsME.Minor) Then
    '   bool9x = True
    ' End If
    ' If bool9x Then
    '   WScript.Echo("Current OS is Windows 9x. It's 2020 (or later). What are you thinking?")
    ' Else
    '   WScript.Echo("Thank the maker! This OS is not Windows 9x.")
    ' End If
    'endregion VersionClassMetadata ####################################################

    'region License ####################################################
    ' Copyright 2021 Frank Lesniak
    '
    ' Permission is hereby granted, free of charge, to any person obtaining a copy of this
    ' software and associated documentation files (the "Software"), to deal in the Software
    ' without restriction, including without limitation the rights to use, copy, modify, merge,
    ' publish, distribute, sublicense, and/or sell copies of the Software, and to permit
    ' persons to whom the Software is furnished to do so, subject to the following conditions:
    '
    ' The above copyright notice and this permission notice shall be included in all copies or
    ' substantial portions of the Software.
    '
    ' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
    ' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
    ' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
    ' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
    ' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
    ' DEALINGS IN THE SOFTWARE.
    'endregion License ####################################################

    'region DownloadLocationNotice ####################################################
    ' The most up-to-date version of this script can be found on the author's GitHub repository
    ' at https://github.com/franklesniak/sysadmin-accelerator
    'endregion DownloadLocationNotice ####################################################

    'region Acknowledgements ####################################################
    ' Andrew Clinick, for writing the MSDN article "Clinick's Clinic on Scripting: Take Five
    ' What's New in the Version 5.0 Script Engines" - which confirmed that a VBScript class
    ' requires 5.0 of the script engine.
    '
    ' Jerry Lee Ford, Jr., for providing a history of VBScript and Windows Scripting Host in
    ' his book, "Microsoft WSH and VBScript Programming for the Absolute Beginner".
    '
    ' Gunter Born, for providing a history of Windows Scripting Host in his book "Microsoft
    ' Windows Script Host 2.0 Developer's Guide" that corrected some points.
    'endregion Acknowledgements ####################################################

    'region DependsOn ####################################################
    ' None - this class is entirely self-contained. However, this class contains a private
    ' function TestObjectForData() that should be identical to the public TestObjectForData()
    'endregion DependsOn ####################################################
```



### Function TestObjectForData(ByVal objToCheck)

```VB
'region FunctionMetadata ####################################################
        ' Checks an object or variable to see if it "has data".
        ' If any of the following are true, then objToCheck is regarded as NOT having data:
        '   VarType(objToCheck) = 0
        '   VarType(objToCheck) = 1
        '   objToCheck Is Nothing
        '   IsEmpty(objToCheck)
        '   IsNull(objToCheck)
        '   objToCheck = vbNullString (or "")
        '   IsArray(objToCheck) = True And UBound(objToCheck) throws an error
        '   IsArray(objToCheck) = True And UBound(objToCheck) < 0
        ' In any of these cases, the function returns False. Otherwise, it returns True.
        '
        ' Version: 1.1.20210115.0
        'endregion FunctionMetadata ####################################################
    
        'region License ####################################################
        ' Copyright 2021 Frank Lesniak
        '
        ' Permission is hereby granted, free of charge, to any person obtaining a copy of this
        ' software and associated documentation files (the "Software"), to deal in the Software
        ' without restriction, including without limitation the rights to use, copy, modify, merge,
        ' publish, distribute, sublicense, and/or sell copies of the Software, and to permit
        ' persons to whom the Software is furnished to do so, subject to the following conditions:
        '
        ' The above copyright notice and this permission notice shall be included in all copies or
        ' substantial portions of the Software.
        '
        ' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
        ' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
        ' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
        ' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
        ' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
        ' DEALINGS IN THE SOFTWARE.
        'endregion License ####################################################
    
        'region DownloadLocationNotice ####################################################
        ' The most up-to-date version of this script can be found on the author's GitHub repository
        ' at https://github.com/franklesniak/Test_Object_For_Data
        'endregion DownloadLocationNotice ####################################################
    
        'region Acknowledgements ####################################################
        ' Thanks to Scott Dexter for writing the article "Empty Nothing And Null How Do You Feel
        ' Today", which inspired me to create this function. https://evolt.org/node/346
        '
        ' Thanks also to "RhinoScript" for the article "Testing for Empty Arrays" for providing
        ' guidance for how to test for the empty array condition in VBScript.
        ' https://wiki.mcneel.com/developer/scriptsamples/emptyarray
        '
        ' Thanks also "iamresearcher" who posted this and inspired the test case for vbNullString:
        ' https://www.vbforums.com/showthread.php?684799-The-Differences-among-Empty-Nothing-vbNull-vbNullChar-vbNullString-and-the-Zero-L
        'endregion Acknowledgements ####################################################
    
```