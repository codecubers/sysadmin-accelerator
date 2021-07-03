Include(".\src\utils\Version.vbs")

Dim oVersion
Set oVersion = new Version
WScript.Echo oVersion.toString()