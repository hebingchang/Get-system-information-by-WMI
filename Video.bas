Attribute VB_Name = "Video"

Public Function Video() As String
    Dim wmiObjSet As SWbemObjectSet
    Dim obj As SWbemObject
    Set wmiObjSet = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_VideoController")
    On Local Error Resume Next
    For Each obj In wmiObjSet
    Video = obj.videoprocessor
    Next
End Function

