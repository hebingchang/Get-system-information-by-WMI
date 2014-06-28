Attribute VB_Name = "Sound"
Public Function Sound() As String
    Dim wmiObjSet As SWbemObjectSet
    Dim obj As SWbemObject
    Set wmiObjSet = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_SoundDevice")
    On Local Error Resume Next
    For Each obj In wmiObjSet
    Sound = obj.ProductName
    Next
End Function
