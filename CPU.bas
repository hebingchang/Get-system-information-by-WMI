Attribute VB_Name = "CPU"
Public Function CpuName() As String
   On Error Resume Next
   Dim TmpCode$
   Dim objWMIService As Object, objItem As Object, colItems As Object
   Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor", , 48)
   For Each objItem In colItems
      CpuName = Trim(objItem.Name)
   Next
End Function

Public Function CpuID() As String
   On Error Resume Next
   Dim TmpCode$
   Dim objWMIService As Object, objItem As Object, colItems As Object
   Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor", , 48)
   For Each objItem In colItems
      CpuID = objItem.ProcessorId
   Next
End Function

Public Function CpuAddressWidth() As String
   On Error Resume Next
   Dim TmpCode$
   Dim objWMIService As Object, objItem As Object, colItems As Object
   Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor", , 48)
   For Each objItem In colItems
      CpuAddressWidth = objItem.AddressWidth
   Next
End Function

Public Function CpuArchitecture() As String
   On Error Resume Next
   Dim TmpCode$
   Dim objWMIService As Object, objItem As Object, colItems As Object
   Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor", , 48)
   For Each objItem In colItems
      CpuArchitecture = objItem.Architecture
   Next
End Function
Public Function CpuAvailability() As String
   On Error Resume Next
   Dim TmpCode$
   Dim objWMIService As Object, objItem As Object, colItems As Object
   Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor", , 48)
   For Each objItem In colItems
      CpuAvailability = objItem.Availability
   Next
End Function
Public Function CpuCaption() As String
   On Error Resume Next
   Dim TmpCode$
   Dim objWMIService As Object, objItem As Object, colItems As Object
   Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor", , 48)
   For Each objItem In colItems
      CpuCaption = objItem.Caption
   Next
End Function

Public Function DataWidth() As String
   On Error Resume Next
   Dim TmpCode$
   Dim objWMIService As Object, objItem As Object, colItems As Object
   Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor", , 48)
   For Each objItem In colItems
      DataWidth = objItem.DataWidth
   Next
End Function
Public Function L2CacheSize() As String
   On Error Resume Next
   Dim TmpCode$
   Dim objWMIService As Object, objItem As Object, colItems As Object
   Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor", , 48)
   For Each objItem In colItems
      L2CacheSize = objItem.L2CacheSize
   Next
End Function
Public Function L3CacheSize() As String
   On Error Resume Next
   Dim TmpCode$
   Dim objWMIService As Object, objItem As Object, colItems As Object
   Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor", , 48)
   For Each objItem In colItems
      L3CacheSize = objItem.L3CacheSize
   Next
End Function

