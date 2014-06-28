Attribute VB_Name = "Memory"
Public Sub AddPhysicalMemory()
   On Error Resume Next
   Dim TmpCode$
   Dim objWMIService As Object, objItem As Object, colItems As Object
   Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory", , 48)
   Dim xItem As ListItem
   For Each objItem In colItems
       Form1.ListView3.ListItems.Add , , "插槽"
        frmSplash.Label3.Caption = "正在获取内存插槽"
        Set xItem = Form1.ListView3.ListItems.Item(1)
        xItem.SubItems(1) = objItem.BankLabel
        Form1.ListView3.ListItems.Add , , "容量"
        frmSplash.Label3.Caption = "正在获取内存容量"
        Set xItem = Form1.ListView3.ListItems.Item(2)
        xItem.SubItems(1) = objItem.Capacity
        Form1.ListView3.ListItems.Add , , "厂商"
        frmSplash.Label3.Caption = "正在获取内存厂商"
        Set xItem = Form1.ListView3.ListItems.Item(3)
        xItem.SubItems(1) = objItem.Manufacturer
        Form1.ListView3.ListItems.Add , , "设备位置"
        frmSplash.Label3.Caption = "正在获取设备位置"
        Set xItem = Form1.ListView3.ListItems.Item(4)
        xItem.SubItems(1) = objItem.DeviceLocator
        Form1.ListView3.ListItems.Add , , "序列号"
        frmSplash.Label3.Caption = "正在获取内存序列号"
        Set xItem = Form1.ListView3.ListItems.Item(5)
        xItem.SubItems(1) = objItem.serialNumber
        Form1.ListView3.ListItems.Add , , "PartNumber"
        frmSplash.Label3.Caption = "正在获取PartNumber"
        Set xItem = Form1.ListView3.ListItems.Item(6)
        xItem.SubItems(1) = objItem.PartNumber
        Form1.ListView3.ListItems.Add , , "速度"
        frmSplash.Label3.Caption = "正在获取速度"
        Set xItem = Form1.ListView3.ListItems.Item(7)
        xItem.SubItems(1) = objItem.Speed
        Form1.ListView3.ListItems.Add , , "总宽度"
        frmSplash.Label3.Caption = "正在获取总宽度"
        Set xItem = Form1.ListView3.ListItems.Item(8)
        xItem.SubItems(1) = objItem.TotalWidth
   Next
End Sub

