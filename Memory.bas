Attribute VB_Name = "Memory"
Public Sub AddPhysicalMemory()
   On Error Resume Next
   Dim TmpCode$
   Dim objWMIService As Object, objItem As Object, colItems As Object
   Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory", , 48)
   Dim xItem As ListItem
   For Each objItem In colItems
       Form1.ListView3.ListItems.Add , , "���"
        frmSplash.Label3.Caption = "���ڻ�ȡ�ڴ���"
        Set xItem = Form1.ListView3.ListItems.Item(1)
        xItem.SubItems(1) = objItem.BankLabel
        Form1.ListView3.ListItems.Add , , "����"
        frmSplash.Label3.Caption = "���ڻ�ȡ�ڴ�����"
        Set xItem = Form1.ListView3.ListItems.Item(2)
        xItem.SubItems(1) = objItem.Capacity
        Form1.ListView3.ListItems.Add , , "����"
        frmSplash.Label3.Caption = "���ڻ�ȡ�ڴ泧��"
        Set xItem = Form1.ListView3.ListItems.Item(3)
        xItem.SubItems(1) = objItem.Manufacturer
        Form1.ListView3.ListItems.Add , , "�豸λ��"
        frmSplash.Label3.Caption = "���ڻ�ȡ�豸λ��"
        Set xItem = Form1.ListView3.ListItems.Item(4)
        xItem.SubItems(1) = objItem.DeviceLocator
        Form1.ListView3.ListItems.Add , , "���к�"
        frmSplash.Label3.Caption = "���ڻ�ȡ�ڴ����к�"
        Set xItem = Form1.ListView3.ListItems.Item(5)
        xItem.SubItems(1) = objItem.serialNumber
        Form1.ListView3.ListItems.Add , , "PartNumber"
        frmSplash.Label3.Caption = "���ڻ�ȡPartNumber"
        Set xItem = Form1.ListView3.ListItems.Item(6)
        xItem.SubItems(1) = objItem.PartNumber
        Form1.ListView3.ListItems.Add , , "�ٶ�"
        frmSplash.Label3.Caption = "���ڻ�ȡ�ٶ�"
        Set xItem = Form1.ListView3.ListItems.Item(7)
        xItem.SubItems(1) = objItem.Speed
        Form1.ListView3.ListItems.Add , , "�ܿ��"
        frmSplash.Label3.Caption = "���ڻ�ȡ�ܿ��"
        Set xItem = Form1.ListView3.ListItems.Item(8)
        xItem.SubItems(1) = objItem.TotalWidth
   Next
End Sub

