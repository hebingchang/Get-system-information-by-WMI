Attribute VB_Name = "BIOS"
Public Sub AddBIOS()
   On Error Resume Next
   Dim TmpCode$
   Dim objWMIService As Object, objItem As Object, colItems As Object
   Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_BIOS", , 48)
   Dim xItem As ListItem
   For Each objItem In colItems
       Form1.ListView2.ListItems.Add , , "����"
        frmSplash.Label3.Caption = "���ڻ�ȡBIOS����"
        Set xItem = Form1.ListView2.ListItems.Item(1)
        xItem.SubItems(1) = objItem.Caption
        Form1.ListView2.ListItems.Add , , "��ǰ����"
        frmSplash.Label3.Caption = "���ڻ�ȡBIOS����"
        Set xItem = Form1.ListView2.ListItems.Item(2)
        xItem.SubItems(1) = objItem.CurrentLanguage
        Form1.ListView2.ListItems.Add , , "����"
        frmSplash.Label3.Caption = "���ڻ�ȡBIOS����"
        Set xItem = Form1.ListView2.ListItems.Item(3)
        xItem.SubItems(1) = objItem.Manufacturer
        Form1.ListView2.ListItems.Add , , "��������"
        frmSplash.Label3.Caption = "���ڻ�ȡBIOS��������"
        Set xItem = Form1.ListView2.ListItems.Item(4)
        xItem.SubItems(1) = objItem.ReleaseDate
        Form1.ListView2.ListItems.Add , , "���к�"
        frmSplash.Label3.Caption = "���ڻ�ȡBIOS���к�"
        Set xItem = Form1.ListView2.ListItems.Item(5)
        xItem.SubItems(1) = objItem.serialNumber
        Form1.ListView2.ListItems.Add , , "SMBIOS�汾"
        frmSplash.Label3.Caption = "���ڻ�ȡBIOS�汾"
        Set xItem = Form1.ListView2.ListItems.Item(6)
        xItem.SubItems(1) = objItem.SMBIOSBIOSVersion
        Form1.ListView2.ListItems.Add , , "�汾"
        frmSplash.Label3.Caption = "���ڻ�ȡBIOS�汾"
        Set xItem = Form1.ListView2.ListItems.Item(7)
        xItem.SubItems(1) = objItem.Version
   Next
End Sub
