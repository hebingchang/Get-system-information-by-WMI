VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "System Information"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   8655
   StartUpPosition =   3  '窗口缺省
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   10
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "系统信息"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(6)=   "Text1(0)"
      Tab(0).Control(7)=   "Text1(1)"
      Tab(0).Control(8)=   "Text1(2)"
      Tab(0).Control(9)=   "Text1(3)"
      Tab(0).Control(10)=   "Text1(4)"
      Tab(0).Control(11)=   "Text1(5)"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "CPU"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ListView1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "BIOS"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "内存"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ListView3"
      Tab(3).ControlCount=   1
      Begin MSComctlLib.ListView ListView3 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "属性"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "值"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "属性"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "值"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "属性"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "值"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   -72480
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   " "
         Top             =   3180
         Width           =   6000
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   -72480
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   " "
         Top             =   2700
         Width           =   6000
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   -72480
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   " "
         Top             =   2220
         Width           =   6000
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   -72480
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "  "
         Top             =   1740
         Width           =   6000
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   -72480
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   " "
         Top             =   1260
         Width           =   6000
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   -72480
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   780
         Width           =   6000
      End
      Begin VB.Label Label9 
         Caption         =   "Program Files文件夹"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   9
         Top             =   3180
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Windows文件夹"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   8
         Top             =   2700
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "用户名"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   7
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "用户域"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   6
         Top             =   1740
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "系统盘"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   5
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "系统内核"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   4
         Top             =   780
         Width           =   1215
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "系统/硬件信息"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI Light"
         Size            =   10.5
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "System Information"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI Light"
         Size            =   18
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00900000&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Module1.Attach Me.hwnd
    frmSplash.Label3.Caption = "正在获取系统信息"
    Dim str(100)
    str(0) = "OS"
    str(1) = "SystemRoot"
    str(2) = "USERDOMAIN"
    str(3) = "USERNAME"
    str(4) = "windir"
    str(5) = "ProgramFiles"
    For I = 0 To 5
        Text1(I).Text = Environ(str(I))
    Next
    ListView1.ListItems.Add , , "CPU"
    ListView1.ListItems.Add , , "CPU ID"
    ListView1.ListItems.Add , , "CPU 地址宽度"
    ListView1.ListItems.Add , , "CPU结构"
    ListView1.ListItems.Add , , "可用"
    ListView1.ListItems.Add , , "标题"
    ListView1.ListItems.Add , , "数据宽度"
    ListView1.ListItems.Add , , "二级缓存大小"
    ListView1.ListItems.Add , , "三级缓存大小"
    Dim xItem As ListItem
    frmSplash.Label3.Caption = "正在获取CPU名称"
    Set xItem = ListView1.ListItems.Item(1)
    xItem.SubItems(1) = CPU.CpuName
    frmSplash.Label3.Caption = "正在获取CPU ID"
    Set xItem = ListView1.ListItems.Item(2)
    xItem.SubItems(1) = CPU.CpuID
    frmSplash.Label3.Caption = "正在获取CPU地址宽度"
    Set xItem = ListView1.ListItems.Item(3)
    xItem.SubItems(1) = CPU.CpuAddressWidth
    frmSplash.Label3.Caption = "正在获取CPU架构"
    Set xItem = ListView1.ListItems.Item(4)
    xItem.SubItems(1) = CPU.CpuArchitecture
    frmSplash.Label3.Caption = "正在获取可用的CPU"
    Set xItem = ListView1.ListItems.Item(5)
    xItem.SubItems(1) = CPU.CpuAvailability
    frmSplash.Label3.Caption = "正在获取CPU标题"
    Set xItem = ListView1.ListItems.Item(6)
    xItem.SubItems(1) = CPU.CpuCaption
    frmSplash.Label3.Caption = "正在获取CPU地址宽度"
    Set xItem = ListView1.ListItems.Item(7)
    xItem.SubItems(1) = CPU.DataWidth
    frmSplash.Label3.Caption = "正在获取CPU二级缓存"
    Set xItem = ListView1.ListItems.Item(8)
    xItem.SubItems(1) = CPU.L2CacheSize
    frmSplash.Label3.Caption = "正在获取CPU三级缓存"
    Set xItem = ListView1.ListItems.Item(9)
    xItem.SubItems(1) = CPU.L3CacheSize
    BIOS.AddBIOS
    Memory.AddPhysicalMemory
End Sub
