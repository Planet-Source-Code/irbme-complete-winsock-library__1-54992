VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEnumeration 
   Caption         =   "W I S P A (Copyright (c) Chris Waddell) - Enum Demo"
   ClientHeight    =   5760
   ClientLeft      =   165
   ClientTop       =   780
   ClientWidth     =   5880
   Icon            =   "frmEnumeration.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   70
      Left            =   5040
      ScaleHeight     =   75
      ScaleWidth      =   1590
      TabIndex        =   6
      Top             =   5145
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.PictureBox picMiddle 
      BorderStyle     =   0  'None
      Height          =   2640
      Left            =   105
      ScaleHeight     =   2640
      ScaleWidth      =   9915
      TabIndex        =   5
      Top             =   1365
      Width           =   9915
      Begin VB.PictureBox picServices 
         Height          =   1800
         Left            =   7560
         ScaleHeight     =   1740
         ScaleWidth      =   1740
         TabIndex        =   27
         Top             =   1890
         Visible         =   0   'False
         Width           =   1800
         Begin VB.TextBox txtService 
            Height          =   285
            Left            =   210
            TabIndex        =   31
            Text            =   "tcp"
            Top             =   420
            Width           =   1275
         End
         Begin VB.OptionButton optServByPort 
            Caption         =   "By Port"
            Height          =   225
            Left            =   105
            TabIndex        =   30
            Top             =   1050
            Width           =   1590
         End
         Begin VB.OptionButton optServByName 
            Caption         =   "By Name"
            Height          =   195
            Left            =   105
            TabIndex        =   29
            Top             =   840
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.CommandButton cmdGetService 
            Caption         =   "Get Service"
            Height          =   330
            Left            =   105
            TabIndex        =   28
            Top             =   1365
            Width           =   1590
         End
         Begin VB.Label lblService 
            Caption         =   "Service Name:"
            Height          =   225
            Left            =   105
            TabIndex        =   32
            Top             =   105
            Width           =   1485
         End
      End
      Begin VB.PictureBox picProtocols 
         Height          =   1800
         Left            =   7560
         ScaleHeight     =   1740
         ScaleWidth      =   1740
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   1800
         Begin VB.CommandButton cmdGetProtocol 
            Caption         =   "Get Protocol"
            Height          =   330
            Left            =   105
            TabIndex        =   26
            Top             =   1365
            Width           =   1590
         End
         Begin VB.OptionButton optProtoByName 
            Caption         =   "By Name"
            Height          =   195
            Left            =   105
            TabIndex        =   25
            Top             =   840
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton optProtoByNumber 
            Caption         =   "By Number"
            Height          =   225
            Left            =   105
            TabIndex        =   24
            Top             =   1050
            Width           =   1590
         End
         Begin VB.TextBox txtProtocol 
            Height          =   285
            Left            =   210
            TabIndex        =   22
            Text            =   "tcp"
            Top             =   420
            Width           =   1275
         End
         Begin VB.Label lblProtocol 
            Caption         =   "Protocol Name:"
            Height          =   225
            Left            =   105
            TabIndex        =   23
            Top             =   105
            Width           =   1485
         End
      End
      Begin MSComctlLib.ListView lvInterfaces 
         Height          =   750
         Left            =   5460
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1323
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.PictureBox picInterfaceSplitter 
         Height          =   2010
         Left            =   7140
         MousePointer    =   9  'Size W E
         ScaleHeight     =   1950
         ScaleWidth      =   15
         TabIndex        =   20
         Top             =   1155
         Visible         =   0   'False
         Width           =   70
      End
      Begin VB.ListBox lstInterfaces 
         Height          =   1035
         Left            =   4200
         TabIndex        =   18
         Top             =   2205
         Visible         =   0   'False
         Width           =   2000
      End
      Begin VB.PictureBox picProtocolSplitter 
         Height          =   2010
         Left            =   7350
         MousePointer    =   9  'Size W E
         ScaleHeight     =   1950
         ScaleWidth      =   15
         TabIndex        =   17
         Top             =   1155
         Visible         =   0   'False
         Width           =   70
      End
      Begin VB.ListBox lstProtocols 
         Height          =   1035
         Left            =   4200
         TabIndex        =   16
         Top             =   1155
         Visible         =   0   'False
         Width           =   2000
      End
      Begin MSComctlLib.ListView lvServices 
         Height          =   750
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1323
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvProtocols 
         Height          =   750
         Left            =   4095
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1323
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvTransProtocols 
         Height          =   750
         Left            =   2730
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1323
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvStartup 
         Height          =   750
         Left            =   1365
         TabIndex        =   8
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1323
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.TabStrip tsMiddle 
         Height          =   1590
         Left            =   105
         TabIndex        =   12
         Top             =   840
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   2805
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picBottomSplitter 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   100
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   105
      ScaleWidth      =   5880
      TabIndex        =   4
      Top             =   3255
      Width           =   5880
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Height          =   2400
      Left            =   0
      ScaleHeight     =   2340
      ScaleWidth      =   5820
      TabIndex        =   3
      Top             =   3360
      Width           =   5880
      Begin RichTextLib.RichTextBox txtErrors 
         Height          =   1275
         Left            =   2835
         TabIndex        =   11
         Top             =   735
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   2249
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmEnumeration.frx":030A
      End
      Begin RichTextLib.RichTextBox txtInfo 
         Height          =   1275
         Left            =   210
         TabIndex        =   10
         Top             =   735
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2249
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmEnumeration.frx":038C
      End
      Begin MSComctlLib.TabStrip tsBottom 
         Height          =   2010
         Left            =   105
         TabIndex        =   7
         Top             =   210
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   3545
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   5880
      TabIndex        =   0
      Top             =   0
      Width           =   5880
      Begin VB.PictureBox picLogo 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2610
         Left            =   6825
         Picture         =   "frmEnumeration.frx":040E
         ScaleHeight     =   174
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   140
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Label lblHeader2 
         BackStyle       =   0  'Transparent
         Caption         =   "Windows Internetwork Socket Programming  API"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   1365
         TabIndex        =   2
         Top             =   630
         Width           =   3795
      End
      Begin VB.Label lblHeader 
         BackStyle       =   0  'Transparent
         Caption         =   "W  I   S   P   A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   435
         Left            =   1365
         TabIndex        =   1
         Top             =   210
         Width           =   3795
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuRestart 
         Caption         =   "&Restart Winsock"
      End
      Begin VB.Menu mnuRefreshProtocols 
         Caption         =   "Refresh Transport &Protocols"
      End
      Begin VB.Menu mnuRefreshInterfaces 
         Caption         =   "Refresh Network &Interfaces"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmEnumeration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Private WithEvents Startup          As CWinsockStartup
Attribute Startup.VB_VarHelpID = -1
Private WithEvents Protocols        As CProtocolEnumeration
Attribute Protocols.VB_VarHelpID = -1
Private WithEvents Socket           As CSocket
Attribute Socket.VB_VarHelpID = -1
Private WithEvents Interfaces       As CInterfaceEnumeration
Attribute Interfaces.VB_VarHelpID = -1
Private WithEvents ProtocolEntries  As CProtocolEntry
Attribute ProtocolEntries.VB_VarHelpID = -1
Private WithEvents ServiceEntries   As CServiceEntry
Attribute ServiceEntries.VB_VarHelpID = -1


Private Sub StartWinsock()

    ' Attempt to start up winsock
    
    AddLog txtInfo, "Starting up winsock"
    
    Set Startup = New CWinsockStartup
    
    If Not Startup.SuccessfulyStartedUp Then
        AddLog txtInfo, "Could not start up winsock"
        Exit Sub
    End If

    AddLog txtInfo, "Winsock successfuly started"
    
End Sub


Private Sub RefreshTransportProtocols()

  Dim i As Long

    ' Attempt to enumerate the protocols
    
    lstProtocols.Clear
    
    AddLog txtInfo, "Enumerating protocols"
    
    Set Protocols = New CProtocolEnumeration
    Protocols.Initialize
    
    For i = 1 To Protocols.Count
        If Protocols.Item(i) Is Nothing Then Exit For
        lstProtocols.AddItem Protocols.Item(i).ProtocolName
    Next i
    
    AddLog txtInfo, "Finished enumerating protocols"

    Set Socket = New CSocket
    
    lstProtocols.ListIndex = 0
    Call lstProtocols_Click

End Sub


Private Sub RefreshInterfaces()

  Dim i As Long

    ' Attempt to enumerate network interfaces

    lstInterfaces.Clear
    
    AddLog txtInfo, "Enumerating network interfaces"
    
    Set Interfaces = New CInterfaceEnumeration

    Interfaces.Initialize

    For i = 1 To Interfaces.Count
        If Interfaces.Item(i).Address.IPAddress Is Nothing Then
            Exit For
        End If
        
        lstInterfaces.AddItem Interfaces.Item(i).Address.IPAddress.StringAddress
    Next i

    AddLog txtInfo, "Finished enumerating network interfaces"

    lstInterfaces.ListIndex = 0
    Call lstInterfaces_Click

End Sub


Private Sub cmdGetProtocol_Click()

  Dim i As Long

    If optProtoByName.Value = True Then
        ProtocolEntries.ProtocolName = txtProtocol.Text
    Else
        If Not IsNumeric(txtProtocol.Text) Then Exit Sub
        ProtocolEntries.ProtocolNumber = CLng(txtProtocol.Text)
    End If
    
    lvProtocols.ListItems.Clear
    
    lvProtocols.ListItems.Add , , "Name"
    lvProtocols.ListItems(1).SubItems(1) = ProtocolEntries.ProtocolName
    
    lvProtocols.ListItems.Add , , "Number"
    lvProtocols.ListItems(2).SubItems(1) = ProtocolEntries.ProtocolNumber
    
    lvProtocols.ListItems.Add , , "Aliases"
    
    If ProtocolEntries.Aliases Is Nothing Then Exit Sub
    
    lvProtocols.ListItems(3).SubItems(1) = ProtocolEntries.Aliases(1)

    If ProtocolEntries.Aliases.Count >= 2 Then
        For i = 2 To ProtocolEntries.Aliases.Count
            lvProtocols.ListItems.Add , , vbNullString
            lvProtocols.ListItems(i + 2).SubItems(1) = ProtocolEntries.Aliases(i)
        Next i
    End If

End Sub


Private Sub cmdGetService_Click()

  Dim i As Long

    If optServByName.Value = True Then
        ServiceEntries.ServiceName = txtService.Text
    Else
        If Not IsNumeric(txtService.Text) Then Exit Sub
        ServiceEntries.Port = CLng(txtService.Text)
    End If
    
    lvServices.ListItems.Clear
    
    lvServices.ListItems.Add , , "Name"
    lvServices.ListItems(1).SubItems(1) = ServiceEntries.ServiceName
    
    lvServices.ListItems.Add , , "Port"
    lvServices.ListItems(2).SubItems(1) = ServiceEntries.Port
    
    lvServices.ListItems.Add , , "Protocol"
    lvServices.ListItems(3).SubItems(1) = ServiceEntries.ProtocolName
    
    lvServices.ListItems.Add , , "Aliases"
    
    If ServiceEntries.Aliases Is Nothing Then Exit Sub
    
    If ServiceEntries.Aliases.Count = 0 Then Exit Sub
    
    lvServices.ListItems(4).SubItems(1) = ServiceEntries.Aliases(1)

    If ServiceEntries.Aliases.Count >= 2 Then
        For i = 2 To ServiceEntries.Aliases.Count
            lvServices.ListItems.Add , , vbNullString
            lvServices.ListItems(i + 3).SubItems(1) = ServiceEntries.Aliases(i)
        Next i
    End If

End Sub


Private Sub Form_Load()

    picInterfaceSplitter.BorderStyle = vbBSNone
    picProtocolSplitter.BorderStyle = vbBSNone
    
    tsBottom.Tabs(1).Key = "info"
    tsBottom.Tabs(1).Caption = "Information"
    tsBottom.Tabs.Add 2, "error", "Errors"
    
    tsMiddle.Tabs(1).Key = "startup"
    tsMiddle.Tabs(1).Caption = "Startup Info"
    tsMiddle.Tabs.Add 2, "interfaces", "Network Interfaces"
    tsMiddle.Tabs.Add 3, "transprotocols", "Transport Protocols"
    tsMiddle.Tabs.Add 4, "protocols", "Protocols"
    tsMiddle.Tabs.Add 5, "services", "Services"
    
    lvStartup.ColumnHeaders.Add , , "Property"
    lvStartup.ColumnHeaders.Add , , "Value"
    
    lvInterfaces.ColumnHeaders.Add , , "Property", 2000
    lvInterfaces.ColumnHeaders.Add , , "Value", 2000
    
    lvTransProtocols.ColumnHeaders.Add , , "Property", 2500
    lvTransProtocols.ColumnHeaders.Add , , "Value", 4500
    
    lvProtocols.ColumnHeaders.Add , , "Property"
    lvProtocols.ColumnHeaders.Add , , "Value"
    
    lvServices.ColumnHeaders.Add , , "Property"
    lvServices.ColumnHeaders.Add , , "Value"
    
    StartWinsock
    
    ' Add the infromation about the startup to the listview
    
    lvStartup.ListItems.Add , , "Description"
    lvStartup.ListItems(1).SubItems(1) = Startup.Description
    
    lvStartup.ListItems.Add , , "Strated Up"
    lvStartup.ListItems(2).SubItems(1) = IIf(Startup.SuccessfulyStartedUp, "Yes", "No")
    
    lvStartup.ListItems.Add , , "System Status"
    lvStartup.ListItems(3).SubItems(1) = Startup.SystemStatus
    
    lvStartup.ListItems.Add , , "Version"
    lvStartup.ListItems(4).SubItems(1) = Startup.Version.StringVersion
    
    lvStartup.ListItems.Add , , "Highest Version"
    lvStartup.ListItems(5).SubItems(1) = Startup.HighestVersion.StringVersion
    
    lvStartup.ListItems.Add , , "Maximum Sockets"
    lvStartup.ListItems(6).SubItems(1) = IIf(Startup.MaxSockets = 0, "Unlimited", Startup.MaxSockets)

    RefreshInterfaces
    RefreshTransportProtocols

    AddLog txtInfo, "Getting Protocol Entry for TCP"
    Set ProtocolEntries = New CProtocolEntry
    txtProtocol.Text = "tcp"
    Call cmdGetProtocol_Click

    AddLog txtInfo, "Getting Service Entry for HTTP"
    Set ServiceEntries = New CServiceEntry
    txtService.Text = "http"
    Call cmdGetService_Click
    
End Sub


Private Sub Form_Resize()

  On Error Resume Next
    
    ' Don't try to resize if the window is minimized
    If Me.WindowState = vbMinimized Then Exit Sub

    ' Ensure the window sticks to the minimum size of 6000*6500
    If Me.Width < 6000 Then Me.Width = 6000
    If Me.Height < 6500 Then Me.Height = 6500
    
    ' The header always has a fixed height
    picHeader.Height = 1300

    ' Resize the heading text
    lblHeader.Left = 1400
    lblHeader.Top = (picHeader.ScaleHeight \ 2) - (lblHeader.Height \ 2)
    lblHeader.Width = picHeader.ScaleWidth - lblHeader.Left - 100
    
    ' Resize the second part of the heading text
    lblHeader2.Top = lblHeader.Top + lblHeader.Height - 50
    lblHeader2.Left = lblHeader.Left + 500
    lblHeader2.Width = picHeader.ScaleWidth - lblHeader2.Left - 100

    ' Resize the middle pic, the main part of the application
    picMiddle.Left = 0
    picMiddle.Top = picHeader.Height
    picMiddle.Height = Me.ScaleHeight - picHeader.Height - picBottom.Height - picBottomSplitter.Height
    picMiddle.Width = Me.ScaleWidth

End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    Set ProtocolEntries = Nothing
    Set ServiceEntries = Nothing
    Set Interfaces = Nothing
    Set Protocols = Nothing
    Set Socket = Nothing
    
    ' Terminate the startup object, cleaning up winsock
    Set Startup = Nothing
End Sub

Private Sub lstInterfaces_Click()

  Dim i As Long
  Dim Interface As CInterface

    If lvInterfaces.ListItems.Count > 0 Then
        If lvInterfaces.ListItems(1).SubItems(1) = lstInterfaces.List(lstInterfaces.ListIndex) Then Exit Sub
    End If
    
    For i = 1 To Interfaces.Count
        Set Interface = Interfaces.Item(i)
        
        If Interfaces Is Nothing Then Exit For
        If Interface.Address.IPAddress.StringAddress = lstInterfaces.List(lstInterfaces.ListIndex) Then
            
            lvInterfaces.ListItems.Clear
            
            lvInterfaces.ListItems.Add , , "Address"
            lvInterfaces.ListItems(1).SubItems(1) = Interface.Address.IPAddress.StringAddress
            
            lvInterfaces.ListItems.Add , , "Subnet Mask"
            lvInterfaces.ListItems(2).SubItems(1) = Interface.SubnetMask.IPAddress.StringAddress
            
            lvInterfaces.ListItems.Add , , "Enabled"
            lvInterfaces.ListItems(3).SubItems(1) = IIf(Interface.InterfaceUp, "Yes", "No")
            
            lvInterfaces.ListItems.Add , , "Is Loopback"
            lvInterfaces.ListItems(4).SubItems(1) = IIf(Interface.IsLoopbackInterface, "Yes", "No")
            
            lvInterfaces.ListItems.Add , , "Is Point To Point Link"
            lvInterfaces.ListItems(5).SubItems(1) = IIf(Interface.IsPointToPointLink, "Yes", "No")
            
            lvInterfaces.ListItems.Add , , "Broadcasting Supported"
            lvInterfaces.ListItems(6).SubItems(1) = IIf(Interface.BroadcastSupported, "Yes", "No")
   
            lvInterfaces.ListItems.Add , , "Broadcast Address"
            lvInterfaces.ListItems(7).SubItems(1) = Interface.BroadcastAddress.IPAddress.StringAddress
            
            lvInterfaces.ListItems.Add , , "Multicast Supported"
            lvInterfaces.ListItems(8).SubItems(1) = IIf(Interface.MulticastSupported, "Yes", "No")
  
            Exit For
        End If
        
    Next i
    
    Set Interface = Nothing
    
End Sub


Private Sub lstProtocols_Click()

  Dim i As Long
  Dim Protocol As cprotocolinfo
  Dim ProtocolChain As CProtocolChain
  Dim s As String

    If lvTransProtocols.ListItems.Count > 0 Then
        If lvTransProtocols.ListItems(1).SubItems(1) = lstProtocols.List(lstProtocols.ListIndex) Then Exit Sub
    End If
    
    For i = 1 To Protocols.Count
        Set Protocol = Protocols.Item(i)
        
        If Protocol Is Nothing Then Exit For
        If Protocol.ProtocolName = lstProtocols.List(lstProtocols.ListIndex) Then
            
            Set ProtocolChain = Protocol.ProtocolChain
            
            lvTransProtocols.ListItems.Clear
            
            lvTransProtocols.ListItems.Add , , "Protocol Name"
            lvTransProtocols.ListItems(1).SubItems(1) = Protocol.ProtocolName
        
            lvTransProtocols.ListItems.Add , , "Address Family"
            lvTransProtocols.ListItems(2).SubItems(1) = Socket.AddressFamilyName(Protocol.AddressFamily)
            
            lvTransProtocols.ListItems.Add , , "Socket Type"
            lvTransProtocols.ListItems(3).SubItems(1) = Socket.SocketTypeName(Protocol.SocketType)
            
            lvTransProtocols.ListItems.Add , , "Protocol Type"
            lvTransProtocols.ListItems(4).SubItems(1) = Socket.ProtocolName(Protocol.Protocol) & " Version " & Protocol.Version.StringVersion
        
            lvTransProtocols.ListItems.Add , , "Address Size"
            lvTransProtocols.ListItems(5).SubItems(1) = Protocol.MinSockAddr & " bytes" & IIf(Protocol.MinSockAddr < Protocol.MaxSockAddr, " to " & Protocol.MaxSockAddr & " bytes", vbNullString)
            
            lvTransProtocols.ListItems.Add , , "Network Byte Order"
            lvTransProtocols.ListItems(6).SubItems(1) = IIf(Protocol.NetworkByteOrder = BOrd_BigEndian, "Big Endian", "Little Endian")
        
            lvTransProtocols.ListItems.Add , , "Message Size"
            lvTransProtocols.ListItems(7).SubItems(1) = IIf(Protocol.MessageSize > 0, Protocol.MessageSize & " bytes", "Unlimited")
        
            lvTransProtocols.ListItems.Add , , "Provider ID"
            lvTransProtocols.ListItems(8).SubItems(1) = Protocol.ProviderId
        
            lvTransProtocols.ListItems.Add , , "Catalog Entry ID"
            lvTransProtocols.ListItems(9).SubItems(1) = Protocol.CatalogEntryId
            
            lvTransProtocols.ListItems.Add , , "Connection Oriented"
            lvTransProtocols.ListItems(10).SubItems(1) = IIf(Protocol.Connectionless, "No", "Yes")
            
            lvTransProtocols.ListItems.Add , , "Supports Connnect Data"
            lvTransProtocols.ListItems(11).SubItems(1) = IIf(Protocol.ConnectData, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Supports Disconnect Data"
            lvTransProtocols.ListItems(12).SubItems(1) = IIf(Protocol.DisconnectData, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Supports Expedited (Urgent) Data"
            lvTransProtocols.ListItems(13).SubItems(1) = IIf(Protocol.ExpeditedData, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Supports Graceful Closing"
            lvTransProtocols.ListItems(14).SubItems(1) = IIf(Protocol.GracefulClose, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Guaranteed Delivery"
            lvTransProtocols.ListItems(15).SubItems(1) = IIf(Protocol.GuaranteedDelivery, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Guaranteed Order"
            lvTransProtocols.ListItems(16).SubItems(1) = IIf(Protocol.GuaranteedOrder, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Matches Protocol Zero"
            lvTransProtocols.ListItems(17).SubItems(1) = IIf(Protocol.MatchesProtocolZero, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Message Oriented"
            lvTransProtocols.ListItems(18).SubItems(1) = IIf(Protocol.MessageOriented, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Pseudo Stream Oriented"
            lvTransProtocols.ListItems(19).SubItems(1) = IIf(Protocol.PseudoStream, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Supports Broadcast Data"
            lvTransProtocols.ListItems(20).SubItems(1) = IIf(Protocol.SupportBroadcast, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Supports Multipoint Data"
            lvTransProtocols.ListItems(21).SubItems(1) = IIf(Protocol.SupportMultipoint, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Multiple Protocol Entrues"
            lvTransProtocols.ListItems(22).SubItems(1) = IIf(Protocol.MultipleProtoEntries, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Multipoint Control Plane"
            lvTransProtocols.ListItems(23).SubItems(1) = IIf(Protocol.MultipointControlPlane, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Multipoint Data Plane"
            lvTransProtocols.ListItems(24).SubItems(1) = IIf(Protocol.MultipointDataPlane, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Partial Messages Supported"
            lvTransProtocols.ListItems(25).SubItems(1) = IIf(Protocol.PartialMessage, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Supports Quality Of Service"
            lvTransProtocols.ListItems(26).SubItems(1) = IIf(Protocol.QoSSupport, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Recommended Protocol Entry"
            lvTransProtocols.ListItems(27).SubItems(1) = IIf(Protocol.RecommendedProtoEntry, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Unidirectional In Recieve Direction"
            lvTransProtocols.ListItems(28).SubItems(1) = IIf(Protocol.UnidirectionalRecv, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Unidirectional In Send Direction"
            lvTransProtocols.ListItems(29).SubItems(1) = IIf(Protocol.UnidirectionalSend, "Yes", "No")
            
            lvTransProtocols.ListItems.Add , , "Uses Installable File System Handles"
            lvTransProtocols.ListItems(30).SubItems(1) = IIf(Protocol.UsesInstallableFileSystemHandles, "Yes", "No")
            
            Exit For
        End If
    Next i

    lvTransProtocols.ListItems.Add , , "Chain Type"
    Select Case ProtocolChain.ChainType
        Case protocolchaintype.BaseProtocol
            lvTransProtocols.ListItems(31).SubItems(1) = "Base Protocol"
        Case protocolchaintype.LayeredProtocol
            lvTransProtocols.ListItems(31).SubItems(1) = "Layered Protocol"
        Case protocolchaintype.LayeredProtocolChain
            lvTransProtocols.ListItems(31).SubItems(1) = "Layered Protocol Chain"
    End Select
     
    For i = 1 To ProtocolChain.ChainEntryCount
        s = s & ProtocolChain.ChainEntry(i) & ","
    Next i
    
    s = Left$(s, Len(s) - 1)

    lvTransProtocols.ListItems.Add , , "Chain Entry Count"
    lvTransProtocols.ListItems(32).SubItems(1) = CStr(ProtocolChain.ChainEntryCount)

    If s <> vbNullString Then
        lvTransProtocols.ListItems.Add , , "Chain Entries"
        lvTransProtocols.ListItems(33).SubItems(1) = s
    End If
    
    Set Protocol = Nothing
    Set ProtocolChain = Nothing

    
End Sub

Private Sub mnuExit_Click()
    ' Exit application
    Unload Me
End Sub

Private Sub mnuRefreshInterfaces_Click()
    RefreshInterfaces
End Sub

Private Sub mnuRefreshProtocols_Click()
    RefreshTransportProtocols
End Sub

Private Sub mnuRestart_Click()
    Set Startup = Nothing
    StartWinsock
End Sub

Private Sub optProtoByName_Click()
    lblProtocol.Caption = "Protocol Name:"
End Sub

Private Sub optProtoByNumber_Click()
    lblProtocol.Caption = "Protocol Number:"
End Sub

Private Sub optServByName_Click()
    lblService.Caption = "Service Name:"
End Sub

Private Sub optServByPort_Click()
    lblService.Caption = "Service Port:"
End Sub

Private Sub picBottom_Resize()

    ' Resize the tab strip at the bottom
    tsBottom.Left = 100
    tsBottom.Top = 100
    tsBottom.Width = picBottom.ScaleWidth - tsBottom.Left - 50
    tsBottom.Height = picBottom.ScaleHeight - tsBottom.Top - 50
    
    ' Resize the information log textbox
    txtInfo.Left = 200
    txtInfo.Top = 500
    txtInfo.Width = picBottom.ScaleWidth - txtInfo.Left - 100
    txtInfo.Height = picBottom.ScaleHeight - txtInfo.Top - 100
    
    ' Resize the error log textbox
    txtErrors.Left = 200
    txtErrors.Top = 500
    txtErrors.Width = picBottom.ScaleWidth - txtErrors.Left - 100
    txtErrors.Height = picBottom.ScaleHeight - txtErrors.Top - 100
End Sub



Private Sub picHeader_Paint()
    ' Do the gradient fill and draw the logo
    GradientFill picHeader, True, RGB(150, 150, 150), RGB(250, 250, 250)
    TransparentBlt picHeader.hdc, 10, 10, 80, 70, picLogo.hdc, 0, 0, picLogo.ScaleWidth, picLogo.ScaleHeight, vbMagenta
End Sub


Private Sub picMiddle_Resize()

  On Error GoTo ResizeError

    ' Tabstrip

    tsMiddle.Left = 100
    tsMiddle.Top = 100
    tsMiddle.Width = picMiddle.ScaleWidth - tsMiddle.Left - 50
    tsMiddle.Height = picMiddle.ScaleHeight - tsMiddle.Top - 50

    ' Startup Data

    lvStartup.Left = 200
    lvStartup.Top = 500
    lvStartup.Width = picMiddle.ScaleWidth - lvStartup.Left - 100
    lvStartup.Height = picMiddle.ScaleHeight - lvStartup.Top - 100
    
    ' Network Interfaces

    lstInterfaces.Left = 200
    lstInterfaces.Top = 500
    lstInterfaces.Height = picMiddle.ScaleHeight - lvInterfaces.Top - 100

    picInterfaceSplitter.Left = lstInterfaces.Left + lstInterfaces.Width
    picInterfaceSplitter.Top = 500
    picInterfaceSplitter.Height = picMiddle.ScaleHeight - picInterfaceSplitter.Top - 100

    lvInterfaces.Left = picInterfaceSplitter.Left + picInterfaceSplitter.Width
    lvInterfaces.Top = 500
    lvInterfaces.Width = picMiddle.ScaleWidth - lvInterfaces.Left - 100
    lvInterfaces.Height = picMiddle.ScaleHeight - lvInterfaces.Top - 100

    ' Transport Protocols

    lstProtocols.Left = 200
    lstProtocols.Top = 500
    lstProtocols.Height = picMiddle.ScaleHeight - lvTransProtocols.Top - 100

    picProtocolSplitter.Left = lstProtocols.Left + lstProtocols.Width
    picProtocolSplitter.Top = 500
    picProtocolSplitter.Height = picMiddle.ScaleHeight - picProtocolSplitter.Top - 100

    lvTransProtocols.Left = picProtocolSplitter.Left + picProtocolSplitter.Width
    lvTransProtocols.Top = 500
    lvTransProtocols.Width = picMiddle.ScaleWidth - lvTransProtocols.Left - 100
    lvTransProtocols.Height = picMiddle.ScaleHeight - lvTransProtocols.Top - 100

    ' Protocol Entries

    picProtocols.Left = 200
    picProtocols.Top = 500
    picProtocols.Height = picMiddle.ScaleHeight - picProtocols.Top - 100
    
    lvProtocols.Left = picProtocols.Left + picProtocols.Width
    lvProtocols.Top = 500
    lvProtocols.Width = picMiddle.ScaleWidth - lvProtocols.Left - 100
    lvProtocols.Height = picMiddle.ScaleHeight - lvProtocols.Top - 100
    
    ' Services
    
    picServices.Left = 200
    picServices.Top = 500
    picServices.Height = picMiddle.ScaleHeight - picServices.Top - 100
    
    lvServices.Left = picServices.Left + picServices.Width
    lvServices.Top = 500
    lvServices.Width = picMiddle.ScaleWidth - lvServices.Left - 100
    lvServices.Height = picMiddle.ScaleHeight - lvServices.Top - 100
    
    Exit Sub
ResizeError:
    lstProtocols.Width = 2000
    lstInterfaces.Width = 2000
    Call picMiddle_Resize
    
End Sub


Private Sub picBottomSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    picSplitter.Left = 0
    picSplitter.Width = Me.Width
    picSplitter.Height = 70

    If picBottomSplitter.Top + y > picHeader.Top + picHeader.Height + 1000 And _
        picBottomSplitter.Top + y < Me.ScaleHeight - 1000 Then
        picSplitter.Top = picBottomSplitter.Top + y
    End If

    picSplitter.Visible = True
End Sub

Private Sub picBottomSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    ' Make sure the splitter keeps within the desired range
    If Button = MouseButtonConstants.vbLeftButton Then
        If picBottomSplitter.Top + y > picHeader.Top + picHeader.Height + 1000 And _
           picBottomSplitter.Top + y < Me.ScaleHeight - 1000 Then
            picSplitter.Top = picBottomSplitter.Top + y
        End If
    End If
    
End Sub

Private Sub picBottomSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picSplitter.Visible = False
    picBottom.Height = Me.ScaleHeight - picSplitter.Top - picBottomSplitter.Height
    picMiddle.Height = Me.ScaleHeight - picHeader.Height - picBottom.Height - picBottomSplitter.Height
    picMiddle.Top = picHeader.Height
End Sub



Private Sub picProtocolSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    picSplitter.Top = picProtocolSplitter.Top + picMiddle.Top
    picSplitter.Height = picProtocolSplitter.Height
    picSplitter.Left = picProtocolSplitter.Left
    picSplitter.Width = 70

    If picProtocolSplitter.Left + x > 1000 And _
        picProtocolSplitter.Left + x < Me.ScaleWidth - 1000 Then
        picSplitter.Left = picProtocolSplitter.Left + x
    End If

    picSplitter.Visible = True
End Sub

Private Sub picProtocolSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Make sure the splitter keeps within the desired range
    If Button = MouseButtonConstants.vbLeftButton Then
        If picProtocolSplitter.Left + x > 1000 And _
           picProtocolSplitter.Left + x < Me.ScaleWidth - 1000 Then
            picSplitter.Left = picProtocolSplitter.Left + x
        End If
    End If
    
End Sub

Private Sub picProtocolSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picSplitter.Visible = False
    picProtocolSplitter.Left = picSplitter.Left
    lstProtocols.Width = picProtocolSplitter.Left - lstProtocols.Left
    lvTransProtocols.Left = picProtocolSplitter.Left + picProtocolSplitter.Width
    lvTransProtocols.Width = Me.ScaleWidth - lvTransProtocols.Left - 100
End Sub



Private Sub picInterfaceSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    picSplitter.Top = picInterfaceSplitter.Top + picMiddle.Top
    picSplitter.Height = picInterfaceSplitter.Height
    picSplitter.Left = picInterfaceSplitter.Left
    picSplitter.Width = 70

    If picInterfaceSplitter.Left + x > 1000 And _
        picInterfaceSplitter.Left + x < Me.ScaleWidth - 1000 Then
        picSplitter.Left = picInterfaceSplitter.Left + x
    End If

    picSplitter.Visible = True
End Sub

Private Sub picInterfaceSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Make sure the splitter keeps within the desired range
    If Button = MouseButtonConstants.vbLeftButton Then
        If picInterfaceSplitter.Left + x > 1000 And _
           picInterfaceSplitter.Left + x < Me.ScaleWidth - 1000 Then
            picSplitter.Left = picInterfaceSplitter.Left + x
        End If
    End If
    
End Sub

Private Sub picInterfaceSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picSplitter.Visible = False
    picInterfaceSplitter.Left = picSplitter.Left
    lstInterfaces.Width = picInterfaceSplitter.Left - lstInterfaces.Left
    lvInterfaces.Left = picInterfaceSplitter.Left + picInterfaceSplitter.Width
    lvInterfaces.Width = Me.ScaleWidth - lvInterfaces.Left - 100
End Sub


Private Sub ProtocolEntries_OnError(Exception As WISPA.CWinsockException)

    ' Add the error to the log
    AddLog txtErrors, "[" & Exception.Source & "] " & Exception.ErrorDescription, vbRed
End Sub

Private Sub Protocols_OnError(Exception As WISPA.CWinsockException)

    ' Add the error to the log
    AddLog txtErrors, "[" & Exception.Source & "] " & Exception.ErrorDescription, vbRed
End Sub

Private Sub ServiceEntries_OnError(Exception As WISPA.CWinsockException)

    ' Add the error to the log
    AddLog txtErrors, "[" & Exception.Source & "] " & Exception.ErrorDescription, vbRed
End Sub

Private Sub Socket_OnError(Exception As WISPA.CWinsockException)

    ' Add the error to the log
    AddLog txtErrors, "[" & Exception.Source & "] " & Exception.ErrorDescription, vbRed
End Sub

Private Sub Startup_OnError(Exception As WISPA.CWinsockException)

    ' Add the error to the log
    AddLog txtErrors, "[" & Exception.Source & "] " & Exception.ErrorDescription, vbRed
End Sub

Private Sub Interfaces_OnError(Exception As WISPA.CWinsockException)

    MsgBox "[" & Exception.Source & "] " & Exception.ErrorDescription

    ' Add the error to the log
    AddLog txtErrors, "[" & Exception.Source & "] " & Exception.ErrorDescription, vbRed
End Sub


Private Sub tsBottom_Click()

    ' Show the correct panel
    Select Case tsBottom.SelectedItem.Key
        Case "info"
            txtErrors.Visible = False
            txtInfo.Visible = True
        Case "error"
            txtInfo.Visible = False
            txtErrors.Visible = True
    End Select

End Sub

Private Sub tsMiddle_Click()

    ' Show the correct panel
    Select Case tsMiddle.SelectedItem.Key
        Case "startup"
            picProtocols.Visible = False
            picServices.Visible = False
            lstInterfaces.Visible = False
            lvInterfaces.Visible = False
            picInterfaceSplitter.Visible = False
            lstProtocols.Visible = False
            picProtocolSplitter.Visible = False
            lvTransProtocols.Visible = False
            lvProtocols.Visible = False
            lvServices.Visible = False
            lvStartup.Visible = True
        Case "transprotocols"
            picProtocols.Visible = False
            picServices.Visible = False
            lstInterfaces.Visible = False
            lvInterfaces.Visible = False
            picInterfaceSplitter.Visible = False
            lstProtocols.Visible = True
            picProtocolSplitter.Visible = True
            lvTransProtocols.Visible = True
            lvProtocols.Visible = False
            lvServices.Visible = False
            lvStartup.Visible = False
        Case "protocols"
            picProtocols.Visible = True
            picServices.Visible = False
            lstInterfaces.Visible = False
            lvInterfaces.Visible = False
            picInterfaceSplitter.Visible = False
            lstProtocols.Visible = False
            picProtocolSplitter.Visible = False
            lvTransProtocols.Visible = False
            lvProtocols.Visible = True
            lvServices.Visible = False
            lvStartup.Visible = False
        Case "services"
            picProtocols.Visible = False
            picServices.Visible = True
            lstInterfaces.Visible = False
            lvInterfaces.Visible = False
            picInterfaceSplitter.Visible = False
            lstProtocols.Visible = False
            picProtocolSplitter.Visible = False
            lvTransProtocols.Visible = False
            lvProtocols.Visible = False
            lvServices.Visible = True
            lvStartup.Visible = False
        Case "interfaces"
            picProtocols.Visible = False
            picServices.Visible = False
            lstInterfaces.Visible = True
            lvInterfaces.Visible = True
            picInterfaceSplitter.Visible = True
            lstProtocols.Visible = False
            picProtocolSplitter.Visible = False
            lvTransProtocols.Visible = False
            lvProtocols.Visible = False
            lvServices.Visible = False
            lvStartup.Visible = False
    End Select

End Sub



Private Sub AddLog(Dest As RichTextBox, Text As String, Optional Colour As ColorConstants)
    Dest.SelStart = Len(Dest.Text)
    If Not IsMissing(Colour) Then Dest.SelColor = Colour
    Dest.SelText = Text & vbCrLf
    Dest.SelStart = Len(Dest.Text)
End Sub
