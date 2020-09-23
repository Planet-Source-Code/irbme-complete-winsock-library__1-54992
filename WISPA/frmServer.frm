VERSION 5.00
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "W I S P A (Copyright (c) Chris Waddell) - Simple Server"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6030
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMain 
      Height          =   3345
      Left            =   0
      ScaleHeight     =   3285
      ScaleWidth      =   5820
      TabIndex        =   4
      Top             =   1275
      Width           =   5880
      Begin VB.ListBox lstClients 
         Height          =   1425
         Left            =   105
         TabIndex        =   9
         Top             =   630
         Width           =   5580
      End
      Begin VB.CommandButton cmdListen 
         Caption         =   "Listen"
         Height          =   330
         Left            =   4725
         TabIndex        =   8
         Top             =   105
         Width           =   960
      End
      Begin VB.TextBox txtPort 
         Height          =   330
         Left            =   3570
         MaxLength       =   5
         TabIndex        =   7
         Text            =   "4000"
         Top             =   105
         Width           =   1065
      End
      Begin VB.ComboBox cmbInterfaces 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   105
         Width           =   2640
      End
      Begin VB.Label Label1 
         Caption         =   "Bind to:"
         Height          =   225
         Left            =   105
         TabIndex        =   6
         Top             =   105
         Width           =   960
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   6030
      TabIndex        =   0
      Top             =   0
      Width           =   6030
      Begin VB.PictureBox picLogo 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2610
         Left            =   6825
         Picture         =   "frmServer.frx":030A
         ScaleHeight     =   174
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   140
         TabIndex        =   1
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   210
         Width           =   3795
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Private WithEvents Startup      As CWinsockStartup
Attribute Startup.VB_VarHelpID = -1
Private WithEvents Interfaces   As CInterfaceEnumeration
Attribute Interfaces.VB_VarHelpID = -1
Private WithEvents ServerSocket As CSocket
Attribute ServerSocket.VB_VarHelpID = -1
Private WithEvents SmartSender  As CSmartSender
Attribute SmartSender.VB_VarHelpID = -1

Private WithEvents ClientSocket As CSocket
Attribute ClientSocket.VB_VarHelpID = -1


Private Sub ClientSocket_Closed()
    lstClients.Clear
End Sub

Private Sub ClientSocket_DataArrived()
  Dim Data() As Byte
    ClientSocket.Recieve Data
    Debug.Print StrConv(Data, vbUnicode)
End Sub

Private Sub ClientSocket_OnError(Exception As WISPA.CWinsockException)
    Exception.Display
    Unload Me
End Sub

Private Sub cmdListen_Click()
  
  Dim BindAddress As CIP4Address

    If cmdListen.Caption = "Listen" Then
        Set BindAddress = New CIP4Address
        BindAddress.StringAddress = cmbInterfaces.List(cmbInterfaces.ListIndex)
        
        ServerSocket.OpenSocket AddFam_InterNetwork, SockType_stream, Proto_Tcp, True
        
        ServerSocket.Bind BindAddress, txtPort.Text
        ServerSocket.listen
        cmdListen.Caption = "Stop"
    Else
        ServerSocket.CloseSocket
        cmdListen.Caption = "Listen"
    End If
        
End Sub


Private Sub Form_Load()

  Dim i As Integer

    Set Startup = New CWinsockStartup
    Set Interfaces = New CInterfaceEnumeration

    Interfaces.Initialize
    
    cmbInterfaces.AddItem "0.0.0.0"
    
    For i = 1 To Interfaces.Count
        cmbInterfaces.AddItem Interfaces.Item(i).Address.IPAddress.StringAddress
    Next i
    
    cmbInterfaces.ListIndex = 0
    
    Set ServerSocket = New CSocket

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set SmartSender = Nothing
    If Not ClientSocket Is Nothing Then
        If ClientSocket.SocketHandle > 0 Then ClientSocket.CloseSocket
    End If
    If ServerSocket.SocketHandle > 0 Then ServerSocket.CloseSocket
    Set ServerSocket = Nothing
    Set Interfaces = Nothing
    Set Startup = Nothing
End Sub

Private Sub picHeader_Paint()
    ' Do the gradient fill and draw the logo
    GradientFill picHeader, True, RGB(150, 150, 150), RGB(250, 250, 250)
    TransparentBlt picHeader.hdc, 10, 10, 80, 70, picLogo.hdc, 0, 0, picLogo.ScaleWidth, picLogo.ScaleHeight, vbMagenta
End Sub

Private Sub Interfaces_OnError(Exception As WISPA.CWinsockException)
    Exception.Display
    Unload Me
End Sub

Private Sub ServerSocket_ConnectionRequest()

  Dim Address As CIP4Address
  Dim Port As Long

    Set ClientSocket = ServerSocket.Accept(Address, Port)
    lstClients.AddItem CStr(ClientSocket.SocketHandle) & " " & Address.StringAddress & ":" & CStr(Port)
    
    Set SmartSender = New CSmartSender
    Set SmartSender.SocketObject = ClientSocket
    
End Sub

Private Sub ServerSocket_OnError(Exception As WISPA.CWinsockException)
    Exception.Display
    Unload Me
End Sub

Private Sub Startup_OnError(Exception As WISPA.CWinsockException)
    Exception.Display
    Unload Me
End Sub

