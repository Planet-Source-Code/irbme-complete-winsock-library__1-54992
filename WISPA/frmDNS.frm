VERSION 5.00
Begin VB.Form frmDNS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "W I S P A (Copyright (c) Chris Waddell) - DNS"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4920
   Icon            =   "frmDNS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMain 
      Height          =   3345
      Left            =   0
      ScaleHeight     =   3285
      ScaleWidth      =   4875
      TabIndex        =   4
      Top             =   1275
      Width           =   4935
      Begin VB.CheckBox chkAsync 
         Caption         =   "Asynchronous"
         Height          =   225
         Left            =   1365
         TabIndex        =   14
         Top             =   630
         Value           =   1  'Checked
         Width           =   3270
      End
      Begin VB.TextBox txtDomainName 
         Height          =   285
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2835
         Width           =   3480
      End
      Begin VB.ListBox lstIPs 
         Height          =   840
         Left            =   1365
         TabIndex        =   10
         Top             =   1890
         Width           =   3480
      End
      Begin VB.ListBox lstAliases 
         Height          =   840
         Left            =   1365
         TabIndex        =   8
         Top             =   1050
         Width           =   3480
      End
      Begin VB.CommandButton cmdResolve 
         Caption         =   "&Resolve"
         Height          =   330
         Left            =   105
         TabIndex        =   7
         Top             =   630
         Width           =   1065
      End
      Begin VB.TextBox txtHostName 
         Height          =   285
         Left            =   1365
         TabIndex        =   6
         Text            =   "www.microsoft.com"
         Top             =   210
         Width           =   3480
      End
      Begin VB.Label Label4 
         Caption         =   "Fully Qualified Domain Name:"
         Height          =   435
         Left            =   105
         TabIndex        =   12
         Top             =   2835
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "IPs:"
         Height          =   330
         Left            =   105
         TabIndex        =   11
         Top             =   1890
         Width           =   1485
      End
      Begin VB.Label Label2 
         Caption         =   "Aliases:"
         Height          =   330
         Left            =   105
         TabIndex        =   9
         Top             =   1050
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Host Name or IP:"
         Height          =   225
         Left            =   105
         TabIndex        =   5
         Top             =   210
         Width           =   1275
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   4920
      TabIndex        =   0
      Top             =   0
      Width           =   4920
      Begin VB.PictureBox picLogo 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2610
         Left            =   6825
         Picture         =   "frmDNS.frx":030A
         ScaleHeight     =   174
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   140
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   2100
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
         TabIndex        =   3
         Top             =   210
         Width           =   3795
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
   End
   Begin VB.Menu mnuPop 
      Caption         =   "mnuPop"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyAddr 
         Caption         =   "Copy Address"
      End
      Begin VB.Menu mnuCopyIP 
         Caption         =   "Copy IP"
      End
   End
End
Attribute VB_Name = "frmDNS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean


Private WithEvents Startup As CWinsockStartup
Attribute Startup.VB_VarHelpID = -1
Private WithEvents HostResolve As CDns
Attribute HostResolve.VB_VarHelpID = -1
Private Exception As CWinsockException
Attribute Exception.VB_VarHelpID = -1


Private Sub cmdResolve_Click()

  Dim DnsResult As CDnsResult

    lstAliases.Clear
    lstIPs.Clear
    txtDomainName.Text = vbNullString

    Set DnsResult = HostResolve.Resolve(txtHostName.Text, IIf(chkAsync.Value = vbChecked, True, False))

    If chkAsync.Value <> vbChecked Then
    
        If DnsResult Is Nothing Then
            Set Exception = New CWinsockException
            Exception.Source = "Resolve"
            Exception.Display
            Set Exception = Nothing
        Else
            HostResolve_HostResolved DnsResult
        End If
    End If

End Sub

Private Sub Form_Load()
    Set Startup = New CWinsockStartup
    Set HostResolve = New CDns
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set HostResolve = Nothing
    Set Startup = Nothing
End Sub


Private Sub HostResolve_HostResolved(HostEntry As WISPA.CDnsResult)
    
  Dim i As Integer
    
    For i = 1 To HostEntry.Aliases.Count
        lstAliases.AddItem HostEntry.Aliases.Item(i)
    Next i
    
    For i = 1 To HostEntry.AddressList.Count
        lstIPs.AddItem HostEntry.AddressList.Item(i).StringAddress
    Next i
    
    txtDomainName.Text = HostEntry.FullyQualifiedDomainName
    
End Sub

Private Sub HostResolve_OnError(Exception As WISPA.CWinsockException)
    Exception.Display
End Sub

Private Sub lstAliases_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 
    If Button = MouseButtonConstants.vbRightButton Then
        mnuCopyAddr.Visible = True
        mnuCopyIP.Visible = False
        PopupMenu mnuPop, , x + lstAliases.Left + picMain.Left, y + lstAliases.Top + picMain.Top
    End If
    
End Sub

Private Sub lstIPs_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 
    If Button = MouseButtonConstants.vbRightButton Then
        mnuCopyIP.Visible = True
        mnuCopyAddr.Visible = False
        PopupMenu mnuPop, , x + lstIPs.Left + picMain.Left, y + lstIPs.Top + picMain.Top
    End If
    
End Sub

Private Sub mnuCopyAddr_Click()
    Clipboard.SetText lstAliases.List(lstAliases.ListIndex)
End Sub

Private Sub mnuCopyIP_Click()
    Clipboard.SetText lstIPs.List(lstIPs.ListIndex)
End Sub

Private Sub picHeader_Paint()
    ' Do the gradient fill and draw the logo
    GradientFill picHeader, True, RGB(150, 150, 150), RGB(250, 250, 250)
    TransparentBlt picHeader.hdc, 10, 10, 80, 70, picLogo.hdc, 0, 0, picLogo.ScaleWidth, picLogo.ScaleHeight, vbMagenta
End Sub

Private Sub Startup_OnError(Exception As WISPA.CWinsockException)
    Exception.Display
    Unload Me
End Sub
