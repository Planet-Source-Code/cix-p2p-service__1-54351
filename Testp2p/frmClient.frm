VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P2P Services"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   3120
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   3120
      TabIndex        =   0
      Top             =   0
      Width           =   3120
      Begin VB.PictureBox picTray 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3840
         Picture         =   "frmClient.frx":27A2
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSWinsockLib.Winsock sckListen 
         Left            =   4080
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Image Image2 
         Height          =   195
         Index           =   0
         Left            =   600
         Picture         =   "frmClient.frx":2AAC
         Stretch         =   -1  'True
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sharing is fun."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009F6539&
         Height          =   240
         Left            =   840
         TabIndex        =   2
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P2P SERVICE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1905
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   2280
         Picture         =   "frmClient.frx":2DF0
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame frameMain 
      BackColor       =   &H00E0E0E0&
      Height          =   4455
      Left            =   0
      TabIndex        =   3
      Top             =   920
      Width           =   3135
      Begin MSComctlLib.ImageList imgOnline 
         Left            =   2160
         Top             =   3600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":4FF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":5348
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":569C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":59F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":5D44
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox picCon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   120
         ScaleHeight     =   3825
         ScaleWidth      =   2865
         TabIndex        =   8
         Top             =   480
         Width           =   2895
         Begin MSComctlLib.TreeView Tree1 
            Height          =   3615
            Left            =   240
            TabIndex        =   9
            Top             =   120
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   6376
            _Version        =   393217
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "imgOnline"
            Appearance      =   0
         End
      End
      Begin VB.Label Label6 
         BackColor       =   &H009F6539&
         Caption         =   "  Online"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frameButtom 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   5280
      Width           =   3135
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Height          =   320
         Left            =   120
         Picture         =   "frmClient.frx":6098
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   320
      End
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "Hidden"
      Visible         =   0   'False
      Begin VB.Menu MnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu MnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu MnuHide 
         Caption         =   "Hide"
      End
      Begin VB.Menu MnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents MyTray As frmSysTray
Attribute MyTray.VB_VarHelpID = -1
Dim XNode As Node
Private Sub Command1_Click()
' txtPath = modSettings.BrowsePath(Me)
End Sub

Private Sub Command2_Click()
 
 'If Command2.Caption = "&Start" Then
 '  ReListen
 '  Command2.Caption = "&Stop"
 'Else
 '  StopListen
'   Command2.Caption = "&Start"
 'End If
 
End Sub

Private Sub Command3_Click()
Dim x As New frmSend

With x
  '.Show
 ' .Caption = "Sending file to " & txtRemoteHost
 ' .lblStatus = "Connection to " & txtRemoteHost
 ' .P2PClient1.Connect txtRemoteHost, Val(txtRemotePort)
 ' .P2PClient1.SendFile txtFilename
End With

End Sub

Private Sub Command4_Click()
  modSettings.SaveSettings Me
 End
End Sub

Private Sub Command5_Click()
  Dim FileDialog As CFileDialog
  Set FileDialog = New CFileDialog
  
  With FileDialog
    .DialogTitle = "P2P Service File Open"
    .Filter = "All Files (*.*)|*.*"
    .FilterIndex = 0
    .Flags = FleFileMustExist + FleHideReadOnly + FleCreatePrompt
    .hWndParent = Me.hwnd
    
'    If .Show(True) Then
'      txtFilename = .FileName
'    End If
    
  End With
  

End Sub

Function ReListen()
 With sckListen
   .Close
   While .State <> sckClosed
    DoEvents
    .Close
   Wend
   '.LocalPort = Val(txtServerPort)
   .Listen
 End With
End Function
Function StopListen()
 With sckListen
   .Close
   While .State <> sckClosed
    DoEvents
    .Close
   Wend
 End With
End Function

Private Sub Command6_Click()
 Me.Hide
End Sub

Private Sub Form_GotFocus()
 Debug.Print "Unload frmMenu"
End Sub

Private Sub Form_Load()
 modSettings.LoadSettings Me
 
 Set MyTray = New frmSysTray
 MyTray.Picture = Me.picTray.Picture
 MyTray.Hide
 MyTray.IconHandle = Me.picTray.Picture.Handle
 
 Dim x As Node
 
 Set x = Tree1.Nodes.Add(, , "mix", "Cintix", 3, 2)
 Set x = Tree1.Nodes.Add(, , "rj", "RjIsinspired", 3, 2)
 Set x = Tree1.Nodes.Add(, , "x", "The-Exterminator", 1, 2)
 Set x = Tree1.Nodes.Add(, , "dt", "Dorthe", 4, 2)
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload MyTray
End Sub

Private Sub Form_Terminate()
 modSettings.SaveSettings Me
Unload MyTray
End Sub

Private Sub Form_Unload(Cancel As Integer)
 modSettings.SaveSettings Me
 Unload MyTray
 
End Sub

Private Sub MyTray_LButtonDblClick()
End Sub

Private Sub MyTray_RButtonDown()
End Sub

Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)
 Dim x As New frmDownload
 
 With x
  .Show
  .Caption = "Downloading file from " & sckListen.RemoteHost
  .lblStatus = "Waiting for file information "
  '.P2PServer1.DownloadFolder = txtPath
  .P2PServer1.Accept requestID
End With

End Sub

Private Sub Tree1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 On Error GoTo Err
 Unload frmMenu
 If Button = 2 And Tree1.SelectedItem.Text <> "" Then
  
  Dim getXY As POINTAPI
  
  GetCursorPos getXY
  
  Load frmMenu
  frmMenu.Move (getXY.x * Screen.TwipsPerPixelX) - (frmMenu.Width), (getXY.y * Screen.TwipsPerPixelY) - 200
  frmMenu.Show
  modForm.FormOnTop frmMenu, True
  frmMenu.lblUser = "             " & Tree1.SelectedItem.Text
  Exit Sub
 End If
 
Err:
Debug.Print Err.Description
End Sub

Private Sub Tree1_NodeClick(ByVal Node As MSComctlLib.Node)
 Set XNode = Node
End Sub
