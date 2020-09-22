VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P2P Services"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Client Information"
      Height          =   1815
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Width           =   6615
      Begin VB.PictureBox picTray 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   5280
         Picture         =   "Form1.frx":27A2
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSWinsockLib.Winsock sckListen 
         Left            =   5880
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Height          =   320
         Left            =   3740
         Picture         =   "Form1.frx":2AAC
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1320
         Width           =   320
      End
      Begin VB.TextBox txtFilename 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   360
         TabIndex        =   13
         Text            =   "C:\file.txt"
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtRemotePort 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   11
         Text            =   "4100"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtRemoteHost 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   360
         TabIndex        =   9
         Text            =   "127.0.0.1"
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Send &File"
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H009F6539&
         Caption         =   "  Filename"
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
         TabIndex        =   14
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H009F6539&
         Caption         =   "  Remote Port"
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
         Left            =   2400
         TabIndex        =   12
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H009F6539&
         Caption         =   "  Remote Host"
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
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Server Information"
      Height          =   1695
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   6615
      Begin VB.TextBox txtPath 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "C:\Temp"
         Top             =   1200
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Height          =   320
         Left            =   3735
         Picture         =   "Form1.frx":2B60
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1200
         Width           =   320
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Exit"
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtServerPort 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Text            =   "4100"
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Hide"
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Start"
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H009F6539&
         Caption         =   "  Download folder"
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
         TabIndex        =   19
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H009F6539&
         Caption         =   "  Server Port"
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
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.PictureBox picLogo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   6615
      TabIndex        =   0
      Top             =   0
      Width           =   6615
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
         Left            =   1320
         TabIndex        =   2
         Top             =   520
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
         Height          =   870
         Left            =   5520
         Picture         =   "Form1.frx":2C04
         Top             =   120
         Width           =   1125
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
Dim WithEvents MyTray As CTray
Attribute MyTray.VB_VarHelpID = -1

Private Sub Command1_Click()
 txtPath = modSettings.BrowsePath(Me)
End Sub

Private Sub Command2_Click()
 
 If Command2.Caption = "&Start" Then
   ReListen
   Command2.Caption = "&Stop"
 Else
   StopListen
   Command2.Caption = "&Start"
 End If
 
End Sub

Private Sub Command3_Click()
Dim X As New frmSend

With X
  .Show
  .Caption = "Sending file to " & txtRemoteHost
  .lblStatus = "Connection to " & txtRemoteHost
  .P2PClient1.Connect txtRemoteHost, Val(txtRemotePort)
  .P2PClient1.SendFile txtFilename
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
    .hWndParent = Me.hWnd
    
    If .Show(True) Then
      txtFilename = .FileName
    End If
    
  End With
  

End Sub

Function ReListen()
 With sckListen
   .Close
   While .State <> sckClosed
    DoEvents
    .Close
   Wend
   .LocalPort = Val(txtServerPort)
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

Private Sub Form_Load()
 modSettings.LoadSettings Me
 Set MyTray = New CTray
 MyTray.PicBox = picTray
 MyTray.ShowIcon
 MyTray.TipText = "P2PServies"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 MyTray.DeleteIcon
End Sub

Private Sub Form_Terminate()
 modSettings.SaveSettings Me
 MyTray.DeleteIcon
End Sub

Private Sub Form_Unload(Cancel As Integer)
 modSettings.SaveSettings Me
 MyTray.DeleteIcon
End Sub

Private Sub MnuExit_Click()
 MyTray.DeleteIcon
 End
End Sub

Private Sub MnuHide_Click()
 Me.Hide
End Sub

Private Sub MnuShow_Click()
 Me.Show
End Sub

Private Sub MyTray_LButtonDblClick()
 Me.Show
End Sub

Private Sub MyTray_RButtonDown()
 PopupMenu mnuHidden
End Sub

Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)
 Dim X As New frmDownload
 
 With X
  .Show
  .Caption = "Downloading file from " & sckListen.RemoteHost
  .lblStatus = "Waiting for file information "
  .P2PServer1.DownloadFolder = txtPath
  .P2PServer1.Accept requestID
End With

End Sub

