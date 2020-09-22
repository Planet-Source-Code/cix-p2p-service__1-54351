VERSION 5.00
Begin VB.Form frmDownload 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Downloading file"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin P2PProject.P2PServer P2PServer1 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   450
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FillColor       =   10446137
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
      ScaleWidth      =   6210
      TabIndex        =   1
      Top             =   0
      Width           =   6210
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
         TabIndex        =   3
         Top             =   520
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Downloading File"
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
         TabIndex        =   2
         Top             =   240
         Width           =   2490
      End
      Begin VB.Image Image1 
         Height          =   870
         Left            =   5160
         Picture         =   "frmDownload.frx":0000
         Top             =   120
         Width           =   1125
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   120
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
 P2PServer1.closeSocket
 Unload Me
End Sub

Private Sub P2PServer1_BytesDownloadedSoFare(lngBytesDownloaded As Long, lngTotalToDownloaded As Long)
 lblStatus.Caption = "Downloaded " & Int(lngBytesDownloaded / 1000) & " kb out of " & Int(lngTotalToDownloaded / 1000) & " kb "
End Sub

Private Sub P2PServer1_FileCancelled()
 Unload Me
End Sub

Private Sub P2PServer1_FileDownloaded()
 MsgBox "File " & P2PServer1.DownloadFile & " is now downloaded", vbInformation, "P2PServer"
 Unload Me
End Sub
