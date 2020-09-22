VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4125
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2760
   ControlBox      =   0   'False
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
   ScaleHeight     =   4125
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image2 
      Height          =   195
      Index           =   4
      Left            =   360
      Picture         =   "frmMenu.frx":0000
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   195
   End
   Begin VB.Image Image2 
      Height          =   195
      Index           =   3
      Left            =   360
      Picture         =   "frmMenu.frx":00B3
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   195
   End
   Begin VB.Image Image2 
      Height          =   195
      Index           =   2
      Left            =   360
      Picture         =   "frmMenu.frx":015D
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   195
   End
   Begin VB.Image Image2 
      Height          =   195
      Index           =   1
      Left            =   360
      Picture         =   "frmMenu.frx":0206
      Stretch         =   -1  'True
      Top             =   960
      Width           =   195
   End
   Begin VB.Image Image2 
      Height          =   200
      Index           =   0
      Left            =   360
      Picture         =   "frmMenu.frx":02BB
      Stretch         =   -1  'True
      Top             =   720
      Width           =   200
   End
   Begin VB.Image imgMenu 
      Height          =   240
      Left            =   360
      Picture         =   "frmMenu.frx":0336
      Top             =   420
      Width           =   240
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   2520
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   2520
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   2520
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "        Delete User"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   285
      TabIndex        =   8
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lblMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   285
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   285
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   285
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "        View History   "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   285
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "        View Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   285
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label lblMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "        Send Web-Link"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   285
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "        Send File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   285
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "        Send Message"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   285
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   0
      Picture         =   "frmMenu.frx":05A0
      Top             =   -360
      Width           =   525
   End
   Begin VB.Label lblUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const ColorCode = &H9F6539
Dim MouseOver As Boolean
Private Sub Form_LostFocus()
 Debug.Print "unloaded"
' If MouseOver Then Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim I As Integer
  MouseOver = True
  For I = 0 To lblMenu.UBound
    lblMenu(I).BackColor = vbWhite
    lblMenu(I).ForeColor = vbBlack
  Next I

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblMenu_Click(Index As Integer)
 End
End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim I As Integer
  MouseOver = True
  lblMenu(Index).BackColor = ColorCode
  lblMenu(Index).ForeColor = vbWhite
  For I = 0 To lblMenu.UBound
   If I <> Index Then
    lblMenu(I).BackColor = vbWhite
    lblMenu(I).ForeColor = vbBlack
   End If
  Next I

End Sub

Private Sub lblUser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MouseOver = True
End Sub
