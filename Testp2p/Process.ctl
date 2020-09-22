VERSION 5.00
Begin VB.UserControl Process 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7980
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   7980
   ToolboxBitmap   =   "Process.ctx":0000
   Begin VB.PictureBox MinBAr 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000040&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Index           =   0
      Left            =   240
      ScaleHeight     =   1665
      ScaleWidth      =   7425
      TabIndex        =   0
      Top             =   1080
      Width           =   7455
      Begin VB.PictureBox MinBAr 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00000040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   975
         Index           =   1
         Left            =   360
         ScaleHeight     =   975
         ScaleWidth      =   6615
         TabIndex        =   1
         Top             =   360
         Width           =   6615
      End
   End
End
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ProcessLineMaxValue As Variant
Function ProcessLine_ValueChange(ByVal ProcessValue As Variant) As String

Dim PicLen As Variant
Dim Procents As Variant
Dim PicDrawWidth As Variant

MinBAr(1).ZOrder
PicLen = MinBAr(0).ScaleWidth

PicDrawWidth = PicLen / 100

Procents = ProcessValue * 100
Procents = Procents / ProcessLineMaxValue
Procents = Int(Procents)
ProcessLine_ValueChange = Procents & "%"

If Procents >= 100 Or ProcessValue = ProcessLineMaxValue Then
 MinBAr(1).Move 0, 0
 MinBAr(1).Cls
 MinBAr(0).Cls
 
 If Procents > 100 Then
  MinBAr(1).CurrentX = ((MinBAr(1).ScaleWidth - MinBAr(1).TextWidth("- OVER 100% (" & Procents & ") -")) / 2) - (Procents * PicDrawWidth)
  MinBAr(1).CurrentY = ((MinBAr(1).ScaleHeight - MinBAr(1).TextHeight("- OVER 100% (" & Procents & ") -")) / 2)
  MinBAr(1).Move Procents * PicDrawWidth, 0
  MinBAr(1).Print "- OVER 100% (" & Procents & ") -"

  MinBAr(0).CurrentX = (MinBAr(0).ScaleWidth - MinBAr(0).TextWidth("- OVER 100% (" & Procents & ") -")) / 2
  MinBAr(0).CurrentY = (MinBAr(0).ScaleHeight - MinBAr(0).TextHeight("- OVER 100% (" & Procents & ") -")) / 2
  MinBAr(0).Print "- OVER 100% (" & Procents & ") -"
 End If
 Exit Function
End If

 MinBAr(1).Cls
 MinBAr(0).Cls

MinBAr(1).CurrentX = ((MinBAr(1).ScaleWidth - MinBAr(1).TextWidth(Procents & "%")) / 2) - (Procents * PicDrawWidth)
MinBAr(1).CurrentY = ((MinBAr(1).ScaleHeight - MinBAr(1).TextHeight(Procents & "%")) / 2)
MinBAr(1).Move Procents * PicDrawWidth, 0
MinBAr(1).Print Procents & "%"

MinBAr(0).CurrentX = (MinBAr(0).ScaleWidth - MinBAr(0).TextWidth(Procents & "%")) / 2
MinBAr(0).CurrentY = (MinBAr(0).ScaleHeight - MinBAr(0).TextHeight(Procents & "%")) / 2
MinBAr(0).Print Procents & "%"

 MinBAr(1).Refresh
 MinBAr(0).Refresh
End Function

Private Sub UserControl_Resize()
 
 MinBAr(0).Left = 0
 MinBAr(0).Top = 0
 
 MinBAr(1).Left = 0
 MinBAr(1).Top = 0
 
 MinBAr(0).Width = UserControl.Width
 MinBAr(0).Height = UserControl.Height
 
 MinBAr(1).Width = UserControl.Width
 MinBAr(1).Height = UserControl.Height

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MinBAr(0),MinBAr,0,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = MinBAr(1).BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    
    MinBAr(0).FillColor = New_BackColor
    MinBAr(0).ForeColor = New_BackColor
    MinBAr(1).BackColor = New_BackColor
    
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MinBAr(1),MinBAr,1,FillColor
Public Property Get FillColor() As OLE_COLOR
    FillColor = MinBAr(1).FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)

    MinBAr(1).FillColor = New_FillColor
    MinBAr(1).ForeColor = New_FillColor
    MinBAr(0).BackColor = New_FillColor
  
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MinBAr(0),MinBAr,0,Font
Public Property Get Font() As Font
    Set Font = MinBAr(0).Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set MinBAr(0).Font = New_Font
    Set MinBAr(1).Font = New_Font
    
    PropertyChanged "Font"
    
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  
    
    MinBAr(0).FillColor = PropBag.ReadProperty("BackColor", &H40&)
    MinBAr(0).ForeColor = PropBag.ReadProperty("BackColor", &H40&)
    MinBAr(0).BackColor = PropBag.ReadProperty("FillColor", &H40&)
    
    MinBAr(1).BackColor = PropBag.ReadProperty("BackColor", &H40&)
    MinBAr(1).FillColor = PropBag.ReadProperty("FillColor", &H40&)
    MinBAr(1).ForeColor = PropBag.ReadProperty("FillColor", &H40&)
    
    Set MinBAr(0).Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set MinBAr(1).Font = PropBag.ReadProperty("Font", Ambient.Font)
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", MinBAr(1).BackColor, &H40&)
    Call PropBag.WriteProperty("FillColor", MinBAr(1).FillColor, &H40&)
    Call PropBag.WriteProperty("Font", MinBAr(0).Font, Ambient.Font)
End Sub


