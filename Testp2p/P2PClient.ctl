VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl P2PClient 
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   ScaleHeight     =   450
   ScaleWidth      =   4830
   ToolboxBitmap   =   "P2PClient.ctx":0000
   Begin P2PProject.Process BAr1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckSend 
      Left            =   4440
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "P2PClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'Default Property Values:
Const m_def_FillColor = vbBlue
'Property Variables:
Dim m_FillColor As Long
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event FileSent()
Public Event FileCancelled()
Event BytesSentSoFare(lngBytesSent As Long, lngTotalToSent As Long)
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event ReadProperties(PropBag As PropertyBag) 'MappingInfo=UserControl,UserControl,-1,ReadProperties
Attribute ReadProperties.VB_Description = "Occurs when a user control or user document is asked to read its data from a file."

Dim blHeaderReceived As Boolean
Dim blConnectionClosed As Boolean
Dim blConnectionConnected As Boolean
Dim isSent As Boolean
Public SendingFile As String

Private Function ShortFileName(StrFilename As String) As String
 If InStr(1, StrFilename, "\", vbTextCompare) Then
  ShortFileName = Mid(StrFilename, InStrRev(StrFilename, "\") + 1)
 Else
  ShortFileName = StrFilename
 End If
End Function
Function SendFile(StrFilename As String)
   
 On Error Resume Next
 
 Do Until blConnectionConnected Or blConnectionClosed
  DoEvents
 Loop

 If blConnectionClosed Then Exit Function
 
 Dim intFreeFile As Integer
 Dim strHeader As String
 Dim LngFileSize As Long
 Dim Data As String * 2048
 Dim LastData As String
 Dim Sentl As Long
 
 blConnectionClosed = False
 LngFileSize = FileLen(StrFilename)
 
 strHeader = ShortFileName(StrFilename) & "#" & LngFileSize
 SendingFile = ShortFileName(StrFilename)
 sckSend.SendData strHeader
 
 Do Until blHeaderReceived Or blConnectionClosed
  DoEvents
 Loop
 
 If blConnectionClosed Then Exit Function
 
 intFreeFile = FreeFile
 
 Open StrFilename For Binary As intFreeFile
  
  
  BAr1.ProcessLineMaxValue = LngFileSize
  
  Do Until EOF(intFreeFile)
   DoEvents
   
   Get #intFreeFile, , Data
   
   Sentl = Sentl + Len(Data)
  
   If (Sentl > LngFileSize) Then
    Sentl = Sentl - Len(Data)
    LastData = Mid(Data, 1, (LngFileSize - Sentl))
    Data = ""
    Sentl = Sentl + Len(LastData)
   End If
  
   DoEvents
   BAr1.ProcessLine_ValueChange Sentl
    
   If sckSend.State <> sckClosed Then
   
     If LastData <> "" Then
       sckSend.SendData LastData
     Else
      sckSend.SendData Data
     End If
   Else
    MsgBox "Connection been cancelled", vbInformation, "P2P Client"
    closeSocket
    RaiseEvent FileCancelled
    Exit Function
   End If
    
    isSent = False
    
    Do Until isSent
     DoEvents
    Loop
    
    RaiseEvent BytesSentSoFare(Sentl, LngFileSize)
    
 Loop
 
 blHeaderReceived = False
 
 Close #intFreeFile
  
 Do Until blHeaderReceived Or blConnectionClosed
  DoEvents
 Loop
 
 
 closeSocket
 
 FileIsSent

End Function

Function closeSocket()
On Error Resume Next
 sckSend.Close
 While sckSend.State <> sckClosed
  DoEvents
 Wend
End Function


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Bar1,Bar1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = BAr1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    BAr1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Bar1,Bar1,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
    FillColor = BAr1.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    BAr1.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Bar1,Bar1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = BAr1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set BAr1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub FileIsSent()
   RaiseEvent FileSent
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=sckSend,sckSend,-1,Close
Public Sub ConnectionClose()
    sckSend.Close
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=sckSend,sckSend,-1,Connect
Public Sub Connect(Optional ByVal RemoteHost As Variant, Optional ByVal RemotePort As Variant)
Attribute Connect.VB_Description = "Connect to the remote computer"
    blConnectionConnected = False
    sckSend.Connect RemoteHost, RemotePort
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    RaiseEvent ReadProperties(PropBag)
    BAr1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    BAr1.FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set BAr1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    sckSend.RemoteHost = PropBag.ReadProperty("RemoteHost", "")
    sckSend.RemotePort = PropBag.ReadProperty("RemotePort", 0)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=sckSend,sckSend,-1,RemoteHost
Public Property Get RemoteHost() As String
Attribute RemoteHost.VB_Description = "Returns/Sets the name used to identify the remote computer"
    RemoteHost = sckSend.RemoteHost
End Property

Public Property Let RemoteHost(ByVal New_RemoteHost As String)
    sckSend.RemoteHost() = New_RemoteHost
    PropertyChanged "RemoteHost"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=sckSend,sckSend,-1,RemoteHostIP
Public Property Get RemoteHostIP() As String
Attribute RemoteHostIP.VB_Description = "Returns the remote host IP address"
    RemoteHostIP = sckSend.RemoteHostIP
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=sckSend,sckSend,-1,RemotePort
Public Property Get RemotePort() As Long
Attribute RemotePort.VB_Description = "Returns/Sets the port to be connected to on the remote computer"
    RemotePort = sckSend.RemotePort
End Property

Public Property Let RemotePort(ByVal New_RemotePort As Long)
    sckSend.RemotePort() = New_RemotePort
    PropertyChanged "RemotePort"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_FillColor = m_def_FillColor
End Sub

Private Sub UserControl_Resize()
 With BAr1
  .Left = 0
  .Top = 0
  .Height = UserControl.Height - 10
  .Width = UserControl.Width - 10
 End With
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", BAr1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("FillColor", BAr1.FillColor, m_def_FillColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", BAr1.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("RemoteHost", sckSend.RemoteHost, "")
    Call PropBag.WriteProperty("RemotePort", sckSend.RemotePort, 0)
End Sub

Private Sub sckSend_Close()
 blConnectionClosed = True
End Sub

Private Sub sckSend_Connect()
 blConnectionConnected = True
End Sub
Private Sub sckSend_DataArrival(ByVal bytesTotal As Long)
 Dim Data As String
   sckSend.GetData Data
   
   If Data = "OK+" Then blHeaderReceived = True
   
   If Data = "NO+" Then
    blHeaderReceived = True
    closeSocket
    RaiseEvent FileCancelled
   End If
   
End Sub

Private Sub sckSend_SendComplete()
 isSent = True
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=sckSend,sckSend,-1,Accept
Public Sub Accept(ByVal requestID As Long)
Attribute Accept.VB_Description = "Accept an incoming connection request"
    sckSend.Accept requestID
End Sub

