VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl P2PServer 
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   270
   ScaleWidth      =   4800
   ToolboxBitmap   =   "P2PServer.ctx":0000
   Begin P2PProject.Process Bar1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock SckServer 
      Left            =   4320
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "P2PServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'Default Property Values:
Const m_def_ForeColor = 0
Const m_def_BackStyle = 0
'Property Variables:
Dim m_ForeColor As Long
Dim m_BackStyle As Integer

Dim Header() As String
Dim LngFileSize As Long
Dim TotalGOT As Long
Dim CurrentFileNr As Integer

Public Event FileDownloaded()
Public Event FileCancelled()
Public Event BytesDownloadedSoFare(lngBytesDownloaded As Long, lngTotalToDownloaded As Long)
Public DownloadFile As String
Public DownloadFolder As String
Dim isSent As Boolean
Dim intCount

Function closeSocket()
 SckServer.Close
 While SckServer.State <> sckClosed
  DoEvents
 Wend
End Function

Private Sub DoneDownloading()
 RaiseEvent FileDownloaded
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Bar1,Bar1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = Bar1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Bar1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
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
    Set Font = Bar1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Bar1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
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
'MappingInfo=Bar1,Bar1,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
    FillColor = Bar1.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    Bar1.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=SckServer,SckServer,-1,LocalPort
Public Property Get LocalPort() As Long
Attribute LocalPort.VB_Description = "Returns/Sets the port used on the local computer"
    LocalPort = SckServer.LocalPort
End Property

Public Property Let LocalPort(ByVal New_LocalPort As Long)
    SckServer.LocalPort() = New_LocalPort
    PropertyChanged "LocalPort"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=SckServer,SckServer,-1,Listen
Public Sub Listen()
Attribute Listen.VB_Description = "Listen for incoming connection requests"
    SckServer.Listen
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=SckServer,SckServer,-1,RemoteHost
Public Property Get RemoteHost() As String
Attribute RemoteHost.VB_Description = "Returns/Sets the name used to identify the remote computer"
    RemoteHost = SckServer.RemoteHost
End Property

Public Property Let RemoteHost(ByVal New_RemoteHost As String)
    SckServer.RemoteHost() = New_RemoteHost
    PropertyChanged "RemoteHost"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=SckServer,SckServer,-1,RemoteHostIP
Public Property Get RemoteHostIP() As String
Attribute RemoteHostIP.VB_Description = "Returns the remote host IP address"
    RemoteHostIP = SckServer.RemoteHostIP
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ForeColor = m_def_ForeColor
    m_BackStyle = m_def_BackStyle
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Bar1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Bar1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Bar1.FillColor = PropBag.ReadProperty("FillColor", &HF7DBD7)
    SckServer.LocalPort = PropBag.ReadProperty("LocalPort", 0)
    SckServer.RemoteHost = PropBag.ReadProperty("RemoteHost", "")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", Bar1.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Bar1.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("FillColor", Bar1.FillColor, &HF7DBD7)
    Call PropBag.WriteProperty("LocalPort", SckServer.LocalPort, 0)
    Call PropBag.WriteProperty("RemoteHost", SckServer.RemoteHost, "")
End Sub

Private Sub SckServer_DataArrival(ByVal bytesTotal As Long)
 
   Dim Data As String
   If intCount < 2 Then intCount = intCount + 1
  
  If intCount < 2 Then
   SckServer.GetData Data, vbString
   Header = Split(Data, "#")
   LngFileSize = Val(Header(1))
   DownloadFile = Header(0)
   
   Dim intReply  As Integer
   
   intReply = MsgBox("Do you accept the file [" & Header(0) & "] Size [" & Header(1) & "]", vbQuestion + vbYesNo, "Incoming file")
   
   If intReply = vbYes Then
     
     Bar1.ProcessLineMaxValue = LngFileSize
     SckServer.SendData "OK+"
     CurrentFileNr = FreeFile
     Open DownloadFolder & DownloadFile For Binary As #CurrentFileNr
       
   Else
     
     SckServer.SendData "NO+"
     isSent = False
  
     Do Until isSent
      DoEvents
     Loop
     closeSocket
     RaiseEvent FileCancelled
     Exit Sub
   End If
  Else
   
   DoEvents
   
   If SckServer.State <> sckClosed Then
    SckServer.GetData Data
   Else
    MsgBox "Connection been cancelled", vbInformation, "P2P Server"
    closeSocket
    intCount = 0
    RaiseEvent FileCancelled
    Exit Sub
   End If
   
   
   TotalGOT = TotalGOT + Len(Data)
   Bar1.ProcessLine_ValueChange TotalGOT
   RaiseEvent BytesDownloadedSoFare(TotalGOT, LngFileSize)
   
   Put #CurrentFileNr, , Data
   
   If TotalGOT = LngFileSize Then
    Close #CurrentFileNr
    SckServer.SendData "OK+"
    
    isSent = False
  
    Do Until isSent
     DoEvents
    Loop
    
    closeSocket
    intCount = 0
    DoneDownloading
    
    Exit Sub
   
   End If
  
   
  End If
  
End Sub
Private Sub SckServer_ConnectionRequest(ByVal requestID As Long)
 If SckServer.State <> sckClosed Then closeSocket
 SckServer.Accept requestID
End Sub
Private Sub sckServer_SendComplete()
 isSent = True
End Sub

Private Sub UserControl_Resize()
 With Bar1
  .Left = 0
  .Top = 0
  .Height = UserControl.Height - 10
  .Width = UserControl.Width - 10
 End With
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=SckServer,SckServer,-1,Accept
Public Sub Accept(ByVal requestID As Long)
Attribute Accept.VB_Description = "Accept an incoming connection request"
    SckServer.Accept requestID
End Sub

