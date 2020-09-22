Attribute VB_Name = "modSettings"
Option Explicit
' SHOW THE DIALOG FOLDER
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type POINTAPI
        x As Long
        y As Long
End Type


Public Declare Function SHBrowseForFolder Lib _
        "shell32" (lpbi As BrowseInfo) As Long


Public Declare Function SHGetPathFromIDList Lib _
        "shell32" (ByVal pidList As Long, ByVal lpBuffer _
        As String) As Long


Public Declare Function lstrcat Lib "kernel32" _
        Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
        lpString2 As String) As Long


 Type BrowseInfo
        hWndOwner As Long
        pIDLRoot As Long
        pszDisplayName As Long
        lpszTitle As Long
        ulFlags As Long
        lpfnCallback As Long
        lParam As Long
        iImage As Long
  End Type

Function BrowsePath(frm As Form) As String
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    szTitle = "Output Location"


    With tBrowseInfo
        .hWndOwner = frmMain.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = 1 + 2
        
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)


    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        If Not Right(sBuffer, 1) = "\" Then
         BrowsePath = sBuffer & "\"
        Else
         BrowsePath = sBuffer
        End If
    End If
End Function

Function SaveSettings(frm As Form)
  Dim x As Control
  Dim Xframe As Control
  
  For Each x In frm
   DoEvents
   SaveControlInfo x
  Next x
End Function

Function LoadSettings(frm As Form)
  Dim x As Control
  Dim Xframe As Control
  
  For Each x In frm
   DoEvents
    GetControlInfo x
  Next x
End Function


Function SaveControlInfo(x As Control)
 
Dim Obj As Object
Dim strValue

Set Obj = x

 If TypeOf x Is TextBox Then
    strValue = x.Text
 ElseIf TypeOf x Is CheckBox Then
    strValue = x.Value
 ElseIf TypeOf x Is OptionButton Then
    strValue = x.Value
 End If

 If strValue <> "" Then
  SaveSetting App.Title, "settings", x.Name, strValue
 End If
 
End Function

Function GetControlInfo(x As Control)
Dim Obj As Object
Dim strValue
  
  
Set Obj = x

 strValue = GetSetting(App.Title, "settings", Obj.Name)
 If strValue = "" Then Exit Function
 
 Debug.Print x.Name & " - " & Obj.Name
 If TypeOf x Is TextBox Then
    Obj.Text = GetSetting(App.Title, "settings", Obj.Name)
 ElseIf TypeOf x Is CheckBox Then
    Debug.Print GetSetting(App.Title, "settings", Obj.Name)
    Obj.Value = GetSetting(App.Title, "settings", Obj.Name, Unchecked)
 ElseIf TypeOf x Is OptionButton Then
    Obj.Value = GetSetting(App.Title, "settings", Obj.Name, Unchecked)
 End If
  
End Function

