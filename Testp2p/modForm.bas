Attribute VB_Name = "modForm"
Option Explicit

' Module      : modForms
' Description : Routines for working with VB forms
' Source      : Total VB SourceBook 6
'
Private Declare Function SetWindowLong Lib "user32" _
  Alias "SetWindowLongA" _
  (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) _
  As Long

Private Declare Function SetWindowPos _
  Lib "user32" _
  (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) _
  As Long

Private Declare Function GetSystemMenu _
  Lib "user32" _
  (ByVal hwnd As Long, _
    ByVal bRevert As Long) _
  As Long

Private Declare Function ModifyMenu _
  Lib "user32" _
  Alias "ModifyMenuA" _
  (ByVal hMenu As Long, _
    ByVal nPosition As Long, _
    ByVal wFlags As Long, _
    ByVal wIDNewItem As Long, _
    ByVal lpString As Any) _
  As Long

Private Declare Function GetMenuItemID _
  Lib "user32" _
  (ByVal hMenu As Long, _
    ByVal nPos As Long) _
  As Long

Private Const WM_SYSCOMMAND = &H112
Private Const MOUSE_MOVE = &HF012
Private Const WM_LBUTTONUP = &H202

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type POINTS
  X  As Integer
  Y  As Integer
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOSIZE = 1
Private Const SWP_NOMOVE = 2
Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&
Private Const MF_GRAYED = &H1&
Private Const SC_CLOSE = &HF060
Private Const WS_EX_TRANSPARENT = &H20&
Private Const GWL_EXSTYLE = (-20)

Public Sub CenterForm(frmIn As Form)
  ' Comments  : Centers the form on the screen
  ' Parameters: frmIn - form to center on the screen
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR

  frmIn.Move (Screen.Width - frmIn.Width) / 2, _
    (Screen.Height - frmIn.Height) / 2

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "CenterForm"
  Resume PROC_EXIT

End Sub

Public Sub DisableCloseMenu(frmIn As Form)
  ' Comments  : Grays out the Close item on the form's
  '             system menu
  ' Parameters: frmIn - form to modify
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  Dim lnghMenu As Long
  Dim lnghItem As Long

  On Error GoTo PROC_ERR

  ' get handle to form's system menu
  lnghMenu = GetSystemMenu(frmIn.hwnd, 0)
  
  ' get handle to the 6th item (Close)
  lnghItem = GetMenuItemID(lnghMenu, 6)
  
  ' gray out this item
  lngResult = ModifyMenu( _
    lnghMenu, _
    lnghItem, _
    MF_BYCOMMAND Or MF_GRAYED, _
    -10, _
    "Close")

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "DisableCloseMenu"
  Resume PROC_EXIT

End Sub

Public Sub FormOnTop( _
  frmIn As Form, _
  ByVal fOnTop As Boolean)
  ' Comments  : Sets the form's style to be always on
  '             top, or to remove the always on top style
  ' Parameters: frmIn - the form to modify
  '             fOnTop - true to set the form to be always
  '             on top of other windows. Set to False to
  '             remove this attribute
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim lngState As Long
  
  On Error GoTo PROC_ERR

  If fOnTop Then
    lngState = HWND_TOPMOST
  Else
    lngState = HWND_NOTOPMOST
  End If
  
  SetWindowPos frmIn.hwnd, lngState, 0&, 0&, 0&, 0&, _
    SWP_NOSIZE Or SWP_NOMOVE

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "FormOnTop"
  Resume PROC_EXIT

End Sub

Public Function GetFormNamed(ByVal strForm As String) As Form
  ' Comments  : Retrieves a reference to form based on a string name
  ' Parameters: strForm - the name of the form to locate
  ' Returns   : A reference to the form if found, otherwise Nothing
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR

  Dim frm As Form
  Dim fFound As Boolean
  
  strForm = UCase(strForm)
  
  ' Search in the forms collection to see if a form with the selected name
  ' is found
  For Each frm In Forms
    If UCase(frm.Name) = strForm Then
      fFound = True
      Set GetFormNamed = frm
      Exit For
    End If
  Next frm
  
  If Not fFound Then
    Set GetFormNamed = Nothing
  End If
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetFormNamed"
  Resume PROC_EXIT

End Function

Public Function LoadFormByName(strForm As String) As Form
  ' Comments  : Loads a form by using a string variable containing
  '             the name of the form to avoid hard-coding form references
  ' Parameters: strForm - The name of the form to load
  ' Returns   : A pointer to the form that was loaded
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR

  On Error Resume Next
  Forms.Add strForm
  If Err.Number = 0 Then
    Set LoadFormByName = Forms(Forms.Count - 1)
  Else
    Set LoadFormByName = Nothing
  End If

  On Error GoTo PROC_ERR
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "LoadFormByName"
  Resume PROC_EXIT

End Function

Public Sub MakeTransparent(frmIn As Form)
  ' Comments  : Sets the form's style to be transparent. This call should
  '             be made before the form is shown, for example in the Load
  '             event
  ' Parameters: frmIn - form to modify
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim lngResult As Long
  
  On Error GoTo PROC_ERR
  
  lngResult = SetWindowLong(frmIn.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
  
PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "MakeTransparent"
  Resume PROC_EXIT

End Sub



