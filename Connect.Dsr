VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   13485
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   15735
   _ExtentX        =   27755
   _ExtentY        =   23786
   _Version        =   393216
   Description     =   "Add-In Project Template"
   DisplayName     =   "Code Inserter"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "None"
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSafe     =   -1  'True
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Public VBInstance As VBIDE.VBE
Private mcbMenuCommandBarCtrl As Object
Public WithEvents ErrorBlockMenuHandler As CommandBarEvents
Attribute ErrorBlockMenuHandler.VB_VarHelpID = -1
Public WithEvents OnErrorMenuHandler As CommandBarEvents
Attribute OnErrorMenuHandler.VB_VarHelpID = -1
Public WithEvents ErrorHandlerMenuHandler As CommandBarEvents
Attribute ErrorHandlerMenuHandler.VB_VarHelpID = -1
'

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
'Event runs when AddIn is added to Instance of VB IDE
On Error GoTo ErrorHandler
   Dim cbMenuCommandBar As Object
   Dim cbCmdBarCtrl As Object
   Dim cbMenu As Object
   '
      'Link To Instance of VB IDE
      Set VBInstance = Application
      'Link to PopUp Menu for Code Editor Window
      Set cbMenu = VBInstance.CommandBars("Code Window")
      '
      If cbMenu Is Nothing Then
         Exit Sub
      End If
      '
      'Add My AddIn to Code Window PopUp Menu
      Set mcbMenuCommandBarCtrl = cbMenu.Controls.Add(10, , , 1)
      With mcbMenuCommandBarCtrl
         .Caption = "Insert Code"
      End With
      '
      'Add Sub-Menu's
      'Add Sub-Menu
      Set cbCmdBarCtrl = mcbMenuCommandBarCtrl.Controls.Add(1)
      'Give menu Item a Caption
      cbCmdBarCtrl.Caption = "Complete Error Block"
      'Link the MenuHandler to the Menu Item
      Set Me.ErrorBlockMenuHandler = VBInstance.Events.CommandBarEvents(cbCmdBarCtrl)
      'Add Sub-Menu
      Set cbCmdBarCtrl = mcbMenuCommandBarCtrl.Controls.Add(1)
      cbCmdBarCtrl.Caption = "On Error Block"
      Set Me.OnErrorMenuHandler = VBInstance.Events.CommandBarEvents(cbCmdBarCtrl)
      'Add Sub-Menu
      Set cbCmdBarCtrl = mcbMenuCommandBarCtrl.Controls.Add(1)
      cbCmdBarCtrl.Caption = "Error Handler Block"
      Set Me.ErrorHandlerMenuHandler = VBInstance.Events.CommandBarEvents(cbCmdBarCtrl)
      '
ExitRoutine:
On Error Resume Next
   Set cbMenuCommandBar = Nothing
   Set cbCmdBarCtrl = Nothing
   Set cbMenu = Nothing
   Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
'Event runs when AddIn is removed from Instance of VB IDE
On Error Resume Next
'
   mcbMenuCommandBarCtrl.Delete
   Set ErrorBlockMenuHandler = Nothing
   Set OnErrorMenuHandler = Nothing
   Set ErrorHandlerMenuHandler = Nothing
   Set mcbMenuCommandBarCtrl = Nothing
   Set VBInstance = Nothing
End Sub

Private Sub ErrorBlockMenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'Event runs when menu item is clicked
On Error Resume Next
   Dim strText As String
   '
      strText = GetStringFromFile("OnError")
      Call VBInstance.ActiveCodePane.CodeModule.InsertLines(FirstLine, strText)
      strText = GetStringFromFile("ErrorHandler")
      Call VBInstance.ActiveCodePane.CodeModule.InsertLines(LastLine, strText)
End Sub

Private Sub ErrorHandlerMenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'Event runs when menu item is clicked
On Error Resume Next
   Dim strText As String
   '
      strText = GetStringFromFile("ErrorHandler")
      Call VBInstance.ActiveCodePane.CodeModule.InsertLines(LastLine, strText)
End Sub

Private Sub OnErrorMenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'Event runs when menu item is clicked
On Error Resume Next
   Dim strText As String
   '
      strText = GetStringFromFile("OnError")
      Call VBInstance.ActiveCodePane.CodeModule.InsertLines(FirstLine, strText)
End Sub

Private Function FirstLine() As Long
'Retrieve First Line number of Procedure
On Error Resume Next
   Dim strProcName As String
   Dim lngProcStartLine As Long
   '
      strProcName = VBInstance.ActiveCodePane.CodeModule.ProcOfLine(CurrentStartLine, vbext_pk_Proc)
      If strProcName = "" Then Exit Function
      lngProcStartLine = VBInstance.ActiveCodePane.CodeModule.ProcBodyLine(strProcName, vbext_pk_Proc)
      FirstLine = lngProcStartLine + 1
End Function

Private Function LastLine() As Long
'Retrieve Last Line number of Procedure
On Error Resume Next
   Dim strProcName As String
   Dim lngProcStartLine As Long
   Dim lngProcLineCount As Long
   '
      strProcName = VBInstance.ActiveCodePane.CodeModule.ProcOfLine(CurrentStartLine, vbext_pk_Proc)
      If strProcName = "" Then Exit Function
      lngProcStartLine = VBInstance.ActiveCodePane.CodeModule.ProcStartLine(strProcName, vbext_pk_Proc)
      lngProcLineCount = VBInstance.ActiveCodePane.CodeModule.ProcCountLines(strProcName, vbext_pk_Proc)
      LastLine = lngProcStartLine + lngProcLineCount - 1
End Function

Private Function CurrentStartLine() As Long
'Retrieve Line number of current cursor position or First Line number of selection if block selected Procedure
On Error Resume Next
   Dim lngStartLine As Long
   Dim lngStartColumn As Long
   Dim lngEndLine As Long
   Dim lngEndColumn As Long
   '
      Call VBInstance.ActiveCodePane.CodeModule.CodePane.GetSelection(lngStartLine, lngStartColumn, lngEndLine, lngEndColumn)
      CurrentStartLine = lngStartLine
End Function

Public Function GetStringFromFile(CodeBlockName As String) As String
'Read the code to insert from file
On Error GoTo ErrorHandler
   Dim strLineText As String
   Dim strText As String
   '
      Open App.Path & "\Code.txt" For Input As #1
      Do While Not EOF(1)
         Line Input #1, strLineText
         If strLineText = "[" & CodeBlockName & "]" Then
            Line Input #1, strLineText
            Do Until Left(strLineText, 1) = "[" And Right(strLineText, 1) = "]"
               strText = strText & strLineText & vbCrLf
               Line Input #1, strLineText
            Loop
            Exit Do
         End If
      Loop
      '
Exit_Routine:
On Error Resume Next
   'Remove vbCrLf from end of string
   GetStringFromFile = Left(strText, Len(strText) - 2)
   Close #1
   Exit Function
   '
ErrorHandler:
   Select Case Err.Number
   Case 62
      'do nothing - End of file error
   Case Else
      MsgBox "An Error has occured." & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "Insert Code Error"
   End Select
   Resume Exit_Routine
End Function
