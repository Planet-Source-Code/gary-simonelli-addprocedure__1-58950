VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Procedure/ Error Handler"
   ClientHeight    =   3990
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   7230
   Icon            =   "frmAddIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Properties "
      Height          =   1425
      Left            =   3690
      TabIndex        =   23
      Top             =   1530
      Width           =   1755
      Begin VB.ListBox lstPropTypes 
         Height          =   960
         ItemData        =   "frmAddIn.frx":0442
         Left            =   150
         List            =   "frmAddIn.frx":044F
         Style           =   1  'Checkbox
         TabIndex        =   25
         Top             =   270
         Width           =   1425
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Structure"
      Height          =   735
      Left            =   3690
      TabIndex        =   21
      Top             =   780
      Width           =   1755
      Begin VB.TextBox txtStructure 
         Height          =   285
         Left            =   150
         TabIndex        =   22
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      ItemData        =   "frmAddIn.frx":0462
      Left            =   3690
      List            =   "frmAddIn.frx":0478
      TabIndex        =   19
      Text            =   "String"
      Top             =   390
      Width           =   1755
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   17
      Top             =   3705
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Type:"
            TextSave        =   "Type:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Instance:"
            TextSave        =   "Instance:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Module:"
            TextSave        =   "Module:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5001
            Text            =   "Version:"
            TextSave        =   "Version:"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Caption         =   "Module"
      Height          =   705
      Left            =   150
      TabIndex        =   14
      Top             =   2250
      Width           =   3525
      Begin VB.OptionButton optClass 
         Caption         =   "Class"
         Height          =   195
         Left            =   2340
         TabIndex        =   16
         Top             =   300
         Width           =   1125
      End
      Begin VB.OptionButton optForm 
         Caption         =   "Form/Module"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   300
         Value           =   -1  'True
         Width           =   1305
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Version"
      Height          =   645
      Left            =   150
      TabIndex        =   13
      Top             =   2970
      Width           =   5295
      Begin VB.OptionButton optStand 
         Caption         =   "Standalone"
         Height          =   195
         Left            =   2340
         TabIndex        =   6
         Top             =   270
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optDLL 
         Caption         =   "DLL Version"
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.TextBox txtRoutineName 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   390
      Width           =   3525
   End
   Begin VB.Frame Frame2 
      Caption         =   "Scope"
      Height          =   705
      Left            =   150
      TabIndex        =   11
      Top             =   1530
      Width           =   3525
      Begin VB.OptionButton optFriend 
         Caption         =   "Friend"
         Height          =   195
         Left            =   2370
         TabIndex        =   24
         Top             =   330
         Width           =   855
      End
      Begin VB.OptionButton optPublic 
         Caption         =   "Public"
         Height          =   195
         Left            =   1230
         TabIndex        =   4
         Top             =   330
         Width           =   1215
      End
      Begin VB.OptionButton optPrivate 
         Caption         =   "Private"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Procedure Type"
      Height          =   735
      Left            =   150
      TabIndex        =   10
      Top             =   780
      Width           =   3525
      Begin VB.OptionButton optProperty 
         Caption         =   "Property"
         Height          =   195
         Left            =   2340
         TabIndex        =   18
         Top             =   330
         Width           =   1095
      End
      Begin VB.OptionButton optFunction 
         Caption         =   "Function"
         Height          =   255
         Left            =   1170
         TabIndex        =   2
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton optSub 
         Caption         =   "Sub"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   5790
      TabIndex        =   9
      Top             =   1380
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5790
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5790
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Type"
      Height          =   225
      Left            =   3690
      TabIndex        =   20
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Routine Name"
      Height          =   225
      Left            =   150
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'Component Name : frmAddIn
'Author         : Gary Simonelli
'Created        : Wednesday, July 31, 2002
'Version        : VB 6.0
'Description    :
'*******************************************************************************
'Modified:             Date:                 By:
'*******************************************************************************
Option Explicit

Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private Type udtInstance
    insPrivate As Boolean
    insPublic  As Boolean
    insFriend  As Boolean
End Type

Private Type udtOptionButtons
    ModuleType          As Boolean
    Sub_Routine         As Boolean
    Function_Routine    As Boolean
    Property_Routine    As Boolean
    Instance            As udtInstance
    ErrCode             As String
    Version             As Boolean
End Type

Private m_udtOptionButtons      As udtOptionButtons

Private Type udtPropType
    lstGet As Boolean
    lstLet As Boolean
    lstSet As Boolean
End Type

Private m_udtPropType As udtPropType


Private Const IDX_FUNCTION      As Integer = 0
Private Const IDX_SUB           As Integer = 1
Private Const IDX_PROPERTY      As Integer = 2
Private Const INST_PRIVATE      As Boolean = True
Private Const INST_PUBLIC       As Boolean = False
Private Const INST_FRIEND       As Boolean = False
Private Const VER_DLL           As Boolean = True
Private Const VER_STANDALONE    As Boolean = False
Private Const MODULE_TYPE_FORM  As Boolean = True
Private Const MODULE_TYPE_CLASS As Boolean = False


Private m_bDirty                As Boolean

Private Sub CancelButton_Click()
    Unload Me
End Sub

'**********************************************************************
'Procedure   : Private Method cmdApply_Click
'Created on  : Wednesday, July 31, 2002
'Description :
'**********************************************************************
Private Sub cmdApply_Click()
    
    Const PROCEDURE_NAME As String = "cmdApply_Click"
    On Error GoTo ErrHandler
    
    Dim rtn
    Dim Ln          As Long
    Dim lCodePane   As Long
    Dim sCodePane   As String
    Dim CurrLine    As Long
    Dim vCurrLine   As Variant
    Dim sItems      As String
    Dim bFound      As Boolean
    Dim lStartLn    As Long
    Dim lStartCol   As Long
    Dim lEndCol     As Long
    
    Dim lEndLn      As Long
    Dim iLn         As Integer
    Dim sList       As String
    Dim cnt         As Integer
    
    bFound = False
    
    If Len(m_udtOptionButtons.ErrCode) > 0 Then
        If optProperty.Value Then
            '//some flag
            If Len(txtStructure.Text) > 0 Then
                VBInstance.ActiveCodePane.GetSelection lStartLn, lStartCol, lEndLn, lEndCol
                sItems = VBInstance.ActiveCodePane.CodeModule.Lines(lStartLn, (lEndLn - lStartLn) + 1)
                bFound = True
                For iLn = 1 To UBound(Split(sItems, vbCrLf))
                    sList = Trim$(Split(sItems, vbCrLf)(iLn))
                    If sList <> "End Type" Then
                        txtRoutineName.Text = Split(sList, Chr$(32))(0)
                        cnt = UBound(Split(sList, Chr$(32)))
                        cboType.Text = Split(sList, Chr$(32))(cnt)
                        
                        Routine IDX_PROPERTY
                    
                        '>>> set the code pane to the current one displayed <<<
                        sCodePane = VBInstance.ActiveCodePane.CodeModule
                        VBInstance.ActiveCodePane.CodeModule = sCodePane
                        '>>> insert the code into the project <<<
                        lCodePane = VBInstance.ActiveCodePane.CodePaneView
                        
                        vCurrLine = (VBInstance.ActiveCodePane.CodeModule.CountOfLines + 1)
                        
                        CurrLine = vCurrLine '(CurrentLine)
                        VBInstance.ActiveCodePane.CodeModule.InsertLines CurrLine, m_udtOptionButtons.ErrCode
                        m_bDirty = False
                    End If
                Next
                txtRoutineName.Text = vbNullString
                cboType.Text = vbNullString
            End If
        End If
        
        If Not bFound Then
            Select Case True
            Case optFunction.Value
                Routine IDX_FUNCTION
            Case optSub.Value
                Routine IDX_SUB
            Case optProperty.Value
                Routine IDX_PROPERTY
            End Select
        
            If m_bDirty And Len(txtRoutineName.Text) > 0 Then
                If VBInstance.ActiveCodePane Is Nothing Then
                    MsgBox "No active code window.", vbExclamation, App.Title
                Else
                    '>>> set the code pane to the current one displayed <<<
                    sCodePane = VBInstance.ActiveCodePane.CodeModule
                    VBInstance.ActiveCodePane.CodeModule = sCodePane
                    
                    '>>> insert the code into the project <<<
                    lCodePane = VBInstance.ActiveCodePane.CodePaneView
                    
                    CurrLine = CurrentLine
                    
                    VBInstance.ActiveCodePane.CodeModule.InsertLines CurrLine, m_udtOptionButtons.ErrCode
                    m_bDirty = False
                End If
            End If
        End If
    End If
    
Procedure_Exit:
    On Error GoTo 0
    Exit Sub
ErrHandler:
   Select Case Err
   Case Else
       Select Case MsgBox("[" & Err.Number & "]" & Err.Description, vbCritical + vbAbortRetryIgnore, App.Title)
       Case vbAbort: GoTo Procedure_Exit
       Case vbRetry: Resume 0
       Case vbIgnore: Resume Next
       End Select
   End Select
End Sub

Private Function CurrentLine() As Long
    
    ' Return the current line of the
    ' active code pane.
    
    Dim lCurrentLine     As Long
    Dim lCurrentColumn   As Long
    Dim lEndLine         As Long
    Dim lEndColumn       As Long
    
    On Error Resume Next
    VBInstance.ActiveCodePane.GetSelection lCurrentLine, lCurrentColumn, lEndLine, lEndColumn
    CurrentLine = lEndLine
    
End Function

Private Sub Form_Load()

    On Error Resume Next

    '>>> Set  default values <<<
    optSub.Value = True
    optPrivate.Value = True
    m_udtOptionButtons.Instance.insPrivate = INST_PRIVATE
    m_udtOptionButtons.Sub_Routine = True
    m_udtOptionButtons.Version = VER_STANDALONE
    m_udtOptionButtons.ModuleType = MODULE_TYPE_FORM
    
    Update_Statusbar 1, "SUBROUTINE"
    Update_Statusbar 2, "PRIVATE"
    Update_Statusbar 3, "FORM/MODULE"
    Update_Statusbar 4, "STANDALONE"
    
    Routine IDX_SUB
     
    lstPropTypes.Selected(0) = True
    lstPropTypes.Selected(1) = True
    lstPropTypes.Selected(2) = False
     
    m_bDirty = True
    
End Sub

Private Sub lstPropTypes_ItemCheck(Item As Integer)
    
    With m_udtPropType
        Select Case Item
        Case 0
            .lstGet = lstPropTypes.Selected(Item)
        Case 1
            .lstLet = lstPropTypes.Selected(Item)
        Case 2
            .lstSet = lstPropTypes.Selected(Item)
        End Select
    End With
End Sub

Private Sub OKButton_Click()
    
    Call cmdApply_Click
    Unload Me
    'Connect.Hide

End Sub

Private Sub optClass_Click()

    m_udtOptionButtons.ModuleType = MODULE_TYPE_CLASS
    Update_Statusbar 3, "CLASS MODULE"
    m_bDirty = True

End Sub

Private Sub optDLL_Click()
    
    m_udtOptionButtons.Version = optDLL.Value
    Update_Statusbar 4, "DLL"
    m_bDirty = True
    
End Sub

Private Sub optForm_Click()

    m_udtOptionButtons.ModuleType = MODULE_TYPE_FORM
    Update_Statusbar 3, "FORM/MODULE"
    m_bDirty = True
    
End Sub



Private Sub optFunction_Click()

    m_udtOptionButtons.Function_Routine = optFunction.Value
    Update_Statusbar 1, "FUNCTION"
    Routine IDX_FUNCTION
    m_bDirty = True
    
End Sub

Private Sub optPrivate_Click()

     m_udtOptionButtons.Instance.insPublic = False
    m_udtOptionButtons.Instance.insFriend = False
    m_udtOptionButtons.Instance.insPrivate = optPrivate.Value
    Update_Statusbar 2, "PRIVATE"
    m_bDirty = True
    
End Sub

Private Sub optPublic_Click()
    
    m_udtOptionButtons.Instance.insPrivate = False
    m_udtOptionButtons.Instance.insFriend = False
    m_udtOptionButtons.Instance.insPublic = optPublic.Value
    Update_Statusbar 2, "PUBLIC"
    m_bDirty = True
    
End Sub

Private Sub optFriend_Click()

    m_udtOptionButtons.Instance.insPrivate = False
    m_udtOptionButtons.Instance.insPublic = False
    m_udtOptionButtons.Instance.insFriend = optFriend.Value
    Update_Statusbar 2, "FRIEND"
    m_bDirty = True

End Sub

Private Sub optStand_Click()
    
    m_udtOptionButtons.Version = optStand.Value
    Update_Statusbar 4, "STANDALONE"
    m_bDirty = True
    
End Sub

Private Sub optSub_Click()
    
    m_udtOptionButtons.Sub_Routine = optSub.Value
    Update_Statusbar 1, "SUBROUTINE"
    Routine IDX_SUB
    m_bDirty = True
    
End Sub

Private Sub optProperty_Click()
    
    m_udtOptionButtons.Sub_Routine = optProperty.Value
    Update_Statusbar 1, "PROPERTY"
    Routine IDX_PROPERTY
    m_bDirty = True

End Sub

'**********************************************************************
'Procedure   : Private Method Routine
'Created on  : Wednesday, July 31, 2002
'Description :
'**********************************************************************
Private Function Routine(Index As Integer)
    
    Const PROCEDURE_NAME As String = "Routine"
    On Error GoTo ErrHandler
    
    Dim sRoutine As String
    Dim sGUID As String
    Dim sType As String
    
    '//what is the return type
    sType = Trim$(cboType.Text)
    
    '>>> allow no spaces in the procedure name <<<
    txtRoutineName.Text = Replace(Trim(txtRoutineName.Text), Chr(32), "_")
    
    Select Case Index
    Case IDX_FUNCTION
        If m_udtOptionButtons.ModuleType = MODULE_TYPE_FORM Then
            If m_udtOptionButtons.Instance.insPrivate Then
                If m_udtOptionButtons.Version = VER_STANDALONE Then
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Private Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Private Function " & txtRoutineName.Text & " () as " & sType & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME as String = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto ErrHandler" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto 0" & vbCrLf
                    sRoutine = sRoutine & "    Exit Function" & vbCrLf
                    sRoutine = sRoutine & "ErrHandler:" & vbCrLf
                    sRoutine = sRoutine & "    Select Case Err" & vbCrLf
                    sRoutine = sRoutine & "    Case Else" & vbCrLf
                    sRoutine = sRoutine & "        Select Case MsgBox(" + """[""" + " & " + "Err.Number" + " & " + """]""" + " & " + "Err.Description, _ " & vbCrLf
                    sRoutine = sRoutine & "        vbCritical" + "+" + "vbAbortRetryIgnore, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "        Case vbAbort: GoTo Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "        Case vbRetry: Resume 0" & vbCrLf
                    sRoutine = sRoutine & "        Case vbIgnore: Resume Next" & vbCrLf
                    sRoutine = sRoutine & "        End Select" & vbCrLf
                    sRoutine = sRoutine & "    End Select" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Function" & vbCrLf
                Else
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Private Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Private Function " & txtRoutineName.Text & " () As " & sType & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error GoTo Error_Handler" & vbCrLf
                    sRoutine = sRoutine & "    Call atsPUSHERR(MODULE_NAME, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    Call atsPOPERR(MODULE_NAME, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "    Exit Function" & vbCrLf
                    sRoutine = sRoutine & "Error_Handler:" & vbCrLf
                    sRoutine = sRoutine & "    Select Case Err" & vbCrLf
                    sRoutine = sRoutine & "    Case Else" & vbCrLf
                    sRoutine = sRoutine & "        Select Case atsERROR(MODULE_NAME, PROCEDURE_NAME, Err, Erl)" & vbCrLf
                    sRoutine = sRoutine & "        Case atsResume: Resume" & vbCrLf
                    sRoutine = sRoutine & "        Case atsResumeNext: Resume Next" & vbCrLf
                    sRoutine = sRoutine & "        Case atsBreak: Stop: Resume" & vbCrLf
                    sRoutine = sRoutine & "        Case atsAbort: Resume Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "        End Select" & vbCrLf
                    sRoutine = sRoutine & "    End Select" & vbCrLf
                    sRoutine = sRoutine & "    Resume Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Function" & vbCrLf
                End If
            ElseIf m_udtOptionButtons.Instance.insPublic Then
                If m_udtOptionButtons.Version = VER_STANDALONE Then
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Public Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Public Function " & txtRoutineName.Text & " () as " & sType & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME as String = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto ErrHandler" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto 0 " & vbCrLf
                    sRoutine = sRoutine & "    Exit Function" & vbCrLf
                    sRoutine = sRoutine & "ErrHandler:" & vbCrLf
                    sRoutine = sRoutine & "    Select Case Err" & vbCrLf
                    sRoutine = sRoutine & "    Case Else" & vbCrLf
                    sRoutine = sRoutine & "        Select Case MsgBox(" + """[""" + " & " + "Err.Number" + " & " + """]""" + " & " + "Err.Description, _ " & vbCrLf
                    sRoutine = sRoutine & "        vbCritical" + "+" + "vbAbortRetryIgnore, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "        Case vbAbort: GoTo Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "        Case vbRetry: Resume 0" & vbCrLf
                    sRoutine = sRoutine & "        Case vbIgnore: Resume Next" & vbCrLf
                    sRoutine = sRoutine & "        End Select" & vbCrLf
                    sRoutine = sRoutine & "    End Select" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Function" & vbCrLf
                Else
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Public Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Public Function " & txtRoutineName.Text & " () As " & sType & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error GoTo Error_Handler" & vbCrLf
                    sRoutine = sRoutine & "    Call atsPUSHERR(MODULE_NAME, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    Call atsPOPERR(MODULE_NAME, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "    Exit Function" & vbCrLf
                    sRoutine = sRoutine & "Error_Handler:" & vbCrLf
                    sRoutine = sRoutine & "    Select Case Err" & vbCrLf
                    sRoutine = sRoutine & "    Case Else" & vbCrLf
                    sRoutine = sRoutine & "        Select Case atsERROR(MODULE_NAME, PROCEDURE_NAME, Err, Erl)" & vbCrLf
                    sRoutine = sRoutine & "        Case atsResume: Resume" & vbCrLf
                    sRoutine = sRoutine & "        Case atsResumeNext: Resume Next" & vbCrLf
                    sRoutine = sRoutine & "        Case atsBreak: Stop: Resume" & vbCrLf
                    sRoutine = sRoutine & "        Case atsAbort: Resume Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "        End Select" & vbCrLf
                    sRoutine = sRoutine & "    End Select" & vbCrLf
                    sRoutine = sRoutine & "    Resume Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Function" & vbCrLf
                End If
            Else
                If m_udtOptionButtons.Version = VER_STANDALONE Then
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Friend Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Friend Function " & txtRoutineName.Text & " () as " & sType & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME as String = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto ErrHandler" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto 0 " & vbCrLf
                    sRoutine = sRoutine & "    Exit Function" & vbCrLf
                    sRoutine = sRoutine & "ErrHandler:" & vbCrLf
                    sRoutine = sRoutine & "    Select Case Err" & vbCrLf
                    sRoutine = sRoutine & "    Case Else" & vbCrLf
                    sRoutine = sRoutine & "        Select Case MsgBox(" + """[""" + " & " + "Err.Number" + " & " + """]""" + " & " + "Err.Description, _ " & vbCrLf
                    sRoutine = sRoutine & "        vbCritical" + "+" + "vbAbortRetryIgnore, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "        Case vbAbort: GoTo Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "        Case vbRetry: Resume 0" & vbCrLf
                    sRoutine = sRoutine & "        Case vbIgnore: Resume Next" & vbCrLf
                    sRoutine = sRoutine & "        End Select" & vbCrLf
                    sRoutine = sRoutine & "    End Select" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Function" & vbCrLf
                Else
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Friend Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Friend Function " & txtRoutineName.Text & " () As " & sType & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error GoTo Error_Handler" & vbCrLf
                    sRoutine = sRoutine & "    Call atsPUSHERR(MODULE_NAME, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    Call atsPOPERR(MODULE_NAME, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "    Exit Function" & vbCrLf
                    sRoutine = sRoutine & "Error_Handler:" & vbCrLf
                    sRoutine = sRoutine & "    Select Case Err" & vbCrLf
                    sRoutine = sRoutine & "    Case Else" & vbCrLf
                    sRoutine = sRoutine & "        Select Case atsERROR(MODULE_NAME, PROCEDURE_NAME, Err, Erl)" & vbCrLf
                    sRoutine = sRoutine & "        Case atsResume: Resume" & vbCrLf
                    sRoutine = sRoutine & "        Case atsResumeNext: Resume Next" & vbCrLf
                    sRoutine = sRoutine & "        Case atsBreak: Stop: Resume" & vbCrLf
                    sRoutine = sRoutine & "        Case atsAbort: Resume Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "        End Select" & vbCrLf
                    sRoutine = sRoutine & "    End Select" & vbCrLf
                    sRoutine = sRoutine & "    Resume Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Function" & vbCrLf
                End If
            End If
        Else
            If m_udtOptionButtons.Instance.insPrivate Then
                If m_udtOptionButtons.Version = VER_STANDALONE Or m_udtOptionButtons.Version = VER_DLL Then
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Private Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Private Function " & txtRoutineName.Text & " () as " & sType & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME as String = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto ErrHandler" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto 0" & vbCrLf
                    sRoutine = sRoutine & "    Exit Function" & vbCrLf
                    sRoutine = sRoutine & "ErrHandler:" & vbCrLf
                    sRoutine = sRoutine & "    Err.Raise Err.Number,PROCEDURE_NAME,Err.Description" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Function" & vbCrLf
                End If
            ElseIf m_udtOptionButtons.Instance.insPublic Then
                If m_udtOptionButtons.Version = VER_STANDALONE Or m_udtOptionButtons.Version = VER_DLL Then
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Public Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Public Function " & txtRoutineName.Text & " () as " & sType & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME as String = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto ErrHandler" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto 0 " & vbCrLf
                    sRoutine = sRoutine & "    Exit Function" & vbCrLf
                    sRoutine = sRoutine & "ErrHandler:" & vbCrLf
                    sRoutine = sRoutine & "    Err.Raise Err.Number,PROCEDURE_NAME,Err.Description" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Function" & vbCrLf
                End If
            Else
                If m_udtOptionButtons.Version = VER_STANDALONE Or m_udtOptionButtons.Version = VER_DLL Then
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Friend Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Friend Function " & txtRoutineName.Text & " () as " & sType & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME as String = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto ErrHandler" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto 0 " & vbCrLf
                    sRoutine = sRoutine & "    Exit Function" & vbCrLf
                    sRoutine = sRoutine & "ErrHandler:" & vbCrLf
                    sRoutine = sRoutine & "    Err.Raise Err.Number,PROCEDURE_NAME,Err.Description" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Function" & vbCrLf
                End If
            End If
        End If
    Case IDX_SUB
        If m_udtOptionButtons.ModuleType = MODULE_TYPE_FORM Then
            If m_udtOptionButtons.Instance.insPrivate Then
                If m_udtOptionButtons.Version = VER_STANDALONE Then
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Private Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Private Sub " & txtRoutineName.Text & " ()" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME as String = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto ErrHandler" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto 0 " & vbCrLf
                    sRoutine = sRoutine & "    Exit Sub" & vbCrLf
                    sRoutine = sRoutine & "ErrHandler:" & vbCrLf
                    sRoutine = sRoutine & "    Select Case Err" & vbCrLf
                    sRoutine = sRoutine & "    Case Else" & vbCrLf
                    sRoutine = sRoutine & "        Select Case MsgBox(" + """[""" + " & " + "Err.Number" + " & " + """]""" + " & " + "Err.Description, _ " & vbCrLf
                    sRoutine = sRoutine & "        vbCritical" + "+" + "vbAbortRetryIgnore, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "        Case vbAbort: GoTo Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "        Case vbRetry: Resume 0" & vbCrLf
                    sRoutine = sRoutine & "        Case vbIgnore: Resume Next" & vbCrLf
                    sRoutine = sRoutine & "        End Select" & vbCrLf
                    sRoutine = sRoutine & "    End Select" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Sub" & vbCrLf
                Else
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Private Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Private Sub " & txtRoutineName.Text & " ()" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error GoTo Error_Handler" & vbCrLf
                    sRoutine = sRoutine & "    Call atsPUSHERR(MODULE_NAME, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    Call atsPOPERR(MODULE_NAME, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "    Exit Sub" & vbCrLf
                    sRoutine = sRoutine & "Error_Handler:" & vbCrLf
                    sRoutine = sRoutine & "    Select Case Err" & vbCrLf
                    sRoutine = sRoutine & "    Case Else" & vbCrLf
                    sRoutine = sRoutine & "        Select Case atsERROR(MODULE_NAME, PROCEDURE_NAME, Err, Erl)" & vbCrLf
                    sRoutine = sRoutine & "        Case atsResume: Resume" & vbCrLf
                    sRoutine = sRoutine & "        Case atsResumeNext: Resume Next" & vbCrLf
                    sRoutine = sRoutine & "        Case atsBreak: Stop: Resume" & vbCrLf
                    sRoutine = sRoutine & "        Case atsAbort: Resume Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "        End Select" & vbCrLf
                    sRoutine = sRoutine & "    End Select" & vbCrLf
                    sRoutine = sRoutine & "    Resume Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Sub" & vbCrLf
                End If
            ElseIf m_udtOptionButtons.Instance.insPublic Then
                If m_udtOptionButtons.Version = VER_STANDALONE Then
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Public Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Public Sub " & txtRoutineName.Text & " ()" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME as String = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto ErrHandler" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto 0 " & vbCrLf
                    sRoutine = sRoutine & "    Exit Sub" & vbCrLf
                    sRoutine = sRoutine & "ErrHandler:" & vbCrLf
                    sRoutine = sRoutine & "    Select Case Err" & vbCrLf
                    sRoutine = sRoutine & "    Case Else" & vbCrLf
                    sRoutine = sRoutine & "        Select Case MsgBox(" + """[""" + " & " + "Err.Number" + " & " + """]""" + " & " + "Err.Description, _ " & vbCrLf
                    sRoutine = sRoutine & "        vbCritical" + "+" + "vbAbortRetryIgnore, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "        Case vbAbort: GoTo Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "        Case vbRetry: Resume 0" & vbCrLf
                    sRoutine = sRoutine & "        Case vbIgnore: Resume Next" & vbCrLf
                    sRoutine = sRoutine & "        End Select" & vbCrLf
                    sRoutine = sRoutine & "    End Select" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Sub" & vbCrLf
                Else
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Public Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Public Sub " & txtRoutineName.Text & " ()" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error GoTo Error_Handler" & vbCrLf
                    sRoutine = sRoutine & "    Call atsPUSHERR(MODULE_NAME, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    Call atsPOPERR(MODULE_NAME, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "    Exit Sub" & vbCrLf
                    sRoutine = sRoutine & "Error_Handler:" & vbCrLf
                    sRoutine = sRoutine & "    Select Case Err" & vbCrLf
                    sRoutine = sRoutine & "    Case Else" & vbCrLf
                    sRoutine = sRoutine & "        Select Case atsERROR(MODULE_NAME, PROCEDURE_NAME, Err, Erl)" & vbCrLf
                    sRoutine = sRoutine & "        Case atsResume: Resume" & vbCrLf
                    sRoutine = sRoutine & "        Case atsResumeNext: Resume Next" & vbCrLf
                    sRoutine = sRoutine & "        Case atsBreak: Stop: Resume" & vbCrLf
                    sRoutine = sRoutine & "        Case atsAbort: Resume Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "        End Select" & vbCrLf
                    sRoutine = sRoutine & "    End Select" & vbCrLf
                    sRoutine = sRoutine & "    Resume Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Sub" & vbCrLf
                End If
            Else
                If m_udtOptionButtons.Version = VER_STANDALONE Then
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Friend Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Friend Sub " & txtRoutineName.Text & " ()" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME as String = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto ErrHandler" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto 0 " & vbCrLf
                    sRoutine = sRoutine & "    Exit Sub" & vbCrLf
                    sRoutine = sRoutine & "ErrHandler:" & vbCrLf
                    sRoutine = sRoutine & "    Select Case Err" & vbCrLf
                    sRoutine = sRoutine & "    Case Else" & vbCrLf
                    sRoutine = sRoutine & "        Select Case MsgBox(" + """[""" + " & " + "Err.Number" + " & " + """]""" + " & " + "Err.Description, _ " & vbCrLf
                    sRoutine = sRoutine & "        vbCritical" + "+" + "vbAbortRetryIgnore, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "        Case vbAbort: GoTo Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "        Case vbRetry: Resume 0" & vbCrLf
                    sRoutine = sRoutine & "        Case vbIgnore: Resume Next" & vbCrLf
                    sRoutine = sRoutine & "        End Select" & vbCrLf
                    sRoutine = sRoutine & "    End Select" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Sub" & vbCrLf
                Else
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Friend Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Friend Sub " & txtRoutineName.Text & " ()" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error GoTo Error_Handler" & vbCrLf
                    sRoutine = sRoutine & "    Call atsPUSHERR(MODULE_NAME, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    Call atsPOPERR(MODULE_NAME, PROCEDURE_NAME)" & vbCrLf
                    sRoutine = sRoutine & "    Exit Sub" & vbCrLf
                    sRoutine = sRoutine & "Error_Handler:" & vbCrLf
                    sRoutine = sRoutine & "    Select Case Err" & vbCrLf
                    sRoutine = sRoutine & "    Case Else" & vbCrLf
                    sRoutine = sRoutine & "        Select Case atsERROR(MODULE_NAME, PROCEDURE_NAME, Err, Erl)" & vbCrLf
                    sRoutine = sRoutine & "        Case atsResume: Resume" & vbCrLf
                    sRoutine = sRoutine & "        Case atsResumeNext: Resume Next" & vbCrLf
                    sRoutine = sRoutine & "        Case atsBreak: Stop: Resume" & vbCrLf
                    sRoutine = sRoutine & "        Case atsAbort: Resume Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "        End Select" & vbCrLf
                    sRoutine = sRoutine & "    End Select" & vbCrLf
                    sRoutine = sRoutine & "    Resume Procedure_Exit" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Sub" & vbCrLf
                End If
            End If
        Else
            If m_udtOptionButtons.Instance.insPrivate Then
                If m_udtOptionButtons.Version = VER_STANDALONE Or m_udtOptionButtons.Version = VER_DLL Then
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Private Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Private Sub " & txtRoutineName.Text & " ()" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME as String = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto ErrHandler" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto 0 " & vbCrLf
                    sRoutine = sRoutine & "    Exit Sub" & vbCrLf
                    sRoutine = sRoutine & "ErrHandler:" & vbCrLf
                    sRoutine = sRoutine & "    Err.Raise Err.Number,PROCEDURE_NAME,Err.Description" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Sub" & vbCrLf
                End If
            ElseIf m_udtOptionButtons.Instance.insPublic Then
                If m_udtOptionButtons.Version = VER_STANDALONE Or m_udtOptionButtons.Version = VER_DLL Then
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Public Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Public Sub " & txtRoutineName.Text & " ()" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME as String = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto ErrHandler" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto 0 " & vbCrLf
                    sRoutine = sRoutine & "    Exit Sub" & vbCrLf
                    sRoutine = sRoutine & "ErrHandler:" & vbCrLf
                    sRoutine = sRoutine & "    Err.Raise Err.Number,PROCEDURE_NAME,Err.Description" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Sub" & vbCrLf
                End If
            Else
                If m_udtOptionButtons.Version = VER_STANDALONE Or m_udtOptionButtons.Version = VER_DLL Then
                    sRoutine = sRoutine & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "'Procedure   : Friend Method " & txtRoutineName.Text & vbCrLf
                    sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                    sRoutine = sRoutine & "Friend Sub " & txtRoutineName.Text & " ()" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "    Const PROCEDURE_NAME as String = " & Chr(34) & txtRoutineName.Text & Chr(34) & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto ErrHandler" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "'//TODO:<<Place code here>>" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "Procedure_Exit:" & vbCrLf
                    sRoutine = sRoutine & "    On Error Goto 0 " & vbCrLf
                    sRoutine = sRoutine & "    Exit Sub" & vbCrLf
                    sRoutine = sRoutine & "ErrHandler:" & vbCrLf
                    sRoutine = sRoutine & "    Err.Raise Err.Number,PROCEDURE_NAME,Err.Description" & vbCrLf
                    sRoutine = sRoutine & "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
                    sRoutine = sRoutine & "End Sub" & vbCrLf
                End If
            End If
        End If
    Case IDX_PROPERTY
        '//Property in class or form has no reason
        'If m_udtOptionButtons.ModuleType = MODULE_TYPE_FORM Then
            If m_udtOptionButtons.Instance.insPrivate Then
                sRoutine = sRoutine & vbCrLf
                sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                sRoutine = sRoutine & "'Property   : Private Property " & txtRoutineName.Text & vbCrLf
                sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                
                If m_udtPropType.lstGet Then
                sRoutine = sRoutine & "Private Property Get " & txtRoutineName.Text & " () as " & sType & vbCrLf
                sRoutine = sRoutine & vbTab & txtRoutineName.Text & " = " & txtStructure.Text & "." & txtRoutineName.Text & vbCrLf
                sRoutine = sRoutine & "End Property" & vbCrLf
                End If
                
                If m_udtPropType.lstLet Then
                sRoutine = sRoutine & "Private Property Let " & txtRoutineName.Text & " (ByVal NewValue as " & sType & " )" & vbCrLf
                sRoutine = sRoutine & vbTab & txtStructure.Text & "." & txtRoutineName.Text & " = NewValue " & vbCrLf
                sRoutine = sRoutine & "End Property" & vbCrLf
                End If
                   
                If m_udtPropType.lstSet Then
                sRoutine = sRoutine & "Private Property Set " & txtRoutineName.Text & " (ByVal NewValue as " & sType & " )" & vbCrLf
                sRoutine = sRoutine & vbTab & "Set " & txtStructure.Text & "." & txtRoutineName.Text & " = NewValue " & vbCrLf
                sRoutine = sRoutine & "End Property" & vbCrLf
                End If
                
            ElseIf m_udtOptionButtons.Instance.insPublic Then
                sRoutine = sRoutine & vbCrLf
                sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                sRoutine = sRoutine & "'Property   : Public Property " & txtRoutineName.Text & vbCrLf
                sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                
                If m_udtPropType.lstGet Then
                sRoutine = sRoutine & "Public Property Get " & txtRoutineName.Text & " () as " & sType & vbCrLf
                sRoutine = sRoutine & vbTab & txtRoutineName.Text & " = " & txtStructure.Text & "." & txtRoutineName.Text & vbCrLf
                sRoutine = sRoutine & "End Property" & vbCrLf
                End If
                
                If m_udtPropType.lstLet Then
                sRoutine = sRoutine & "Public Property Let " & txtRoutineName.Text & " (ByVal NewValue as " & sType & " )" & vbCrLf
                sRoutine = sRoutine & vbTab & txtStructure.Text & "." & txtRoutineName.Text & " = NewValue " & vbCrLf
                sRoutine = sRoutine & "End Property" & vbCrLf
                End If
                
                If m_udtPropType.lstSet Then
                sRoutine = sRoutine & "Public Property Set " & txtRoutineName.Text & " (ByVal NewValue as " & sType & " )" & vbCrLf
                sRoutine = sRoutine & vbTab & "Set " & txtStructure.Text & "." & txtRoutineName.Text & " = NewValue " & vbCrLf
                sRoutine = sRoutine & "End Property" & vbCrLf
                End If
            Else
                sRoutine = sRoutine & vbCrLf
                sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                sRoutine = sRoutine & "'Property   : Friend Property " & txtRoutineName.Text & vbCrLf
                sRoutine = sRoutine & "'**********************************************************************" & vbCrLf
                
                If m_udtPropType.lstGet Then
                sRoutine = sRoutine & "Friend Property Get " & txtRoutineName.Text & " () as " & sType & vbCrLf
                sRoutine = sRoutine & vbTab & txtRoutineName.Text & " = " & txtStructure.Text & "." & txtRoutineName.Text & vbCrLf
                sRoutine = sRoutine & "End Property" & vbCrLf
                End If
                
                If m_udtPropType.lstLet Then
                sRoutine = sRoutine & "Friend Property Let " & txtRoutineName.Text & " (ByVal NewValue as " & sType & " )" & vbCrLf
                sRoutine = sRoutine & vbTab & txtStructure.Text & "." & txtRoutineName.Text & " = NewValue " & vbCrLf
                sRoutine = sRoutine & "End Property" & vbCrLf
                End If
                
                If m_udtPropType.lstSet Then
                sRoutine = sRoutine & "Friend Property Set " & txtRoutineName.Text & " (ByVal NewValue as " & sType & " )" & vbCrLf
                sRoutine = sRoutine & vbTab & "Set " & txtStructure.Text & "." & txtRoutineName.Text & " = NewValue " & vbCrLf
                sRoutine = sRoutine & "End Property" & vbCrLf
                End If
            End If
        'End If
    End Select
    
    m_udtOptionButtons.ErrCode = sRoutine
    
    If m_udtOptionButtons.Version = VER_DLL Then
        MsgBox "Using this version will require a reference to the ATSError.DLL", vbInformation, App.Title
    End If
    
Procedure_Exit:
    On Error GoTo 0
    Exit Function
ErrHandler:
   Select Case Err
   Case Else
       Select Case MsgBox("[" & Err.Number & "]" & Err.Description, vbCritical + vbAbortRetryIgnore, App.Title)
       Case vbAbort: GoTo Procedure_Exit
       Case vbRetry: Resume 0
       Case vbIgnore: Resume Next
       End Select
   End Select
End Function

Private Sub Update_Statusbar(IPanel As Integer, sText As String)
    
    On Error Resume Next

    With StatusBar1
        Select Case IPanel
        Case 1
            .Panels(1).Text = "Type: " & sText
        Case 2
            .Panels(2).Text = "Instance: " & sText
        Case 3
            .Panels(3).Text = "Module: " & sText
        Case 4
            .Panels(4).Text = "Version: " & sText
        End Select
    End With

End Sub

Private Sub txtRoutineName_Change()

    m_bDirty = True

End Sub
