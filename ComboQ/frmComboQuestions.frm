VERSION 5.00
Begin VB.Form frmComboQuestions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Combo Questions"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmComboQuestions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkValue 
      Caption         =   "Name,Age / Name,Age,Country"
      Height          =   195
      Left            =   135
      TabIndex        =   8
      ToolTipText     =   "Double Value OR Tripal Value"
      Top             =   1740
      Value           =   2  'Grayed
      Width           =   3420
   End
   Begin VB.CommandButton Command1 
      Height          =   570
      Left            =   3720
      Picture         =   "frmComboQuestions.frx":0E52
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1950
      Width           =   840
   End
   Begin VB.Frame fraPerson 
      Caption         =   "Person Info"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   750
         Width           =   840
      End
      Begin VB.TextBox txtCountry 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1170
         Width           =   3375
      End
      Begin VB.ComboBox cboName 
         Height          =   315
         Left            =   960
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   285
         Width           =   3375
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Age"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   765
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Country"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmComboQuestions.frx":1294
      Stretch         =   -1  'True
      Top             =   2340
      Width           =   330
   End
End
Attribute VB_Name = "frmComboQuestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboName_Click()
Dim idx As Integer
    idx = cboName.ListIndex
    
    If chkValue.Value = vbChecked Then
        txtAge.Text = CStr(cboName.ItemData(idx))
    ElseIf chkValue.Value = vbUnchecked Then
            
        txtAge.Text = CStr(MyPerson(cboName.ItemData(idx)).pAge)
        txtCountry.Text = MyPerson(cboName.ItemData(idx)).pCountry
    End If
    
End Sub

Private Sub chkValue_Click()
    If chkValue.Value = vbChecked Then
        lblTitle(1).Enabled = False
        txtCountry.Enabled = False
        LoadValues True
        
    ElseIf chkValue.Value = vbUnchecked Then
        lblTitle(1).Enabled = True
        txtCountry.Enabled = True
        LoadValues False
        
    End If
    
    
End Sub

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()
    ReDim MyPerson(TotalPerson)
    Dim i As Integer
    
    GetMyPersons
    chkValue.Value = vbChecked
End Sub
'*** Real mechanism for the loading starts here ***
Private Sub LoadValues(ByVal vValue As Boolean)
    Dim i As Integer
    cboName.Clear
    
    For i = 0 To TotalPerson - 1
        cboName.AddItem MyPerson(i).pName '*** Adding a new Item ***
        '*** Adding Numeric Value/Array Index depending upon the request ***
        cboName.ItemData(cboName.NewIndex) = IIf(vValue, MyPerson(i).pAge, i)
    Next i
cboName.Refresh

    '*** Setting default Value into the combobox ***
    cboName.Text = cboName.List(DefaultText)
End Sub


Private Sub Image1_Click()
    MsgBox "Please feel free to write your Comments/Suggestions." & vbCrLf & " Thnx!" & vbCrLf & "-Deepakk_2k@yahoo.com", vbInformation, "Thanks!"
End Sub
