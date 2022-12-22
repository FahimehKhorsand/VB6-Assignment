VERSION 5.00
Begin VB.Form frmInterestCalculation 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Interest Calculation"
   ClientHeight    =   5280
   ClientLeft      =   6735
   ClientTop       =   3285
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   6210
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2295
      TabIndex        =   4
      Top             =   4665
      Width           =   1440
   End
   Begin VB.CommandButton cmdCalculation 
      Caption         =   "Calculation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2370
      TabIndex        =   3
      Top             =   3825
      Width           =   1335
   End
   Begin VB.ListBox lstbxDisplay 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   495
      TabIndex        =   2
      Top             =   2085
      Width           =   5220
   End
   Begin VB.TextBox txtPrincipal 
      Height          =   405
      Left            =   2385
      TabIndex        =   1
      Top             =   1005
      Width           =   2040
   End
   Begin VB.TextBox txtInterestRate 
      Height          =   420
      Left            =   2400
      TabIndex        =   0
      Top             =   330
      Width           =   2010
   End
   Begin VB.Frame framNeededInformation 
      Caption         =   "Needed Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6000
      Begin VB.Label lblPrincipal 
         BackColor       =   &H80000004&
         Caption         =   "Principal:              $"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   765
         TabIndex        =   8
         Top             =   915
         Width           =   1500
      End
      Begin VB.Label lblInterestRate 
         BackColor       =   &H80000004&
         Caption         =   "Interest Rate:      %"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   720
         TabIndex        =   7
         Top             =   255
         Width           =   1560
      End
   End
   Begin VB.Frame framResult 
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "frmInterestCalculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalculation_Click()
    Dim intYears As Integer
    Dim dblInterestRate As Double
    Dim curAmount As Currency
    Dim curPrincipal As Currency
    
    
''''''Check Empty Fields
  
    If txtInterestRate.Text = "" Then
        MsgBox "Interest Rate field cannot be empty"
        txtInterestRate.SetFocus
    Exit Sub
    End If
    
    If txtPrincipal.Text = "" Then
        MsgBox "Principal field cannot be empty"
        txtPrincipal.SetFocus
    Exit Sub
    End If
    
''''''Formula
    
    lstbxDisplay.Clear
    intYears = 0
    
    dblInterestRate = txtInterestRate.Text / 100
    curPrincipal = txtPrincipal.Text
    
    lstbxDisplay.AddItem ("Year" & vbTab & "Amount on Diposit")
    
    
    For intYears = 0 To 10
        curAmount = curPrincipal * (1 + dblInterestRate) ^ intYears
        lstbxDisplay.AddItem (Format$(intYears, "@@@@") & vbTab & _
                             Format$(Format$(curAmount, "Currency"), String$(17, "@")))
    Next intYears
    
   cmdExit.SetFocus
    
End Sub

Private Sub cmdExit_Click()

Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

Unload Me

End Sub

Private Sub txtInterestRate_KeyDown(KeyCode As Integer, Shift As Integer)
 
 If KeyCode = vbKeyReturn Then
     txtPrincipal.SetFocus

End If
End Sub

Private Sub txtInterestRate_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

    If txtInterestRate.Text = "" Then
    MsgBox ("This field cannot be empty")
    txtInterestRate.SetFocus
    End If
End If

End Sub

Private Sub txtPrincipal_KeyDown(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyReturn Then
    cmdCalculation.SetFocus

End If
End Sub

Private Sub txtPrincipal_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

    If txtPrincipal.Text = "" Then
    MsgBox ("This field cannot be empty")
    txtPrincipal.SetFocus
    End If
End If

End Sub
