VERSION 5.00
Begin VB.Form frmLoanPaymentCalculation 
   BackColor       =   &H00FFFFFF&
   Caption         =   "LoanPaymentCalculation"
   ClientHeight    =   4590
   ClientLeft      =   7080
   ClientTop       =   3555
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   6765
   Begin VB.TextBox txtMonthlyPayment 
      BackColor       =   &H80000003&
      Enabled         =   0   'False
      Height          =   405
      Left            =   2280
      TabIndex        =   5
      Top             =   3090
      Width           =   2730
   End
   Begin VB.TextBox txtInterestRate 
      Height          =   390
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   2730
   End
   Begin VB.TextBox txtLoanTerm 
      Height          =   435
      Left            =   2400
      TabIndex        =   3
      Top             =   1680
      Width           =   2745
   End
   Begin VB.TextBox txtAmntOfLoan 
      Height          =   420
      Left            =   2400
      TabIndex        =   1
      Top             =   465
      Width           =   2700
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
      Left            =   5145
      TabIndex        =   4
      Top             =   3030
      Width           =   1335
   End
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
      Left            =   2670
      TabIndex        =   0
      Top             =   3990
      Width           =   1515
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
      Height          =   2625
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   6660
      Begin VB.Label lblAmntOfLoan 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   " Amount of Loan:    $"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   420
         TabIndex        =   11
         Top             =   405
         Width           =   1785
      End
      Begin VB.Label lblAnnualInterestRate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "Annual Interest Rate:"
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
         Left            =   375
         TabIndex        =   10
         Top             =   1065
         Width           =   1815
      End
      Begin VB.Label lblTerm 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "Term: (years)"
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
         Left            =   345
         TabIndex        =   9
         Top             =   1710
         Width           =   1800
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
      Height          =   1170
      Left            =   45
      TabIndex        =   7
      Top             =   2715
      Width           =   6675
      Begin VB.Label lblMonthlyPayment 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "Monthly Payment:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   405
         TabIndex        =   8
         Top             =   390
         Width           =   1770
      End
   End
End
Attribute VB_Name = "frmLoanPaymentCalculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculation_Click()

Dim intTrm As Integer
Dim dblAmnt As Double
Dim dblPayment As Double
Dim dblRate As Double

''''''Check Empty Fields

If txtAmntOfLoan.Text = "" Then
    MsgBox "Amnt Of Loan field cannot be empty"
    txtAmntOfLoan.SetFocus
Exit Sub
End If

If txtInterestRate.Text = "" Then
    MsgBox "Interest Rate field cannot be empty"
    txtInterestRate.SetFocus
Exit Sub
End If

If txtLoanTerm.Text = "" Then
    MsgBox "Loan Term field cannot be empty"
    LoanTerm.SetFocus
Exit Sub
End If

'''''''Formula

dblAmnt = Val(txtAmntOfLoan.Text)
dblRate = (Val(txtInterestRate.Text) / 100) / 12
 intTrm = Val(txtLoanTerm.Text) * 12
dblPayment = Pmt(dblRate, intTrm, -dblAmnt, 0, 0)

txtMonthlyPayment.Text = Format(dblPayment, "$#,##0.00")

cmdExit.SetFocus

End Sub



Private Sub cmdExit_Click()

Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

Unload Me

End Sub



Private Sub txtAmntOfLoan_KeyDown(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyReturn Then
    txtInterestRate.SetFocus

End If

End Sub

Private Sub txtAmntOfLoan_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

    If txtAmntOfLoan.Text = "" Then
    MsgBox ("This field cannot be empty")
    txtAmntOfLoan.SetFocus
    End If
End If

End Sub

Private Sub txtInterestRate_KeyDown(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyReturn Then
    txtLoanTerm.SetFocus

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

Private Sub txtLoanTerm_KeyDown(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyReturn Then
    cmdCalculation.SetFocus

End If

End Sub

Private Sub txtLoanTerm_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

    If txtLoanTerm.Text = "" Then
    MsgBox ("This field cannot be empty")
    txtLoanTerm.SetFocus
    End If
End If
End Sub
