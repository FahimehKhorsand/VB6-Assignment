VERSION 5.00
Begin VB.Form frmMainMenu 
   ClientHeight    =   5505
   ClientLeft      =   5730
   ClientTop       =   2535
   ClientWidth     =   9825
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   9825
   Begin VB.PictureBox Picture1 
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   -2640
      Picture         =   "frmMainMenu.frx":0000
      ScaleHeight     =   5955
      ScaleWidth      =   12435
      TabIndex        =   0
      Top             =   -480
      Width           =   12495
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   2910
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2835
         Width           =   3885
      End
      Begin VB.CommandButton cmdLoanPaymentCalculation 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Loan Payment Calculation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2940
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2100
         Width           =   3870
      End
      Begin VB.CommandButton cmdInterestCalculation 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Interest Calculation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2955
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1335
         Width           =   3870
      End
      Begin VB.Label lblCalcOfLoanInterest 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Calculation of loan and bank deposit interest"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   495
         Left            =   2865
         TabIndex        =   4
         Top             =   600
         Width           =   9525
      End
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

frmInterestCalculation.Show

End Sub

Private Sub cmdExit_Click()

Unload Me

End Sub

Private Sub cmdInterestCalculation_Click()

frmInterestCalculation.Show

End Sub

Private Sub cmdLoanPaymentCalculation_Click()

frmLoanPaymentCalculation.Show

End Sub

Private Sub Exit_Click()

Unload Me

End Sub

Private Sub InterestCalculation_Click()

frmInterestCalculation.Show

End Sub


Private Sub LoanPaymentCalculation_Click()

frmLoanPaymentCalculation.Show

End Sub
