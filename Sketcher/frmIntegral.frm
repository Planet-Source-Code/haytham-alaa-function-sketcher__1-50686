VERSION 5.00
Begin VB.Form frmIntegral 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Evaluate Integrals"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Smart accuracy changer"
      Height          =   375
      Left            =   2700
      TabIndex        =   10
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtAcc 
      Height          =   285
      Left            =   3600
      TabIndex        =   8
      Text            =   "100"
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdEval 
      Caption         =   "Evaluate"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Text            =   "2"
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   3600
      TabIndex        =   1
      Text            =   "1"
      Top             =   960
      Width           =   615
   End
   Begin VB.ListBox lstFunctions 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Accuracy :"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Choose a function to evaluate the integral for it :"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "To :"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "From :"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "frmIntegral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
   If Check1.Value = vbUnchecked Then
      txtAcc.Enabled = True
   Else
      txtAcc.Enabled = False
   End If
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdEval_Click()
   If lstFunctions.ListIndex = -1 Then MsgBox "No Function displayed": Exit Sub
   If IsNumeric(txtAcc.Text) = False Or IsNumeric(txtFrom.Text) = False Or IsNumeric(txtTo.Text) = False Then MsgBox "You must enter numbers in the form and to boxes"
   MsgBox CalcIntg(lstFunctions.List(lstFunctions.ListIndex))
End Sub

Public Function CalcIntg(ByVal Funct)
   Dim X, Part, Final
   X = (Val(txtTo) - Val(txtFrom)) / Val(txtAcc)
   Noww = (Val(txtFrom) + (Val(txtFrom) + X)) / 2
   For i = txtFrom To (Val(txtFrom) / X)
      fSubs = Form1.Subs(Funct, Noww)
'      If fSubs = 4 / 6 Then
'         MsgBox "Done"
'      End If
      Part = Part + fSubs ' Form1.Subs(Funct, Noww)
      Noww = Noww + X
   Next
   Final = X * Part
   CalcIntg = Final
End Function

Private Sub txtFrom_Change()
   If Check1.Value = vbChecked Then
      txtAcc.Text = (Val(txtTo.Text) - Val(txtFrom.Text)) * 100
      If Val(txtAcc.Text) < 0 Then txtAcc.Text = -1 * Val(txtAcc.Text)
   End If
End Sub

Private Sub txtTo_Change()
   If Check1.Value = vbChecked Then
      txtAcc.Text = (Val(txtTo.Text) - Val(txtFrom.Text)) * 100
      If Val(txtAcc.Text) < 0 Then txtAcc.Text = -1 * Val(txtAcc.Text)
   End If

End Sub
