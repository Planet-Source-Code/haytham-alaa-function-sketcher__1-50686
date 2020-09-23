VERSION 5.00
Begin VB.Form frmFuncExp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Function Explorer"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2310
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstFunc 
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmFuncExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Left = 0 'MDIForm1.ScaleWidth - Me.Width
    Me.Top = frmProperties.Top + frmProperties.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIForm1.mnuWindExp.Checked = False
End Sub

