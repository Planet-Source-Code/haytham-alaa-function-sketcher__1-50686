VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "* Properties"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   1560
   ClientWidth     =   2355
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   2355
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbScale 
      Height          =   315
      ItemData        =   "frmProperties.frx":0000
      Left            =   240
      List            =   "frmProperties.frx":0013
      TabIndex        =   11
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CheckBox chkClear 
      Caption         =   "Clear Graph before sketching"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox txtyScale 
      Height          =   285
      Left            =   840
      TabIndex        =   9
      Text            =   "50"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtxScale 
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Text            =   "50"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtStep 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Text            =   "1"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Text            =   "50"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "-50"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblyScale 
      Caption         =   "Y Scale :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblxScale 
      Caption         =   "X Scale :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Step :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "To :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "From :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbScale_Click()
   Select Case cmbScale.ListIndex
      Case 0
         txtxScale.Text = 40
         txtyScale.Text = 40
         txtStep.Text = 0.09
      Case 1
         txtxScale.Text = 45
         txtyScale.Text = 45
         txtStep.Text = 0.1

      Case 2
         txtxScale.Text = 40
         txtyScale.Text = 25
         txtStep.Text = 0.1
      
      Case 3
         txtxScale.Text = 25
         txtyScale.Text = 40
         txtStep.Text = 0.1
      
      Case 4
         txtxScale.Text = 30
         txtyScale.Text = 30
         txtStep.Text = 0.1
      
   End Select
      
End Sub

Private Sub Form_Load()
    frmProperties.Top = 0
    frmProperties.Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIForm1.mnuWindProp.Checked = False
End Sub
