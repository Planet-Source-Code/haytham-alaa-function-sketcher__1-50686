VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   5580
   ClientTop       =   1290
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox txtFunct 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Text            =   "x"
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Sketch Now"
      Height          =   320
      Left            =   3120
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "-50"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Text            =   "50"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtStep 
      Height          =   285
      Left            =   4560
      TabIndex        =   1
      Text            =   "1"
      Top             =   360
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Draw in the same Graph"
      Height          =   255
      Left            =   4680
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Function :"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "From :"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "To :"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub
