VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Function Sketcher"
   ClientHeight    =   4800
   ClientLeft      =   1800
   ClientTop       =   3000
   ClientWidth     =   6840
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   1560
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar2 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   5
      Top             =   420
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   1482
      BandCount       =   1
      _CBWidth        =   6840
      _CBHeight       =   840
      _Version        =   "6.0.8169"
      Child1          =   "Toolbar1"
      MinHeight1      =   780
      Width1          =   3375
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   780
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   1376
         ButtonWidth     =   1191
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "imgIcons"
         DisabledImageList=   "imgIcons"
         HotImageList    =   "imgIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "New"
               Object.ToolTipText     =   "Add Graph"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "d/dx"
               Object.ToolTipText     =   "Evaluate Derevative"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Integral"
               Object.ToolTipText     =   "Evaluate definite Integral"
            EndProperty
         EndProperty
         MousePointer    =   1
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   741
      BandCount       =   1
      _CBWidth        =   6840
      _CBHeight       =   420
      _Version        =   "6.0.8169"
      MinHeight1      =   360
      Width1          =   2880
      NewRow1         =   0   'False
      Begin VB.TextBox label1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   120
         TabIndex        =   4
         Text            =   "Y = "
         Top             =   120
         Width           =   330
      End
      Begin VB.CommandButton cmdDraw 
         Caption         =   "Draw"
         Height          =   255
         Left            =   4450
         TabIndex        =   3
         Top             =   75
         Width           =   615
      End
      Begin VB.ComboBox cmbDrawMode 
         Height          =   315
         ItemData        =   "frmMain.frx":0452
         Left            =   5100
         List            =   "frmMain.frx":045C
         TabIndex        =   2
         Top             =   70
         Width           =   1695
      End
      Begin VB.TextBox txtFunct 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   480
         TabIndex        =   1
         Top             =   75
         Width           =   3975
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Begin VB.Menu mnuFileNewGraph 
            Caption         =   "Graph"
         End
         Begin VB.Menu mnuFileNewProject 
            Caption         =   "Project"
         End
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Begin VB.Menu mnuFileOpenProj 
            Caption         =   "Existing Project"
         End
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "Close Project"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAdv 
      Caption         =   "Advanced Operations"
      Begin VB.Menu mnuAdvExt 
         Caption         =   "Show Local Maximum and Local Minimum"
      End
      Begin VB.Menu mnuAdvInteg 
         Caption         =   "Evaluate Integrals"
      End
      Begin VB.Menu mnuAdvDer 
         Caption         =   "Evaluate Derevative"
      End
      Begin VB.Menu mnuAdvTg 
         Caption         =   "Draw Tangent"
      End
      Begin VB.Menu mnuAdvIntCpt 
         Caption         =   "X-Interception"
      End
      Begin VB.Menu mnuAdvYIntCpt 
         Caption         =   "Y-Interception"
      End
   End
   Begin VB.Menu mnuWind 
      Caption         =   "Window"
      Begin VB.Menu mnuWindFuncCreator 
         Caption         =   "View Function Creator window"
      End
      Begin VB.Menu mnuWindProp 
         Caption         =   "View Properties window"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuWindExp 
         Caption         =   "View Function Explorer"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frmGraph() As Form1

' **************************************************************
' This program was made by Haytham Alaa, FCIS, AinShams Uni, Cairo, Egypt
' Thankx for everyone who helped me in some way or another
' **************************************************************

Private Sub cmdDraw_Click()
   Dim FFunct, i
   ActiveFrm.DrawGraph txtFunct.Text, frmProperties.txtFrom, frmProperties.txtTo, frmProperties.txtStep, frmProperties.chkClear.Value + 1, frmProperties.txtxScale, frmProperties.txtyScale
   FFunct = Split(txtFunct.Text, ";")
    For i = 0 To UBound(FFunct)
        frmFuncExp.lstFunc.AddItem FFunct(i) & "  " & ActiveFrm.Caption
        frmIntegral.lstFunctions.AddItem FFunct(i)
    Next
End Sub

Private Sub MDIForm_Load()
   ReDim frmGraph(1)
   Set frmGraph(0) = New Form1
   
   frmGraph(0).Enabled = True
   frmGraph(0).Caption = "Graph 1"
   frmGraph(0).Show
   Debug.Print label1.Left
   Debug.Print txtFunct.Left
   Debug.Print txtFunct.Width / Me.Width
   Debug.Print cmdDraw.Left / Me.Width
   Debug.Print cmbDrawMode.Left / Me.Width
   frmGraph(0).SetFocus
End Sub

Private Sub MDIForm_Resize()
   txtFunct.Width = 0.6220657 * Me.Width
   cmdDraw.Left = 0.713615 * Me.Width
   cmbDrawMode.Left = 0.798122 * Me.Width
   label1.Left = 120
End Sub

Private Sub MDIForm_Terminate()
   End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   End
End Sub

Private Sub mnuAdvDer_Click()
'   If Not ActiveForm.Caption = "Graph*" Then MsgBox "No Graph selected": Exit Sub
   MsgBox ActiveForm.Def(txtFunct.Text)
End Sub

Private Sub mnuAdvIntCpt_Click()
'   If Not ActiveForm.Caption = "Graph*" Then MsgBox "No Graph selected": Exit Sub
   MsgBox ActiveForm.Caption
   MsgBox ActiveForm.InCpt(txtFunct.Text, frmProperties.txtFrom, frmProperties.txtTo, 1)
End Sub

Private Sub mnuAdvInteg_Click()
   frmIntegral.Show 1
End Sub

Private Sub mnuAdvTg_Click()
   Dim ss
'   If Not ActiveForm.Caption = "Graph*" Then MsgBox "No Graph selected": Exit Sub
   ss = ActiveForm.Def(txtFunct.Text)
   ActiveForm.DrawGraph ss, frmProperties.txtFrom, frmProperties.txtTo, frmProperties.txtStep, frmProperties.chkClear.Value + 1, frmProperties.txtxScale, frmProperties.txtyScale
End Sub

Private Sub mnuAdvYIntCpt_Click()
   MsgBox ActiveForm.Subs(txtFunct.Text, 0)
End Sub

Private Sub mnuFileExit_Click()
   Dim a
   a = MsgBox("Are you sure you want to exit all opened windows?", vbYesNo)
   If a = vbNo Then Exit Sub
   End
End Sub

Private Sub mnuFileNewGraph_Click()
   ReDim Preserve frmGraph(UBound(frmGraph) + 1)
   Set frmGraph(UBound(frmGraph)) = New Form1
   frmGraph(UBound(frmGraph)).Enabled = True
   frmGraph(UBound(frmGraph)).Caption = "Graph " & UBound(frmGraph)
   frmGraph(UBound(frmGraph)).Show
   frmGraph(UBound(frmGraph)).SetFocus
End Sub

Private Sub mnuWindExp_Click()
'    frmFuncExp.Show
   Static frmExp As frmFuncExp
   Set frmExp = frmFuncExp
   If mnuWindExp.Checked = False Then
      mnuWindExp.Checked = True
      frmExp.Enabled = True
      frmExp.Show
   Else
      mnuWindExp.Checked = False
      frmExp.Hide
   End If

End Sub

Private Sub mnuWindProp_Click()
   Static frmProp As frmProperties
   Set frmProp = frmProperties
   If mnuWindProp.Checked = False Then
      mnuWindProp.Checked = True
      frmProp.Enabled = True
      frmProp.Show
   Else
      mnuWindProp.Checked = False
      frmProp.Hide
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
   Select Case Button.Caption
      Case "New"
         mnuFileNewGraph_Click
      Case "d/dx"
         mnuAdvDer_Click
      Case "Integral"
         mnuAdvInteg_Click
   End Select
   
End Sub

Private Sub txtFunct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdDraw_Click
End Sub
