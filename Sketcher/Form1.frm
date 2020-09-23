VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Sketch Functions"
   ClientHeight    =   6975
   ClientLeft      =   135
   ClientTop       =   345
   ClientWidth     =   11400
   DrawWidth       =   2
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawStyle       =   6  'Inside Solid
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   11355
      TabIndex        =   0
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function CalcStr(ByVal TxT)

' This function was made by "Haytham Alaa"
' In order to use this function you must contact me at kasparov03@hotmail.com

On Error Resume Next
   If IsNumeric(TxT) Then
      CalcStr = Val(TxT)
      Exit Function
   End If
   If TxT Like "*--*" Then
      TxT = Replace(TxT, "--", "+")
   End If
   If InStr(TxT, "[") = 0 Then
      If InStr(TxT, "|") = 0 Then
         If Not TxT Like "*(*" Then

            
            For i = Len(TxT) To 1 Step -1
               If Mid(TxT, i, 1) = "+" Then CalcStr = CalcStr(Mid(TxT, 1, i - 1)) + CalcStr(Mid(TxT, i + 1))
               If Mid(TxT, i, 1) = "-" Then CalcStr = CalcStr(Mid(TxT, 1, i - 1)) - CalcStr(Mid(TxT, i + 1))
            Next
          
            For i = 1 To Len(TxT)
               If Mid(TxT, i, 1) = "*" Then CalcStr = CalcStr(Mid(TxT, 1, i - 1)) * CalcStr(Mid(TxT, i + 1, Len(TxT) - i))
               If Mid(TxT, i, 1) = "/" Then CalcStr = CalcStr(Mid(TxT, 1, i - 1)) / CalcStr(Mid(TxT, i + 1, Len(TxT) - i))
            Next
            
            P = Split(TxT, "^")
            For i = 0 To UBound(P) Step 2
               CalcStr = P(i) ^ P(i + 1)
            Next
            
            Si = Split(TxT, "sin")
            For i = 0 To UBound(Si) Step 2
               CalcStr = Sin(Si(i + 1))
            Next
            
            Cs = Split(TxT, "csc")
            For i = 0 To UBound(Cs) Step 2
               CalcStr = 1 / (Sin(Cs(i + 1)))
            Next
            
            Co = Split(TxT, "cos")
            For i = 0 To UBound(Co) Step 2
               CalcStr = Cos(Co(i + 1))
            Next
         
            Se = Split(TxT, "sec")
            For i = 0 To UBound(Se) Step 2
               CalcStr = 1 / (Cos(Se(i + 1)))
            Next
            
            Ta = Split(TxT, "tan")
            For i = 0 To UBound(Ta) Step 2
               CalcStr = 1 / (Tan(Ta(i + 1)))
            Next
            
            Cotn = Split(TxT, "cot")
            For i = 0 To UBound(Cotn) Step 2
               CalcStr = 1 / (Tan(Cotn(i + 1)))
            Next
      
            logn = Split(TxT, "log")
            For i = 0 To UBound(logn) Step 2
               CalcStr = Log(logn(i + 1))
            Next
            
            lnn = Split(TxT, "ln")
            For i = 0 To UBound(lnn) Step 2
               CalcStr = Log(lnn(i + 1)) / Log(2.7182812846)
            Next
      
         Else
            For i = 1 To Len(TxT)
               If Mid(TxT, i, 1) = "(" Then bstart = i
               If Mid(TxT, i, 1) = ")" Then
                  bend = i
                  s1 = CalcStr(Mid(TxT, bstart + 1, bend - bstart - 1))
                  TxT = Replace(TxT, Mid(TxT, bstart, bend - bstart + 1), s1)
                  i = 0
               End If
            Next
            CalcStr = CalcStr(TxT)
         End If
      Else
         For i = 1 To Len(TxT)
            If Mid(TxT, i, 1) = "|" And bFirst = "" Then AbStart = i: i = i + 1: bFirst = "Done"
            If Mid(TxT, i, 1) = "|" Then
               AbEnd = i
               s1 = CalcStr(Mid(TxT, AbStart + 1, AbEnd - AbStart - 1))
               If s1 < 0 Then
                  s1 = -1 * s1
               End If
               TxT = Replace(TxT, Mid(TxT, AbStart, AbEnd - AbStart + 1), s1)
               i = 0
            End If
         Next
         CalcStr = CalcStr(TxT)
      End If
   Else
      For i = 1 To Len(TxT)
         If Mid(TxT, i, 1) = "[" Then bstart = i
         If Mid(TxT, i, 1) = "]" Then
            bend = i
            s1 = CalcStr(Mid(TxT, bstart + 1, bend - bstart - 1))
            If Val(s1 / 2) <> Int(s1 / 2) Then s1 = Int(s1)
            TxT = Replace(TxT, Mid(TxT, bstart, bend - bstart + 1), s1)
            i = 0
         End If
      Next
      CalcStr = CalcStr(TxT)
   End If
End Function

Public Function Subs(ByVal Funct As String, ByVal Num As Double)   'As Double
   If CalcStr(Replace(Funct, "x", Num)) <> "" Then
      Subs = CalcStr(Replace(Funct, "x", Num))
   Else
      Subs = "1/0"
      For i = 0 To Pic1.Height Step 320
         Pic1.Line (Pic1.Width / 2 + Num * 50, i)-(Pic1.Width / 2 + Num * 50, i + 160)
      Next
   End If
End Function

Public Sub DrawGraph(ByVal Funct As String, txtFrom As Double, txtTo As Double, txtStep As Double, Clr As Long, xScale As Double, yScale As Double)
On Error Resume Next
   Dim Fun, i, ii
   If Clr = 2 Then Pic1.Cls
   DrawGrid
   
   Fun = Split(Funct, ";")
   For ii = 0 To UBound(Fun)
      R1 = Int(Rnd * 255)
      G1 = Int(Rnd * 255)
      B1 = Int(Rnd * 255)
      For i = Val(txtFrom) To Val(txtTo) Step txtStep
         If i = 0 Then
            shi = 2
         End If
         DoEvents
         If Subs(Fun(ii), i) <> "1/0" And Subs(Fun(ii), i + Val(txtStep)) <> "1/0" Then
            Pic1.Line ((Pic1.Width / 2) + (i * xScale * 10), Pic1.Height / 2 - Subs(Fun(ii), i) * yScale * 10)-((Pic1.Width / 2) + (i + Val(txtStep)) * xScale * 10, Pic1.Height / 2 - Val(txtStep) - Subs(Fun(ii), i + Val(txtStep)) * yScale * 10), RGB(R1, G1, B1)
         End If
         PCounter = Int((i - Val(txtFrom)) / (Val(txtTo) - Val(txtFrom)) * 100)
         MDIForm1.Caption = PCounter & " % Complete"
      Next
   Next
   MDIForm1.Caption = "Function Sketcher"
End Sub

Private Sub Form_Activate()
   Set ActiveFrm = Me
End Sub

Private Sub Form_Load()
   Me.Width = 9200
   Me.Height = 7000
   Me.Left = frmProperties.Width
   DrawGrid
End Sub

Private Sub Form_Resize()
   Pic1.Width = Me.Width
   Pic1.Height = Me.Height
   Pic1.Cls
   DrawGrid
End Sub

Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Xx1 = (X - Pic1.Left - (Pic1.Width / 2)) / 50
   Yy1 = (-Y + Pic1.Top + (Pic1.Height / 2)) / 50
   Pic1.ToolTipText = "( " & Xx1 & " , " & Yy1 & " )"
End Sub

Private Function DrawGrid()
   For i = -50 To Pic1.Width Step (160 + frmProperties.txtxScale * 10)
      Pic1.Line (i, 0)-(i, Pic1.Height), RGB(217, 217, 217)
      Pic1.Line (0, i)-(Pic1.Width, i), RGB(217, 217, 217)
   Next
   Pic1.Line (Pic1.Width / 2, 0)-(Pic1.Width / 2, Pic1.Height)
   Pic1.Line (0, Pic1.Height / 2)-(Pic1.Width, Pic1.Height / 2)

End Function

Public Function DrawPoint(ByVal X, ByVal Y)
    Functt = "((" & (Y ^ 2 + X ^ 2) * 50 & ")-(x^2))^0.5"
    DrawGraph Functt, -50, 50, 1, 0, 30, 30
    DrawGraph "-1*((" & (Y ^ 2 + X ^ 2) * 50 & ")-(x^2))^0.5", -50, 50, 1, 0, 50, 50
End Function

Public Function deff(ByVal FFunct As String)
   If Not FFunct Like "*(*" Then
      s1 = Split(FFunct, "+")
      For i = 0 To UBound(s1) - 1
         deff = deff & "+" & deff(s1(i))
      Next

      s2 = Split(FFunct, "-")
      For i = 0 To UBound(s2) - 1
         deff = deff(s2(i))
      Next
      
      s3 = Split(FFunct, "*")
      For i = 0 To UBound(s3) - 1 Step 2
         deff = deff(s3(i)) & "*" & s3(i + 1) & "+" & s3(i) & "*" & deff(s3(i + 1))
      Next
      
      For i = 1 To Len(FFunct)
         If Mid(FFunct, i, 1) = "^" Then deff = (Mid(FFunct, i + 1, Len(FFunct) - i) & "*(" & Mid(FFunct, 1, i) & Mid(FFunct, i + 1, Len(FFunct) - i) - 1) & ")"
      Next
   Else
   End If
End Function

Public Function Def(ByVal FFunct As String)
'   On Error Resume Next
   If IsNumeric(FFunct) = True Then Def = "0": Exit Function
   If LCase(FFunct) = "x" Then Def = "1": Exit Function
   If Not FFunct Like "*(*" Then
      For i = 1 To Len(FFunct)
         If Mid$(FFunct, i, 1) = "+" Then Def = "(" & Def(Mid(FFunct, 1, i - 1)) & ")" & "+" & "(" & Def(Mid(FFunct, i + 1, Len(FFunct) - i)) & ")": Exit Function
         If Mid$(FFunct, i, 1) = "-" Then Def = "(" & Def(Mid(FFunct, 1, i - 1)) & ")" & "-" & "(" & Def(Mid(FFunct, i + 1, Len(FFunct) - i)) & ")": Exit Function
      Next
      
      For i = 1 To Len(FFunct)
         If Mid(FFunct, i, 1) = "*" Then Def = "(" & Def(Mid(FFunct, 1, i - 1)) & "*" & Mid(FFunct, i + 1, Len(FFunct) - i) & ")" & "+" & "(" & Mid(FFunct, 1, i - 1) & "*" & Def(Mid(FFunct, i + 1, Len(FFunct) - i)) & ")"
      Next
      
      For i = 1 To Len(FFunct)
         If Mid(FFunct, i, 1) = "^" Then Def = (Mid(FFunct, i + 1, Len(FFunct) - i)) & "*" & "(" & Mid(FFunct, 1, i - 1) & "^" & (Mid(FFunct, i + 1, Len(FFunct) - i) - 1) & ")"
      Next
   Else
      For i = 1 To Len(FFunct)
         If Mid$(FFunct, i, 1) = "(" Then
            If sStart = "" Then
               sStart = i
            End If
            sBrk = sBrk + 1
         End If
         If Mid$(FFunct, i, 1) = ")" Then
            eBrk = eBrk + 1
            If eBrk = sBrk Then eEnd = i: Exit For
         End If
      Next
      s1 = Mid(FFunct, sStart + 1, eEnd - 1 - sStart)
      s2 = Mid(FFunct, eEnd + 2, Len(FFunct) - eEnd)
      s3 = Mid(FFunct, eEnd + 2, Len(FFunct) - eEnd)
      s4 = Mid(FFunct, sStart + 1, eEnd - 1 - sStart)
      If Mid(FFunct, eEnd + 1, 1) = "*" Then Def = Def(s1) & "*" & s2 & "+" & Def(s3) & "*" & s4
      If Mid(FFunct, eEnd + 1, 1) = "+" Or Mid(FFunct, eEnd + 1, 1) = "-" Then Def = Def(s1) & Mid(FFunct, eEnd + 1, 1) & "(" & Def(s2) & ")": Exit Function

      If sStart = 1 Then Def = Def(s1): Exit Function
      If Mid(FFunct, sStart - 1, 1) = "+" Or Mid(FFunct, sStart - 1, 1) = "-" Then Def = Def(Mid(FFunct, 1, sStart - 2)) & Mid(FFunct, sStart - 1, 1) & "(" & Def(s1) & ")": Exit Function
      If Mid(FFunct, sStart - 1, 1) = "*" Then Def = Mid(FFunct, 1, sStart - 2) & "*" & "(" & Def(s1) & ")" & "+" & Def(Mid(FFunct, 1, sStart - 2)) & "*" & "(" & s1 & ")": Exit Function
      If Len(FFunct) <= eEnd Then Def = Def(FFunct)
   End If
End Function

Public Function InCpt(ByVal FFunct, ByVal sFrom, ByVal sTo, ByVal sStep)
   For i = sFrom To sTo Step sStep
      If Subs(FFunct, i) = 0 Then InCpt = i: Exit Function
      If Subs(FFunct, i) > 0 And Subs(FFunct, i + sStep) < 0 Then
         InCpt = InCpt(FFunct, i, i + sStep, sStep / 10)
      Else
         If Subs(FFunct, i + sStep) > 0 And Subs(FFunct, i) < 0 Then
            InCpt = InCpt(FFunct, i, i + sStep, sStep / 10)
         End If
      End If
   Next
End Function
