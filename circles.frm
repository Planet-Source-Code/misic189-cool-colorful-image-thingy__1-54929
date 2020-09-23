VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   Caption         =   "Form1"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5280
      Top             =   3480
   End
   Begin VB.Menu file 
      Caption         =   "file"
      Begin VB.Menu print 
         Caption         =   "print"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim changer As Integer
Dim intTime As Integer
Dim randcir As Integer
Dim n1 As Integer
Dim n2 As Integer
Dim m1 As Integer
Dim m2 As Integer
Dim j1 As Integer
Dim j2 As Integer
Dim r1 As Integer
Dim r2 As Integer
Dim y1 As Integer
Dim y2 As Integer
Dim u1 As Integer
Dim u2 As Integer

Private Sub Form_Click()
Rem:goes to sub that resets all values if you click on screen
restart
End Sub

Private Sub Form_Load()
Form1.ForeColor = vbBlack
Rem:goes to sub that resets all values
restart
End Sub

Private Sub print_Click()
Form1.PrintForm
End Sub

Private Sub Timer1_Timer()
Rem:makes circles at random places
Circle (n1, n2), changer
Circle (m1, m2), changer
Circle (j1, j2), changer
Circle (r1, r2), changer
Circle (y1, y2), changer
Circle (u1, u2), changer

Rem:increases circle diameter by random value
changer = changer + randcir

Rem:if circles get too big then goes to restart to clear values
If changer > 12000 Then
restart
End If
Form1.ForeColor = Form1.ForeColor + 1
End Sub
Public Sub restart()
Cls
Form1.ForeColor = vbBlack
changer = 0

Rem:makes a very random numbers for positions of circles
Rem:according to the time
intTime = (Hour(Time) + Minute(Time) + Second(Time))
Do
randcir = Int(Rnd * 40)
n1 = Int(Rnd * Form1.Width)
n2 = Int(Rnd * Form1.Height)
m1 = Int(Rnd * Form1.Width)
m2 = Int(Rnd * Form1.Height)
j1 = Int(Rnd * Form1.Width)
j2 = Int(Rnd * Form1.Height)
r1 = Int(Rnd * Form1.Width)
r2 = Int(Rnd * Form1.Height)
y1 = Int(Rnd * Form1.Width)
y2 = Int(Rnd * Form1.Height)
u1 = Int(Rnd * Form1.Width)
u2 = Int(Rnd * Form1.Height)
intTime = intTime - 1
Loop Until intTime = 0
End Sub
