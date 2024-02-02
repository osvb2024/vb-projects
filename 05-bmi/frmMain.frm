VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H007A6E54&
   Caption         =   "MBI Calculator - Ali Daei"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3675
   BeginProperty Font 
   EndProperty
   Font            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Appearance      =   0  'Flat
      BackColor       =   &H00C868BA&
      Caption         =   "&R"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0018
      Height          =   510
      Left            =   2452
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2925
      Width           =   690
   End
   Begin VB.CommandButton cmdCalculate 
      Appearance      =   0  'Flat
      BackColor       =   &H002257FF&
      Caption         =   "&Calculate"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0028
      Height          =   510
      Left            =   532
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2925
      Width           =   1935
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0038
      ForeColor       =   &H00205E1B&
      Height          =   510
      Left            =   690
      TabIndex        =   3
      Text            =   "92.6"
      Top             =   2130
      Width           =   2295
   End
   Begin VB.TextBox txtHeight 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0048
      ForeColor       =   &H00205E1B&
      Height          =   510
      Left            =   690
      TabIndex        =   2
      Text            =   "1.8"
      Top             =   900
      Width           =   2295
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0058
      ForeColor       =   &H00F1EFEC&
      Height          =   405
      Left            =   630
      TabIndex        =   11
      Top             =   5235
      Width           =   2400
   End
   Begin VB.Label lblBMITitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "YOUR BMI IS"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0068
      ForeColor       =   &H00F1EFEC&
      Height          =   405
      Left            =   630
      TabIndex        =   10
      Top             =   3915
      Width           =   2400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   270
      X2              =   3390
      Y1              =   3735
      Y2              =   3735
   End
   Begin VB.Label lblBMI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0078
      ForeColor       =   &H00F1EFEC&
      Height          =   585
      Left            =   1680
      TabIndex        =   9
      Top             =   4485
      Width           =   300
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   270
      X2              =   3390
      Y1              =   5895
      Y2              =   5895
   End
   Begin VB.Label lblAppVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 2.0"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0088
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1410
      TabIndex        =   8
      Top             =   6375
      Width           =   855
   End
   Begin VB.Label lblAppName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BMI Calculator"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0098
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1080
      TabIndex        =   7
      Top             =   6120
      Width           =   1515
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "© 2023 Ali Daei. All rights reserved."
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":00A8
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   540
      TabIndex        =   6
      Top             =   6585
      Width           =   2595
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Weight (kg):"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":00B8
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   690
      TabIndex        =   5
      Top             =   1650
      Width           =   1725
   End
   Begin VB.Label lblMass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Height (m):"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":00C8
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   690
      TabIndex        =   4
      Top             =   420
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalculate_Click()
    Dim height As Double
    Dim weight As Double
    Dim bmi As Double

    height = Val(txtHeight.Text)
    weight = Val(txtWeight.Text)
    
    If height = 0 Then Exit Sub

    bmi = weight / (height * height)

    lblBMI.Caption = Round(bmi, 1)
    Me.Caption = "YOUR BMI IS " + Str(Round(bmi, 1))
    
    If bmi < 18.5 Then

        lblResult.Caption = "(underweight)"
        lblResult.ForeColor = &HFFC4C4
        lblBMI.ForeColor = &HFFC4C4
        Me.BackColor = &HCC6E6E
        
    ElseIf bmi < 25 Then

        lblResult.Caption = "(normal)"
        lblResult.ForeColor = &H53C800
        lblBMI.ForeColor = &H53C800
        Me.BackColor = &H205E1B
        
    ElseIf bmi < 30 Then

        lblResult.Caption = "(overweight)"
        lblResult.ForeColor = &HAAD3FF
        lblBMI.ForeColor = &HAAD3FF
        Me.BackColor = &H6CE0
        
    Else

        lblResult.Caption = "(obese)"
        lblResult.ForeColor = &HAAAAFF
        lblBMI.ForeColor = &HAAAAFF
        Me.BackColor = &HC4

    End If

End Sub

Private Sub cmdReset_Click()
    Dim answer As Integer
    
    answer = MsgBox("The form will be reset!", vbOKCancel + vbExclamation + vbDefaultButton2, "Reset")
    
    If answer <> vbOK Then Exit Sub
    
    txtHeight.Text = "0"
    txtWeight.Text = "0"
    lblBMI.Caption = "0"
    lblBMI.ForeColor = &HF1EFEC
    lblResult.Caption = "-"
    lblResult.ForeColor = &HF1EFEC
    Me.BackColor = &H7A6E54
    Me.Caption = "MBI Calculator - Ali Daei"
End Sub
