VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H007A6E54&
   Caption         =   "MBI Calculator - Ali Daei"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3675
   BeginProperty Font 
   EndProperty
   Font            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6195
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   2
      Text            =   "92.600"
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
      TabIndex        =   0
      Text            =   "1.80"
      Top             =   900
      Width           =   2295
   End
   Begin VB.Label lblBMITitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "YOUR BMI IS"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0058
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
   Begin VB.Label lblBMIResult 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0068
      ForeColor       =   &H00F1EFEC&
      Height          =   585
      Left            =   1680
      TabIndex        =   9
      Top             =   4410
      Width           =   300
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   270
      X2              =   3390
      Y1              =   5115
      Y2              =   5115
   End
   Begin VB.Label lblAppVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 0.1"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0078
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1410
      TabIndex        =   8
      Top             =   5595
      Width           =   855
   End
   Begin VB.Label lblAppName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BMI Calculator"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0088
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1080
      TabIndex        =   7
      Top             =   5340
      Width           =   1515
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "© 2023 Ali Daei. All rights reserved."
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0098
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   540
      TabIndex        =   6
      Top             =   5805
      Width           =   2595
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Weight (kg):"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":00A8
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   690
      TabIndex        =   3
      Top             =   1650
      Width           =   1725
   End
   Begin VB.Label lblMass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Height (m):"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":00B8
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   690
      TabIndex        =   1
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
    Dim sngHeight As Single
    Dim sngWeight As Single
    Dim sngBMI As Single

    sngHeight = Val(txtHeight.Text)
    sngWeight = Val(txtWeight.Text)

    sngBMI = sngWeight / (sngHeight * sngHeight)
    
    lblBMIResult.Caption = Round(sngBMI, 1)
    lblBMIResult.ForeColor = &H53C800
    Me.BackColor = &H205E1B
    Me.Caption = "YOUR BMI IS " & Round(sngBMI, 1)
End Sub

Private Sub cmdReset_Click()
    txtHeight.Text = "0"
    txtWeight.Text = "0"
    lblBMIResult.Caption = "0"
    lblBMIResult.ForeColor = &HF1EFEC
    Me.BackColor = &H7A6E54
    Me.Caption = "MBI Calculator - Ali Daei"
End Sub
