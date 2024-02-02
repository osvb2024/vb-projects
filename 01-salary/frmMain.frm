VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00D39530&
   Caption         =   "Salary"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5085
   BeginProperty Font 
   EndProperty
   Font            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0018
      Height          =   500
      Left            =   3495
      TabIndex        =   10
      Top             =   2580
      Width           =   1200
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0028
      Height          =   500
      Left            =   2092
      TabIndex        =   9
      Top             =   2580
      Width           =   1200
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0038
      Height          =   500
      Left            =   390
      TabIndex        =   8
      Top             =   2580
      Width           =   1500
   End
   Begin VB.TextBox txtSalary 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0048
      Height          =   315
      Left            =   1620
      TabIndex        =   6
      Top             =   1215
      Width           =   3105
   End
   Begin VB.TextBox txtLastName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0058
      Height          =   315
      Left            =   2250
      TabIndex        =   2
      Top             =   705
      Width           =   2500
   End
   Begin VB.TextBox txtFirstName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0068
      Height          =   315
      Left            =   2250
      TabIndex        =   1
      Top             =   225
      Width           =   2500
   End
   Begin VB.Label lblPayment 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0078
      ForeColor       =   &H80000001&
      Height          =   315
      Left            =   2205
      TabIndex        =   7
      Top             =   1785
      Width           =   2535
   End
   Begin VB.Label lblPaymentTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " Payment:"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0088
      ForeColor       =   &H80000001&
      Height          =   315
      Left            =   465
      TabIndex        =   5
      Top             =   1785
      Width           =   4260
   End
   Begin VB.Label lblSalary 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " Salary:"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0098
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   465
      TabIndex        =   4
      Top             =   1215
      Width           =   1155
   End
   Begin VB.Label lblLastName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " Last name: "
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":00A8
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   480
      TabIndex        =   3
      Top             =   705
      Width           =   1770
   End
   Begin VB.Label lblFirstName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " First name: "
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":00B8
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   465
      TabIndex        =   0
      Top             =   225
      Width           =   1785
   End
   Begin VB.Image imgFormPattern 
      Height          =   10575
      Left            =   -405
      Picture         =   "frmMain.frx":00C8
      Top             =   2595
      Width           =   10650
   End
   Begin VB.Image Imagebgyalda 
      Height          =   10575
      Left            =   -345
      Picture         =   "frmMain.frx":44990
      Top             =   -10110
      Width           =   10650
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalculate_Click()
    Dim salary As Double
    Dim insurance As Double
    Dim tax As Double
    Dim payment As Double

    salary = Val(txtSalary.Text)

    insurance = salary * 0.03
    tax = salary * 0.05
    payment = salary - insurance - tax

    lblPayment.Caption = payment
    lblPayment.ForeColor = &H2C9300
    Me.BackColor = &H67A319
End Sub

Private Sub cmdReset_Click()
    lblPayment.Caption = 0
    lblPayment.ForeColor = &HE0
    txtFirstName.Text = ""
    txtLastName.Text = ""
    txtSalary.Text = ""
    txtFirstName.SetFocus
    Me.BackColor = &H239AD8
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

