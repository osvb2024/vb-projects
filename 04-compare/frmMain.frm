VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   Caption         =   "Compare"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
   EndProperty
   Font            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4290
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumX 
      Alignment       =   2  'Center
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0018
      ForeColor       =   &H80000002&
      Height          =   750
      Left            =   1073
      TabIndex        =   1
      Text            =   "0"
      Top             =   705
      Width           =   2535
   End
   Begin VB.TextBox txtNumY 
      Alignment       =   2  'Center
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0029
      ForeColor       =   &H80000002&
      Height          =   750
      Left            =   1073
      TabIndex        =   2
      Text            =   "0"
      Top             =   2085
      Width           =   2535
   End
   Begin VB.CommandButton cmdCompare 
      Caption         =   "Compare"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":003A
      Height          =   750
      Left            =   1073
      TabIndex        =   0
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label lblY 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":004B
      ForeColor       =   &H8000000A&
      Height          =   525
      Left            =   2220
      TabIndex        =   4
      Top             =   1515
      Width           =   255
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":005C
      ForeColor       =   &H8000000A&
      Height          =   525
      Left            =   2213
      TabIndex        =   3
      Top             =   150
      Width           =   255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCompare_Click()
    Dim x As Integer
    Dim y As Integer
    
    x = Val(txtNumX.Text)
    y = Val(txtNumY.Text)
    
    If x < y Then
        MsgBox "X is less than Y"
    ElseIf x > y Then
        MsgBox "X is greater than Y"
    Else
        MsgBox "X is equal to Y"
    End If
End Sub

Private Sub txtNumY_GotFocus()
    If txtNumY.Text = "0" Then
        txtNumY.Text = ""
    End If
End Sub

Private Sub txtNumY_LostFocus()
    If txtNumY.Text = "" Then
        txtNumY.Text = "0"
    End If
End Sub

Private Sub txtNumY_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
