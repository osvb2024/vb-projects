VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   Caption         =   "Odd or Even?"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
   EndProperty
   Font            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOddOrEven 
      Caption         =   "Odd or Even"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0018
      Height          =   750
      Left            =   1073
      TabIndex        =   0
      Top             =   1688
      Width           =   2535
   End
   Begin VB.TextBox txtNumber 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0029
      ForeColor       =   &H80000002&
      Height          =   750
      Left            =   1073
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "Please enter a number"
      Top             =   653
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOddOrEven_Click()
    Dim n As Integer
    Dim r As Integer
    
    n = Val(txtNumber.Text)
    r = n Mod 2
    
    If r = 0 Then
        MsgBox "Even"
    Else
        MsgBox "Odd"
    End If
End Sub

Private Sub txtNumber_GotFocus()
    txtNumber.Text = ""
End Sub
