VERSION 5.00
Begin VB.Form frmFTPSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FTP Settings"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmFTPSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2970
      TabIndex        =   9
      Top             =   1575
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1545
      TabIndex        =   8
      Top             =   1575
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1125
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1200
      Width           =   3465
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1125
      TabIndex        =   5
      Top             =   825
      Width           =   3465
   End
   Begin VB.TextBox txtTarget 
      Height          =   285
      Left            =   1125
      TabIndex        =   3
      Top             =   450
      Width           =   3465
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1125
      TabIndex        =   1
      Top             =   75
      Width           =   3465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   75
      TabIndex        =   6
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Username:"
      Height          =   195
      Left            =   75
      TabIndex        =   4
      Top             =   825
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Target file:"
      Height          =   195
      Left            =   75
      TabIndex        =   2
      Top             =   450
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Host address:"
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   975
   End
End
Attribute VB_Name = "frmFTPSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cancel As Boolean

Private Sub cmdCancel_Click()
    Cancel = True
    Me.Hide
End Sub

Private Sub cmdOk_Click()

    txtHost.Text = Trim$(txtHost.Text)
    txtTarget.Text = Trim$(txtTarget.Text)
    
    If txtHost.Text = "" Then
        MsgBox "You failed to enter an FTP host address.", vbExclamation, "Try Again"
        Exit Sub
    End If
    
    If txtTarget.Text = "" Then
        MsgBox "You failed to enter a target file name/path.", vbExclamation, "Try Again"
        Exit Sub
    End If
    
    If txtUsername.Text = "" Then
        MsgBox "You failed to enter a user name.", vbExclamation, "Try Again"
        Exit Sub
    End If
    
    If txtPassword.Text = "" Then
        MsgBox "You failed to enter an password.", vbExclamation, "Try Again"
        Exit Sub
    End If
    
    Cancel = False
    Me.Hide
    
End Sub

