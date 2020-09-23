VERSION 5.00
Object = "{DF6D6558-5B0C-11D3-9396-008029E9B3A6}#1.0#0"; "ezVidC60.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmCamera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ZangCam Pro"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4875
   Icon            =   "frmCamera.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFTPSettings 
      Caption         =   "FTP Settings"
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   4050
      Width           =   1065
   End
   Begin VB.TextBox txtMsg 
      Height          =   315
      Left            =   1125
      TabIndex        =   4
      Text            =   "%d - %t"
      Top             =   3675
      Width           =   3690
   End
   Begin InetCtlsObjects.Inet netFTP 
      Left            =   4200
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin VB.CommandButton cmdOnOff 
      Caption         =   "Start Cam"
      Height          =   315
      Left            =   3900
      TabIndex        =   2
      Top             =   4050
      Width           =   915
   End
   Begin VB.CommandButton cmdCamSettings 
      Appearance      =   0  'Flat
      Caption         =   "Cam Settings"
      Height          =   315
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3675
      Width           =   1065
   End
   Begin VB.Timer tmrCapture 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4350
      Top             =   75
   End
   Begin VB.PictureBox picBMP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3615
      Left            =   0
      ScaleHeight     =   237
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   3
      Top             =   0
      Width           =   4815
   End
   Begin vbVidC60.ezVidCap VidCap 
      Height          =   3600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6350
      StreamMaster    =   1
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "Next snapshot:"
      Height          =   195
      Left            =   1125
      TabIndex        =   5
      Top             =   4050
      Width           =   1065
   End
End
Attribute VB_Name = "frmCamera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private c As New cDIBSection
Private nTime As Integer

Public Username As String
Public Password As String
Public Host As String
Public TargetFile As String

Private Sub cmdFTPSettings_Click()
    With frmFTPSettings
    .txtHost.Text = Host
    .txtPassword.Text = Password
    .txtTarget.Text = TargetFile
    .txtUsername.Text = Username
    .Show vbModal
    If .Cancel Then Exit Sub
    Username = .txtUsername.Text
    Password = .txtPassword.Text
    Host = .txtHost.Text
    TargetFile = .txtTarget.Text
    End With
    SaveFTPSettings
End Sub

Private Sub cmdOnOff_Click()
    If tmrCapture.Enabled = False Then
        tmrCapture.Enabled = True
        cmdOnOff.Caption = "On"
    Else
        tmrCapture.Enabled = False
        cmdOnOff.Caption = "Off"
    End If
    
End Sub

Private Sub cmdCamSettings_Click()
    VidCap.ShowDlgVideoSource
End Sub

Private Sub Form_Load()

    VidCap.CaptureFile = App.Path & "\webcam.bmp"
    picBMP.FontBold = True
    picBMP.FontName = "Tahoma"
    picBMP.FontSize = 10
    
    LoadFTPSettings
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
    
End Sub

Private Sub tmrCapture_Timer()
    
    nTime = nTime + 1
    lblTime.Caption = "Next snapshot: " & (10 - nTime) & " seconds"
    If nTime = 10 Then
        nTime = 0
        If VidCap.CapSingleFrame Then
            VidCap.SaveDIB VidCap.CaptureFile
            picBMP.Cls
            Set picBMP.Picture = LoadPicture(VidCap.CaptureFile)
            
            WriteText
            SavePicture picBMP.Image, VidCap.CaptureFile
            picBMP.Cls
            Set picBMP.Picture = LoadPicture(VidCap.CaptureFile)
            c.CreateFromPicture picBMP.Picture
            SaveJPG c, App.Path & "\webcam.jpg"
            Kill VidCap.CaptureFile
            
            'On Error Resume Next
            '// Upload the file
            netFTP.RemoteHost = Host
            netFTP.Username = Username
            netFTP.Password = Password
            netFTP.Execute , "PUT """ & App.Path & "\webcam.jpg"" """ & TargetFile & """"
            Do While netFTP.StillExecuting: DoEvents: Loop
            'On Error GoTo 0
            
        End If
    End If
    
End Sub

Private Sub WriteText()
    Static szMsg As String

    szMsg = txtMsg.Text
    szMsg = Replace$(szMsg, "%d", Format$(Now, "dddd, mmm d yyyy"))
    szMsg = Replace$(szMsg, "%t", Format$(Now, "hh:mm:ss AMPM"))
    
    '// Upper Left      (*|/)
    '                   (-+-)
    '                   (/|\)
    
    picBMP.CurrentY = 214
    picBMP.CurrentX = 4
    picBMP.ForeColor = 0
    picBMP.Print szMsg
    
    '// Middle Left     (\|/)
    '                   (*+-)
    '                   (/|\)
    
    picBMP.CurrentY = 215
    picBMP.CurrentX = 4
    picBMP.Print szMsg
    
    '// Lower Left      (\|/)
    '                   (-+-)
    '                   (*|\)
    
    picBMP.CurrentY = 216
    picBMP.CurrentX = 4
    picBMP.Print szMsg
    
    '// Lower Middle    (\|/)
    '                   (-+-)
    '                   (/*\)
    
    picBMP.CurrentY = 216
    picBMP.CurrentX = 5
    picBMP.Print szMsg
    

    '// Lower Right     (\|/)
    '//                 (-+-)
    '//                 (/|*)
    
    picBMP.CurrentY = 216
    picBMP.CurrentX = 6
    picBMP.Print szMsg
    
    '// Middle Right    (\|/)
    '                   (-+*)
    '                   (/|\)
    
    picBMP.CurrentY = 215
    picBMP.CurrentX = 6
    picBMP.Print szMsg
    
    '// Upper Right     (\|*)
    '                   (-+-)
    '                   (/|\)
    
    picBMP.CurrentY = 214
    picBMP.CurrentX = 6
    picBMP.Print szMsg
    
    '// Upper Middle    (\*/)
    '                   (-+-)
    '                   (/|\)
    
    picBMP.CurrentY = 214
    picBMP.CurrentX = 5
    picBMP.Print szMsg

    '// Middle Middle   (\|/)
    '                   (-*-)
    '                   (/|\)
    
    picBMP.CurrentY = 215
    picBMP.CurrentX = 5
    picBMP.ForeColor = RGB(&HFF, &HCC, 0)
    picBMP.Print szMsg
    
End Sub

Public Sub SaveFTPSettings()
    SaveSetting App.Title, App.Major & "." & App.Minor, "Host", Host
    SaveSetting App.Title, App.Major & "." & App.Minor, "Username", Username
    SaveSetting App.Title, App.Major & "." & App.Minor, "Password", StringNot(Password)
    SaveSetting App.Title, App.Major & "." & App.Minor, "Target", TargetFile
    
End Sub

Public Sub LoadFTPSettings()
    
    Host = GetSetting(App.Title, App.Major & "." & App.Minor, "Host")
    Username = GetSetting(App.Title, App.Major & "." & App.Minor, "Username")
    Password = StringNot(GetSetting(App.Title, App.Major & "." & App.Minor, "Password"))
    TargetFile = GetSetting(App.Title, App.Major & "." & App.Minor, "Target")
    
    If Host = "" Or Username = "" Or Password = "" Or TargetFile = "" Then
        MsgBox "You must first configure FTP settings.", vbInformation, "ZangCam Pro"
        frmFTPSettings.Show vbModal
        If frmFTPSettings.Cancel Then
            Unload Me
            End
        Else
            With frmFTPSettings
            Username = .txtUsername.Text
            Password = .txtPassword.Text
            Host = .txtHost.Text
            TargetFile = .txtTarget.Text
            End With
            SaveFTPSettings
        End If
            
    End If
    
    
End Sub
