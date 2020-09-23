VERSION 5.00
Begin VB.Form ftpConnect 
   Caption         =   "FTP transfer"
   ClientHeight    =   2952
   ClientLeft      =   2148
   ClientTop       =   3144
   ClientWidth     =   6264
   ScaleHeight     =   2952
   ScaleWidth      =   6264
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3720
      TabIndex        =   7
      Top             =   2160
      Width           =   1332
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1080
      TabIndex        =   6
      Top             =   2160
      Width           =   1332
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   2052
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   2052
   End
   Begin VB.TextBox txtFTP 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   2652
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3240
      TabIndex        =   5
      Top             =   1320
      Width           =   2052
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Width           =   2052
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "FTP site"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   2652
   End
End
Attribute VB_Name = "ftpConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()
 
 On Error GoTo errorhandler
 
 Dim varTemp As Variant

  With ftpList.Inet1
   .URL = txtFTP.Text
   .UserName = txtUser.Text
   .Password = txtPassword.Text
   .Execute , "DIR"

   varTemp = .GetChunk(1024)
   Load ftpList
   ftpList.Show
   ftpList.subShowFiles (varTemp)
   Unload Me
  End With
  
errorhandler:
 Select Case Err.Number
 
  Case 35764        '  Still executes last command
   DoEvents
   Resume
   
 End Select
End Sub

Private Sub cmdExit_Click()

 Unload Me
 End

End Sub
