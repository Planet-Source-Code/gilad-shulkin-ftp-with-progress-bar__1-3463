VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ftpList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FTP Transfer"
   ClientHeight    =   5088
   ClientLeft      =   1392
   ClientTop       =   1608
   ClientWidth     =   6876
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5088
   ScaleWidth      =   6876
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCopyTo 
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
      Left            =   3480
      TabIndex        =   2
      Top             =   1800
      Width           =   3132
   End
   Begin VB.ListBox lstFiles 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3696
      ItemData        =   "ftpList.frx":0000
      Left            =   240
      List            =   "ftpList.frx":0002
      TabIndex        =   1
      Top             =   1200
      Width           =   3012
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   445
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   9
      Scrolling       =   1
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4440
      Top             =   480
      _ExtentX        =   804
      _ExtentY        =   804
      _Version        =   393216
      AccessType      =   1
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
      RequestTimeout  =   1800
   End
   Begin VB.Label Label2 
      Caption         =   "To download a file specify output and doubleclick doublckicking on directory name will show directory content."
      Height          =   612
      Left            =   3480
      TabIndex        =   4
      Top             =   3000
      Width           =   3372
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Copy to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3480
      TabIndex        =   3
      Top             =   1320
      Width           =   3132
   End
End
Attribute VB_Name = "ftpList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' lists the files from current directory

Public Sub subShowFiles(var As Variant)

  Dim strArray() As String
  Dim intTemp As Integer
   lstFiles.Clear
   
   strArray = Split(CStr(var), Chr(13) & Chr(10))
   lstFiles.AddItem ("../")  ' to go one level up on non UNIX based stations
   For intTemp = 0 To UBound(strArray)
      lstFiles.AddItem (strArray(intTemp))
   Next
   

End Sub


Private Sub lstFiles_DblClick()

On Error GoTo errorhandler

Dim strFile As String
Dim bolFlag As Boolean

MousePointer = vbHourglass

With Inet1
 If (Left(lstFiles.Text, 2) = "./" Or Left(lstFiles.Text, 3) = "../" Or Right(lstFiles.Text, 1) = "/") Then
   ' Clicking on the directory name
   .Execute , "cd " & lstFiles.Text
   .Execute , "DIR"
   subShowFiles (.GetChunk(1024))
   
 Else       ' clicking on the file name
   bolFlag = True
   strFile = subDetermineOutputFileName
   .Execute , "size " & lstFiles.Text  ' size of file to download
   ProgressBar1.Max = .GetChunk(1024)
   .Execute , "get " & lstFiles.Text & " " & strFile
   .Execute , "pwd"  ' Forcing trappable error
   MsgBox "Download complete"
   ProgressBar1.Value = 0
   bolFlag = False
 End If
End With

MousePointer = vbDefault

errorhandler:
   Select Case Err.Number
   Case 35764        '  Still executes last command
   DoEvents
   If bolFlag Then        ' File transfer
    If Not (Dir(strFile) = "") Then
     ProgressBar1.Value = FileLen(strFile)       ' Updating progress bar
     ProgressBar1.ToolTipText = CInt(ProgressBar1.Value * 100 / ProgressBar1.Max) & "% transmitted"
    End If
   End If
   Resume
  End Select
End Sub

' determines the name of output file since in ftp there are several syntax options

Private Function subDetermineOutputFileName() As String

 If Right(txtCopyTo.Text, 1) = "\" Then
   subDetermineOutputFileName = txtCopyTo.Text & lstFiles.Text
 ElseIf Right(txtCopyTo.Text, 1) = ":" Then
   subDetermineOutputFileName = txtCopyTo.Text & "\" & lstFiles.Text
 Else
   subDetermineOutputFileName = txtCopyTo.Text
 End If
 
End Function
