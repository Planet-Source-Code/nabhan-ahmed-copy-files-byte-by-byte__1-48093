VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy File"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton CmdCopy 
      Caption         =   "Copy"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   5415
   End
   Begin VB.TextBox TxtDest 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.TextBox TxtSrc 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Destination :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Source       :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded by   : Nabhan Ahmed
'Date       : 8-25-2003
'Description: This program shows you how to copy a file byte by byte.
'             It reads 4 kbs from the source file and write them in
'             the destination file until it reads all byte in the source
'             file. There a bar that shows the copying progress.


Private Sub CmdCopy_Click()
On Error GoTo CopyErr

'Declare variables
Dim SrcFile As String
Dim DestFile As String
Dim SrcFileLen As Long
Dim nSF, nDF As Integer
Dim Chunk As String
Dim BytesToGet As Integer
Dim BytesCopied As Long

'The source file the you want to copy
SrcFile = TxtSrc
'The destination file name
DestFile = TxtDest

'Get source file length
SrcFileLen = FileLen(SrcFile)
'Progress bar settings
ProgressBar1.Min = 0
ProgressBar1.Max = SrcFileLen

'Open both files
nSF = 1
nDF = 2
Open SrcFile For Binary As nSF
Open DestFile For Binary As nDF

'How many bytes to get each time
BytesToGet = 4096 '4kb
BytesCopied = 0
'Show Progress
ProgressBar1.Value = 0
'ProgressBar1.Visible = True

'Keep copying until finishing all bytes
Do While BytesCopied < SrcFileLen
    'Check how many bytes left
    If BytesToGet < (SrcFileLen - BytesCopied) Then
        'Copy 4 KBytes
        Chunk = Space(BytesToGet)
        Get #nSF, , Chunk
    Else
        'Copy the rest
        Chunk = Space(SrcFileLen - BytesCopied)
        Get #nSF, , Chunk
    End If
    BytesCopied = BytesCopied + Len(Chunk)
    
    'Show progress
    ProgressBar1.Value = BytesCopied
        
    'Put data in destination file
    Put #nDF, , Chunk
Loop

'Hide progress bar
ProgressBar1.Value = 0
'ProgressBar1.Visible = False

'Close files
Close #nSF
Close #nDF
Exit Sub

CopyErr:
MsgBox Err.Description, vbCritical, "Error"
End Sub

