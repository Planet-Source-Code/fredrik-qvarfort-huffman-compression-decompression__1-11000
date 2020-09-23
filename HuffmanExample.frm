VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Huffman Encoding/Decoding Example Project"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Text            =   "C:\Saol.txt.new"
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Information"
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   4215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Progress:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   18
         Top             =   1450
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<unknown>"
         Height          =   195
         Index           =   4
         Left            =   1920
         TabIndex        =   17
         Top             =   1450
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   290
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   580
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ratio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   870
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time spent:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   13
         Top             =   1160
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<unknown>"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   12
         Top             =   290
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<unknown>"
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   11
         Top             =   580
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<unknown>"
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   10
         Top             =   870
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<unknown>"
         Height          =   195
         Index           =   3
         Left            =   1920
         TabIndex        =   9
         Top             =   1160
         Width           =   840
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decompress"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Text            =   "C:\Saol.huf"
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Text            =   "C:\Saol.txt"
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compress"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Decompress To File:"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compressed File:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Original File:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Huffman As clsHuffman
Attribute Huffman.VB_VarHelpID = -1

Private Sub Command1_Click()

  Dim OldTimer As Single

  On Error GoTo ErrorHandler

  'Store the current timer for later use
  OldTimer = Timer

  'Compress the source file
  Call Huffman.EncodeFile(Text1(0).Text, Text1(1).Text)

  'Update the statistics
  Label2(3).Caption = Timer - OldTimer & " s"
  Label2(0).Caption = FileLen(Text1(0).Text) & " bytes"
  Label2(1).Caption = FileLen(Text1(1).Text) & " bytes"
  Label2(2).Caption = Int(Val(Label2(1).Caption) / Val(Label2(0).Caption) * 100) & "%"

  'Show a nice dialog to the user
  Call MsgBox("Compression successful.", vbInformation)
  
  Exit Sub
  
ErrorHandler:
  Call MsgBox("The compression was not successful. Something went terribly wrong." & vbCrLf & vbCrLf & Err.Description, vbExclamation)

End Sub
Private Sub Command2_Click()

  Dim Filenr As Integer
  Dim OldTimer As Single

  On Error GoTo ErrorHandler
  
  'Store the time for later use
  OldTimer = Timer
  
  'Uncompress the compressed file
  Call Huffman.DecodeFile(Text1(1).Text, Text1(2).Text)
  
  'Update decompression statistics
  Label2(3).Caption = Timer - OldTimer & " s"
  Label2(0).Caption = FileLen(Text1(1).Text) & " bytes"
  Label2(1).Caption = FileLen(Text1(2).Text) & " bytes"
  Label2(2).Caption = Int(Val(Label2(1).Caption) / Val(Label2(0).Caption) * 100) & "%"
  
  'Show a nice dialog to the user
  Call MsgBox("Decompression successful.", vbInformation)
  
  Exit Sub
  
ErrorHandler:
  Call MsgBox("The decompression was not successful. Something went terribly wrong." & vbCrLf & vbCrLf & Err.Description, vbExclamation)

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()

  Set Huffman = New clsHuffman
  
End Sub


Private Sub Huffman_Progress(Procent As Integer)

  Label2(4).Caption = Procent & "%"
  'Label2(4).Refresh
  DoEvents
  
End Sub


