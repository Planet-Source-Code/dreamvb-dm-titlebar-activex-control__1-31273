VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   1860
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   3900
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   2610
      TabIndex        =   4
      Top             =   1320
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   300
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   300
      Left            =   330
      TabIndex        =   2
      Top             =   1320
      Width           =   945
   End
   Begin Project1.DMTitleBar DMTitleBar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cool Title bar"
      ForeColor       =   65535
      MinsizeButToolTipText=   "Even has tool tips"
      CloseButToolTipText=   "Click here to exit.."
      TitleBarIcon    =   "demo.frx":0000
      TitleIconToolTipText=   "This some text"
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "demo.frx":059A
      Top             =   495
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "You could use this Title-bar Control for a custom message box."
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   750
      TabIndex        =   1
      Top             =   510
      Width           =   2715
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Function addslash(lzpath As String) As String
    If Right(lzpath, 1) = "\" Then addslash = lzpath Else addslash = lzpath & "\"
    
End Function

Private Sub DMTitleBar1_ButTitleIcon()
    MsgBox "This is were you can add a menu"
    
End Sub

Private Sub DMTitleBar1_Buttonclose(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgBox "Hope you like this example and please vote for me", vbInformation
    End
    
    
End Sub

Private Sub Form_Load()
    DMTitleBar1.LoadSkin addslash(App.Path) & "skin\main.ini"
    
End Sub
