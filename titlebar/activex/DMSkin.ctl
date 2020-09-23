VERSION 5.00
Begin VB.UserControl DMTitleBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4785
   ScaleHeight     =   24
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   319
   Begin VB.PictureBox resbutclose 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   1335
      TabIndex        =   3
      Top             =   2205
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox resbutmin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   1335
      TabIndex        =   2
      Top             =   1905
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox restop 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   1335
      TabIndex        =   1
      Top             =   1605
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picTitleBar 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   319
      TabIndex        =   0
      Top             =   0
      Width           =   4785
      Begin VB.PictureBox PicMin 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4170
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   5
         Top             =   60
         Width           =   255
      End
      Begin VB.PictureBox PicClose 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4455
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   4
         Top             =   60
         Width           =   255
      End
      Begin VB.Image picicon 
         Height          =   240
         Left            =   60
         Top             =   60
         Width           =   240
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DM Skin ActiveX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   375
         TabIndex        =   6
         Top             =   60
         Width           =   1380
      End
   End
End
Attribute VB_Name = "DMTitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Title bar Skin ActiveX Doucement by dreamvb

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()


Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Error_Description As String

Enum CaptionTextAlignment
    vsLeft = 0
    vsCentere = 1
    vsRight = 2
End Enum

Public Enum SkinBorder
    vsNone = 0
    vsFixedSingle = 1
End Enum

Private Type SkinProp
    TopSkin As String
    ButtonMinSkin As String
    ButtonCloseSkin As String
    TopBarSkinHeight As Integer
    ButtonWidth As Integer
    ButtonTop As Integer
    ButtonHeight As Integer
    ButtonsDownPos As Integer
    ButtonsUpPos As Integer
    BaseColour As String
    CaptionForeCol As String
    TMaskOnOff As Boolean
    TMask1 As String
    TMask2 As String
    RoundWindows As String
End Type

Private SkinProp As SkinProp

Event Buttonclose(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Buttonminsize(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Event Declarations:
Event ButTitleIcon() 'MappingInfo=picicon,picicon,-1,Click
Attribute ButTitleIcon.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."

'Event IconButton(Button As Integer, Shift As Integer, X As Single, Y As Single)
Function HideButtons(nShow As Boolean)
    If nShow Then
        PicClose.Visible = False
        PicMin.Visible = False
    Else
        PicMin.Visible = True
        PicClose.Visible = True
    End If
    
End Function
Function GetPathFromFile(lzFilePath As String) As String
Dim I As Integer, IPart As Integer
    lzFilePath = Trim(lzFilePath)
    For I = Len(lzFilePath) To 1 Step -1
        ch = Mid(lzFilePath, I, 1)
        If ch = "\" Then IPart = I: Exit For
    Next
    GetPathFromFile = Left(lzFilePath, IPart)
    
End Function
Function GetRGB(StrRGB As String) As Long
Dim VRGB As Variant
    If Trim(Len(StrRGB)) <= 0 Then StrRGB = "0,0,0"
    VRGB = Split(StrRGB, ",")
    If VRGB(0) <= 0 Then VRGB(0) = 0
    If VRGB(0) > 255 Then VRGB(0) = 0
    
    If VRGB(1) <= 0 Then VRGB(1) = 0
    If VRGB(1) > 255 Then VRGB(1) = 0
    
    If VRGB(2) <= 0 Then VRGB(2) = 0
    If VRGB(2) > 255 Then VRGB(2) = 0
    GetRGB = RGB(VRGB(0), VRGB(1), VRGB(2))
    
End Function
Public Function ReadConfig(AppName As String, StrKey As String, lzFileName As String) As String
Dim StrBuff As String
Dim Xpos As Integer
    StrBuff = String(255, Chr(0))
    GetPrivateProfileString AppName, StrKey, "ERROR", StrBuff, 255, lzFileName
    ReadConfig = Left(StrBuff, InStr(StrBuff, Chr(0)) - 1)
    
End Function

Function LoadSkin(SkinIni As String)
Dim ICnt As Integer
    If FindFile(SkinIni) = False Then
        Error_Description = "unable to locatate main skin ini"
        Exit Function
    Else
        ' This loads in all the skins that we need
        SkinProp.TopSkin = ReadConfig("Main", "SknTop", SkinIni)
        SkinProp.ButtonMinSkin = ReadConfig("Main", "ButMin", SkinIni)
        SkinProp.ButtonCloseSkin = ReadConfig("Main", "ButClose", SkinIni)
        
        ' Check that the files are in the Main Skin ini file
        If Len(Trim(SkinProp.TopSkin)) <= 0 Then
            Error_Description = "unable to locate Top Bar value in ini selection"
            Exit Function
        ElseIf Len(Trim(SkinProp.ButtonMinSkin)) <= 0 Then
            Error_Description = "unable to locate Button Minsize Bar value in ini selection"
            Exit Function
        ElseIf Len(Trim(SkinProp.ButtonCloseSkin)) <= 0 Then
            Error_Description = "unable to locate Button Close value in ini selection"
            Exit Function
        End If
        
        ' Check that each file exists
        If FindFile(GetPathFromFile(SkinIni) & SkinProp.TopSkin) = False Then
            Error_Description = "unable to locate top bar skin"
            Exit Function
        ElseIf FindFile(GetPathFromFile(SkinIni) & SkinProp.ButtonMinSkin) = False Then
            Error_Description = "unable to locate minsize button skin"
            Exit Function
        ElseIf FindFile(GetPathFromFile(SkinIni) & SkinProp.ButtonCloseSkin) = False Then
            Error_Description = "unable to locate button close skin"
            Exit Function
        End If
        
        ' This loads in the skin width and height values
        
        SkinProp.TopBarSkinHeight = Val(ReadConfig("Skins.sizes", "SknTopHeight", SkinIni))
        SkinProp.ButtonWidth = Val(ReadConfig("Skins.sizes", "Butwidth", SkinIni))
        SkinProp.ButtonHeight = Val(ReadConfig("Skins.sizes", "Butheight", SkinIni))
        SkinProp.ButtonTop = Val(ReadConfig("Skins.sizes", "Buttop", SkinIni))
        
        SkinProp.RoundWindows = UCase(ReadConfig("Main", "RoundWnd", SkinIni))
        If SkinProp.RoundWindows = "YES" Then MakeRoundWindow UserControl.Parent.hwnd
        ' This load the minsize and close button click postions
        
        SkinProp.ButtonsDownPos = Val(ReadConfig("Button.Events", "ButtonDownPos", SkinIni))
        SkinProp.ButtonsUpPos = Val(ReadConfig("Button.Events", "ButtonUpPos", SkinIni))
        ' This loads in all the colour values for caption base colour masking colours etc.
        
        SkinProp.TMask1 = ReadConfig("Base.Colours", "Mask1", SkinIni)
        SkinProp.TMask2 = ReadConfig("Base.Colours", "Mask2", SkinIni)
        
        restop.Picture = LoadPicture(GetPathFromFile(SkinIni) & SkinProp.TopSkin) ' loads top bar
        resbutmin.Picture = LoadPicture(GetPathFromFile(SkinIni) & SkinProp.ButtonMinSkin) ' loads minsize buttons
        resbutclose.Picture = LoadPicture(GetPathFromFile(SkinIni) & SkinProp.ButtonCloseSkin) ' loads close buttons

        ' Setup all the picture boxes widths & heights
        
        picTitleBar.Height = SkinProp.TopBarSkinHeight
        
        ' Tool bar buttons heights and widths
        PicMin.Width = SkinProp.ButtonWidth
        PicMin.Height = SkinProp.ButtonHeight
        PicClose.Width = SkinProp.ButtonWidth
        PicClose.Height = SkinProp.ButtonHeight
        PicMin.Top = SkinProp.ButtonTop
        PicClose.Top = SkinProp.ButtonTop
        UserControl.Parent.ScaleMode = 3
        
        ' Check if masking colour is enabled
        If UCase(ReadConfig("Base.Colours", "UseMask", SkinIni)) = "YES" Then SkinProp.TMaskOnOff = True
        
        For ICnt = 0 To picTitleBar.Width
            ' this will apply the top bar and botton bar
            BitBlt picTitleBar.hDC, ICnt, 0, restop.Width, SkinProp.TopBarSkinHeight, restop.hDC, 0, 0, vbSrcCopy
        Next
        
        ICnt = 0 ' reset counter
        
        BitBlt PicClose.hDC, 0, 0, SkinProp.ButtonHeight, SkinProp.ButtonWidth, resbutclose.hDC, 0, 0, vbSrcCopy ' Close button
        BitBlt PicMin.hDC, 0, 0, SkinProp.ButtonHeight, SkinProp.ButtonWidth, resbutmin.hDC, 0, 0, vbSrcCopy  ' Minsize button
        
        ' Refresh all picture boxes
        picTitleBar.Refresh
        PicClose.Refresh
        PicMin.Refresh
    End If
    
End Function
Private Function FixPath(lzpath As String) As String
    If Right(lzpath, 1) = "\" Then FixPath = lzpath Else FixPath = lzpath & "\"
    
End Function

Private Function FindFile(lzFileName As String) As Boolean
    If Dir(lzFileName) = "" Then FindFile = False Else FindFile = True
    
End Function
Private Function MakeRoundWindow(lzHangle As Long)
    Dim Rw As Long
    UserControl.Parent.ScaleMode = vbPixels
    Rw = CreateRoundRectRgn(0, 0, UserControl.Parent.ScaleWidth, UserControl.Parent.ScaleHeight, 20, 20) ' Create the window handle
    SetWindowRgn lzHangle, Rw, True ' set up the window and displays it
    DeleteObject Rw ' Delete the object
    
End Function
Private Function MoveWindowFrm(lzHangle As Long)
Dim ReVal As Long
    ReleaseCapture
    ReVal = SendMessage(lzHangle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    
End Function

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        MoveWindowFrm UserControl.Parent.hwnd
    End If
    
End Sub

Private Sub PicClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        BitBlt PicClose.hDC, 0, 0, SkinProp.ButtonHeight, SkinProp.ButtonWidth, resbutclose.hDC, SkinProp.ButtonWidth + SkinProp.ButtonsDownPos, 0, vbSrcCopy ' Close button
        PicClose.Refresh
    End If
    
End Sub

Private Sub PicClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        RaiseEvent Buttonclose(Button, Shift, X, Y)
        BitBlt PicClose.hDC, 0, 0, SkinProp.ButtonHeight, SkinProp.ButtonWidth, resbutclose.hDC, SkinProp.ButtonsUpPos, 0, vbSrcCopy ' Close button
        PicClose.Refresh
    End If

End Sub

Private Sub PicMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        BitBlt PicMin.hDC, 0, 0, SkinProp.ButtonHeight, SkinProp.ButtonWidth, resbutmin.hDC, SkinProp.ButtonWidth + SkinProp.ButtonsDownPos, 0, vbSrcCopy
        PicMin.Refresh
    End If
    
End Sub

Private Sub PicMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        RaiseEvent Buttonminsize(Button, Shift, X, Y)
        BitBlt PicMin.hDC, 0, 0, SkinProp.ButtonHeight, SkinProp.ButtonWidth, resbutmin.hDC, SkinProp.ButtonsUpPos, 0, vbSrcCopy
        PicMin.Refresh
    End If

End Sub

Private Sub picTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        MoveWindowFrm UserControl.Parent.hwnd
    End If
    
End Sub

Private Sub UserControl_Resize()
    PicClose.Left = picTitleBar.Width - PicClose.Width - 8
    PicMin.Left = PicClose.Left - PicMin.Width - 4
    UserControl.Height = 345
    
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Sub Size(ByVal Width As Single, ByVal Height As Single)
Attribute Size.VB_Description = "Changes the width and height of a User Control."
    UserControl.Size Width, Height
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "DM Skin ActiveX")
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
    PicMin.ToolTipText = PropBag.ReadProperty("MinsizeButToolTipText", "")
    PicClose.ToolTipText = PropBag.ReadProperty("CloseButToolTipText", "")
    Set TitleBarIcon = PropBag.ReadProperty("TitleBarIcon", Nothing)
    
    picicon.ToolTipText = PropBag.ReadProperty("TitleIconToolTipText", "")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "DM Skin ActiveX")
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("MinsizeButToolTipText", PicMin.ToolTipText, "")
    Call PropBag.WriteProperty("CloseButToolTipText", PicClose.ToolTipText, "")
    Call PropBag.WriteProperty("TitleBarIcon", Nothing)
    Call PropBag.WriteProperty("TitleBarIcon", TitleBarIcon, Nothing)
    Call PropBag.WriteProperty("TitleIconToolTipText", picicon.ToolTipText, "")
End Sub
Public Property Get MinsizeButToolTipText() As String
    MinsizeButToolTipText = PicMin.ToolTipText
End Property

Public Property Let MinsizeButToolTipText(ByVal New_ToolTipText As String)
    PicMin.ToolTipText() = New_ToolTipText
    PropertyChanged "MinsizeButToolTipText"
End Property

Public Property Get CloseButToolTipText() As String
    CloseButToolTipText = PicClose.ToolTipText
End Property

Public Property Let CloseButToolTipText(ByVal New_ToolTipText As String)
    PicClose.ToolTipText() = New_ToolTipText
    PropertyChanged "CloseButToolTipText"
End Property





'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picicon,picicon,-1,Picture
Public Property Get TitleBarIcon() As Picture
Attribute TitleBarIcon.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set TitleBarIcon = picicon.Picture
End Property

Public Property Set TitleBarIcon(ByVal New_Picture As Picture)
    Set picicon.Picture = New_Picture
    PropertyChanged "TitleBarIcon"
End Property

Private Sub picicon_Click()
    RaiseEvent ButTitleIcon
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picicon,picicon,-1,ToolTipText
Public Property Get TitleIconToolTipText() As String
Attribute TitleIconToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    TitleIconToolTipText = picicon.ToolTipText
End Property

Public Property Let TitleIconToolTipText(ByVal New_TitleIconToolTipText As String)
    picicon.ToolTipText() = New_TitleIconToolTipText
    PropertyChanged "TitleIconToolTipText"
End Property

