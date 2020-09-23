VERSION 5.00
Begin VB.UserControl Skin 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   DataSourceBehavior=   1  'vbDataSource
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   2730
   ScaleWidth      =   7920
   Begin VB.PictureBox Skinpic 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   1350
      Picture         =   "Skin.ctx":0000
      ScaleHeight     =   795
      ScaleWidth      =   4275
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   990
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.PictureBox TitleBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7935
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title Bar"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   60
         Width           =   585
      End
   End
End
Attribute VB_Name = "Skin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------
'
'
'           Author : Rajeev P
'           Email ID : Rajeev_Punnalil@hotmail.com
'
'           All of u might have used vbskinner which is not free
'           for pro version .Here i have included all skinner featured
'           Except for rounded edges . If u guys have any suggestions
'           please contact me at Rajeev_punnalil@hotmail.com. U may make
'           and may redistribute this code as long as this commented lines
'           are retainded in all of them.
'
'----------------------------------------------------------------------------------------------------------
'           Note : Retain The above lines in all redistributed versions
'
'
'           This code uses skins from vbskinner so u can go there and download
'           more skin files if u want . Enjoy!
'
'           IMPORTANT !
'           ------------
'           Remember to change form borderstyle to 0-none
'           Use 'send to back' on the skinner actvex
'
'           Improvements over last version
'           ---------------
'           Thanks a lot for ur wonderful response which made me lookback to my code
'           and i found some wonderful imporvements over the last one!
'           1) It doesnot use iterative method any more and hence it is really fast now
'               ,Thanks to the guy who suggested c++ which made me think about optimising the code
'           2) Form resize has been added ,Thanks to
'
'           Thanks
'           ------
'               Thanks for all ur suggestions . A special thanks to merlin,he got
'           the idea of the project correct ! plz add suggestions !
'----------------------------------------------------------------------------------------------------------
Option Explicit
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



Private Type ButtonLocation
    Close_Left As Integer
    Close_Right As Integer
    Min_Left As Integer
    Min_Right As Integer
    Max_Left As Integer
    Max_Right As Integer
End Type

Private skinned As Boolean
Private skinfile As String
Private Bool_Min As Boolean
Private Bool_Max As Boolean
Private frm As Form
Private initok As Boolean
Private locate As ButtonLocation
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Private inrgn As Boolean

Private Sub Init_PaintSkin() ' Runs Routines which paints the skin
If initok = True Then
    UserControl.Cls
    DrawTitleBar
    Draw_Close_Defaults
    Draw_Min_Defaults
    Draw_Max_Defaults frm.WindowState
    Draw_BackGround
End If
End Sub
Public Sub ChangeSkin(filename As String) ' For changing skin during run time
    If initok = False Then Exit Sub
    SaveSetting frm.Caption, "skin", "main", filename
    On Error Resume Next
    If Not filename = "" Then
        Skinpic.Picture = LoadPicture(filename)
    End If
    Init_PaintSkin
End Sub
Private Sub InitAllocate() 'avoiding some redundant code from allocating effort to increase speed
    frm.BorderStyle = 0
    TitleBar.Top = 0
    TitleBar.Left = 0
    CapLabel.Top = 75
    CapLabel.Caption = frm.Caption
    TitleBar.Height = 300
End Sub
Public Sub allocate() ' allocate locations of various buttons depending on settings
    Dim level As Integer
    UserControl.Width = frm.Width
    UserControl.Height = frm.Height
    TitleBar.Width = frm.Width
    locate.Close_Left = frm.Width - 300
    locate.Close_Right = frm.Width - 105
    
    If Bool_Max Then 'allocate Max button if present
        locate.Max_Left = frm.Width - 510
        locate.Max_Right = frm.Width - 330
        level = level + 1
    End If
    
    If Bool_Min Then 'allocate min button if present depend on weather maxbutton is present or not
        If level = 1 Then
            locate.Min_Left = frm.Width - 720
            locate.Min_Right = frm.Width - 480
        Else
            locate.Min_Left = frm.Width - 510
            locate.Min_Right = frm.Width - 330
        End If
    End If
    Init_PaintSkin
    
End Sub

Public Function GetSkinTheme() As Long ' Returns usercontrol's backcolor
    GetSkinTheme = UserControl.BackColor
End Function
Private Sub Draw_BackGround()
    'This Part has been drastically improved from last version
    'iteration techniques are avoided producing faster o/p
If initok = True Then
    Skinpic.ScaleMode = 3
    UserControl.BackColor = GetPixel(Skinpic.hdc, 2400 / 15, 0)
    Skinpic.ScaleMode = 1
    UserControl.PaintPicture Skinpic.Picture, 0, 0, 75, frm.ScaleHeight, 2370, 240, 75, 240
    UserControl.PaintPicture Skinpic.Picture, frm.ScaleWidth - 75, 0, 75, frm.ScaleHeight, 2955, 210, 75, 285
    UserControl.PaintPicture Skinpic.Picture, 0, TitleBar.Height, frm.ScaleWidth, 75, 2595, 0, 270, 75
    UserControl.PaintPicture Skinpic.Picture, 0, frm.ScaleHeight - 75, frm.ScaleWidth, 75, 2595, 645, 270, 75
Else
    MsgBox "Add The Following Lines To your code" & vbNewLine & "skinner1.init_skin me", vbOKOnly
End If
End Sub
Private Sub DrawTitleBar() 'Draws Title Bar, iteration avoided !
If initok = True Then
    TitleBar.PaintPicture Skinpic.Picture, 0, 0, frm.ScaleWidth, 375, 300, 210, 270, 435
Else
    MsgBox "Add The Following Lines To your code" & vbNewLine & "skinner1.init_skin me", vbOKOnly
End If
End Sub
Private Sub CapLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'Helps in moving the form
If initok = True Then
    TitleBar_MouseMove Button, 0, X, Y
Else
    MsgBox "Add The Following Lines To your code" & vbNewLine & "skinner1.init_skin me", vbOKOnly
End If
End Sub



Private Sub TitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) ' Determines titlebar mousedown
If initok = True Then
    If X > locate.Close_Left And X < locate.Close_Right Then
        Unload frm
    Else
    If X > locate.Min_Left And X < locate.Min_Right Then
        frm.WindowState = 1
        Draw_Min_Defaults
    Else
    If X > locate.Max_Left And X < locate.Max_Right Then
        If frm.WindowState = 0 Then
            UserControl.Cls
            TitleBar.Cls
            frm.WindowState = 2
            allocate
            Init_PaintSkin
        Else
            TitleBar.Cls
            UserControl.Cls
            frm.WindowState = 0
            allocate
            Init_PaintSkin
        End If
    Draw_Max_Defaults frm.WindowState
    End If
    End If
    End If
Else
    MsgBox "Add The Following Lines To your code" & vbNewLine & "skinner1.init_skin me", vbOKOnly
End If
End Sub

Private Sub ResetBar() 'Resetes to default
If initok = True Then
    Draw_Close_Defaults
    Draw_Min_Defaults
    Draw_Max_Defaults frm.WindowState
Else
    MsgBox "Add The Following Lines To your code" & vbNewLine & "skinner1.init_skin me", vbOKOnly
End If
End Sub


Private Sub Draw_Close_Defaults() 'Draws close default
If locate.Close_Right = 0 Then Exit Sub
If initok = True Then
    TitleBar.PaintPicture Skinpic.Picture, locate.Close_Left, 90, 195, 195, 0, 0, 195, 195
Else
    MsgBox "Add The Following Lines To your code" & vbNewLine & "skinner1.init_skin me", vbOKOnly
End If
End Sub
Private Sub Draw_Close_Move() 'Draws close mouse over
If locate.Close_Right = 0 Then Exit Sub
If initok = True Then
    TitleBar.PaintPicture Skinpic.Picture, locate.Close_Left, 90, 195, 195, 210, 0, 195, 195
Else
    MsgBox "Add The Following Lines To your code" & vbNewLine & "skinner1.init_skin me", vbOKOnly
End If
End Sub
Private Sub Draw_Close_Down() 'Draws Close Mouse Down
If locate.Close_Right = 0 Then Exit Sub
If initok = True Then
    TitleBar.PaintPicture Skinpic.Picture, locate.Close_Left, 90, 195, 195, 420, 0, 195, 195
Else
    MsgBox "Add The Following Lines To your code" & vbNewLine & "skinner1.init_skin me", vbOKOnly
End If
End Sub

Private Sub Draw_Min_Defaults() 'Draws Min Default
If locate.Min_Right = 0 Then Exit Sub
If initok = True Then
    TitleBar.PaintPicture Skinpic.Picture, locate.Min_Left, 90, 195, 195, 1890, 0, 195, 195
Else
    MsgBox "Add The Following Lines To your code" & vbNewLine & "skinner1.init_skin me", vbOKOnly
End If
End Sub
Private Sub Draw_Min_Move() 'Draws min mouse move
If locate.Min_Right = 0 Then Exit Sub
If initok = True Then
    TitleBar.PaintPicture Skinpic.Picture, locate.Min_Left, 90, 195, 195, 2100, 0, 195, 195
Else
    MsgBox "Add The Following Lines To your code" & vbNewLine & "skinner1.init_skin me", vbOKOnly
End If
End Sub
Private Sub Draw_Max_Defaults(ByVal State As Integer) 'Draws max defaults
If locate.Max_Right = 0 Then Exit Sub
If State = 2 Then
    TitleBar.PaintPicture Skinpic.Picture, locate.Max_Left, 90, 195, 195, 630, 0, 195, 195
Else
    TitleBar.PaintPicture Skinpic.Picture, locate.Max_Left, 90, 195, 195, 1260, 0, 195, 195
End If
End Sub
Private Sub Draw_Max_Move(ByVal State As Integer) 'Draws max mouse move
If locate.Max_Right = 0 Then Exit Sub
If State = 2 Then
    TitleBar.PaintPicture Skinpic.Picture, locate.Max_Left, 90, 195, 195, 840, 0, 195, 195
Else
    TitleBar.PaintPicture Skinpic.Picture, locate.Max_Left, 90, 195, 195, 1470, 0, 195, 195
End If
End Sub

Private Sub TitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) ' Title Bar Mouse move
If initok = True Then
If X > locate.Close_Left And X < locate.Close_Right Then
    Draw_Close_Move
    Draw_Min_Defaults
    Draw_Max_Defaults frm.WindowState
Else
    Draw_Close_Defaults
If X > locate.Min_Left And X < locate.Min_Right Then
    Draw_Min_Move
    Draw_Max_Defaults frm.WindowState
    Draw_Close_Defaults
Else
If X > locate.Max_Left And X < locate.Max_Right Then
    Draw_Max_Move frm.WindowState
    Draw_Min_Defaults
    Draw_Close_Defaults
Else
    ReleaseCapture
    SendMessage frm.hwnd, &HA1, 2, 0
End If
End If
End If
Else
    MsgBox "Add The Following Lines To your code" & vbNewLine & "skinner1.init_skin me", vbOKOnly
        
End If
End Sub



Private Sub UserControl_Initialize()
    TitleBar.Width = UserControl.Width
End Sub

' Various propertis of skinner
Public Property Let ButtonMin(Temp As Boolean)
    Bool_Min = Temp
End Property

Public Property Let ButtonMax(Temp As Boolean)
    Bool_Max = Temp
End Property
Public Property Let Caption(Temp As String)
    CapLabel = Temp
End Property
Public Property Get ButtonMin() As Boolean
    ButtonMin = Bool_Min
End Property

Public Property Get ButtonMax() As Boolean
    ButtonMax = Bool_Max
End Property

Public Property Get Caption() As String
    Caption = CapLabel.Caption
End Property
Public Property Set Skin(Skn As Picture)
    Set Skinpic.Picture = Skn
    If initok Then Init_PaintSkin
End Property
Public Property Get Skin() As Picture
       Set Skin = Skinpic.Picture
End Property


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
'Form resize added in routine usercontrol - mouseup,mousemove!
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If X > UserControl.Width - 200 Or Y > UserControl.Height / 15 - 200 Then
            inrgn = True
        Else
            If Not inrgn Then
                ReleaseCapture
                SendMessage frm.hwnd, &HA1, 2, 0
            End If
        End If
    Else
        If X > UserControl.Width - 200 And Y > UserControl.Height - 200 Then
        UserControl.MousePointer = 8
    Else
        If X > UserControl.Width - 200 And Y < UserControl.Height - 200 Then
            UserControl.MousePointer = 9
        Else
            If X < UserControl.Width - 200 And Y > UserControl.Height - 200 Then
                UserControl.MousePointer = 7
            Else
                UserControl.MousePointer = 0
            End If
        End If
    End If
        RaiseEvent MouseMove(Button, Shift, X, Y)
    End If
    ResetBar
End Sub



Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
       On Error Resume Next
       If inrgn = True Then
            If UserControl.MousePointer = 8 Or UserControl.MousePointer = 9 Then frm.Width = X
            If UserControl.MousePointer = 8 Or UserControl.MousePointer = 7 Then frm.Height = Y
            allocate
        inrgn = False
       End If
       UserControl.MousePointer = 0
End Sub

Private Sub UserControl_Resize()
    TitleBar.Height = 300
End Sub

' Property Bag Additions
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 With PropBag
    Call .ReadProperty("Caption", "")
    Bool_Max = .ReadProperty("ButtonMax", True)
    Bool_Min = .ReadProperty("ButtonMin", True)
    Set Skinpic = .ReadProperty("Skin", Skinpic.Picture)
    
 End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
    Call .WriteProperty("Caption", CapLabel.Caption, "")
    Call .WriteProperty("Skin", Skinpic, Skinpic.Picture)
    Call .WriteProperty("ButtonMax", Bool_Max, True)
    Call .WriteProperty("ButtonMin", Bool_Min, True)
    
End With
End Sub

'Initializes with form control .. Vital part of the code

Public Sub Init_Skin(frm1 As Form)
    initok = True
    Set frm = frm1
    Dim filename As String
    filename = GetSetting(frm.Caption, "skin", "main", "")
    InitAllocate
    
    allocate
    If Not filename = "" Then ChangeSkin filename
End Sub

