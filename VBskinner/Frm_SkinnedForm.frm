VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_SkinnedForm 
   BorderStyle     =   0  'None
   Caption         =   "Improved Skinner !"
   ClientHeight    =   6420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   1830
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change skin During Run time"
      Height          =   525
      Left            =   1950
      TabIndex        =   1
      Top             =   2670
      Width           =   2295
   End
   Begin Skinned_Form.Skin Skin1 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   450
      Caption         =   "Title Bar"
      Skin            =   "Frm_SkinnedForm.frx":0000
   End
End
Attribute VB_Name = "Frm_SkinnedForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cd.Filter = "*.bmp"
cd.InitDir = App.Path & "\skins"
cd.ShowOpen
Skin1.ChangeSkin cd.filename
End Sub

Private Sub Form_Load()
Skin1.Init_Skin Me
End Sub
