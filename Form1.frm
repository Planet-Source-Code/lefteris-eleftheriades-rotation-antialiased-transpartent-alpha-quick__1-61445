VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Rotate, Scale, Alpha"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   318
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   259
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox S 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   1215
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   0
      Top             =   1185
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   15
      Picture         =   "Form1.frx":6C12
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   6
      Top             =   30
      Width           =   3840
   End
   Begin VB.PictureBox Picture1 
      Height          =   870
      Left            =   30
      ScaleHeight     =   810
      ScaleWidth      =   3735
      TabIndex        =   1
      Top             =   3885
      Width           =   3795
      Begin VB.HScrollBar vScroll1 
         Height          =   270
         Left            =   645
         Max             =   360
         TabIndex        =   4
         Top             =   0
         Width           =   2895
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   270
         Index           =   0
         Left            =   645
         Max             =   100
         TabIndex        =   3
         Top             =   285
         Value           =   100
         Width           =   2895
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   270
         Index           =   1
         Left            =   645
         Max             =   100
         Min             =   1
         TabIndex        =   2
         Top             =   585
         Value           =   50
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   $"Form1.frx":16DC0
         Height          =   570
         Left            =   150
         TabIndex        =   5
         Top             =   135
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Rad As Currency = 1.74532925199433E-02
Private Sub Form_Load()
vScroll1_Change
If App.LogMode = 0 Then Me.Caption = "Compile for best results!!!"
End Sub

Private Sub HScroll1_Change(index%)
vScroll1_Change
End Sub

Private Sub HScroll1_Scroll(index%)
vScroll1_Change
End Sub

Private Sub vScroll1_Change()
DoEvents
Picture2.Cls
RotBlt Picture2.hdc, vScroll1.Value * Rad, 128, 128, S.Width, S.Height, S.Image.handle, &HFF00FF, HScroll1(0).Value / 100, HScroll1(1).Value / 50
Picture2.Refresh
End Sub

Private Sub vScroll1_Scroll()
vScroll1_Change
End Sub
