VERSION 5.00
Begin VB.Form FrmStyle 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Style Settings."
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Font Color:"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
      Begin VB.OptionButton OptFont 
         Caption         =   "From BMP File: (16x30 px)"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   2235
      End
      Begin VB.OptionButton OptFont 
         Caption         =   "Random Color"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton OptFont 
         Caption         =   "Fixed Color:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton FntColor 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame FraBack 
      Caption         =   "BackColor:"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton BckColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblCurrentBck 
         Caption         =   "Current BackColor:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Image PIC 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "FrmStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Dim bLoading As Boolean

Private Sub BckColor_Click()

Dim SColor As CHOOSECOLOR
Dim CC() As Byte
       
SColor.lStructSize = Len(SColor)
SColor.lpCustColors = StrConv(CC, vbUnicode)
SColor.hInstance = App.hInstance
SColor.hwndOwner = Me.hwnd
SColor.flags = 0
If CHOOSECOLOR(SColor) <> 0 Then
    BckColor.BackColor = SColor.rgbResult
End If

End Sub

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdOk_Click()
Dim I As Integer

lBackColor = BckColor.BackColor
lForeColor = FntColor.BackColor

For I = 0 To 2
If OptFont(I).Value = True Then
    StereoMode = I
    Exit For
End If
Next

Unload Me
End Sub

Private Sub FntColor_Click()
Dim SColor As CHOOSECOLOR
Dim CC() As Byte
       
SColor.lStructSize = Len(SColor)
SColor.lpCustColors = StrConv(CC, vbUnicode)
SColor.hInstance = App.hInstance
SColor.hwndOwner = Me.hwnd
SColor.flags = 0
If CHOOSECOLOR(SColor) <> 0 Then
    FntColor.BackColor = SColor.rgbResult
End If

End Sub

Private Sub Form_Load()

bLoading = True

If lBackColor = -1 Then
    BckColor.BackColor = vbWhite
Else
    BckColor.BackColor = lBackColor
End If

If lForeColor = -1 Then
    FntColor.BackColor = 0
Else
    FntColor.BackColor = lForeColor
End If

Dim I As Integer
For I = 0 To 2
    If StereoMode = I Then
        OptFont(I).Value = True
    Else
        OptFont(I).Value = False
    End If
Next

bLoading = False
End Sub

Private Sub OptFont_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index <> 2 Or bLoading = True Then Exit Sub

Dim CD As New FileCommonDialog
Dim sfile As String

CD.WindowTitle = "Choose Background Picture"
CD.Filter = "Bitmap Files (*.bmp)|*.bmp"
CD.DefaultExtension = "*.*"
If IsNull(CD.InitialDirectory) Then
    CD.InitialDirectory = App.Path
End If

Me.Enabled = False
sfile = CD.GetFileOpenName()
Me.Enabled = True
sfile = Replace(sfile, Chr(0), "")

If Replace(sfile, " ", "") = "" Then Exit Sub

Dim I As Integer
I = 1
Do While Left(Right(sfile, I), 1) = " "
    I = I + 1
Loop

sfile = Left(sfile, Len(sfile) - I + 1)

If FileExist(sfile) Then
    PIC.Picture = LoadPicture(sfile)
    DoEvents
    If PIC.Width / Screen.TwipsPerPixelX <> 16 Or PIC.Height / Screen.TwipsPerPixelY <> 30 Then
        MsgBox "The picture must have width 16px and height 30px", vbOKOnly + vbExclamation, "ERROR!"
        OptFont(2).Value = False
        Exit Sub
    End If
    Set pMINI = LoadPicture(sfile)
    OptFont(2).Value = True
Else
    OptFont(2).Value = False
End If


End Sub

