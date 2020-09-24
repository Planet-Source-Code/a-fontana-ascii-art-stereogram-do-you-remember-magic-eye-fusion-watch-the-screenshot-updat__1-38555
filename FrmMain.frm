VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Sterogram Creator!"
   ClientHeight    =   5370
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   358
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   611
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraTools 
      Caption         =   "Tools:"
      Height          =   3015
      Left            =   8040
      TabIndex        =   4
      Top             =   120
      Width           =   975
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   180
         TabIndex        =   10
         ToolTipText     =   "Foreground & Background colors."
         Top             =   2220
         Width           =   615
      End
      Begin VB.Shape cTab 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   240
         Top             =   2280
         Width           =   375
      End
      Begin VB.Image Tool 
         Height          =   480
         Index           =   2
         Left            =   240
         Picture         =   "FrmMain.frx":030A
         ToolTipText     =   "Invert Picture!"
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image Tool 
         Height          =   480
         Index           =   1
         Left            =   240
         Picture         =   "FrmMain.frx":0BD4
         ToolTipText     =   "Fill Region!"
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Tool 
         Height          =   480
         Index           =   0
         Left            =   240
         Picture         =   "FrmMain.frx":189E
         ToolTipText     =   "Pencil (to draw!)"
         Top             =   360
         Width           =   480
      End
      Begin VB.Shape Sh 
         BorderColor     =   &H00808080&
         Height          =   495
         Index           =   1
         Left            =   240
         Top             =   960
         Width           =   495
      End
      Begin VB.Shape Sh 
         BorderColor     =   &H00808080&
         Height          =   495
         Index           =   0
         Left            =   240
         Top             =   360
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   360
         Top             =   2400
         Width           =   375
      End
      Begin VB.Shape Sh 
         BorderColor     =   &H00808080&
         Height          =   495
         Index           =   2
         Left            =   240
         Top             =   1560
         Width           =   495
      End
   End
   Begin VB.Frame FraDeep 
      Caption         =   "Deep:"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   3135
      Begin VB.Label Label2 
         Caption         =   "HIGH"
         Height          =   255
         Left            =   2460
         TabIndex        =   6
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "LOW"
         Height          =   255
         Left            =   285
         TabIndex        =   5
         Top             =   315
         Width           =   345
      End
      Begin VB.Image Sposta 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   2280
         Picture         =   "FrmMain.frx":2568
         Top             =   315
         Width           =   480
      End
      Begin VB.Image Sposta 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   1
         Left            =   720
         Picture         =   "FrmMain.frx":2E32
         Top             =   330
         Width           =   480
      End
      Begin VB.Shape RPos 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   135
         Left            =   960
         Top             =   360
         Width           =   375
      End
      Begin VB.Shape Bar 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   135
         Left            =   960
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame FraPic 
      Caption         =   "Model To Render:"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7755
      Begin VB.PictureBox PicFrame 
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   220
         ScaleHeight     =   3015
         ScaleWidth      =   7275
         TabIndex        =   2
         Top             =   360
         Width           =   7275
         Begin VB.PictureBox Sfondo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3015
            Left            =   0
            ScaleHeight     =   201
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   481
            TabIndex        =   3
            Top             =   0
            Width           =   7215
            Begin VB.Shape Pix 
               FillColor       =   &H00FFFFFF&
               FillStyle       =   0  'Solid
               Height          =   105
               Index           =   0
               Left            =   0
               Top             =   0
               Width           =   90
            End
         End
      End
   End
   Begin VB.Label lblVote 
      BackStyle       =   0  'Transparent
      Caption         =   "here!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7680
      MouseIcon       =   "FrmMain.frx":36FC
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Vote for this code at planet source code right now clicking"
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   4920
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Stereogram Creator. By Andrea Fontana © 2002."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      MouseIcon       =   "FrmMain.frx":384E
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4560
      Width           =   4215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Model..."
      End
      Begin VB.Menu csep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Load Model..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Model..."
         Shortcut        =   ^S
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu CmdImp 
         Caption         =   "&Import Picture..."
      End
      Begin VB.Menu mnuEsporta 
         Caption         =   "Export Picture..."
      End
      Begin VB.Menu Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCrea 
         Caption         =   "&Create Stereogram!"
         Shortcut        =   {F5}
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEsci 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuMod 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCancel 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRip 
         Caption         =   "&Redo"
         Enabled         =   0   'False
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnueditsep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStyle 
         Caption         =   "&Style Settings..."
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "&?"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


''
''   ******
''   IF YOU WANT TO USE COLORIZED CURSOR YOU HAVE TO
''   COMPILE THIS CODE AND NOT TO RUN IT FROM VISUAL BASIC
''   ******
''
''







Dim FraPicWidth As Long     ''  \
Dim FraPicHeight As Long    ''   |-> Infos for main
Dim MinWidth As Long        ''   |-> window's resize
Dim MinHeight As Long       ''  /

Private Const SW_SHOW = 5       ' Displays Window in its current size
                                ' and position
Private Const SW_SHOWNORMAL = 1 ' Restores Window if Minimized or
                                ' Maximized
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
         "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
         String, ByVal lpFile As String, ByVal lpParameters As String, _
         ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias _
         "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
         String, ByVal lpResult As String) As Long


Private Sub CmdImp_Click()
Dim sFile As String
CD.DefaultExtension = "bmp"
CD.Filter = "Pictures (*.bmp,*.jpg,*.gif,*.ico)|*.bmp;*.jpg;*.gif;*.ico"
If IsNull(CD.InitialDirectory) Then
    CD.InitialDirectory = App.Path
End If
CD.WindowTitle = "Open Picture..."
FrmMain.Enabled = False
sFile = CD.GetFileOpenName()
FrmMain.Enabled = True
sFile = Replace(sFile, Chr(0), "")
If Replace(sFile, " ", "") = "" Then Exit Sub
Dim I As Integer
I = 1
Do While Left(Right(sFile, I), 1) = " "
    I = I + 1
Loop

sFile = Left(sFile, Len(sFile) - I + 1)

If FileExist(sFile) = False Then
    MsgBox "File doesn't exists!", vbCritical + vbOKOnly, "Il file non esiste!"
    Exit Sub
End If
Call ImportImage(sFile)

End Sub

Private Sub Form_Load()
Dim I As Integer, K As Integer, T As Integer
Dim hpos As Integer, vpos As Integer

lBackColor = -1
lForeColor = -1

StereoMode = Normal

FraPicWidth = FrmMain.ScaleWidth - FraPic.Width
FraPicHeight = FrmMain.ScaleHeight - FraPic.Height
MinWidth = FrmMain.Width
MinHeight = FrmMain.Height

lDizio = Split(cAlpha, ",")         ''  Create dictionary's and
lLevs = Split(cLevs, ",")           ''  colors' array

Set CD = New FileCommonDialog      ''  New Com. Dlg class
Call Randomize(Timer)              ''  Init random genarator
ReDim sBackup(0)                   ''  Init backup array.

curLev = 3                         ''  Pos on lLevs()
curBackPos = 0                     ''  Pos on sBackup()

For I = 0 To 1                     ''  Scrollbar
    Sposta(I).Width = 120          ''  size
    Sposta(I).Height = 240
Next

Sfondo.Width = nWidth * 6 * Screen.TwipsPerPixelX + 10   ''  Editor
Sfondo.Height = nHeight * 8 * Screen.TwipsPerPixelY + 10  ''  Size

PicFrame.Width = FraPic.Width * Screen.TwipsPerPixelX - 600
PicFrame.Height = FraPic.Height * Screen.TwipsPerPixelY - 600


Pix(0).Width = 7                   ''  "Pixels" width &
Pix(0).Height = 9                  ''  height

'' Pixels' Loader >>>

For K = 0 To nHeight - 1
vpos = 8 * K
hpos = 0
For I = 0 To nWidth - 1
     T = K * nWidth + I
     If T > Pix.UBound Then
     Load Pix(T)
     End If
     Pix(T).Left = hpos
     Pix(T).Top = vpos
     Pix(T).FillColor = &HFFFFFF   ''  The pixel are white!.
     Pix(T).Visible = True
     hpos = hpos + 6               ''  Oriz. Movement
Next
DoEvents
Next

Tool_MouseDown 0, 0, 0, 0, 0              '' Set current tool
cTab.BackColor = CLng(lLevs(curLev - 1))  '' & current color
RPos.Width = Bar.Width


Call Decode(DecodeRLE(cStartUp))        ''  Load Default Picture

End Sub


Private Sub mnuStyle_Click()
FrmStyle.Show 1
End Sub

Private Sub PicFrame_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw x, y, Button, curTool, True
End Sub

Private Sub PicFrame_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw x, y, Button, curTool
End Sub

Private Sub PicFrame_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw x, y, Button, curTool
End Sub

Private Sub Form_Resize()

''' Resize function
''' (not longer used)

If Me.WindowState = vbMinimized Then Exit Sub

If Me.Width < MinWidth Then Me.Width = MinWidth
If Me.Height < MinHeight Then Me.Height = MinHeight

FraPic.Width = FrmMain.ScaleWidth - FraPicWidth
FraPic.Height = FrmMain.ScaleHeight - FraPicHeight

PicFrame.Width = FraPic.Width * Screen.TwipsPerPixelX - 440
PicFrame.Height = FraPic.Height * Screen.TwipsPerPixelY - 600

Sfondo.Left = (PicFrame.Width - Sfondo.Width) / 2
Sfondo.Top = (PicFrame.Height - Sfondo.Height) / 2


FraTools.Left = FraPic.Width + 19

FraDeep.Top = FraPic.Height + 15

Label3.Top = FraDeep.Top + 8
Label4.Top = Label3.Top + 24
lblVote.Top = Label3.Top + 24
End Sub



Private Sub Label3_Click()
OpenLink "http://www.vbp.it/trikko"
End Sub

Private Sub lblVote_Click()
OpenLink "http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=38555&lngWId=1"
End Sub

Private Sub mnuCancel_Click()

curBackPos = curBackPos - 1        ''  Decode previus
Call Decode(sBackup(curBackPos))   ''  picture

If curBackPos = 0 Then mnuCancel.Enabled = False
mnuRip.Enabled = True
End Sub

Private Sub mnuCrea_Click()   '' Create the stereogram
Dim sFile As String
CD.DefaultExtension = "bmp"
CD.Filter = "Stereogram (*.bmp)|*.bmp"
If IsNull(CD.InitialDirectory) Then
    CD.InitialDirectory = App.Path
End If
CD.WindowTitle = "Create stereogram ..."
FrmMain.Enabled = False
sFile = CD.GetFileSaveName()
FrmMain.Enabled = True
sFile = Replace(sFile, Chr(0), "")
If Replace(sFile, " ", "") = "" Then Exit Sub
Dim I As Integer
I = 1
Do While Left(Right(sFile, I), 1) = " "
    I = I + 1
Loop

sFile = Left(sFile, Len(sFile) - I + 1)

Dim sBuffer As String
Dim fId As Integer
If FileExist(sFile) = True Then
    If MsgBox("The file" & vbCrLf & sFile & vbCrLf & "already exists. Overwrite?", vbQuestion + vbYesNo, "Overwrite?") = vbNo Then Exit Sub
    Kill sFile
End If
sBuffer = GeneraStereo()    '' Create Stereogram
Call GeneraBmp(sFile, sBuffer) '' Save steregoram

If MsgBox("Do you want to see the result?", vbYesNo + vbQuestion, "Open Picture.") = vbYes Then
    ShellExecute Me.hwnd, "OPEN", sFile, vbNullString, App.Path, SW_SHOW
End If

End Sub

Private Sub mnuEsci_Click()
Unload Me
End
End Sub

Private Sub mnuEsporta_Click()
Dim sFile As String
CD.DefaultExtension = "bmp"
CD.Filter = "Bitmap (*.bmp)|*.bmp"
If IsNull(CD.InitialDirectory) Then
    CD.InitialDirectory = App.Path
End If
CD.WindowTitle = "Save picture..."
FrmMain.Enabled = False
sFile = CD.GetFileSaveName()
FrmMain.Enabled = True
sFile = Replace(sFile, Chr(0), "")
If Replace(sFile, " ", "") = "" Then Exit Sub
Dim I As Integer
I = 1
Do While Left(Right(sFile, I), 1) = " "
    I = I + 1
Loop

sFile = Left(sFile, Len(sFile) - I + 1)

If FileExist(sFile) = True Then
    If MsgBox("The file" & vbCrLf & sFile & vbCrLf & "already exists. Overwrite?", vbQuestion + vbYesNo, "Overwrite?") = vbNo Then Exit Sub
    Kill sFile
End If
Call ExportImage(sFile)
End Sub

Private Sub mnuInfo_Click()
MsgBox "Stereogram Creator!" & vbCrLf & "By Andrea Fontana © 2002." & vbCrLf & "Homepage: http://www.it.owns.it" & vbCrLf & "E-Mail: Trikko@katamail.com", vbOKOnly + vbInformation, "Credits!"
End Sub

Private Sub mnuNew_Click()
Dim I As Integer, K As Integer, T As Integer
Dim hpos As Integer, vpos As Integer

ReDim sBackup(0)                   ''  Init Backup array

curLev = 3                         ''  Current lLevs() pos
curBackPos = 0                     ''  Current sBackup() pos


'' Pixel Loader >>>

For K = 0 To nHeight - 1
vpos = 6 * K
hpos = 0
For I = 0 To nWidth - 1
     T = K * nWidth + I
     If T > Pix.UBound Then
     Load Pix(T)
     End If
     Pix(T).FillColor = &HFFFFFF   ''  I pixel are white.
     Pix(T).Visible = True
Next
Next

Tool_MouseDown 0, 0, 0, 0, 0            '' Set the current tool
cTab.BackColor = CLng(lLevs(curLev - 1))  '' and current color
RPos.Width = Bar.Width

mnuCancel.Enabled = False
mnuRip.Enabled = False

End Sub

Private Sub mnuOpen_Click()
Dim sFile As String
CD.DefaultExtension = "mdl"
CD.Filter = "Models (*.mdl)|*.mdl"
If IsNull(CD.InitialDirectory) Then
    CD.InitialDirectory = App.Path
End If
CD.WindowTitle = "Open Model..."
FrmMain.Enabled = False
sFile = CD.GetFileOpenName()
FrmMain.Enabled = True
sFile = Replace(sFile, Chr(0), "")
If Replace(sFile, " ", "") = "" Then Exit Sub
Dim I As Integer
I = 1
Do While Left(Right(sFile, I), 1) = " "
    I = I + 1
Loop

sFile = Left(sFile, Len(sFile) - I + 1)

Dim sBuffer As String
Dim fId As Integer
If FileExist(sFile) = False Then
    MsgBox "File doesn't exists!", vbCritical + vbOKOnly, "Il file non esiste!"
    Exit Sub
End If
fId = FreeFile()
Open sFile For Binary Access Read As fId
sBuffer = Space(LOF(fId))
Get fId, 1, sBuffer
Close
sBuffer = DecodeRLE(sBuffer)
Decode sBuffer
End Sub

Private Sub mnuRip_Click()
curBackPos = curBackPos + 1
Decode sBackup(curBackPos)

If curBackPos = UBound(sBackup) Then mnuRip.Enabled = False
mnuCancel.Enabled = True
End Sub

Private Sub mnuSave_Click()
Dim sFile As String
CD.DefaultExtension = "mdl"
CD.Filter = "Models (*.mdl)|*.mdl"
If IsNull(CD.InitialDirectory) Then
    CD.InitialDirectory = App.Path
End If
CD.WindowTitle = "Save Model..."
FrmMain.Enabled = False
sFile = CD.GetFileSaveName()
FrmMain.Enabled = True
sFile = Replace(sFile, Chr(0), "")
If Replace(sFile, " ", "") = "" Then Exit Sub
Dim I As Integer
I = 1
Do While Left(Right(sFile, I), 1) = " "
    I = I + 1
Loop

sFile = Left(sFile, Len(sFile) - I + 1)

Dim sBuffer As String
Dim fId As Integer
If FileExist(sFile) = True Then
    If MsgBox("The file" & vbCrLf & sFile & vbCrLf & "already exists. Overwrite?", vbQuestion + vbYesNo, "Overwrite?") = vbNo Then Exit Sub
    Kill sFile
End If
sBuffer = EncodeRLE(Compatta())
fId = FreeFile()
Open sFile For Binary Access Write As fId
Put fId, 1, sBuffer
Close
End Sub

Private Sub Sfondo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw x, y, Button, curTool, True
End Sub

Private Sub Sfondo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Draw x, y, Button, curTool
End Sub

Private Sub Sfondo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Call BackUp
End Sub

Private Sub Sposta_DblClick(Index As Integer)

If Index = 0 Then
curLev = curLev + 1
Else
curLev = curLev - 1
End If

If curLev < 1 Then curLev = 1
If curLev > 3 Then curLev = 3

RPos.Width = (Bar.Width / 2) * (curLev - 1)  '' Tick on the scrollbar
cTab.BackColor = CLng(lLevs(curLev - 1)) '' Set Fillcolor!
End Sub

Private Sub Sposta_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Index = 0 Then
curLev = curLev + 1
Else
curLev = curLev - 1
End If
If curLev < 1 Then curLev = 1
If curLev > 3 Then curLev = 3

RPos.Width = (Bar.Width / 2) * (curLev - 1)  '' Tick on scrollbar
cTab.BackColor = CLng(lLevs(curLev - 1)) '' Set fillcolor
End Sub


Private Sub Tool_Click(Index As Integer)
Dim I As Integer
''' If the tool is flood i use it!

If Index = 2 Then
    For I = 0 To Pix.Count - 1
        Pix(I).FillColor = lLevs(2 - fLevel(Pix(I).FillColor))
    Next
    Call BackUp
End If
End Sub

Private Sub Tool_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim I As Integer
For I = 0 To Tool.Count - 1
Sh(I).BackStyle = 0
Next
curTool = Index         ''  Set Current Tool
Sh(Index).BackStyle = 1
If Index = 2 Then Exit Sub
Sfondo.MouseIcon = LoadResPicture(101 + Index, vbResCursor)
Sfondo.MousePointer = 99
End Sub


Private Sub Tool_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Index = 2 Then
    Tool_MouseDown 0, 0, 0, 0, 0
End If
End Sub

Public Sub OpenLink(Link As String)
Dim FileName As String, Dummy As String
Dim BrowserExec As String * 255
Dim RetVal As Long
Dim Filenumber As Integer
If LCase(Left(Link, 7)) = "mailto:" Then
ShellExecute FrmMain.hwnd, "OPEN", Link, vbNullString, App.Path, 1
Exit Sub
End If
FileName = App.Path + "\temphtm.HTM"
Filenumber = FreeFile                    ' Get unused file number
Open FileName For Output As #Filenumber  ' Create temp HTML file
    Write #Filenumber, "<HTML> <\HTML>"  ' Output text
Close #Filenumber                        ' Close file
BrowserExec = Space(255)
RetVal = FindExecutable(FileName, Dummy, BrowserExec)
BrowserExec = Trim(BrowserExec)
Kill FileName
' If an application is found, launch it!
If RetVal <= 32 Or IsEmpty(BrowserExec) Then ' Error
ShellExecute FrmMain.hwnd, "OPEN", Link, vbNullString, App.Path, 1
Exit Sub
Else
RetVal = ShellExecute(FrmMain.hwnd, "open", BrowserExec, _
Link, Dummy, SW_SHOWNORMAL)
Exit Sub
If RetVal <= 32 Then        ' Error
    ShellExecute FrmMain.hwnd, "OPEN", Link, vbNullString, App.Path, 1
    Exit Sub
End If
End If
ShellExecute FrmMain.hwnd, "OPEN", Link, vbNullString, App.Path, 1
End Sub
