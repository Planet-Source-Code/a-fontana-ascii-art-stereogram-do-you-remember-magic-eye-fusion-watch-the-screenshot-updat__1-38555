Attribute VB_Name = "ModStereo"
Option Explicit

'''''''''''''''''''''
''    CONSTS >>>   ''
'''''''''''''''''''''

''  DEFAULT PICTURE >>>
Public Const cStartUp As String = "069A5i1\14A5i1\14A5i1\14A5i1\14A5i1\10A5𓚽i1\10A5𓚽i1\10A5𓚽i1\10A5𓚽i1\10A5𓚽i1\10A5𓚽i1\10A5𓚽i1\10A5𓚽i1\10A5𓚽i1\10A5𓚽i1\10A5𓚽i1\10A5𓚽i1\10A5𓚽i1\10A5𓚽i1\10A5𓚽i1\10A5𓚽i1\10A5𓚽i1\14A5i1\14A5i1\14A5i1\45A"

''  DICTONARY >>>
Public Const cAlpha As String = "1,2,3,4,5,6,7,8,9,0,q,w,e,r,t,y,u,i,o,p,a,s,d,f,g,h,j,k,l,z,x,c,v,b,n,m,Q,W,E,R,T,Y,U,I,O,P,A,S,D,F,G,H,J,K,L,Z,X,C,V,B,N,M"

''  Color Level >>>
Public Const cLevs As String = "&H00FFFFFF,&H00999999,&H00000000"
Public Const nWidth As Integer = 80    '' Picture Width
Public Const nHeight As Integer = 30   '' Picture Height
Public Const minString As Integer = 16 '' Tile Lenght
Const RC_PALETTE As Long = &H100       '' \
Const SIZEPALETTE As Long = 104        ''  --> BMP CONSTS
Const RASTERCAPS As Long = 38          '' /
Public Const WM_GETFONT = &H31         '' Sendmessage Consts


'''''''''''''''''''''
''     Types >>    ''
'''''''''''''''''''''

'' A simple rect :) >>>
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'' These are all BMP STRUCTURE >>>
Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY ' Enough for 256 colors
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Public Enum sMode
    Normal = 0
    Randomized = 1
    WithPicture = 2
End Enum

'''''''''''''''''''
''    ARRAY >>   ''
'''''''''''''''''''

Public lLevs() As String       ''  Levels Color's Array.
Public lDizio() As String      ''  All symbol we can use.
Public sBackup() As String     ''  Backup. (for undo's use)

''  OTHER VARS >>

Public curBackPos As Integer  ''  Current backup state.
Public curLev As Integer      ''  Current color level.
Public curTool As Integer     ''  Current tool used.
Public StereoMode As sMode    ''  Are we using picture as back?
Public sPictureFile As String ''  The filename of picture
Public lBackColor As Long     ''  The BackColor
Public lForeColor As Long     ''  The Forecolor
Public pMINI As StdPicture

'' CLASSES >>

Public CD As FileCommonDialog ''  Save & Open Dialogs

'' API DECLARATION >>>

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal HDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal HDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal HDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal HDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal HDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal HDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal HDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Public Sub ImportImage(PicFile As String)
Dim X As Integer, Y As Integer        ''  "For-Next" Vars
Dim pPic As StdPicture, HDC As Long   ''  To import picture
Dim PixColor As Long                  ''  Pixel Color

Set pPic = LoadPicture(PicFile)       ''  Load The Picture
HDC = CreateCompatibleDC(0)                  ''\
DeleteObject SelectObject(HDC, pPic.Handle)  '' --> Whe Read here.

With FrmMain

For Y = 0 To nHeight - 1
    For X = 0 To nWidth - 1
        PixColor = GetPixel(HDC, X, Y)
        Select Case GetPixel(HDC, X, Y)
            Case 0, &HFFFFFF, &H999999
                .Pix(Y * nWidth + X).FillColor = GetPixel(HDC, X, Y)
            Case Else
                '' We ignore the color if is not one of three standard color
        End Select
    Next
Next

DeleteObject HDC       '' Free MEM!
End With
End Sub

Public Sub ExportImage(PicFile As String)
Dim X As Integer, Y As Integer        ''  Per Scanning
Dim pPic As StdPicture, HDC As Long, hbit As Long   ''  Per importare immagini
Dim PixColor As Long                  ''  Colore del pixel
Dim nWidth As Long
Dim nHeight As Long
With FrmMain

nWidth = .Sfondo.Width / 6 / Screen.TwipsPerPixelX     '' BMP's
nHeight = .Sfondo.Height / 6 / Screen.TwipsPerPixelY   '' SIZES


HDC = CreateCompatibleDC(GetDC(FrmMain.hwnd))            ''\
hbit = CreateCompatibleBitmap(GetDC(0), nWidth, nHeight) '' --> We draw here
DeleteObject SelectObject(HDC, hbit)                     ''/


For Y = 0 To nHeight - 1
    For X = 0 To nWidth - 1
        PixColor = .Pix(Y * nWidth + X).FillColor
        SetPixel HDC, X, Y, PixColor
    Next
Next


Set pPic = hDCToPicture(HDC, 0, 0, nWidth, nHeight) '' --> Copy Picture

End With
SavePicture pPic, PicFile  '' --> Save Picture
DeleteDC HDC               '' \
DeleteObject hbit          ''  --> Free MEM!
End Sub


Public Function GeneraStereo() As String
Dim tStr As String, fStr As String, tCol As Long
Dim T As Integer, I As Integer, K As Integer
Dim Diff As Integer

With FrmMain

For K = 0 To nHeight - 1                  ''  It really Generate the stereogram
tCol = &HFFFFFF                           ''  It's a little (... little...) difficult
tStr = RandomString(minString)            ''  to explain how it works on a few lines
Diff = minString                          ''  >>>>>>>>>>>
For I = 0 To nWidth - 1                   ''  >>>>>>>>>>>
    T = K * nWidth + I
     If .Pix(T).FillColor <> tCol Then
        Diff = minString - fLevel(.Pix(T).FillColor)
        tCol = .Pix(T).FillColor
    End If
    tStr = tStr + Left(Right(tStr, Diff), 1)
Next
fStr = fStr + tStr + vbCrLf
Next
GeneraStereo = fStr
End With

End Function

Public Sub Flood(T As Long, Col As Long, sCol As Long)

'' Recursive FLOOD!

With FrmMain

If (T > .Pix.UBound) Or (T < .Pix.LBound) Then Exit Sub

If .Pix(T).FillColor <> sCol Then Exit Sub
If .Pix(T).FillColor = Col Then Exit Sub

.Pix(T).FillColor = Col

If (T Mod 80) > 0 Then
Flood T - 1, Col, sCol
End If

If (T Mod 80) < 79 Then
Flood T + 1, Col, sCol
End If

Flood T + 80, Col, sCol
Flood T - 80, Col, sCol

End With
End Sub

Public Function RandomString(strLen As Integer) As String
Dim I As Integer
Dim nChr As String

For I = 1 To strLen                           ''  GENERATE A RADOM
nChr = lDizio(Int(Rnd(1) * UBound(lDizio)))   ''  STRING !!!!
RandomString = RandomString + nChr
Next

End Function

Public Function Draw(X As Single, Y As Single, Button As Integer, TL As Integer, Optional bClick = False)
Dim W As Long, H As Long
Dim T As Long, Col As Long

With FrmMain

If Button = 0 Then Exit Function



W = X \ 6
H = Y \ 8
T = H * nWidth + W

If Button = 1 Then Col = .cTab.BackColor
If Button = 2 Then Col = &HFFFFFF '' Cancello

If Not (X > .Sfondo.ScaleWidth - 2 Or Y > .Sfondo.ScaleHeight - 2 Or Y < 0 Or X < 0) Then
    If Not (T > .Pix.UBound Or T < .Pix.LBound) Then
        Select Case TL
        Case 0
            If .Pix(T).FillColor <> Col Then
            .Pix(T).FillColor = Col
            End If
        Case 1
            If bClick = True Then
                Flood T, Col, .Pix(T).FillColor
                Call BackUp
            End If
        End Select
    End If
End If

End With
End Function

Public Function fLevel(tCol As Long) As Integer
Dim I As Integer

For I = 0 To UBound(lLevs())
If tCol = lLevs(I) Then
    fLevel = I
    Exit Function
End If
Next

End Function

Public Function Compatta() As String
Dim I As Integer
Dim tStr As String

''  IT COMPRESS INFORMATION

With FrmMain

For I = 0 To .Pix.Count - 1
tStr = tStr & fLevel(.Pix(I).FillColor)
If Len(tStr) = 4 Then
    Compatta = Compatta + Encode(tStr)
    If tStr <> "0000" Then
    DoEvents
    End If
    tStr = ""
End If
Next
If Len(tStr) > 0 Then
    Compatta = Compatta + Encode(tStr & String(4 - Len(tStr), "0"))
    tStr = ""
End If
End With
End Function
Public Function Encode(tStr As String) As String
Dim tVal As Integer
Dim Pot As Integer
Dim I As Integer
Pot = 1
For I = 1 To 4
tVal = tVal + Val(Mid(tStr, 5 - I, 1)) * Pot
Pot = Pot * 3
Next
Encode = Chr(65 + tVal)
End Function
Public Sub Decode(Text As String)
Dim I As Integer
Dim tVal As Integer
Dim fCode As String
Dim tStr As String

For I = 1 To Len(Text)
tVal = Asc(Mid(Text, I, 1)) - 65
tStr = ""

Do While tVal > 0
tStr = (tVal Mod 3) & tStr
tVal = tVal \ 3
Loop

If Len(tStr) < 4 Then
tStr = (tVal Mod 3) & tStr
tStr = String(4 - Len(tStr), "0") & tStr
End If

fCode = fCode & tStr
Next

With FrmMain

For I = 0 To .Pix.Count - 1
.Pix(I).FillColor = lLevs(Val(Mid(fCode, I + 1, 1)))
Next

End With
End Sub

'' A SIMPLE RLE CODING / DECONDING >>>>>>

Public Function EncodeRLE(txt As String) As String
Dim I As Integer
Dim iCount As Integer
Dim tChr As String
txt = txt + Chr(0)
tChr = ""
For I = 1 To Len(txt)
If Mid(txt, I, 1) <> tChr Then
EncodeRLE = EncodeRLE & iCount & tChr
tChr = Mid(txt, I, 1)
iCount = 1
Else
iCount = iCount + 1
End If
Next

End Function
Public Function DecodeRLE(txt As String) As String
Dim tNum As String
Dim tChr As String * 1
Dim I As Integer
I = 1
Do While (I <= Len(txt))
tNum = ""
tChr = Mid(txt, I, 1)
Do While Asc(tChr) > 47 And Asc(tChr) < 58
tNum = tNum & tChr
I = I + 1
tChr = Mid(txt, I, 1)
Loop
DecodeRLE = DecodeRLE & String(Val(tNum), tChr)
I = I + 1
Loop
End Function

Public Sub BackUp()     ''''' CREATE A BACKUP
Dim sBack As String
sBack = Compatta()
If sBack <> sBackup(curBackPos) Then
    curBackPos = curBackPos + 1
    ReDim Preserve sBackup(curBackPos)
    sBackup(curBackPos) = sBack
End If
FrmMain.mnuCancel.Enabled = True
FrmMain.mnuRip.Enabled = False

End Sub

Public Function FileExist(FileName As String) As Boolean
On Error GoTo fine:
Dim fLen As Long
FileExist = False
fLen = FileLen(FileName)
FileExist = True
fine:
End Function


Public Function GeneraBmp(sfile As String, sBuffer As String)
Dim pPic As StdPicture
Dim MiniDC As Long   '' Handle For Device Context
Dim HDC As Long, hbit As Long, hFont As Long   ''  Image's Settings
Dim hPen As Long, hBrush As Long               ''  Image's Settings
Dim nnWidth As Long, nnHeight As Long          ''  Image's Width & Height
Dim I As Integer, K As Integer                 ''  For Vars
Dim R As RECT                                  ''  Rect To Write
Dim tRect As RECT                              ''  Rect To Fill Background

FrmMain.Font.Name = "Terminal"
FrmMain.Font.Size = "6"

nnWidth = (nWidth + minString) * 6 + 5
nnHeight = nHeight * 8 + 4

'' If you don't select a background color it uses white
If lBackColor = -1 Then lBackColor = vbWhite
hBrush = CreateSolidBrush(lBackColor)

HDC = CreateCompatibleDC(GetDC(0))
hbit = CreateCompatibleBitmap(GetDC(0), nnWidth, nnHeight)
hFont = SendMessage(FrmMain.hwnd, WM_GETFONT, 0&, ByVal 0&)
    
DeleteObject SelectObject(HDC, hbit)
DeleteObject SelectObject(HDC, hFont)

'' If you don't select a foreground color it uses black
If lForeColor = -1 Then lForeColor = 0
SetTextColor HDC, lForeColor
SetBkMode HDC, 1

SetRect tRect, 0, 0, nnWidth, nnHeight
FillRect HDC, tRect, hBrush

DeleteObject hBrush

Select Case StereoMode
    Case Normal
        SetRect R, 2, 2, nnWidth - 2, nnHeight - 2
        DrawTextEx HDC, sBuffer, Len(sBuffer), R, 0, ByVal 0&
    Case Randomized
        Dim MiniBit As Long  '' Handel For Bitmap Picture
        Dim rColor As Long   '' Random Color
        
        MiniDC = CreateCompatibleDC(GetDC(0))
        MiniBit = CreateCompatibleBitmap(GetDC(0), minString, 30)
        DeleteObject SelectObject(MiniDC, MiniBit)
        
        For I = 0 To 15
            
            For K = 0 To 30
                rColor = RGB(Int(Rnd(1) * 255), Int(Rnd(1) * 255), Int(Rnd(1) * 255))
                
                Do While rColor = lBackColor
                    rColor = RGB(Int(Rnd(1) * 255), Int(Rnd(1) * 255), Int(Rnd(1) * 255))
                    DoEvents
                Loop
                
                SetPixel MiniDC, I, K, rColor
            Next
            DoEvents
            
        Next
        
        CreateFromPic MiniDC, HDC, sBuffer
        DeleteDC MiniDC
        
    Case WithPicture
    
        sBuffer = Replace(sBuffer, vbCrLf, "")
        MiniDC = CreateCompatibleDC(GetDC(0))
        DeleteObject SelectObject(MiniDC, pMINI.Handle)
        
        CreateFromPic MiniDC, HDC, sBuffer
        DeleteDC MiniDC
End Select
FrmMain.Font.Name = "Arial"
FrmMain.Font.Size = "8"



Set pPic = hDCToPicture(HDC, 0, 0, nnWidth, nnHeight)
SavePicture pPic, sfile

DeleteObject hbit
DeleteObject hFont
DeleteDC HDC
End Function

Public Sub CreateFromPic(MiniDC As Long, HDC As Long, sBuffer As String)
Dim T As Integer, T1 As Integer, I As Integer, K As Integer
Dim tCol As Long, nCol As Long
Dim Diff As Integer
Dim SupportDc As Long
Dim SupportBitmap As Long

SupportDc = CreateCompatibleDC(GetDC(0))
SupportBitmap = CreateCompatibleBitmap(GetDC(0), nWidth + minString, nHeight)
DeleteObject SelectObject(SupportDc, SupportBitmap)

sBuffer = Replace(sBuffer, vbCrLf, "")

With FrmMain

T = 0
T1 = 0
For K = 0 To nHeight - 1                '' Now it color the stereogram...
tCol = &HFFFFFF
Diff = minString
For I = 0 To minString + nWidth - 1
    T = T + 1
    If I < minString Then
        nCol = GetPixel(MiniDC, I, K)
        SetTextColor HDC, nCol
        SetPixel SupportDc, I, K, nCol
    Else
        If .Pix(T1).FillColor <> tCol Then
        Diff = minString - fLevel(.Pix(T1).FillColor)
        tCol = .Pix(T1).FillColor
        End If
        T1 = T1 + 1
        nCol = GetPixel(SupportDc, I - Diff, K)
        SetTextColor HDC, nCol
        SetPixel SupportDc, I, K, nCol
    End If
    TextOut HDC, 2 + I * 6, 2 + K * 8, Mid(sBuffer, T + 1, 1), 1
Next
Next

DeleteDC SupportDc
DeleteObject SupportBitmap
End With



End Sub

''' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''
''' This Last Two Subs is taken from
''' Allapi.net. Thanks to them!!!!
'''
''' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim R As Long, PIC As PicBmp, IPic As IPicture, IID_IDispatch As GUID

    'Fill GUID info
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    'Fill picture info
    With PIC
        .Size = Len(PIC) ' Length of structure
        .Type = vbPicTypeBitmap ' Type of Picture (bitmap)
        .hBmp = hBmp ' Handle to bitmap
        .hPal = hPal ' Handle to palette (may be null)
    End With

    'Create the picture
    R = OleCreatePictureIndirect(PIC, IID_IDispatch, 1, IPic)

    'Return the new picture
    Set CreateBitmapPicture = IPic
End Function
Function hDCToPicture(ByVal hDCSrc As Long, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    Dim hDCMemory As Long, hBmp As Long, hBmpPrev As Long, R As Long
    Dim hPal As Long, hPalPrev As Long, RasterCapsScrn As Long, HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long, LogPal As LOGPALETTE

    'Create a compatible device context
    hDCMemory = CreateCompatibleDC(hDCSrc)
    'Create a compatible bitmap
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    'Select the compatible bitmap into our compatible device context
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    'Raster capabilities?
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
    'Does our picture use a palette?
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette
    'What's the size of that palette?
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Set the palette version
        LogPal.palVersion = &H300
        'Number of palette entries
        LogPal.palNumEntries = 256
        'Retrieve the system palette entries
        R = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        'Create the palette
        hPal = CreatePalette(LogPal)
        'Select the palette
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        'Realize the palette
        R = RealizePalette(hDCMemory)
    End If

    'Copy the source image to our compatible device context
    R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)

    'Restore the old bitmap
    hBmp = SelectObject(hDCMemory, hBmpPrev)

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Select the palette
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If

    'Delete our memory DC
    R = DeleteDC(hDCMemory)

    Set hDCToPicture = CreateBitmapPicture(hBmp, hPal)
End Function

