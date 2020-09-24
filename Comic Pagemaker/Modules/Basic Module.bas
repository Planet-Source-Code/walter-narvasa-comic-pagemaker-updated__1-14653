Attribute VB_Name = "Module1"
Option Explicit

Global PicName, lb, rb, zoomactive As Boolean, BrushType, RepRed, RepGre, repBlu, progress, NumSides, AtAngle, Rx, Ry, PolyX, PolyY, CopyX, CopyY
Global ImageArray(4, 1500, 1500) As Integer
Global X, Y As Integer
Global larrCol() As Long
Global Const CB_HEIGHT = 400
Global Const Pi = 3.14159265359
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const SRCERASE = &H4400328
Public Const WHITENESS = &HFF0062
Public Const BLACKNESS = &H42
Public Const ThisApp = "Stu Paint V2"
Public Const ThisKey = "Recent Files"
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Declare Function CloseClipBoard Lib "user32" Alias "CloseClipboard" () As Long
Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As String) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal nCount As Long) As Long
Declare Function TWAIN_AcquireToFilename Lib "EZTW32.DLL" (ByVal hwndApp%, ByVal bmpFileName$) As Integer
Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwndApp&) As Long
Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwndApp As Long, ByVal wPixTypes As Long) As Long
Declare Function TWAIN_IsAvailable Lib "EZTW32.DLL" () As Long
Declare Function TWAIN_EasyVersion Lib "EZTW32.DLL" () As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_MAXIMIZE = 3

Dim ret As Long
Global TileSize
    'the size of each tile
Type Tile
    Type As String
        'type, (ie: 0 = non walkable)
    ImageNum As Integer
        'the index of the picturebox
    Name As String
        'string that labels the tile (ie: grass)
    Src As Long
        'the handle to bitblt from
    Goto As String
        'uses only if it's a door
End Type
Public Tiles(0 To 51) As Tile
    'total of 3 tiles being represented
Public WorldMap(8, 8) As Tile
    'the world map consists of 8 x 8
Global ShowBalloonDialog As Boolean

'Application info for display in the Splash and About forms.
Public Const pcstrAppPlatform As String = "Win 95/98/NT4"

'API declaration used to ensure Splash screen stays on top.
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_HWNDPARENT = (-8)


Public Function Oldpic()
Dim a
For a = 4 To 1 Step -1
Form1.UndoPicBox(a).Picture = Form1.UndoPicBox(a - 1).Image
Form1.UndoPicBox(a).Refresh
Next a
Form1.UndoPicBox(0).Picture = Form1.MainPic.Image
Form1.UndoPicBox(0).Refresh
End Function
Public Function RGBRed(RGBCol As Long) As Integer
    RGBRed = RGBCol And &HFF
End Function
Public Function RGBGreen(RGBCol As Long) As Integer
    RGBGreen = ((RGBCol And &H100FF00) / &H100)
End Function
Public Function RGBBlue(RGBCol As Long) As Integer
    RGBBlue = (RGBCol And &HFF0000) / &H10000
End Function
Public Function replace_routine()
Dim r2, g2, b2, color1, r, g, b
r2 = RepRed / 100
g2 = RepGre / 100
b2 = repBlu / 100
For X = 0 To Form1.MainPic.ScaleWidth - 1
For Y = 0 To Form1.MainPic.ScaleHeight - 1
color1 = GetPixel(Form1.MainPic.hdc, X, Y)
r = (color1 Mod 256)
b = (Int(color1 / 65536))
g = ((color1 - (b * 65536) - r) / 256)
r = Abs(r * r2)
b = Abs(b * b2)
g = Abs(g * g2)
SetPixelV Form1.MainPic.hdc, X, Y, RGB(r, g, b)
Next Y
progress = (100 / (Form1.MainPic.ScaleWidth - 1)) * X
Call progressbar
Next X
Form1.MainPic.Refresh
End Function
Public Sub progressbar()
If Form1.Proview.Checked = False Then Exit Sub
Form1.Progpic(0).Cls
Form1.Progpic(0).ForeColor = RGB(192, 192, 192)
Form1.Progpic(0).Line (CByte(progress), 0)-(100, 200), , BF
Form1.Progpic(0).Line (45, 0)-(45, 0)
Form1.Progpic(0).ForeColor = RGB(0, 0, 0)
Form1.Progpic(0).Print CByte(progress)
Form1.StatusBar1.Panels.Item(1).Picture = Form1.Progpic(0).Image
End Sub
Public Function FileExist(sFileN As String) As Boolean
    Dim tmpRv As Long
    
    On Error Resume Next
    tmpRv = GetAttr(sFileN)
    If Err Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function
Sub GetRecentFiles()
    Dim i, j As Integer
    Dim varFiles As Variant
    If GetSetting(ThisApp, ThisKey, "RecentFile1") = Empty Then Exit Sub
    varFiles = GetAllSettings(ThisApp, ThisKey)
    For i = 0 To UBound(varFiles, 1)
        Form1.sepfil.Visible = True
        Form1.mnuRecentFile(0).Visible = True
        Form1.mnuRecentFile(i).Caption = varFiles(i, 1)
        Form1.mnuRecentFile(i).Visible = True
    Next i
End Sub
Sub UpdateFileMenu(Filename)
        Dim intRetVal As Integer
        intRetVal = OnRecentFilesList(Filename)
        If Not intRetVal Then
            WriteRecentFiles (Filename)
        End If
        GetRecentFiles
End Sub
Function OnRecentFilesList(Filename) As Integer
  Dim i
  For i = 1 To 4
    If Form1.mnuRecentFile(i).Caption = Filename Then
      OnRecentFilesList = True
      Exit Function
    End If
  Next i
    OnRecentFilesList = False
End Function
Sub WriteRecentFiles(OpenFileName)
    Dim i, j As Integer
    Dim strFile, key As String
    For i = 3 To 1 Step -1
        key = "RecentFile" & i
        strFile = GetSetting(ThisApp, ThisKey, key)
        If strFile <> "" Then
            key = "RecentFile" & (i + 1)
            SaveSetting ThisApp, ThisKey, key, strFile
        End If
    Next i
    SaveSetting ThisApp, ThisKey, "RecentFile1", OpenFileName
End Sub

Public Sub CreateTiles()
    'this sub just basically creates all the necessary tiles
    'and their info. The map is basically saved as the
    'picture's index number, that way, we could also load
    'from it
    Tiles(0).Src = Form1.picTile(0).hdc
    Tiles(1).Src = Form1.picTile(1).hdc
    Tiles(2).Src = Form1.picTile(2).hdc
    Tiles(3).Src = Form1.picTile(3).hdc
    Tiles(4).Src = Form1.picTile(4).hdc
    Tiles(5).Src = Form1.picTile(5).hdc
    Tiles(6).Src = Form1.picTile(6).hdc
    Tiles(7).Src = Form1.picTile(7).hdc
    Tiles(8).Src = Form1.picTile(8).hdc
End Sub

Sub Main()
    On Error GoTo HandleErrors
    Form9.Platform = pcstrAppPlatform
    Form9.Show
    'Ensure the Splash form is refreshed prior to displaying the Main form.
    DoEvents
    '---------------------------------------------------------------------------------------------------------------------
    'Perform other start up tasks here...
    'For demo purposes we add a delay to simulate a typical applications initialisation.
    Call SplashDelay
  '---------------------------------------------------------------------------------------------------------------------
    Form1.Show
    DoEvents
    Unload Form9
ExitHandleErrors:
  Exit Sub
HandleErrors:
  MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, App.Title & " Error"
  Resume ExitHandleErrors
End Sub

Public Sub SplashDelay()
    On Error Resume Next
    Dim sngStartTime As Single
    sngStartTime = Timer
    Do Until (Timer - sngStartTime) > 4
          DoEvents
    Loop
End Sub
