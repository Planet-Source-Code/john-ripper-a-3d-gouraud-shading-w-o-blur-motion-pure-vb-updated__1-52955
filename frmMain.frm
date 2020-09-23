VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "BlurMotion"
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   238
   ScaleMode       =   3  'Píxel
   ScaleWidth      =   318
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicBuff 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   3600
      Left            =   -15
      ScaleHeight     =   240
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   -15
      Width           =   4800
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ret As Boolean


Dim objDx As DirectX7
Dim objDraw7 As DirectDraw7
Dim Mays As Boolean


Private Sub Form_Click()
    EndIt
End Sub


Private Sub Form_Load()
    
    Unload frmInit
    
    Show
    
    MakeSinTable
    MakeCosTable
    
    rHeight = 240
    rWidth = 320
    
    XScreen = rWidth
    YScreen = rHeight
    
    XCenter = XScreen \ 2
    YCenter = YScreen \ 2

    ReadFileMesh App.Path & "\Meshes\" & MeshName, Points, Faces, True
    
    PicBuff.Picture = LoadPicture(App.Path & "\" & PaletteName & ".gif")
    
    SpeedXangle = 2
    SpeedYangle = 2
    SpeedZangle = 2
    
    Init
End Sub

Public Sub Init()

    On Local Error GoTo errOut
    If FullScreen = True Then
        Set objDx = New DirectX7
        Set objDraw7 = objDx.DirectDrawCreate("")
        objDraw7.SetCooperativeLevel Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE
        objDraw7.SetDisplayMode 320, 240, 16, 0, DDSDM_STANDARDVGAMODE
    End If
    GetObjectAPI frmMain.PicBuff.Picture, Len(bmpBuff), bmpBuff
    
    If bmpBuff.bmPlanes <> 1 Or bmpBuff.bmBitsPixel <> 8 Then
        MsgBox " 256-color bitmaps only", vbCritical
        EndIt
    End If

    With saBuff
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = bmpBuff.bmHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = bmpBuff.bmWidthBytes
        .pvData = bmpBuff.bmBits
    End With

    Do
        Render
        DoEvents
    Loop

    Exit Sub

errOut:
    EndIt
End Sub

Sub EndIt()
    If FullScreen = True Then
        objDraw7.RestoreDisplayMode
        objDraw7.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL
    End If
    'frmInit.Show vbModal
    End
End Sub

Private Sub PicBuff_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 16 Then
        Mays = True
    ElseIf KeyCode = 27 Then
        EndIt
    End If

End Sub

Private Sub PicBuff_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 16 Then
        Mays = False
    End If
    
    Select Case KeyCode
        Case 90 'z
            If Mays = True Then
                SpeedZangle = SpeedZangle - 1
            Else
                SpeedZangle = SpeedZangle + 1
            End If
        Case 88 'x
            If Mays = True Then
                SpeedXangle = SpeedXangle - 1
            Else
                SpeedXangle = SpeedXangle + 1
            End If
        
        Case 89 'y
            If Mays = True Then
                SpeedYangle = SpeedYangle - 1
            Else
                SpeedYangle = SpeedYangle + 1
            End If
    
    End Select
'    Debug.Print "X:" & SpeedXangle
'    Debug.Print "Y:" & SpeedYangle
'    Debug.Print "Z:" & SpeedZangle
End Sub
