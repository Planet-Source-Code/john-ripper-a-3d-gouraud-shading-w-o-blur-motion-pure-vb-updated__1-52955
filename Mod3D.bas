Attribute VB_Name = "Mod3D"
Option Explicit

Dim ry1 As Long '"real" Y Point Values
Dim ry2 As Long
Dim ry3 As Long
  
' Edge arrays:
'   0 - Edge
'   1 - Color

'segun como esten las 3 componentes del triangulo unas veces un lado sera
'Edge1 y otras veces lo sera otro lado.
'En este ejemplo el primer lado que se pinta sera siempre el Edge1 y los
'otros 2 seran Edge2. Comprueba que si se varian las coordenadas del
'triangulo unas veces Edge1 sera un lado y otras veces sera otro lado.
'En estos arrays se guarda en la 1º componente la 'X para esa Y' y en la
'segunda que color tiene esa X. (es decir, a partir de que X se ha de pintar
'en la coordenada Y.)

'Horrible English:
'Depending with 3D triangle coord., some times one segment is Edge1 and some
'times Edge1 is the other side.
'In this source, we draw first the side Edge1, and the other two triagle sides
'are Edge2. Take a look, that if you change coordinates of triangle, occurs that
'some times Edge1 its a side or is the other.
'In those arrays, we stored on first component, the X value for "that" Y value
'On the second component of array, we stored the "color" for this X.

Dim Edge1(0 To 240, 0 To 1) As Long
Dim Edge2(0 To 240, 0 To 1) As Long

Const ClipX1 As Long = 1
Const ClipY1 As Long = 1
Const ClipX2 As Long = 318
Const ClipY2 As Long = 238

Public EnableBlurMotion As Boolean

Public PaletteName As String
Public FullScreen  As Boolean
Public MeshName    As String

Public pictBuff()  As Byte
Public saBuff      As SAFEARRAY2D
Public bmpBuff     As BitMap

Const XOrg         As Long = 0
Const YOrg         As Long = 0
Const ZOrg         As Long = 260

Const NumSinVal    As Long = 1024

Public NumPoints   As Long
Public NumFaces    As Long

Public XCenter     As Long
Public YCenter     As Long

Public XScreen     As Long
Public YScreen     As Long

Public Type Point3D
    X   As Long
    Y   As Long
    Z   As Long
    aux As Long
End Type

Public Points()     As Point3D
Public TempPoints() As Point3D

Public Type Face3D
    A  As Long
    B  As Long
    C  As Long
    Z  As Long
    AB As Long
    BC As Long
    CA As Long
End Type
Public Faces() As Face3D

Public CosTable(1025) As Long
Public SinTable(1025) As Long

Public Const PI As Single = 3.141592654

Public Xangle      As Long
Public Yangle      As Long
Public Zangle      As Long

Public SpeedXangle As Long
Public SpeedYangle As Long
Public SpeedZangle As Long

Public rHeight     As Long
Public rWidth      As Long



Public Sub ReadFileMesh(ByVal FileMesh As String, arrPoints() As Point3D, arrFaces() As Face3D, Optional ByVal ReadFaces As Boolean = False)
    
  Dim dataFile    As String
  Dim nF          As Integer
  Dim i           As Integer
  Dim FilePoints  As Long
  Dim FileFaces   As Long
  Dim FlagFaces   As Boolean
  Dim CounterFile As Long
  Dim Pos1        As Long
  Dim Pos2        As Long
  Dim Pos3        As Long
  Dim Pos4        As Long
  Dim Pos5        As Long

    nF = FreeFile
    Open FileMesh For Input As #nF
    
    'read "header"
    For i = 1 To 8
        Line Input #nF, dataFile
    Next i

    Line Input #nF, dataFile
    Pos1 = InStr(1, dataFile, "=")
    FilePoints = Mid(dataFile, Pos1 + 1)
    NumPoints = FilePoints
    ReDim arrPoints(FilePoints)
    ReDim TempPoints(FilePoints)
    
    Line Input #nF, dataFile
    FlagFaces = True
    If (InStr(1, dataFile, "Not Available") <> 0) Then
        FlagFaces = False
      Else
        Pos1 = InStr(1, dataFile, "=")
        FileFaces = Mid(dataFile, Pos1 + 1)
        NumFaces = FileFaces
    End If
    
    Line Input #nF, dataFile        '""
    Line Input #nF, dataFile        '"--------------------------POINTS-------------------------"
    
    CounterFile = 0
    Do Until CounterFile = FilePoints + 1
        Line Input #nF, dataFile 'X!Y@Z format
        Pos1 = InStr(1, dataFile, "!")
        Pos2 = InStr(1, dataFile, "@")
        Pos3 = InStr(1, dataFile, "*")
        
        arrPoints(CounterFile).X = Mid(dataFile, 1, Pos1 - 1)
        arrPoints(CounterFile).Y = Mid(dataFile, Pos1 + 1, Pos2 - Pos1 - 1)
        If (Pos3 = 0) Then
            arrPoints(CounterFile).Z = Mid(dataFile, Pos2 + 1)
          Else
            arrPoints(CounterFile).Z = Mid(dataFile, Pos2 + 1, Pos3 - Pos2 - 1)
            arrPoints(CounterFile).aux = Mid(dataFile, Pos3 + 1)
        End If
        CounterFile = CounterFile + 1
    Loop
    
    If (ReadFaces And FlagFaces) Then
        ReDim arrFaces(FileFaces)
        
        Line Input #nF, dataFile    '--------------------------FACES--------------------------
    
        CounterFile = 0
        Do Until CounterFile = FileFaces + 1
            Line Input #nF, dataFile    'A!B@C format
            Pos1 = InStr(1, dataFile, "!")
            Pos2 = InStr(1, dataFile, "@")
            Pos3 = InStr(1, dataFile, "*")
            Pos4 = InStr(1, dataFile, "%")
            Pos5 = InStr(1, dataFile, "(")
            
            arrFaces(CounterFile).A = Mid(dataFile, 1, Pos1 - 1)
            arrFaces(CounterFile).B = Mid(dataFile, Pos1 + 1, Pos2 - Pos1 - 1)
            arrFaces(CounterFile).C = Mid(dataFile, Pos2 + 1, Pos3 - Pos2 - 1)
            arrFaces(CounterFile).Z = 0
            arrFaces(CounterFile).AB = Mid(dataFile, Pos3 + 1, Pos4 - Pos3 - 1)
            arrFaces(CounterFile).BC = Mid(dataFile, Pos4 + 1, Pos5 - Pos4 - 1)
            arrFaces(CounterFile).CA = Mid(dataFile, Pos5 + 1)
            CounterFile = CounterFile + 1
        Loop
    End If
    
    Close #nF
End Sub



Public Sub MakeCosTable()

  Dim CntVal As Long
  Dim CntAng As Single
  Dim IncDeg As Single
  
    IncDeg = 2 * PI / NumSinVal
    CntAng = IncDeg
    CntVal = 0
    
    Do Until CntVal > 1024
        CosTable(CntVal) = CInt((255 * Cos(CntAng)))
        CntAng = CntAng + IncDeg
        CntVal = CntVal + 1
    Loop
End Sub

Public Sub MakeSinTable()

  Dim CntVal As Long
  Dim CntAng As Single
  Dim IncDeg As Single

    IncDeg = 2 * PI / NumSinVal
    CntAng = IncDeg
    CntVal = 0
    
    Do Until CntVal > 1024
        SinTable(CntVal) = CInt((255 * Sin(CntAng)))
        CntAng = CntAng + IncDeg
        CntVal = CntVal + 1
    Loop
End Sub

Public Sub Calc3DRotations(ByVal SinX As Long, ByVal CosX As Long, ByVal SinY As Long, ByVal CosY As Long, ByVal SinZ As Long, ByVal CosZ As Long, _
                           OrgPoints() As Point3D, DesPoints() As Point3D, _
                           ByVal NumPoints As Long)

  Dim X1  As Long
  Dim Y1  As Long
  Dim z1  As Long
  Dim cnt As Long

    For cnt = 0 To NumPoints
        
'     X1 := (cos(YAngle) * X  - sin(YAngle) * Z)
        X1 = (CosY * OrgPoints(cnt).X - SinY * OrgPoints(cnt).Z) \ 256
'     Z1 := (sin(YAngle) * X  + cos(YAngle) * Z)
        z1 = (SinY * OrgPoints(cnt).X + CosY * OrgPoints(cnt).Z) \ 256
'     X  := (cos(ZAngle) * X1 + sin(ZAngle) * Y)
        DesPoints(cnt).X = (CosZ * X1 + SinZ * OrgPoints(cnt).Y) \ 256
'     Y1 := (cos(ZAngle) * Y  - sin(ZAngle) * X1)
        Y1 = (CosZ * OrgPoints(cnt).Y - SinZ * X1) \ 256
'     Z  := (cos(XAngle) * Z1 - sin(XAngle) * Y1)
        DesPoints(cnt).Z = (CosX * z1 - SinX * Y1) \ 256
'     Y  := (sin(XAngle)) * Z1 + cos(XAngle) * Y1)
        DesPoints(cnt).Y = (SinX * z1 + CosX * Y1) \ 256
      
        DesPoints(cnt).aux = OrgPoints(cnt).aux
   Next cnt
End Sub

Public Sub Proyect3D(ByVal XScreen As Long, ByVal YScreen As Long, ByVal NumPoints As Long, _
                     OrgPoints() As Point3D, DesPoints() As Point3D)

  Dim cnt As Long

    For cnt = 0 To NumPoints
        
        With OrgPoints(cnt)
            DesPoints(cnt).X = XScreen + ((XOrg * .Z - .X * ZOrg) / (.Z - ZOrg))
            DesPoints(cnt).Y = YScreen + ((YOrg * .Z - .Y * ZOrg) / (.Z - ZOrg))
        End With
    Next cnt
End Sub


Public Sub QuickSortZFaces(ByVal NumPoints As Long, Points2qS() As Point3D, Faces2qS() As Face3D)
  
  Dim cnt As Long
    
    For cnt = 0 To NumPoints
     
        With Faces2qS(cnt)
            .Z = (Points2qS(.A).Z + Points2qS(.B).Z + Points2qS(.C).Z) \ 3
        End With
    Next cnt
    
    Call QuickSortFaces(Faces2qS, 0, NumPoints)
End Sub

Private Sub QuickSortFaces(vntArr() As Face3D, ByVal lngLeft As Long, ByVal lngRight As Long)

  Dim i          As Long
  Dim j          As Long
  Dim lngMid     As Long
  Dim vntTestVal As Variant
  Dim vntTemp    As Face3D
    
    If (lngLeft < lngRight) Then
        
        lngMid = (lngLeft + lngRight) \ 2
        vntTestVal = vntArr(lngMid).Z
        i = lngLeft
        j = lngRight
        
        Do
            Do While vntArr(i).Z < vntTestVal
                i = i + 1
            Loop
            Do While vntArr(j).Z > vntTestVal
                j = j - 1
            Loop
            If (i <= j) Then
                vntTemp = vntArr(j)
                vntArr(j) = vntArr(i)
                vntArr(i) = vntTemp
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j

        If (j <= lngMid) Then
            Call QuickSortFaces(vntArr, lngLeft, j)
            Call QuickSortFaces(vntArr, i, lngRight)
          Else
            Call QuickSortFaces(vntArr, i, lngRight)
            Call QuickSortFaces(vntArr, lngLeft, j)
        End If
    End If
End Sub

Public Function FaceVisible(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, X3 As Long, Y3 As Long) As Boolean
'Simple escalar product:
'Return TRUE if face is visible
'if FaceVisible=False NOT PAINT anything. Increase speed!

  Dim A As Long
  Dim B As Long

    A = (X2 - X1) * (Y3 - Y1)
    B = (X3 - X1) * (Y2 - Y1)
    
    FaceVisible = (A - B >= 0)
End Function

Public Sub Render()

  Dim i   As Long

  Dim px  As Long
  Dim py  As Long

  Dim colp1 As Long 'Color Points
  Dim colp2 As Long
  Dim colp3 As Long


    CopyMemory ByVal VarPtrArray(pictBuff), VarPtr(saBuff), 4
    
    If EnableBlurMotion = False Then
        For py = ClipY1 To ClipY2
            For px = ClipX1 To ClipX2
                pictBuff(px, py) = 1
            Next px
        Next py
    End If
    
    Call Calc3DRotations(SinTable(Xangle), CosTable(Xangle), SinTable(Yangle), CosTable(Yangle), SinTable(Zangle), CosTable(Zangle), Points, TempPoints, UBound(Points))
    Call Proyect3D(XCenter, YCenter, UBound(Points), TempPoints, TempPoints)
    Call QuickSortZFaces(UBound(Faces), TempPoints, Faces)

    For i = 0 To UBound(Faces)
    
        If (FaceVisible(TempPoints(Faces(i).A).X, TempPoints(Faces(i).A).Y, TempPoints(Faces(i).B).X, TempPoints(Faces(i).B).Y, TempPoints(Faces(i).C).X, TempPoints(Faces(i).C).Y)) Then
        
            'Select point color by Z coord:
            
            colp1 = (TempPoints(Faces(i).A).Z \ 2) + 32
            'To avoid Division by Zero
            If colp1 <= 0 Then
                colp1 = 1
            End If
            
            colp2 = (TempPoints(Faces(i).B).Z \ 2) + 32
            If colp2 <= 0 Then
                colp2 = 1
            End If
            
            colp3 = (TempPoints(Faces(i).C).Z \ 2) + 32
            If colp3 <= 0 Then
                colp3 = 1
            End If
            
            GouraudFill pictBuff, TempPoints(Faces(i).A).X, TempPoints(Faces(i).A).Y, _
                    TempPoints(Faces(i).B).X, TempPoints(Faces(i).B).Y, _
                    TempPoints(Faces(i).C).X, TempPoints(Faces(i).C).Y, _
                    colp1, colp2, colp3
            
        End If
    Next i

    Xangle = Xangle + SpeedXangle
    If (Xangle > 1024) Then Xangle = Xangle - 1024 Else If (Xangle < 0) Then Xangle = 0
    
    Yangle = Yangle + SpeedYangle
    If (Yangle > 1024) Then Yangle = Yangle - 1024 Else If (Yangle < 0) Then Yangle = 0
    
    Zangle = Zangle + SpeedZangle
    If (Zangle > 1024) Then Zangle = Zangle - 1024 Else If (Zangle < 0) Then Zangle = 0
    
    If EnableBlurMotion = True Then
        For py = ClipY1 To ClipY2
            For px = ClipX1 To ClipX2
                pictBuff(px, py) = (CLng(pictBuff(px + 1, py)) + pictBuff(px - 1, py) + pictBuff(px, py + 1) + pictBuff(px, py - 1)) \ 4
            Next px
        Next py
    End If
    
    CopyMemory ByVal VarPtrArray(pictBuff), 0&, 4
    frmMain.PicBuff.Refresh

End Sub

'This routine stored data on Edge array for Gradiet Triangle
'(modification of Breseham algorithm)
Private Sub GouraudFill(BitMap() As Byte, ByVal X1 As Long, ByVal Y1 As Long, _
                        ByVal X2 As Long, ByVal Y2 As Long, _
                        ByVal X3 As Long, ByVal Y3 As Long, _
                        ByVal Col1 As Long, _
                        ByVal Col2 As Long, _
                        ByVal Col3 As Long)
                        
  Dim SwapVar As Long
  Dim FixX As Long
  Dim FixCol As Long
  Dim AuxFix As Long
  Dim IncLine As Long
  Dim IncCol As Long
  Dim CntY As Long

    If Y1 > Y2 Then  ' change (X1,Y1) (X2,Y2)
        SwapVar = X1
        X1 = X2
        X2 = SwapVar
        SwapVar = Y1
        Y1 = Y2
        Y2 = SwapVar
        SwapVar = Col1
        Col1 = Col2
        Col2 = SwapVar
    End If
  
    If Y1 > Y3 Then  ' change (X1,Y1) (X3,Y3)
        SwapVar = X1
        X1 = X3
        X3 = SwapVar
        SwapVar = Y1
        Y1 = Y3
        Y3 = SwapVar
        SwapVar = Col1
        Col1 = Col3
        Col3 = SwapVar
    End If

    If (Y2 - Y1) < (Y3 - Y1) Then ' change (X2,Y2) (X3,Y3)
        SwapVar = X2
        X2 = X3
        X3 = SwapVar
        SwapVar = Y2
        Y2 = Y3
        Y3 = SwapVar
        SwapVar = Col2
        Col2 = Col3
        Col3 = SwapVar
    End If
    
' First LINE -------------------------------------------------------------------
    
    IncLine = (X2 - X1)
    IncLine = 64 * IncLine
    'asm mov ax,Incline;sal ax,6;mov Incline,ax end;
    IncCol = (Col2 - Col1)
    IncCol = 256 * IncCol
    'asm mov ax,IncCol;sal ax,8;mov IncCol,ax end;

    If (Y2 - Y1) <> 0 Then
        IncLine = IncLine \ (Y2 - Y1)
        IncCol = IncCol \ (Y2 - Y1)
    Else
        IncLine = 0
        IncCol = 0
    End If
    
    FixX = X1
    FixX = 64 * FixX
    'asm mov ax,FixX;sal ax,6;mov FixX,ax end;
    'FixCol:=Col1 shl 8;
    FixCol = Col1 * 256
    ry1 = Y1
  
    For CntY = Y1 To Y2
        If (CntY >= ClipY1) And (CntY <= ClipY2) Then
            AuxFix = FixX
            'asm mov ax,AuxFix;sar ax,6;mov AuxFix,ax end;
            AuxFix = AuxFix \ 64
            Edge1(CntY, 0) = AuxFix
            Edge1(CntY, 1) = FixCol \ 256
        End If
        FixX = FixX + IncLine
        FixCol = FixCol + IncCol
    Next CntY
    
' Second LINE -------------------------------------------------------------------

    IncLine = (X3 - X1)
    IncLine = 64 * IncLine
    'asm mov ax,Incline;sal ax,6;mov Incline,ax end;
    IncCol = (Col3 - Col1)
    IncCol = 256 * IncCol
    'asm mov ax,IncCol;sal ax,8;mov IncCol,ax end;
  
    If (Y3 - Y1) <> 0 Then
        IncLine = IncLine \ (Y3 - Y1)
        IncCol = IncCol \ (Y3 - Y1)
    Else
        IncLine = 0
        IncCol = 0
    End If
    
    FixX = X1
    FixX = 64 * FixX
    'asm mov ax,FixX;sal ax,6;mov FixX,ax end;
    'FixCol:=Col1 shl 8;
    FixCol = Col1 * 256
    ry2 = Y3
  
    For CntY = Y1 To Y3
        If (CntY >= ClipY1) And (CntY <= ClipY2) Then
            AuxFix = FixX
            'asm mov ax,AuxFix;sar ax,6;mov AuxFix,ax end;
            AuxFix = AuxFix \ 64
            Edge2(CntY, 0) = AuxFix
            Edge2(CntY, 1) = FixCol \ 256
        End If
        FixX = FixX + IncLine
        FixCol = FixCol + IncCol
    Next CntY
    

' Thrird and last Triangle LINE -------------------------------------------------------------------

    IncLine = (X2 - X3)
    IncLine = 64 * IncLine
    'asm mov ax,Incline;sal ax,6;mov Incline,ax end;
    IncCol = (Col2 - Col3)
    IncCol = 256 * IncCol
    'asm mov ax,IncCol;sal ax,8;mov IncCol,ax end;
  
    If (Y2 - Y3) <> 0 Then
        IncLine = IncLine \ (Y2 - Y3)
        IncCol = IncCol \ (Y2 - Y3)
    Else
        IncLine = 0
        IncCol = 0
    End If
    FixX = X3
    FixX = 64 * FixX
    'asm mov ax,FixX;sal ax,6;mov FixX,ax end;
    'FixCol:=Col1 shl 8;
    FixCol = Col3 * 256
    
    For CntY = Y3 To Y2
        If (CntY >= ClipY1) And (CntY <= ClipY2) Then
            AuxFix = FixX
            'asm mov ax,AuxFix;sar ax,6;mov AuxFix,ax end;
            AuxFix = AuxFix \ 64
            Edge2(CntY, 0) = AuxFix
            Edge2(CntY, 1) = FixCol \ 256
        End If
        FixX = FixX + IncLine
        FixCol = FixCol + IncCol
    Next CntY
    

'Cliiiiiiiiiiping Y
  If Y1 < ClipY1 Then Y1 = ClipY1
  If Y2 > ClipY2 Then Y2 = ClipY2
    
    ry1 = Y1
    ry3 = Y2
    
    'Now, fill the BitMap pointer by Grouraud Gradient
    MyFill BitMap

End Sub

Private Sub MyFill(BitMap() As Byte)
'Este procedimiento va rellenando el triangulo mediante lineas Horizontales
'de tal modo que la linea horizontal se va degradando desde su color
'inicial, que se guarda en Edge?[?,1] hasta su color final Edge?[?,1]
'y su coordenada X se guarda en Edge?[?,0].La coordenada Y vendra dada
'por la variable CNT33, que va desde la coordanada Y1 hasta la Y3

'This sub, fills the triangle by Horizontal "lines", degradating the color
'Initial Color: Edge?(?,1)
'Finish Color: Edge?(?,1)

Dim cnt33   As Long
Dim cini    As Long
Dim cfin    As Long
Dim IncCol  As Long
Dim FixCol  As Long
Dim aux2    As Long
    
    For cnt33 = ry1 To ry3
        If Edge1(cnt33, 0) <= Edge2(cnt33, 0) Then
            cini = Edge1(cnt33, 1)
            cfin = Edge2(cnt33, 1)
            IncCol = cfin - cini
            IncCol = 256 * IncCol
            If (cfin - cini) <> 0 And (Edge2(cnt33, 0) - Edge1(cnt33, 0) <> 0) Then
                IncCol = IncCol \ (Edge2(cnt33, 0) - Edge1(cnt33, 0))
            Else
                IncCol = 0
            End If
            FixCol = 256 * cini
            For aux2 = Edge1(cnt33, 0) To Edge2(cnt33, 0)
                If (aux2 < ClipX1 Or aux2 > ClipX2 Or cnt33 < ClipY1 Or cnt33 > ClipY2) Then
                Else
                    BitMap(aux2, cnt33) = FixCol \ 256
                End If
                FixCol = FixCol + IncCol
            Next aux2
         
         Else
            cini = Edge2(cnt33, 1)
            cfin = Edge1(cnt33, 1)
            IncCol = cfin - cini
            IncCol = 256 * IncCol
            If (cfin - cini) <> 0 And (Edge1(cnt33, 0) - Edge2(cnt33, 0)) <> 0 Then
                IncCol = IncCol \ (Edge1(cnt33, 0) - Edge2(cnt33, 0))
            Else
                IncCol = 0
            End If
            FixCol = 256 * cini
            For aux2 = Edge2(cnt33, 0) To Edge1(cnt33, 0)
                If (aux2 < ClipX1 Or aux2 > ClipX2 Or cnt33 < ClipY1 Or cnt33 > ClipY2) Then
                Else
                    BitMap(aux2, cnt33) = FixCol \ 256
                End If
                FixCol = FixCol + IncCol
            Next aux2
         End If
    Next cnt33
End Sub
