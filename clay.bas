Attribute VB_Name = "clay"
Option Explicit

Public diameter As Double 'unit in (m)
Public modulus As Double 'unit in (kN/m2)
Public inertia As Double 'unit in (m4)
Public pt As Double 'unit in (kN)
Const pi = 3.14159265358979
Public i, j, k, l, m, n As Integer
Const consJ = 0.5

Function area(diameter)

    area = 0.25 * pi * (diameter) ^ 2

End Function

Sub py_clay()

    Dim wsInput As Worksheet
    Dim wsReport As Worksheet
    Dim wsCalc As Worksheet

    Set wsInput = Worksheets("Input")
    Set wsReport = Worksheets("Report")
    Set wsCalc = Worksheets("Calc")

    Dim numData As Integer
    
    numData = WorksheetFunction.Count(wsInput.Range(wsInput.Range("a6"), wsInput.Range("A6").End(xlDown)))
    
    diameter = wsInput.Cells(1, 2)
    modulus = wsInput.Cells(2, 2)
    inertia = wsInput.Cells(3, 2)
    pt = wsInput.Cells(1, 6)
    
    
    ReDim zi(numData)
    ReDim zo(numData)
    
    
    Dim km0 As Double
    Dim deltaX As Double
    Dim segmen As Integer
    Dim point As Integer
    Dim deltaPt As Integer
    Dim nodal As Integer
    Dim mi As Integer
    Dim MM As Integer
    
    segmen = wsInput.Cells(2, 6)
    point = 15
    nodal = (2 * segmen) + 1
    deltaPt = 5
    
    ReDim Pult0(point)
    ReDim valP0(point)
    ReDim valY0(point)
    
    ReDim Pult1(segmen)
    ReDim Pult2(segmen)
    ReDim Pult(segmen)
    ReDim epsilon50(segmen)
    ReDim valY50(segmen)
    
    ReDim valY(segmen, point)
    ReDim valP(segmen, point)
    
    
    ReDim cu(numData)
    ReDim cus(segmen)

    ReDim Gamma(numData)
    ReDim gammaS(segmen)
    
    ReDim depthI(segmen)
    
    ReDim km(segmen)
    
    ReDim valA(nodal)
    ReDim valPt(deltaPt)
    ReDim valC1(deltaPt)
    
    ReDim valB(nodal)
    
    Dim valD1 As Double
    Dim valD2 As Double
    Dim valD3 As Double
    
    ReDim valV(segmen + 4, deltaPt)
     
     
    For i = 1 To numData
        zi(i) = wsInput.Cells(5 + i, 3)
        zo(i) = wsInput.Cells(5 + i, 2)
        Gamma(i) = wsInput.Cells(5 + i, 4)
        cu(i) = wsInput.Cells(5 + i, 5)
    Next i
    
    
    deltaX = zi(numData) / segmen
    
    depthI(0) = 0
    For k = 1 To segmen
        depthI(k) = deltaX + depthI(0)
        depthI(0) = depthI(k)
        
        'Cells(17 - k, 39) = depthI(k)
    Next k
    

    For j = 1 To segmen
        For i = 1 To numData
            If j = 1 Then
                cus(j) = cu(1)
                gammaS(j) = Gamma(1)
            ElseIf j > 1 And j < segmen Then
                If depthI(j) > zo(i) And depthI(j) < zi(i) Then
                    cus(j) = cu(i)
                    gammaS(j) = Gamma(i)
                End If
            ElseIf j = segmen Then
                cus(j) = cu(numData)
                gammaS(j) = Gamma(numData)
            End If
        Next i
    Next j
    
    
    
    
    For j = 1 To segmen
        Pult1(j) = 9 * cus(j) * diameter
        Pult2(j) = (3 + (gammaS(j) * depthI(j) / cus(j)) + (consJ * depthI(j) / diameter)) * cus(j) * diameter

        If (Pult1(j) <= Pult2(j)) Then
            Pult(j) = Pult1(j)
        Else
            Pult(j) = Pult2(j)
        End If
        'Cells(3 + j, 23) = Pult(j)
    Next j


    For j = 1 To segmen
        If cus(j) < 48 Then
            epsilon50(j) = 0.02
        ElseIf 48 <= cus(j) And cus(j) <= 96 Then
            epsilon50(j) = 0.01
        ElseIf 96 <= cus(j) And cus(j) <= 192 Then
            epsilon50(j) = 0.005
        Else
            epsilon50(j) = 0.005
        End If
        'Cells(3 + j, 24) = epsilon50(j)
    Next j


    For j = 1 To segmen
        valY50(j) = 2.5 * diameter * epsilon50(j)
    Next j


    For j = 1 To segmen
        For k = 0 To point
            valY(j, k) = valY50(j) * k
        Next k
    Next j


    For j = 1 To segmen
        For k = 0 To point
            If k <= 8 Then
                valP(j, k) = 0.5 * (valY(j, k) / valY50(j)) ^ (1 / 3) * Pult(j)
            ElseIf k > 8 Then
                valP(j, k) = Pult(j)
            End If
        Next k
    Next j


    For k = 0 To point

        Pult0(k) = 3 * cu(1) * diameter
        valY0(k) = k * valY50(1)

        If k <= 8 Then
            valP0(k) = 0.5 * (k) ^ (1 / 3) * Pult0(k)
        ElseIf k > 8 Then
            valP0(k) = Pult0(k)
        End If
    Next k


    wsCalc.Cells.Clear


    For j = 1 To segmen
        For k = 0 To point

            wsCalc.Cells(1, 1) = 0

            wsCalc.Cells(1, (2 * j + 1)) = depthI(j)

            wsCalc.Cells(2, 1) = "Y (m)"
            wsCalc.Cells(2, 2) = "P (kN/m)"

            wsCalc.Cells(2, (2 * j + 1)) = "Y (m)"
            wsCalc.Cells(2, (2 * j + 2)) = "P (kN/m)"

            wsCalc.Cells(3 + k, 1) = valY0(k)
            wsCalc.Cells(3 + k, 2) = valP0(k)

            wsCalc.Cells(3 + k, (2 * j + 1)) = valY(j, k)
            wsCalc.Cells(3 + k, (2 * j + 2)) = valP(j, k)
        Next k
    Next j


    km0 = (valP0(1) - valP0(0)) / (valY0(1) - valY0(0))

    'Cells(5, 12) = km0

    For j = segmen To 1 Step -1         'dibalik karena nodal 0 dimulai dari bawah
        km(j) = (valP(j, 1) - valP(j, 0)) / (valY(j, 1) - valY(j, 0))
        'Cells(5 + j, 12) = km(j)
    Next j




    For j = segmen To 1 Step -1         'dibalik karena nodal 0 dimulai dari bawah

        valA(0) = km0 * (deltaX) ^ 4 / (modulus * inertia)
        valA(j) = km(j) * (deltaX) ^ 4 / (modulus * inertia)
        'Cells(5, 9) = valA(0)
        'Cells(5 + j, 9) = valA(j)

    Next j

    For l = 1 To deltaPt

        valPt(l) = pt / (deltaPt + 1 - l)
        'Cells(5 + l, 8) = valPt(l)
    Next l

    For l = 1 To deltaPt

        valC1(l) = 2 * valPt(l) * (deltaX ^ 3) / (modulus * inertia)
        'Cells(5 + l, 15) = valC1(l)
    Next l


    For m = 0 To nodal
        If m = 0 Then
            valB(m) = 2 / (valA(segmen) + 2)
        ElseIf m = 1 Then
            valB(m) = 2 * valB(m - 1)
        ElseIf m = 2 Then
            valB(m) = 1 / (5 + valA(segmen - (m - 1)) - (2 * valB(m - 1)))
        ElseIf m > 2 Then
            If m Mod 2 > 0 Then
                mi = (m - 1) / 2
                valB(m) = valB(2 * mi) * (4 - valB((2 * mi) - 1))
            ElseIf m Mod 2 = 0 Then
                mi = (m / 2)
                valB(m) = 1 / (6 + valA(segmen - mi) - valB(2 * mi - 4) - (valB(2 * mi - 1) * (4 - valB(2 * mi - 3))))
            End If
        End If

        'Cells(5 + m, 10) = valB(m)
    Next m

'Cells(5 + 20, 12) = valB(0)

    valD1 = 1 / valB(2 * segmen)


    valD2 = valD1 * valB(2 * segmen + 1) - valB(2 * segmen - 2) * (2 - valB(2 * segmen - 3)) - 2


    valD3 = valD1 - valB(2 * segmen - 4) - valB(2 * segmen - 1) * (2 - valB(2 * segmen - 3))



    'point 12
    For l = 1 To deltaPt

        m = segmen

        valV(m + 2, l) = valC1(l) * (1 + valB(2 * m - 2)) / (valD3 * (1 + valB(2 * m - 2)) - valD2 * valB(2 * m - 1))

        'Cells(17, 33 + l) = valV(MM + 2, l)

    Next l


    'point 13
    For l = 1 To deltaPt

        m = segmen

        valV(m + 3, l) = valB(2 * m - 1) * valV(m + 2, l) / (1 + valB(2 * m - 2))

        'Cells(18, 33 + l) = valV(MM + 3, l)

    Next l


    'point 14
    For l = 1 To deltaPt

        m = segmen

        valV(m + 4, l) = valD1 * valB(2 * m + 1) * valV(m + 1, l) - valD1 * valV(m + 2, l)

        'Cells(19, 33 + l) = valV(MM + 4, l)

    Next l

    'point 11 sampai 2
    For l = 1 To deltaPt
        For m = segmen - 1 To 0 Step -1

            valV(m + 2, l) = -valB(2 * m) * valV(m + 3, l) + valB(2 * m + 1) * valV(m + 1, l)

            'Cells(7 + MM, 33 + l) = valV(MM + 2, l)

        Next m
    Next l






    'point 1 dan 0
    For l = 1 To deltaPt

        valV(1, l) = 2 * valV(2, l) - valV(3, l)

        'Cells(6, 33 + l) = valV(1, l)
        'Cells(6, 40 + l) = 1

        valV(0, l) = valV(4, l) - 4 * (valV(3, l)) + 4 * (valV(2, l))

        'Cells(5, 33 + l) = valV(0, l)
        'Cells(5, 40 + l) = 0
    Next l






    'PRINT........................
    
    wsReport.Cells.Clear
    
    
    wsReport.Cells(1, 1) = "Depth (m)"

    For m = 1 To segmen

        wsReport.Cells(2, 1) = 0
        wsReport.Cells(2 + m, 1) = depthI(m)

    Next m


    For l = 1 To deltaPt
        For m = 0 To segmen - 1

            wsReport.Cells(segmen + 2 - m, 1 + l) = valV(m + 2, l) * 100

        Next m
    Next l


    For l = 1 To deltaPt

        m = segmen

        wsReport.Cells(2, 1 + l) = valV(m + 2, l) * 100

    Next l


    wsReport.Cells(1, 24) = "Load (kN)"
    wsReport.Cells(2, 24) = 0

    wsReport.Cells(1, 25) = "Deflection (cm)"
    wsReport.Cells(2, 25) = 0

    For l = 1 To deltaPt

        wsReport.Cells(1, 1 + l) = "Load " & l

        wsReport.Cells(2 + l, 24) = valPt(l)
        wsReport.Cells(2 + l, 25) = valV(m + 2, l) * 100

    Next l
    
    ReDim Momen(segmen, deltaPt)
    ReDim Shear(segmen, deltaPt)
    ReDim soilResistance(segmen, deltaPt)
    
    
    For l = 1 To deltaPt
    wsReport.Cells(1, 6 + l) = "Momen (kN.m)"
    wsReport.Cells(1, 11 + l) = "Shear (kN)"
    wsReport.Cells(1, 16 + l) = "Soil Resistance (kN/m)"
        For m = segmen To 0 Step -1
            Momen(m, l) = (valV(m + 1, l) - 2 * valV(m + 2, l) + valV(m + 3, l)) * (modulus * inertia / deltaX ^ 2)
            wsReport.Cells(segmen + 2 - m, 6 + l) = Momen(m, l)
            
            Shear(m, l) = (valV(m, l) - 2 * valV(m + 2, l) + 2 * valV(m + 3, l) - valV(m + 4, l)) * (modulus * inertia / (2 * deltaX ^ 3))
            wsReport.Cells(segmen + 2 - m, 11 + l) = Shear(m, l)
            
            soilResistance(m, l) = (valV(m, l) - 4 * valV(m + 1, l) + 6 * valV(m + 2, l) - 4 * valV(m + 3, l) + valV(m + 4, l)) * (modulus * inertia / deltaX ^ 4)
            wsReport.Cells(segmen + 2 - m, 16 + l) = soilResistance(m, l)
        Next m
    Next l
    
    
    
    
End Sub




     




