Attribute VB_Name = "Clay03"
Option Explicit

public diameter as Double 'unit in m
public modulus as double 'unit in m
public inertia as double 'unit in m4
public pt as double 'unit in kN
public i, j, k, l, n, m as integer

dim wsInput as worksheet
dim wsReport as worksheet
dim wsCalc as worksheet

dim deltaPt as integer
dim point as integer

const pi = 3.14159265358979
const consJ = 0.5

function valDiameter()

    set wsInput = worksheets("Input")
    valDiameter = wsInput.cells(1,2)

end function

function valModulus()

    set wsInput = worksheets("Input")
    valModulus = wsInput.cells(2,2)

end function

function valInertia()

    set wsInput = worksheets("Input")
    valInertia = wsInput.cells(3,2)

end function

function valPt()

    set wsInput = worksheets("Input")
    valPt = wsInput.cells(1,6)

end function

function valSegmen()

    set wsInput = worksheets("Input")
    valSegmen = wsInput.cells(2,6)

end function

function nData()

    set wsInput = worksheets("Input")
    nData = wsInput.application.range(cells(6,1), cells(6,1).end(xlDown)).rows.count

end function

function zoI(nI as integer)

    set wsInput = worksheets("Input")
    zoI = wsInput.application.range(cells(6,2), cells(6,2).end(xlDown)).cells(nI)

end function

function ziI(nI as integer)

    set wsInput = worksheets("Input")
    ziI = wsInput.application.range(cells(6,3), cells(6,3).end(xlDown)).cells(nI)

end function

function gammaI(nI as integer)

    set wsInput = worksheets("Input")
    gammaI = wsInput.application.range(cells(6,4), cells(6,4).end(xlDwon)).cells(nI)

end function

function cuI(nI as integer)

    set wsInput = worksheets("Input")
    cuI = wsInput.application.range(cells(6,5), cells(6,5).end(xlDown)).cells(nI)

end function

function nodal()

    nodal = (2*valSegmen()) + 1

end function

function deltaX()

    deltaX = ziI(nData())/valSegmen()

end function


sub softClay()

redim nDepthI(nodal())
redim gammas(valSegmen())
redim cus(valSegmen())
redim layer(valSegmen())
redim gammaDepth(valSegmen())
redim gammaDepth(valSegmen())
redim gammaAvg(valSegmen())


    nDepthI(0) = 0                      'menghitung kedalaman tiap titik nodal
    for i = 1 to (nodal() - 1)
        nDepth(i) = deltaX() + nDepthI(0)
        nDepthI(0) = nDepth(i) 
    next i

    for i = 0 to (valSegmen() - 1)      'menentukan nilai c dan gamma untuk setiap lapisan
        for j = 0 to (nData() - 1)
            if i = 0 then
                cus(i) = cuI(0)
                gammas(i) = gammaI(0)
            elseif i > 0 and i < (valSegmen() - 1) then
                if nDepth(i) > zoI(j) and nDepth(i) < ziI(j) then
                    cus(i) = cuI(j)
                    gammas(i) = gammaI(j)
                end if
            elseif i = (valSegmen() - 1) then
                cus(i) = cuI(nData() - 1)
                gammas(i) = gammaI(nData() - 1)
            end if
        next j
    next i

    gammaDepth(0) = 0                   'menghitung gamma average
    for i = 0 to (valSegmen() - 1)
        if i = 0 then
            layer(i) = nDepth(0)
        elseif i = (valSegmen() - 1) then
            layer(i) = nDepth(valSegmen() - 1) - nDepth(valSegmen() - 2)
        else
            layer(i) = nDepth(i+1) - nDepth(i)
        end if

        gammaDepth(i) = gammaDepth(0) + (layer(i)*gammas(i))
        gammaDepth(0) = gammaDepth(i)

        gammaAvg(i) = gammaDepth(i) / nDepth(i)
    next i



end sub
