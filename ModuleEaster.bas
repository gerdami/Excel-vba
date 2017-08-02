Attribute VB_Name = "ModuleEaster"
Option Explicit

Public Function EASTER(year As Single) As Date
    'Based on Claus Tondering algorithm interpretation.
    'See http://www.tondering.dk/claus/cal/calendar29.html
    'Norman Harker 10-Jul-2004
    '
    'Peltier: http://peltiertech.com/calculating-easter/
    '
    Dim G As Integer: Dim C As Integer: Dim H As Integer
    Dim I As Integer: Dim J As Integer: Dim L As Integer
    Dim EM As Integer: Dim ED As Integer
    Dim Adj1904 As Integer
    G = year Mod 19
    C = year \ 100
    H = (C - C \ 4 - (8 * C + 13) \ 25 + 19 * G + 15) Mod 30
    I = H - (H \ 28) * (1 - (29 \ (H + 1)) * ((21 - G) \ 11))
    J = (year + year \ 4 + I + 2 - C + C \ 4) Mod 7
    L = I - J
    EM = 3 + (L + 40) \ 44
    ED = L + 28 - (31 * (EM \ 4))
    If ActiveWorkbook.Date1904 = True Then
        Adj1904 = 365 * 4 + 2
    End If
    EASTER = DateSerial(year, EM, ED) - Adj1904
End Function
