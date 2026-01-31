Function SelectHuaweiInverter(kWp As Double, pvModel As String) As String ' Ghi chu: Hay dam bao ham co tham so pvModel
    ' Ghi chu: Xu ly truong hop cong suat khong hop le hoac rat nho
    If kWp <= 0 Then
        SelectHuaweiInverter = "Cong suat khong hop le"
        Exit Function
    End If
    '        End If
    '    Case Else ' Ghi chu: Logic mac dinh cho cac model PV khac
            If kWp <= 9.6 Then ' 8kW inverter: 8kW * 1.2 = 9.6kW.
                SelectHuaweiInverter = "SUN2000-8K-MAP0"
            ElseIf kWp <= 10.6 Then ' 10kW inverter: Gi?i h?n du?i 11.3 kWp.
                SelectHuaweiInverter = "SUN2000-10K-MAP0"
            ElseIf kWp <= 14.4 Then ' 12kW inverter: Gi?i h?n du?i 14.2 kWp.
                SelectHuaweiInverter = "SUN2000-12K-MAP0"
            ElseIf kWp <= 16 Then ' 15kW inverter: Bao gom 20 PV (14.2kWp) va 21, 22, 24, 25, 26 PV (t?i 18.46kWp).
                SelectHuaweiInverter = "SUN2000-15KTL-M5"
            ElseIf kWp <= 25 Then ' 20kW inverter: Bao gom 27, 28, 30, 32, 34, 35, 36 PV (t?i 25.56kWp).
                SelectHuaweiInverter = "SUN2000-20KTL-M5"
            ' Ghi chu: Cap nhat NGUONG QUAN TRONG nay cho 30KTL-M3
            ElseIf kWp <= 36 Then ' Neu kWp <= 39.5, chon 30KTL-M3. Vd: 55 PV * 0.665 = 36.575 kWp.
                SelectHuaweiInverter = "SUN2000-30KTL-M3"
            ' Ghi chu: Cap nhat NGUONG QUAN TRONG nay cho 40KTL-M3
            ElseIf kWp <= 48 Then ' Neu kWp <= 51.5, chon 40KTL-M3. Vd: 72 PV * 0.665 = 47.88 kWp.
                SelectHuaweiInverter = "SUN2000-40KTL-M3"
            ElseIf kWp <= 60 Then ' 50kW inverter: 50kW * 1.2 = 60kW.
                SelectHuaweiInverter = "SUN2000-50KTL-M3"
            Else ' Neu kWp lon hon tat ca cac muc tren
                SelectHuaweiInverter = "Khong tim thay Inverter phu hop"
            End If
    ' End Select ' Ghi chu: Dong End Select nay chi su dung neu ban co Select Case pvModel
End Function
' Ham dinh dang o
Sub FormatCell(rng As Range, fontName As String, fontSize As Double, Optional isBold As Boolean = False, _
                Optional isItalic As Boolean = False, Optional hAlign As XlHAlign = xlLeft)
    On Error GoTo ErrHandler
    With rng
        .Font.Name = fontName
        .Font.Size = fontSize
        .Font.Bold = isBold
        .Font.Italic = isItalic
        .HorizontalAlignment = hAlign
    End With
    Exit Sub
ErrHandler:
    MsgBox "Loi khi dinh dang o: " & Err.description & vbCrLf & "So loi: " & Err.Number, vbCritical
End Sub

' Ham dinh dang tieu de nhom
Sub FormatGroupHeader(ws As Worksheet, row As Long)
    With ws.Range("A" & row & ":G" & row)
        .Merge
        .Font.Name = "Cambria"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .Interior.Color = vbYellow
    End With
End Sub

' Ham lay so luong tu BOQ
Function GetBOQQuantity(materials As Collection, materialCode As String) As Variant
    On Error Resume Next
    GetBOQQuantity = materials(materialCode)
    If Err.Number <> 0 Then GetBOQQuantity = ""
    On Error GoTo 0
End Function

' Ham tra ve key tuong ung voi cot
Function GetMaterialKey(columnIndex As Long) As String
    Select Case columnIndex
        Case 4: GetMaterialKey = "ER-R-ECO/5200"
        Case 5: GetMaterialKey = "ER-SP-ECO"
        Case 6: GetMaterialKey = "I-05-6.3/85/CS"
        Case 7: GetMaterialKey = "ER-EC-ST35"
        Case 8: GetMaterialKey = "ER-IC-ST35"
        Case 9: GetMaterialKey = "EZ-CC-PV/4"
        Case 10: GetMaterialKey = "EZ-GC-ST"
        Case 11: GetMaterialKey = "EZ-GL-ST"
        Case 12: GetMaterialKey = "CS7N-665MS"
        Case 13: GetMaterialKey = "T4-T6"
        Case 14: GetMaterialKey = "SUN2000"
        Case 15: GetMaterialKey = "Sdongle5A"
        Case 16: GetMaterialKey = "DTSU"
        Case 17: GetMaterialKey = "CXV-3x16-10sqrt"
        Case 18: GetMaterialKey = "C10sqrt"
        Case 19: GetMaterialKey = "Leader"
        Case 20: GetMaterialKey = "LS"
        Case 21: GetMaterialKey = "1/2-3/4"
        Case 22: GetMaterialKey = "JUNC_BOX"
        Case 23: GetMaterialKey = "SUNTREE"
        Case Else: GetMaterialKey = ""
    End Select
End Function
Function GetPVInfo(pvModel As String) As Collection
    Dim pvInfo As New Collection
    Dim description As String
    Dim manufacturer As String
    Dim unit As String
    
    ' Khoi tao gia tri mac dinh de xu ly truong hop khong tim thay
    description = "Mo ta PV khong tim thay cho model: " & pvModel
    manufacturer = "Khong xac dinh"
    unit = "panel" ' Mac dinh don vi tinh la panel

    Select Case UCase(Trim(pvModel)) ' Su dung UCase va Trim de xu ly khong phan biet chu hoa/thuong va khoang trang
        Case "CS7N-665MS"
            manufacturer = "CANADIAN SOLAR"
            description = "HiKu7 Mono PERC, Dimensions 2384 x 1303 x 35 mm," & vbLf & _
                          "Nominal Max. Power (Pmax) 665 W, Module efficiency 21.4%," & vbLf & _
                          "Opt. Operating Voltage (Vmp) 38.5 V, Opt. Operating Current (Imp) 17.28 A," & vbLf & _
                          "Open Circuit Voltage (Voc) 45.6 V, Short Circuit Current (Isc) 18.51 A"
        
        Case "CHSM72M-HC-555W"
            manufacturer = "Astronergy"
            description = "ASTRO 5 Semi, CHSM72M-HC Monofacial Series, Dimensions 2278 x 1134 x 35 mm," & vbLf & _
                          "Nominal Power (Pmax) 555 W, Module efficiency 21.5%," & vbLf & _
                          "Rated Voltage (Vmpp) 42.27 V, Rated Current (Impp) 13.13 A," & vbLf & _
                          "Open circuit voltage (Voc) 50.30 V, Short circuit current (Isc) 13.98 A"
                          
        Case "CHSM72M-HC-550W"
            manufacturer = "Astronergy"
            description = "ASTRO 5 Semi, CHSM72M-HC Monofacial Series, Dimensions 2278 x 1134 x 35 mm," & vbLf & _
                          "Nominal Power (Pmax) 550 W, Module efficiency 21.3%," & vbLf & _
                          "Rated Voltage (Vmpp) 42.10 V, Rated Current (Impp) 13.06 A," & vbLf & _
                          "Open circuit voltage (Voc) 50.10 V, Short circuit current (Isc) 13.90 A"
                          
        Case "CHSM78M-HC-590W"
            manufacturer = "Astronergy"
            description = "ASTRO 5 Twins, CHSM78M(DG)/F-BH Bifacial Series, Dimensions 2465 x 1134 x 30 mm," & vbLf & _
                          "Rated output (Pmpp) 590 Wp, Module efficiency 21.3%," & vbLf & _
                          "Rated voltage (Vmpp) 45.60 V, Rated current (Impp) 12.94 A," & vbLf & _
                          "Open circuit voltage (Voc) 54.26 V, Short circuit current (Isc) 13.75 A"
                          
        Case "HSM-TH665W"
            ' Vui long cap nhat thong tin tu datasheet Huasun chinh thuc neu co
            manufacturer = "TCL/China"
            description = "Himalaya Series (TH-Type), N-type Bifacial Module," & vbLf & _
                          "Maximum Power 665 W, TOPCon technology, (Thong so kich thuoc va dien chi tiet can cap nhat tu datasheet)"
