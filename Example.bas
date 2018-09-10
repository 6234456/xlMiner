' Dicts/Lists required

Sub fetch()
    Dim m As New xlMiner
    m.SZSE_Dividend("2").dump "1", 1, 2
  
    Dim f As New FormatUtil
    f.formatRng Worksheets("1").Cells(2, 1).CurrentRegion
    
    Set m = Nothing
    Set f = Nothing
End Sub
