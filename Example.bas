

Sub fetch()
    Dim m As New xlMiner
    m.SZSE_Dividend("2").dump "1", 1, 2
  
    Dim f As New FormatUtil
    f.formatRng Worksheets("1").Cells(2, 1).CurrentRegion
    
    Set m = Nothing
    Set f = Nothing
End Sub


Sub trail()
    
    Dim f As New FormatUtil
    
    Dim m As New xlMiner
    m.fs("2", 2018, 2).toRng Cells(1, 1)
    f.formatRng Cells(1, 1).CurrentRegion
   
   
    m.fs("2", 2018, 2, 1).toRng Cells(1, 3)
    f.formatRng Cells(1, 3).CurrentRegion

    
     m.fs("2", 2018, 2, 2).toRng Cells(1, 5)
    f.formatRng Cells(1, 5).CurrentRegion

End Sub





