Sub query()
    main Worksheets("Overview").Cells(1, 4).value
End Sub

Function main(ByVal code As String)

    If Len(Trim(code)) > 0 Then
    
        Dim f As New FormatUtil
        Dim m As New xlMiner
        
        Dim y As Integer
        y = 2018
        
        Dim q As Integer
        q = 4
        
        Dim d As Dicts
        
        Set d = m.profile(code)
        d.p
    
        With Worksheets("Overview")
            .Cells.clear
            .Cells(1, 1) = d.dict.Item("ORGNAME")
            .Cells(1, 2) = y & "Q" & q
            m.fs(code, 2018, 4, fsType.BALANCE_STMT).toRng .Cells(2, 1)
            
            f.formatRng Cells(1, 1).CurrentRegion
            
            With .Cells(1, 1).CurrentRegion
                .Columns.AutoFit
                
                With .Columns(2)
                    .NumberFormat = "#,###"
                    .HorizontalAlignment = xlRight
                End With
            End With
        End With
    
    End If

End Function
