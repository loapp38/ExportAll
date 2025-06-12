Sub ExportAllPTQueries()
    Dim exportList As Variant
    Dim i As Integer
    Dim exportPath As String
    Dim queryName As String
    Dim fileName As String
    
    exportList = Array( _
    "Query01", "Query02", "Query03", "Query04", "Query05", _
    "Query06", "Query07", "Query08", "Query09", "Query10", _
    "Query11", "Query12", "Query13", "Query14", "Query15", _
    "Query16", "Query17", "Query18", "Query19", "Query20", _
    "Query21", "Query22", "Query24", "Query25", _
    "Query26", "Query27", "Query28", "Query29", "Query30", _
    "Query31", "Query32", "Query33", "Query34", "Query35", _
    "Query36" _
    )
    
    exportPath = "\\F02PRDRCASVM01.sunlifecorp.com\cdss_prd_dmu\Member Health Check\TABLEAU - OASIS ADM Project\Access Extract\CSV Extract\"
    
    
    For i = LBound(exportList) To UBound(exportList)
        queryName = exportList(i)
        fileName = exportPath & queryName & ".csv"
        
        Debug.Print "Exporting " & queryName & " to " & fileName
        
        DoCmd.TransferText _
            TransferType:=acExportDelim, _
            TableName:=queryName, _
            fileName:=fileName, _
            HasFieldNames:=True
        Next i
        
End Sub
