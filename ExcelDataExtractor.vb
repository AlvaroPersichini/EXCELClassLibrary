Option Explicit On
Public Class ExcelDataExtractor

    Function ExtractData(oSheet As Microsoft.Office.Interop.Excel.Worksheet) As Dictionary(Of String, (NewPartNumber As String, DescriptionRef As String, Quantity As Integer, Source As Integer, Nomenclature As String, Definition As String))

        Dim lastRow As Integer = GetLastRow(oSheet)

        Dim oDic As New Dictionary(Of String, (NewPartNumber As String, DescriptionRef As String, Quantity As Integer, Source As Integer, Nomenclature As String, Definition As String))

        ' CUIDADO esto hay que manejarlo de otra manera.
        ' esta hardcodeado el 3 porque es la fila donde empiezan los datos,
        ' pero si cambia el formato del excel esto puede fallar.
        If lastRow < 3 Then Return oDic

        For i As Integer = 3 To lastRow

            Dim key As String = oSheet.Cells(i, 4).Text.ToString().Trim()

            If Not String.IsNullOrWhiteSpace(key) AndAlso Not oDic.ContainsKey(key) Then
                Dim itemData = (
                NewPartNumber:=oSheet.Cells(i, 5).Text.ToString().Trim(),
                DescriptionRef:=oSheet.Cells(i, 6).Text.ToString().Trim(),
                Quantity:=CInt(Val(oSheet.Cells(i, 7).Text.ToString())),
                Source:=CInt(Val(oSheet.Cells(i, 8).Text.ToString())),
                Nomenclature:=oSheet.Cells(i, 10).Text.ToString().Trim(),
                Definition:=oSheet.Cells(i, 11).Text.ToString().Trim()
            )
                oDic.Add(key, itemData)
            End If
        Next

        Return oDic

    End Function



    Private Function GetLastRow(oSheet As Microsoft.Office.Interop.Excel.Worksheet) As Integer
        Try
            Dim lastCell = oSheet.Cells.Find("*", , , ,
                Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious)
            If lastCell Is Nothing Then Return 0
            Return lastCell.Row
        Catch
            Return 0
        End Try
    End Function
End Class