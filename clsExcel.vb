
Imports System
Imports System.Diagnostics
Imports System.Collections.Generic
Imports System.Text
Imports Microsoft.Office.Interop.Excel

Public Class clsExcel
    Private xlApp As Application
    Private xlSheet As Worksheet
    Private processId As Integer = 0

    ''' <summary>
    ''' Launch excel application and add new workbook and link active worksheet.
    ''' </summary>
    ''' <param name="con"></param>
    ''' <remarks></remarks>
    Public Sub New(ByRef con As Integer)
        Try
            Dim oId As Integer = excelAppId()
            xlApp = Interaction.CreateObject("Excel.Application")
            Dim nId As Integer = excelAppId()
            'Store current excelapplication Id.
            processId = nId - oId
            con = 1
        Catch
            con = -1
            MsgBox("Excel application is not properly installed in this computer.")
            Return
        End Try

        xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
        xlSheet = xlApp.ActiveWorkbook.ActiveSheet
    End Sub

    ''' <summary>
    ''' Launch Excel application and open the given excel file.
    ''' </summary>
    ''' <param name="con"></param>
    ''' <param name="fileName"></param>
    ''' <remarks></remarks>
    Public Sub New(ByRef con As Integer, ByVal fileName As String)
        Try
            Dim oId As Integer = excelAppId()
            xlApp = Interaction.CreateObject("Excel.Application")
            Dim nId As Integer = excelAppId()
            'Store current excelapplication Id.
            processId = nId - oId
            con = 1
        Catch
            con = -1
            MsgBox("Excel application is not properly installed in this computer.")
            Return
        End Try

        xlApp.Workbooks.Open(fileName)
        xlSheet = xlApp.ActiveWorkbook.ActiveSheet
    End Sub

    ''' <summary>
    ''' Link already opened active excel application.
    ''' </summary>
    ''' <param name="linkXls"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal linkXls As Boolean)
        Try
            xlApp = Interaction.GetObject(, "Excel.Application")
        Catch
            MsgBox("Excel application is not properly installed in this computer.")
            Return
        End Try

        xlSheet = xlApp.ActiveWorkbook.ActiveSheet
    End Sub

    Public Sub quitExcel()
        'if the normal process does not kill the process then second time process killing
        If (processId > 0) Then
            Dim proc As Process = Process.GetProcessById(processId)
            proc.Kill()
        End If
    End Sub

    Protected Overrides Sub Finalize()
        xlApp = Nothing
        xlSheet = Nothing
        MyBase.Finalize()
    End Sub

    ' <summary>
    ' Find the sum of Id of all the excel application.
    ' </summary>
    ' <returns></returns>
    Public Function excelAppId() As Integer
        Dim pId As Integer = 0

        For Each cP As Process In Process.GetProcessesByName("EXCEL")
            pId += cP.Id
        Next

        Return pId
    End Function

    ' <summary>
    ' Find tthe Id of current excel application.
    ' </summary>
    ' <returns></returns>
    Public Function currentExcelAppId() As Integer
        For Each cP As Process In Process.GetProcessesByName("EXCEL")
            Return cP.Id
        Next

        Return 0
    End Function

    ' <summary>
    ' Add new excel sheet before the activesheet to activeworkbook.
    ' </summary>
    ' <param name="sheetName"></param>
    ' <returns></returns>
    Public Function addnewSheet(ByVal sheetName As String) As Worksheet
        Dim curSheet As Worksheet = xlApp.ActiveWorkbook.Sheets.Add()
        If (sheetName <> String.Empty) Then
            sheetName = sheetName.Replace("/", "|")
            sheetName = sheetName.Replace("\", "|")
            sheetName = sheetName.Replace("*", "|")
            sheetName = sheetName.Replace("?", "|")
            sheetName = sheetName.Replace("[", "|")
            sheetName = sheetName.Replace("]", "|")
            curSheet.Name = sheetName
        End If

        Return curSheet
    End Function

    ' <summary>
    ' Set current excel sheet for working.
    ' if sheet Name is given that will be current sheet otherwise activesheet will be the current.
    ' </summary>
    ' <param name="sheetName"></param>
    Public Sub setCurrentSheet(ByVal sheetName As String)
        If (sheetName = String.Empty) Then
            xlSheet = xlApp.ActiveWorkbook.ActiveSheet
        Else
            'xlSheet = xlApp.ActiveWorkbook.Sheets.Item(sheetName)
            xlSheet = xlApp.ActiveWorkbook.Sheets(sheetName)
        End If
    End Sub

    Public Sub copySheet(ByVal sheetToCopy As String, ByVal newSheetName As String)
        'Excel spreadsheet syntax
        Try
            xlApp.Worksheets(sheetToCopy).Copy(After:=xlApp.Worksheets(sheetToCopy))
            If newSheetName.Trim() <> String.Empty Then
                Dim nSheet As Worksheet = xlApp.ActiveWorkbook.ActiveSheet
                nSheet.Name = newSheetName
            End If
        Catch
            'MsgBox("Wrong template file.")
        End Try
    End Sub

    Public Sub copySheet(ByVal sheetToCopy As String, ByVal newSheetName As String, ByVal boolAfter As Boolean)
        'Excel spreadsheet syntax
        Try
            If boolAfter Then
                xlApp.Worksheets(sheetToCopy).Copy(After:=xlApp.ActiveSheet)
            Else
                xlApp.Worksheets(sheetToCopy).Copy(Before:=xlApp.ActiveSheet)
            End If

            If newSheetName.Trim() <> String.Empty Then
                Dim nSheet As Worksheet = xlApp.ActiveWorkbook.ActiveSheet
                nSheet.Name = newSheetName
            End If
        Catch
            'MsgBox("Wrong template file.")
        End Try
    End Sub

    ''' <summary>
    ''' Insert the sheet after the active sheet.
    ''' </summary>
    ''' <param name="newSheetName"></param>
    ''' <remarks></remarks>
    Public Sub addWorkSheet(ByVal newSheetName As String)
        'Excel spreadsheet syntax
        Try
            Dim nSheet As Worksheet = xlApp.ActiveWorkbook.ActiveSheet
            Dim curSheet As Worksheet = xlApp.ActiveWorkbook.Sheets.Add(, nSheet)
            If newSheetName.Trim() <> String.Empty Then
                curSheet.Name = newSheetName
            End If
        Catch
            'MsgBox("Wrong template file.")
        End Try
    End Sub

    Public Sub hideSheet(ByVal sheetTohide As String)
        Try
            xlApp.Worksheets(sheetTohide).Visible = False
        Catch
            'do nothing.
        End Try
    End Sub

    Public Sub hideColumn(ByVal xlSheet As Worksheet, ByVal c As Integer)
        Try
            xlSheet.Columns(c).Hidden = True
        Catch
            'do nothing.
        End Try
    End Sub

    Public Sub hideRow(ByVal xlSheet As Worksheet, ByVal r As Integer)
        Try
            xlSheet.Rows(r).Hidden = True
        Catch
            'do nothing.
        End Try
    End Sub

    ' <summary>
    ' Write range of cells of activeworksheet.
    ' </summary>
    ' <param name="uD"></param>
    ' <param name="R"></param>
    ' <param name="C"></param>
    Public Sub writeRangeData(ByVal uD()() As Object, ByVal R As Integer, ByVal C As Integer)
        Try
            Dim n As Long = uD.GetLength(0)
            For i As Long = 0 To n - 1
                For j As Long = 0 To uD.GetLength(1) - 1
                    xlSheet.Cells(R + i, C + j).Value = uD(i)(j)
                Next
            Next
        Catch
            'do nothing.
        End Try
    End Sub

    ' <summary>
    ' Write single cell of the activeworkbook.
    ' </summary>
    ' <param name="uD"></param>
    ' <param name="R"></param>
    ' <param name="C"></param>
    Public Sub writeCell(ByVal uD As Object, ByVal R As Integer, ByVal C As Integer)
        xlSheet.Cells(R, C).Value = uD
    End Sub

    Public Sub writeCell(ByVal uD As Object, ByVal R As Integer, ByVal C As Integer, ByVal fontSize As Integer, _
                ByVal bold As Boolean, ByVal italic As Boolean, ByVal border As Boolean)
        Dim cCell As Range = xlSheet.Cells(R, C)

        cCell.Value2 = uD
        cCell.Font.Size = fontSize
        cCell.Font.Italic = italic
        cCell.Font.Bold = bold
        If border Then
            cCell.BorderAround(, XlBorderWeight.xlThick)
        End If
    End Sub

    ' <summary>
    ' Write single cell of activeworkbook with multiple attributes for cell.
    ' </summary>
    ' <param name="uD"></param>
    ' <param name="R"></param>
    ' <param name="C"></param>
    ' <param name="fontSize"></param>
    ' <param name="bold"></param>
    ' <param name="italic"></param>
    ' <param name="fontColor"></param>
    ' <param name="backColor"></param>
    ' <param name="align"></param>
    ' <param name="txtFormat"></param>
    Public Sub writeCell(ByVal uD As Object, ByVal R As Integer, ByVal C As Integer, ByVal fontSize As Integer, _
            ByVal bold As Boolean, ByVal italic As Boolean, ByVal fontColor As Integer, ByVal backColor As Integer, ByVal align As Integer, ByVal txtFormat As String)
        Dim cCell As Range = xlSheet.Cells(R, C)

        If (txtFormat <> String.Empty) Then
            If (Convert.IsDBNull(uD)) Then uD = 0
            cCell.Value2 = Microsoft.VisualBasic.Format(uD, txtFormat) 'String.Format(txtFormat, uD)
        Else
            cCell.Value2 = uD
        End If

        cCell.Font.Size = fontSize
        cCell.Font.Italic = italic
        cCell.Font.Bold = bold
    End Sub

    Public Sub writeCell(ByVal uD As Object, ByVal R As Integer, ByVal C As Integer, ByVal txtFormat As String)
        Dim cCell As Range = xlSheet.Cells(R, C)

        If (txtFormat <> String.Empty) Then
            If (Convert.IsDBNull(uD)) Then uD = 0
            cCell.Value2 = Microsoft.VisualBasic.Format(uD, txtFormat) 'String.Format(txtFormat, uD)
        Else
            cCell.Value2 = uD
        End If
    End Sub

    Public Sub writeCell(ByVal uD As Object, ByVal R As Integer, ByVal C As Integer, ByVal txtFormat As String, ByVal bold As Boolean)
        Dim cCell As Range = xlSheet.Cells(R, C)

        If (txtFormat <> String.Empty) Then
            If (Convert.IsDBNull(uD)) Then uD = 0
            cCell.Value2 = Microsoft.VisualBasic.Format(uD, txtFormat) 'String.Format(txtFormat, uD)
        Else
            cCell.Value2 = uD
        End If

        cCell.Font.Bold = bold
    End Sub

    ' <summary>
    ' But border around the cell.
    ' </summary>
    ' <param name="uD"></param>
    ' <param name="R"></param>
    ' <param name="C"></param>
    Public Sub putBorder(ByVal R As Integer, ByVal C As Integer)
        Dim cCell As Range = xlSheet.Cells(R, C)
        cCell.BorderAround(XlLineStyle.xlDot, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic)
    End Sub

    ' <summary>
    ' Set the width of row column.
    ' </summary>
    ' <param name="uD"></param>
    ' <param name="R"></param>
    ' <param name="C"></param>
    ' <param name="cWidth"></param>
    Public Sub setColumnWidth(ByVal R As Integer, ByVal C As Integer, ByVal cWidth As Integer)
        Dim cCell As Range = xlSheet.Cells(R, C)
        cCell.ColumnWidth = cWidth
    End Sub

    ' <summary>
    ' Set the height of the row.
    ' </summary>
    ' <param name="uD"></param>
    ' <param name="R"></param>
    ' <param name="C"></param>
    ' <param name="rHeight"></param>
    Public Sub setRowHeight(ByVal R As Integer, ByVal C As Integer, ByVal rHeight As Integer)
        Dim cCell As Range = xlSheet.Cells(R, C)
        cCell.RowHeight = rHeight
    End Sub

    ' <summary>
    ' Merges the cells.
    ' </summary>
    ' <param name="R1"></param>
    ' <param name="C1"></param>
    ' <param name="R2"></param>
    ' <param name="C2"></param>
    Public Sub mergeCells(ByVal R1 As Integer, ByVal C1 As Integer, ByVal R2 As Integer, ByVal C2 As Integer)
        Dim cCell As Range = makeRange(R1, C1, R2, C2)
        cCell.Merge()
    End Sub

    ' <summary>
    ' Create Range object for active sheet.
    ' </summary>
    ' <param name="R1"></param>
    ' <param name="C1"></param>
    ' <param name="R2"></param>
    ' <param name="C2"></param>
    ' <returns></returns>
    Public Function makeRange(ByVal R1 As Integer, ByVal C1 As Integer, ByVal R2 As Integer, ByVal C2 As Integer) As Range
        Dim sCell As Range = xlSheet.Cells(R1, C1)
        Dim eCell As Range = xlSheet.Cells(R2, C2)

        Dim cRange As Range = xlSheet.Cells.Range(sCell, eCell)

        Return cRange
    End Function

    ' <summary>
    ' Save As the given file.
    ' </summary>
    ' <param name="fileName"></param>
    Public Sub saveAsFile(ByVal fileName As String)
        xlApp.ActiveWorkbook.SaveCopyAs(fileName)
    End Sub

    Public Sub saveFile()
        xlApp.ActiveWorkbook.Save()
    End Sub

    Public Sub showExcel()
        xlApp.Visible = True
    End Sub

    ' <summary>
    ' Return an array with base index 1 e.g. starting elements from [1,1].
    ' </summary>
    ' <returns></returns>
    Public Function readSelectionRange(ByVal sheetName As String) As Object(,)
        setCurrentSheet(sheetName)

        Dim curRange As Range = xlApp.Selection
        Dim cVals(,) As Object = curRange.Value2
        'returns all the data in form of array object within selection range.
        Return cVals
    End Function

    Public Sub writeRange(ByVal d As Object(,), ByVal r As Long, ByVal c As Long)
        If (d Is Nothing Or d.GetLength(0) < 1 Or d.GetLength(1) < 1) Then Return

        Dim nr As Long = d.GetLength(0) - 1
        Dim nc As Long = d.GetLength(1) - 1

        Dim urange As Range = makeRange(r, c, r + nr, c + nc)
        urange.Value = d
    End Sub

    ' <summary>
    ' Read the cell value.
    ' </summary>
    ' <param name="R"></param>
    ' <param name="C"></param>
    ' <param name="sheetName"></param>
    ' <returns></returns>
    Public Function readCellValue(ByVal R As Integer, ByVal C As Integer, ByVal sheetName As String) As Object
        setCurrentSheet(sheetName)

        Dim cCell As Range = xlSheet.Cells(R, C)
        Return cCell.Value2
    End Function

    ' <summary>
    ' Read data range from excel application.
    ' </summary>
    ' <param name="R1"></param>
    ' <param name="C1"></param>
    ' <param name="R2"></param>
    ' <param name="C2"></param>
    ' <param name="sheetName"></param>
    ' <returns></returns>
    Public Function readRange(ByVal R1 As Integer, ByVal C1 As Integer, ByVal R2 As Integer, ByVal C2 As Integer, ByVal sheetName As String) As Object(,)
        setCurrentSheet(sheetName)

        Dim cCell As Range = makeRange(R1, C1, R2, C2)
        cCell.Select()

        Return readSelectionRange(sheetName)
    End Function

    ''' <summary>
    ''' Destination cell address(r,c)
    ''' Data Source Cell address(r1,c1)
    ''' </summary>
    ''' <param name="R"></param>
    ''' <param name="C"></param>
    ''' <param name="xlsSheet"></param>
    ''' <param name="R1"></param>
    ''' <param name="C1"></param>
    ''' <remarks></remarks>
    Public Sub writeFormula(ByVal R As Integer, ByVal C As Integer, _
                ByVal xlsSheet As String, ByVal R1 As Integer, ByVal C1 As Integer)
        Dim cCell As Range = xlSheet.Cells(R, C)
        Dim xCell As Range = xlSheet.Cells(R1, C1)

        cCell.Formula = "=" & "'" & xlsSheet & "'" & "!" & xCell.Address
    End Sub

    ''' <summary>
    ''' Destination cell address(r,c)
    ''' Data Source Cell address(r1,c1)
    ''' </summary>
    ''' <param name="R"></param>
    ''' <param name="C"></param>
    ''' <param name="fac"></param>
    ''' <param name="R1"></param>
    ''' <param name="C1"></param>
    ''' <remarks></remarks>
    Public Sub writeFormula(ByVal R As Integer, ByVal C As Integer, _
                ByVal fac As Double, ByVal R1 As Integer, ByVal C1 As Integer)
        Dim cCell As Range = xlSheet.Cells(R, C)
        Dim xCell As Range = xlSheet.Cells(R1, C1)

        cCell.Formula = "=" & fac & " * " & xCell.Address
    End Sub

    ''' <summary>
    ''' Destination cell address(r,c)
    ''' Data Source first Cell address(r1,c1) for the range
    ''' Data Source last Cell address(r2,c2) for the range
    ''' function name --> func.
    ''' </summary>
    ''' <param name="R"></param>
    ''' <param name="C"></param>
    ''' <param name="func"></param>
    ''' <param name="R1"></param>
    ''' <param name="C1"></param>
    ''' <param name="R2"></param>
    ''' <param name="C2"></param>
    ''' <remarks></remarks>
    Public Sub writeFormula(ByVal R As Integer, ByVal C As Integer, _
                ByVal func As String, ByVal R1 As Integer, ByVal C1 As Integer, ByVal R2 As Integer, ByVal C2 As Integer)
        Dim cCell As Range = xlSheet.Cells(R, C)
        Dim xCell1 As Range = xlSheet.Cells(R1, C1)
        Dim xCell2 As Range = xlSheet.Cells(R2, C2)

        cCell.Formula = "=" & func & "(" & xCell1.Address & ":" & xCell2.Address & ")"
    End Sub

    ''' <summary>
    ''' Destination cell address(r,c)
    ''' Data Source first Cell address(r1,c1) for the range
    ''' Data Source last Cell address(r2,c2) for the range
    ''' function name --> +,-,*,/.
    ''' </summary>
    ''' <param name="R"></param>
    ''' <param name="C"></param>
    ''' <param name="func"></param>
    ''' <param name="R1"></param>
    ''' <param name="C1"></param>
    ''' <param name="R2"></param>
    ''' <param name="C2"></param>
    ''' <remarks></remarks>
    Public Sub writeFormula(ByVal R As Integer, ByVal C As Integer, _
                ByVal func As Char, ByVal R1 As Integer, ByVal C1 As Integer, ByVal R2 As Integer, ByVal C2 As Integer)
        Dim cCell As Range = xlSheet.Cells(R, C)
        Dim xCell1 As Range = xlSheet.Cells(R1, C1)
        Dim xCell2 As Range = xlSheet.Cells(R2, C2)

        cCell.Formula = "=" & "(" & xCell1.Address & func & xCell2.Address & ")"
    End Sub

    ''' <summary>
    ''' Copy range of cell from same worksheet.
    ''' </summary>
    ''' <param name="rangeToCopy"></param>
    ''' <param name="R1"></param>
    ''' <param name="C1"></param>
    ''' <param name="R2"></param>
    ''' <param name="C2"></param>
    ''' <remarks></remarks>
    Public Sub copyCellRange(ByVal rangeToCopy As String, ByVal R1 As Integer, ByVal C1 As Long, ByVal R2 As Integer, ByVal C2 As Long)
        Dim cCell As Range = xlSheet.Cells(R1, C1)
        Dim cCell1 As Range = xlSheet.Cells(R2, C2)

        Try
            xlSheet.Range(rangeToCopy).Copy(Destination:=xlSheet.Range(cCell, cCell1))
        Catch ex As Exception
            'Do nothing.
        End Try
    End Sub

    ''' <summary>
    ''' Copy range of cell within an active workbook.
    ''' </summary>
    ''' <param name="rangeToCopy"></param>
    ''' <param name="R1"></param>
    ''' <param name="C1"></param>
    ''' <remarks></remarks>
    Public Sub copyCellRange(ByVal rangeToCopy As String, ByVal sheetNa As String, ByVal R1 As Integer, ByVal C1 As Long, ByVal R2 As Integer, ByVal C2 As Long)
        Dim cCell As Range = xlSheet.Cells(R1, C1)
        Dim cCell1 As Range = xlSheet.Cells(R2, C2)

        Try
            xlApp.Worksheets.Item(sheetNa).Range(rangeToCopy).Copy(Destination:=xlSheet.Range(cCell, cCell1))
        Catch ex As Exception
            'Do nothing.
        End Try
    End Sub

    Public Sub clearCellRange(ByVal R1 As Integer, ByVal C1 As Long, ByVal R2 As Integer, ByVal C2 As Long)
        Dim cCell As Range = xlSheet.Cells(R1, C1)
        Dim cCell1 As Range = xlSheet.Cells(R2, C2)

        Dim xlRange As Range = xlSheet.Range(cCell, cCell1)
        xlRange.Clear()
    End Sub
End Class