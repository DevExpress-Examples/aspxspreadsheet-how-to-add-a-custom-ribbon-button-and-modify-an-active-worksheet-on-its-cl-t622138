Imports DevExpress.Spreadsheet
Imports DevExpress.Web.ASPxSpreadsheet
Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Linq
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls

Partial Public Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
        If Not IsPostBack Then
            ASPxSpreadsheet1.Open(Server.MapPath("~/WorkDirectory/Book1.xlsx"))
        End If
    End Sub

    Protected Sub ASPxSpreadsheet1_Callback(ByVal sender As Object, ByVal e As DevExpress.Web.CallbackEventArgsBase)
        Dim ss As ASPxSpreadsheet = TryCast(sender, ASPxSpreadsheet)
        Dim worksheet As Worksheet = ss.Document.Worksheets.ActiveWorksheet
        If e.Parameter = "custom" Then
            PrepareTitleRange(worksheet)
            PrepareHeaderCells(worksheet)
            InitializeDataCellsValues(worksheet)
        End If
        If e.Parameter = "clear" Then
            worksheet.Clear(worksheet.GetUsedRange())
        End If
    End Sub

    Private Sub PrepareTitleRange(ByVal worksheet As Worksheet)
        worksheet.Cells("B1").FillColor = Color.LightBlue
        worksheet.Cells("B1").Value = "Cell value types"
        Dim range As CellRange = worksheet.Range("A1:B1")
        range.Style = worksheet.Workbook.Styles("Title")
        range.Merge()
    End Sub
    Private Sub PrepareHeaderCells(ByVal worksheet As Worksheet)

        Dim header_Renamed As CellRange = worksheet.Range("A2:B2")
        header_Renamed(0).Value = "Type"
        header_Renamed(1).Value = "Value"
        header_Renamed.ColumnWidthInCharacters = 25
        header_Renamed.Style = worksheet.Workbook.Styles("Heading 2")
    End Sub
    Private Sub InitializeDataCellsValues(ByVal worksheet As Worksheet)
        ' Add data of different types to cells.
        worksheet.Cells("B3").Value = Date.Now
        worksheet.Cells("B4").Value = Math.PI
        worksheet.Cells("B5").Value = "Have a nice day!"
        worksheet.Cells("B6").Value = CellValue.ErrorReference
        worksheet.Cells("B7").Value = True
        worksheet.Cells("B8").Value = Single.MaxValue
        worksheet.Cells("B9").Value = "a"c
        worksheet.Cells("B10").Value = Int32.MaxValue

        worksheet.Cells("A3").Value = "dateTime"
        worksheet.Cells("A4").Value = "double"
        worksheet.Cells("A5").Value = "string"
        worksheet.Cells("A6").Value = "error constant"
        worksheet.Cells("A7").Value = "boolean"
        worksheet.Cells("A8").Value = "float"
        worksheet.Cells("A9").Value = "char"
        worksheet.Cells("A10").Value = "int32"
        worksheet.Cells("A13").Value = "fill range"

        ' Fill all cells of the range with 10.
        worksheet.Range("B13:C13").Value = 10
    End Sub
End Class