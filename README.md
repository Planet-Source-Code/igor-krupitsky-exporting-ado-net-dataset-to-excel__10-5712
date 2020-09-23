<div align="center">

## Exporting ADO\.NET DataSet To Excel


</div>

### Description

Each table within the ADO.NET DataSet will be exported as a worksheet. All column headers are frozen.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Igor Krupitsky](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/igor-krupitsky.md)
**Level**          |Intermediate
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB\.NET, ASP\.NET
**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__10-9.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/igor-krupitsky-exporting-ado-net-dataset-to-excel__10-5712/archive/master.zip)





### Source Code

```
Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
	Dim sConn As String = "Password=test;User ID=test;Initial Catalog=Northwind;Data Source=(local);"
	Dim cn As SqlConnection = New SqlConnection(sConn)
	cn.Open()
	Dim ds As DataSet = New DataSet("Order")
	Dim da1 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Employees", cn)
	da1.Fill(ds, "Employees")
	Dim da2 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Products", cn)
	da2.Fill(ds, "Products")
	cn.Close()
	ExportToExcel(ds)
End Sub
Sub ExportToExcel(ByRef ds As DataSet)
	Dim oTable As DataTable
	Dim oRow As DataRow
	Dim oColumn As DataColumn
	'Header
	Response.ContentType = "application/vnd.ms-excel"
	Response.Write("<?xml version=""1.0"" encoding=""iso-8859-1""?>" & vbCrLf)
	Response.Write("<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet""" & vbCrLf)
	Response.Write("xmlns:o=""urn:schemas-microsoft-com:office:office""" & vbCrLf)
	Response.Write("xmlns:x=""urn:schemas-microsoft-com:office:excel""" & vbCrLf)
	Response.Write("xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet""" & vbCrLf)
	Response.Write("xmlns:html=""http://www.w3.org/TR/REC-html40"">" & vbCrLf)
	'Style
	Response.Write("<Styles>")
	Response.Write("<Style ss:ID=""s21"">")
	Response.Write("<Font ss:Bold=""1""/>")
	Response.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Bottom""/>")
	Response.Write("</Style>")
	Response.Write("</Styles>")
	For Each oTable In ds.Tables
		'Start Worksheet
		Response.Write("<Worksheet ss:Name=""" & oTable.TableName & """>" & vbCrLf)
		Response.Write("<Table>" & vbCrLf)
		'Column Width
		For Each oColumn In oTable.Columns
			Response.Write("<Column ss:AutoFitWidth=""1"" ss:Width=""150""/>")
		Next
		'Columns
		Response.Write("<Row>" & vbCrLf)
		For Each oColumn In oTable.Columns
			If oColumn.DataType.ToString() <> "System.Byte[]" Then
				Response.Write("<Cell ss:StyleID=""s21"">")
				Response.Write("<Data ss:Type=""String"">")
				Response.Write(oColumn.ColumnName)
				Response.Write("</Data>")
				Response.Write("</Cell>" & vbCrLf)
			End If
		Next
		Response.Write("</Row>" & vbCrLf)
		'Data
		For Each oRow In oTable.Rows
			Response.Write("<Row>")
			For i As Integer = 0 To oTable.Columns.Count - 1
				Dim sType As String = oTable.Columns(i).DataType.ToString()
				If sType <> "System.Byte[]" Then
					Dim sValue As String = oRow(i) & ""
					Response.Write("<Cell>")
					Response.Write("<Data ss:Type=""" & GetExcelDataType(sType) & """>")
					Response.Write("<![CDATA[" & sValue & "]]>")
					Response.Write("</Data>")
					Response.Write("</Cell>" & vbCrLf)
				End If
			Next
			Response.Write("</Row>" & vbCrLf)
		Next
		Response.Write("</Table>" & vbCrLf)
		'Options
		Response.Write("<WorksheetOptions xmlns=""urn:schemas-microsoft-com:office:excel"">" & vbCrLf)
		Response.Write("<FreezePanes/>" & vbCrLf)
		Response.Write("<FrozenNoSplit/>" & vbCrLf)
		Response.Write("<SplitHorizontal>1</SplitHorizontal>" & vbCrLf)
		Response.Write("<TopRowBottomPane>1</TopRowBottomPane>" & vbCrLf)
		Response.Write("<ActivePane>2</ActivePane>" & vbCrLf)
		Response.Write("</WorksheetOptions>" & vbCrLf)
		Response.Write("</Worksheet>" & vbCrLf)
	Next
	Response.Write("</Workbook>" & vbCrLf)
End Sub
Function GetExcelDataType(ByVal sType As String) As String
	Select Case sType
		Case "System.Int32" : Return "Number"
		Case "System.Int16" : Return "Number"
		Case "System.Decimal" : Return "Number"
	End Select
	Return "String"
End Function
```

