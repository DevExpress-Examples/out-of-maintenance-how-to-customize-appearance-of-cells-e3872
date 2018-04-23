Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Xml.Serialization
Imports DevExpress.Xpf.PivotGrid

Namespace DXPivotGrid_CellAppearance
	Partial Public Class MainPage
		Inherits UserControl
		Private dataFileName As String = "nwind.xml"
		Private minValue, maxValue, minTotalValue, maxTotalValue As Decimal
		Private maxMinCalculated As Boolean
		Public Sub New()
			InitializeComponent()

			' Parses an XML file and creates a collection of data items.
			Dim [assembly] As System.Reflection.Assembly = _
				System.Reflection.Assembly.GetExecutingAssembly()
			Dim stream As Stream = [assembly].GetManifestResourceStream(dataFileName)
			Dim s As New XmlSerializer(GetType(OrderData))
			Dim dataSource As Object = s.Deserialize(stream)

			' Binds a pivot grid to this collection.
			pivotGridControl1.DataSource = dataSource
		End Sub
		Private Sub UserControl_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
			pivotGridControl1.CollapseAll()
		End Sub
		Private Sub pivotGridControl1_GridLayout(ByVal sender As Object, ByVal e As RoutedEventArgs)
			ResetMaxMin()
		End Sub
		Private Sub pivotGridControl1_CustomCellAppearance(ByVal sender As Object, _
			ByVal e As PivotCustomCellAppearanceEventArgs)
			If Not(TypeOf e.Value Is Decimal) Then
				Return
			End If
			Dim value As Decimal = CDec(e.Value)
			Dim isGrandTotal As Boolean = Me.IsGrandTotal(e)
			If IsValueNotFit(value, isGrandTotal) Then
				ResetMaxMin()
			End If
			EnsureMaxMin()
			If IsValueNotFit(value, isGrandTotal) Then
				Return
			End If
			e.Foreground = New SolidColorBrush(GetColorByValue(value, Me.IsGrandTotal(e)))
		End Sub
		Private Function IsValueNotFit(ByVal value As Decimal, ByVal isGrandToatl As Boolean) As Boolean
			If isGrandToatl Then
				Return value < minTotalValue OrElse value > maxTotalValue
			Else
				Return value < minValue OrElse value > maxValue
			End If
		End Function

		' Generates a custom color for a cell based on the cell value's share 
		' in the maximum summary or total value (for summary and total cells),
		' or in the maximum Grand Total value (for Grand Total cells).
		Private Function GetColorByValue(ByVal value As Decimal, ByVal isGrandTotal As Boolean) As Color
			Dim variation As Integer
			If isGrandTotal Then
				variation = _
					Convert.ToInt32(510 * (value - minTotalValue) / (maxTotalValue - minTotalValue))
			Else
				variation = _
					Convert.ToInt32(510 * (value - minValue) / (maxValue - minValue))
			End If
			Dim r, b As Byte
			If variation >= 255 Then
				r = Convert.ToByte(510 - variation)
				b = 255
			Else
				r = 255
				b = Convert.ToByte(variation)
			End If
			Return Color.FromArgb(255, r, 0, b)
		End Function
		Private Function IsGrandTotal(ByVal e As PivotCustomSummaryEventArgs) As Boolean
			Return e.RowField Is Nothing OrElse e.ColumnField Is Nothing
		End Function
		Private Function IsGrandTotal(ByVal e As PivotCellBaseEventArgs) As Boolean
			Return e.RowValueType = FieldValueType.GrandTotal OrElse _
				e.ColumnValueType = FieldValueType.GrandTotal
		End Function

		' Calculates the maximum and minimum summary and Grand Total values.
		Private Sub EnsureMaxMin()
			If maxMinCalculated Then
				Return
			End If
			For i As Integer = 0 To pivotGridControl1.RowCount - 1
				For j As Integer = 0 To pivotGridControl1.ColumnCount - 1
					Dim val As Object = pivotGridControl1.GetCellValue(j, i)
					If Not(TypeOf val Is Decimal) Then
						Continue For
					End If
					Dim value As Decimal = CDec(val)
					Dim isGrandTotal As Boolean = _
						pivotGridControl1.GetFieldValueType(True, j) = FieldValueType.GrandTotal OrElse _
						pivotGridControl1.GetFieldValueType(False, i) = FieldValueType.GrandTotal
					If isGrandTotal Then
						If value > maxTotalValue Then
							maxTotalValue = value
						End If
						If value < minTotalValue Then
							minTotalValue = value
						End If
					Else
						If value > maxValue Then
							maxValue = value
						End If
						If value < minValue Then
							minValue = value
						End If
					End If
				Next j
			Next i
			If minTotalValue = maxTotalValue Then
				maxTotalValue += 1
			End If
			If minValue = maxValue Then
				maxValue += 1
			End If
			maxMinCalculated = True
		End Sub

		' Resets the maximum and minimum summary and Grand Total values.
		Private Sub ResetMaxMin()
			minValue = Decimal.MaxValue
			maxValue = Decimal.MinValue
			minTotalValue = Decimal.MaxValue
			maxTotalValue = Decimal.MinValue
			maxMinCalculated = False
		End Sub
	End Class
End Namespace
