Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Xml.Serialization

<XmlRoot("NewDataSet")> _
Public Class OrderData
	Inherits List(Of Orders)
	Public Sub New()
	End Sub
End Class
Public Class Orders
	Private privateOrderID As Integer
	Public Property OrderID() As Integer
		Get
			Return privateOrderID
		End Get
		Set(ByVal value As Integer)
			privateOrderID = value
		End Set
	End Property
	Private privateCustomerID As String
	Public Property CustomerID() As String
		Get
			Return privateCustomerID
		End Get
		Set(ByVal value As String)
			privateCustomerID = value
		End Set
	End Property
	Private privateEmployeeID As Integer
	Public Property EmployeeID() As Integer
		Get
			Return privateEmployeeID
		End Get
		Set(ByVal value As Integer)
			privateEmployeeID = value
		End Set
	End Property
	Private privateOrderDate As DateTime
	Public Property OrderDate() As DateTime
		Get
			Return privateOrderDate
		End Get
		Set(ByVal value As DateTime)
			privateOrderDate = value
		End Set
	End Property
	Private privateRequiredDate As DateTime
	Public Property RequiredDate() As DateTime
		Get
			Return privateRequiredDate
		End Get
		Set(ByVal value As DateTime)
			privateRequiredDate = value
		End Set
	End Property
	Private privateShippedDate As DateTime
	Public Property ShippedDate() As DateTime
		Get
			Return privateShippedDate
		End Get
		Set(ByVal value As DateTime)
			privateShippedDate = value
		End Set
	End Property
	Private privateShipVia As Integer
	Public Property ShipVia() As Integer
		Get
			Return privateShipVia
		End Get
		Set(ByVal value As Integer)
			privateShipVia = value
		End Set
	End Property
	Private privateFreight As Decimal
	Public Property Freight() As Decimal
		Get
			Return privateFreight
		End Get
		Set(ByVal value As Decimal)
			privateFreight = value
		End Set
	End Property
	Private privateShipName As String
	Public Property ShipName() As String
		Get
			Return privateShipName
		End Get
		Set(ByVal value As String)
			privateShipName = value
		End Set
	End Property
	Private privateShipAddress As String
	Public Property ShipAddress() As String
		Get
			Return privateShipAddress
		End Get
		Set(ByVal value As String)
			privateShipAddress = value
		End Set
	End Property
	Private privateShipCity As String
	Public Property ShipCity() As String
		Get
			Return privateShipCity
		End Get
		Set(ByVal value As String)
			privateShipCity = value
		End Set
	End Property
	Private privateShipRegion As String
	Public Property ShipRegion() As String
		Get
			Return privateShipRegion
		End Get
		Set(ByVal value As String)
			privateShipRegion = value
		End Set
	End Property
	Private privateShipPostalCode As String
	Public Property ShipPostalCode() As String
		Get
			Return privateShipPostalCode
		End Get
		Set(ByVal value As String)
			privateShipPostalCode = value
		End Set
	End Property
	Private privateShipCountry As String
	Public Property ShipCountry() As String
		Get
			Return privateShipCountry
		End Get
		Set(ByVal value As String)
			privateShipCountry = value
		End Set
	End Property
End Class

