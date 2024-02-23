#Region "Imports"
Imports System
Imports System.Collections.Generic
Imports YahooAccessData.ExtensionService
Imports System.IO
#End Region



<Serializable()>
Partial Public Class FinancialHighlight
	Implements IEquatable(Of FinancialHighlight)
	Implements IRegisterKey(Of Date)
	Implements IComparable(Of FinancialHighlight)
	Implements IMemoryStream
	Implements IFormatData
	Implements IDateUpdate

	Private MyException As Exception
	Private Shared MyListHeaderInfo As List(Of HeaderInfo)

#Region "New"
	Public Sub New(ByVal DateUpdate As Date)
		With Me
			.DateUpdate = DateUpdate
		End With
		If MyListHeaderInfo Is Nothing And LIST_OF_HEADER_FILE_ENABLED Then
			Dim ThisFile = My.Application.Info.DirectoryPath & "\HeaderInfo\" & TypeName(Me) & ".HeaderInfo.json"
			MyListHeaderInfo = FileHeaderRead(ThisFile, ListOfHeader, Me.Exception)
		Else
			MyListHeaderInfo = ListOfHeader()
		End If
	End Sub

	Public Sub New()
    ReportDate.DateNullValue()
	End Sub

	Public Sub New(ByRef Parent As Record, ByRef Stream As Stream)
		Me.New()
		With Me
			.SerializeLoadFrom(Stream)
			.Record = Parent
			.RecordID = .Record.ID
			.Record.FinancialHighlights.Add(Me)
		End With
	End Sub
#End Region	 'New
#Region "Main"
	Friend Function CopyDeep(ByRef Parent As Record, Optional ByVal IsIgnoreID As Boolean = False) As FinancialHighlight
		Dim ThisFinancialHighlight = New FinancialHighlight

		With ThisFinancialHighlight
			If IsIgnoreID = False Then .ID = Me.ID
			.BookValuePerShare = Me.BookValuePerShare
			.CurrentRatio = Me.CurrentRatio
			.DilutedEPS = Me.DilutedEPS
			.EBITDAInMillion = Me.EBITDAInMillion
			.FiscalYearEnds = Me.FiscalYearEnds
			.GrossProfitInMillion = Me.GrossProfitInMillion
			.LeveredFreeCashFlowInMillion = Me.LeveredFreeCashFlowInMillion
			.MostRecentQuarter = Me.MostRecentQuarter
			.NetIncomeAvlToCommonInMillion = Me.NetIncomeAvlToCommonInMillion
			.OperatingCashFlowInMillion = Me.OperatingCashFlowInMillion
			.OperatingMarginPercent = Me.OperatingMarginPercent
			.ProfitMarginPercent = Me.ProfitMarginPercent
			.QuarterlyRevenueGrowthPercent = Me.QuarterlyRevenueGrowthPercent
			.QuaterlyEarningsGrowthPercent = Me.QuaterlyEarningsGrowthPercent
			.ReturnOnAssetsPercent = Me.ReturnOnAssetsPercent
			.ReturnOnEquityPercent = Me.ReturnOnEquityPercent
			.RevenueInMillion = Me.RevenueInMillion
			.RevenuePerShare = Me.RevenuePerShare
			.TotalCashInMillion = Me.TotalCashInMillion
			.TotalCashPerShare = Me.TotalCashPerShare
			.TotalDeptInMillion = Me.TotalDeptInMillion
			.TotalDeptPerEquity = Me.TotalDeptPerEquity
			.DateUpdate = Me.DateUpdate
			.Record = Parent
			.RecordID = .Record.ID
			.Record.FinancialHighlights.Add(ThisFinancialHighlight)
		End With
		Return ThisFinancialHighlight
	End Function

	''' <summary>
	''' If an exception occurs, the exception object will be stored here. If no exception occurs, this property is null/Nothing.
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Property Exception() As Exception
		Get
				Return MyException
		End Get
		Set(Exception As Exception)
			MyException = Exception
		End Set
	End Property

	Public Overrides Function ToString() As String
		Return String.Format("Object:{0}:{1}:{2}", TypeName(Me), Me.KeyID, Me.KeyValue)
	End Function
#End Region	'Main
#Region "IComparable"
''' <summary>
''' </summary>
''' <param name="other"></param>
''' <returns>
''' Less than zero: This object is less than the other parameter. 
''' Zero : This object is equal to other. 
''' Greater than zero : This object is greater than other. 
''' </returns>
''' <remarks></remarks>
	Public Function CompareTo(other As FinancialHighlight) As Integer Implements System.IComparable(Of FinancialHighlight).CompareTo
		Return Me.KeyValue.CompareTo(other.KeyValue)
	End Function
#End Region
#Region "Register Key"
	Public Property KeyID As Integer Implements IRegisterKey(Of Date).KeyID
		Get
			Return Me.ID
		End Get
		Set(value As Integer)
			Me.ID = value
		End Set
	End Property

	Public Property KeyValue As Date Implements IRegisterKey(Of Date).KeyValue
		Get
			Return Me.DateUpdate
		End Get
		Set(value As Date)

		End Set
	End Property
#End Region
#Region "IMemoryStream"
	Public Sub SerializeSaveTo(ByRef Stream As System.IO.Stream) Implements IMemoryStream.SerializeSaveTo
		Me.SerializeSaveTo(Stream, IMemoryStream.enuFileType.Standard)
	End Sub

	Public Sub SerializeSaveTo(ByRef Stream As System.IO.Stream, FileType As IMemoryStream.enuFileType) Implements IMemoryStream.SerializeSaveTo
		Dim ThisBinaryWriter As New BinaryWriter(Stream)
		Dim ThisListException As List(Of Exception)
		Dim ThisException As Exception

		With ThisBinaryWriter
			.Write(VERSION_MEMORY_STREAM)
			.Write(Me.ID)
			ThisListException = Me.Exception.ToList
			ThisListException.Reverse()
			.Write(ThisListException.Count)
			For Each ThisException In ThisListException
				.Write(ThisException.Message)
			Next
			.Write(Me.BookValuePerShare)
			.Write(Me.CurrentRatio)
			.Write(Me.DilutedEPS)
			.Write(Me.EBITDAInMillion)
			.Write(Me.FiscalYearEnds.ToBinary)
			.Write(Me.GrossProfitInMillion)
			.Write(Me.LeveredFreeCashFlowInMillion)
			.Write(Me.MostRecentQuarter.ToBinary)
			.Write(Me.NetIncomeAvlToCommonInMillion)
			.Write(Me.OperatingCashFlowInMillion)
			.Write(Me.OperatingMarginPercent)
			.Write(Me.ProfitMarginPercent)
			.Write(Me.QuarterlyRevenueGrowthPercent)
			.Write(Me.QuaterlyEarningsGrowthPercent)
			.Write(Me.ReturnOnAssetsPercent)
			.Write(Me.ReturnOnEquityPercent)
			.Write(Me.RevenueInMillion)
			.Write(Me.RevenuePerShare)
			.Write(Me.TotalCashInMillion)
			.Write(Me.TotalCashPerShare)
			.Write(Me.TotalDeptInMillion)
			.Write(Me.TotalDeptPerEquity)
			.Write(Me.RecordID)
			.Write(Me.DateUpdate.ToBinary)
		End With
	End Sub

	Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream, IsRecordVirtual As Boolean) Implements IMemoryStream.SerializeLoadFrom
    Dim ThisBinaryReader As New BinaryReader(Stream, New System.Text.UTF8Encoding(), leaveOpen:=True)
		Dim ThisVersion As Single

		With ThisBinaryReader
			ThisVersion = .ReadSingle
			Me.ID = .ReadInt32
			Me.Exception = Nothing
			Dim I As Integer
			For I = 1 To .ReadInt32
				MyException = New Exception(.ReadString, MyException)
			Next
			Me.BookValuePerShare = .ReadSingle
			Me.CurrentRatio = .ReadSingle
			Me.DilutedEPS = .ReadSingle
			Me.EBITDAInMillion = .ReadSingle
			Me.FiscalYearEnds = DateTime.FromBinary(.ReadInt64)
			Me.GrossProfitInMillion = .ReadSingle
			Me.LeveredFreeCashFlowInMillion = .ReadSingle
			Me.MostRecentQuarter = DateTime.FromBinary(.ReadInt64)
			Me.NetIncomeAvlToCommonInMillion = .ReadSingle
			Me.OperatingCashFlowInMillion = .ReadSingle
			Me.OperatingMarginPercent = .ReadSingle
			Me.ProfitMarginPercent = .ReadSingle
			Me.QuarterlyRevenueGrowthPercent = .ReadSingle
			Me.QuaterlyEarningsGrowthPercent = .ReadSingle
			Me.ReturnOnAssetsPercent = .ReadSingle
			Me.ReturnOnEquityPercent = .ReadSingle
			Me.RevenueInMillion = .ReadSingle
			Me.RevenuePerShare = .ReadSingle
			Me.TotalCashInMillion = .ReadSingle
			Me.TotalCashPerShare = .ReadSingle
			Me.TotalDeptInMillion = .ReadSingle
			Me.TotalDeptPerEquity = .ReadSingle
			Me.RecordID = .ReadInt32
			Me.DateUpdate = DateTime.FromBinary(.ReadInt64)
    End With
    ThisBinaryReader.Dispose()
	End Sub

	Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream) Implements IMemoryStream.SerializeLoadFrom
		Call SerializeLoadFrom(Stream, False)
	End Sub

	Public Sub SerializeLoadFrom(ByRef Data() As Byte) Implements IMemoryStream.SerializeLoadFrom
		Dim ThisStream As Stream = New System.IO.MemoryStream(Data, writable:=True)
		Me.SerializeLoadFrom(ThisStream)
	End Sub

	Public Function SerializeSaveTo() As Byte() Implements IMemoryStream.SerializeSaveTo
		Dim ThisStream As Stream = New System.IO.MemoryStream
		Dim ThisBinaryReader As New BinaryReader(ThisStream)
		Me.SerializeSaveTo(ThisStream)
		Return ThisBinaryReader.ReadBytes(CInt(ThisStream.Length))
	End Function
#End Region
#Region "IFormatData"
	Public Function ToStingOfData() As String() Implements IFormatData.ToStingOfData
		Return Extensions.ToStingOfData(Of FinancialHighlight)(Me)
	End Function

	Public Function ToListOfHeader() As List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
		Return MyListHeaderInfo
	End Function

  Public Shared Function ListOfHeader() As List(Of HeaderInfo)
    Dim ThisListHeaderInfo As New List(Of HeaderInfo)
    With ThisListHeaderInfo
      .Add(New HeaderInfo With {.Name = "ID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateUpdate", .Title = .Name, .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "BookValuePerShare", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "CurrentRatio", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DilutedEPS", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "EBITDAInMillion", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "FiscalYearEnds", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "GrossProfitInMillion", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "LeveredFreeCashFlowInMillion", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "MostRecentQuarter", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "NetIncomeAvlToCommonInMillion", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "OperatingCashFlowInMillion", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "OperatingMarginPercent", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "ProfitMarginPercent", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "QuarterlyRevenueGrowthPercent", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "QuaterlyEarningsGrowthPercent", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "ReturnOnAssetsPercent", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "ReturnOnEquityPercent", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "RevenueInMillion", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "RevenuePerShare", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "TotalCashInMillion", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "TotalCashPerShare", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "TotalDeptInMillion", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "TotalDeptPerEquity", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
    End With
    Return ThisListHeaderInfo
  End Function
#End Region
#Region "IEquatable"
	Public Function EqualsDeep(other As FinancialHighlight, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
		With other
			If IsIgnoreID = False Then
				If .ID <> Me.ID Then Return False
				If .RecordID <> Me.RecordID Then Return False
			End If
			If .BookValuePerShare <> Me.BookValuePerShare Then Return False
			If .CurrentRatio <> Me.CurrentRatio Then Return False
			If .DilutedEPS <> Me.DilutedEPS Then Return False
			If .EBITDAInMillion <> Me.EBITDAInMillion Then Return False
			If .FiscalYearEnds <> Me.FiscalYearEnds Then Return False
			If .GrossProfitInMillion <> Me.GrossProfitInMillion Then Return False
			If .LeveredFreeCashFlowInMillion <> Me.LeveredFreeCashFlowInMillion Then Return False
			If .MostRecentQuarter <> Me.MostRecentQuarter Then Return False
			If .NetIncomeAvlToCommonInMillion <> Me.NetIncomeAvlToCommonInMillion Then Return False
			If .OperatingCashFlowInMillion <> Me.OperatingCashFlowInMillion Then Return False
			If .OperatingMarginPercent <> Me.OperatingMarginPercent Then Return False
			If .ProfitMarginPercent <> Me.ProfitMarginPercent Then Return False
			If .QuarterlyRevenueGrowthPercent <> Me.QuarterlyRevenueGrowthPercent Then Return False
			If .QuaterlyEarningsGrowthPercent <> Me.QuaterlyEarningsGrowthPercent Then Return False
			If .ReturnOnAssetsPercent <> Me.ReturnOnAssetsPercent Then Return False
			If .ReturnOnEquityPercent <> Me.ReturnOnEquityPercent Then Return False
			If .RevenueInMillion <> Me.RevenueInMillion Then Return False
			If .RevenuePerShare <> Me.RevenuePerShare Then Return False
			If .TotalCashInMillion <> Me.TotalCashInMillion Then Return False
			If .TotalCashPerShare <> Me.TotalCashPerShare Then Return False
			If .TotalDeptInMillion <> Me.TotalDeptInMillion Then Return False
			If .TotalDeptPerEquity <> Me.TotalDeptPerEquity Then Return False
			If .DateUpdate <> Me.DateUpdate Then Return False
		End With
		Return True
	End Function

	Public Overloads Function Equals(other As FinancialHighlight) As Boolean Implements IEquatable(Of FinancialHighlight).Equals
		If other Is Nothing Then Return False
		If Me.DateUpdate = other.DateUpdate Then
			If Me.Record IsNot Nothing Then
				Return Me.Record.Equals(other.Record)
			Else
				Return False
			End If
		Else
			Return False
		End If
	End Function

	Public Overrides Function Equals(obj As Object) As Boolean
		If (TypeOf obj Is FinancialHighlight) Then
			Return Me.Equals(DirectCast(obj, FinancialHighlight))
		Else
			Return (False)
		End If
	End Function

	Public Overrides Function GetHashCode() As Integer
		If Me.Record IsNot Nothing Then
			Return Me.Record.GetHashCode()
		Else
			Return Me.DateUpdate.GetHashCode()
		End If
	End Function
#End Region
#Region "IDateUpdate"
	Private Property IDateUpdate_DateStart As Date Implements IDateUpdate.DateStart
		Get
			Return Me.DateUpdate
		End Get
		Set(value As Date)
		End Set
	End Property

	Private Property IDateUpdate_DateStop As Date Implements IDateUpdate.DateStop
		Get
			Return Me.DateUpdate
		End Get
		Set(value As Date)
		End Set
	End Property

	Private ReadOnly Property IDateUpdate_DateUpdate As Date Implements IDateUpdate.DateUpdate
		Get
			Return Me.DateUpdate
		End Get
  End Property

  Private ReadOnly Property IDateUpdate_DateLastTrade As Date Implements IDateUpdate.DateLastTrade
    Get
      Return IDateUpdate_DateStop
    End Get
  End Property

  Private ReadOnly Property IDateUpdate_DateDay As Date Implements IDateUpdate.DateDay
    Get
      Return IDateUpdate_DateStop.Date
    End Get
  End Property
#End Region
End Class
#Region "EqualityComparerOfFinancialHighlight"
	<Serializable()>
	Friend Class EqualityComparerOfFinancialHighlight
		Implements IEqualityComparer(Of FinancialHighlight)

		Public Overloads Function Equals(x As FinancialHighlight, y As FinancialHighlight) As Boolean Implements IEqualityComparer(Of FinancialHighlight).Equals
			If (x Is Nothing) And (y Is Nothing) Then
				Return True
			ElseIf (x Is Nothing) Xor (y Is Nothing) Then
				Return False
			Else
				If x.DateUpdate = y.DateUpdate Then
					If (x.Record Is Nothing) And (y.Record Is Nothing) Then
						Return True
					ElseIf (x.Record Is Nothing) Xor (y.Record Is Nothing) Then
						Return False
					Else
						If x.Record.Equals(y) Then
							Return True
						Else
							Return False
						End If
					End If
				Else
					Return False
				End If
			End If
		End Function

		Public Overloads Function GetHashCode(obj As FinancialHighlight) As Integer Implements IEqualityComparer(Of FinancialHighlight).GetHashCode
			If obj IsNot Nothing Then
				Return obj.DateUpdate.GetHashCode
			Else
				Return obj.GetHashCode
			End If
		End Function
	End Class
#End Region

'------------------------------------------------------------------------------
' <auto-generated>
'    This code was generated from a template.
'
'    Manual changes to this file may cause unexpected behavior in your application.
'    Manual changes to this file will be overwritten if the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

#Region "FinancialHighlight template"
Partial Public Class FinancialHighlight
  Public Property ID As Integer
  Public Property BookValuePerShare As Single
  Public Property CurrentRatio As Single
  Public Property DilutedEPS As Single
  Public Property EBITDAInMillion As Single
  Public Property FiscalYearEnds As Date
  Public Property GrossProfitInMillion As Single
  Public Property LeveredFreeCashFlowInMillion As Single
  Public Property MostRecentQuarter As Date
  Public Property NetIncomeAvlToCommonInMillion As Single
  Public Property OperatingCashFlowInMillion As Single
  Public Property OperatingMarginPercent As Single
  Public Property ProfitMarginPercent As Single
  Public Property QuarterlyRevenueGrowthPercent As Single
  Public Property QuaterlyEarningsGrowthPercent As Single
  Public Property ReturnOnAssetsPercent As Single
  Public Property ReturnOnEquityPercent As Single
  Public Property RevenueInMillion As Single
  Public Property RevenuePerShare As Single
  Public Property TotalCashInMillion As Single
  Public Property TotalCashPerShare As Single
  Public Property TotalDeptInMillion As Single
  Public Property TotalDeptPerEquity As Single
  Public Property RecordID As Integer
  Public Property DateUpdate As Date

  Public Overridable Property Record As Record

End Class
#End Region