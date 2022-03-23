#Region "Imports"
Imports System
Imports System.Collections.Generic
Imports YahooAccessData.ExtensionService
Imports System.IO
#End Region

<Serializable()>
Partial Public Class TradeInfo
	Implements IEquatable(Of TradeInfo)
	Implements IRegisterKey(Of Date)
	Implements IComparable(Of TradeInfo)
	Implements IMemoryStream
	Implements IFormatData
	Implements IDateUpdate

	Private MyException As Exception
	Private Shared MyListHeaderInfo As List(Of HeaderInfo)

	Public Sub New(ByVal DateUpdate As Date)
		With Me
			.DateUpdate = DateUpdate
			.DividendDate = DateSerial(1900, 1, 1)
			.ExDividendDate = DateSerial(1900, 1, 1)
			.LastSplitDate = DateSerial(1900, 1, 1)
		End With
		If MyListHeaderInfo Is Nothing Then
			Dim ThisFile = My.Application.Info.DirectoryPath & "\HeaderInfo\" & TypeName(Me) & ".HeaderInfo.xml"
      MyListHeaderInfo = FileHeaderRead(ThisFile, ListOfHeader, Me.Exception)
		End If
	End Sub

	Public Sub New()
    Me.New(Now)
	End Sub

	Public Sub New(ByRef Parent As Record, ByRef Stream As Stream)
		Me.New()
		With Me
			.SerializeLoadFrom(Stream)
			.Record = Parent
			.RecordID = .Record.ID
			.Record.TradeInfoes.Add(Me)
		End With
	End Sub

	Friend Function CopyDeep(ByRef Parent As Record, Optional ByVal IsIgnoreID As Boolean = False) As TradeInfo
		Dim ThisTradeInfo = New TradeInfo

		With ThisTradeInfo
			If IsIgnoreID = False Then .ID = Me.ID
			.AverageVolumeTenDaysInThousand = Me.AverageVolumeTenDaysInThousand
			.AverageVolumeThreeMonthInThousand = Me.AverageVolumeThreeMonthInThousand
			.Beta = Me.Beta
			.DividendDate = Me.DividendDate
			.ExDividendDate = Me.ExDividendDate
			.FiftyDayMovingAverage = Me.FiftyDayMovingAverage
			.FiveYearAverageDividendYieldPercent = Me.FiveYearAverageDividendYieldPercent
			.FloatInMillion = Me.FloatInMillion
			.ForwardAnnualDividendRate = Me.ForwardAnnualDividendRate
			.ForwardAnnualDividendYieldPercent = Me.ForwardAnnualDividendYieldPercent
			.LastSplitDate = Me.LastSplitDate
			.LastSplitFactor = Me.LastSplitFactor
			.OneYearChangePercent = Me.OneYearChangePercent
			.OneYearHigh = Me.OneYearHigh
			.OneYearLow = Me.OneYearLow
			.PayoutRatio = Me.PayoutRatio
			.PercentHeldByInsiders = Me.PercentHeldByInsiders
			.PercentHeldByInstitutions = Me.PercentHeldByInstitutions
			.SharesOutstandingInMillion = Me.SharesOutstandingInMillion
			.SharesShortInMillion = Me.SharesShortInMillion
			.SharesShortPriorMonthInMillion = Me.SharesShortPriorMonthInMillion
			.ShortPercentOfFloat = Me.ShortPercentOfFloat
			.ShortRatio = Me.ShortRatio
			.TrailingAnnualDividendYield = Me.TrailingAnnualDividendYield
			.TrailingAnnualDividendYieldPercent = Me.TrailingAnnualDividendYieldPercent
			.TwoHundredDayMovingAverage = Me.TwoHundredDayMovingAverage
			.RecordID = Me.RecordID
			.DateUpdate = Me.DateUpdate
			.Record = Parent
			.Record.TradeInfoes.Add(ThisTradeInfo)
		End With
		Return ThisTradeInfo
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
	Public Function CompareTo(other As TradeInfo) As Integer Implements System.IComparable(Of TradeInfo).CompareTo
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
			.Write(Me.AverageVolumeTenDaysInThousand)
			.Write(Me.AverageVolumeThreeMonthInThousand)
			.Write(Me.Beta)
			.Write(Me.DividendDate.ToBinary)
			.Write(Me.ExDividendDate.ToBinary)
			.Write(Me.FiftyDayMovingAverage)
			.Write(Me.FiveYearAverageDividendYieldPercent)
			.Write(Me.FloatInMillion)
			.Write(Me.ForwardAnnualDividendRate)
			.Write(Me.ForwardAnnualDividendYieldPercent)
			.Write(Me.LastSplitDate.ToBinary)
			.Write(Me.LastSplitFactor.ToBinary)
			.Write(Me.OneYearChangePercent)
			.Write(Me.OneYearHigh)
			.Write(Me.OneYearLow)
			.Write(Me.PayoutRatio)
			.Write(Me.PercentHeldByInsiders)
			.Write(Me.PercentHeldByInstitutions)
			.Write(Me.SharesOutstandingInMillion)
			.Write(Me.SharesShortInMillion)
			.Write(Me.SharesShortPriorMonthInMillion)
			.Write(Me.ShortPercentOfFloat)
			.Write(Me.ShortRatio)
			.Write(Me.TrailingAnnualDividendYield)
			.Write(Me.TrailingAnnualDividendYieldPercent)
			.Write(Me.TwoHundredDayMovingAverage)
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
			Me.AverageVolumeTenDaysInThousand = .ReadSingle
			Me.AverageVolumeThreeMonthInThousand = .ReadSingle
			Me.Beta = .ReadSingle
			Me.DividendDate = DateTime.FromBinary(.ReadInt64)
			Me.ExDividendDate = DateTime.FromBinary(.ReadInt64)
			Me.FiftyDayMovingAverage = .ReadSingle
			Me.FiveYearAverageDividendYieldPercent = .ReadSingle
			Me.FloatInMillion = .ReadSingle
			Me.ForwardAnnualDividendRate = .ReadSingle
			Me.ForwardAnnualDividendYieldPercent = .ReadSingle
			Me.LastSplitDate = DateTime.FromBinary(.ReadInt64)
			Me.LastSplitFactor = DateTime.FromBinary(.ReadInt64)
			Me.OneYearChangePercent = .ReadSingle
			Me.OneYearHigh = .ReadSingle
			Me.OneYearLow = .ReadSingle
			Me.PayoutRatio = .ReadSingle
			Me.PercentHeldByInsiders = .ReadSingle
			Me.PercentHeldByInstitutions = .ReadSingle
			Me.SharesOutstandingInMillion = .ReadSingle
			Me.SharesShortInMillion = .ReadSingle
			Me.SharesShortPriorMonthInMillion = .ReadSingle
			Me.ShortPercentOfFloat = .ReadSingle
			Me.ShortRatio = .ReadSingle
			Me.TrailingAnnualDividendYield = .ReadSingle
			Me.TrailingAnnualDividendYieldPercent = .ReadSingle
			Me.TwoHundredDayMovingAverage = .ReadSingle
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
		Return Extensions.ToStingOfData(Of TradeInfo)(Me)
	End Function

	Public Function ToListOfHeader() As List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
		Return MyListHeaderInfo
	End Function

  Public Shared Function ListOfHeader() As List(Of HeaderInfo)
    Dim ThisListHeaderInfo As New List(Of HeaderInfo)
    With ThisListHeaderInfo
      .Add(New HeaderInfo With {.Name = "ID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateUpdate", .Title = .Name, .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "AverageVolumeTenDaysInThousand", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "AverageVolumeThreeMonthInThousand", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Beta", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DividendDate", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "ExDividendDate", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "FiftyDayMovingAverage", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "FiveYearAverageDividendYieldPercent", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "FloatInMillion", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "ForwardAnnualDividendRate", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "ForwardAnnualDividendYieldPercent", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "LastSplitDate", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "LastSplitFactor", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "OneYearChangePercent", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "OneYearHigh", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "OneYearLow", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "PayoutRatio", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "PercentHeldByInsiders", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "PercentHeldByInstitutions", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "SharesOutstandingInMillion", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "SharesShortInMillion", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "SharesShortPriorMonthInMillion", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "ShortPercentOfFloat", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "ShortRatio", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "TrailingAnnualDividendYield", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "TrailingAnnualDividendYieldPercent", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "TwoHundredDayMovingAverage", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
    End With
    Return ThisListHeaderInfo
  End Function
#End Region
#Region "IEquatable"
	Public Function EqualsDeep(other As TradeInfo, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
		With other
			If IsIgnoreID = False Then
				If .ID <> Me.ID Then Return False
				If .RecordID <> Me.RecordID Then Return False
			End If
			If .AverageVolumeTenDaysInThousand <> Me.AverageVolumeTenDaysInThousand Then Return False
			If .AverageVolumeThreeMonthInThousand <> Me.AverageVolumeThreeMonthInThousand Then Return False
			If .Beta <> Me.Beta Then Return False
			If .DividendDate <> Me.DividendDate Then Return False
			If .ExDividendDate <> Me.ExDividendDate Then Return False
			If .FiftyDayMovingAverage <> Me.FiftyDayMovingAverage Then Return False
			If .FiveYearAverageDividendYieldPercent <> Me.FiveYearAverageDividendYieldPercent Then Return False
			If .FloatInMillion <> Me.FloatInMillion Then Return False
			If .ForwardAnnualDividendRate <> Me.ForwardAnnualDividendRate Then Return False
			If .ForwardAnnualDividendYieldPercent <> Me.ForwardAnnualDividendYieldPercent Then Return False
			If .LastSplitDate <> Me.LastSplitDate Then Return False
			If .LastSplitFactor <> Me.LastSplitFactor Then Return False
			If .OneYearChangePercent <> Me.OneYearChangePercent Then Return False
			If .OneYearHigh <> Me.OneYearHigh Then Return False
			If .OneYearLow <> Me.OneYearLow Then Return False
			If .PayoutRatio <> Me.PayoutRatio Then Return False
			If .PercentHeldByInsiders <> Me.PercentHeldByInsiders Then Return False
			If .PercentHeldByInstitutions <> Me.PercentHeldByInstitutions Then Return False
			If .SharesOutstandingInMillion <> Me.SharesOutstandingInMillion Then Return False
			If .SharesShortInMillion <> Me.SharesShortInMillion Then Return False
			If .SharesShortPriorMonthInMillion <> Me.SharesShortPriorMonthInMillion Then Return False
			If .ShortPercentOfFloat <> Me.ShortPercentOfFloat Then Return False
			If .ShortRatio <> Me.ShortRatio Then Return False
			If .TrailingAnnualDividendYield <> Me.TrailingAnnualDividendYield Then Return False
			If .TrailingAnnualDividendYieldPercent <> Me.TrailingAnnualDividendYieldPercent Then Return False
			If .TwoHundredDayMovingAverage <> Me.TwoHundredDayMovingAverage Then Return False
			If .DateUpdate <> Me.DateUpdate Then Return False
		End With
		Return True
	End Function

	Public Overloads Function Equals(other As TradeInfo) As Boolean Implements IEquatable(Of TradeInfo).Equals
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
		If (TypeOf obj Is TradeInfo) Then
			Return Me.Equals(DirectCast(obj, TradeInfo))
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
#End Region	 'IDateUpdate
End Class

#Region "EqualityComparerOfTradeInfo"
	<Serializable()>
	Friend Class EqualityComparerOfTradeInfo
		Implements IEqualityComparer(Of TradeInfo)

		Public Overloads Function Equals(x As TradeInfo, y As TradeInfo) As Boolean Implements IEqualityComparer(Of TradeInfo).Equals
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

		Public Overloads Function GetHashCode(obj As TradeInfo) As Integer Implements IEqualityComparer(Of TradeInfo).GetHashCode
			If obj IsNot Nothing Then
				Return obj.DateUpdate.GetHashCode
			Else
				Return obj.GetHashCode
			End If
		End Function
	End Class
#End Region

#Region "Template"
'------------------------------------------------------------------------------
' <auto-generated>
'    This code was generated from a template.
'
'    Manual changes to this file may cause unexpected behavior in your application.
'    Manual changes to this file will be overwritten if the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Partial Public Class TradeInfo
  Public Property ID As Integer
  Public Property AverageVolumeTenDaysInThousand As Single
  Public Property AverageVolumeThreeMonthInThousand As Single
  Public Property Beta As Single
  Public Property DividendDate As Date
  Public Property ExDividendDate As Date
  Public Property FiftyDayMovingAverage As Single
  Public Property FiveYearAverageDividendYieldPercent As Single
  Public Property FloatInMillion As Single
  Public Property ForwardAnnualDividendRate As Single
  Public Property ForwardAnnualDividendYieldPercent As Single
  Public Property LastSplitDate As Date
  Public Property LastSplitFactor As Date
  Public Property OneYearChangePercent As Single
  Public Property OneYearHigh As Single
  Public Property OneYearLow As Single
  Public Property PayoutRatio As Single
  Public Property PercentHeldByInsiders As Single
  Public Property PercentHeldByInstitutions As Single
  Public Property SharesOutstandingInMillion As Single
  Public Property SharesShortInMillion As Single
  Public Property SharesShortPriorMonthInMillion As Single
  Public Property ShortPercentOfFloat As Single
  Public Property ShortRatio As Single
  Public Property TrailingAnnualDividendYield As Single
  Public Property TrailingAnnualDividendYieldPercent As Single
  Public Property TwoHundredDayMovingAverage As Single
  Public Property RecordID As Integer
  Public Property DateUpdate As Date

  Public Overridable Property Record As Record
End Class
#End Region