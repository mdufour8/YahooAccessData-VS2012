#Region "Imports"
  Imports YahooAccessData.ExtensionService
  Imports System
  Imports System.IO
#End Region

Public Class RecordDaily
  Implements IEquatable(Of RecordDaily)
  Implements IRegisterKey(Of Date)
  Implements IComparable(Of RecordDaily)
  Implements IMemoryStream
  Implements IFormatData
  Implements IPriceVol
  Implements IDateTrade
  Implements IDateUpdate

  Private MyException As Exception
  Private Shared MyListHeaderInfo As List(Of HeaderInfo)

#Region "New"
  Public Sub New(ByVal DateUpdate As Date)
    With Me
      .DateUpdate = Now
      .DateDay = .DateUpdate.Date
    End With
    If MyListHeaderInfo Is Nothing And LIST_OF_HEADER_FILE_ENABLED Then
      Dim ThisFile = My.Application.Info.DirectoryPath & "\HeaderInfo\" & TypeName(Me) & ".HeaderInfo.json"
      MyListHeaderInfo = FileHeaderRead(ThisFile, ListOfHeader, MyException)
    Else
      MyListHeaderInfo = ListOfHeader()
    End If
  End Sub

  Public Sub New()
    Me.New(Now)
  End Sub

  Public Sub New(ByRef Parent As Stock, ByRef Stream As Stream)
    Me.New()
    With Me
      .SerializeLoadFrom(Stream)
      .Stock = Parent
      If .Stock IsNot Nothing Then
        .StockID = .Stock.ID
        .Stock.RecordsDaily.Add(Me)
      End If
    End With
  End Sub
#End Region
#Region "Relational Properties"
  Public Property ID As Integer
  Public Property StockID As Integer
  Public Property Stock As Stock
#End Region
#Region "IPriceVol"
  Public ReadOnly Property AsIPriceVol As IPriceVol Implements IPriceVol.AsIPriceVol
    Get
      Return Me
    End Get
  End Property
  Public Property DateDay As Date Implements IPriceVol.DateDay
  Public Property DateUpdate As Date Implements IPriceVol.DateUpdate
  Public Property LastAdjusted As Single Implements IPriceVol.LastAdjusted
  Public Property Open As Single Implements IPriceVol.Open
  Public Property OpenNext As Single Implements IPriceVol.OpenNext
  Public Property Last As Single Implements IPriceVol.Last
  Public Property High As Single Implements IPriceVol.High
  Public Property Low As Single Implements IPriceVol.Low
  Public Property Vol As Integer Implements IPriceVol.Vol
  Public Property IsIntraDay As Boolean Implements IPriceVol.IsIntraDay
  Public Property VolMinus As Integer Implements IPriceVol.VolMinus
  Public Property VolPlus As Integer Implements IPriceVol.VolPlus
  Public Property LastWeighted As Single Implements IPriceVol.LastWeighted
  Public Property Range As Single Implements IPriceVol.Range
  Public Property LastPrevious As Single Implements IPriceVol.LastPrevious

  Public Property IsSpecialDividendPayout As Boolean Implements IPriceVol.IsSpecialDividendPayout
  Public Property SpecialDividendPayoutValue As Single Implements IPriceVol.SpecialDividendPayoutValue
#End Region
#Region "Others Functions"
  Friend Function CopyDeep(ByRef Parent As Stock, Optional ByVal IsIgnoreID As Boolean = False) As YahooAccessData.RecordDaily
    Dim ThisRecordDaily = New RecordDaily
    With ThisRecordDaily
      If IsIgnoreID = False Then .ID = Me.ID
      .DateDay = Me.DateDay
      .DateUpdate = Me.DateUpdate
      .Open = Me.Open
      .Last = Me.Last
      .LastAdjusted = Me.LastAdjusted
      .High = Me.High
      .Low = Me.Low
      .Vol = Me.Vol
      .Stock = Parent
      If .Stock IsNot Nothing Then
        .StockID = .Stock.ID
        .Stock.RecordsDaily.Add(ThisRecordDaily)
      End If
    End With
    Return ThisRecordDaily
  End Function

  ''' <summary>
  ''' This function check that the transaction high low are consistent with the open and last
  ''' </summary>
  ''' <returns>Return True if the record is consistent</returns>
  ''' <remarks>This function can also optionally fix the issues </remarks>
  Public Function CheckPrice(Optional ByVal IsApplyCorrection As Boolean = False) As Boolean
    Return YahooAccessData.Record.CheckPrice(Me, IsApplyCorrection)
  End Function
#End Region
#Region "IFormatData"
  Public Function ToStingOfData() As String() Implements IFormatData.ToStingOfData
    Return Extensions.ToStingOfData(Of RecordDaily)(Me)
  End Function

  Public Function ToListOfHeader() As List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
    Return MyListHeaderInfo
  End Function

  Public Shared Function ListOfHeader() As List(Of HeaderInfo)
    Dim ThisListHeaderInfo As New List(Of HeaderInfo)
    With ThisListHeaderInfo
      .Add(New HeaderInfo With {.Name = "ID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateDay", .Title = .Name, .Format = "{0:d}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateUpdate", .Title = .Name, .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Open", .Title = .Name, .Format = "{0:0.00}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "High", .Title = .Name, .Format = "{0:0.00}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Low", .Title = .Name, .Format = "{0:0.00}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Last", .Title = .Name, .Format = "{0:0.00}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "LastAdjusted", .Title = .Name, .Format = "{0:0.00}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Vol", .Title = "Volume", .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
    End With
    Return ThisListHeaderInfo
  End Function
#End Region
#Region "IEquatable"
  Public Function EqualsDeep(other As RecordDaily, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
    With other
      If IsIgnoreID = False Then
        If .ID <> Me.ID Then Return False
        If .StockID <> Me.StockID Then Return False
      End If
      If .DateDay <> Me.DateDay Then Return False
      If .DateUpdate <> Me.DateUpdate Then Return False
      If .Open <> Me.Open Then Return False
      If .Last <> Me.Last Then Return False
      If .LastAdjusted <> Me.LastAdjusted Then Return False
      If .High <> Me.High Then Return False
      If .Low <> Me.Low Then Return False
      If .Vol <> Me.Vol Then Return False
    End With
    Return True
  End Function

  Public Overloads Function Equals(other As RecordDaily) As Boolean Implements IEquatable(Of RecordDaily).Equals
    If other Is Nothing Then Return False
    If Me.DateUpdate = other.DateUpdate Then
      If Me.Stock IsNot Nothing Then
        Return Me.Stock.Equals(other.Stock)
      Else
        Return False
      End If
    Else
      Return False
    End If
  End Function

  Public Overrides Function Equals(obj As Object) As Boolean
    If (TypeOf obj Is Record) Then
      Return Me.Equals(DirectCast(obj, Record))
    Else
      Return (False)
    End If
  End Function

  Public Overrides Function GetHashCode() As Integer
    Return Me.DateDay.GetHashCode()
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
      Return Me.DateDay
    End Get
    Set(value As Date)

    End Set
  End Property
#End Region
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
  Public Function CompareTo(other As RecordDaily) As Integer Implements System.IComparable(Of RecordDaily).CompareTo
    Return Me.KeyValue.CompareTo(other.KeyValue)
  End Function
#End Region
#Region "IMemoryStream"
  Public Sub SerializeSaveTo(ByRef Stream As System.IO.Stream) Implements IMemoryStream.SerializeSaveTo
    Me.SerializeSaveTo(Stream, IMemoryStream.enuFileType.Standard)
  End Sub

  Public Sub SerializeSaveTo(ByRef Stream As System.IO.Stream, FileType As IMemoryStream.enuFileType) Implements IMemoryStream.SerializeSaveTo
    Dim ThisBinaryWriter As New BinaryWriter(Stream)
    With ThisBinaryWriter
      .Write(VERSION_MEMORY_STREAM)
      .Write(Me.ID)
      .Write(Me.StockID)
      .Write(Me.DateDay.ToBinary)
      .Write(Me.DateUpdate.ToBinary)
      .Write(Me.Open)
      .Write(Me.Last)
      .Write(Me.High)
      .Write(Me.Low)
      .Write(Me.Vol)
      .Write(Me.LastAdjusted)
    End With
  End Sub

  Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream) Implements IMemoryStream.SerializeLoadFrom
    Me.SerializeLoadFrom(Stream, False)
  End Sub

  Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream, IsRecordVirtual As Boolean) Implements IMemoryStream.SerializeLoadFrom
    Dim ThisBinaryReader As New BinaryReader(Stream, New System.Text.UTF8Encoding(), leaveOpen:=True)
    Dim ThisVersion As Single

    With ThisBinaryReader
      ThisVersion = .ReadSingle
      Me.ID = .ReadInt32
      Me.StockID = .ReadInt32
      Me.DateDay = DateTime.FromBinary(.ReadInt64)
      Me.DateUpdate = DateTime.FromBinary(.ReadInt64)
      Me.Open = .ReadSingle
      Me.Last = .ReadSingle
      Me.High = .ReadSingle
      Me.Low = .ReadSingle
      Me.Vol = .ReadInt32
      Me.LastAdjusted = .ReadSingle
    End With
    ThisBinaryReader.Dispose()
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
#Region "IDateUpdate"
  Private Property IDateUpdate_DateStart As Date Implements IDateUpdate.DateStart
    Get
      Return Me.DateDay
    End Get
    Set(value As Date)

    End Set
  End Property
  Private Property IDateUpdate_DateStop As Date Implements IDateUpdate.DateStop
    Get
      Return Me.DateDay
    End Get
    Set(value As Date)

    End Set
  End Property
  Private ReadOnly Property IDateUpdate_DateUpdate As Date Implements IDateUpdate.DateUpdate
    Get
      Return Me.DateUpdate
    End Get
  End Property

  Public ReadOnly Property IDateUpdate_DateLastTrade As Date Implements IDateUpdate.DateLastTrade
    Get
      Return Me.DateDay
    End Get
  End Property

  Public ReadOnly Property IDateUpdate_DateDay As Date Implements IDateUpdate.DateDay
    Get
      Return Me.DateDay
    End Get
  End Property
#End Region 'IDateUpdate
#Region "IDateTrade"
  Private Property IDateTrade_DateStart As Date Implements IDateTrade.DateStart
    Get
      Return Me.DateDay
    End Get
    Set(value As Date)

    End Set
  End Property
  Private Property IDateTrade_DateStop As Date Implements IDateTrade.DateStop
    Get
      Return Me.DateDay
    End Get
    Set(value As Date)
    End Set
  End Property
#End Region   'IDateTrade
End Class
