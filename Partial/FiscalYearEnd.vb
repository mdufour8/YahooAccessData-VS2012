#Region "Imports"
  Imports YahooAccessData.ExtensionService
  Imports System
  Imports System.IO
  Imports System.Collections.Generic
#End Region

''' <summary>
''' Not Completed yet
''' </summary>
''' <remarks></remarks>
'<Serializable()>
'Public Class FiscalYearEnd
'  Implements IEquatable(Of FiscalYearEnd)
'  Implements IRegisterKey(Of Date)
'  Implements IComparable(Of FiscalYearEnd)
'  Implements IMemoryStream
'  Implements IFormatData
'  Implements IDateUpdate

'  Public Property ID As          
'  Public Property DateYear As Date
'  Public Property DateUpdate As Date
'  Public Property FiscalYearEnd As Date
'  Public Property FiscalYearEndAnnounced As Date
'  Public Property StockID As Integer
'  Public Property Stock As Stock

'#Region "IEquatable"
'  Public Function EqualsDeep(other As FiscalYearEnd, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
'    If other Is Nothing Then Return False
'    With other
'      If IsIgnoreID = False Then
'        If .ID <> Me.ID Then Return False
'      End If
'      If .DateYear <> Me.DateYear Then Return False
'      If .DateUpdate <> Me.DateUpdate Then Return False
'      If .FiscalYearEnd <> Me.FiscalYearEnd Then Return False
'      If .FiscalYearEndAnnounced <> Me.FiscalYearEndAnnounced Then Return False
'    End With
'    Return True
'  End Function

'  Public Overloads Function Equals(other As FiscalYearEnd) As Boolean Implements IEquatable(Of FiscalYearEnd).Equals
'    If other Is Nothing Then Return False
'    If Me.KeyValue = other.KeyValue Then
'      Return True
'    Else
'      Return False
'    End If
'  End Function

'  Public Overrides Function Equals(obj As Object) As Boolean
'    If obj Is Nothing Then Return False
'    If (TypeOf obj Is StockSymbol) Then
'      Return Me.Equals(DirectCast(obj, StockSymbol))
'    Else
'      Return False
'    End If
'  End Function

'  Public Overrides Function GetHashCode() As Integer
'    Return Me.KeyValue.GetHashCode
'  End Function
'#End Region
'#Region "IComparable"
' ''' <summary>
' ''' </summary>
' ''' <param name="other"></param>
' ''' <returns>
' ''' Less than zero: This object is less than the other parameter. 
' ''' Zero : This object is equal to other. 
' ''' Greater than zero : This object is greater than other. 
' ''' </returns>
' ''' <remarks></remarks>
'  Public Function CompareTo(other As FiscalYearEnd) As Integer Implements System.IComparable(Of FiscalYearEnd).CompareTo
'    Return Me.KeyValue.CompareTo(other.KeyValue)
'  End Function
'#End Region
'#Region "Register Key"
'  Public Property KeyID As Integer Implements IRegisterKey(Of Date).KeyID
'    Get
'      Return Me.ID
'    End Get
'    Set(value As Integer)
'      Me.ID = value
'    End Set
'  End Property

'  Public Property KeyValue As Date Implements IRegisterKey(Of Date).KeyValue
'    Get
'      Return Me.DateYear
'    End Get
'    Set(value As Date)

'    End Set
'  End Property
'#End Region
'#Region "IMemoryStream"
'  Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream) Implements IMemoryStream.SerializeLoadFrom
'    Me.SerializeLoadFrom(Stream, False)
'  End Sub

'  Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream, IsRecordVirtual As Boolean) Implements IMemoryStream.SerializeLoadFrom
'    Dim ThisBinaryReader As New BinaryReader(Stream)
'    Dim ThisVersion As Single

'    With ThisBinaryReader
'      ThisVersion = .ReadSingle
'      Me.ID = .ReadInt32
'      Dim I As Integer
'      For I = 1 To .ReadInt32
'        MyException = New Exception(.ReadString, MyException)
'      Next
'      Me.Name = .ReadString
'      Me.Symbol = .ReadString
'      Me.SymbolNew = .ReadString
'      Me.Exchange = .ReadString
'      Me.DateUpdate = DateTime.FromBinary(.ReadInt64)
'      Me.IndustryID = .ReadInt32
'      Me.SectorID = .ReadInt32
'      Me.StockID = .ReadInt32
'    End With
'  End Sub

'  Public Sub SerializeSaveTo(ByRef Stream As System.IO.Stream) Implements IMemoryStream.SerializeSaveTo
'    Me.SerializeSaveTo(Stream, IMemoryStream.enuFileType.Standard)
'  End Sub

'  Public Sub SerializeSaveTo(ByRef Stream As System.IO.Stream, FileType As IMemoryStream.enuFileType) Implements IMemoryStream.SerializeSaveTo
'    Dim ThisBinaryWriter As New BinaryWriter(Stream)

'    With ThisBinaryWriter
'      .Write(VERSION_MEMORY_STREAM)
'      .Write(Me.ID)
'      .Write(Me.DateYear.ToBinary)
'      .Write(Me.DateUpdate.ToBinary)
'      .Write(Me.FiscalYearEnd.ToBinary)
'      .Write(Me.FiscalYearEndAnnounced.ToBinary)
'      .Write(Me.StockID)
'    End With
'  End Sub

'  Public Sub SerializeLoadFrom(ByRef Data() As Byte) Implements IMemoryStream.SerializeLoadFrom
'    Dim ThisStream As Stream = New System.IO.MemoryStream(Data, writable:=True)
'    Me.SerializeLoadFrom(ThisStream)
'  End Sub

'  Public Function SerializeSaveTo() As Byte() Implements IMemoryStream.SerializeSaveTo
'    Dim ThisStream As Stream = New System.IO.MemoryStream
'    Dim ThisBinaryReader As New BinaryReader(ThisStream)
'    Me.SerializeSaveTo(ThisStream)
'    Return ThisBinaryReader.ReadBytes(CInt(ThisStream.Length))
'  End Function
'#End Region
'End Class
