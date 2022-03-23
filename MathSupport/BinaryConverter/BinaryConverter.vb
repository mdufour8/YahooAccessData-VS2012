Imports System.Runtime.InteropServices

Namespace MathPlus.BinaryConverter
  ''' <summary>
  ''' use to convert an signed and unsigned number directly 
  ''' </summary>
  ''' <remarks></remarks>
  <StructLayout(LayoutKind.Explicit)>
  Public Structure UnionForInteger
    Sub New(ByVal Value As Integer)
      Me.Int32 = Value
    End Sub

    Public Sub New(ByVal Value As UInteger)
      Me.UInt32 = Value
    End Sub

    Public Sub New(ByVal Value As UnionForInteger)
      Me.Int32 = Value.Int32
    End Sub

    Sub New(ByVal ValueLow As Short, ByVal ValueHigh As Short)
      Me.Int16Low = ValueLow
      Me.Int16High = ValueHigh
    End Sub

    Sub New(ByVal ValueLow As UShort, ByVal ValueHigh As UShort)
      Me.UInt16Low = ValueLow
      Me.UInt16High = ValueHigh
    End Sub

    <FieldOffset(0)>
    Public Int32 As Integer
    <FieldOffset(0)>
    Public UInt32 As UInteger
    <FieldOffset(0)>
    Public Int16Low As Short
    <FieldOffset(2)>
    Public Int16High As Short
    <FieldOffset(0)>
    Public UInt16Low As UShort
    <FieldOffset(2)>
    Public UInt16High As UShort
  End Structure
  '<StructLayout(LayoutKind.Explicit)>
  'Public Structure UnionForShort
  '  <FieldOffset(0)>
  '  Public Int16 As Short
  '  <FieldOffset(0)>
  '  Public UInt16 As UShort
  'End Structure
  <StructLayout(LayoutKind.Explicit)>
  Public Structure UnionForLong
    Sub New(ByVal Value As Long)
      Me.Int64 = Value
    End Sub

    Sub New(ByVal Value As ULong)
      Me.UInt64 = Value
    End Sub

    Sub New(ByVal ValueLow As Integer, ByVal ValueHigh As Integer)
      Me.Int32Low = ValueLow
      Me.Int32High = ValueHigh
    End Sub

    Sub New(ByVal ValueLow As UInteger, ByVal ValueHigh As UInteger)
      Me.UInt32Low = ValueLow
      Me.UInt32High = ValueHigh
    End Sub

    Sub New(ByVal ValueLow As UnionForInteger, ByVal ValueHigh As UnionForInteger)
      Me.Int32LowUnion = ValueLow
      Me.Int32HighUnion = ValueHigh
    End Sub

    Sub New(ByVal Value As UnionForLong)
      Me.Int64 = Value.Int64
    End Sub

    <FieldOffset(0)>
    Public Int64 As Long
    <FieldOffset(0)>
    Public UInt64 As ULong
    <FieldOffset(0)>
    Public Int32Low As Integer
    <FieldOffset(4)>
    Public Int32High As Integer
    <FieldOffset(0)>
    Public UInt32Low As UInteger
    <FieldOffset(4)>
    Public UInt32High As UInteger
    <FieldOffset(0)>
    Public Int32LowUnion As UnionForInteger
    <FieldOffset(4)>
    Public Int32HighUnion As UnionForInteger
  End Structure

  'Public Class BinaryConverterToGrayCode
  '  ''' <summary>
  '  ''' use to convert an signed and unsigned number directly 
  '  ''' </summary>
  '  ''' <remarks></remarks>
  '  <StructLayout(LayoutKind.Explicit)>
  '  Public Structure UnionForInteger
  '    <FieldOffset(0)>
  '    Public Int32 As Integer
  '    <FieldOffset(0)>
  '    Public UInt32 As UInteger
  '    <FieldOffset(0)>
  '    Public Int16Low As Short
  '    <FieldOffset(2)>
  '    Public Int16High As Short
  '    <FieldOffset(0)>
  '    Public UInt16Low As UShort
  '    <FieldOffset(2)>
  '    Public UInt16High As UShort
  '  End Structure
  '  <StructLayout(LayoutKind.Explicit)>
  '  Private Structure UnionForShort
  '    <FieldOffset(0)>
  '    Public Int16 As Short
  '    <FieldOffset(0)>
  '    Public UInt16 As UShort
  '  End Structure
  '  <StructLayout(LayoutKind.Explicit)>
  '  Public Structure UnionForLong
  '    <FieldOffset(0)>
  '    Public Int64 As Long
  '    <FieldOffset(0)>
  '    Public UInt64 As ULong
  '    <FieldOffset(0)>
  '    Public Int32Low As Integer
  '    <FieldOffset(4)>
  '    Public Int32High As Integer
  '    <FieldOffset(0)>
  '    Public UInt32Low As UInteger
  '    <FieldOffset(4)>
  '    Public UInt32High As UInteger
  '    <FieldOffset(0)>
  '    Public Int32LowUnion As UnionForInteger
  '    <FieldOffset(4)>
  '    Public Int32HighUnion As UnionForInteger
  '  End Structure

  '  Public Shared Function BinaryToGray(ByVal Value As UInteger) As UInteger
  '    Return (Value >> 1) Xor Value
  '  End Function

  '  Public Shared Function GrayToBinary(ByVal Value As UInteger) As UInteger
  '    Dim ThisMask As UInteger
  '    ThisMask = Value >> 1
  '    While ThisMask <> 0
  '      Value = Value Xor ThisMask
  '      ThisMask = ThisMask >> 1
  '    End While
  '    Return Value
  '  End Function

  '  Public Shared Function BinaryToGray(ByVal Value As UShort) As UShort
  '    Return (Value >> 1) Xor Value
  '  End Function

  '  Public Shared Function GrayToBinary(ByVal Value As UShort) As UShort
  '    Dim ThisMask As UShort
  '    ThisMask = Value >> 1
  '    While ThisMask <> 0
  '      Value = Value Xor ThisMask
  '      ThisMask = ThisMask >> 1
  '    End While
  '    Return Value
  '  End Function

  '  Public Shared Function ToBitArray(ByVal Value As UInteger) As BitArray
  '    Return New BitArray(New Integer() {CInt(Value)})
  '  End Function

  '  Public Shared Function ToBooleanArray(ByVal Value As UInteger) As Boolean()
  '    Dim ThisBitArray = BinaryConverterToGrayCode.ToBitArray(Value)
  '    Dim ThisArray(0 To ThisBitArray.Length - 1) As Boolean
  '    ThisBitArray.CopyTo(ThisArray, 0)
  '    Return ThisArray
  '  End Function

  '  Public Shared Function ToBitString(ByVal Value As UInteger) As String
  '    Return BinaryConverterToGrayCode.ToBitString(Value, 16)
  '  End Function

  '  Public Shared Function ToBitString(ByVal Value As UInteger, ByVal NumberBit As Integer) As String
  '    Dim ThisStringBuilder As New System.Text.StringBuilder(capacity:=16)
  '    If NumberBit < 1 Then
  '      NumberBit = 1
  '    ElseIf NumberBit > 16 Then
  '      NumberBit = 16
  '    End If
  '    With BinaryConverterToGrayCode.ToBitArray(Value)
  '      For I = (NumberBit - 1) To 0 Step -1
  '        If .Item(I) Then
  '          ThisStringBuilder.Append("1")
  '        Else
  '          ThisStringBuilder.Append("0")
  '        End If
  '      Next
  '    End With
  '    Return ThisStringBuilder.ToString
  '  End Function

  '  Public Shared Function ToBitInteger(ByVal Value As UInteger) As Integer()
  '    Return BinaryConverterToGrayCode.ToBitInteger(Value, 16)
  '  End Function

  '  Public Shared Function ToBitInteger(ByVal Value As UInteger, ByVal NumberBit As Integer) As Integer()
  '    Dim ThisData() As Integer

  '    If NumberBit < 1 Then
  '      NumberBit = 1
  '    ElseIf NumberBit > 16 Then
  '      NumberBit = 16
  '    End If
  '    ReDim ThisData(0 To NumberBit - 1)
  '    With BinaryConverterToGrayCode.ToBitArray(Value)
  '      For I = 0 To (NumberBit - 1)
  '        If .Item(I) Then
  '          ThisData(I) = 1
  '        Else
  '          ThisData(I) = 0
  '        End If
  '      Next
  '    End With
  '    Return ThisData
  '  End Function

  '  Public Shared Function IsBitSet(Value As UInteger, pos As Integer) As Boolean
  '    Return (Value And (1 << pos)) <> 0
  '  End Function

  '  Public Shared Function IsBitSet(Value As Integer, pos As Integer) As Boolean
  '    Return (Value And (1 << pos)) <> 0
  '  End Function

  '  Public Shared Function IsBitSet(Value As Short, pos As Integer) As Boolean
  '    Return (Value And (1 << pos)) <> 0
  '  End Function

  '  Public Shared Function IsBitSet(Value As UShort, pos As Integer) As Boolean
  '    Return (Value And (1 << pos)) <> 0
  '  End Function

  '  ''' <summary>
  '  ''' Convert directly signed an unsigned binary number
  '  ''' </summary>
  '  ''' <param name="Value"></param>
  '  ''' <returns></returns>
  '  ''' <remarks></remarks>
  '  Public Shared Function ToUnsigned(ByVal Value As Integer) As UInteger
  '    Dim ThisConversionForInteger As UnionForInteger
  '    ThisConversionForInteger.Int32 = Value
  '    Return ThisConversionForInteger.UInt32
  '  End Function

  '  ''' <summary>
  '  ''' Convert directly signed an unsigned binary number
  '  ''' </summary>
  '  ''' <param name="Value"></param>
  '  ''' <returns></returns>
  '  ''' <remarks></remarks>
  '  Public Shared Function ToUnsigned(ByVal Value As Short) As UShort
  '    Dim ThisConversionForShort As UnionForShort
  '    ThisConversionForShort.Int16 = Value
  '    Return ThisConversionForShort.UInt16
  '  End Function

  '  Public Shared Function ToUnionForInteger(ByVal Value As Integer) As UnionForInteger
  '    Dim ThisConversionForInteger As UnionForInteger
  '    ThisConversionForInteger.Int32 = Value
  '    Return ThisConversionForInteger
  '  End Function
  'End Class

  'Public Class BinaryUnionForLong
  '  Implements IUnionForLong

  '  Private ThisUnionForlong As BinaryConverterToGrayCode.UnionForLong

  '  Public Sub New()
  '    Me.New(0)
  '  End Sub

  '  Public Sub New(ByVal Value As Long)
  '    ThisUnionForlong.Int64 = Value
  '  End Sub

  '  Public Sub New(ByVal Value As ULong)
  '    ThisUnionForlong.UInt64 = Value
  '  End Sub

  '  Public Property Int64 As Long Implements IUnionForLong.Int64
  '    Get
  '      Return ThisUnionForlong.Int64
  '    End Get
  '    Set(value As Long)
  '      ThisUnionForlong.Int64 = value
  '    End Set
  '  End Property

  '  Public Property UInt64 As ULong Implements IUnionForLong.UInt64
  '    Get
  '      Return ThisUnionForlong.UInt64
  '    End Get
  '    Set(value As ULong)
  '      ThisUnionForlong.UInt64 = value
  '    End Set
  '  End Property

  '  Public Property Int32Low As Integer Implements IUnionForLong.Int32Low
  '    Get
  '      Return ThisUnionForlong.Int32Low
  '    End Get
  '    Set(value As Integer)
  '      ThisUnionForlong.Int32Low = value
  '    End Set
  '  End Property

  '  Public Property Int32High As Integer Implements IUnionForLong.Int32High
  '    Get
  '      Return ThisUnionForlong.Int32High
  '    End Get
  '    Set(value As Integer)
  '      ThisUnionForlong.Int32High = value
  '    End Set
  '  End Property

  '  Public Property UInt32Low As UInteger Implements IUnionForLong.UInt32Low
  '    Get
  '      Return ThisUnionForlong.UInt32Low
  '    End Get
  '    Set(value As UInteger)
  '      ThisUnionForlong.UInt32Low = value
  '    End Set
  '  End Property

  '  Public Property UInt32High As UInteger Implements IUnionForLong.UInt32High
  '    Get
  '      Return ThisUnionForlong.UInt32High
  '    End Get
  '    Set(value As UInteger)
  '      ThisUnionForlong.UInt32High = value
  '    End Set
  '  End Property

  '  Public Property Int32LowUnion As BinaryConverterToGrayCode.UnionForInteger Implements IUnionForLong.Int32LowUnion
  '    Get
  '      Return ThisUnionForlong.Int32LowUnion
  '    End Get
  '    Set(value As BinaryConverterToGrayCode.UnionForInteger)
  '      ThisUnionForlong.Int32LowUnion = value
  '    End Set
  '  End Property

  '  Public Property Int32HighUnion As BinaryConverterToGrayCode.UnionForInteger Implements IUnionForLong.Int32HighUnion
  '    Get
  '      Return ThisUnionForlong.Int32HighUnion
  '    End Get
  '    Set(value As BinaryConverterToGrayCode.UnionForInteger)
  '      ThisUnionForlong.Int32HighUnion = value
  '    End Set
  '  End Property
  'End Class

  'Public Class BinaryUnionForInteger
  '  Implements IUnionForInteger

  '  Private ThisUnionForInteger As BinaryConverterToGrayCode.UnionForInteger

  '  Public Sub New()
  '    Me.New(0)
  '  End Sub

  '  Public Sub New(ByVal Value As Integer)
  '    ThisUnionForInteger.Int32 = Value
  '  End Sub

  '  Public Sub New(ByVal Value As UInteger)
  '    ThisUnionForInteger.UInt32 = Value
  '  End Sub

  '  Public Property Int32 As Integer Implements IUnionForInteger.Int32
  '    Get
  '      Return ThisUnionForInteger.Int32
  '    End Get
  '    Set(value As Integer)
  '      ThisUnionForInteger.Int32 = value
  '    End Set
  '  End Property

  '  Public Property UInt32 As UInteger Implements IUnionForInteger.UInt32
  '    Get
  '      Return ThisUnionForInteger.UInt32
  '    End Get
  '    Set(value As UInteger)
  '      ThisUnionForInteger.UInt32 = value
  '    End Set
  '  End Property

  '  Public Property Int16Low As Short Implements IUnionForInteger.Int16Low
  '    Get
  '      Return ThisUnionForInteger.Int16Low
  '    End Get
  '    Set(value As Short)
  '      ThisUnionForInteger.Int16Low = value
  '    End Set
  '  End Property

  '  Public Property Int16High As Short Implements IUnionForInteger.Int16High
  '    Get
  '      Return ThisUnionForInteger.Int16High
  '    End Get
  '    Set(value As Short)
  '      ThisUnionForInteger.Int16High = value
  '    End Set
  '  End Property

  '  Public Property UInt16Low As UShort Implements IUnionForInteger.UInt16Low
  '    Get
  '      Return ThisUnionForInteger.UInt16Low
  '    End Get
  '    Set(value As UShort)
  '      ThisUnionForInteger.UInt16Low = value
  '    End Set
  '  End Property

  '  Public Property UInt16High As UShort Implements IUnionForInteger.UInt16High
  '    Get
  '      Return ThisUnionForInteger.UInt16High
  '    End Get
  '    Set(value As UShort)
  '      ThisUnionForInteger.UInt16High = value
  '    End Set
  '  End Property
  'End Class

  'Public Interface IUnionForInteger
  '  Property Int32 As Integer
  '  Property UInt32 As UInteger
  '  Property Int16Low As Short
  '  Property Int16High As Short
  '  Property UInt16Low As UShort
  '  Property UInt16High As UShort
  'End Interface
  'Public Interface IUnionForShort
  '  Property Int16 As Short
  '  Property UInt16 As UShort
  'End Interface
  'Public Interface IUnionForLong
  '  Property Int64 As Long
  '  Property UInt64 As ULong
  '  Property Int32Low As Integer
  '  Property Int32High As Integer
  '  Property UInt32Low As UInteger
  '  Property UInt32High As UInteger
  '  Property Int32LowUnion As BinaryConverterToGrayCode.UnionForInteger
  '  Property Int32HighUnion As BinaryConverterToGrayCode.UnionForInteger
  'End Interface
End Namespace