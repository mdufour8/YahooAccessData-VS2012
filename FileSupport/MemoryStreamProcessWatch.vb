Imports System
Imports System.IO

Public Class MemoryStreamWatch
  Inherits MemoryStream

  Private MyListOfWatch As List(Of IProcessTimeMeasurement)

#Region "New"
  Public Sub New()
    MyBase.New()
    MyListOfWatch = New List(Of IProcessTimeMeasurement)
  End Sub

  Public Sub New(buffer() As Byte)
    MyBase.New(buffer:=buffer)
    MyListOfWatch = New List(Of IProcessTimeMeasurement)
  End Sub

  Public Sub New(capacity As Integer)
    MyBase.New(capacity:=capacity)
    MyListOfWatch = New List(Of IProcessTimeMeasurement)
  End Sub

  Public Sub New(buffer() As Byte, writable As Boolean)
    MyBase.New(buffer:=buffer, writable:=writable)
    MyListOfWatch = New List(Of IProcessTimeMeasurement)
  End Sub

  Public Sub New(buffer() As Byte, index As Integer, count As Integer)
    MyBase.New(buffer:=buffer, index:=index, count:=count)
    MyListOfWatch = New List(Of IProcessTimeMeasurement)
  End Sub

  Public Sub New(buffer() As Byte, index As Integer, count As Integer, writable As Boolean)
    MyBase.New(buffer:=buffer, index:=index, count:=count, writable:=writable)
    MyListOfWatch = New List(Of IProcessTimeMeasurement)
  End Sub

  Public Sub New(buffer() As Byte, index As Integer, count As Integer, writable As Boolean, publiclyVisible As Boolean)
    MyBase.New(buffer:=buffer, index:=index, count:=count, writable:=writable, publiclyVisible:=publiclyVisible)
    MyListOfWatch = New List(Of IProcessTimeMeasurement)
  End Sub
#End Region
  Public Property KeyValue As String

  Public ReadOnly Property ToListOfProcessWatch() As IList(Of IProcessTimeMeasurement)
    Get
      Return MyListOfWatch
    End Get
  End Property
End Class
