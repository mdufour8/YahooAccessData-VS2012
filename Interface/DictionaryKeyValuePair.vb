''' <summary>
''' Use to form a data link to the key and value of an element in a dictionary 
''' </summary>
''' <typeparam name="TKey"></typeparam>
''' <typeparam name="TValue"></typeparam>
''' <remarks></remarks>
<Serializable>
Public Class DictionaryKeyValuePair(Of TKey, TValue)
  Implements IDictionaryKeyValuePair(Of TKey, TValue)

  Public Sub New()
  End Sub

  Public Sub New(ByRef KeyValuePair As KeyValuePair(Of TKey, TValue))
    Me.Key = KeyValuePair.Key
    Me.Value = KeyValuePair.Value
  End Sub

  Public Sub New(ByRef KeyValuePair As IDictionaryKeyValuePair(Of TKey, TValue))
    Me.Key = KeyValuePair.Key
    Me.Value = KeyValuePair.Value
  End Sub

  Public Property Key As TKey Implements IDictionaryKeyValuePair(Of TKey, TValue).Key
  Public Property Value As TValue Implements IDictionaryKeyValuePair(Of TKey, TValue).Value
End Class

Public Interface IDictionaryKeyValuePair(Of TKey, TValue)
  Property Key As TKey
  Property Value As TValue
End Interface
