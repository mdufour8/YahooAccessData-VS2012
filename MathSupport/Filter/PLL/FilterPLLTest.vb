''' <summary>
''' An example of a file that can be use as a test
''' </summary>
Module FilterPLLTest
	Sub Main()
		' Create an instance of the FilterPLL class
		Dim filterRate As Double = 10.0
		Dim dampingFactor As Double = 0.5
		Dim pllFilter As New FilterPLL(filterRate, dampingFactor)

		' Define a series of input values
		Dim inputValues As Double() = {1.0, 2.0, 3.0, 4.0, 5.0}
		Dim filteredValues As New List(Of Double)

		' Process each input value using the FilterRun method
		For Each value As Double In inputValues
			Dim filteredValue As Double = pllFilter.FilterRun(value)
			filteredValues.Add(filteredValue)
		Next

		' Output the filtered values
		For Each filteredValue As Double In filteredValues
			Console.WriteLine(filteredValue)
		Next
	End Sub
End Module
