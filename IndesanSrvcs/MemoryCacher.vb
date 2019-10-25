Imports Microsoft.VisualBasic
Public Class MemoryCacher
	Public Function GetValue(key As String) As Object
		Dim mc As MemoryCache = MemoryCache.Default
		Return mc.Get(key)
	End Function

	Public Function Add(key As String, value As Object, absExpiration As DateTimeOffset) As Boolean
		Dim mc As MemoryCache = MemoryCache.Default
		Return mc.Add(key, value, absExpiration)
	End Function

	Public Sub Delete(key As String)
		Dim mc As MemoryCache = MemoryCache.Default
		If mc.Contains(key) Then
			mc.Remove(key)
		End If
	End Sub
End Class
