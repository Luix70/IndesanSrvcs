Imports System
Namespace Models
	Public Class LoginRequest
		Private newUsername As String
		Private newPassword As String
		Public Property Username() As String
			Get
				Return newUsername
			End Get
			Set(ByVal value As String)
				newUsername = value
			End Set
		End Property

		Public Property Password() As String
			Get
				Return newPassword
			End Get
			Set(ByVal value As String)
				newPassword = value
			End Set
		End Property
	End Class

End Namespace