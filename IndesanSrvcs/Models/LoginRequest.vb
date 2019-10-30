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

	Public Class Credencial
		Private IdCredencial_ As Long
		Private TipoEntidad_ As String
		Private NombreUsuario_ As String
		Private Password_ As String
		Private AccesoCli_ As String
		Private AccesoRep_ As String
		Private Email_ As String
		Private Idioma_ As String

		Public Property IdCredencial As Long
			Get
				Return IdCredencial_
			End Get
			Set(value As Long)
				IdCredencial_ = value
			End Set
		End Property

		Public Property TipoEntidad As String
			Get
				Return TipoEntidad_
			End Get
			Set(value As String)
				TipoEntidad_ = value
			End Set
		End Property

		Public Property NombreUsuario As String
			Get
				Return NombreUsuario_
			End Get
			Set(value As String)
				NombreUsuario_ = value
			End Set
		End Property

		Public Property Password As String
			Get
				Return Password_
			End Get
			Set(value As String)
				Password_ = value
			End Set
		End Property

		Public Property AccesoCli As String
			Get
				Return AccesoCli_
			End Get
			Set(value As String)
				AccesoCli_ = value
			End Set
		End Property

		Public Property AccesoRep As String
			Get
				Return AccesoRep_
			End Get
			Set(value As String)
				AccesoRep_ = value
			End Set
		End Property

		Public Property Email As String
			Get
				Return Email_
			End Get
			Set(value As String)
				Email_ = value
			End Set
		End Property

		Public Property Idioma As String
			Get
				Return Idioma_
			End Get
			Set(value As String)
				Idioma_ = value
			End Set
		End Property
	End Class

End Namespace