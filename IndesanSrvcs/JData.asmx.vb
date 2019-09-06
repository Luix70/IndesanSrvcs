Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web
Imports System.Web.Services
Imports System.Web.Script.Services
Imports System.ComponentModel
Imports System.Web.Script.Serialization
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Web.Services.Protocols
' Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente.

<System.Web.Services.WebService(Namespace:="http://servicios.indesan.com/")>
<System.Web.Script.Services.ScriptService>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class JData
	Inherits System.Web.Services.WebService

	Private Const StrResponseNew As String = "Hello World"
	Public SoapHeader As TokenAuthUser

	<WebMethod()>
	<ScriptMethod(ResponseFormat:=ResponseFormat.Json)>
	Public Sub JPedidos()

		Dim parCodRep As String = HttpContext.Current.Request.Params("cr")
		Dim parCodCli As String = HttpContext.Current.Request.Params("cc")
		Dim parCodigodoc As String = HttpContext.Current.Request.Params("cd")
		Dim parTipodoc As String = HttpContext.Current.Request.Params("td")

		Dim qj As New QueryJson()

		Dim resultados As String

		resultados = qj.GenerarJson(parCodRep, parCodCli, parTipodoc, parCodigodoc)

		HttpContext.Current.Response.Clear()
		HttpContext.Current.Response.AppendHeader("Access-Control-Allow-Origin", "*")
		HttpContext.Current.Response.ContentType = "application/json;charset=UTF-8"
		HttpContext.Current.Response.Write(resultados)


	End Sub
	<WebMethod()>
	<ScriptMethod(ResponseFormat:=ResponseFormat.Json)>
	Public Sub JOps()

		Dim parUser As String = HttpContext.Current.Request.Params("user")
		Dim parPassWord As String = HttpContext.Current.Request.Params("password")

		Dim qj As New QueryJson()

		Dim resultados As String

		resultados = qj.GenerarJson2(parUser, parPassWord)

		HttpContext.Current.Response.Clear()
		HttpContext.Current.Response.AppendHeader("Access-Control-Allow-Origin", "*")
		' HttpContext.Current.Response.AppendHeader("Content-Encoding", "identity")
		HttpContext.Current.Response.ContentType = "application/json;charset=UTF-8"
		HttpContext.Current.Response.Write(resultados)


	End Sub

	<WebMethod()>
	<ScriptMethod(ResponseFormat:=ResponseFormat.Json)>
	Public Sub JSaluda()

		Dim js As New JavaScriptSerializer
		Dim saludo As New atomObject(StrResponseNew)
		HttpContext.Current.Response.Clear()
		HttpContext.Current.Response.AppendHeader("Access-Control-Allow-Origin", "*")
		' HttpContext.Current.Response.AppendHeader("Content-Encoding", "identity")
		HttpContext.Current.Response.ContentType = "application/json;charset=UTF-8"
		HttpContext.Current.Response.Write(js.Serialize(saludo))



	End Sub
	'<SoapHeader("user")>
	<WebMethod()>
	<ScriptMethod(ResponseFormat:=ResponseFormat.Json)>
	Public Sub JAuth()

		Dim parUser As String = HttpContext.Current.Request.Params("user")
		Dim parPassWord As String = HttpContext.Current.Request.Params("password")
		Dim strRespuesta As String = "Usuario inválido"
		Dim usr As New AuthUser
		usr.Username = parUser
		usr.Password = parPassWord
		Dim NombreUsuario As String = ""

		If usr.isValid(NombreUsuario) Then

			strRespuesta = "Bienvenido, " + NombreUsuario

		End If

		Dim js As New JavaScriptSerializer
		Dim saludo As New atomObject(strRespuesta)
		HttpContext.Current.Response.Clear()
		HttpContext.Current.Response.AppendHeader("Access-Control-Allow-Origin", "*")
		' HttpContext.Current.Response.AppendHeader("Content-Encoding", "identity")
		HttpContext.Current.Response.ContentType = "application/json;charset=UTF-8"
		HttpContext.Current.Response.Write(js.Serialize(saludo))

	End Sub

	<WebMethod()>
	<ScriptMethod(ResponseFormat:=ResponseFormat.Json)>
	<SoapHeader("SoapHeader")>
	Public Function JTokenAuth() As String

		If SoapHeader Is Nothing Then
			Return "Necesitas autentificarte antes de acceder a este método. Llama a AuthenticationMethod()"
		End If

		If Not SoapHeader.isUserCredentialsValid(SoapHeader) Then
			Return "Necesitas autentificarte antes de acceder a este método. Llama a AuthenticationMethod()"
		Else
			Return "Hola, " + HttpRuntime.Cache(SoapHeader.AuthenticationToken)
		End If
	End Function

	<WebMethod()>
	<ScriptMethod(ResponseFormat:=ResponseFormat.Json)>
	<SoapHeader("SoapHeader")>
	Public Function AuthenticationMethod() As String
		If SoapHeader Is Nothing Then
			Return "Por favor, suministra un SoapHeader"
		End If

		If String.IsNullOrEmpty(SoapHeader.Username) OrElse String.IsNullOrEmpty(SoapHeader.Password) Then
			Return "Por favor, suministra nombre de ususario y contraseña"
		End If

		If Not SoapHeader.isUserCredentialsValid(SoapHeader.Username, SoapHeader.Password) Then

			Return "Usuario o contraseña inválidos"

		End If

		Dim token As String = Guid.NewGuid().ToString()

		HttpRuntime.Cache.Add(token,
							  SoapHeader.Username,
							  Nothing,
							  System.Web.Caching.Cache.NoAbsoluteExpiration,
							  TimeSpan.FromMinutes(30),
							  System.Web.Caching.CacheItemPriority.NotRemovable,
							  Nothing)
		Return token

	End Function
	Private Class atomObject
		Private strResponse As String
		Public Sub New(ByVal strResponseNew As String)
			Me.strResponse = strResponseNew
		End Sub
		Public Property response As String
			Get
				Return strResponse
			End Get


			Set(value As String)
				strResponse = value
			End Set
		End Property
	End Class

End Class

Public Class AuthUser
	Inherits System.Web.Services.Protocols.SoapHeader
	Private strUsername As String
	Public Property Username() As String
		Get
			Return strUsername
		End Get
		Set(ByVal value As String)
			strUsername = value
		End Set
	End Property

	Private newPassword As String
	Public Property Password() As String
		Get
			Return newPassword
		End Get
		Set(ByVal value As String)
			newPassword = value
		End Set
	End Property

	Public Function isValid(ByRef strNombreCliente As String) As Boolean
		'Realizar toda la lógica para determinar si el usuario es válido
		'(Consultar database etc)
		Dim count As Int16 = 0
		Dim config As String = ConfigurationManager.ConnectionStrings("myCon").ConnectionString

		Dim strCadenaConsulta As String

		strCadenaConsulta = "Select * from CLIENTES_RST where nombreusuario ='" + Username + "' AND contras ='" + Password + "';"
		Dim Cons As New OleDb.OleDbConnection

		Cons.ConnectionString = config

		Cons.Open()

		Dim ndt As New DataTable
		Dim dad As New OleDb.OleDbDataAdapter(strCadenaConsulta, Cons)

		Try
			dad.Fill(ndt)
		Catch ex As Exception
			MsgBox(ex.Message)
		End Try

		count = ndt.Rows.Count

		Dim result As Boolean

		If count > 0 Then

			result = True
			strNombreCliente = ndt.Rows(0).Item("rzs")
		Else
			result = False
			strNombreCliente = ""
		End If


		Cons.Close()
		Cons = Nothing
		Return result

	End Function
End Class

Public Class TokenAuthUser
	Inherits System.Web.Services.Protocols.SoapHeader
	Private strUsername As String
	Public Property Username() As String
		Get
			Return strUsername
		End Get
		Set(ByVal value As String)
			strUsername = value
		End Set
	End Property

	Private newPassword As String
	Public Property Password() As String
		Get
			Return newPassword
		End Get
		Set(ByVal value As String)
			newPassword = value
		End Set
	End Property
	Private newAuthenticationToken As String
	Public Property AuthenticationToken() As String
		Get
			Return newAuthenticationToken
		End Get
		Set(ByVal value As String)
			newAuthenticationToken = value
		End Set
	End Property
	Public Function isUserCredentialsValid(username As String, password As String) As Boolean
		'Realizar toda la lógica para determinar si el usuario es válido
		'(Consultar database etc)
		Dim count As Int16 = 0
		Dim config As String = ConfigurationManager.ConnectionStrings("myCon").ConnectionString

		Dim strCadenaConsulta As String

		strCadenaConsulta = "Select * from CLIENTES_RST where nombreusuario ='" + username + "' AND contras ='" + password + "';"
		Dim Cons As New OleDb.OleDbConnection

		Cons.ConnectionString = config

		Cons.Open()

		Dim ndt As New DataTable
		Dim dad As New OleDb.OleDbDataAdapter(strCadenaConsulta, Cons)

		Try
			dad.Fill(ndt)
		Catch ex As Exception
			MsgBox(ex.Message)
		End Try

		count = ndt.Rows.Count

		Dim result As Boolean

		If count > 0 Then

			result = True

		Else
			result = False

		End If


		Cons.Close()
		Cons = Nothing
		Return result

	End Function

	Public Function isUserCredentialsValid(SoapHeader As TokenAuthUser) As Boolean
		If SoapHeader Is Nothing Then
			Return False
		End If

		If Not String.IsNullOrEmpty(SoapHeader.AuthenticationToken) Then
			Return HttpRuntime.Cache(SoapHeader.AuthenticationToken) IsNot Nothing
		End If

		Return False

	End Function
End Class