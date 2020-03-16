Imports System
Imports System.Net
Imports System.Security.Principal
Imports System.Threading
Imports System.Web.Http
Imports IndesanSrvcs.Models


Namespace Controllers
	<AllowAnonymous>
	<RoutePrefix("api/login")>
	Public Class LoginController
		Inherits ApiController

		<HttpPost>
		<Route("mensaje")>
		Public Function mensaje(<FromBody> msg As Newtonsoft.Json.Linq.JObject) As IHttpActionResult
			Dim qj As New QueryJson

			msg.Add("status", "recibido")
			Return Ok(qj.RegistrarMensaje(msg))


		End Function

		<HttpGet>
		<Route("echoping")>
		Public Function EchoPing() As IHttpActionResult
			Return Ok(True)
		End Function

		<HttpGet>
		<Route("echouser")>
		Public Function EchoUser() As IHttpActionResult
			Dim identity As IIdentity = Thread.CurrentPrincipal.Identity
			Return Ok($" IPrincipal-user: {identity.Name} - IsAuthenticated: {identity.IsAuthenticated}")
		End Function

		<HttpPost>
		<Route("authenticate")>
		Public Function Authenticate(login As LoginRequest) As IHttpActionResult
			If login Is Nothing Then
				Throw New HttpResponseException(HttpStatusCode.BadRequest)
			End If

			Dim cr As Credencial = VerificarCredenciales(login)
			Dim isCredentialValid As Boolean = (cr.IdCredencial <> -1)

			If isCredentialValid Then
				Dim token As String = TokenGenerator.GenerateTokenJwt(login.username, cr)
				Return Ok(token)
			Else

				Return Ok(cr.Password)
				'Return Content(HttpStatusCode.Unauthorized, cr.Password) 'cr.password contiene el mansaje de error

			End If

		End Function
		<HttpPost>
		<Route("register")>
		Public Function Register(candidato As RegisterRequest) As IHttpActionResult
			Dim strResultado As String
			If candidato Is Nothing Then
				Throw New HttpResponseException(HttpStatusCode.BadRequest)
			End If

			strResultado = VerificarCandidato(candidato)


			'If strResultado = "OK" Then

			Return Ok(strResultado)
			'Else

			'Return Content(HttpStatusCode.Unauthorized, strResultado)
			'End If

		End Function
		<HttpPost>
		<Route("activate")>
		Public Function Activate(candidato As ActivationRequest) As IHttpActionResult
			Dim strResultado As String
			If candidato Is Nothing Then
				Throw New HttpResponseException(HttpStatusCode.BadRequest)
			End If

			strResultado = VerificarActivacion(candidato)


			'If strResultado = "OK" Then

			Return Ok(strResultado)
			'Else

			'Return Content(HttpStatusCode.Unauthorized, strResultado)
			'End If

		End Function
		<HttpPost>
		<Route("passwordRecovery")>
		Public Function Recovery(candidato As RegisterRequest) As IHttpActionResult
			Dim strResultado As String
			If candidato Is Nothing Then
				Throw New HttpResponseException(HttpStatusCode.BadRequest)
			End If

			strResultado = VerificarPassword(candidato)


			'If strResultado = "OK" Then

			Return Ok(strResultado)
			'Else

			'Return Content(HttpStatusCode.Unauthorized, strResultado)
			'End If

		End Function

		Private Function VerificarCredenciales(login As LoginRequest) As Credencial
			Dim qj As New QueryJson()

			Return qj.GeneraCredencial(Username:=login.username, Password:=login.password)

		End Function
		Private Function VerificarCandidato(candidato As RegisterRequest) As String
			Dim qj As New QueryJson()

			Return qj.VerificarCandidato(Username:=candidato.username, Cif:=candidato.cif, Password:=candidato.password, lan:=candidato.lan)

		End Function
		Private Function VerificarPassword(candidato As RegisterRequest) As String
			Dim qj As New QueryJson()

			Return qj.VerificarPassword(Username:=candidato.username, lan:=candidato.lan)

		End Function

		Private Function VerificarActivacion(candidato As ActivationRequest) As String
			Dim qj As New QueryJson()

			Return qj.VerificarActivacion(candidato:=candidato.cli, codigo:=candidato.cod)

		End Function
	End Class
End Namespace