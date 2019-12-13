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
				Dim errorMsg As String
				If cr.Email = "NOOK" Then
					errorMsg = "UI"

				Else
					errorMsg = "CI"
				End If
				Return Content(HttpStatusCode.Unauthorized, errorMsg)
			End If

		End Function

		Private Function VerificarCredenciales(login As LoginRequest) As Credencial
			Dim qj As New QueryJson()

			Return qj.GeneraCredencial(Username:=login.username, Password:=login.password)

		End Function
	End Class
End Namespace