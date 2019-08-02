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

			Dim isCredentialValid As Boolean = (login.Password = "123456")
			'TODO: Validate credentials Correctly, this code Is only for demo !!
			If isCredentialValid Then
				Dim token As String = TokenGenerator.GenerateTokenJwt(login.Username)
				Return Ok(token)
			Else
				Return Unauthorized()
			End If

		End Function
	End Class
End Namespace