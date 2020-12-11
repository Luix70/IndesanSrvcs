Imports System.IdentityModel.Tokens.Jwt
Imports System.Web.Http
Imports System.Web.Script.Serialization
Imports Microsoft.IdentityModel.Tokens
Imports IndesanSrvcs.Models
Imports System.Net

Namespace Controllers
	<Authorize>
	<RoutePrefix("api/customers")>
	Public Class CustomersController
		Inherits ApiController

		<HttpGet>
		<Route("GetId/{id}")>
		Public Function GetId(ByVal id As String) As IHttpActionResult
			'Dim customer As String = "{" & $"""id"": ""{id}""" & "}"
			Dim qj As New QueryJson
			Dim customer As String = qj.DatosCliente(id)
			Return Ok(customer)
		End Function

		<HttpGet>
		<Route("GetRecogidas/{id}")>
		Public Function GetRecogidas(ByVal id As String) As IHttpActionResult
			'Dim customer As String = "{" & $"""id"": ""{id}""" & "}"
			Dim qj As New QueryJson
			Dim customer As String = qj.RecogidasCliente(id)
			Return Ok(customer)
		End Function

		<HttpGet>
		<Route("GetAll")>
		Public Function GetAll() As IHttpActionResult

			Dim jwt As JwtSecurityToken = ExtraerToken()


			Dim qj As New QueryJson()

			Dim resultados As String

			resultados = qj.ObtenerDocs(jwt.Payload("AccesoRep"), jwt.Payload("AccesoCli"), jwt.Payload("TipoEntidad"))



			Return Ok(resultados)


		End Function


		<HttpPost>
		<Route("savePrefs")>
		Public Function Recovery(usuario As Credencial) As IHttpActionResult

			If usuario Is Nothing Then
				Throw New HttpResponseException(HttpStatusCode.BadRequest)
				Return Ok("NoOK")
			End If

			Dim qj As New QueryJson
			Dim strres As String

			strres = qj.GuardarPreferencias(usuario)


			'Si no hay errores...
			Return Ok("OK")

		End Function
	End Class
End Namespace