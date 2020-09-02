Imports System.IdentityModel.Tokens.Jwt
Imports System.Web.Http
Imports System.Web.Script.Serialization
Imports Microsoft.IdentityModel.Tokens


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
		<Route("GetAll")>
		Public Function GetAll() As IHttpActionResult

			Dim jwt As JwtSecurityToken = ExtraerToken()


			Dim qj As New QueryJson()

			Dim resultados As String

			resultados = qj.ObtenerDocs(jwt.Payload("AccesoRep"), jwt.Payload("AccesoCli"), jwt.Payload("TipoEntidad"))



			Return Ok(resultados)


		End Function

	End Class
End Namespace