Imports System.IdentityModel.Tokens.Jwt
Imports System.Web.Http
Imports System.Web.Script.Serialization
Imports Microsoft.IdentityModel.Tokens
Imports IndesanSrvcs.Controllers
Namespace Ofertas
	<Authorize>
	<RoutePrefix("api/ofertas")>
	Public Class OfertasController
		Inherits ApiController


		<HttpGet>
		<Route("GetAll")>
		Public Function GetAll() As IHttpActionResult

			Dim jwt As JwtSecurityToken = ExtraerToken()


			Dim qj As New QueryJson()

			Dim resultados As String

			resultados = qj.ObtenerOfertas()

			Return Ok(resultados)


		End Function

	End Class
End Namespace