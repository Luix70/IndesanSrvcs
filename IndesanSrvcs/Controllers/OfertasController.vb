Imports System.IdentityModel.Tokens.Jwt
Imports System.Web.Http

Imports IndesanSrvcs.Controllers


Namespace Ofertas
	<Authorize>
	<RoutePrefix("api/ofertas")>
	Public Class OfertasController
		Inherits ApiController


		<HttpGet>
		<Route("GetAll")>
		Public Function GetAll() As IHttpActionResult

			'Dim jwt As JwtSecurityToken = ExtraerToken()


			Dim qj As New QueryJson()

			Dim resultados As String

			resultados = qj.ObtenerOfertas()

			Return Ok(resultados)


		End Function

		<HttpGet>
		<Route("custData")>
		Public Function CustData() As IHttpActionResult

			Dim jwt As JwtSecurityToken = ExtraerToken()


			Dim qj As New QueryJson()

			Dim resultados As String

			Dim cliente As String = jwt.Payload("AccesoCli")


			If IsNumeric(cliente) Then

				resultados = qj.CustData(cliente)
			Else
				resultados = "{codCliente: undefined }"
			End If


			Return Ok(resultados)


		End Function

		<HttpPost>
		<Route("cursarPedido")>
		Public Function cursarPedido(<FromBody()> pedido) As IHttpActionResult




			Return Ok("{""data"": {""resultado"":""OK"" }, ""output"":" & pedido.ToString & "}")




		End Function


	End Class
End Namespace