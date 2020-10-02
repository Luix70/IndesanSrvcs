Imports System.IdentityModel.Tokens.Jwt
Imports System.Web.Http
Imports IndesanSrvcs.QueryJson
Imports IndesanSrvcs.Controllers
Imports System.Web.Script.Serialization

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
		Public Function cursarPedido(<FromBody()> pedido As PedidoCarrito) As IHttpActionResult

			Dim js As JavaScriptSerializer
			js = New JavaScriptSerializer()
			js.MaxJsonLength = 100000000

			Dim qj As New QueryJson()

			'Comprobamos la existencia de los articulos y realizamos la reserva
			If Not qj.reservarOfertas(pedido.ListaArticulos) Then
				Return Ok("{""data"": {""resultado"":""Fail"", ""Motivo"":""STOCK"" }, ""output"":" & js.Serialize(pedido) & "}")
			End If


			'Hemos llegado hasta aqui, es por que hay existencias.
			'Procedemos a generar el pedido

			If Not qj.generarPedido(pedido) Then
				' hay que cancelar las ofertas
				For Each ofr As Oferta In pedido.ListaArticulos
					ofr.Reservadas = -1 * ofr.Reservadas
				Next

				qj.reservarOfertas(pedido.ListaArticulos)

				Return Ok("{""data"": {""resultado"":""Fail"", ""Motivo"":""PEDIDO""  }, ""output"":" & js.Serialize(pedido) & "}")

			End If





			'En el caso en que todo lo anterior 
			Return Ok("{""data"": {""resultado"":""OK"" }, ""output"":" & js.Serialize(pedido) & "}")

		End Function


	End Class
End Namespace