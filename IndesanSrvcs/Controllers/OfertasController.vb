Imports System.IdentityModel.Tokens.Jwt
Imports System.Web.Http
Imports IndesanSrvcs.QueryJson
Imports IndesanSrvcs.Controllers
Imports System.Web.Script.Serialization
Imports System.Net.Mail

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

			Dim strbody As String
			Dim strListaArticulos As String = ""

			pedido.ListaArticulos.ForEach(Sub(ped)
											  strListaArticulos = strListaArticulos & $"<li style='margin-top:3em'>{ped.Cod} - {ped.Desc.es} <p style='color:blue;font-size:10 px; margin-left:5em;'> <em>{ped.Desc2.es}</em></li>"
										  End Sub)
			strbody = $"<div>Pedido:{pedido.CodPedido}</div><div style='margin-top:3em'>Cliente: {pedido.DatosCliente.CodCliente} - {pedido.DatosCliente.Rzs} ({pedido.DatosCliente.Nombrecomercial})</div> <div><ul>{strListaArticulos}<ul></div>"

			enviarCorreo("compras@indesan.com", "Pedido Realizado desde la web", strbody)



			Return Ok("{""data"": {""resultado"":""OK"" }, ""output"":" & js.Serialize(pedido) & "}")



		End Function

		Function enviarCorreo(strcuenta, strSubject, strBody) As String

			'Enviamos un correo
			'En el caso en que todo lo anterior 


			Dim SMTP_SSL As Boolean = Boolean.Parse(ConfigurationManager.AppSettings("SMTP_SSL"))

			Dim SMTP_SERVER As String = Environment.GetEnvironmentVariable("SMTP_SERVER")
			Dim SMTP_PORT As Integer = Integer.Parse(Environment.GetEnvironmentVariable("SMTP_PORT"))
			Dim SMTP_USER As String = Environment.GetEnvironmentVariable("SMTP_USER")
			Dim SMTP_PASSWORD As String = Environment.GetEnvironmentVariable("SMTP_PASSWORD")


			Dim client As New SmtpClient()




			client.DeliveryMethod = SmtpDeliveryMethod.Network
			client.EnableSsl = SMTP_SSL
			client.Host = SMTP_SERVER
			client.Port = SMTP_PORT

			' setup Smtp authentication
			Dim credentials As System.Net.NetworkCredential = New System.Net.NetworkCredential(SMTP_USER, SMTP_PASSWORD)
			client.UseDefaultCredentials = False
			client.Credentials = credentials



			Dim msg As MailMessage = New MailMessage()



			msg.From = New MailAddress(SMTP_USER)
			msg.To.Add(New MailAddress(strcuenta))


			msg.Subject = strSubject
			msg.IsBodyHtml = True
			msg.Body = strBody

			Try
				client.Send(msg)
				Return "OK"
			Catch ex As Exception
				Return "EMAIL_FAILED"
			End Try


		End Function

	End Class
End Namespace