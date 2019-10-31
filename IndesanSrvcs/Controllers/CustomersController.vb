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
		<Route("GetId")>
		Public Function GetId(ByVal id As Int32) As IHttpActionResult
			Dim customerFake = "customer-fake"
			Return Ok(customerFake)
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
		Public Function LifetimeValidator(ByVal notBefore As DateTime?, ByVal expires As DateTime?, ByVal securityToken As SecurityToken, ByVal validationParameters As TokenValidationParameters) As Boolean
			If (Not (expires) Is Nothing) Then
				If (DateTime.UtcNow < expires) Then
					Return True
				End If

			End If

			Return False
		End Function

		Public Function ExtraerToken() As JwtSecurityToken
			Try
				Dim token As String = HttpContext.Current.Request.Headers("Authorization").Substring(7)
				Dim tokenHandler As New JwtSecurityTokenHandler()
				Dim jwToken As JwtSecurityToken = tokenHandler.ReadJwtToken(token)
				If jwToken Is Nothing Then
					Return New JwtSecurityToken()
				End If

				Dim secretKey = ConfigurationManager.AppSettings("JWT_SECRET_KEY")
				Dim audienceToken = ConfigurationManager.AppSettings("JWT_AUDIENCE_TOKEN")
				Dim issuerToken = ConfigurationManager.AppSettings("JWT_ISSUER_TOKEN")
				Dim securityKey = New SymmetricSecurityKey(System.Text.Encoding.Default.GetBytes(secretKey))
				Dim securityToken As JwtSecurityToken = Nothing

				Dim validationParameters As New TokenValidationParameters() With {
						.ValidAudience = audienceToken,
						.ValidIssuer = issuerToken,
						.ValidateLifetime = True,
						.ValidateIssuerSigningKey = True,
						.LifetimeValidator = AddressOf Me.LifetimeValidator,
						.IssuerSigningKey = securityKey
					}

				' Extract and assign Current Principal and user
				HttpContext.Current.User = tokenHandler.ValidateToken(token, validationParameters, securityToken)
				Return securityToken
			Catch ex As SecurityTokenValidationException
				Return New JwtSecurityToken()
			Catch ex As Exception
				Return New JwtSecurityToken()

			End Try

		End Function
	End Class
End Namespace