Imports System
Imports System.Collections.Generic
Imports System.Configuration
Imports System.IdentityModel.Tokens.Jwt
Imports System.Linq
Imports System.Net
Imports System.Net.Http
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports Microsoft.IdentityModel.Tokens

Namespace Controllers

	''' <summary>
	''' Token validator for Authorization Request using a DelegatingHandler
	''' </summary>
	Class TokenValidationHandler
		Inherits DelegatingHandler

		Private Shared Function TryRetrieveToken(ByVal request As HttpRequestMessage, ByRef token As String) As Boolean
			token = Nothing
			Dim authzHeaders As IEnumerable(Of String) = Nothing
			If (Not request.Headers.TryGetValues("Authorization", authzHeaders) OrElse (authzHeaders.Count > 1)) Then
				Return False
			End If

			Dim bearerToken = authzHeaders.ElementAt(0)
			If bearerToken.StartsWith("Bearer ") Then
				token = bearerToken.Substring(7)
			Else
				token = bearerToken
			End If

			Return True
		End Function

		Protected Overrides Function SendAsync(ByVal request As HttpRequestMessage, ByVal cancellationToken As CancellationToken) As Task(Of HttpResponseMessage)
			Dim statusCode As HttpStatusCode
			Dim token As String = ""
			' determine whether a jwt exists or not
			If Not TokenValidationHandler.TryRetrieveToken(request, token) Then
				statusCode = HttpStatusCode.Unauthorized
				Return MyBase.SendAsync(request, cancellationToken)
			End If

			Try
				Dim secretKey = ConfigurationManager.AppSettings("JWT_SECRET_KEY")
				Dim audienceToken = ConfigurationManager.AppSettings("JWT_AUDIENCE_TOKEN")
				Dim issuerToken = ConfigurationManager.AppSettings("JWT_ISSUER_TOKEN")
				Dim securityKey = New SymmetricSecurityKey(System.Text.Encoding.Default.GetBytes(secretKey))
				Dim securityToken As JwtSecurityToken = Nothing

				Dim tokenHandler As New System.IdentityModel.Tokens.Jwt.JwtSecurityTokenHandler()
				Dim validationParameters As New TokenValidationParameters() With {
					.ValidAudience = audienceToken,
					.ValidIssuer = issuerToken,
					.ValidateLifetime = True,
					.ValidateIssuerSigningKey = True,
					.LifetimeValidator = AddressOf Me.LifetimeValidator,
					.IssuerSigningKey = securityKey
				}

				' Extract and assign Current Principal and user
				Thread.CurrentPrincipal = tokenHandler.ValidateToken(token, validationParameters, securityToken)
				HttpContext.Current.User = tokenHandler.ValidateToken(token, validationParameters, securityToken)




				Return MyBase.SendAsync(request, cancellationToken)
			Catch ex As SecurityTokenValidationException
				statusCode = HttpStatusCode.Unauthorized
			Catch ex As Exception
				statusCode = HttpStatusCode.InternalServerError
			End Try



			Return Task(Of HttpResponseMessage).Factory.StartNew(Function() New HttpResponseMessage(statusCode)) 'equivalente de funcion lambda en VB.net

		End Function

		Public Function LifetimeValidator(ByVal notBefore As DateTime?, ByVal expires As DateTime?, ByVal securityToken As SecurityToken, ByVal validationParameters As TokenValidationParameters) As Boolean
			If (Not (expires) Is Nothing) Then
				If (DateTime.UtcNow < expires) Then
					Return True
				End If

			End If

			Return False
		End Function
	End Class
End Namespace
