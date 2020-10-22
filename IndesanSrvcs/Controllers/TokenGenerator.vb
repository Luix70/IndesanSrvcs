Imports System
Imports System.Configuration
Imports System.Security.Claims
Imports Microsoft.IdentityModel.Tokens
Imports IndesanSrvcs.Models
Imports System.Reflection
Imports System.IdentityModel.Tokens.Jwt

Namespace Controllers
	Friend Module TokenGenerator
		Public Function GenerateTokenJwt(username As String, cr As Credencial) As String
			Dim secretKey As String = Environment.GetEnvironmentVariable("JWT_SECRET_KEY")
			Dim audienceToken As String = ConfigurationManager.AppSettings("JWT_AUDIENCE_TOKEN")
			Dim issuerToken As String = ConfigurationManager.AppSettings("JWT_ISSUER_TOKEN")
			Dim expiretime As String = ConfigurationManager.AppSettings("JWT_EXPIRE_MINUTES")

			Dim securityKey As New SymmetricSecurityKey(System.Text.Encoding.Default.GetBytes(secretKey))
			Dim signingCredentials As New SigningCredentials(securityKey, SecurityAlgorithms.HmacSha256Signature)

			' create a claimsIdentity

			Dim claimsIdentity As ClaimsIdentity = New ClaimsIdentity({New Claim(ClaimTypes.Name, username)})
			Dim tokenHandler As New System.IdentityModel.Tokens.Jwt.JwtSecurityTokenHandler()
			Dim jwtSecurityToken = tokenHandler.CreateJwtSecurityToken(
				audience:=audienceToken,
				issuer:=issuerToken,
				subject:=claimsIdentity,
				notBefore:=DateTime.UtcNow,
				expires:=DateTime.UtcNow.AddMinutes(Convert.ToInt32(expiretime)),
				signingCredentials:=signingCredentials)

			'Añadimos datos al payload






			Dim info() As PropertyInfo = cr.GetType().GetProperties()

			For Each prop As PropertyInfo In info
				jwtSecurityToken.Payload(prop.Name) = prop.GetValue(cr, Nothing)

			Next

			Dim jwtTokenString As String = tokenHandler.WriteToken(jwtSecurityToken)

			Return jwtTokenString

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

				Dim secretKey = Environment.GetEnvironmentVariable("JWT_SECRET_KEY")
				Dim audienceToken = ConfigurationManager.AppSettings("JWT_AUDIENCE_TOKEN")
				Dim issuerToken = ConfigurationManager.AppSettings("JWT_ISSUER_TOKEN")
				Dim securityKey = New SymmetricSecurityKey(System.Text.Encoding.Default.GetBytes(secretKey))
				Dim securityToken As JwtSecurityToken = Nothing

				Dim validationParameters As New TokenValidationParameters() With {
						.ValidAudience = audienceToken,
						.ValidIssuer = issuerToken,
						.ValidateLifetime = True,
						.ValidateIssuerSigningKey = True,
						.LifetimeValidator = AddressOf LifetimeValidator,
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

	End Module

End Namespace
