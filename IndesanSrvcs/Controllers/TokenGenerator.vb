Imports System
Imports System.Configuration
Imports System.Security.Claims
Imports Microsoft.IdentityModel.Tokens
Imports IndesanSrvcs.Models
Imports System.Reflection

Namespace Controllers
	Friend Module TokenGenerator
		Public Function GenerateTokenJwt(username As String, cr As Credencial) As String
			Dim secretKey As String = ConfigurationManager.AppSettings("JWT_SECRET_KEY")
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
	End Module

End Namespace
