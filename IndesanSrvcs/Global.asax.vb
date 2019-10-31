Imports System
Imports System.Web.Http
Imports System.Web.Http.Cors

Public Class WebApiApplication
	Inherits System.Web.HttpApplication

	Protected Sub Application_Start()
		GlobalConfiguration.Configure(AddressOf WebApiConfig.Register)

	End Sub

	Protected Sub Application_BeginRequest(sender As Object, e As EventArgs)
		If (HttpContext.Current.Request.HttpMethod = "OPTIONS") Then

			HttpContext.Current.Response.Flush()
		End If


	End Sub
End Class
