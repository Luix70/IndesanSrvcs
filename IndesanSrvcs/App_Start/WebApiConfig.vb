Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web.Http
Imports System.Web.Http.Cors
Imports IndesanSrvcs.Controllers
Public Module WebApiConfig
	Public Sub Register(ByVal config As HttpConfiguration)
		' Configuración y servicios de API web
		Dim corsAttr As New EnableCorsAttribute("*", "*", "*")
		config.EnableCors(corsAttr)
		' Rutas de API web
		config.MapHttpAttributeRoutes()
		config.MessageHandlers.Add(New TokenValidationHandler())
		config.Routes.MapHttpRoute(
			name:="DefaultApi",
			routeTemplate:="api/{controller}/{id}",
			defaults:=New With {.id = RouteParameter.Optional}
		)
	End Sub
End Module
