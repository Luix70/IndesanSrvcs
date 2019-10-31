Imports System.Web.Http


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
			Dim customersFake = New String() {"customer-1", "customer-2", "customer-3"}
			Return Ok(customersFake)
		End Function
	End Class
End Namespace