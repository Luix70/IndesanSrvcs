Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web
Imports System.Web.Services
Imports System.Web.Script.Services
Imports System.ComponentModel
Imports System.Web.Script.Serialization
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Web.Services.Protocols
Imports System.Web.Hosting
Imports System.IO
Imports PdfSharp.Pdf
Imports PdfSharp.Drawing
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Web.Http


' Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente.

<System.Web.Services.WebService(Namespace:="http://servicios.indesan.com/")>
<System.Web.Script.Services.ScriptService>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class JData
	Inherits System.Web.Services.WebService

	Private Const StrResponseNew As String = "Hello World"
	Public SoapHeader As TokenAuthUser

	Dim cache As New MemoryCacher

	<WebMethod()>
	<ScriptMethod(ResponseFormat:=ResponseFormat.Json)>
	Public Sub JPedidos()

		Dim parCodRep As String = HttpContext.Current.Request.Params("cr")
		Dim parCodCli As String = HttpContext.Current.Request.Params("cc")
		Dim parCodigodoc As String = HttpContext.Current.Request.Params("cd")
		Dim parTipodoc As String = HttpContext.Current.Request.Params("td")
		Dim Json As String

		Dim js As New JavaScriptSerializer()

		js.MaxJsonLength = 50000000



		Dim qj As New QueryJson()

		Dim resultados As Object

		resultados = cache.GetValue("datos")

		If IsNothing(resultados) Then

			resultados = qj.GenerarJson(Nothing, parCodCli, parTipodoc, parCodigodoc)
			cache.Add("datos", resultados, New DateTimeOffset(DateTime.Now.AddMinutes(ConfigurationManager.AppSettings("CACHE_PERSISTENCE"))))


		End If

		Json = js.Serialize(resultados)

		HttpContext.Current.Response.Clear()
		HttpContext.Current.Response.AppendHeader("Access-Control-Allow-Origin", "*")
		HttpContext.Current.Response.ContentType = "application/json;charset=UTF-8"
		HttpContext.Current.Response.Write(Json)


	End Sub

	<WebMethod()>
	<ScriptMethod(ResponseFormat:=ResponseFormat.Json)>
	Public Sub JScans()

		Dim parCodigodoc As String = HttpContext.Current.Request.Params("cd")
		Dim parTipodoc As String = HttpContext.Current.Request.Params("td")

		Dim Json As String

		Dim js As New JavaScriptSerializer()

		js.MaxJsonLength = 50000000


		If parCodigodoc = "" Or parTipodoc = "" Then


			Json = "{error: 'no se han suministrado los parámetros necesarios'}"

			HttpContext.Current.Response.Clear()
			HttpContext.Current.Response.AppendHeader("Access-Control-Allow-Origin", "*")
			HttpContext.Current.Response.ContentType = "application/json;charset=UTF-8"
			HttpContext.Current.Response.Write(Json)
		Else

			Dim qj As New QueryJson()

			Dim resultados As Object

			resultados = qj.GenerarScanJson(parTipodoc, parCodigodoc)

			Json = js.Serialize(resultados)

			HttpContext.Current.Response.Clear()
			HttpContext.Current.Response.AppendHeader("Access-Control-Allow-Origin", "*")
			HttpContext.Current.Response.ContentType = "application/json;charset=UTF-8"
			HttpContext.Current.Response.Write(Json)

		End If



	End Sub

	<WebMethod()>
	<ScriptMethod(ResponseFormat:=ResponseFormat.Json)>
	Public Sub JTransferScan()

		Dim parRuta As String = HttpContext.Current.Request.Params("ruta")
		Dim parTipodoc As String = HttpContext.Current.Request.Params("td")
		Dim parCodigoc As String = HttpContext.Current.Request.Params("cd")
		Dim parTipoArchivo As String = HttpContext.Current.Request.Params("tipoArchivo")

		Dim Json As String


		Dim nombreDocumento As String = parTipodoc + parCodigoc + "-" + parTipoArchivo.Split(" ")(0) + ".pdf"

		If parRuta = "" Or parRuta = "" Then
			Dim filePath As String = "no se han suminsitrado los parámetros necesarios"

			Json = "{error: " + filePath + "}"

			HttpContext.Current.Response.Clear()
			HttpContext.Current.Response.AppendHeader("Access-Control-Allow-Origin", "*")
			HttpContext.Current.Response.ContentType = "application/json;charset=UTF-8"
			HttpContext.Current.Response.Write(Json)
		Else



			Dim filePath As String = ConfigurationManager.AppSettings("LOCAL_VAULT") + parRuta

			If parRuta.Substring(parRuta.Length - 3, 3) = "tif" Then

				Dim uri As New Uri(filePath)
				If uri.IsFile Then
					Dim filename As String = System.IO.Path.GetFileName(uri.LocalPath)

					filePath = tiff2PDF(filePath)
				End If
			End If

			With HttpContext.Current.Response
				.Clear()
				.AddHeader("Content-Disposition", "attachment;filename=" + nombreDocumento)
				.AppendHeader("Access-Control-Allow-Origin", "*")
				.BinaryWrite(System.IO.File.ReadAllBytes(filePath))

				.Flush()
				.End()

			End With



		End If





	End Sub
	<WebMethod()>
	<Authorize>
	<HttpPost>
	<ScriptMethod(ResponseFormat:=ResponseFormat.Json)>
	Public Function JAddScan(<FromBody> nombre) As String



		Return ("{OK : " & nombre & "}")


	End Function
	<WebMethod()>
	<ScriptMethod(ResponseFormat:=ResponseFormat.Json)>
	Public Sub JOps()

		Dim parUser As String = HttpContext.Current.Request.Params("user")
		Dim parPassWord As String = HttpContext.Current.Request.Params("password")

		Dim qj As New QueryJson()

		Dim resultados As String

		resultados = qj.GenerarJson2(parUser, parPassWord)

		HttpContext.Current.Response.Clear()
		HttpContext.Current.Response.AppendHeader("Access-Control-Allow-Origin", "*")
		' HttpContext.Current.Response.AppendHeader("Content-Encoding", "identity")
		HttpContext.Current.Response.ContentType = "application/json;charset=UTF-8"
		HttpContext.Current.Response.Write(resultados)


	End Sub

	<WebMethod()>
	<ScriptMethod(ResponseFormat:=ResponseFormat.Json)>
	Public Sub JSaluda()

		Dim js As New JavaScriptSerializer
		Dim saludo As New atomObject(StrResponseNew)
		HttpContext.Current.Response.Clear()
		HttpContext.Current.Response.AppendHeader("Access-Control-Allow-Origin", "*")
		' HttpContext.Current.Response.AppendHeader("Content-Encoding", "identity")
		HttpContext.Current.Response.ContentType = "application/json;charset=UTF-8"
		HttpContext.Current.Response.Write(js.Serialize(saludo))



	End Sub
	'<SoapHeader("user")>
	<WebMethod()>
	<ScriptMethod(ResponseFormat:=ResponseFormat.Json)>
	Public Sub JAuth()

		Dim parUser As String = HttpContext.Current.Request.Params("user")
		Dim parPassWord As String = HttpContext.Current.Request.Params("password")
		Dim strRespuesta As String = "Usuario inválido"
		Dim usr As New AuthUser
		usr.Username = parUser
		usr.Password = parPassWord
		Dim NombreUsuario As String = ""

		If usr.isValid(NombreUsuario) Then

			strRespuesta = "Bienvenido, " + NombreUsuario

		End If

		Dim js As New JavaScriptSerializer
		Dim saludo As New atomObject(strRespuesta)
		HttpContext.Current.Response.Clear()
		HttpContext.Current.Response.AppendHeader("Access-Control-Allow-Origin", "*")
		' HttpContext.Current.Response.AppendHeader("Content-Encoding", "identity")
		HttpContext.Current.Response.ContentType = "application/json;charset=UTF-8"
		HttpContext.Current.Response.Write(js.Serialize(saludo))

	End Sub

	<WebMethod()>
	<ScriptMethod(ResponseFormat:=ResponseFormat.Json)>
	<SoapHeader("SoapHeader")>
	Public Function JTokenAuth() As String

		If SoapHeader Is Nothing Then
			Return "Necesitas autentificarte antes de acceder a este método. Llama a AuthenticationMethod()"
		End If

		If Not SoapHeader.isUserCredentialsValid(SoapHeader) Then
			Return "Necesitas autentificarte antes de acceder a este método. Llama a AuthenticationMethod()"
		Else
			Return "Hola, " + HttpRuntime.Cache(SoapHeader.AuthenticationToken)
		End If
	End Function

	<WebMethod()>
	<ScriptMethod(ResponseFormat:=ResponseFormat.Json)>
	<SoapHeader("SoapHeader")>
	Public Function AuthenticationMethod() As String
		If SoapHeader Is Nothing Then
			Return "Por favor, suministra un SoapHeader"
		End If

		If String.IsNullOrEmpty(SoapHeader.Username) OrElse String.IsNullOrEmpty(SoapHeader.Password) Then
			Return "Por favor, suministra nombre de ususario y contraseña"
		End If

		If Not SoapHeader.isUserCredentialsValid(SoapHeader.Username, SoapHeader.Password) Then

			Return "Usuario o contraseña inválidos"

		End If

		Dim token As String = Guid.NewGuid().ToString()

		HttpRuntime.Cache.Add(token,
							  SoapHeader.Username,
							  Nothing,
							  System.Web.Caching.Cache.NoAbsoluteExpiration,
							  TimeSpan.FromMinutes(30),
							  System.Web.Caching.CacheItemPriority.NotRemovable,
							  Nothing)
		Return token

	End Function
	Private Class atomObject
		Private strResponse As String
		Public Sub New(ByVal strResponseNew As String)
			Me.strResponse = strResponseNew
		End Sub
		Public Property response As String
			Get
				Return strResponse
			End Get


			Set(value As String)
				strResponse = value
			End Set
		End Property
	End Class

	Public Function tiff2PDF(ByVal fileName As String) As String
		Dim doc As PdfDocument = New PdfDocument
		Dim tiff As New TiffImageSplitter()
		Dim nfilename As String = "error"
		Dim pageCount As Integer = tiff.getPageCount(fileName)

		For i = 0 To pageCount - 1
			Dim page As PdfPage = New PdfPage
			Dim tiffImg As Image = tiff.getTiffImage(fileName, i)
			Dim img As XImage = XImage.FromGdiPlusImage(tiffImg)
			page.Width = img.PointWidth
			page.Height = img.PointHeight
			doc.Pages.Add(page)
			Dim xgr As XGraphics = XGraphics.FromPdfPage(doc.Pages(i))
			xgr.DrawImage(img, 0, 0)
		Next
		nfilename = fileName.Replace(".tif", ".pdf")

		Try
			doc.Save(nfilename)
		Catch ex As Exception
			nfilename = nfilename.Replace(".pdf", CStr(Int(Rnd(100) * 100)) & ".pdf")
			doc.Save(nfilename)
		End Try

		doc.Close()

		Return nfilename

	End Function


End Class

Public Class AuthUser
	Inherits System.Web.Services.Protocols.SoapHeader
	Private strUsername As String
	Public Property Username() As String
		Get
			Return strUsername
		End Get
		Set(ByVal value As String)
			strUsername = value
		End Set
	End Property

	Private newPassword As String
	Public Property Password() As String
		Get
			Return newPassword
		End Get
		Set(ByVal value As String)
			newPassword = value
		End Set
	End Property

	Public Function isValid(ByRef strNombreCliente As String) As Boolean
		'Realizar toda la lógica para determinar si el usuario es válido
		'(Consultar database etc)
		Dim count As Int16 = 0
		Dim config As String = ConfigurationManager.ConnectionStrings("myCon").ConnectionString

		Dim strCadenaConsulta As String

		strCadenaConsulta = "Select * from CLIENTES_RST where nombreusuario ='" + Username + "' AND contras ='" + Password + "';"
		Dim Cons As New OleDb.OleDbConnection

		Cons.ConnectionString = config

		Cons.Open()

		Dim ndt As New DataTable
		Dim dad As New OleDb.OleDbDataAdapter(strCadenaConsulta, Cons)

		Try
			dad.Fill(ndt)
		Catch ex As Exception
			MsgBox(ex.Message)
		End Try

		count = ndt.Rows.Count

		Dim result As Boolean

		If count > 0 Then

			result = True
			strNombreCliente = ndt.Rows(0).Item("rzs")
		Else
			result = False
			strNombreCliente = ""
		End If


		Cons.Close()
		Cons = Nothing
		Return result

	End Function
End Class

Public Class TokenAuthUser
	Inherits System.Web.Services.Protocols.SoapHeader
	Private strUsername As String
	Public Property Username() As String
		Get
			Return strUsername
		End Get
		Set(ByVal value As String)
			strUsername = value
		End Set
	End Property

	Private newPassword As String
	Public Property Password() As String
		Get
			Return newPassword
		End Get
		Set(ByVal value As String)
			newPassword = value
		End Set
	End Property
	Private newAuthenticationToken As String
	Public Property AuthenticationToken() As String
		Get
			Return newAuthenticationToken
		End Get
		Set(ByVal value As String)
			newAuthenticationToken = value
		End Set
	End Property
	Public Function isUserCredentialsValid(username As String, password As String) As Boolean
		'Realizar toda la lógica para determinar si el usuario es válido
		'(Consultar database etc)
		Dim count As Int16 = 0
		Dim config As String = ConfigurationManager.ConnectionStrings("myCon").ConnectionString

		Dim strCadenaConsulta As String

		strCadenaConsulta = "Select * from CLIENTES_RST where nombreusuario ='" + username + "' AND contras ='" + password + "';"
		Dim Cons As New OleDb.OleDbConnection

		Cons.ConnectionString = config

		Cons.Open()

		Dim ndt As New DataTable
		Dim dad As New OleDb.OleDbDataAdapter(strCadenaConsulta, Cons)

		Try
			dad.Fill(ndt)
		Catch ex As Exception
			MsgBox(ex.Message)
		End Try

		count = ndt.Rows.Count

		Dim result As Boolean

		If count > 0 Then

			result = True

		Else
			result = False

		End If


		Cons.Close()
		Cons = Nothing
		Return result

	End Function

	Public Function isUserCredentialsValid(SoapHeader As TokenAuthUser) As Boolean
		If SoapHeader Is Nothing Then
			Return False
		End If

		If Not String.IsNullOrEmpty(SoapHeader.AuthenticationToken) Then
			Return HttpRuntime.Cache(SoapHeader.AuthenticationToken) IsNot Nothing
		End If

		Return False

	End Function
End Class

Public Class TiffImageSplitter
	Public Function getPageCount(ByVal fileName As String) As Integer
		Dim pageCount = -1

		Try
			Dim img As Image = Bitmap.FromFile(fileName)
			pageCount = img.GetFrameCount(FrameDimension.Page)
			img.Dispose()
		Catch ex As Exception
			pageCount = 0
		End Try

		Return pageCount
	End Function

	Public Function getPageCount(ByVal img As Image) As Integer
		Dim pageCount = -1

		Try
			pageCount = img.GetFrameCount(FrameDimension.Page)
		Catch ex As Exception
			pageCount = 0
		End Try

		Return pageCount
	End Function

	Public Function getTiffImage(ByVal sourceFile As String, ByVal pageNumber As Integer) As Image
		Dim returnImage As Image = Nothing

		Try
			Dim sourceIamge As Image = Bitmap.FromFile(sourceFile)
			returnImage = getTiffImage(sourceIamge, pageNumber)
			sourceIamge.Dispose()
		Catch ex As Exception
			returnImage = Nothing
		End Try

		Return returnImage
	End Function

	Public Function getTiffImage(ByVal sourceImage As Image, ByVal pageNumber As Integer) As Image
		Dim ms As MemoryStream = Nothing
		Dim returnImage As Image = Nothing

		Try
			ms = New MemoryStream
			Dim objGuid As Guid = sourceImage.FrameDimensionsList(0)
			Dim objDimension As FrameDimension = New FrameDimension(objGuid)
			sourceImage.SelectActiveFrame(objDimension, pageNumber)
			sourceImage.Save(ms, ImageFormat.Tiff)
			returnImage = Image.FromStream(ms)
		Catch ex As Exception
			returnImage = Nothing
		End Try

		Return returnImage
	End Function
End Class
