Imports System.Web.Script.Serialization
Imports System
Imports System.IO
Imports System.Text
Imports System.Configuration
Imports System.Globalization
Imports IndesanSrvcs.Models
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Threading
Imports System.Data.OleDb

Public Class QueryJson

	Shared strConexion As String = ConfigurationManager.ConnectionStrings("myCon").ConnectionString
	Shared dt As DataTable
	Shared intCurrentRow As Integer = 0
	Private Class Resultado
		Public consulta As String
		Public FechaConsulta As String
		Public representantes As New Collection
		Public totalRepresentantes As Integer
		Public Status As String
		Public errorcode As String
		Public Sub FiltrarRepresentante(repre As Long)

			'Return
			'TO DO: implementar el filtro


			Dim rep As Object

			For i = representantes.Count To 1 Step -1
				rep = representantes(i)

				If rep("codrep") <> repre Then
					representantes.Remove(i)
				End If
			Next


		End Sub
		Public Sub FiltrarCliente(codcl As String)


			Dim i As Integer

			Dim j As Integer



			For i = representantes.Count To 1 Step -1
				For j = representantes(i)("Clientes").Length - 1 To 0
					If representantes(i)("Clientes")(j)("codigo") <> codcl Then
						representantes(i)("Clientes")(j).Remove(j)
					End If
				Next
			Next


		End Sub
		Function ObtenerRepresentantes() As Integer

			'Poblar la coleccion de representantes y dar valor al totalRepresentantes
			Dim reps As Integer = 0
			Dim objRep As Representante

			Dim intRows As Integer = dt.Rows.Count


			While intRows > 0 And intCurrentRow < intRows

				reps += 1

				Dim codrep As Integer = dt.Rows(intCurrentRow).Item("codrep")
				Dim NombreRepresentante As String = dt.Rows(intCurrentRow).Item("nombre")
				objRep = New Representante
				objRep.codrep = codrep
				objRep.nombre = NombreRepresentante

				' El método obtenerClientes debe depositar el cursor del DbReader en el ultimo registro
				' correspondiente al representante actual
				objRep.totalClientes = objRep.obtenerClientes(codrep)

				representantes.Add(objRep, codrep.ToString)

			End While

			Return reps

		End Function
	End Class
	Private Class Scans
		Public consulta As String
		Public FechaConsulta As String
		Public Scanners As New Collection
		Public totalDocs As Integer
		Public Status As String
		Public errorcode As String

		Function ObtenerScans() As Integer

			'Poblar la coleccion de documentos y dar valor al totalRepresentantes
			Dim docs As Integer = 0
			Dim objDoc As Scanner

			Dim intRows As Integer = dt.Rows.Count


			While intRows > 0 And intCurrentRow < intRows

				docs += 1

				Dim numerador As Integer = dt.Rows(intCurrentRow).Item("numerador")
				Dim documento As String = dt.Rows(intCurrentRow).Item("documento")
				Dim tipodoc As String = dt.Rows(intCurrentRow).Item("tipodoc")
				Dim codigodoc As Integer = dt.Rows(intCurrentRow).Item("codigodoc")
				Dim Archivo As Integer = dt.Rows(intCurrentRow).Item("Archivo")
				Dim codTipo As Integer = dt.Rows(intCurrentRow).Item("codTipo")
				Dim TipoImagen As String = dt.Rows(intCurrentRow).Item("TipoImagen")
				Dim ruta As String = dt.Rows(intCurrentRow).Item("ruta")

				objDoc = New Scanner
				objDoc.numerador = numerador
				objDoc.documento = documento
				objDoc.tipodoc = tipodoc
				objDoc.codigodoc = codigodoc
				objDoc.Archivo = Archivo
				objDoc.codTipo = codTipo
				objDoc.TipoImagen = TipoImagen
				objDoc.ruta = ruta


				' El método obtenerClientes debe depositar el cursor del DbReader en el ultimo registro
				' correspondiente al representante actual


				Scanners.Add(objDoc, numerador.ToString)
				intCurrentRow += 1
			End While

			Return docs

		End Function
	End Class
	Private Class Scanner
		Public tipodoc As String
		Public codigodoc As Long
		Public documento As String
		Public ruta As String
		Public Archivo As Long
		Public numerador As Long
		Public TipoImagen As String
		Public codTipo As Long
	End Class
	Private Class Representante
		Public codrep As Long
		Public nombre As String
		Private _clientes As New Collection
		Public totalClientes As Integer

		Public Property Clientes As Collection
			Get
				Return _clientes
			End Get
			Set(value As Collection)
				_clientes = value
			End Set
		End Property


		Function obtenerClientes(codRepActual As Integer) As Integer

			Dim n_clientes As Integer = 0
			Dim objCliente As Cliente

			While intCurrentRow < dt.Rows.Count AndAlso (dt.Rows(intCurrentRow).Item("codrep") = codRepActual Or codRepActual = 0)

				n_clientes += 1
				objCliente = New Cliente

				objCliente.codigo = dt.Rows(intCurrentRow).Item("codigo")
				objCliente.rzs = dt.Rows(intCurrentRow).Item("rzs")
				If Not IsDBNull(dt.Rows(intCurrentRow).Item("nomComercial")) Then
					objCliente.nomComercial = dt.Rows(intCurrentRow).Item("nomComercial")
				Else
					objCliente.nomComercial = ""
				End If

				objCliente.poblacion = dt.Rows(intCurrentRow).Item("poblacion")

				objCliente.totalDocumentos = objCliente.obtenerDocumentos(objCliente.codigo)

				Clientes.Add(objCliente, objCliente.codigo.ToString)


			End While

			' devolvemos el cursor al ultimo registro del cliente o al ultimo de la tabla

			Return n_clientes

		End Function

	End Class

	Private Class Cliente
		Public codigo As String
		Public rzs As String
		Public nomComercial As String
		Public poblacion As String
		Public documentos As New Collection
		Public totalDocumentos As Integer

		Function obtenerDocumentos(codigo As String) As Integer

			Dim n_documentos As Integer = 0
			Dim objDoc As Documento

			While intCurrentRow < dt.Rows.Count AndAlso dt.Rows(intCurrentRow).Item("codigo") = codigo

				objDoc = New Documento
				n_documentos += 1

				objDoc.tipodoc = dt.Rows(intCurrentRow).Item("tipodoc")
				objDoc.codigodoc = dt.Rows(intCurrentRow).Item("codigodoc")
				objDoc.fechadoc = dt.Rows(intCurrentRow).Item("fecha1").ToString
				objDoc.fecha2 = dt.Rows(intCurrentRow).Item("fecha2").ToString
				objDoc.Importebruto = dt.Rows(intCurrentRow).Item("Bruto")
				If Not IsDBNull(dt.Rows(intCurrentRow).Item("FECHAPEDIDO")) Then
					objDoc.fechapedido = dt.Rows(intCurrentRow).Item("FECHAPEDIDO")
				Else
					objDoc.fechapedido = dt.Rows(intCurrentRow).Item("fecha1")

				End If
				If Not IsDBNull(dt.Rows(intCurrentRow).Item("fecha_confirmacion")) Then
					objDoc.fechaConfirmacion = dt.Rows(intCurrentRow).Item("fecha_confirmacion")

				End If
				If Not IsDBNull(dt.Rows(intCurrentRow).Item("emailConfirmacion")) Then
					objDoc.confirmadoA = dt.Rows(intCurrentRow).Item("emailConfirmacion")

				End If

				If IsDBNull(dt.Rows(intCurrentRow).Item("referencia")) Then
					objDoc.referencia = ""
				Else
					objDoc.referencia = dt.Rows(intCurrentRow).Item("referencia")
				End If

				objDoc.totalLineas = objDoc.obtenerLineas(objDoc.tipodoc, objDoc.codigodoc)

				documentos.Add(objDoc, objDoc.tipodoc & objDoc.codigodoc.ToString)



			End While

			Return n_documentos

		End Function

	End Class

	Private Class Documento
		Public tipodoc As String
		Public codigodoc As Long
		Public fechadoc As String
		Public fecha2 As String
		Public fechapedido As String
		Public fechaEntrega As String
		Public fechaConfirmacion As String
		Public confirmadoA As String
		Public Importebruto As Single
		Public referencia As String
		Public lineas As New Collection
		Public totalLineas As Integer
		Public primeraPintura As String
		Public ultimoEmbalaje As String

		Function obtenerLineas(tipodoc As String, codigodoc As Long) As Integer
			Dim n_lineas As Integer = 0
			Dim objLinea As Linea

			While intCurrentRow < dt.Rows.Count AndAlso dt.Rows(intCurrentRow).Item("tipodoc") = tipodoc AndAlso dt.Rows(intCurrentRow).Item("codigodoc") = codigodoc
				n_lineas += 1
				objLinea = New Linea

				objLinea.linea = dt.Rows(intCurrentRow).Item("linea")
				objLinea.coart = dt.Rows(intCurrentRow).Item("coart")

				objLinea.codagrupacion = dt.Rows(intCurrentRow).Item("codagrupacion")

				objLinea.Bultos = Convert.ToInt64(dt.Rows(intCurrentRow).Item("Bultos"))

				If IsDBNull(dt.Rows(intCurrentRow).Item("fechaped")) Then
					objLinea.fechapedido = ""
				Else
					objLinea.fechapedido = dt.Rows(intCurrentRow).Item("fechaped")
				End If

				If IsDBNull(dt.Rows(intCurrentRow).Item("fecha_entrada")) Then
					objLinea.fechapedido = ""
				Else
					objLinea.fechapedido = dt.Rows(intCurrentRow).Item("fecha_entrada")
				End If

				If dt.Rows(intCurrentRow).Item("codagrupacion").ToString.Contains("COLOR=") Then
					If IsDate(dt.Rows(intCurrentRow).Item("fecha_muestra")) Then
						objLinea.fecha_muestra = dt.Rows(intCurrentRow).Item("fecha_muestra")
						If IsDate(primeraPintura) Then

							If CDate(primeraPintura) >= CDate(objLinea.fecha_muestra) Then
								primeraPintura = CDate(objLinea.fecha_muestra)
							End If
						Else
							primeraPintura = CDate(objLinea.fecha_muestra)
						End If

					End If

				End If

				If objLinea.Bultos > 0 Then
					If IsDate(dt.Rows(intCurrentRow).Item("fecha_emb")) Then
						objLinea.fecha_emb = dt.Rows(intCurrentRow).Item("fecha_emb")
						If IsDate(ultimoEmbalaje) Then
							If CDate(ultimoEmbalaje) <= CDate(objLinea.fecha_emb) Then
								ultimoEmbalaje = CDate(objLinea.fecha_emb)
							End If

						Else
							ultimoEmbalaje = CDate(objLinea.fecha_emb)
						End If
					Else
						ultimoEmbalaje = "31/12/2100"
					End If

				End If

				If IsDate(dt.Rows(intCurrentRow).Item("fechaalbaran")) Then
					objLinea.fechaalbaran = dt.Rows(intCurrentRow).Item("fechaalbaran")
					If IsDate(fechaEntrega) Then
						If CDate(fechaEntrega) <= CDate(objLinea.fechaalbaran) Then
							fechaEntrega = CDate(objLinea.fechaalbaran)
						End If
					Else
						fechaEntrega = CDate(objLinea.fechaalbaran)
					End If

				End If

				If IsDate(dt.Rows(intCurrentRow).Item("fechaped")) Then
					objLinea.fechapedido = dt.Rows(intCurrentRow).Item("fechaped")
					'	If IsDate(fechapedido) Then
					'		If CDate(fechapedido) >= CDate(objLinea.fechapedido) Then
					'			fechapedido = CDate(objLinea.fechapedido)
					'		End If
					'	Else
					'		fechapedido = CDate(objLinea.fechapedido)
					'	End If

				End If


				Try
					If IsDBNull(dt.Rows(intCurrentRow).Item("cantidad")) Then
						objLinea.cantidad = 0
					Else
						objLinea.cantidad = dt.Rows(intCurrentRow).Item("cantidad")
					End If



				Catch e As Exception
					objLinea.cantidad = 0
				End Try

				Try
					objLinea.fecha_entrega_agencia = dt.Rows(intCurrentRow).Item("fecha_entrega_agencia")

				Catch ex As Exception

				End Try

				Try

					objLinea.fecha_entrada = dt.Rows(intCurrentRow).Item("fecha_entrada")

				Catch ex As Exception

				End Try


				Try

					objLinea.albaran = dt.Rows(intCurrentRow).Item("albaran")
				Catch ex As Exception

				End Try


				objLinea.descripcion = dt.Rows(intCurrentRow).Item("descripcion")

				If IsDBNull(dt.Rows(intCurrentRow).Item("ref_linea")) Then
					objLinea.ref_linea = ""
				Else
					objLinea.ref_linea = dt.Rows(intCurrentRow).Item("ref_linea")
				End If

				If IsDBNull(dt.Rows(intCurrentRow).Item("referencia2")) Then
					objLinea.referencia = ""
				Else
					objLinea.referencia = dt.Rows(intCurrentRow).Item("referencia2")
				End If

				objLinea.precio = dt.Rows(intCurrentRow).Item("precio")
				objLinea.pedido = dt.Rows(intCurrentRow).Item("pedido")
				objLinea.codLinea = dt.Rows(intCurrentRow).Item("codLinea")


				lineas.Add(objLinea, n_lineas.ToString)

				intCurrentRow += 1

			End While

			Return n_lineas



		End Function

	End Class
	Private Class Linea
		Public linea As Integer
		Public coart As String
		Public codagrupacion As String
		Public cantidad As Single
		Public descripcion As String
		Public ref_linea As String
		Public precio As Single
		Public pedido As String
		Public albaran As String
		Public codLinea As String
		Public fechapedido As String
		Public referencia As String
		Public fecha_muestra As String
		Public fecha_emb As String
		Public fecha_entrega_agencia As String
		Public fecha_entrada As String
		Public fechaalbaran As String
		Public Bultos As Long

	End Class

	Public Function RegistrarMensaje(msg As JObject) As JObject

		Dim strCadenaConsulta As String

		Dim txt As String = "Mensaje Recibido via WEB:" & vbCrLf & "============================" & vbCrLf & msg("mensaje").ToString() & vbCrLf & "============================" & vbCrLf & "Recibido de: " & msg("nombre").ToString() & vbCrLf & "email: " & msg("email").ToString() & vbCrLf & "Teléfono de: " & msg("telefono").ToString()
		Dim dateMessage As String = Date.Now.ToShortDateString & " " & Date.Now.ToLongTimeString

		strCadenaConsulta = "INSERT INTO NotasCliente ( Nota, fecha, Usuario, avisar, fecha_aviso, UsuarioAviso ) Select '" & txt & "' As Nota, '" & dateMessage & "' As fecha, 'WEB' As Usuario, True As avisar, '" & dateMessage & "' As fecha_aviso, 'Irene' As UsuarioAviso"

		Try
			Dim i As Integer
			Dim Cons As New OleDb.OleDbConnection
			Cons.ConnectionString = strConexion
			Cons.Open()
			Dim cmd As New OleDbCommand(strCadenaConsulta, Cons)
			i = cmd.ExecuteNonQuery()

			Cons.Close()
			Cons = Nothing
			msg.Add("registrado", Date.Now.ToShortDateString & " " & Date.Now.ToLongTimeString)


		Catch ex As Exception
			msg.Add("error: " & ex.ToString)
		End Try

		Return msg

	End Function
	Public Function GenerarJson(parCodRep As String, parCodCli As String, parTipodoc As String, parCodigodoc As String) As Object



		dt = New DataTable
		intCurrentRow = 0



		Dim strCadenaConsulta As String
		strCadenaConsulta = CadenaConsulta(parCodRep, parCodCli, parTipodoc, parCodigodoc)

		Dim Cons As New OleDb.OleDbConnection
		Cons.ConnectionString = strConexion
		Cons.Open()

		Using dad As New OleDb.OleDbDataAdapter(strCadenaConsulta, Cons)

			dad.Fill(dt)

		End Using

		Cons.Close()
		Cons = Nothing

		Dim res As New Resultado
		res.consulta = strCadenaConsulta

		Dim culture As New CultureInfo("en-US")
		res.FechaConsulta = Now.ToString("o", culture)
		res.totalRepresentantes = res.ObtenerRepresentantes()


		dt = Nothing


		Return res

	End Function
	Public Function GenerarScanJson(parTipodoc As String, parCodigodoc As String) As Object



		dt = New DataTable
		intCurrentRow = 0



		Dim strCadenaConsulta As String
		strCadenaConsulta = "SELECT Scan_docs_imgs.tipodoc, Scan_docs_imgs.codigodoc, Scan_imgs.documento, Scan_imgs.Archivo, Scan_imgs.numerador, scan_tipos_imagenes.TipoImagen, scan_tipos_imagenes.codTipo, Scan_Archivos.Nombre AS ruta
FROM Scan_Archivos INNER JOIN ((scan_tipos_imagenes INNER JOIN Scan_imgs ON scan_tipos_imagenes.codTipo = Scan_imgs.tipoImagen) INNER JOIN Scan_docs_imgs ON Scan_imgs.numerador = Scan_docs_imgs.Cod_img) ON Scan_Archivos.Archivo = Scan_imgs.Archivo
								WHERE (((Scan_docs_imgs.tipodoc)='" & parTipodoc & "') AND ((Scan_docs_imgs.codigodoc)=" & parCodigodoc & ") AND ((Scan_imgs.Archivo)<>2));"

		Dim Cons As New OleDb.OleDbConnection
		Cons.ConnectionString = strConexion
		Cons.Open()

		Using dad As New OleDb.OleDbDataAdapter(strCadenaConsulta, Cons)

			dad.Fill(dt)

		End Using

		Cons.Close()
		Cons = Nothing

		Dim res As New Scans
		res.consulta = strCadenaConsulta
		res.totalDocs = res.ObtenerScans

		Dim culture As New CultureInfo("en-US")
		res.FechaConsulta = Now.ToString("o", culture)


		dt = Nothing


		Return res

	End Function
	Public Function GenerarJson2(parUsuario As String, parPassword As String) As String

		Dim Json As String = ""
		Dim culture As New CultureInfo("en-US")
		intCurrentRow = 0
		Dim js As JavaScriptSerializer
		js = New JavaScriptSerializer()
		js.MaxJsonLength = 100000000

		Dim strCadenaConsulta As String

		strCadenaConsulta = "SELECT * FROM Credenciales_rst WHERE Credenciales_rst.Email ='" & parUsuario & "';"
		Dim Cons As New OleDb.OleDbConnection
		Cons.ConnectionString = strConexion
		Cons.Open()

		dt = New DataTable

		Using dad As New OleDb.OleDbDataAdapter(strCadenaConsulta, Cons)

			Try
				dad.Fill(dt)
			Catch ex As Exception
				MsgBox(ex.Message)
			End Try


		End Using

		Cons.Close()
		Cons = Nothing

		Dim res As New Resultado

		If dt.Rows.Count = 0 Then
			res.consulta = "nulo"
			res.Status = "noOK"
			res.errorcode = "Usuario Inexistente"
			res.FechaConsulta = Now.ToString("o", culture)
			Json = js.Serialize(res)
			Return Json

		Else
			If dt.Rows(0).Field(Of String)("Password") <> parPassword Then
				res.consulta = "nulo"
				res.Status = "noOK"
				res.errorcode = "Contraseña errónea"
				res.consulta = strCadenaConsulta


				res.FechaConsulta = Now.ToString("o", culture)
				res.totalRepresentantes = res.ObtenerRepresentantes()
				Json = js.Serialize(res)
				Return Json
			Else
				'Todo correcto
				res.consulta = CadenaConsulta2("{tipoentidad:'" & dt.Rows(0).Field(Of String)("TipoEntidad") & "', AccesoCli:'" & dt.Rows(0).Field(Of String)("AccesoCli") & "', AccesoRep: '" & dt.Rows(0).Field(Of String)("AccesoRep") & "'}")
				res.Status = "OK"
				res.FechaConsulta = Now.ToString("o", culture)
				res.errorcode = ""
				'primero vamos a ver si existe una copia en el cache reciente, para evitar llamadas a la base de datos
				If File.Exists(HttpContext.Current.Server.MapPath(".") & "\JSON\result.json") And (res.consulta = "SELECT * FROM listadoOperaciones where (codrep > 0);") Then

					Dim strJson As String
					Using sr As New StreamReader(HttpContext.Current.Server.MapPath(".") & "\JSON\result.json")

						' Read the stream to a string and write the string to the console.
						strJson = sr.ReadToEnd()

					End Using

					res = js.Deserialize(Of Resultado)(strJson)

					If DateDiff(DateInterval.Minute, Convert.ToDateTime(res.FechaConsulta), Date.Now()) < 60 Then
						Return strJson
					Else ' file is old

						Json = ConsultarDB(res.consulta)
						Return Json

					End If
				Else 'file doesn't exists

					Json = ConsultarDB(res.consulta)
					Return Json

				End If
			End If
		End If

	End Function
	Public Function ObtenerDocs(repre As String, cliente As String, tipoEntidad As String) As String

		Dim Json As String = ""
		Dim Job As New JObject

		Dim culture As New CultureInfo("en-US")
		intCurrentRow = 0

		'Dim js As New JavaScriptSerializer()
		'js.MaxJsonLength = 50000000



		Dim res As New Resultado
		res.consulta = "SELECT * FROM listadoOperaciones;"
		res.Status = "OK"
		res.FechaConsulta = Now.ToString("o", culture)
		res.errorcode = ""
		'primero vamos a ver si existe una copia en el cache reciente, para evitar llamadas a la base de datos
		If Not File.Exists(HttpContext.Current.Server.MapPath(".") & "\JSON\result.json") Then


			Job = JObject.Parse(ConsultarDB(res.consulta))

		End If

		Using sr As New StreamReader(HttpContext.Current.Server.MapPath(".") & "\JSON\result.json")

			' Read the stream to a string and write the string to the console.

			Job = JObject.Parse(sr.ReadToEnd())

		End Using

		'res = JsonConvert.DeserializeObject(Of Resultado)(Json)

		If DateDiff(DateInterval.Minute, Convert.ToDateTime(Job("FechaConsulta")), Date.Now()) > 60 Then


			Job = JObject.Parse(ConsultarDB(res.consulta))

		End If



		If repre = "*" And cliente = "*" Then

			Return Job.ToString
		Else

			If repre <> "*" Then
				Dim arrRepres As Array = repre.Split(",")

				Dim JRep As JArray = Job("representantes")
				Dim nJrep As New JArray

				For Each r As JObject In JRep
					For r_index = 0 To arrRepres.Length - 1
						If r("codrep") = arrRepres(r_index) Then
							nJrep.Add(r)
							Exit For
						End If
					Next
				Next

				Job("representantes") = nJrep

				Job("totalRepresentantes") = arrRepres.Length
				Return Job.ToString

			Else
				If cliente <> "*" Then
					Dim JRep As JArray = Job("representantes")
					Dim nJrep As New JArray
					Dim JCli As New JArray
					Dim nJCli As New JArray
					For Each r As JObject In JRep
						JCli = r("Clientes")
						For Each c As JObject In JCli
							If c("codigo") = cliente Then
								nJCli.Add(c)
								r("Clientes") = nJCli
								r("totalClientes") = 1
								nJrep.Add(r)
								Job("representantes") = nJrep
								Job("totalRepresentantes") = 1
								Return Job.ToString

							End If
						Next
					Next

					Job("representantes") = nJrep 'Por si está vacio

					Return Job.ToString
				Else
					Return "{}"
				End If
			End If



		End If

	End Function
	Public Function GeneraCredencial(Username As String, Password As String) As Credencial
		Dim nc As New Credencial
		Dim strConsulta As String
		Dim cdt As New DataTable
		Dim Cons As New OleDb.OleDbConnection

		strConsulta = "SELECT Credenciales_rst.* FROM Credenciales_rst WHERE (((Credenciales_rst.email)='" & Username & "'));"
		Cons.ConnectionString = strConexion
		Cons.Open()

		Using dad As New OleDb.OleDbDataAdapter(strConsulta, Cons)

			dad.Fill(cdt)

		End Using


		Cons.Close()
		Cons = Nothing

		If cdt.Rows.Count = 1 Then
			If cdt.Rows(0).Item("Password") = Password Then
				nc.IdCredencial = cdt.Rows(0).Item("IdCredencial")
				nc.TipoEntidad = cdt.Rows(0).Item("TipoEntidad")
				nc.NombreUsuario = cdt.Rows(0).Item("NombreUsuario")
				nc.Password = "OK"
				nc.AccesoCli = cdt.Rows(0).Item("AccesoCli")
				nc.AccesoRep = cdt.Rows(0).Item("AccesoRep")
				nc.Email = cdt.Rows(0).Item("Email")
				nc.Idioma = cdt.Rows(0).Item("Idioma")
			Else
				nc.IdCredencial = -1
				nc.Email = cdt.Rows(0).Item("Email")
				nc.Password = "NOOK"
			End If

		Else
			nc.IdCredencial = -1
			nc.Email = "NOOK"
		End If




		Return nc

	End Function
	Public Function VerificarCandidato(Username As String, Cif As String, Password As String) As String

		Dim strConsulta As String
		Dim cdt As New DataTable
		Dim Cons As New OleDb.OleDbConnection

		strConsulta = "SELECT Clientes_rst.email, Clientes_rst.cif FROM Clientes_rst WHERE (((Clientes_rst.cif)='" & Cif & "'));"
		Cons.ConnectionString = strConexion
		Cons.Open()

		Using dad As New OleDb.OleDbDataAdapter(strConsulta, Cons)

			dad.Fill(cdt)

		End Using


		Cons.Close()
		Cons = Nothing

		If cdt.Rows.Count = 1 Then
			If cdt.Rows(0).Item("email") = Username Then

				Return "OK"
			Else
				Return "email_invalido"
			End If

		Else
			Return "Usuario_desconocido"
		End If





	End Function

	Function ConsultarDB(sql As String) As String
		' hay que leerlo de nuevo
		Dim Res = New Resultado
		Dim json As String = ""
		dt = New DataTable
		'Consultar de nuevo para llenar dt
		Dim Cons2 As New OleDb.OleDbConnection
		Cons2.ConnectionString = strConexion
		Cons2.Open()
		Using dad As New OleDb.OleDbDataAdapter(sql, Cons2)

			Try
				dad.Fill(dt)
			Catch ex As Exception
				MsgBox(ex.Message)
			End Try

		End Using

		Cons2.Close()
		Cons2 = Nothing


		Res.Status = "OK"
		Res.errorcode = ""
		Res.consulta = sql
		Res.totalRepresentantes = Res.ObtenerRepresentantes()
		Dim originalCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
		' Change culture to en-US.
		Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
		Res.FechaConsulta = Now.ToShortDateString & " " & Now.ToShortTimeString
		Thread.CurrentThread.CurrentCulture = New CultureInfo("es-ES")
		json = JsonConvert.SerializeObject(Res)

		If sql = "SELECT * FROM listadoOperaciones;" Then

			'Aprovechamos que hacemos una consulta completa y la guardamos

			If File.Exists(HttpContext.Current.Server.MapPath(".") & "\JSON\result.json") Then
				File.Delete(HttpContext.Current.Server.MapPath(".") & "\JSON\result.json")
			End If
			Dim fs As FileStream = File.Create(HttpContext.Current.Server.MapPath(".") & "\JSON\result.json")

			AddText(fs, json)

			fs.Close()

		End If

		dt = Nothing

		Return json

	End Function
	Private Shared Sub AddText(ByVal fs As FileStream, ByVal value As String)
		Dim info As Byte() = New UTF8Encoding(True).GetBytes(value)
		fs.Write(info, 0, info.Length)
	End Sub

	Private Function CadenaConsulta(parCodRep As String, parCodCli As String, parTipodoc As String, parCodigodoc As String) As String
		Dim strCadenaConsulta As String
		Dim strCadenaConsultaBasica As String

		strCadenaConsultaBasica = "SELECT * FROM listadopedidos WHERE ("
		strCadenaConsulta = strCadenaConsultaBasica

		If (IsNothing(parCodRep) And IsNothing(parCodCli) And IsNothing(parTipodoc) And IsNothing(parCodigodoc) And parCodigodoc <> "0") Then

			strCadenaConsulta = "SELECT * FROM listadopedidos"

		Else

			If Not IsNothing(parCodRep) And parCodRep <> "0" Then
				strCadenaConsulta = strCadenaConsulta & "(codrep = " & parCodRep & ")"
			End If

			If Not IsNothing(parCodCli) Then
				If Not strCadenaConsulta.Equals(strCadenaConsultaBasica) Then
					strCadenaConsulta = strCadenaConsulta & " AND "
				End If
				strCadenaConsulta = strCadenaConsulta & "(codigo = '" & parCodCli & "')"
			End If


			If Not IsNothing(parTipodoc) Then
				If Not strCadenaConsulta.Equals(strCadenaConsultaBasica) Then
					strCadenaConsulta = strCadenaConsulta & " AND "
				End If
				strCadenaConsulta = strCadenaConsulta & "(tipodoc = '" & parTipodoc & "')"
			End If

			If Not IsNothing(parTipodoc) Then
				If Not strCadenaConsulta.Equals(strCadenaConsultaBasica) Then
					strCadenaConsulta = strCadenaConsulta & " AND "
				End If
				strCadenaConsulta = strCadenaConsulta & "(codigodoc = " & parCodigodoc & ")"
			End If

			strCadenaConsulta = strCadenaConsulta & ");"

		End If

		Return strCadenaConsulta

	End Function

	Private Function CadenaConsulta2(params As String) As String
		Dim strCadenaConsulta As String
		Dim strCadenaConsultaBasica As String

		Dim jss As New JavaScriptSerializer()
		Dim dict As Dictionary(Of String, String) = jss.Deserialize(Of Dictionary(Of String, String))(params)



		strCadenaConsultaBasica = "SELECT * FROM listadoOperaciones WHERE ("
		strCadenaConsulta = strCadenaConsultaBasica

		If (dict("AccesoCli") = "*" And dict("AccesoRep") = "*") Then

			strCadenaConsulta = "SELECT * FROM listadoOperaciones where (codrep > 0);"

		Else

			If dict("AccesoCli") = "*" And dict("AccesoRep") <> "*" Then
				strCadenaConsulta = strCadenaConsulta & "(codrep IN (" & dict("AccesoRep") & "))"
			End If

			If dict("AccesoCli") <> "*" And dict("AccesoRep") = "*" Then
				strCadenaConsulta = strCadenaConsulta & "(codigo = '" & dict("AccesoCli") & "')"
			End If


			strCadenaConsulta = strCadenaConsulta & ");"

		End If

		Return strCadenaConsulta

	End Function

End Class









