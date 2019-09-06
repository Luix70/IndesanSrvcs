Imports System.Web.Script.Serialization
Imports System
Imports System.IO
Imports System.Text
Imports System.Configuration
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

	Private Class Representante
		Public codrep As Long
		Public nombre As String
		Public clientes As New Collection
		Public totalClientes As Integer

		Function obtenerClientes(codRepActual As Integer) As Integer

			Dim n_clientes As Integer = 0
			Dim objCliente As Cliente

			While intCurrentRow < dt.Rows.Count AndAlso dt.Rows(intCurrentRow).Item("codrep") = codRepActual

				n_clientes += 1
				objCliente = New Cliente

				objCliente.codigo = dt.Rows(intCurrentRow).Item("codigo")
				objCliente.rzs = dt.Rows(intCurrentRow).Item("rzs")
				objCliente.poblacion = dt.Rows(intCurrentRow).Item("poblacion")

				objCliente.totalDocumentos = objCliente.obtenerDocumentos(objCliente.codigo)

				clientes.Add(objCliente, objCliente.codigo.ToString)


			End While

			' devolvemos el cursor al ultimo registro del cliente o al ultimo de la tabla

			Return n_clientes

		End Function

	End Class

	Private Class Cliente
		Public codigo As String
		Public rzs As String
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
				objDoc.fechapedido = dt.Rows(intCurrentRow).Item("fecha1").ToString
				objDoc.Importebruto = dt.Rows(intCurrentRow).Item("Bruto")

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
		Public fechapedido As String
		Public Importebruto As Single
		Public referencia As String
		Public lineas As New Collection
		Public totalLineas As Integer

		Function obtenerLineas(tipodoc As String, codigodoc As Long) As Integer
			Dim n_lineas As Integer = 0
			Dim objLinea As Linea

			While intCurrentRow < dt.Rows.Count AndAlso dt.Rows(intCurrentRow).Item("tipodoc") = tipodoc AndAlso dt.Rows(intCurrentRow).Item("codigodoc") = codigodoc
				n_lineas += 1
				objLinea = New Linea

				objLinea.linea = dt.Rows(intCurrentRow).Item("linea")
				objLinea.coart = dt.Rows(intCurrentRow).Item("coart")
				Try
					If IsDBNull(dt.Rows(intCurrentRow).Item("cantidad")) Then
						objLinea.cantidad = 0
					Else
						objLinea.cantidad = dt.Rows(intCurrentRow).Item("cantidad")
					End If
				Catch e As Exception
					objLinea.cantidad = 0
				End Try


				objLinea.descripcion = dt.Rows(intCurrentRow).Item("descripcion")
				If IsDBNull(dt.Rows(intCurrentRow).Item("ref_linea")) Then
					objLinea.ref_linea = ""
				Else
					objLinea.ref_linea = dt.Rows(intCurrentRow).Item("ref_linea")
				End If

				objLinea.precio = dt.Rows(intCurrentRow).Item("precio")

				lineas.Add(objLinea, n_lineas.ToString)

				intCurrentRow += 1

			End While

			Return n_lineas



		End Function

	End Class
	Private Class Linea
		Public linea As Integer
		Public coart As String
		Public cantidad As Single
		Public descripcion As String
		Public ref_linea As String
		Public precio As Single
	End Class

	Public Function GenerarJson(parCodRep As String, parCodCli As String, parTipodoc As String, parCodigodoc As String) As String


		Dim Json As String = ""
		dt = New DataTable
		intCurrentRow = 0
		Dim js As JavaScriptSerializer
		js = New JavaScriptSerializer()

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
		res.FechaConsulta = Now.ToShortDateString & " " & Now.ToShortTimeString
		res.totalRepresentantes = res.ObtenerRepresentantes()

		Json = js.Serialize(res)

		dt = Nothing


		Return Json

	End Function

	Public Function GenerarJson2(parUsuario As String, parPassword As String) As String

		Dim Json As String = ""

		intCurrentRow = 0
		Dim js As JavaScriptSerializer
		js = New JavaScriptSerializer()
		js.MaxJsonLength = 10000000

		Dim strCadenaConsulta As String

		strCadenaConsulta = "SELECT * FROM Credenciales_rst WHERE Credenciales_rst.NombreUsuario ='" & parUsuario & "';"
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
			res.FechaConsulta = Now.ToShortDateString & " " & Now.ToShortTimeString
			Json = js.Serialize(res)
			Return Json

		Else
			If dt.Rows(0).Field(Of String)("Contraseña") <> parPassword Then
				res.consulta = "nulo"
				res.Status = "noOK"
				res.errorcode = "Contraseña errónea"
				res.FechaConsulta = Now.ToShortDateString & " " & Now.ToShortTimeString
				Json = js.Serialize(res)
				Return Json
			Else
				'Todo correcto
				res.consulta = CadenaConsulta2("{tipoentidad:'" & dt.Rows(0).Field(Of String)("TipoEntidad") & "', AccesoCli:'" & dt.Rows(0).Field(Of String)("AccesoCli") & "', AccesoRep: '" & dt.Rows(0).Field(Of String)("AccesoRep") & "'}")
				res.Status = "OK"
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
	Function ConsultarDB(sql As String) As String
		' hay que leerlo de nuevo
		Dim Res = New Resultado
		Dim json As String = ""
		Dim js As New JavaScriptSerializer()
		js.MaxJsonLength = 10000000

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
		Res.FechaConsulta = Now.ToShortDateString & " " & Now.ToShortTimeString
		json = js.Serialize(Res)

		If sql = "SELECT * FROM listadoOperaciones where (codrep > 0);" Then

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

		If (IsNothing(parCodRep) And IsNothing(parCodCli) And IsNothing(parTipodoc) And IsNothing(parCodigodoc)) Then

			strCadenaConsulta = "SELECT * FROM listadopedidos"

		Else

			If Not IsNothing(parCodRep) Then
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

			If dict("AccesoCli") = "*" And dict("AccesoRep") <> "-" Then
				strCadenaConsulta = strCadenaConsulta & "(codrep IN (" & dict("AccesoRep") & "))"
			End If

			If dict("AccesoCli") <> "*" And dict("AccesoRep") = "-" Then
				strCadenaConsulta = strCadenaConsulta & "(codigo = '" & dict("AccesoCli") & "')"
			End If


			strCadenaConsulta = strCadenaConsulta & ");"

		End If

		Return strCadenaConsulta

	End Function

End Class









