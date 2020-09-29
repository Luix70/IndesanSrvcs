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
Imports System.Net.Mail
Imports System.Security.Cryptography
Imports System.Reflection
Imports IndesanSrvcs

Public Class QueryJson

	Shared strConexion As String = ConfigurationManager.ConnectionStrings("myCon").ConnectionString
	Shared dt As DataTable
	Shared intCurrentRow As Integer = 0

	Public Class Literales
		Private es_ As String
		Private en_ As String
		Private fr_ As String

		Public Property es As String
			Get
				Return es_
			End Get
			Set(value As String)
				es_ = value
			End Set
		End Property

		Public Property en As String
			Get
				Return en_
			End Get
			Set(value As String)
				en_ = value
			End Set
		End Property

		Public Property fr As String
			Get
				Return fr_
			End Get
			Set(value As String)
				fr_ = value
			End Set
		End Property
	End Class
	Public Class Tarifa

		Private tarifa_ As String
		Private precio_ As Single
		Private moneda_ As String
		Private descripcion_ As String


		Public Property Tarifa As String
			Get
				Return tarifa_
			End Get
			Set(value As String)
				tarifa_ = value
			End Set
		End Property

		Public Property Precio As Single
			Get
				Return precio_
			End Get
			Set(value As Single)
				precio_ = value
			End Set
		End Property

		Public Property Moneda As String
			Get
				Return moneda_
			End Get
			Set(value As String)
				moneda_ = value
			End Set
		End Property

		Public Property Descripcion As String
			Get
				Return descripcion_
			End Get
			Set(value As String)
				descripcion_ = value
			End Set
		End Property
	End Class

	Public Class Oferta

		Private Id_ As Long
		Private Cod_ As String
		Private Desc_ As Literales
		Private Desc2_ As Literales
		Private Disponibles_ As Long
		Private Imagen_ As String
		Private Precios_ As New List(Of Tarifa)
		Private Codagrupacion_ As String
		Private Reservadas_ As Long
		Private Codembalaje_ As String
		Private Bultos_ As Long
		Private Peso_ As Single
		Private Volumen_ As Single



		Public Property Id As Long
			Get
				Return Id_
			End Get
			Set(value As Long)
				Id_ = value
			End Set
		End Property

		Public Property Cod As String
			Get
				Return Cod_
			End Get
			Set(value As String)
				Cod_ = value
			End Set
		End Property

		Public Property Disponibles As Long
			Get
				Return Disponibles_
			End Get
			Set(value As Long)
				Disponibles_ = value
			End Set
		End Property

		Public Property Imagen As String
			Get
				Return Imagen_
			End Get
			Set(value As String)
				Imagen_ = value
			End Set
		End Property

		Public Property Desc As Literales
			Get
				Return Desc_
			End Get
			Set(value As Literales)
				Desc_ = value
			End Set
		End Property

		Public Property Desc2 As Literales
			Get
				Return Desc2_
			End Get
			Set(value As Literales)
				Desc2_ = value
			End Set
		End Property

		Public Property Precios As List(Of Tarifa)
			Get
				Return Precios_
			End Get
			Set(value As List(Of Tarifa))
				Precios_ = value
			End Set
		End Property

		Public Property Codagrupacion As String
			Get
				Return Codagrupacion_
			End Get
			Set(value As String)
				Codagrupacion_ = value
			End Set
		End Property

		Public Property Reservadas As Long
			Get
				Return Reservadas_
			End Get
			Set(value As Long)
				Reservadas_ = value
			End Set
		End Property

		Public Property Codembalaje As String
			Get
				Return Codembalaje_
			End Get
			Set(value As String)
				Codembalaje_ = value
			End Set
		End Property

		Public Property Bultos As Long
			Get
				Return Bultos_
			End Get
			Set(value As Long)
				Bultos_ = value
			End Set
		End Property

		Public Property Peso As Single
			Get
				Return Peso_
			End Get
			Set(value As Single)
				Peso_ = value
			End Set
		End Property

		Public Property Volumen As Single
			Get
				Return Volumen_
			End Get
			Set(value As Single)
				Volumen_ = value
			End Set
		End Property
	End Class

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

	Public Class ClienteCarrito

		Private codCliente_ As String
		Private rzs_ As String
		Private nombrecomercial_ As String
		Private cif_ As String
		Private Direccion_predet_ As Long
		Private FormaPago_ As String
		Private DirEnvio_ As Long
		Private DirFacturacion_ As Direccion
		Private DireccionesEnvio_ As List(Of Direccion)
		Public Sub New()

		End Sub

		Public Sub New(codCliente As String, rzs As String, nombrecomercial As String, cif As String, direccion_predet As Long, dirEnvio As Long, formaPago As String, dirFacturacion As Direccion, direccionesEnvio As List(Of Direccion))
			Me.CodCliente = codCliente
			Me.Rzs = rzs
			Me.Nombrecomercial = nombrecomercial
			Me.Cif = cif
			Me.Direccion_predet = direccion_predet
			Me.DirEnvio = dirEnvio
			Me.FormaPago = formaPago
			Me.DirFacturacion = dirFacturacion
			Me.DireccionesEnvio = direccionesEnvio
		End Sub

		Public Property CodCliente As String
			Get
				Return codCliente_
			End Get
			Set(value As String)
				codCliente_ = value
			End Set
		End Property

		Public Property Rzs As String
			Get
				Return rzs_
			End Get
			Set(value As String)
				rzs_ = value
			End Set
		End Property

		Public Property Nombrecomercial As String
			Get
				Return nombrecomercial_
			End Get
			Set(value As String)
				nombrecomercial_ = value
			End Set
		End Property

		Public Property Cif As String
			Get
				Return cif_
			End Get
			Set(value As String)
				cif_ = value
			End Set
		End Property

		Public Property Direccion_predet As Long
			Get
				Return Direccion_predet_
			End Get
			Set(value As Long)
				Direccion_predet_ = value
			End Set
		End Property
		Public Property DirEnvio As Long
			Get
				Return DirEnvio_
			End Get
			Set(value As Long)
				DirEnvio_ = value
			End Set
		End Property
		Public Property FormaPago As String
			Get
				Return FormaPago_
			End Get
			Set(value As String)
				FormaPago_ = value
			End Set
		End Property

		Public Property DirFacturacion As Direccion
			Get
				Return DirFacturacion_
			End Get
			Set(value As Direccion)
				DirFacturacion_ = value
			End Set
		End Property

		Public Property DireccionesEnvio As List(Of Direccion)
			Get
				Return DireccionesEnvio_
			End Get
			Set(value As List(Of Direccion))
				DireccionesEnvio_ = value
			End Set
		End Property


	End Class
	Public Class Direccion
		Private codigo_ As String
		Private codsucursal_ As Long
		Private NombreSucursal_ As String
		Private direccion_ As String
		Private codpostal_ As String
		Private poblacion_ As String
		Private provincia_ As String
		Private telef_ As String
		Private email_ As String
		Private Observaciones_ As String
		Private demora_Agencia_ As Long
		Private Agencia_ As String
		Public Sub New()

		End Sub

		Public Sub New(codigo As String, codsucursal As Long, nombreSucursal As String, direccion As String, codpostal As String, poblacion As String, provincia As String, telef As String, email As String, observaciones As String, demora_Agencia As Long, agencia As String)
			Me.Codigo = codigo
			Me.Codsucursal = codsucursal
			Me.NombreSucursal = nombreSucursal
			Me.Direccion = direccion
			Me.Codpostal = codpostal
			Me.Poblacion = poblacion
			Me.Provincia = provincia
			Me.Telef = telef
			Me.Email = email
			Me.Observaciones = observaciones
			Me.Demora_Agencia = demora_Agencia
			Me.Agencia = agencia
		End Sub

		Public Property Codigo As String
			Get
				Return codigo_
			End Get
			Set(value As String)
				codigo_ = value
			End Set
		End Property

		Public Property Codsucursal As Long
			Get
				Return codsucursal_
			End Get
			Set(value As Long)
				codsucursal_ = value
			End Set
		End Property

		Public Property NombreSucursal As String
			Get
				Return NombreSucursal_
			End Get
			Set(value As String)
				NombreSucursal_ = value
			End Set
		End Property

		Public Property Direccion As String
			Get
				Return direccion_
			End Get
			Set(value As String)
				direccion_ = value
			End Set
		End Property

		Public Property Codpostal As String
			Get
				Return codpostal_
			End Get
			Set(value As String)
				codpostal_ = value
			End Set
		End Property

		Public Property Poblacion As String
			Get
				Return poblacion_
			End Get
			Set(value As String)
				poblacion_ = value
			End Set
		End Property

		Public Property Provincia As String
			Get
				Return provincia_
			End Get
			Set(value As String)
				provincia_ = value
			End Set
		End Property

		Public Property Telef As String
			Get
				Return telef_
			End Get
			Set(value As String)
				telef_ = value
			End Set
		End Property

		Public Property Email As String
			Get
				Return email_
			End Get
			Set(value As String)
				email_ = value
			End Set
		End Property

		Public Property Observaciones As String
			Get
				Return Observaciones_
			End Get
			Set(value As String)
				Observaciones_ = value
			End Set
		End Property

		Public Property Demora_Agencia As Long
			Get
				Return demora_Agencia_
			End Get
			Set(value As Long)
				demora_Agencia_ = value
			End Set
		End Property

		Public Property Agencia As String
			Get
				Return Agencia_
			End Get
			Set(value As String)
				Agencia_ = value
			End Set
		End Property


	End Class


	Public Class PedidoCarrito
		Private observaciones_ As String
		Private referencia_ As String
		Private codPedido_ As Long

		Private datosCliente_ As ClienteCarrito
		Private listaArticulos_ As List(Of Oferta)
		Public Property CodPedido As Long
			Get
				Return codPedido_
			End Get
			Set(value As Long)
				codPedido_ = value
			End Set
		End Property
		Public Property Observaciones As String
			Get
				Return observaciones_
			End Get
			Set(value As String)
				observaciones_ = value
			End Set
		End Property

		Public Property Referencia As String
			Get
				Return referencia_
			End Get
			Set(value As String)
				referencia_ = value
			End Set
		End Property

		Public Property DatosCliente As ClienteCarrito
			Get
				Return datosCliente_
			End Get
			Set(value As ClienteCarrito)
				datosCliente_ = value
			End Set
		End Property

		Public Property ListaArticulos As List(Of Oferta)
			Get
				Return listaArticulos_
			End Get
			Set(value As List(Of Oferta))
				listaArticulos_ = value
			End Set
		End Property


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
	Private Class FichaCliente
		'Tabla Clientes
		Private _codigo1 As String
		Private _rzs1 As String
		Private _nombrecomercial1 As String
		Private _direcc1 As String
		Private _poblacion1 As String
		Private _cp11 As String
		Private _telef1 As String
		Private _fax1 As String
		Private _cif1 As String
		Private _tarifa1 As String
		Private _email1 As String
		Private _idioma1 As String
		Private _latitud_cli1 As String
		Private _longitud_cli1 As String
		Private _direccionFraElectronica1 As String



		'Tabla Provincias
		Private _nomProvincia1 As String

		'Tabla Representantes
		Private _nomRepresentante1 As String
		Private _telRepresentante1 As String
		Private _emailRepresentante1 As String


		'Control
		Public Found As Boolean

		Public Property Codigo As String
			Get
				Return _codigo1
			End Get
			Set(value As String)
				_codigo1 = value
			End Set
		End Property

		Public Property Rzs As String
			Get
				Return _rzs1
			End Get
			Set(value As String)
				_rzs1 = value
			End Set
		End Property

		Public Property Nombrecomercial As String
			Get
				Return _nombrecomercial1
			End Get
			Set(value As String)
				_nombrecomercial1 = value
			End Set
		End Property

		Public Property Direcc As String
			Get
				Return _direcc1
			End Get
			Set(value As String)
				_direcc1 = value
			End Set
		End Property

		Public Property Poblacion As String
			Get
				Return _poblacion1
			End Get
			Set(value As String)
				_poblacion1 = value
			End Set
		End Property

		Public Property Cp1 As String
			Get
				Return _cp11
			End Get
			Set(value As String)
				_cp11 = value
			End Set
		End Property

		Public Property Telef As String
			Get
				Return _telef1
			End Get
			Set(value As String)
				_telef1 = value
			End Set
		End Property

		Public Property Fax As String
			Get
				Return _fax1
			End Get
			Set(value As String)
				_fax1 = value
			End Set
		End Property

		Public Property Cif As String
			Get
				Return _cif1
			End Get
			Set(value As String)
				_cif1 = value
			End Set
		End Property

		Public Property Tarifa As String
			Get
				Return _tarifa1
			End Get
			Set(value As String)
				_tarifa1 = value
			End Set
		End Property

		Public Property Email As String
			Get
				Return _email1
			End Get
			Set(value As String)
				_email1 = value
			End Set
		End Property

		Public Property Idioma As String
			Get
				Return _idioma1
			End Get
			Set(value As String)
				_idioma1 = value
			End Set
		End Property

		Public Property Latitud_cli As String
			Get
				Return _latitud_cli1
			End Get
			Set(value As String)
				_latitud_cli1 = value
			End Set
		End Property

		Public Property Longitud_cli As String
			Get
				Return _longitud_cli1
			End Get
			Set(value As String)
				_longitud_cli1 = value
			End Set
		End Property

		Public Property DireccionFraElectronica As String
			Get
				Return _direccionFraElectronica1
			End Get
			Set(value As String)
				_direccionFraElectronica1 = value
			End Set
		End Property

		Public Property NomProvincia As String
			Get
				Return _nomProvincia1
			End Get
			Set(value As String)
				_nomProvincia1 = value
			End Set
		End Property

		Public Property NomRepresentante As String
			Get
				Return _nomRepresentante1
			End Get
			Set(value As String)
				_nomRepresentante1 = value
			End Set
		End Property

		Public Property TelRepresentante As String
			Get
				Return _telRepresentante1
			End Get
			Set(value As String)
				_telRepresentante1 = value
			End Set
		End Property

		Public Property EmailRepresentante As String
			Get
				Return _emailRepresentante1
			End Get
			Set(value As String)
				_emailRepresentante1 = value
			End Set
		End Property
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
	Public Function DatosCliente(id As String) As String
		Dim Json As String = ""
		Dim culture As New CultureInfo("en-US")
		intCurrentRow = 0
		Dim js As JavaScriptSerializer
		js = New JavaScriptSerializer()
		js.MaxJsonLength = 100000000



		Dim res As New FichaCliente ' valores por defecto
		res.Found = False
		res.Codigo = id


		Dim strCadenaConsulta As String

		strCadenaConsulta = "SELECT CLIENTES_RST.codigo, CLIENTES_RST.rzs, CLIENTES_RST.nombrecomercial, CLIENTES_RST.direcc, CLIENTES_RST.poblacion, CLIENTES_RST.cp1, CLIENTES_RST.telef, CLIENTES_RST.fax, CLIENTES_RST.cif, CLIENTES_RST.tarifa , CLIENTES_RST.email , CLIENTES_RST.Idioma, CLIENTES_RST.latitud_cli, CLIENTES_RST.longitud_cli, CLIENTES_RST.direccionFraElectronica, Provincias.provincia AS nomProvincia, REPRESENTANTES_RST.nombre AS nomRepresentante, REPRESENTANTES_RST.telefono AS telRepresentante, REPRESENTANTES_RST.email AS emailRepresentante " &
								"FROM REPRESENTANTES_RST INNER JOIN (Provincias INNER JOIN CLIENTES_RST ON Provincias.codprov = CLIENTES_RST.codprov) ON REPRESENTANTES_RST.codigo = CLIENTES_RST.codrep1 " &
								$"WHERE (((CLIENTES_RST.codigo)='{id}'));"
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



		If dt.Rows.Count = 1 Then


			With dt.Rows(0)

				Dim type As Type = res.GetType()
				Dim properties() As Reflection.PropertyInfo = type.GetProperties()
				For Each _property As Reflection.PropertyInfo In properties




					If IsDBNull(.Item(_property.Name)) Then
						_property.SetValue(res, "")
					Else
						_property.SetValue(res, .Item(_property.Name).ToString())
					End If



				Next


				res.Found = True
				'res.codigo = id
				'res.cif = .Item("cif")
				'res.cp1 = .Item("cp1")
				'res.direcc = .Item("direcc")
				'res.direccionFraElectronica = .Item("direccionFraElectronica")
				'res.email = .Item("email")
				'res.emailRepresentante = .Item("emailRepresentante")
				'res.fax = .Item("fax")
				'res.Idioma = .Item("Idioma")
				'res.latitud_cli = .Item("latitud_cli")
				'res.longitud_cli = .Item("longitud_cli")
				'res.nombrecomercial = .Item("nombrecomercial")
				'res.nomProvincia = .Item("nomProvincia")
				'res.nomRepresentante = .Item("nomRepresentante")
				'res.poblacion = .Item("poblacion")
				'res.rzs = .Item("rzs")
				'res.tarifa = .Item("tarifa")
				'res.telef = .Item("telef")
				'res.telRepresentante = .Item("telRepresentante")
			End With


		End If

		Json = js.Serialize(res)
		Return Json



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

			If Not CBool(cdt.Rows(0).Item("Activada")) Then
				nc.IdCredencial = -1
				nc.Email = cdt.Rows(0).Item("Email")
				nc.Password = "NO_ACTIVADA"
			Else
				Dim hashAlg As HashAlgorithm = New SHA256CryptoServiceProvider()
				Dim bytValue As Byte() = System.Text.Encoding.UTF8.GetBytes(cdt.Rows(0).Item("salt") & "." & Password)
				Dim bytHash As Byte() = hashAlg.ComputeHash(bytValue)
				Dim passBase64 As String = Convert.ToBase64String(bytHash)

				If cdt.Rows(0).Item("Password") = passBase64 Then
					nc.IdCredencial = cdt.Rows(0).Item("IdCredencial")
					nc.TipoEntidad = cdt.Rows(0).Item("TipoEntidad")
					nc.NombreUsuario = cdt.Rows(0).Item("NombreUsuario")
					nc.Password = "OK"
					nc.AccesoCli = cdt.Rows(0).Item("AccesoCli")
					nc.AccesoRep = cdt.Rows(0).Item("AccesoRep")
					nc.Email = cdt.Rows(0).Item("Email")
					nc.Idioma = cdt.Rows(0).Item("Idioma")
					nc.VerPrecios = cdt.Rows(0).Item("verPrecios")
					nc.Tarifa = cdt.Rows(0).Item("Tarifa")
					nc.TarifaP = cdt.Rows(0).Item("TarifaP")

					If cdt.Rows(0).Item("verPedidos") = True Then
						nc.AccesoDocumentos.Add("1")
						nc.AccesoDocumentos.Add("42")
					End If

					If cdt.Rows(0).Item("verConfClientes") = True Then
						nc.AccesoDocumentos.Add("8")

					End If

					If cdt.Rows(0).Item("verAlbaranes") = True Then

						nc.AccesoDocumentos.Add("2")

					End If

					If cdt.Rows(0).Item("verConfRepresentantes") = True Then
						nc.AccesoDocumentos.Add("40")

					End If

					If cdt.Rows(0).Item("verFacturas") = True Then
						nc.AccesoDocumentos.Add("3")


					End If

					If cdt.Rows(0).Item("verAlbaranesFirmados") = True Then

						nc.AccesoDocumentos.Add("45")

					End If

					If cdt.Rows(0).Item("verFacturasEnviadas") = True Then

						nc.AccesoDocumentos.Add("34")

					End If


					'Actualizamos la fecha del ultimo acceso

					Dim strFecha As String
					Dim dtFecha As Date = Date.Now().ToLocalTime()
					strFecha = $"{dtFecha.Month}/{dtFecha.Day}/{dtFecha.Year} {dtFecha.Hour}:{dtFecha.Minute}:{dtFecha.Second}"

					Dim strCad As String = $"UPDATE Credenciales_rst Set Credenciales_rst.FechaUltimoAcceso = #{strFecha}# " &
												$"WHERE (((Credenciales_rst.TipoEntidad)='CL') AND ((Credenciales_rst.AccesoCli)='{nc.AccesoCli}'));"

					Cons = New OleDb.OleDbConnection

					Cons.ConnectionString = strConexion
					Cons.Open()

					Dim cmd As New OleDbCommand(strCad, Cons)

					Dim i As Long
					i = cmd.ExecuteNonQuery()
					Cons.Close()
					Cons = Nothing




				Else
					nc.IdCredencial = -1
					nc.Email = cdt.Rows(0).Item("Email")
					nc.Password = "BAD_PASSWORD"
				End If
			End If



		Else
			nc.IdCredencial = -1
			nc.Email = "NO_USER"
			nc.Password = "NO_USER"
		End If




		Return nc


	End Function
	Public Function VerificarActivacion(candidato As String, codigo As String, lan As String) As String
		Dim strConsulta As String
		Dim cdt As New DataTable
		Dim Cons As New OleDb.OleDbConnection

		strConsulta = "SELECT Credenciales_rst.TipoEntidad, Credenciales_rst.AccesoCli, Credenciales_rst.codigoActivacion, Credenciales_rst.Activada " &
						"FROM Credenciales_rst " &
						$"WHERE (((Credenciales_rst.TipoEntidad)='CL') AND ((Credenciales_rst.AccesoCli)='{candidato}'));"

		Cons.ConnectionString = strConexion
		Cons.Open()

		Using dad As New OleDb.OleDbDataAdapter(strConsulta, Cons)

			dad.Fill(cdt)


		End Using


		Cons.Close()
		Cons = Nothing

		If cdt.Rows.Count = 1 Then
			If CBool(cdt.Rows(0).Item("Activada")) Then
				Return "ALREADY_ACTIVATED"
			End If

			If cdt.Rows(0).Item("codigoActivacion") <> codigo Then
				Return "CODE_MISMATCH"
			End If

			Try

				Dim strFecha As String
				Dim dtFecha As Date = Date.Now().ToLocalTime()
				strFecha = $"{dtFecha.Month}/{dtFecha.Day}/{dtFecha.Year} {dtFecha.Hour}:{dtFecha.Minute}:{dtFecha.Second}"

				Dim strCad As String = $"UPDATE Credenciales_rst Set Credenciales_rst.codigoActivacion = Null, Credenciales_rst.Activada = True, Credenciales_rst.FechaActivacion = #{strFecha}# " &
										$"WHERE (((Credenciales_rst.TipoEntidad)='CL') AND ((Credenciales_rst.AccesoCli)='{candidato}') AND ((Credenciales_rst.codigoActivacion)='{codigo}'));"

				Cons = New OleDb.OleDbConnection

				Cons.ConnectionString = strConexion
				Cons.Open()

				Dim cmd As New OleDbCommand(strCad, Cons)

				Dim i As Long
				i = cmd.ExecuteNonQuery()
				Cons.Close()
				Cons = Nothing
				Return "OK"

			Catch ex As Exception
				Return "ACTIVATION_ERROR"
			End Try



		Else
			Return "ACTIVATION_CODE_NOT_FOUND"
		End If
	End Function


	Public Function VerificarCandidato(Username As String, Cif As String, Password As String, lan As String) As String

		Dim strConsulta As String
		Dim cdt As New DataTable
		Dim Cons As New OleDb.OleDbConnection

		strConsulta = "SELECT Clientes_rst.email, Clientes_rst.cif, Clientes_rst.codigo , Clientes_rst.rzs, Clientes_rst.nombrecomercial, Clientes_rst.tarifa FROM Clientes_rst WHERE (((Clientes_rst.cif)='" & Cif & "' AND  (Clientes_rst.Activo = False) ));"
		Cons.ConnectionString = strConexion
		Cons.Open()

		Using dad As New OleDb.OleDbDataAdapter(strConsulta, Cons)

			dad.Fill(cdt)

		End Using


		Cons.Close()
		Cons = Nothing

		If cdt.Rows.Count > 0 Then
			If Not IsDBNull(cdt.Rows(0).Item("email")) Then
				If cdt.Rows(0).Item("email") = Username Then
					'creamos una nueva credencial
					Dim CodCliente As String = cdt.Rows(0).Item("codigo")
					Dim rzs As String = cdt.Rows(0).Item("rzs")
					Dim nc As String = rzs
					Dim tar As String = cdt.Rows(0).Item("Tarifa")

					If Not IsDBNull(cdt.Rows(0).Item("nombrecomercial")) Then
						nc = cdt.Rows(0).Item("nombrecomercial")
					End If

					Dim status As String
					status = NuevaCredencial(Username, Cif, Password, CodCliente, rzs, nc, lan, tar)

					Return status
				Else
					Return "INVALID_EMAIL"
				End If
			Else
				Return "NO_EMAIL"
			End If

		Else
			Return "USER_UNKNOW"
		End If





	End Function
	Public Function VerificarPassword(Username As String, lan As String) As String

		'Comprobar que si credencial no existe ya.
		Dim strConsulta As String
		Dim cdt As New DataTable
		Dim Cons As New OleDb.OleDbConnection
		Dim codCliente As String
		strConsulta = "SELECT Credenciales_rst.NombreUsuario, Credenciales_rst.email, Credenciales_rst.Password, Credenciales_rst.TipoEntidad, Credenciales_rst.AccesoCli, Credenciales_rst.AccesoRep, Credenciales_rst.Idioma, Credenciales_rst.salt, Credenciales_rst.Activada , Credenciales_rst.codigoActivacion " &
						"From Credenciales_rst " &
						"Where (((Credenciales_rst.email) = '" & Username & "'));"

		Cons.ConnectionString = strConexion
		Cons.Open()

		Using dad As New OleDb.OleDbDataAdapter(strConsulta, Cons)

			dad.Fill(cdt)

		End Using


		Cons.Close()
		Cons = Nothing

		' 1.- Si existe daremos al usuario al opcion de recuperar contraseña
		If cdt.Rows.Count = 1 Then
			'VER SI ESTÁ ACTIVADA
			If CBool(cdt.Rows(0).Item("Activada")) Then
				' 2.- Si no existe generamos el hash para la contraseña
				Dim passwordHash As String = CreateSalt(24) 'importante que sea 8 o cualquier numero de bytes qye no genere un padding (= ó ==)
				Dim Password As String = CreateSalt(9) ' generamos un password aleatorio
				Dim codActivacion As String = CreateSalt(24)
				Dim hashAlg As HashAlgorithm = New SHA256CryptoServiceProvider()
				Dim bytValue As Byte() = System.Text.Encoding.UTF8.GetBytes(passwordHash & "." & Password)
				Dim bytHash As Byte() = hashAlg.ComputeHash(bytValue)
				Dim passBase64 As String = Convert.ToBase64String(bytHash)
				codCliente = cdt.Rows(0).Item("accesoCli")

				'vamos a intentar guardar la contraseña en la base de datos


				Try
					Dim strCad As String = $"UPDATE Credenciales_rst Set Credenciales_rst.Salt = '{passwordHash}', Credenciales_rst.[Password] = '{passBase64}',  Credenciales_rst.[codigoActivacion] = '{codActivacion}', Credenciales_rst.[Activada] = False, Credenciales_rst.FechaActivacion = Null " &
											$"WHERE(((Credenciales_rst.email) = '{Username}'));"


					Cons = New OleDb.OleDbConnection

					Cons.ConnectionString = strConexion
					Cons.Open()

					Dim cmd As New OleDbCommand(strCad, Cons)

					Dim i As Long
					i = cmd.ExecuteNonQuery()
					Cons.Close()
					Cons = Nothing


				Catch ex As Exception
					Return "CREDENTIAL_FAILED"
				End Try


				Return SendNewPassword(Username, Password, codActivacion, codCliente, lan)

			Else
				Return "CREDENTIAL_NOT_ACTIVATED"
			End If
		Else

			Return "CREDENTIAL_NOT_EXISTS"
		End If




	End Function





	Function NuevaCredencial(Username As String, Cif As String, Password As String, CodCliente As String, rzs As String, nc As String, lan As String, tarifa As String) As String

		Dim strCodigoTemporal As String = GeneraRegistroCredencial(Username, Cif, Password, CodCliente, rzs, nc, lan, tarifa)

		If (strCodigoTemporal = "CREDENTIAL_FAILED" Or strCodigoTemporal = "CREDENTIAL_EXISTS_ACTIVATED") Then

			Return strCodigoTemporal

		End If


		Dim SMTP_SERVER As String = ConfigurationManager.AppSettings("SMTP_SERVER")
		Dim SMTP_PORT As Integer = Integer.Parse(ConfigurationManager.AppSettings("SMTP_PORT"))
		Dim SMTP_USER As String = ConfigurationManager.AppSettings("SMTP_USER")
		Dim SMTP_PASSWORD As String = ConfigurationManager.AppSettings("SMTP_PASSWORD")
		Dim SMTP_SSL As Boolean = Boolean.Parse(ConfigurationManager.AppSettings("SMTP_SSL"))

		Dim client As New SmtpClient()




		client.DeliveryMethod = SmtpDeliveryMethod.Network
		client.EnableSsl = SMTP_SSL
		client.Host = SMTP_SERVER
		client.Port = SMTP_PORT

		' setup Smtp authentication
		Dim credentials As System.Net.NetworkCredential = New System.Net.NetworkCredential(SMTP_USER, SMTP_PASSWORD)
		client.UseDefaultCredentials = False
		client.Credentials = credentials



		Dim msg As MailMessage = New MailMessage()

		Dim strBody(3) As String
		strBody(0) = $"<html><head></head><body><div style=' font-family:Arial, Helvetica, sans-serif; font-size:12px; color:#666666; padding:20px;'><h1>Bienvenido a INDESAN. </h1><p>Este email es para verificar que es Vd el propietario de la cuenta de correo empleada para el registro</p><p> Para activar su cuenta haga click en el siguiente enlace:</p> <ul><li><a href=https://indesan.com/activate?cod={strCodigoTemporal}&cli={CodCliente}&lan=es>Activar Cuenta</a></li></ul><hr/><p>INDESAN SL</p><p>webmaster@indesan.com</p></div></body>"
		strBody(1) = $"<html><head></head><body><div style=' font-family:Arial, Helvetica, sans-serif; font-size:12px; color:#666666; padding:20px;'><h1>Bienvenu sur INDESAN. </h1><p>Le but de cet messagee est de vérifier que vous êtes le proprietaire de l'adresse électronique employée pour l'inscription</p><p> Pour activer votre compte suivez le lien:</p> <ul><li><a href=https://indesan.com/activate?cod={strCodigoTemporal}&cli={CodCliente}&lan=fr>Activer Compte</a></li></ul><hr/><p>INDESAN SL</p><p>webmaster@indesan.com</p></div></body>"
		strBody(2) = $"<html><head></head><body><div style=' font-family:Arial, Helvetica, sans-serif; font-size:12px; color:#666666; padding:20px;'><h1>Welcome to INDESAN.</h1><p> This is to verify that you are the owner of the e-mail account used in the registration process</p><p>To activate your account, please follow the link:</p> <ul><li><a href=https://indesan.com/activate?com={strCodigoTemporal}&cli={CodCliente}&lan=en>Activate Account</a></li></ul><hr/><p>INDESAN SL</p><p>webmaster@indesan.com</p></div></body>"

		Dim strSubject(3) As String
		strSubject(0) = "Sus credenciales para el Área de Cliente de INDESAN"
		strSubject(1) = "Vos coordonnées pour l'Espace Client de INDESAN"
		strSubject(2) = "Your login for INDESAN's User Area"

		msg.From = New MailAddress(SMTP_USER)
		msg.To.Add(New MailAddress(Username))
		msg.To.Add(New MailAddress("webmaster@indesan.com"))

		Dim intLan As Long
		Select Case lan.ToUpper()
			Case "ES"
				intLan = 0
			Case "FR"
				intLan = 1
			Case "EN"
				intLan = 2
		End Select

		msg.Subject = strSubject(intLan)
		msg.IsBodyHtml = True
		msg.Body = strBody(intLan)

		Try
			client.Send(msg)
			Return "OK"
		Catch ex As Exception
			Return "EMAIL_FAILED"
		End Try


	End Function
	Private Function SendNewPassword(Username As String, password As String, strCodigoTemporal As String, codcliente As String, lan As String) As String
		Dim SMTP_SERVER As String = ConfigurationManager.AppSettings("SMTP_SERVER")
		Dim SMTP_PORT As Integer = Integer.Parse(ConfigurationManager.AppSettings("SMTP_PORT"))
		Dim SMTP_USER As String = ConfigurationManager.AppSettings("SMTP_USER")
		Dim SMTP_PASSWORD As String = ConfigurationManager.AppSettings("SMTP_PASSWORD")
		Dim SMTP_SSL As Boolean = Boolean.Parse(ConfigurationManager.AppSettings("SMTP_SSL"))

		Dim client As New SmtpClient()




		client.DeliveryMethod = SmtpDeliveryMethod.Network
		client.EnableSsl = SMTP_SSL
		client.Host = SMTP_SERVER
		client.Port = SMTP_PORT

		' setup Smtp authentication
		Dim credentials As System.Net.NetworkCredential = New System.Net.NetworkCredential(SMTP_USER, SMTP_PASSWORD)
		client.UseDefaultCredentials = False
		client.Credentials = credentials



		Dim msg As MailMessage = New MailMessage()

		Dim strBody(3) As String
		strBody(0) = $"<html><head></head><body><div style=' font-family:Arial, Helvetica, sans-serif; font-size:12px; color:#666666; padding:20px;'><h1>MENSAJE DE INDESAN S.L. </h1><p>A continuación encontrará un juego de credenciales para identificarse en el área de usuario de la web de INDESAN</p><ul><li> Nombre de Usuario: {Username}</li><li> Contraseña: {password}</li></ul><p> Para activar tu cuenta haz click en el siguiente enlace:</p> <ul><li><a href=https://indesan.com/activate?cod={strCodigoTemporal}&cli={codcliente}&lan=es>Activar Cuenta</a></li></ul><p>Le recordamos que desde el panel de control de su área de usuario podrá cambiar la contraseña si lo desea.</p><hr/><p>INDESAN SL</p><p>webmaster@indesan.com</p></div></body"
		strBody(1) = $"<html><head></head><body><div style=' font-family:Arial, Helvetica, sans-serif; font-size:12px; color:#666666; padding:20px;'><h1>CELUI-CI EST UN MESSAGE D'INDESAN S.L. </h1><p>Voici un nouveau jeu de coordonnées pour vous identifier sur la web d'INDESAN</p><ul><li> Nom d'utilisateur: {Username}</li><li> Mot de passe: {password}</li></ul><p> Pour activer votre compte suivez le lien:</p> <ul><li><a href=https://indesan.com/activate?cod={strCodigoTemporal}&cli={codcliente}&lan=fr>Activer Compte</a></li></ul><p>Vous aurez la possibilité de changer le mot de passe dans votre tableau de bord</p><hr/><p>INDESAN SL</p><p>webmaster@indesan.com</p></div></body"
		strBody(2) = $"<html><head></head><body><div style=' font-family:Arial, Helvetica, sans-serif; font-size:12px; color:#666666; padding:20px;'><h1>THIS IS A MESSAGE FROM INDESAN S.L. </h1><p>You will find here your new credentials for the User Area at the INDESAN Web</p><ul><li> User Name: {Username}</li><li> Password: {password}</li></ul><p>To activate your account, please follow the link:</p> <ul><li><a href=https://indesan.com/activate?cod={strCodigoTemporal}&cli={codcliente}&lan=en>Activate Account</a></li></ul><p>You'll be able to change it from your dashboard once logged in</p><hr/><p>INDESAN SL</p><p>webmaster@indesan.com</p></div></body"

		Dim strSubject(3) As String
		strSubject(0) = "Sus nuevas credenciales para el Área de Cliente de INDESAN"
		strSubject(1) = "Vos nouvelles coordonnées pour l'Espace Client de INDESAN"
		strSubject(2) = "Your new credentials for INDESAN's User Area"

		msg.From = New MailAddress(SMTP_USER)
		msg.To.Add(New MailAddress(Username))
		'msg.To.Add(New MailAddress("compras@indesan.com"))

		Dim intLan As Long
		Select Case lan.ToUpper()
			Case "ES"
				intLan = 0
			Case "FR"
				intLan = 1
			Case "EN"
				intLan = 2
		End Select

		msg.Subject = strSubject(intLan)
		msg.IsBodyHtml = True
		msg.Body = strBody(intLan)

		Try
			client.Send(msg)
			Return "OK"
		Catch ex As Exception
			Return "EMAIL_FAILED"
		End Try
	End Function
	Private Function GeneraRegistroCredencial(Username As String, Cif As String, Password As String, CodCliente As String, rzs As String, nc As String, lan As String, tarifa As String) As String

		'Comprobar que si credencial no existe ya.
		Dim strConsulta As String
		Dim cdt As New DataTable
		Dim Cons As New OleDb.OleDbConnection

		strConsulta = "SELECT Credenciales_rst.NombreUsuario, Credenciales_rst.email, Credenciales_rst.Password, Credenciales_rst.TipoEntidad, Credenciales_rst.AccesoCli, Credenciales_rst.AccesoRep, Credenciales_rst.Idioma, Credenciales_rst.salt, Credenciales_rst.Activada , Credenciales_rst.codigoActivacion , Credenciales_rst.Tarifa, Credenciales_rst.TarifaP " &
						"From Credenciales_rst " &
						"Where (((Credenciales_rst.email) = '" & Username & "'));"

		Cons.ConnectionString = strConexion
		Cons.Open()

		Using dad As New OleDb.OleDbDataAdapter(strConsulta, Cons)

			dad.Fill(cdt)

		End Using


		Cons.Close()
		Cons = Nothing

		' 1.- Si existe daremos al usuario al opcion de recuperar contraseña
		If cdt.Rows.Count > 0 Then
			'VER SI ESTÁ ACTIVADA
			If CBool(cdt.Rows(0).Item("Activada")) = "True" Then
				Return "CREDENTIAL_EXISTS_ACTIVATED" 'La credencial Ya existe
			Else
				Return cdt.Rows(0).Item("codigoActivacion")
			End If
		Else

			' 2.- Si no existe generamos el hash para la contraseña
			Dim passwordHash As String = CreateSalt(24) 'importante que sea 24 o cualquier numero de bytes qye no genere un padding (= ó ==)
			Dim codigoActivacion As String = CreateSalt(24)

			Dim hashAlg As HashAlgorithm = New SHA256CryptoServiceProvider()
			Dim bytValue As Byte() = System.Text.Encoding.UTF8.GetBytes(passwordHash & "." & Password)
			Dim bytHash As Byte() = hashAlg.ComputeHash(bytValue)
			Dim passBase64 As String = Convert.ToBase64String(bytHash)

			'vamos a intentar guardar la contraseña en la base de datos

			If IsNothing(lan) Then
				lan = "ES"
			End If

			Dim strFecha As String
			Dim dtFecha As Date = Date.Now().ToLocalTime()
			strFecha = $"{dtFecha.Month}/{dtFecha.Day}/{dtFecha.Year} {dtFecha.Hour}:{dtFecha.Minute}:{dtFecha.Second}"

			Try
				Dim strCad As String

				strCad = "INSERT INTO Credenciales_rst ( TipoEntidad, NombreUsuario, Salt, [Password], AccesoCli, AccesoRep, email, Idioma, Activada, codigoActivacion, FechaAlta, Tarifa, TarifaP ) " &
										   $"SELECT 'CL' AS te, '{nc}' AS nc, '{passwordHash}' AS sa, '{passBase64}' AS pa, '{CodCliente}' AS ac, '*' AS ar, '{Username}' AS email, '{lan.ToUpper()}' AS lan, False AS act, '{codigoActivacion}' as codigoActivacion, #{strFecha}# As FechaAlta, '{tarifa}' as Tarifa, '{tarifa & "P" }' as TarifaP  ;"



				Cons = New OleDb.OleDbConnection

				Cons.ConnectionString = strConexion
				Cons.Open()

				Dim cmd As New OleDbCommand(strCad, Cons)

				Dim i As Long
				i = cmd.ExecuteNonQuery
				Cons.Close()
				Cons = Nothing


			Catch ex As Exception
				Return "CREDENTIAL_FAILED"
			End Try


			Return codigoActivacion
		End If




	End Function
	Private Shared Function CreateSalt(ByVal size As Integer) As String
		Dim rng As RNGCryptoServiceProvider = New RNGCryptoServiceProvider()
		Dim buff As Byte() = New Byte(size - 1) {}
		rng.GetBytes(buff)
		Return Convert.ToBase64String(buff)
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

	Public Function CambiarPassword(usuario As String, pass As String, newpass As String) As String
		'Comprobar que si credencial no existe ya.
		Dim strConsulta As String
		Dim cdt As New DataTable
		Dim Cons As New OleDb.OleDbConnection

		strConsulta = "SELECT Credenciales_rst.NombreUsuario, Credenciales_rst.email, Credenciales_rst.Password, Credenciales_rst.TipoEntidad, Credenciales_rst.AccesoCli, Credenciales_rst.AccesoRep, Credenciales_rst.Idioma, Credenciales_rst.salt, Credenciales_rst.Activada , Credenciales_rst.codigoActivacion " &
						"From Credenciales_rst " &
						"Where (((Credenciales_rst.email) = '" & usuario & "'));"

		Cons.ConnectionString = strConexion
		Cons.Open()

		Using dad As New OleDb.OleDbDataAdapter(strConsulta, Cons)

			dad.Fill(cdt)






			' 1.- Si existe daremos al usuario al opcion de recuperar contraseña
			If cdt.Rows.Count = 1 Then
				Dim passwordHash As String = CreateSalt(24) 'importante que sea 24 o cualquier numero de bytes qye no genere un padding (= ó ==)

				Dim hashAlg As HashAlgorithm = New SHA256CryptoServiceProvider()
				Dim bytValue As Byte() = System.Text.Encoding.UTF8.GetBytes(passwordHash & "." & newpass)
				Dim bytHash As Byte() = hashAlg.ComputeHash(bytValue)
				Dim passBase64 As String = Convert.ToBase64String(bytHash)



				Try
					Dim strCad As String

					strCad = $"UPDATE Credenciales_rst Set Credenciales_rst.Salt = '{passwordHash}', Credenciales_rst.[Password] = '{passBase64}' " &
								$"WHERE(((Credenciales_rst.email) = '{usuario}'));"


					Dim cmd As New OleDbCommand(strCad, Cons)

					Dim i As Long
					i = cmd.ExecuteNonQuery

					If i = 1 Then Return newpass


				Catch ex As Exception
					Return pass
				End Try



				Return newpass
			Else
				Return pass
			End If

		End Using

		Cons.Close()
		Cons = Nothing

	End Function

	Public Function ObtenerOfertas() As String

		Dim Json As String = ""
		Dim culture As New CultureInfo("en-US")
		intCurrentRow = 0
		Dim js As JavaScriptSerializer
		js = New JavaScriptSerializer()
		js.MaxJsonLength = 100000000

		Dim strCadenaConsulta As String

		strCadenaConsulta = "SELECT * FROM OfertasTarifas;"
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

		Dim res As New List(Of Oferta)

		Dim curID As Long = -1
		Dim totalOfertas As Long = 0
		Dim ofertaActual As Oferta
		Dim totalPrecios As Long = 0
		Dim precios As New List(Of Tarifa)

		For Each row As DataRow In dt.Rows

			If row.Item("id") <> curID Then
				totalOfertas += 1
				ofertaActual = New Oferta

				res.Add(ofertaActual)

				precios = New List(Of Tarifa)

				curID = row.Item("id")

				With ofertaActual

					.Id = curID
					.Cod = row.Item("cod")
					.Disponibles = row.Item("disponibles")
					.Imagen = row.Item("imagen")
					.Codagrupacion = row.Item("codagrupacion")
					.Desc = New Literales With {
						.es = row.Item("desc_es"),
						.en = row.Item("desc_en"),
						.fr = row.Item("desc_fr")
					}
					.Desc2 = New Literales With {
						.es = row.Item("desc2_es"),
						.en = row.Item("desc2_en"),
						.fr = row.Item("desc2_fr")
					}

					.Precios = precios
					.Bultos = row.Item("Bultos")
					.Codembalaje = row.Item("Codembalaje")
					.Peso = row.Item("Peso")
					.Volumen = row.Item("Volumen")

				End With
			End If

			totalPrecios += 1

			Dim precioActual As New Tarifa
			precios.Add(precioActual)

			With precioActual
				.Tarifa = row.Item("CodTarifa")
				.Moneda = row.Item("Moneda")
				.Precio = row.Item("Precio")
				.Descripcion = row.Item("Descripcion")

			End With

		Next


		Json = js.Serialize(res)
		Return Json


	End Function
	Public Function CustData(cliente As String) As String

		'Return "{""codCliente"": """ & cliente & """}"

		Dim Json As String = ""
		Dim culture As New CultureInfo("en-US")
		intCurrentRow = 0
		Dim js As JavaScriptSerializer
		js = New JavaScriptSerializer()
		js.MaxJsonLength = 100000000

		Dim strCadenaConsulta As String

		strCadenaConsulta = $"SELECT * FROM ClientesCarrito WHERE codCliente = '{cliente}';"
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



		Dim CodCliente As String = dt.Rows(0).Item("CodCliente")
		Dim rzs As String = dt.Rows(0).Item("rzs")

		Dim nombrecomercial As String = ""
		If Not IsDBNull(dt.Rows(0).Item("nombrecomercial")) Then
			nombrecomercial = dt.Rows(0).Item("nombrecomercial")
		End If

		Dim cif As String = dt.Rows(0).Item("cif")
		Dim direccion_predet As String = dt.Rows(0).Item("direccion_predet")
		Dim direccionEnvio As String = direccion_predet
		Dim formaPago As String = dt.Rows(0).Item("DescFormaPago")

		'Direccion principal
		Dim codigo As String = dt.Rows(0).Item("CodCliente")
		Dim codsucursal As Long = 0

		Dim nombreSucursal As String = ""

		Dim dire As String = dt.Rows(0).Item("direcc")
		Dim codpostal As String = dt.Rows(0).Item("cp1")
		Dim poblacion As String = dt.Rows(0).Item("pobrzs")
		Dim provincia As String = dt.Rows(0).Item("provrzs")

		Dim telef As String = ""
		Dim email As String = ""
		Dim observaciones As String = ""

		If Not IsDBNull(dt.Rows(0).Item("telefrzs")) Then
			telef = dt.Rows(0).Item("telefrzs")
		End If

		If Not IsDBNull(dt.Rows(0).Item("email")) Then
			email = dt.Rows(0).Item("email")
		End If


		Dim demora_Agencia As Long = 0
		Dim agencia As String = ""


		Dim dirfacturacion As New Direccion(codigo:=codigo, codsucursal:=codsucursal, nombreSucursal:=nombreSucursal, direccion:=dire, codpostal:=codpostal, poblacion:=poblacion, provincia:=provincia, telef:=telef, email:=email, observaciones:=observaciones, demora_Agencia:=demora_Agencia, agencia:=agencia)

		Dim direccionesEnvio As New List(Of Direccion)

		Dim ccar As New ClienteCarrito(codCliente:=CodCliente, rzs:=rzs, nombrecomercial:=nombrecomercial, cif:=cif, direccion_predet:=direccion_predet, dirEnvio:=direccionEnvio, formaPago:=formaPago, dirFacturacion:=dirfacturacion, direccionesEnvio:=direccionesEnvio)




		For Each row As DataRow In dt.Rows

			codigo = row.Item("CodCliente")
			codsucursal = row.Item("codsucursal")



			If Not IsDBNull(row.Item("NombreSucursal")) Then
				nombreSucursal = row.Item("NombreSucursal")
			End If


			dire = row.Item("direccion")

			codpostal = row.Item("codpostal")
			poblacion = row.Item("pobsuc")
			provincia = row.Item("provsuc")



			If Not IsDBNull(row.Item("telefsuc")) Then
				telef = row.Item("telefsuc")
			End If
			If Not IsDBNull(row.Item("e-mail")) Then
				email = row.Item("e-mail")
			End If
			If Not IsDBNull(row.Item("Observaciones")) Then
				observaciones = row.Item("Observaciones")
			End If

			demora_Agencia = row.Item("Demora_agencia")
			agencia = row.Item("Agencia")


			dirfacturacion = New Direccion(codigo:=codigo, codsucursal:=codsucursal, nombreSucursal:=nombreSucursal, direccion:=dire, codpostal:=codpostal, poblacion:=poblacion, provincia:=provincia, telef:=telef, email:=email, observaciones:=observaciones, demora_Agencia:=demora_Agencia, agencia:=agencia)

			ccar.DireccionesEnvio.Add(dirfacturacion)

		Next


		Json = js.Serialize(ccar)
		Return Json

	End Function

	Public Function reservarOfertas(articulos As List(Of Oferta)) As Boolean

		Dim strCadenaConsulta As String
		strCadenaConsulta = "SELECT * FROM Ofertas;"

		'1.- Verificamos disponibilidad


		Dim Cons As New OleDb.OleDbConnection
		Cons.ConnectionString = strConexion
		Cons.Open()

		dt = New DataTable
		Dim blnDisponibles As Boolean = True

		Using dad As New OleDb.OleDbDataAdapter(strCadenaConsulta, Cons)

			Try
				dad.Fill(dt)
			Catch ex As Exception
				MsgBox(ex.Message)
			End Try



			For Each articulo As Oferta In articulos

				For Each row As DataRow In dt.Rows

					If row.Item("id") = articulo.Id Then

						blnDisponibles = blnDisponibles And (row.Item("disponibles") >= articulo.Reservadas)
						If blnDisponibles Then
							row.Item("disponibles") = row.Item("disponibles") - articulo.Reservadas
							If row.Item("disponibles") = 0 Then
								row.Delete()
							End If

						End If
					End If



				Next

			Next

			'Ahora si hemos comprobado la disponiblidad de todos los articulos, actualizamos la base de datos de ofertas

			If blnDisponibles Then

				Dim cd As New OleDbCommandBuilder(dad)
				dad.UpdateCommand = cd.GetUpdateCommand()
				dad.Update(dt)

			End If

		End Using



		Cons.Close()
		Cons = Nothing

		Return blnDisponibles

	End Function

End Class










