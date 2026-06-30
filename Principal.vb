Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms
Imports ClosedXML.Excel

Public Class Principal

    Public x, y, CantPantallas, vPantallas, vCodigo, vContador, vCalculoVersion1, vCalculoVersion2, vCalculoVersion3 As Integer
    Public tipoDsn, tipoSql, vtipoSql, vWidth, vHeigth, vPosicion, respuesta, vNumeroVersion As String
    Public vConcepto, vDescripcion, vNotas, vCuenta, vImporte, vDescripcionAPU As String
    Public vImporteAPU, vNotasAPU, vCuentaAPU, vCompactada, appDataPath, carpetaDB As String
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    ' 1. Constructor: Es el mejor sitio para fijar el idioma antes de que se vea nada
    Public Sub New()
        ' Configura el idioma inicial (ej. Español)
        Dim cultura As String = My.Settings.CulturaUsuario
        System.Threading.Thread.CurrentThread.CurrentUICulture = New System.Globalization.CultureInfo(cultura)
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo(cultura)
        ' Esta llamada es vital
        InitializeComponent()
    End Sub

    Public Sub RefrescarTextos()
        ' 1. Refrescar Labels, Botones y Título del Formulario
        'Buscamos la Key en el recurso que se llame igual que el nombre del control
        For Each ctrl As Control In Me.Controls
            ' Si el control tiene un nombre que coincide con una Key en recursos, lo traduce
            Dim textoTraducido As String = My.Resources.ResourceManager.GetString(ctrl.Name)
            If Not String.IsNullOrEmpty(textoTraducido) Then
                ctrl.Text = textoTraducido
            End If
            If textoTraducido Is Nothing Then
                textoTraducido = rmse.GetString(ctrl.Name)
            End If
        Next
        ' 2. Llamar al refresco de menús
        RefrescarMenus()
    End Sub

    Public Sub RefrescarMenus()
        ' 1. Inicializa el gestor de recursos del Formulario Principal
        Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(Principal))
        ' 2. Bucle para los menús principales que están directamente en la barra
        For Each menuPrincipal As ToolStripMenuItem In Me.BarraDeMenu.Items.OfType(Of ToolStripMenuItem)()
            ' Traduce el texto del menú principal (ej: "Archivo")
            resources.ApplyResources(menuPrincipal, menuPrincipal.Name)
            ' 3. Llama a la función mágica para que revise todos los subniveles hacia abajo
            TraducirSubMenusRecursivo(menuPrincipal, resources)
        Next
        ' 3. Recorre todos los elementos de la barra de herramientas
        For Each item As ToolStripItem In BarraDeHerramientas.Items
            ' Busca si existe la traducción para el ToolTipText de este botón específico
            Dim textoTraducido As String = resources.GetString(item.Name & ".Text")
            ' Si encuentra la traducción, la aplica de inmediato
            If Not String.IsNullOrEmpty(textoTraducido) Then
                item.ToolTipText = textoTraducido
            End If
            'MsgBox("Buscando traducción para: " & item.Name & ".ToolTipText" & vbCrLf & "Traducción encontrada: " & If(String.IsNullOrEmpty(textoTraducido), "No encontrada", textoTraducido))
        Next
        ' 4. Refresca la barra de inmediato para aplicar los cambios visuales
        BarraDeHerramientas.Refresh()
        Dim textoLimpio As String = rmse.GetString("LblAvisoActivacion")

        If Not String.IsNullOrEmpty(textoLimpio) Then
            LblNotificacion.Text = textoLimpio
        End If
    End Sub

    Private Sub TraducirSubMenusRecursivo(menuPadre As ToolStripMenuItem, resources As System.ComponentModel.ComponentResourceManager)
        ' Recorre todos los subelementos del menú actual
        For Each subMenu As ToolStripMenuItem In menuPadre.DropDownItems.OfType(Of ToolStripMenuItem)()
            ' Traduce el submenú actual (Nivel 2, Nivel 3, etc.)
            resources.ApplyResources(subMenu, subMenu.Name)
            ' RECURSIVIDAD: El submenú se convierte en padre y busca si tiene sus propios hijos
            TraducirSubMenusRecursivo(subMenu, resources)
        Next
    End Sub

    Private Sub ActualizarItemMenu(ByVal item As ToolStripItem)
        ' Busca el texto por el nombre del objeto (ej: "ArchivoToolStripMenuItem")
        Dim texto As String = My.Resources.ResourceManager.GetString(item.Name)
        If Not String.IsNullOrEmpty(texto) Then item.Text = texto

        ' Si tiene submenús (hijos), los recorre recursivamente
        If TypeOf item Is ToolStripMenuItem Then
            For Each subItem As ToolStripItem In DirectCast(item, ToolStripMenuItem).DropDownItems
                ActualizarItemMenu(subItem)
            Next
        End If
    End Sub

    Private Sub Principal_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        RefrescarMenus()

        '' Esto te mostrará en la ventana de "Salida" los nombres exactos detectados
        'For Each res In Assembly.GetExecutingAssembly().GetManifestResourceNames()
        '    MsgBox(res)
        'Next

        ' Ejemplo de uso del ResourceManager para obtener una cadena traducida
        'Dim rm As New ResourceManager("Contahogar.Recursos", Assembly.GetExecutingAssembly())
        'Dim mensaje As String = resManager.GetString("SinFiltrar")
        'MsgBox(mensaje)

        '    ' 1. Creamos un lector para recorrer los recursos del componente
        '    Dim conjuntoRecursos As System.Resources.ResourceSet =
        'rmse.GetResourceSet(System.Globalization.CultureInfo.CurrentCulture, True, True)

        '    ' 2. Verificamos que contenga elementos antes de continuar
        '    If conjuntoRecursos IsNot Nothing Then

        '        ' 3. Recorremos cada recurso individualmente
        '        For Each elemento As System.Collections.DictionaryEntry In conjuntoRecursos

        '            ' 4. Mostramos el nombre del recurso y su contenido
        '            MsgBox("Clave: " & elemento.Key.ToString() & vbCrLf &
        '           "Valor: " & elemento.Value.ToString(),
        '           MsgBoxStyle.Information,
        '           "Recurso Detectado")
        '        Next
        '    Else
        '        MsgBox("No se encontraron recursos guardados.", MsgBoxStyle.Exclamation)
        '    End If


        ' El instalador actualiza los archivos, pero este código migra las preferencias
        If My.Settings.UpgradeRequired Then
            My.Settings.Upgrade()
            My.Settings.UpgradeRequired = False
            My.Settings.Save() ' Guarda el cambio para que no lo haga más en esta versión
        End If

        ' Leemos los Settings para mostrar o no la Barra de Herramientas, la Barra de Estado y el Color en las Barras
        If My.Settings.BarraHerramientas = True Then
            BarraDeHerramientasMenu.Checked = True
            Cambiarbarraherramientas()
        Else
            BarraDeHerramientasMenu.Checked = False
            Cambiarbarraherramientas()
        End If
        If My.Settings.BarraEstado = True Then
            BarraDeEstadoMenu.Checked = True
            Cambiarbarraestado()
        Else
            BarraDeEstadoMenu.Checked = False
            Cambiarbarraestado()
        End If
        If My.Settings.MenuColores = True Then
            BarraYMenuConColores.Checked = True
            CambiarColorBarraMenu()
        Else
            BarraYMenuConColores.Checked = False
            CambiarColorBarraMenu()
        End If

        If My.Settings.Codigo = "Codigo Activación: Sin Activar" Then
            LblNotificacion.Visible = True
        Else
            LblNotificacion.Visible = False
        End If

        'Cargamos la conexión vRuta Mdb
        '******************************
        ' Para AppData\Local (Recomendado para datos de una sola máquina)
        appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)

        ' Combinar con el nombre de tu compañía/aplicación y el archivo de base de datos
        carpetaDB = IO.Path.Combine(appDataPath, "A.Oberholzer", "ContaHogar3.0")
        vRuta = IO.Path.Combine(carpetaDB, "ContaHogar.mdb")
        'vRuta = "C:\ContaHogar3.0\ContaHogar.mdb" ' Para la carpeta de instalación (No recomendado, pero para usuarios que no quieran complicarse)"

        My.Settings.RutaBD = vRuta
        My.Settings.Save()
        My.Settings.Reload()

        'Leemos el Settings para saber los Monitores que había la última vez
        '********************************************************************
        CantPantallas = My.Settings.Pantallas

        'Se examinan todas las pantallas y se detectan cuantas hay
        '**********************************************************
        vPantallas = 0
        For Each scrn As Screen In Screen.AllScreens
            vPantallas += 1
        Next

        If vPantallas = 1 And CantPantallas >= 2 Then
            vPosicion = "{x=150,y=0}"
            x = Val(Mid(vPosicion, 4, (InStrRev(vPosicion, ",") - 1)))
            y = Val(Mid(vPosicion, (InStrRev(vPosicion, "=") + 1)))
            vWidth = 1139
            vHeigth = 629
        Else
            vPosicion = My.Settings.Posicion
            x = Val(Mid(vPosicion, 4, (InStrRev(vPosicion, ",") - 1)))
            y = Val(Mid(vPosicion, (InStrRev(vPosicion, "=") + 1)))

            ' PARACHOQUES: Si la posición guardada es el techo total (0), 
            ' le forzamos a bajar un poco (ej. 100) para que no se quede atrapado
            If y = 0 Then y = 100

            vWidth = My.Settings.PantallaAncho
            vHeigth = My.Settings.PantallaAlto
        End If
        ' Tu código original con el CInt que blinda el multiidioma
        Me.Location = New Point(CInt(x), CInt(y))
        Me.Size = New Size(CInt(vWidth), CInt(vHeigth))

        'Si en el ChbPantallaCompleta.Checked = True, se abre Pantalla Completa
        '***********************************************************************
        If My.Settings.PantallaCompleta = True Then
            Me.Location = New Point(0, 0)
            Me.Size = Screen.PrimaryScreen.WorkingArea.Size
            Me.Height = Screen.PrimaryScreen.WorkingArea.Height
        End If

        If My.Settings.PantallaCierre = True Then
            Me.Location = New Point(x, y)
            Me.Size = New Size(vWidth, vHeigth)
        End If

        AnchoFrmPrincipal = Me.Size.Width

        tipoDsn = "AccessMdb" ' Se conecta a Mdb
        Conectarse(tipoDsn)

        VerificarYActualizarEstructuraBD()

        'Buscamos Ejercicio
        '******************
        vAñoActual = Date.Now.Year
        cmdMdb1cr.Connection = conexion1
        cmdMdb1cr.CommandType = CommandType.Text
        cmdMdb1cr.CommandText = "Select * FROM ejercicios"
        cmdMdb1cr.CommandText += " WHERE ejercicios.EjercicioEJE = " & vAñoActual.ToString
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    vAñoEjercicio = vAñoActual
                End While
                'MsgBox("Ya Existe registro del " & vAñoActual.ToString)
            Else
                MsgBox(resManager.GetString("NoExistenRegistros") & " " & vAñoActual.ToString & ", " & rmse.GetString("SeCrearaEjercicio"))
                drMdb1.Close()
                ' 1. Diseñamos la estructura limpia para ejercicios usando el comodín '?'
                tipoSql = "INSERT INTO ejercicios (EjercicioEJE) VALUES (?)"
                cmdMdb1cr.CommandText = tipoSql

                ' 2. Limpiamos parámetros e inyectamos el año como un número entero puro
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("@EjercicioEJE", CInt(vAñoActual))
                cmdMdb1cr.CommandText = tipoSql
                Try
                    cmdMdb1cr.ExecuteNonQuery()
                    vAñoEjercicio = vAñoActual
                    MsgBox(resManager.GetString("Ejercicio") & " " & vAñoActual.ToString & " " & rmse.GetString("CreadoCorrectamente"))
                Catch ex As Exception
                    MsgBox(resManager.GetString("ErrorAlCrearEjercicio") & " " & vAñoActual.ToString)
                    MsgBox(ex.ToString)
                End Try
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorAlBuscarEjercicio") & " " & vAñoActual.ToString)
            MsgBox(ex.ToString)
            Return
        End Try

        My.Settings.Version = "3.1.1"
        My.Settings.Save()

        vActualizar = My.Settings.Actualizar
        vMoneda = My.Settings.Moneda
        vPathExportar = My.Settings.PathExportar

        ' Congelamos el redibujado
        Me.SuspendLayout()
        Me.BarraDeEstado.SuspendLayout()

        ' Alineamos a la derecha el día y la hora, y mostramos el día actual
        TsLabelDia.Alignment = ToolStripItemAlignment.Right
        TsLabelHora.Alignment = ToolStripItemAlignment.Right
        TsLabelDia.Text = DateTime.Today.ToString("d")
        Dim loTimer As New Windows.Forms.Timer With {
               .Interval = 1000
                }
        AddHandler loTimer.Tick, AddressOf IP_Timer
        loTimer.Start()

        ' Forzamos el cálculo de medidas antes de mostrar nada
        Me.BarraDeEstado.ResumeLayout(False)
        Me.BarraDeEstado.PerformLayout()
        Me.ResumeLayout(True)

        'Buscamos si hay una nueva Versión en PCloud
        '*******************************************
        BuscarActualizacion()
        If vHayNuevaVersion = "SI" Then
            VactualVDisponible.Visible = True
            VactualVDisponible.Text = resManager.GetString("VersionInstalada") & " " & My.Settings.Version & " - " & resManager.GetString("VersionDisponible") & " " & vNuevaVersion
        End If

        'Iniciar los Saldos Iniciales del Ejercicio
        IniciarSaldosIniciales(vAñoEjercicio)

        ' Mostramos el mensaje de "En Espera..." en la barra de estado mientras se carga todo
        If vAviso = True Then
            Me.TsLabelFormulario.Text = resManager.GetString("Aviso") & ": " & resManager.GetString("NoHayDatosHistoricos")
        Else
            Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
        End If

        'Buscamos si hay un Apunte Periódico que tenga fecha igual o anterior a la fecha de hoy y cuento los que hay.
        cmdMdb1cr.CommandText = "SELECT * FROM apuper"
        cmdMdb1cr.CommandText += " WHERE apuper.EjercicioAPP = " & vAñoEjercicio.ToString
        cmdMdb1cr.CommandText += " ORDER BY apuper.FechaAPP ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                vContador = 0
                While drMdb1.Read()
                    ' 1. Obtén la fecha directamente como objeto Date (sin pasar por String)
                    vDate1 = Convert.ToDateTime(drMdb1.GetValue(1))
                    ' 2. Asegúrate de que vfechaHoy sea un objeto Date (no String)
                    ' Si ya es Date, úsalo directamente. Si es String, convierte:
                    ' Dim dHoy As Date = Convert.ToDateTime(vfechaHoy)
                    ' 3. Compara los objetos Date directamente
                    ' Es mucho más limpio y preciso
                    If vDate1 <= vfechaHoy Then
                        vContador += 1
                    End If
                End While
            Else
                'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        drMdb1.Close()

        For i = 1 To vContador
            cmdMdb1cr.CommandText = "SELECT * FROM apuper"
            cmdMdb1cr.CommandText += " WHERE apuper.EjercicioAPP = " & vAñoEjercicio.ToString
            cmdMdb1cr.CommandText += " ORDER BY apuper.FechaAPP ASC"
            Try
                drMdb1 = cmdMdb1cr.ExecuteReader()
                If drMdb1.HasRows Then
                    While drMdb1.Read()
                        vCodigo = drMdb1.GetValue(0)
                        vConcepto = drMdb1.GetValue(2).ToString
                        vDescripcion = ApostrofePorAcentoAgudo(drMdb1.GetValue(3))
                        vImporte = drMdb1.GetValue(4).ToString
                        vNotas = drMdb1.GetValue(6).ToString
                        vCuenta = drMdb1.GetValue(7).ToString
                        If vDate1 <= DateTime.Today Then
                            drMdb1.Close()
                            ' 1. Diseñamos la estructura limpia para apuntes usando comodines '?'
                            vAñadirSql = "INSERT INTO apuntes (FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) " &
                                         "VALUES (?, ?, ?, ?, ?, ?, ?)"
                            cmdMdb1cr.CommandText = vAñadirSql

                            ' 2. Inyectamos los parámetros en el orden EXACTO de aparición del SQL
                            cmdMdb1cr.Parameters.Clear()

                            ' Fecha pura en binario (Evita que Windows confunda días con meses)
                            cmdMdb1cr.Parameters.AddWithValue("@FechaAPU", vDate1)

                            ' Cadenas de texto libres (Los parámetros limpian los apóstrofes solos)
                            cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vConcepto)
                            cmdMdb1cr.Parameters.AddWithValue("@DescripcionAPU", vDescripcion)

                            ' Importe blindado en formato Moneda nativa de Access
                            Dim paramImp As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@ImporteAPU", OleDb.OleDbType.Currency)
                            paramImp.Value = Math.Round(ConvertirDecimalSeguro(vImporte), 2)

                            ' Resto de campos del asiento contable
                            cmdMdb1cr.Parameters.AddWithValue("@EjercicioAPU", CInt(vAñoEjercicio))
                            cmdMdb1cr.Parameters.AddWithValue("@NotasAPU", vNotas)
                            cmdMdb1cr.Parameters.AddWithValue("@CuentaAPU", vCuenta)
                            cmdMdb1cr.CommandText = vAñadirSql
                            Try
                                cmdMdb1cr.ExecuteNonQuery()
                                MsgBox(vDate1 & vbNewLine & vConcepto & "     " & vDescripcion & "     " & vImporte & vbNewLine & rmse.GetString("CreadoCorrectamente"))
                            Catch ex As Exception
                                MsgBox(ex.ToString)
                            End Try
                            ' Eliminar Registro Apunte Periódico
                            vtipoSql = "DELETE FROM apuper"
                            vtipoSql += " WHERE apuper.CodigoAPP = " & vCodigo.ToString
                            cmdMdb1cr.CommandText = vtipoSql
                            Try
                                cmdMdb1cr.ExecuteNonQuery()
                                MsgBox(frmEditarApuntesPeriodicos.rmse.GetString("RegistroApuntePeriódicoBorrado"))
                            Catch ex As Exception
                                MsgBox(ex.ToString)
                            End Try
                            Exit While
                        End If
                    End While
                Else
                    'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            drMdb1.Close()
        Next
        drMdb1.Close()

        ' Traducimos este formulario
        ActualizarTextosFormulario(Me)
    End Sub

    Private Sub IP_Timer(ByVal sender As Object, ByVal e As EventArgs)
        Dim lFechaHora As Date = Date.Now
        Dim lsHora As String = lFechaHora.ToLongTimeString()
        Me.TsLabelHora.Text = lsHora
    End Sub

    Private Sub BtnApuntesContables_Click(sender As Object, e As EventArgs) Handles BtnApuntesContables.Click
        VerApuntesToolStripMenuItem.PerformClick()
    End Sub

    Private Sub VerApuntesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VerApuntesToolStripMenuItem.Click
        ' 1. Cambiar el texto de la barra de estado
        TsLabelFormulario.Text = rmse.GetString("VerApuntesToolStripMenuItem.Text")
        ' 2. Controlar la instancia física del formulario de forma tradicional
        If ((frmApuntesContables Is Nothing) OrElse (Not frmApuntesContables.IsHandleCreated)) Then
            frmApuntesContables = New ApuntesContables
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmApuntesContables)
        ' 4. Configurar el posicionamiento manual obligatorio
        frmApuntesContables.StartPosition = FormStartPosition.Manual
        ' 5. Calcular las coordenadas definitivas con tu ajuste de +10
        Dim posX As Integer = Me.Left + (Me.Width - frmApuntesContables.Width) \ 2
        Dim posY As Integer = Me.Top + 10
        If posX < 0 Then posX = 0
        ' 6. Fijar la ubicación calculada y abrir la ventana
        frmApuntesContables.Location = New System.Drawing.Point(posX, posY)
        frmApuntesContables.ShowDialog()
        ' 7. Destrucción explícita al cerrar
        frmApuntesContables.Dispose()
        ' 8. IMPORTANTE: Limpiar la variable manual para evitar el error de objeto destruido
        frmApuntesContables = Nothing
        ' 9. Restaurar el texto de espera de la barra
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub BtnIntroducirApuntes_Click(sender As Object, e As EventArgs) Handles BtnIntroducirApuntes.Click
        IntroducirApuntesToolStripMenuItem.PerformClick()
    End Sub

    Private Sub IntroducirApuntesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IntroducirApuntesToolStripMenuItem.Click
        TsLabelFormulario.Text = rmse.GetString("IntroducirApuntesToolStripMenuItem.Text")
        ' Comprobamos si existe un identificador asociado.
        If ((frmIntroApuntes Is Nothing) OrElse (Not frmIntroApuntes.IsHandleCreated)) Then
            frmIntroApuntes = New IntroApuntes
        End If
        ' Llamamos al formulario de manera modal.
        frmIntroApuntes.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmIntroApuntes.Dispose()
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
        Return
    End Sub

    Private Sub BtnTraspasoCuentas_Click(sender As Object, e As EventArgs) Handles BtnTraspasoCuentas.Click
        ' 🌟 LA CORRECCIÓN CLAVE: Buscamos de forma segura la ventana activa en el sistema
        ' Si la variable global es Nothing o se cerró, volvemos a instanciarla para que no dé el NullReference
        If frmApuntesContables Is Nothing OrElse frmApuntesContables.IsDisposed Then
            frmApuntesContables = New ApuntesContables
        End If

        ' Ahora que es 100% seguro que el objeto existe en la RAM, invocamos la apertura
        frmApuntesContables.BtnTraspasarRegistro_Click(Nothing, Nothing)
    End Sub

    Private Sub IntroducirTraspasosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IntroducirTraspasosToolStripMenuItem.Click
        ' 🌟 LA CORRECCIÓN CLAVE: Buscamos de forma segura la ventana activa en el sistema
        ' Si la variable global es Nothing o se cerró, volvemos a instanciarla para que no dé el NullReference
        If frmApuntesContables Is Nothing OrElse frmApuntesContables.IsDisposed Then
            frmApuntesContables = New ApuntesContables
        End If

        ' Ahora que es 100% seguro que el objeto existe en la RAM, invocamos la apertura
        frmApuntesContables.BtnTraspasarRegistro_Click(Nothing, Nothing)
    End Sub


    Private Sub BtnApuntesPeriodicos_Click(sender As Object, e As EventArgs) Handles BtnApuntesPeriodicos.Click
        ApuntesPeriodicosToolStripMenuItem.PerformClick()
    End Sub

    Private Sub ApuntesPeriodicosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ApuntesPeriodicosToolStripMenuItem.Click
        TsLabelFormulario.Text = rmse.GetString("ApuntesPeriodicosToolStripMenuItem.Text")
        ' Comprobamos si existe un identificador asociado.
        If ((frmApuntesPeriodicos Is Nothing) OrElse (Not frmApuntesPeriodicos.IsHandleCreated)) Then
            frmApuntesPeriodicos = New ApuntesPeriodicos
        End If
        ' Llamamos al formulario de manera NO modal.
        frmApuntesPeriodicos.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmApuntesPeriodicos.Dispose()
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub BtnConceptos_Click(sender As Object, e As EventArgs) Handles BtnConceptos.Click
        ConceptosContablesToolStripMenuItem.PerformClick()
    End Sub

    Private Sub ConceptosContablesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConceptosContablesToolStripMenuItem.Click
        ' 1. Cambiar el texto de la barra de estado
        TsLabelFormulario.Text = rmse.GetString("ConceptosContablesToolStripMenuItem.Text")
        ' 2. Controlar la instancia física del formulario de forma tradicional
        If ((frmConceptosContables Is Nothing) OrElse (Not frmConceptosContables.IsHandleCreated)) Then
            frmConceptosContables = New ConceptosContables
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmConceptosContables)
        ' 4. Configurar el posicionamiento manual obligatorio
        frmConceptosContables.StartPosition = FormStartPosition.Manual
        ' 5. Calcular las coordenadas definitivas con tu ajuste de +10
        Dim posX As Integer = Me.Left + (Me.Width - frmConceptosContables.Width) \ 2
        Dim posY As Integer = Me.Top + 10
        If posX < 0 Then posX = 0
        ' 6. Fijar la ubicación calculada y abrir la ventana
        frmConceptosContables.Location = New System.Drawing.Point(posX, posY)
        frmConceptosContables.ShowDialog()
        ' 7. Destrucción explícita al cerrar
        frmConceptosContables.Dispose()
        ' 8. IMPORTANTE: Limpiar la variable manual para evitar el error de objeto destruido
        frmConceptosContables = Nothing
        ' 9. Restaurar el texto de espera de la barra
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub BtnCuentasBancarias_Click(sender As Object, e As EventArgs) Handles BtnCuentasBancarias.Click
        CuentasToolStripMenuItem.PerformClick()
    End Sub

    Private Sub CuentasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CuentasToolStripMenuItem.Click
        ' 1. Cambiar el texto de la barra de estado
        TsLabelFormulario.Text = rmse.GetString("CuentasToolStripMenuItem.Text")
        ' 2. Controlar la instancia física del formulario de forma tradicional
        If ((frmCuentasBancarias Is Nothing) OrElse (Not frmCuentasBancarias.IsHandleCreated)) Then
            frmCuentasBancarias = New CuentasBancarias
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmCuentasBancarias)
        ' 4. Configurar el posicionamiento manual obligatorio
        frmCuentasBancarias.StartPosition = FormStartPosition.Manual
        ' 5. Calcular las coordenadas definitivas con tu ajuste de +10
        Dim posX As Integer = Me.Left + (Me.Width - frmCuentasBancarias.Width) \ 2
        Dim posY As Integer = Me.Top + 10
        If posX < 0 Then posX = 0
        ' 6. Fijar la ubicación calculada y abrir la ventana
        frmCuentasBancarias.Location = New System.Drawing.Point(posX, posY)
        frmCuentasBancarias.ShowDialog()
        ' 7. Destrucción explícita al cerrar
        frmCuentasBancarias.Dispose()
        ' 8. IMPORTANTE: Limpiar la variable manual para evitar el error de objeto destruido
        frmCuentasBancarias = Nothing
        ' 9. Restaurar el texto de espera de la barra
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub BtnTipoCuentas_Click(sender As Object, e As EventArgs) Handles BtnTipoCuentas.Click
        TiposDeCuentasToolStripMenuItem.PerformClick()
    End Sub

    Private Sub TiposDeCuentasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TiposDeCuentasToolStripMenuItem.Click
        ' 1. Cambiar el texto de la barra de estado
        TsLabelFormulario.Text = rmse.GetString("TiposDeCuentasToolStripMenuItem.Text")
        ' 2. Controlar la instancia física del formulario de forma tradicional
        If ((frmTipoCuentaBancaria Is Nothing) OrElse (Not frmTipoCuentaBancaria.IsHandleCreated)) Then
            frmTipoCuentaBancaria = New TipoCuentaBancaria
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmTipoCuentaBancaria)
        ' 4. Configurar el posicionamiento manual obligatorio
        frmTipoCuentaBancaria.StartPosition = FormStartPosition.Manual
        ' 5. Calcular las coordenadas definitivas con tu ajuste de +10
        Dim posX As Integer = Me.Left + (Me.Width - frmTipoCuentaBancaria.Width) \ 2
        Dim posY As Integer = Me.Top + 10
        If posX < 0 Then posX = 0
        ' 6. Fijar la ubicación calculada y abrir la ventana
        frmTipoCuentaBancaria.Location = New System.Drawing.Point(posX, posY)
        frmTipoCuentaBancaria.ShowDialog()
        ' 7. Destrucción explícita al cerrar
        frmTipoCuentaBancaria.Dispose()
        ' 8. IMPORTANTE: Limpiar la variable manual para evitar el error de objeto destruido
        frmTipoCuentaBancaria = Nothing
        ' 9. Restaurar el texto de espera de la barra
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub BtnPresupuestos_Click(sender As Object, e As EventArgs) Handles BtnPresupuestos.Click
        IntroducirDaToolStripMenuItem.PerformClick()
    End Sub

    Private Sub IntroducirDaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IntroducirDaToolStripMenuItem.Click
        ' 1. Cambiar el texto de la barra de estado
        TsLabelFormulario.Text = rmse.GetString("IntroducirDaToolStripMenuItem.Text")
        ' 2. Controlar la instancia física del formulario de forma tradicional
        If ((frmIntroPresupuestos Is Nothing) OrElse (Not frmIntroPresupuestos.IsHandleCreated)) Then
            frmIntroPresupuestos = New IntroPresupuestos
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmIntroPresupuestos)
        ' 4. Configurar el posicionamiento manual obligatorio
        frmIntroPresupuestos.StartPosition = FormStartPosition.Manual
        ' 5. Calcular las coordenadas definitivas con tu ajuste de +10
        Dim posX As Integer = Me.Left + (Me.Width - frmIntroPresupuestos.Width) \ 2
        Dim posY As Integer = Me.Top + 10
        If posX < 0 Then posX = 0
        ' 6. Fijar la ubicación calculada y abrir la ventana
        frmIntroPresupuestos.Location = New System.Drawing.Point(posX, posY)
        frmIntroPresupuestos.ShowDialog()
        ' 7. Destrucción explícita al cerrar
        frmIntroPresupuestos.Dispose()
        ' 8. IMPORTANTE: Limpiar la variable manual para evitar el error de objeto destruido
        frmIntroPresupuestos = Nothing
        ' 9. Restaurar el texto de espera de la barra
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub BtnDesviacionPresupuestos_Click(sender As Object, e As EventArgs) Handles BtnDesviacionPresupuestos.Click
        VerDesviaciónPresupuestosToolStripMenuItem.PerformClick()
    End Sub

    Private Sub VerDesviaciónPresupuestosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VerDesviaciónPresupuestosToolStripMenuItem.Click
        ' 1. Cambiar el texto de la barra de estado
        TsLabelFormulario.Text = rmse.GetString("VerDesviaciónPresupuestosToolStripMenuItem.Text")
        ' 2. Controlar la instancia física del formulario de forma tradicional
        If ((frmPresupuestos Is Nothing) OrElse (Not frmPresupuestos.IsHandleCreated)) Then
            frmPresupuestos = New Presupuestos
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmPresupuestos)
        ' 4. Configurar el posicionamiento manual obligatorio
        frmPresupuestos.StartPosition = FormStartPosition.Manual
        ' 5. Calcular las coordenadas definitivas con tu ajuste de +10
        Dim posX As Integer = Me.Left + (Me.Width - frmPresupuestos.Width) \ 2
        Dim posY As Integer = Me.Top + 10
        If posX < 0 Then posX = 0
        ' 6. Fijar la ubicación calculada y abrir la ventana
        frmPresupuestos.Location = New System.Drawing.Point(posX, posY)
        frmPresupuestos.ShowDialog()
        ' 7. Destrucción explícita al cerrar
        frmPresupuestos.Dispose()
        ' 8. IMPORTANTE: Limpiar la variable manual para evitar el error de objeto destruido
        frmPresupuestos = Nothing
        ' 9. Restaurar el texto de espera de la barra
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub OrdenadoPorFechasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OrdenadoPorFechasToolStripMenuItem.Click
        vTituloInforme = rmse.GetString("ListadoApuntesOrdenadoFechas") '1
        TsLabelFormulario.Text = rmse.GetString("ListadoApuntesOrdenadoFechas")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vOrdenadoPorFechasAPU = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vOrdenadoPorFechasAPU = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub OrdenadoPorConceptosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OrdenadoPorConceptosToolStripMenuItem.Click
        vTituloInforme = rmse.GetString("ListadoApuntesOrdenadoConceptos") '2
        TsLabelFormulario.Text = rmse.GetString("ListadoApuntesOrdenadoConceptos")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vOrdenadoPorConceptosAPU = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vOrdenadoPorConceptosAPU = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub OrdenadoporImportesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OrdenadoporImportesToolStripMenuItem.Click
        vTituloInforme = rmse.GetString("ListadoApuntesOrdenadoImportes") '3
        TsLabelFormulario.Text = rmse.GetString("ListadoApuntesOrdenadoImportes")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vOrdenadoPorImportesAPU = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vOrdenadoPorImportesAPU = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub SoloIngresosOrdenadoPorFechasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SoloIngresosOrdenadoPorFechasToolStripMenuItem.Click
        vTituloInforme = rmse.GetString("ListadoApuntesSoloIngresosFechas") '4
        TsLabelFormulario.Text = rmse.GetString("ListadoApuntesSoloIngresosFechas")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vSoloIngresosAPU = 1
        vOrdenadoPorFechasAPU = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vSoloIngresosAPU = 0
        vOrdenadoPorFechasAPU = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub SoloIngresosOrdenadoPorConceptosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SoloIngresosOrdenadoPorConceptosToolStripMenuItem.Click
        vTituloInforme = rmse.GetString("ListadoApuntesSoloIngresosConceptos") '5
        TsLabelFormulario.Text = rmse.GetString("ListadoApuntesSoloIngresosConceptos")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vSoloIngresosAPU = 1
        vOrdenadoPorConceptosAPU = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vSoloIngresosAPU = 0
        vOrdenadoPorConceptosAPU = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub SoloIngresosOrdenadoPorImportesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SoloIngresosOrdenadoPorImportesToolStripMenuItem.Click
        vTituloInforme = rmse.GetString("ListadoApuntesSoloIngresosImportes") '6
        TsLabelFormulario.Text = rmse.GetString("ListadoApuntesSoloIngresosImportes")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vSoloIngresosAPU = 1
        vOrdenadoPorImportesAPU = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vSoloIngresosAPU = 0
        vOrdenadoPorImportesAPU = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub SoloGastosOrdenadoPorFechasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SoloGastosOrdenadoPorFechasToolStripMenuItem.Click
        vTituloInforme = rmse.GetString("ListadoApuntesSoloGastosFechas") '7
        TsLabelFormulario.Text = rmse.GetString("ListadoApuntesSoloGastosFechas")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vSoloGastosAPU = 1
        vOrdenadoPorFechasAPU = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vSoloIngresosAPU = 0
        vOrdenadoPorFechasAPU = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub SoloGastosOrdenadoPorConceptosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SoloGastosOrdenadoPorConceptosToolStripMenuItem.Click
        vTituloInforme = rmse.GetString("ListadoApuntesSoloGastosConceptos")  '8
        TsLabelFormulario.Text = rmse.GetString("ListadoApuntesSoloGastosConceptos")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vSoloGastosAPU = 1
        vOrdenadoPorConceptosAPU = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vSoloIngresosAPU = 0
        vOrdenadoPorConceptosAPU = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub SoloGastosOrdenadoPorImportesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SoloGastosOrdenadoPorImportesToolStripMenuItem.Click
        vTituloInforme = rmse.GetString("ListadoApuntesSoloGastosImportes") '9
        TsLabelFormulario.Text = rmse.GetString("ListadoApuntesSoloGastosImportes")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vSoloGastosAPU = 1
        vOrdenadoPorImportesAPU = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vSoloIngresosAPU = 0
        vOrdenadoPorImportesAPU = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub OrdenadoPorFechasToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles OrdenadoPorFechasToolStripMenuItem1.Click
        vTituloInforme = rmse.GetString("ListadoPeriodicosFechas") '10
        TsLabelFormulario.Text = rmse.GetString("ListadoPeriodicosFechas")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vOrdenadoPorFechasAPP = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vOrdenadoPorFechasAPP = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub OrdenadoPorConceptosToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles OrdenadoPorConceptosToolStripMenuItem1.Click
        vTituloInforme = rmse.GetString("ListadoPeriodicosConceptos") '11
        TsLabelFormulario.Text = rmse.GetString("ListadoPeriodicosConceptos")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vOrdenadoPorConceptosAPP = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vOrdenadoPorConceptosAPP = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub OrdenadoPorImportesToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles OrdenadoPorImportesToolStripMenuItem2.Click
        vTituloInforme = rmse.GetString("ListadoPeriodicosImportes") '12
        TsLabelFormulario.Text = rmse.GetString("ListadoPeriodicosImportes")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vOrdenadoPorImportesAPP = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vOrdenadoPorImportesAPP = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub SoloIngresosOrdenadoPorFechasToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SoloIngresosOrdenadoPorFechasToolStripMenuItem1.Click
        vTituloInforme = rmse.GetString("ListadoPeriodicosSoloIngresosFechas") '13
        TsLabelFormulario.Text = rmse.GetString("ListadoPeriodicosSoloIngresosFechas")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vSoloIngresosAPP = 1
        vOrdenadoPorFechasAPP = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vSoloIngresosAPP = 0
        vOrdenadoPorFechasAPP = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub SoloIngresosOrdenadoPorConceptosToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SoloIngresosOrdenadoPorConceptosToolStripMenuItem1.Click
        vTituloInforme = rmse.GetString("ListadoPeriodicosSoloIngresosConceptos") '14
        TsLabelFormulario.Text = rmse.GetString("ListadoPeriodicosSoloIngresosConceptos")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vSoloIngresosAPP = 1
        vOrdenadoPorConceptosAPP = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vSoloIngresosAPP = 0
        vOrdenadoPorConceptosAPP = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub SoloIngresosOrdenadoPorImportesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SoloIngresosOrdenadoPorImportesToolStripMenuItem1.Click
        vTituloInforme = rmse.GetString("ListadoPeriodicosSoloIngresosImportes") '15
        TsLabelFormulario.Text = rmse.GetString("ListadoPeriodicosSoloIngresosImportes")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vSoloIngresosAPP = 1
        vOrdenadoPorImportesAPP = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vSoloIngresosAPP = 0
        vOrdenadoPorImportesAPP = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub SoloGastosOrdenadoPorFechasToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SoloGastosOrdenadoPorFechasToolStripMenuItem1.Click
        vTituloInforme = rmse.GetString("ListadoPeriodicosSoloGastosFechas")  '16
        TsLabelFormulario.Text = rmse.GetString("ListadoPeriodicosSoloGastosFechas")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vSoloGastosAPP = 1
        vOrdenadoPorFechasAPP = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vSoloIngresosAPP = 0
        vOrdenadoPorFechasAPP = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub SoloGastosOrdenadoPorConceptosToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SoloGastosOrdenadoPorConceptosToolStripMenuItem1.Click
        vTituloInforme = rmse.GetString("ListadoPeriodicosSoloGastosConceptos") '17
        TsLabelFormulario.Text = rmse.GetString("ListadoPeriodicosSoloGastosConceptos")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vSoloGastosAPP = 1
        vOrdenadoPorConceptosAPP = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vSoloIngresosAPP = 0
        vOrdenadoPorConceptosAPP = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub SoloGastosOrdenadoPorImportesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SoloGastosOrdenadoPorImportesToolStripMenuItem1.Click
        vTituloInforme = rmse.GetString("ListadoPeriodicosSoloGastosImportes") '18
        TsLabelFormulario.Text = rmse.GetString("ListadoPeriodicosSoloGastosImportes")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionFechas Is Nothing) OrElse (Not frmSeleccionFechas.IsHandleCreated)) Then
            frmSeleccionFechas = New SeleccionFechas
        End If
        ' Llamamos al formulario de manera modal.
        vSoloGastosAPP = 1
        vOrdenadoPorImportesAPP = 1
        frmSeleccionFechas.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionFechas.Dispose()
        vSoloIngresosAPP = 0
        vOrdenadoPorImportesAPP = 0
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub GráficosDeIngresosPorConceptoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GráficosDeIngresosPorConceptoToolStripMenuItem.Click
        TsLabelFormulario.Text = rmse.GetString("GráficosDeIngresosPorConceptoToolStripMenuItem.Text")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionarDatosIngresos Is Nothing) OrElse (Not frmSeleccionarDatosIngresos.IsHandleCreated)) Then
            frmSeleccionarDatosIngresos = New SeleccionDatosIngresos
        End If
        ' Llamamos al formulario de manera modal.
        vGraficoSolo = "IngresosPorConcepto"
        frmSeleccionarDatosIngresos.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionarDatosIngresos.Dispose()
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub GráficosDeGastosPorConceptoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GráficosDeGastosPorConceptoToolStripMenuItem.Click
        TsLabelFormulario.Text = rmse.GetString("GráficosDeGastosPorConceptoToolStripMenuItem.Text")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionarDatosGastos Is Nothing) OrElse (Not frmSeleccionarDatosGastos.IsHandleCreated)) Then
            frmSeleccionarDatosGastos = New SeleccionDatosGastos
        End If
        ' Llamamos al formulario de manera modal.
        vGraficoSolo = "GastosPorConcepto"
        frmSeleccionarDatosGastos.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionarDatosGastos.Dispose()
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub BtnCompactarBaseDatos_Click(sender As Object, e As EventArgs) Handles BtnCompactarBaseDatos.Click
        CompactarBaseDeDatosToolStripMenuItem.PerformClick()
    End Sub

    Private Sub CompactarBaseDeDatosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CompactarBaseDeDatosToolStripMenuItem.Click
        ' Si no existe la carpeta de BackUp la creamos.
        Dim path As String = "C:\ContaHogar3.0\Backup"
        If Directory.Exists(path) Then
            'MsgBox("Ya existe la Ruta C:\ContaHogar3.0\Backup.")
        Else
            Directory.CreateDirectory(path)
            'MsgBox("Ruta C:\ContaHogar3.0\Backup, Creada.")
        End If
        Dim jetEng As JRO.JetEngine
        jetEng = New JRO.JetEngine()
        vCompactada = "C:\ContaHogar3.0\Backup\contahogarcompacted.mdb"
        If File.Exists(vCompactada) Then
            File.Delete(vCompactada)
        End If
        Try
            conexion1.Close()
            jetEng.CompactDatabase("Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & vRuta, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vCompactada & ";Jet OLEDB:Engine Type=5")
            FileCopy(vCompactada, vRuta)
            tipoDsn = "AccessMdb" ' Se conecta a Mdb
            Conectarse(tipoDsn)
            MessageBox.Show(rmse.GetString("CompactacionOk"), rmse.GetString("Compactar"), MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MsgBox(rmse.GetString("CompactarError") & ":  " & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Error"))
        End Try
    End Sub

    Private Sub BtnActivarSoftware_Click(sender As Object, e As EventArgs) Handles BtnActivarSoftware.Click
        BtnActivarSoftwareToolStripMenuItem.PerformClick()
    End Sub

    Private Sub BtnActivarSoftwareToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BtnActivarSoftwareToolStripMenuItem.Click
        ' 1. Cambiar el texto de la barra de estado
        TsLabelFormulario.Text = rmse.GetString("BtnActivarSoftwareToolStripMenuItem.Text")
        ' 2. Controlar la instancia física del formulario de forma tradicional
        If ((frmActivarSoftware Is Nothing) OrElse (Not frmActivarSoftware.IsHandleCreated)) Then
            frmActivarSoftware = New ActivarSoftware
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmActivarSoftware)
        ' 4. Configurar el posicionamiento manual obligatorio
        frmActivarSoftware.StartPosition = FormStartPosition.Manual
        ' 5. Calcular las coordenadas definitivas con tu ajuste de +10
        Dim posX As Integer = Me.Left + (Me.Width - frmActivarSoftware.Width) \ 2
        Dim posY As Integer = Me.Top + 10
        If posX < 0 Then posX = 0
        ' 6. Fijar la ubicación calculada y abrir la ventana
        frmActivarSoftware.Location = New System.Drawing.Point(posX, posY)
        frmActivarSoftware.ShowDialog()
        ' 7. Destrucción explícita al cerrar
        frmActivarSoftware.Dispose()
        ' 8. IMPORTANTE: Limpiar la variable manual para evitar el error de objeto destruido
        frmActivarSoftware = Nothing
        ' 9. Restaurar el texto de espera de la barra
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
        If vActivado Then
            LblNotificacion.Visible = False
        End If
    End Sub

    Private Sub BtnImportarContaHogar_Click(sender As Object, e As EventArgs) Handles BtnImportarContaHogar.Click
        ImportaAntiguoContahogarToolStripMenuItem.PerformClick()
    End Sub

    Private Sub ImportaAntiguoContahogarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportaAntiguoContahogarToolStripMenuItem.Click
        MsgBox(rmse.GetString("MsgImportar1"), MsgBoxStyle.Information, rmse.GetString("ImportarContahogar"))
        ' Si no existe la carpeta la creamos y salimos del Sub.
        Dim RutaArchivo As String = "C:\ContaHogar3.0\CHDB2.mdb"
        If File.Exists(RutaArchivo) Then
            MsgBox(rmse.GetString("ExisteCHDB2"))
        Else
            MsgBox(rmse.GetString("NoExisteCHDB2"))
            If Directory.Exists("C:\ContaHogar3.0") Then
                MsgBox(rmse.GetString("YaExisteRuta"))
            Else
                Directory.CreateDirectory("C:\ContaHogar3.0")
                MsgBox(rmse.GetString("RutaCreada"))
            End If
            Exit Sub
        End If

        TsLabelFormulario.ForeColor = Color.Red
        TsLabelFormulario.Text = rmse.GetString("ImportandoApuntes")
        Dim respuesta As MsgBoxResult = MsgBox(rmse.GetString("RealizarCopiaSeguridad"), vbQuestion + vbYesNo + vbDefaultButton2, rmse.GetString("$this.Text"))
        If respuesta = vbYes Then
            BtnCopiaSeguridad.PerformClick()
        Else
            If MsgBox(rmse.GetString("RecomiendaCopiaSeguridad"), vbExclamation + vbYesNo + vbDefaultButton2, rmse.GetString("Advertencia")) = vbNo Then
                TsLabelFormulario.ForeColor = Color.Black
                Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
                Exit Sub
            End If
        End If
        ' --- CONFIGURACIÓN DE RUTAS ---
        Dim rutaOrigen As String = "C:\ContaHogar3.0\CHDB2.mdb"
        Dim connOrigenString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & rutaOrigen & ";"
        Dim connDestinoString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vRuta & ";"

        Using connOrigen As New OleDbConnection(connOrigenString)
            Using connDestino As New OleDbConnection(connDestinoString)
                Try
                    connOrigen.Open()
                    connDestino.Open()

                    Dim sqlSelectOrigenApuntes As String = "SELECT FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU FROM APUNTES"
                    Using cmdOrigen As New OleDbCommand(sqlSelectOrigenApuntes, connOrigen)
                        Using reader As OleDbDataReader = cmdOrigen.ExecuteReader()
                            ' Preparar comandos para destino (se reutilizan)
                            Dim sqlCheckApuntes As String = "SELECT COUNT(*) FROM APUNTES WHERE FechaAPU = ? AND ConceptoAPU = ? AND DescripcionAPU = ? AND ImporteAPU = ? AND EjercicioAPU = ? AND NotasAPU = ? AND CuentaAPU = ?"
                            Using cmdCheck As New OleDbCommand(sqlCheckApuntes, connDestino)
                                cmdCheck.Parameters.Add("?", OleDbType.Date)
                                cmdCheck.Parameters.Add("?", OleDbType.VarChar)
                                cmdCheck.Parameters.Add("?", OleDbType.VarChar)
                                cmdCheck.Parameters.Add("?", OleDbType.Currency)
                                cmdCheck.Parameters.Add("?", OleDbType.Integer)
                                cmdCheck.Parameters.Add("?", OleDbType.VarChar)
                                cmdCheck.Parameters.Add("?", OleDbType.VarChar)

                                Dim sqlInsert As String = "INSERT INTO APUNTES (FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) VALUES (?, ?, ?, ?, ?, ?, ?)"
                                Using cmdInsert As New OleDbCommand(sqlInsert, connDestino)
                                    cmdInsert.Parameters.Add("?", OleDbType.Date)
                                    cmdInsert.Parameters.Add("?", OleDbType.VarChar)
                                    cmdInsert.Parameters.Add("?", OleDbType.VarChar)
                                    cmdInsert.Parameters.Add("?", OleDbType.Currency)
                                    cmdInsert.Parameters.Add("?", OleDbType.Integer)
                                    cmdInsert.Parameters.Add("?", OleDbType.VarChar)
                                    cmdInsert.Parameters.Add("?", OleDbType.VarChar)

                                    Dim contador As Integer = 0

                                    While reader.Read()
                                        ' Leer valores (comprobando DBNull)
                                        Dim vFecha As Object = If(reader.IsDBNull(0), DBNull.Value, reader.GetValue(0))
                                        Dim vConcepto As Object = If(reader.IsDBNull(1), DBNull.Value, reader.GetValue(1))
                                        Dim vDescripcion As Object = If(reader.IsDBNull(2), DBNull.Value, reader.GetValue(2))
                                        Dim vImporte As Object = If(reader.IsDBNull(3), DBNull.Value, reader.GetValue(3))
                                        Dim vEjercicio As Object = If(reader.IsDBNull(4), DBNull.Value, reader.GetValue(4))
                                        Dim vNotas As Object = If(reader.IsDBNull(5), DBNull.Value, reader.GetValue(5))
                                        Dim vCuenta As Object = If(reader.IsDBNull(6), DBNull.Value, reader.GetValue(6))

                                        ' Comprobar existencia en destino
                                        cmdCheck.Parameters(0).Value = vFecha
                                        cmdCheck.Parameters(1).Value = vConcepto
                                        cmdCheck.Parameters(2).Value = vDescripcion
                                        cmdCheck.Parameters(3).Value = vImporte
                                        cmdCheck.Parameters(4).Value = vEjercicio
                                        cmdCheck.Parameters(5).Value = vNotas
                                        cmdCheck.Parameters(6).Value = vCuenta

                                        Dim existe As Integer = Convert.ToInt32(cmdCheck.ExecuteScalar())

                                        If existe = 0 Then
                                            ' Insertar en destino
                                            cmdInsert.Parameters(0).Value = vFecha
                                            cmdInsert.Parameters(1).Value = vConcepto
                                            cmdInsert.Parameters(2).Value = vDescripcion
                                            cmdInsert.Parameters(3).Value = vImporte
                                            cmdInsert.Parameters(4).Value = vEjercicio
                                            cmdInsert.Parameters(5).Value = vNotas
                                            cmdInsert.Parameters(6).Value = vCuenta

                                            cmdInsert.ExecuteNonQuery()
                                            contador += 1
                                        End If
                                    End While
                                    MsgBox(rmse.GetString("TransferenciaApuntes") & ". " & contador.ToString() & " " & rmse.GetString("RegistrosCopiados"), MsgBoxStyle.Information, rmse.GetString("$this.Text"))
                                End Using
                            End Using
                        End Using
                    End Using
                Catch ex As Exception
                    MsgBox(rmse.GetString("ErrorTransferenciaApuntes") & ": " & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Error"))
                End Try
            End Using
        End Using

        TsLabelFormulario.ForeColor = Color.Red
        TsLabelFormulario.Text = rmse.GetString("ImportandoApuntesPeriodicos")
        Using connOrigen As New OleDbConnection(connOrigenString)
            Using connDestino As New OleDbConnection(connDestinoString)
                Try
                    connOrigen.Open()
                    connDestino.Open()

                    Dim sqlSelectOrigenApuntes As String = "SELECT FechaAPP, ConceptoAPP, DescripcionAPP, ImporteAPP, EjercicioAPP, NotasAPP, CuentaAPP FROM APUPER"
                    Using cmdOrigen As New OleDbCommand(sqlSelectOrigenApuntes, connOrigen)
                        Using reader As OleDbDataReader = cmdOrigen.ExecuteReader()
                            ' Preparar comandos para destino (se reutilizan)
                            Dim sqlCheckApuntes As String = "SELECT COUNT(*) FROM APUPER WHERE FechaAPP = ? AND ConceptoAPP = ? AND DescripcionAPP = ? AND ImporteAPP = ? AND EjercicioAPP = ? AND NotasAPP = ? AND CuentaAPP = ?"
                            Using cmdCheck As New OleDbCommand(sqlCheckApuntes, connDestino)
                                cmdCheck.Parameters.Add("?", OleDbType.Date)
                                cmdCheck.Parameters.Add("?", OleDbType.VarChar)
                                cmdCheck.Parameters.Add("?", OleDbType.VarChar)
                                cmdCheck.Parameters.Add("?", OleDbType.Double)
                                cmdCheck.Parameters.Add("?", OleDbType.Integer)
                                cmdCheck.Parameters.Add("?", OleDbType.VarChar)
                                cmdCheck.Parameters.Add("?", OleDbType.VarChar)

                                Dim sqlInsert As String = "INSERT INTO APUNTES (FechaAPP, ConceptoAPP, DescripcionAPP, ImporteAPP, EjercicioAPP, NotasAPP, CuentaAPP) VALUES (?, ?, ?, ?, ?, ?, ?)"
                                Using cmdInsert As New OleDbCommand(sqlInsert, connDestino)
                                    cmdInsert.Parameters.Add("?", OleDbType.Date)
                                    cmdInsert.Parameters.Add("?", OleDbType.VarChar)
                                    cmdInsert.Parameters.Add("?", OleDbType.VarChar)
                                    cmdInsert.Parameters.Add("?", OleDbType.Double)
                                    cmdInsert.Parameters.Add("?", OleDbType.Integer)
                                    cmdInsert.Parameters.Add("?", OleDbType.VarChar)
                                    cmdInsert.Parameters.Add("?", OleDbType.VarChar)

                                    Dim contador As Integer = 0

                                    While reader.Read()
                                        ' Leer valores (comprobando DBNull)
                                        Dim vFecha As Object = If(reader.IsDBNull(0), DBNull.Value, reader.GetValue(0))
                                        Dim vConcepto As Object = If(reader.IsDBNull(1), DBNull.Value, reader.GetValue(1))
                                        Dim vDescripcion As Object = If(reader.IsDBNull(2), DBNull.Value, reader.GetValue(2))
                                        Dim vImporte As Object = If(reader.IsDBNull(3), DBNull.Value, reader.GetValue(3))
                                        Dim vEjercicio As Object = If(reader.IsDBNull(4), DBNull.Value, reader.GetValue(4))
                                        Dim vNotas As Object = If(reader.IsDBNull(5), DBNull.Value, reader.GetValue(5))
                                        Dim vCuenta As Object = If(reader.IsDBNull(6), DBNull.Value, reader.GetValue(6))

                                        ' Comprobar existencia en destino
                                        cmdCheck.Parameters(0).Value = vFecha
                                        cmdCheck.Parameters(1).Value = vConcepto
                                        cmdCheck.Parameters(2).Value = vDescripcion
                                        cmdCheck.Parameters(3).Value = vImporte
                                        cmdCheck.Parameters(4).Value = vEjercicio
                                        cmdCheck.Parameters(5).Value = vNotas
                                        cmdCheck.Parameters(6).Value = vCuenta

                                        Dim existe As Integer = Convert.ToInt32(cmdCheck.ExecuteScalar())

                                        If existe = 0 Then
                                            ' Insertar en destino
                                            cmdInsert.Parameters(0).Value = vFecha
                                            cmdInsert.Parameters(1).Value = vConcepto
                                            cmdInsert.Parameters(2).Value = vDescripcion
                                            cmdInsert.Parameters(3).Value = vImporte
                                            cmdInsert.Parameters(4).Value = vEjercicio
                                            cmdInsert.Parameters(5).Value = vNotas
                                            cmdInsert.Parameters(6).Value = vCuenta

                                            cmdInsert.ExecuteNonQuery()
                                            contador += 1
                                        End If
                                    End While
                                    MsgBox(rmse.GetString("TransferenciaApuntesPeriodicos") & ". " & contador.ToString() & " " & rmse.GetString("RegistrosCopiados"), MsgBoxStyle.Information, rmse.GetString("$this.Text"))
                                End Using
                            End Using
                        End Using
                    End Using
                Catch ex As Exception
                    MsgBox(rmse.GetString("ErrorTransferenciaApuntesPeriodicos") & ": " & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Error"))
                End Try
            End Using
        End Using

        TsLabelFormulario.ForeColor = Color.Red
        TsLabelFormulario.Text = rmse.GetString("ImportandoEjercicios")

        Dim sqlSeleccion As String = "SELECT EjercicioEJE FROM EJERCICIOS"
        Dim sqlVerificar As String = "SELECT COUNT(*) FROM EJERCICIOS WHERE EjercicioEJE = ?"
        Dim sqlInsercion As String = "INSERT INTO EJERCICIOS (EjercicioEJE) VALUES (?)"
        Using connOrigen As New OleDbConnection(connOrigenString), connDestino As New OleDbConnection(connDestinoString)
            Try
                connOrigen.Open()
                connDestino.Open()

                Dim cmdOrigen As New OleDbCommand(sqlSeleccion, connOrigen)
                Dim reader As OleDbDataReader = cmdOrigen.ExecuteReader()

                ' Comando para verificar existencia
                Dim cmdCheck As New OleDbCommand(sqlVerificar, connDestino)
                cmdCheck.Parameters.Add("?", OleDbType.Integer) ' Cambia el tipo si no es Número

                ' Comando para insertar
                Dim cmdInsert As New OleDbCommand(sqlInsercion, connDestino)
                cmdInsert.Parameters.Add("?", OleDbType.Integer)

                Dim insertados As Integer = 0
                Dim omitidos As Integer = 0

                While reader.Read()
                    Dim valorActual = reader("EjercicioEJE")

                    ' 1. Verificar si ya existe en el destino
                    cmdCheck.Parameters(0).Value = valorActual
                    Dim existe As Integer = CInt(cmdCheck.ExecuteScalar())

                    ' 2. Si no existe (count = 0), insertar
                    If existe = 0 Then
                        cmdInsert.Parameters(0).Value = valorActual
                        cmdInsert.ExecuteNonQuery()
                        insertados += 1
                    Else
                        omitidos += 1
                    End If
                End While
                reader.Close()
                MsgBox(rmse.GetString("TransferenciaEjercicios") & ". " & insertados.ToString() & " " & rmse.GetString("RegistrosCopiados") & ", " & omitidos.ToString() & " " & rmse.GetString("RegistrosOmitidos") & ".", MsgBoxStyle.Information, rmse.GetString("$this.Text"))
            Catch ex As Exception
                MsgBox(rmse.GetString("ErrorTransferenciaEjercicios") & ":  " & ex.Message, MsgBoxStyle.Critical, rmse.GetString("Error"))
            End Try
        End Using

        TsLabelFormulario.ForeColor = Color.Red
        TsLabelFormulario.Text = rmse.GetString("ImportandoConceptos")

        Dim sqlSeleccionCON As String = "SELECT CodigoCON, DescripcionCON, TipoCON, NotasCON FROM CONCEPTOS"
        Dim sqlVerificarCON As String = "SELECT COUNT(*) FROM CONCEPTOS WHERE CodigoCON = ?"
        Dim sqlInsercionCON As String = "INSERT INTO CONCEPTOS (CodigoCON, DescripcionCON, TipoCON, NotasCON) VALUES (?, ?, ?, ?)"

        ' Consultas para la validación y borrado en el DESTINO
        Dim sqlConceptosDestino As String = "SELECT CodigoCON FROM CONCEPTOS"
        Dim sqlCheckApuntesDestino As String = "SELECT COUNT(*) FROM APUNTES WHERE ConceptoAPU = ?"
        Dim sqlCheckApuperDestino As String = "SELECT COUNT(*) FROM APUPER WHERE ConceptoAPP = ?"
        Dim sqlEliminarDestino As String = "DELETE FROM CONCEPTOS WHERE CodigoCON = ?"

        Dim conceptosMDBAplicacion As New System.Collections.Generic.List(Of String)()
        Dim conceptosAEliminarDestino As New System.Collections.Generic.List(Of String)()

        Using connOrigen As New OleDbConnection(connOrigenString), connDestino As New OleDbConnection(connDestinoString)
            Try
                connOrigen.Open()
                connDestino.Open()

                ' === FASE 1: IMPORTACIÓN (ORIGEN A DESTINO) ===
                Dim cmdOrigen As New OleDbCommand(sqlSeleccionCON, connOrigen)
                Dim reader As OleDbDataReader = cmdOrigen.ExecuteReader()

                Dim cmdCheck As New OleDbCommand(sqlVerificarCON, connDestino)
                cmdCheck.Parameters.Add("?", OleDbType.VarChar)

                Dim cmdInsert As New OleDbCommand(sqlInsercionCON, connDestino)
                For i As Integer = 1 To 4
                    cmdInsert.Parameters.Add("?", OleDbType.VarChar)
                Next

                Dim insertados As Integer = 0
                Dim omitidos As Integer = 0

                While reader.Read()
                    Dim codigoActual As String = reader("CodigoCON").ToString()
                    cmdCheck.Parameters(0).Value = codigoActual

                    If CInt(cmdCheck.ExecuteScalar()) = 0 Then
                        cmdInsert.Parameters(0).Value = codigoActual
                        cmdInsert.Parameters(1).Value = reader("DescripcionCON").ToString()
                        cmdInsert.Parameters(2).Value = reader("TipoCON").ToString()
                        cmdInsert.Parameters(3).Value = reader("NotasCON").ToString()
                        cmdInsert.ExecuteNonQuery()
                        insertados += 1
                    Else
                        omitidos += 1
                    End If
                End While
                reader.Close()

                ' === FASE 2: ENCONTRAR MUESTRAS EN EL DESTINO ===
                Dim cmdTodosDestino As New OleDbCommand(sqlConceptosDestino, connDestino)
                Dim readerDestino As OleDbDataReader = cmdTodosDestino.ExecuteReader()
                While readerDestino.Read()
                    Dim codDestino As String = readerDestino("CodigoCON").ToString().ToUpper()

                    ' FILTRO SEGURO: Solo nos interesa si es uno de tus 33 conceptos de fábrica originales
                    If ConceptosMuestraSistema.Contains(codDestino) Then
                        conceptosMDBAplicacion.Add(codDestino)
                    End If
                End While
                readerDestino.Close()

                ' === FASE 3: COMPROBAR APUNTES EN EL DESTINO ===
                Dim cmdApuntesDestino As New OleDbCommand(sqlCheckApuntesDestino, connDestino)
                cmdApuntesDestino.Parameters.Add("?", OleDbType.VarChar)

                Dim cmdApuperDestino As New OleDbCommand(sqlCheckApuperDestino, connDestino)
                cmdApuperDestino.Parameters.Add("?", OleDbType.VarChar)

                For Each codDestino As String In conceptosMDBAplicacion
                    cmdApuntesDestino.Parameters(0).Value = codDestino
                    cmdApuperDestino.Parameters(0).Value = codDestino

                    Dim totalApuntes As Integer = CInt(cmdApuntesDestino.ExecuteScalar()) + CInt(cmdApuperDestino.ExecuteScalar())

                    ' Si es de fábrica Y nadie lo ha usado en el destino, va a la lista de borrado
                    If totalApuntes = 0 Then
                        conceptosAEliminarDestino.Add(codDestino)
                    End If
                Next

                ' === FASE 4: BORRADO SEGURO DE LAS MUESTRAS VACÍAS ===
                If conceptosAEliminarDestino.Count > 0 Then
                    Using cmdDeleteDestino As New OleDbCommand(sqlEliminarDestino, connDestino)
                        cmdDeleteDestino.Parameters.Add("?", OleDbType.VarChar)
                        For Each codBorrar As String In conceptosAEliminarDestino
                            cmdDeleteDestino.Parameters(0).Value = codBorrar
                            cmdDeleteDestino.ExecuteNonQuery()
                        Next
                    End Using
                End If

                MsgBox(rmse.GetString("TransferenciaConceptos") & ". " & insertados.ToString() & " " & rmse.GetString("RegistrosCopiados") & ", " & omitidos.ToString() & " " & rmse.GetString("RegistrosOmitidos") & ".", MsgBoxStyle.Information, rmse.GetString("$this.Text"))
            Catch ex As Exception
                MsgBox(resManager.GetString("Error") & ":   " & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Error"))
            End Try
        End Using

        TsLabelFormulario.ForeColor = Color.Red
        TsLabelFormulario.Text = rmse.GetString("ImportandoCuentas")
        ' --- CONFIGURACIÓN DE CUENTAS ---
        ' Cambia Campo1, Campo2, etc., por los nombres reales de tus columnas
        Dim sqlSeleccionCUE As String = "SELECT NombreCUE, NumeroCUE, TipoCUE, NotasCUE FROM CUENTAS"

        ' Verificamos duplicidad solo por el primer campo
        Dim sqlVerificarCUE As String = "SELECT COUNT(*) FROM CUENTAS WHERE NombreCUE = ?"

        ' Insertamos en los 4 campos de la tabla destino
        Dim sqlInsercionCUE As String = "INSERT INTO CUENTAS (NombreCUE, NumeroCUE, TipoCUE, NotasCUE) VALUES (?, ?, ?, ?)"

        Using connOrigen As New OleDbConnection(connOrigenString), connDestino As New OleDbConnection(connDestinoString)
            Try
                connOrigen.Open()
                connDestino.Open()
                Dim cmdOrigen As New OleDbCommand(sqlSeleccionCUE, connOrigen)
                Dim reader As OleDbDataReader = cmdOrigen.ExecuteReader()
                ' Configurar comando de verificación (solo 1 parámetro de texto)
                Dim cmdCheck As New OleDbCommand(sqlVerificarCUE, connDestino)
                cmdCheck.Parameters.Add("?", OleDbType.VarChar)
                ' Configurar comando de inserción (4 parámetros de texto)
                Dim cmdInsert As New OleDbCommand(sqlInsercionCUE, connDestino)
                For i As Integer = 1 To 4
                    cmdInsert.Parameters.Add("?", OleDbType.VarChar)
                Next
                Dim insertados As Integer = 0
                Dim omitidos As Integer = 0
                While reader.Read()
                    ' 1. Validar duplicado usando solo el primer campo
                    cmdCheck.Parameters(0).Value = reader("NombreCUE").ToString()
                    ' 2. Si el conteo es 0, no existe, procedemos a insertar
                    If CInt(cmdCheck.ExecuteScalar()) = 0 Then
                        cmdInsert.Parameters(0).Value = reader("NombreCUE").ToString()
                        cmdInsert.Parameters(1).Value = reader("NumeroCUE").ToString()
                        cmdInsert.Parameters(2).Value = reader("TipoCUE").ToString()
                        cmdInsert.Parameters(3).Value = reader("NotasCUE").ToString()
                        cmdInsert.ExecuteNonQuery()
                        insertados += 1
                    Else
                        omitidos += 1
                    End If
                End While
                reader.Close()
                MsgBox(rmse.GetString("TransferenciaCuentas") & ". " & insertados.ToString() & " " & rmse.GetString("RegistrosCopiados") & ", " & omitidos.ToString() & " " & rmse.GetString("RegistrosOmitidos") & ".", MsgBoxStyle.Information, rmse.GetString("$this.Text"))
            Catch ex As Exception
                MsgBox(rmse.GetString("ErrorTransferenciaCuentas") & ":  " & ex.Message, MsgBoxStyle.Critical, rmse.GetString("Error"))
            End Try
        End Using

        TsLabelFormulario.ForeColor = Color.Red
        TsLabelFormulario.Text = rmse.GetString("ImportandoPresupuestos")
        ' --- CONSULTAS CON TIPOS MIXTOS ---
        Dim sqlSeleccionPRE As String = "SELECT [ConceptoPRE], [ImportePRE], [FDesdePRE], [EjercicioPRE] FROM PRESUPUESTO"

        ' Verificación de todos los campos para evitar duplicados exactos
        Dim sqlVerificarPRE As String = "SELECT COUNT(*) FROM PRESUPUESTO WHERE [ConceptoPRE]=? AND [ImportePRE]=? AND [FDesdePRE]=? AND [EjercicioPRE]=?"

        Dim sqlInsercionPRE As String = "INSERT INTO PRESUPUESTO ([ConceptoPRE], [ImportePRE], [FDesdePRE], [EjercicioPRE]) VALUES (?, ?, ?, ?)"

        Using connOrigen As New OleDbConnection(connOrigenString), connDestino As New OleDbConnection(connDestinoString)
            Try
                connOrigen.Open()
                connDestino.Open()
                Dim cmdOrigen As New OleDbCommand(sqlSeleccionPRE, connOrigen)
                Dim reader As OleDbDataReader = cmdOrigen.ExecuteReader()

                ' --- CONFIGURAR PARÁMETROS (Orden y Tipo son críticos) ---
                Dim cmdCheck As New OleDbCommand(sqlVerificarPRE, connDestino)
                cmdCheck.Parameters.Add("@p1", OleDbType.VarChar)  ' Concepto
                cmdCheck.Parameters.Add("@p2", OleDbType.Double)   ' Importe (Número/Double)
                cmdCheck.Parameters.Add("@p3", OleDbType.Date)     ' FDesde (Fecha)
                cmdCheck.Parameters.Add("@p4", OleDbType.Integer)  ' Ejercicio (Número)
                Dim cmdInsert As New OleDbCommand(sqlInsercionPRE, connDestino)
                cmdInsert.Parameters.Add("@p1", OleDbType.VarChar)
                cmdInsert.Parameters.Add("@p2", OleDbType.Double)
                cmdInsert.Parameters.Add("@p3", OleDbType.Date)
                cmdInsert.Parameters.Add("@p4", OleDbType.Integer)
                Dim insertados As Integer = 0
                Dim omitidos As Integer = 0
                While reader.Read()
                    ' Extraer valores del Reader
                    Dim vConcepto = If(IsDBNull(reader("ConceptoPRE")), "", reader("ConceptoPRE").ToString())
                    Dim vImporte = If(IsDBNull(reader("ImportePRE")), 0, CDbl(reader("ImportePRE")))
                    Dim vFecha = If(IsDBNull(reader("FDesdePRE")), #1/1/1900#, CDate(reader("FDesdePRE")))
                    Dim vEjercicio = If(IsDBNull(reader("EjercicioPRE")), 0, CInt(reader("EjercicioPRE")))
                    ' 1. Asignar al verificador
                    cmdCheck.Parameters(0).Value = vConcepto
                    cmdCheck.Parameters(1).Value = vImporte
                    cmdCheck.Parameters(2).Value = vFecha
                    cmdCheck.Parameters(3).Value = vEjercicio
                    ' 2. Si no existe la combinación exacta, insertar
                    If CInt(cmdCheck.ExecuteScalar()) = 0 Then
                        cmdInsert.Parameters(0).Value = vConcepto
                        cmdInsert.Parameters(1).Value = vImporte
                        cmdInsert.Parameters(2).Value = vFecha
                        cmdInsert.Parameters(3).Value = vEjercicio
                        cmdInsert.ExecuteNonQuery()
                        insertados += 1
                    Else
                        omitidos += 1
                    End If
                End While
                reader.Close()
                MsgBox(rmse.GetString("TransferenciaPresupuestos") & ". " & insertados.ToString() & " " & rmse.GetString("RegistrosCopiados") & ", " & omitidos.ToString() & " " & rmse.GetString("RegistrosOmitidos") & ".", MsgBoxStyle.Information, rmse.GetString("$this.Text"))
            Catch ex As Exception
                MsgBox(rmse.GetString("ErrorTransferenciaPresupuestos") & ": " & ex.Message, MsgBoxStyle.Critical, rmse.GetString("Error"))
            End Try
        End Using
        TsLabelFormulario.ForeColor = Color.Black
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub BtnCambiarEjercicioActivo_Click(sender As Object, e As EventArgs) Handles BtnCambiarEjercicioActivo.Click
        CambiarEjercicioActivoToolStripMenuItem.PerformClick()
    End Sub

    Private Sub CambiarEjercicioActivoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CambiarEjercicioActivoToolStripMenuItem.Click
        TsLabelFormulario.Text = rmse.GetString("CambiarEjercicioActivoToolStripMenuItem.Text")
        ' Comprobamos si existe un identificador asociado.
        If ((frmSeleccionEjercicio Is Nothing) OrElse (Not frmSeleccionEjercicio.IsHandleCreated)) Then
            frmSeleccionEjercicio = New SeleccionEjercicio
        End If
        ' Llamamos al formulario de manera modal.
        frmSeleccionEjercicio.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmSeleccionEjercicio.Dispose()
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub BtnAcercade_Click(sender As Object, e As EventArgs) Handles BtnAcercade.Click
        AcercaDeToolStripMenuItem.PerformClick()
    End Sub

    Private Sub AcercaDeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AcercaDeToolStripMenuItem.Click
        ' 1. Cambiar el texto de la barra de estado
        TsLabelFormulario.Text = rmse.GetString("AcercaDeToolStripMenuItem.Text")
        ' 2. Controlar la instancia física del formulario de forma tradicional
        If ((frmAcercaDe Is Nothing) OrElse (Not frmAcercaDe.IsHandleCreated)) Then
            frmAcercaDe = New AcercaDe
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmAcercaDe)
        ' 4. Configurar el posicionamiento manual obligatorio
        frmAcercaDe.StartPosition = FormStartPosition.Manual
        ' 5. Calcular las coordenadas definitivas con tu ajuste de +10
        Dim posX As Integer = Me.Left + (Me.Width - frmAcercaDe.Width) \ 2
        Dim posY As Integer = Me.Top + 10
        If posX < 0 Then posX = 0
        ' 6. Fijar la ubicación calculada y abrir la ventana
        frmAcercaDe.Location = New System.Drawing.Point(posX, posY)
        frmAcercaDe.ShowDialog()
        ' 7. Destrucción explícita al cerrar
        frmAcercaDe.Dispose()
        ' 8. IMPORTANTE: Limpiar la variable manual para evitar el error de objeto destruido
        frmAcercaDe = Nothing
        ' 9. Restaurar el texto de espera de la barra
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
    End Sub

    Private Sub BtnPreferencias_Click(sender As Object, e As EventArgs) Handles BtnPreferencias.Click
        PreferenciasToolStripMenuItem.PerformClick()
    End Sub

    Private Sub PreferenciasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PreferenciasToolStripMenuItem.Click
        TsLabelFormulario.Text = rmse.GetString("PreferenciasToolStripMenuItem.Text")
        ' Comprobamos si existe un identificador asociado.
        If ((frmPreferencias Is Nothing) OrElse (Not frmPreferencias.IsHandleCreated)) Then
            frmPreferencias = New Preferencias
        End If
        ' Llamamos al formulario de manera modal.
        frmPreferencias.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmPreferencias.Dispose()
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
        If My.Settings.Codigo = "Codigo Activación: Sin Activar" Then
            Me.LblNotificacion.Visible = True
        Else
            Me.LblNotificacion.Visible = False
        End If
    End Sub

    Private Sub BarraDeHerramientasMenu_Click(sender As Object, e As EventArgs) Handles BarraDeHerramientasMenu.Click
        If BarraDeHerramientasMenu.Checked Then
            My.Settings.BarraHerramientas = True
        Else
            My.Settings.BarraHerramientas = False
        End If
        Cambiarbarraherramientas()
        My.Settings.Save()
    End Sub

    Private Sub BarraDeEstadoMenu_Click(sender As Object, e As EventArgs) Handles BarraDeEstadoMenu.Click
        If BarraDeEstadoMenu.Checked Then
            My.Settings.BarraEstado = True
        Else
            My.Settings.BarraEstado = False
        End If
        Cambiarbarraestado()
        My.Settings.Save()
    End Sub

    Private Sub BarraYMenuConColores_Click(sender As Object, e As EventArgs) Handles BarraYMenuConColores.Click
        If BarraYMenuConColores.Checked Then
            My.Settings.MenuColores = True
        Else
            My.Settings.MenuColores = False
        End If
        CambiarColorBarraMenu()
        My.Settings.Save()
    End Sub

    Private Sub BtnAyuda_Click(sender As Object, e As EventArgs) Handles BtnAyuda.Click
        ArchivoDeAyudaToolStripMenuItem.PerformClick()
    End Sub

    Private Sub ArchivoDeAyudaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ArchivoDeAyudaToolStripMenuItem.Click
        Dim Proceso As New Process
        Proceso.StartInfo.FileName = IO.Path.Combine(carpetaDB, "Ayuda_ContaHogar 3.0.pdf")
        Proceso.StartInfo.Verb = "open"
        Proceso.Start()
    End Sub

    Private Sub BtnIniciarBaseDatos_Click(sender As Object, e As EventArgs) Handles BtnIniciarBaseDatos.Click
        ReiniciarBaseDeDatosToolStripMenuItem.PerformClick()
    End Sub

    Private Sub ReiniciarBaseDeDatosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReiniciarBaseDeDatosToolStripMenuItem.Click
        ' 1. SANEAMIENTO PREVENTIVO: Limpiamos la memoria de consultas previas
        cmdMdb1cr.Parameters.Clear()

        ' 2. PRIMERA ALERTA TRADUCIDA (Con tu confirmador inmune al idioma de Windows)
        Dim msgPregunta1 As String = rmse.GetString("MsgVaciarBD") & ": " & vAñoEjercicio.ToString & "?." & NL & rmse.GetString("MsgVaciarBD2") & NL & rmse.GetString("MsgNoVaciarBD")
        Dim titPregunta1 As String = rmse.GetString("MsgVaciarBD3") & ": " & vAñoEjercicio.ToString

        If ConfirmarAccionTraducida(msgPregunta1, titPregunta1) = MsgBoxResult.No Then
            Exit Sub
        End If

        ' 3. SEGUNDA ALERTA DE SEGURIDAD (Confirmación doble antibloqueos)
        Dim msgPregunta2 As String = rmse.GetString("MsgVaciarBD4") & ": " & vAñoEjercicio.ToString & " ¿Ok?."

        If ConfirmarAccionTraducida(msgPregunta2, titPregunta1) = MsgBoxResult.Yes Then

            ' =========================================================================
            ' 🌟 FASE 1: VACIADO PARAMETRIZADO POR EJERCICIO ANUAL
            ' =========================================================================

            ' A. Eliminar Registro Apuntes Contables (¡Adiós concatenaciones de texto!)
            cmdMdb1cr.CommandText = "DELETE FROM apuntes WHERE EjercicioAPU = ?"
            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.Add("@eje", OleDbType.Integer).Value = Convert.ToInt32(vAñoEjercicio)
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(rmse.GetString("ApuntesContablesVaciado"), MsgBoxStyle.Information)
            Catch ex As Exception
                MsgBox("Error Apuntes: " & ex.Message, MsgBoxStyle.Critical)
            End Try

            ' B. Eliminar Registros Apuntes Periódicos
            cmdMdb1cr.CommandText = "DELETE FROM apuper WHERE EjercicioAPP = ?"
            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.Add("@eje", OleDbType.Integer).Value = Convert.ToInt32(vAñoEjercicio)
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(rmse.GetString("ApuntesPeriodicosVaciado"), MsgBoxStyle.Information)
            Catch ex As Exception
                MsgBox("Error Periódicos: " & ex.Message, MsgBoxStyle.Critical)
            End Try

            ' C. Eliminar Registros Presupuestos
            cmdMdb1cr.CommandText = "DELETE FROM presupuesto WHERE EjercicioPRE = ?"
            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.Add("@eje", OleDbType.Integer).Value = Convert.ToInt32(vAñoEjercicio)
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(rmse.GetString("PresupuestosVaciado"), MsgBoxStyle.Information)
            Catch ex As Exception
                MsgBox("Error Presupuesto: " & ex.Message, MsgBoxStyle.Critical)
            End Try

            '            ' CONCEPTOS Y CUENTAS NO SE ELIMINAN
            '            '***********************************

            ' =========================================================================
            ' FASE 2: TRUNCADO LIMPIO DE TABLAS TEMPORALES DE OPERACIÓN
            ' =========================================================================
            cmdMdb1cr.Parameters.Clear()

            ' Eliminar Registros extracto
            Try
                cmdMdb1cr.CommandText = "DELETE FROM extracto"
                cmdMdb1cr.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End Try

            ' Eliminar Registros Tempapu
            Try
                cmdMdb1cr.CommandText = "DELETE FROM tempapu"
                cmdMdb1cr.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End Try

            ' Eliminar Registros Temppre
            Try
                cmdMdb1cr.CommandText = "DELETE FROM temppre"
                cmdMdb1cr.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End Try

            ' Eliminar Registros Tmpprint
            Try
                cmdMdb1cr.CommandText = "DELETE FROM tmpprint"
                cmdMdb1cr.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End Try

            ' Cierre controlado impecable de fábrica
            MsgBox(resManager.GetString("CerrarApp"), MsgBoxStyle.Information)
            Me.Close()
        End If
    End Sub


    Private Sub BtnCalculadora_Click(sender As Object, e As EventArgs) Handles BtnCalculadora.Click
        CalculadoraToolStripMenuItem.PerformClick()
    End Sub

    Private Sub CalculadoraToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CalculadoraToolStripMenuItem.Click
        Dim Proceso As New Process()
        Proceso.StartInfo.FileName = "calc.exe"
        Proceso.StartInfo.Arguments = ""
        Proceso.Start()
    End Sub

    Private Sub BtnCopiaSeguridad_Click(sender As Object, e As EventArgs) Handles BtnCopiaSeguridad.Click
        HacerCopiaDeSeguridadToolStripMenuItem.PerformClick()
    End Sub

    Private Sub HacerCopiaDeSeguridadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HacerCopiaDeSeguridadToolStripMenuItem.Click
        ' Si no existe la carpeta de BackUp la creamos.
        Dim path As String = "C:\ContaHogar3.0\Backup"
        If Directory.Exists(path) Then
            'MsgBox("Ya existe la Ruta C:\ContaHogar3.0\Backup.")
        Else
            Directory.CreateDirectory(path)
            'MsgBox("Ruta C:\ContaHogar3.0\Backup, Creada.")
        End If
        Dim NombreBaseDatos As String = "ContaHogar3.0" & "[" & Now.ToString("ddMMyyyy") & "]" & "[" & Now.ToString("HHmmss") & "]" & ".mdb"
        Dim DataBaseFile As String = vRuta
        Dim FileDestino As String = "C:\ContaHogar3.0\Backup\" & NombreBaseDatos
        backup.InitialDirectory = "C:\ContaHogar3.0\Backup\"
        backup.Title = "Backup BD Access"
        backup.CheckFileExists = False
        backup.CheckPathExists = False
        backup.DefaultExt = "mdb"
        backup.FileName = NombreBaseDatos
        backup.Filter = "Access (ContaHogar*.mdb)|ContaHogar*.mdb|All files (*.*)|*.*"
        backup.RestoreDirectory = True
        If backup.ShowDialog = Windows.Forms.DialogResult.OK Then
            Try
                FileCopy(DataBaseFile, FileDestino)
                MessageBox.Show(rmse.GetString("BackupOk"), "BACKUP", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub BtnRestaurarCopia_Click(sender As Object, e As EventArgs) Handles BtnRestaurarCopia.Click
        RestaurarCopiaDeSeguridadToolStripMenuItem.PerformClick()
    End Sub

    Private Sub RestaurarCopiaDeSeguridadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestaurarCopiaDeSeguridadToolStripMenuItem.Click
        Dim respuesta As MsgBoxResult = MsgBox(rmse.GetString("PreguntaBackup"), vbQuestion + vbYesNo + vbDefaultButton2, rmse.GetString("RestaurarBD"))
        If respuesta = vbYes Then
            Dim RestoreFile As String = vRuta
            restore.InitialDirectory = "C:\ContaHogar3.0\Backup\"
            restore.Title = rmse.GetString("RestaurarBD")
            restore.CheckFileExists = False
            restore.CheckPathExists = False
            restore.DefaultExt = "mdb"
            restore.Filter = "Access (ContaHogar*.mdb)|ContaHogar*.mdb|All files (*.*)|*.*"
            restore.RestoreDirectory = True
            If restore.ShowDialog = Windows.Forms.DialogResult.OK Then
                Try
                    FileCopy(restore.FileName, RestoreFile)
                    MessageBox.Show(rmse.GetString("RestaurarOk"), rmse.GetString("Restaurar"), MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
                MsgBox(resManager.GetString("CerrarApp"), vbInformation)
                Me.Close()
            End If
        End If
    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        SalirMenu.PerformClick()
    End Sub

    Private Sub SalirMenu_Click(sender As Object, e As EventArgs) Handles SalirMenu.Click
        Me.Close()
        End
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            My.Settings.PantallaAncho = Me.Width
            My.Settings.PantallaAlto = Me.Height
            My.Settings.Posicion = Me.Location.ToString
            CantPantallas = 0
            For Each scrn As Screen In Screen.AllScreens
                CantPantallas += 1
            Next
            My.Settings.Pantallas = CantPantallas
            My.Settings.Save()
            My.Settings.Reload()
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub BtnBizum_MouseHover(sender As Object, e As EventArgs) Handles BtnBizum.MouseHover
        LblDonacion.Visible = True
    End Sub

    Private Sub BtnBizum_MouseLeave(sender As Object, e As EventArgs) Handles BtnBizum.MouseLeave
        LblDonacion.Visible = False
    End Sub

    Private Sub BtnBizum_Click(sender As Object, e As EventArgs) Handles BtnBizum.Click
        System.Diagnostics.Process.Start("https://www.paypal.com/donate/?hosted_button_id=EZCSRQ4QBPVZN")
    End Sub

    Private Sub AportaciónBizumToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AportaciónBizumToolStripMenuItem.Click
        System.Diagnostics.Process.Start("https://www.paypal.com/donate/?hosted_button_id=EZCSRQ4QBPVZN")
    End Sub

    Private Sub BtnHistorialVersiones_Click(sender As Object, e As EventArgs) Handles BtnHistorialVersiones.Click
        HistorialDeVersionesToolStripMenuItem.PerformClick()
    End Sub

    Private Sub HistorialDeVersionesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HistorialDeVersionesToolStripMenuItem.Click
        Dim Proceso As New Process
        Proceso.StartInfo.FileName = IO.Path.Combine(carpetaDB, "Historial_Versiones.pdf")
        Proceso.StartInfo.Verb = "open"
        Proceso.Start()
    End Sub

    Private Sub CambiarColorBarraMenu()
        If My.Settings.MenuColores = True Then
            ArchivoToolStripMenuItem.ForeColor = Color.Green
            VerToolStripMenuItem.ForeColor = Color.Black
            MantenimientosToolStripMenuItem.ForeColor = Color.Magenta
            ApuntesToolStripMenuItem.ForeColor = Color.Blue
            PresupuestosToolStripMenuItem.ForeColor = Color.Red
            InformesToolStripMenuItem.ForeColor = Color.Black
            HerramientasToolStripMenuItem.ForeColor = Color.Brown
            AyudaToolStripMenuItem.ForeColor = Color.Orange
            ToolStripLabel12.BackColor = Color.Green
            ToolStripLabel16.BackColor = Color.Magenta
            ToolStripLabel2.BackColor = Color.Blue
            ToolStripLabel6.BackColor = Color.Red
            ToolStripLabel8.BackColor = Color.Brown
            ToolStripLabel10.BackColor = Color.Orange
            ToolStripLabel4.BackColor = Color.Green
            ToolStripLabel1.BackColor = Color.Magenta
            ToolStripLabel3.BackColor = Color.Blue
            ToolStripLabel7.BackColor = Color.Red
            ToolStripLabel13.BackColor = Color.Brown
            ToolStripLabel5.BackColor = Color.Orange
        Else
            ArchivoToolStripMenuItem.ForeColor = Color.Black
            VerToolStripMenuItem.ForeColor = Color.Black
            MantenimientosToolStripMenuItem.ForeColor = Color.Black
            ApuntesToolStripMenuItem.ForeColor = Color.Black
            PresupuestosToolStripMenuItem.ForeColor = Color.Black
            InformesToolStripMenuItem.ForeColor = Color.Black
            HerramientasToolStripMenuItem.ForeColor = Color.Black
            AyudaToolStripMenuItem.ForeColor = Color.Black
            ToolStripLabel12.BackColor = Color.Transparent
            ToolStripLabel16.BackColor = Color.Transparent
            ToolStripLabel2.BackColor = Color.Transparent
            ToolStripLabel6.BackColor = Color.Transparent
            ToolStripLabel8.BackColor = Color.Transparent
            ToolStripLabel10.BackColor = Color.Transparent
            ToolStripLabel4.BackColor = Color.Transparent
            ToolStripLabel1.BackColor = Color.Transparent
            ToolStripLabel3.BackColor = Color.Transparent
            ToolStripLabel7.BackColor = Color.Transparent
            ToolStripLabel13.BackColor = Color.Transparent
            ToolStripLabel5.BackColor = Color.Transparent
        End If
    End Sub

    Private Sub Cambiarbarraherramientas()
        If My.Settings.BarraHerramientas = True Then
            BarraDeHerramientas.Visible = True
        Else
            BarraDeHerramientas.Visible = False
        End If
    End Sub

    Private Sub Cambiarbarraestado()
        If My.Settings.BarraEstado = True Then
            BarraDeEstado.Visible = True
        Else
            BarraDeEstado.Visible = False
        End If
    End Sub

    Private Sub FrmPrincipal_Move(sender As Object, e As EventArgs) Handles MyBase.Move
        ' Me.Left es la posición X actual en la pantalla
        ' Me.Top es la posición Y actual en la pantalla
        posX = Me.Left
        posY = Me.Top

        ' Ejemplo: Mostrar la posición en la barra de título en tiempo real
        'Me.Text = $"Posición X: {posX} | Y: {posY}"
    End Sub

    'Boton para actualizar resx Manager con los excel

    'Cómo adaptarlo para otros idiomas en el futuroEn el código que tienes en tu botón,
    'solo debes localizar las líneas donde pone "G" y cambiarlas por la letra de la columna
    'del idioma que quieras actualizar.Por ejemplo, si el Inglés está en la columna H de tus archivos de Excel,
    'modificarías estas 3 líneas estratégicas en tu código:En la lectura de 'Casi' (Fila 27 aprox.):vb'
    'Cambias "G" por "H" (o la letra del inglés)
    'Dim valorG As String = wsCasi.Cell(fila, "H").GetString() 
    'En la comparación de 'Todo' (Fila 49 aprox.):vb' Cambias "G" por "H"
    'Dim celdaGTodo As IXLCell = wsTodo.Cell(fila, "H") 
    'En el mensaje de éxito (Fila 62 aprox.):vbMessageBox.Show($"... en la columna H.", ...)
    'De esta manera, el programa hará exactamente el mismo trabajo quirúrgico con el resto de idiomas,
    'fila a fila, protegiendo las Keys duplicadas de los diferentes formularios y ahorrándote volver a
    'pasar por todo este proceso manual.

    Private Sub ButtonActualizar_Click(sender As Object, e As EventArgs) Handles ButtonActualizar.Click
        ' 1. Configura tus rutas reales
        Dim rutaTodo As String = "C:\Dell\Todo.xlsx"
        Dim rutaCasi As String = "C:\Dell\Casi.xlsx"

        Me.Cursor = Cursors.WaitCursor

        Try
            ' 2. Abrir archivos en memoria
            Using wbTodo As New XLWorkbook(rutaTodo)
                Using wbCasi As New XLWorkbook(rutaCasi)

                    Dim wsTodo As IXLWorksheet = wbTodo.Worksheet(1)
                    Dim wsCasi As IXLWorksheet = wbCasi.Worksheet(1)

                    ' 3. Cargamos el archivo 'Casi' en el diccionario usando Clave Compuesta (B + C)
                    ' Clave: "File|Key" (Ej: "\ActivarSoftware|MyKey"), Valor: Columna G (Catalán)
                    Dim diccionarioCasi As New Dictionary(Of String, String)()
                    Dim ultimaFilaCasi As Integer = wsCasi.LastRowUsed().RowNumber()

                    For fila As Integer = 2 To ultimaFilaCasi
                        Dim fileB As String = wsCasi.Cell(fila, "B").GetString().Trim().ToLower()
                        Dim keyC As String = wsCasi.Cell(fila, "C").GetString().Trim().ToLower()
                        Dim valorG As String = wsCasi.Cell(fila, "G").GetString()

                        ' Creamos el identificador único compuesto
                        Dim claveCompuesta As String = $"{fileB}|{keyC}"

                        If Not String.IsNullOrEmpty(fileB) AndAlso Not String.IsNullOrEmpty(keyC) Then
                            ' Al usar la combinación B+C no habrá duplicados, lo guardamos/actualizamos
                            diccionarioCasi(claveCompuesta) = valorG
                        End If
                    Next

                    ' 4. Recorrer 'Todo' y machacar diferencias comparando la Clave Compuesta
                    Dim ultimaFilaTodo As Integer = wsTodo.LastRowUsed().RowNumber()
                    Dim filasModificadas As Integer = 0

                    For fila As Integer = 2 To ultimaFilaTodo
                        Dim fileBTodo As String = wsTodo.Cell(fila, "B").GetString().Trim().ToLower()
                        Dim keyCTodo As String = wsTodo.Cell(fila, "C").GetString().Trim().ToLower()

                        ' Generamos la misma clave compuesta para buscar
                        Dim claveCompuestaTodo As String = $"{fileBTodo}|{keyCTodo}"

                        ' Si la combinación exacta de Formulario + Key existe en 'Casi'
                        If diccionarioCasi.ContainsKey(claveCompuestaTodo) Then
                            Dim valorG_Casi As String = diccionarioCasi(claveCompuestaTodo)
                            Dim celdaGTodo As IXLCell = wsTodo.Cell(fila, "G")

                            ' Si el catalán actual no coincide con el de 'Casi', se machaca
                            If celdaGTodo.GetString() <> valorG_Casi Then
                                celdaGTodo.SetValue(valorG_Casi)
                                filasModificadas += 1
                            End If
                        End If
                    Next

                    ' 5. Guardar los cambios directamente
                    If filasModificadas > 0 Then
                        wbTodo.Save()
                        MessageBox.Show($"¡Proceso completado con éxito!{Environment.NewLine}Se han machacado {filasModificadas} líneas correctamente usando la combinación de Formulario y Key.",
                                    "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("El proceso terminó, pero no se encontraron diferencias para modificar en la columna G.",
                                    "Sin cambios", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If

                End Using
            End Using

        Catch ex As System.IO.IOException
            MessageBox.Show($"Error de acceso: Asegúrate de cerrar los archivos de Excel.{Environment.NewLine}{Environment.NewLine}Detalle: {ex.Message}",
                        "Archivo Bloqueado", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Ocurrió un error inesperado: {ex.Message}",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

End Class