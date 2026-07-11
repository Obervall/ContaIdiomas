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
        ' 1. Refrescar Labels, Botones y Título del Formulario Principal (Tu excelente bucle original)
        For Each ctrl As Control In Me.Controls
            Dim textoTraducido As String = My.Resources.ResourceManager.GetString(ctrl.Name)
            If Not String.IsNullOrEmpty(textoTraducido) Then
                ctrl.Text = textoTraducido
            End If
            If textoTraducido Is Nothing Then
                textoTraducido = rmse.GetString(ctrl.Name)
            End If
        Next

        ' =========================================================================
        ' 🌟 EL PUENTE RELACIONAL MULTIIDIOMA (Sincro en caliente para el extracto)
        ' =========================================================================
        ' Si el usuario cambió el idioma en Preferencias, obligamos a que el formulario 
        ' de Apuntes Contables (si está abierto en la RAM) pase de inmediato el traductor 
        ' por sus celdas para que el 'Saldo Inicial' mutre al alemán/catalán al instante.
        ' (Asegúrate de comprobar si tu clase de formulario se llama exactamente frmApuntesContables)
        For Each f As Form In Application.OpenForms
            If f.Name = "frmApuntesContables" OrElse f.Name = "ApuntesContables" Then
                ' Localizamos la rejilla viva (asumiendo que se llama DgvApuntes)
                Dim dgvContable As DataGridView = f.Controls.Find("DgvApuntes", True).FirstOrDefault()
                If dgvContable IsNot Nothing Then
                    ' Ejecutamos la rutina de tu módulo con la cultura en caliente ya actualizada
                    TraducirGridApuntesBD(dgvContable)
                End If
            End If
        Next

        ' 2. Llamar al refresco de menús de fábrica
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
        ActualizarTextosFormulario(Me)

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

        '********************************================================****************************************
        ' 🚀 ARRANQUE INTELIGENTE MODO MSIX CON PUENTE DE RESCATE (Sustitución de vRuta)
        '****************================================================================================********
        ' 1. Definimos la NUEVA RUTA oficial, libre de derechos, dentro de "Mis Documentos" (Recomendado para MSIX)
        Dim carpetaDocumentos As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim carpetaAppOficial As String = IO.Path.Combine(carpetaDocumentos, "ContaHogar3.0")
        Dim archivoBdDestino As String = IO.Path.Combine(carpetaAppOficial, "ContaHogar.mdb")

        ' 2. Capturamos milimétricamente tu RUTA ANTERIOR de AppData\Roaming para rescatar datos de usuarios viejos
        Dim appDataPathViejo As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Dim carpetaDBVieja As String = IO.Path.Combine(appDataPathViejo, "A.Oberholzer", "ContaHogar3.0")
        Dim archivoBdAppDataVieja As String = IO.Path.Combine(carpetaDBVieja, "ContaHogar.mdb")

        ' 3. Creamos la nueva carpeta física en Mis Documentos si es la primera vez que entra
        Try
            If Not Directory.Exists(carpetaAppOficial) Then
                Directory.CreateDirectory(carpetaAppOficial)
            End If
        Catch ex As Exception
            ' Cortafuegos de emergencia por si Mis Documentos estuviera restringido por OneDrive
            carpetaAppOficial = carpetaDocumentos
            archivoBdDestino = IO.Path.Combine(carpetaAppOficial, "ContaHogar.mdb")
        End Try

        ' 4. 🎯 EL PUENTE DE MIGRACIÓN: Sembramos el archivo respetando el histórico del usuario
        If Not File.Exists(archivoBdDestino) Then
            Try
                ' ESCENARIO B: ¿El usuario es antiguo y tiene sus apuntes reales en tu vieja ruta de Roaming?
                If File.Exists(archivoBdAppDataVieja) Then
                    ' Mudamos su base de datos real intacta a Mis Documentos para no perder ni un céntimo
                    File.Copy(archivoBdAppDataVieja, archivoBdDestino, True)

                    ' ESCENARIO A: Es un usuario 100% nuevo que se descarga la app de la Microsoft Store
                Else
                    ' Extraemos la base limpia y compactada de 488 KB que viaja dentro del paquete MSIX
                    Dim archivoBdOrigenRuta As String = IO.Path.Combine(Application.StartupPath, "ContaHogar.mdb")
                    If File.Exists(archivoBdOrigenRuta) Then
                        File.Copy(archivoBdOrigenRuta, archivoBdDestino, True)
                    End If
                End If
            Catch ex As Exception
                MsgBox(rmse.GetString("ErrorCriticoPuenteRescate") & ": " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If

        ' 5. 🌟 ASENTAMOS LA VARIABLE GLOBAL Y SCONCRONIZAMOS LAS PREFERENCIAS
        vRuta = archivoBdDestino

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

        'Hoy (Versión 3.1.1): Dejas tu código con la "3.1.1".
        'Y cuando configuremos el manifiesto del proyecto ContaHogarPackaging,
        'nos aseguraremos de escribir en su casilla de versión: 3.1.1.0 [1.1].
        'Mañana (Versión 3.1.2): Cambiarás tu línea a My.Settings.Version = "3.1.2",
        'cambiarás el manifiesto del MSIX a 3.1.2.0 y subirás el nuevo paquete reluciente [1.1].

        My.Settings.Version = "3.1.1.0"
        My.Settings.Save()

        vMoneda = My.Settings.Moneda

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

        'Iniciar los Saldos Iniciales del Ejercicio
        IniciarSaldosIniciales(vAñoEjercicio)

        ' Mostramos el mensaje de "En Espera..." en la barra de estado mientras se carga todo
        If vAviso = True Then
            Me.TsLabelFormulario.Text = resManager.GetString("Aviso") & ": " & resManager.GetString("NoHayDatosHistoricos")
        Else
            Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
        End If

        ' =========================================================================
        ' 🌟 PROCESADOR AUTOMÁTICO DE TRANSACCIONES PERIÓDICAS (Nueva Era)
        ' =========================================================================
        cmdMdb1cr.Parameters.Clear()

        ' 1. Aseguramos que la fecha de hoy sea un objeto Date cronológico puro sin horas
        Dim dHoy As Date = DateTime.Today
        If vfechaHoy.Year > 1 Then
            dHoy = vfechaHoy.Date
        End If

        ' 2. Creamos una lista volátil en la RAM para recolectar los registros a procesar
        ' Esto evita mantener el DataReader abierto mientras hacemos INSERT/DELETE en la misma tabla
        Dim listaAsientosAProcesar As New List(Of Dictionary(Of String, Object))()

        cmdMdb1cr.CommandText = "SELECT apuper.CodigoAPP, apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP FROM apuper ORDER BY apuper.FechaAPP ASC"

        Try
            Using drMdb1 As OleDbDataReader = cmdMdb1cr.ExecuteReader()
                While drMdb1.Read()
                    Dim fechaAsiento As Date = Convert.ToDateTime(drMdb1("FechaAPP"))

                    ' 🚀 FILTRO CRONOLÓGICO SEGURO: Si la fecha del apunte periódico es hoy o anterior
                    If fechaAsiento.Date <= dHoy.Date Then
                        Dim registro As New Dictionary(Of String, Object)()
                        registro("CodigoAPP") = drMdb1("CodigoAPP")
                        registro("FechaAPP") = fechaAsiento.Date
                        registro("ConceptoAPP") = drMdb1("ConceptoAPP") ' ID Numérico Entero largo
                        registro("DescripcionAPP") = drMdb1("DescripcionAPP").ToString()
                        registro("ImporteAPP") = drMdb1("ImporteAPP").ToString()
                        registro("NotasAPP") = drMdb1("NotasAPP").ToString()
                        registro("CuentaAPP") = drMdb1("CuentaAPP").ToString()

                        listaAsientosAProcesar.Add(registro)
                    End If
                End While
            End Using
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorEscanearPeriodicos") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try

        ' =========================================================================
        ' 3. BUCLE MÁGICO DE VOLCADO QUIRÚRGICO A DIARIO (Inmune a fallos de tipos)
        ' =========================================================================
        For Each asu In listaAsientosAProcesar
            Dim vCodigo As Integer = Convert.ToInt32(asu("CodigoAPP"))
            Dim vDate1 As Date = CDate(asu("FechaAPP"))
            Dim idConcepto As Integer = Convert.ToInt32(asu("ConceptoAPP"))
            Dim vDescripcion As String = ApostrofePorAcentoAgudo(asu("DescripcionAPP").ToString())
            Dim vImporte As String = asu("ImporteAPP").ToString()
            Dim vNotas As String = asu("NotasAPP").ToString()
            Dim vCuenta As String = asu("CuentaAPP").ToString()

            ' A. INYECCIÓN PARAMETRIZADA PURA EN LA TABLA DE APUNTES DIARIOS
            vAñadirSql = "INSERT INTO apuntes (FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) VALUES (?, ?, ?, ?, ?, ?, ?)"
            cmdMdb1cr.CommandText = vAñadirSql
            cmdMdb1cr.Parameters.Clear()

            cmdMdb1cr.Parameters.AddWithValue("@FechaAPU", vDate1.Date)
            cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", idConcepto) ' 🚀 ¡EXITO!: Inyectamos el ID entero relacional
            cmdMdb1cr.Parameters.AddWithValue("@DescripcionAPU", vDescripcion.Trim())

            Dim paramImp As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@ImporteAPU", OleDb.OleDbType.Currency)
            paramImp.Value = Math.Round(ConvertirDecimalSeguro(vImporte), 2)

            cmdMdb1cr.Parameters.AddWithValue("@EjercicioAPU", CInt(vAñoEjercicio))
            cmdMdb1cr.Parameters.AddWithValue("@NotasAPU", vNotas.Trim())
            cmdMdb1cr.Parameters.AddWithValue("@CuentaAPU", vCuenta.Trim())

            Try
                cmdMdb1cr.ExecuteNonQuery()

                ' Averiguamos el nombre corto legible del concepto para enseñarlo en el cartel traducido
                Dim nombreCortoConcepto As String = "CONCEPTO"
                Using con As New OleDbConnection(conexion1.ConnectionString)
                    Using cmd As New OleDbCommand("SELECT CodigoCON FROM conceptos WHERE IdConceptoCON = ?", con)
                        cmd.Parameters.Add("@id", OleDbType.Integer).Value = idConcepto
                        Try
                            con.Open()
                            Dim r = cmd.ExecuteScalar()
                            If r IsNot Nothing Then nombreCortoConcepto = r.ToString().Replace("_", " ").ToUpper()
                        Catch
                        End Try
                    End Using
                End Using

                ' Mostramos tu cartel oficial con el idioma regional activo
                MsgBox(vDate1.ToShortDateString() & vbNewLine & nombreCortoConcepto & "     " & vDescripcion & "     " & vImporte & vbNewLine & rmse.GetString("CreadoCorrectamente"), MsgBoxStyle.Information)

                ' B. EXTIRPACIÓN DEL VENCIMIENTO YA PROCESADO EN LA TABLA APUPER
                cmdMdb1cr.CommandText = "DELETE FROM apuper WHERE CodigoAPP = ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.Add("@cod", OleDbType.Integer).Value = vCodigo
                cmdMdb1cr.ExecuteNonQuery()

            Catch ex As Exception
                MsgBox(resManager.GetString("ErrorApuntePeriodico") & ": " & ex.Message, MsgBoxStyle.Critical)
            End Try
        Next
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmIntroApuntes)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmApuntesPeriodicos)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionFechas)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionarDatosIngresos)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionarDatosGastos)
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
        ' =========================================================================
        ' 🚀 REPARADO MODO MSIX: ELIMINADAS LAS RUTAS RÍGIDAS DE C:\
        ' =========================================================================
        ' Creamos el archivo temporal de compactación en la misma carpeta segura de Mis Documentos
        Dim carpetaDestinoSafe As String = IO.Path.GetDirectoryName(vRuta)
        Dim vCompactadaReal As String = IO.Path.Combine(carpetaDestinoSafe, "contahogarcompacted.mdb")

        ' Inicializamos el motor nativo de compactación de Microsoft Jet
        Dim jetEng As JRO.JetEngine
        jetEng = New JRO.JetEngine()

        ' Limpiamos cualquier rastro temporal anterior de forma dócil en la RAM
        If File.Exists(vCompactadaReal) Then
            Try : File.Delete(vCompactadaReal) : Catch : End Try
        End If

        Try
            ' 1. Cerramos la compuerta de la conexión principal para liberar el candado del archivo
            conexion1.Close()

            ' 2. 🚀 LA JUGADA MAESTRA: Compactamos en caliente dentro de la zona libre de derechos de usuario
            jetEng.CompactDatabase("Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & vRuta,
                                   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vCompactadaReal & ";Jet OLEDB:Engine Type=5")

            ' 3. Reemplazamos la base de datos oficial por la versión pulida y reducida a 488 KB
            File.Copy(vCompactadaReal, vRuta, True)

            ' 4. Volvemos a levantar el motor de datos relacional para las cuadrículas y rejillas
            tipoDsn = "AccessMdb"
            Conectarse(tipoDsn)

            MessageBox.Show(rmse.GetString("CompactacionOk"), rmse.GetString("Compactar"), MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            ' Cortafuegos de emergencia: si algo falla, intentamos reconectar para no dejar la pantalla colgada
            Try : Conectarse("AccessMdb") : Catch : End Try
            MsgBox(rmse.GetString("CompactarError") & ":  " & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Error"))
        Finally
            ' Limpieza biológica final: trituramos el temporal contahogarcompacted del perfil
            If File.Exists(vCompactadaReal) Then
                Try : File.Delete(vCompactadaReal) : Catch : End Try
            End If
        End Try
    End Sub

    Private Sub BtnImportarContaHogar_Click(sender As Object, e As EventArgs) Handles BtnImportarContaHogar.Click
        ImportaAntiguoContahogarToolStripMenuItem.PerformClick()
    End Sub

    Private Sub ImportaAntiguoContahogarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportaAntiguoContahogarToolStripMenuItem.Click
        TsLabelFormulario.Text = rmse.GetString("BtnImportarContaHogar.Text")

        Dim respuesta As MsgBoxResult = ConfirmarAccionTraducida(rmse.GetString("MsgImportar1"), rmse.GetString("ImportarContahogar"))
        If respuesta = vbYes Then
            ' =========================================================================
            ' 🚀 COMPRESIÓN COMPLETA: BACKUP AUTOMÁTICO A SACO (¡Inmune a Colisiones!)
            ' =========================================================================
            ' 1. Calculamos la ruta de la carpeta segura usando IO.Path.Combine nativo
            Dim carpetaBackupSegura As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ContaHogar_Backups")
            Try
                If Not Directory.Exists(carpetaBackupSegura) Then Directory.CreateDirectory(carpetaBackupSegura)
            Catch
                carpetaBackupSegura = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            End Try
            ' 2. Fabricamos el nombre cronológico simétrico y calculamos la ruta final
            Dim NombreBaseDatos As String = "ContaHogar3.0_PRE_MIGRACION_[" & Now.ToString("ddMMyyyy") & "]_[" & Now.ToString("HHmmss") & "].mdb"
            Dim DataBaseFile As String = vRuta
            Dim FileDestinoReal As String = IO.Path.Combine(carpetaBackupSegura, NombreBaseDatos)
            ' 3. 🚀 EJECUCIÓN DIRECTA EN CALIENTE: Copiamos la base de datos sin abrir ventanas
            Try
                If File.Exists(DataBaseFile) Then
                    FileCopy(DataBaseFile, FileDestinoReal)
                    ' Opcional: Puedes quitar este MessageBox si quieres que el backup sea 100% invisible
                    MessageBox.Show(rmse.GetString("MsgImportarCopiaPreventiva") & ": " & NombreBaseDatos, "BACKUP", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Catch ex As Exception
                ' Cortafuegos silencioso para que un fallo en el backup nunca aborte la importación real
            End Try

            ' =========================================================================
            ' 🌟 FASE A RECTIFICADA: BUSCADOR INTELIGENTE COMPATIBLE CON MSIX
            ' =========================================================================
            ' 1. Definimos la ruta del clon temporal estrictamente en "Mis Documentos" (Donde hay permiso MSIX)
            Dim carpetaDocumentos As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Dim carpetaAppOficial As String = IO.Path.Combine(carpetaDocumentos, "ContaHogar3.0")
            Dim RutaClonMigrada As String = IO.Path.Combine(carpetaAppOficial, "CHDB2_MIGRADA.mdb")
            Dim RutaOriginalVieja As String = ""

            ' 2. Configuramos el buscador oficial de archivos de Windows
            Using ofd As New OpenFileDialog()
                ofd.Title = "Selecciona la base de datos antigua (CHDB2.mdb) para importar"
                ofd.Filter = "Base de datos ContaHogar 2.0 (CHDB2.mdb)|CHDB2.mdb|Todos los archivos (*.*)|*.*"
                ' Sugerimos por defecto la ruta típica donde moría el programa viejo para ahorrarle trabajo al usuario
                Dim rutaPorDefectoVieja As String = "C:\Program Files (x86)\ContaHogar"
                If Directory.Exists(rutaPorDefectoVieja) Then
                    ofd.InitialDirectory = rutaPorDefectoVieja
                Else
                    ofd.InitialDirectory = carpetaDocumentos
                End If

                ' 3. 🎯 LA JUGADA MAESTRA: Si el archivo existe en la ruta típica, lo pre-seleccionamos automáticamente
                Dim rutaFisicaTipica As String = IO.Path.Combine(rutaPorDefectoVieja, "CHDB2.mdb")
                If File.Exists(rutaFisicaTipica) Then
                    RutaOriginalVieja = rutaFisicaTipica
                Else
                    ' Si no está en su sitio de nacimiento, le abrimos la ventana dócilmente para que lo busque a mano
                    If ofd.ShowDialog() = DialogResult.OK Then
                        RutaOriginalVieja = ofd.FileName
                    Else
                        ' Si el usuario cancela la búsqueda, apagamos la redonda y abortamos de forma limpia
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If
                End If
            End Using

            ' 4. Duplicamos el archivo seleccionado hacia el clon seguro de laboratório en Mis Documentos
            If File.Exists(RutaClonMigrada) Then File.Delete(RutaClonMigrada)
                File.Copy(RutaOriginalVieja, RutaClonMigrada)

            ' =========================================================================
            ' 🌟 FLECO: ENCENDEMOS LA REDONDITA GIRATORIA (UX Premium)
            ' =========================================================================
            Me.Cursor = Cursors.WaitCursor

            ' =========================================================================
            ' 🌟 FLECO EL CORTAFUEGOS DEL PASO A (Evita dobles migraciones)
            ' =========================================================================
            ' Creamos el clon temporal en el disco duro
            If File.Exists(RutaClonMigrada) Then File.Delete(RutaClonMigrada)
                File.Copy(RutaOriginalVieja, RutaClonMigrada)
                ' Tu Paso 1 biológico interroga si la columna sigue siendo Texto
                Dim necesitaActualizar As Boolean = False
                Dim stringConexionClon As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & RutaClonMigrada & ";"
                Using conexionClon As New OleDbConnection(stringConexionClon)
                    Using cmdClonVerificar As New OleDbCommand("SELECT TOP 1 ConceptoAPU FROM apuntes", conexionClon)
                        Try
                            conexionClon.Open()
                            Using adapter As New OleDbDataAdapter(cmdClonVerificar)
                                Dim dtPrueba As New DataTable()
                                adapter.Fill(dtPrueba)
                                If dtPrueba.Columns("ConceptoAPU").DataType = GetType(String) Then necesitaActualizar = True
                            End Using
                        Catch
                            necesitaActualizar = False
                        End Try
                    End Using
                End Using
                ' Si la base ya tiene IDs numéricos, limpia el clon y frena el Sub en seco
                If Not necesitaActualizar Then
                    If File.Exists(RutaClonMigrada) Then File.Delete(RutaClonMigrada)
                    MsgBox(rmse.GetString("MsgExisteClonMigracion"), MsgBoxStyle.Information, rmse.GetString("ImportarContahogar"))
                    Exit Sub
                End If
                ' Lanzamos tu rutina específica de alteración estructural aislada
                MigrarEstructuraBaseDatosExterna(RutaClonMigrada)
                ' =========================================================================
                ' 🚀 FASE B: MOTOR DE VOLCADO INTELIGENTE POR TEXTO (Coherencia Total)
                ' =========================================================================
                Dim connClonString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & RutaClonMigrada & ";"
                Dim connDestinoString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vRuta & ";"
                ' Abrimos las dos compuertas principales una sola vez para todas las tablas
                Using connClon As New OleDbConnection(connClonString)
                    Using connDestino As New OleDbConnection(connDestinoString)
                        Try
                            connClon.Open()
                            connDestino.Open()
                            ' =========================================================================
                            ' 📌 TRAMO B.1: VOLCADO DE APUNTES DIARIOS CON AUTO-CREACIÓN EN CALIENTE
                            ' =========================================================================
                            ' Leemos los campos de texto originales directos de la base vieja
                            Dim sqlSelectClon As String = "SELECT FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU FROM APUNTES"
                            Using cmdClon As New OleDbCommand(sqlSelectClon, connClon)
                                Using reader As OleDbDataReader = cmdClon.ExecuteReader()
                                    ' Verificación limpia en tu base de datos destino (IDs numéricos nativos)
                                    Dim sqlCheck As String = "SELECT COUNT(*) FROM APUNTES WHERE FechaAPU = ? AND ConceptoAPU = ? AND DescripcionAPU = ? AND ImporteAPU = ? AND EjercicioAPU = ? AND NotasAPU = ? AND CuentaAPU = ?"
                                    Using cmdCheck As New OleDbCommand(sqlCheck, connDestino)
                                        cmdCheck.Parameters.Clear()
                                        cmdCheck.Parameters.Add("?", OleDbType.Date)
                                        cmdCheck.Parameters.Add("?", OleDbType.Integer)
                                        cmdCheck.Parameters.Add("?", OleDbType.VarChar)
                                        cmdCheck.Parameters.Add("?", OleDbType.Currency)
                                        cmdCheck.Parameters.Add("?", OleDbType.Integer)
                                        cmdCheck.Parameters.Add("?", OleDbType.VarChar)
                                        cmdCheck.Parameters.Add("?", OleDbType.Integer)
                                        ' Inserción relacional pura y directa en la tabla de destino real
                                        Dim sqlInsert As String = "INSERT INTO APUNTES (FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) VALUES (?, ?, ?, ?, ?, ?, ?)"
                                        Using cmdInsert As New OleDbCommand(sqlInsert, connDestino)
                                            cmdInsert.Parameters.Clear()
                                            cmdInsert.Parameters.Add("?", OleDbType.Date)
                                            cmdInsert.Parameters.Add("?", OleDbType.Integer)
                                            cmdInsert.Parameters.Add("?", OleDbType.VarChar)
                                            cmdInsert.Parameters.Add("?", OleDbType.Currency)
                                            cmdInsert.Parameters.Add("?", OleDbType.Integer)
                                            cmdInsert.Parameters.Add("?", OleDbType.VarChar)
                                            cmdInsert.Parameters.Add("?", OleDbType.Integer)
                                            ' TRADUCTORES QUIRÚRGICOS: Buscan el ID real en tu base buena usando la palabra exacta en mayúsculas
                                            Dim cmdIdCON As New OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = ? OR DescripcionCON = ?", connDestino)
                                            cmdIdCON.Parameters.Add("?", OleDbType.VarChar)
                                            cmdIdCON.Parameters.Add("?", OleDbType.VarChar)
                                            Dim cmdIdCUE As New OleDbCommand("SELECT IdCuentaCUE FROM cuentas WHERE NombreCUE = ?", connDestino)
                                            cmdIdCUE.Parameters.Add("?", OleDbType.VarChar)
                                            Dim contador As Integer = 0
                                            While reader.Read()
                                                Dim vFecha As Object = If(reader.IsDBNull(0), DBNull.Value, reader.GetValue(0))
                                                Dim txtConceptoViejo As String = If(reader.IsDBNull(1), "VARIOS", reader.GetValue(1).ToString().Trim().ToUpper())
                                                Dim vDescripcion As Object = If(reader.IsDBNull(2), DBNull.Value, reader.GetValue(2))
                                                Dim vImporte As Object = If(reader.IsDBNull(3), DBNull.Value, reader.GetValue(3))
                                                Dim vEjercicio As Object = If(reader.IsDBNull(4), DBNull.Value, Convert.ToInt32(reader.GetValue(4)))
                                                Dim vNotas As Object = If(reader.IsDBNull(5), DBNull.Value, reader.GetValue(5))
                                                Dim txtCuentaVieja As String = If(reader.IsDBNull(6), "VARIOS", reader.GetValue(6).ToString().Trim().ToUpper())
                                                ' --- RESOLUCIÓN INTELIGENTE DE CONCEPTOS ---
                                                Dim idConceptoReal As Integer = 1
                                                cmdIdCON.Parameters(0).Value = txtConceptoViejo
                                                cmdIdCON.Parameters(1).Value = txtConceptoViejo
                                                Dim resC = cmdIdCON.ExecuteScalar()

                                            If resC IsNot Nothing AndAlso Not IsDBNull(resC) Then
                                                idConceptoReal = Convert.ToInt32(resC)
                                            Else
                                                ' Si NO existe, lo creamos en tu base buena al vuelo interrogando al pasado
                                                Try
                                                    Dim cmdMaxCon As New OleDbCommand("SELECT MAX(IdConceptoCON) FROM conceptos", connDestino)
                                                    Dim maxC = cmdMaxCon.ExecuteScalar()
                                                    Dim nuevoIdCON As Integer = If(maxC IsNot Nothing AndAlso Not IsDBNull(maxC), Convert.ToInt32(maxC) + 1, 2)

                                                    ' 🚀 LA JUGADA MAESTRA: Viajamos al clon a buscar el tipo original real de ese concepto
                                                    Dim tipoOriginalViejo As String = "GASTO" ' Salvavidas base
                                                    Dim cmdBuscaTipo As New OleDbCommand("SELECT TipoCON FROM conceptos WHERE CodigoCON = ?", connClon)
                                                    cmdBuscaTipo.Parameters.AddWithValue("?", txtConceptoViejo)
                                                    Dim resTipo = cmdBuscaTipo.ExecuteScalar()

                                                    ' Si lo encuentra en el pasado, capturamos su tipo biológico (GASTO, INGRESO)
                                                    If resTipo IsNot Nothing AndAlso Not IsDBNull(resTipo) Then
                                                        tipoOriginalViejo = resTipo.ToString().Trim().ToUpper()
                                                    End If

                                                    ' Inserción simétrica perfecta en tu base de datos buena
                                                    Dim cmdInsCon As New OleDbCommand("INSERT INTO conceptos (IdConceptoCON, CodigoCON, DescripcionCON, TipoCON) VALUES (?, ?, ?, ?)", connDestino)
                                                    cmdInsCon.Parameters.AddWithValue("?", nuevoIdCON)
                                                    cmdInsCon.Parameters.AddWithValue("?", txtConceptoViejo)
                                                    cmdInsCon.Parameters.AddWithValue("?", txtConceptoViejo)
                                                    cmdInsCon.Parameters.AddWithValue("?", tipoOriginalViejo) ' 🎯 Inyectamos el tipo real rescatado
                                                    cmdInsCon.ExecuteNonQuery()

                                                    idConceptoReal = nuevoIdCON
                                                Catch
                                                    idConceptoReal = 1
                                                End Try
                                            End If

                                            ' --- RESOLUCIÓN INTELIGENTE DE CUENTAS BANCARIAS ---
                                            Dim idCuentaReal As Integer = 1
                                                cmdIdCUE.Parameters(0).Value = txtCuentaVieja
                                                Dim resQ = cmdIdCUE.ExecuteScalar()

                                                If resQ IsNot Nothing AndAlso Not IsDBNull(resQ) Then
                                                    idCuentaReal = Convert.ToInt32(resQ)
                                                Else
                                                    ' Si el usuario creó una cuenta vieja que no está en la nueva, se da de alta sola (Tipo 1 por defecto)
                                                    Try
                                                        Dim cmdMaxCue As New OleDbCommand("SELECT MAX(IdCuentaCUE) FROM cuentas", connDestino)
                                                        Dim maxQ = cmdMaxCue.ExecuteScalar()
                                                        Dim nuevoIdCUE As Integer = If(maxQ IsNot Nothing AndAlso Not IsDBNull(maxQ), Convert.ToInt32(maxQ) + 1, 2)

                                                        Dim cmdInsCue As New OleDbCommand("INSERT INTO cuentas (IdCuentaCUE, NombreCUE, NumeroCUE, TipoCUE, NotasCUE) VALUES (?, ?, 'MIGRADA', 1, 'Cuenta importada automáticamente')", connDestino)
                                                        cmdInsCue.Parameters.AddWithValue("?", nuevoIdCUE)
                                                        cmdInsCue.Parameters.AddWithValue("?", txtCuentaVieja)
                                                        cmdInsCue.ExecuteNonQuery()
                                                        idCuentaReal = nuevoIdCUE
                                                    Catch
                                                        idCuentaReal = 1
                                                    End Try
                                                End If

                                                ' Sincronizamos los parámetros del verificador en destino
                                                cmdCheck.Parameters(0).Value = vFecha
                                                cmdCheck.Parameters(1).Value = idConceptoReal
                                                cmdCheck.Parameters(2).Value = vDescripcion
                                                cmdCheck.Parameters(3).Value = vImporte
                                                cmdCheck.Parameters(4).Value = vEjercicio
                                                cmdCheck.Parameters(5).Value = vNotas
                                                cmdCheck.Parameters(6).Value = idCuentaReal
                                                If Convert.ToInt32(cmdCheck.ExecuteScalar()) = 0 Then

                                                    cmdInsert.Parameters(0).Value = vFecha
                                                    cmdInsert.Parameters(1).Value = idConceptoReal
                                                    cmdInsert.Parameters(2).Value = vDescripcion
                                                    cmdInsert.Parameters(3).Value = If(IsNumeric(vImporte), Math.Round(Convert.ToDouble(vImporte), 2), vImporte)
                                                    cmdInsert.Parameters(4).Value = vEjercicio
                                                    cmdInsert.Parameters(5).Value = vNotas
                                                    cmdInsert.Parameters(6).Value = idCuentaReal
                                                    cmdInsert.ExecuteNonQuery()
                                                    contador += 1
                                                End If
                                            End While
                                            MsgBox(rmse.GetString("TransferenciaApuntes") & ". " & contador.ToString() & " " & rmse.GetString("RegistrosCopiados"), MsgBoxStyle.Information, rmse.GetString("$this.Text"))
                                        End Using
                                    End Using
                                End Using
                            End Using
                            ' =========================================================================
                            ' 🚀 FASE B.1.5: VOLCADO SIMÉTRICO DE AÑOS (EJERCICIOS)
                            ' =========================================================================
                            Dim sqlSelectEje As String = "SELECT EjercicioEJE FROM ejercicios"
                            Using cmdClonEje As New OleDbCommand(sqlSelectEje, connClon)
                                Using readerEJE As OleDbDataReader = cmdClonEje.ExecuteReader()

                                    Dim sqlCheckEje As String = "SELECT COUNT(*) FROM ejercicios WHERE EjercicioEJE = ?"
                                    Using cmdCheckEje As New OleDbCommand(sqlCheckEje, connDestino)
                                        cmdCheckEje.Parameters.Clear()
                                        cmdCheckEje.Parameters.Add("?", OleDbType.Integer)

                                        Dim sqlInsertEje As String = "INSERT INTO ejercicios (EjercicioEJE) VALUES (?)"
                                        Using cmdInsertEje As New OleDbCommand(sqlInsertEje, connDestino)
                                            cmdInsertEje.Parameters.Clear()
                                            cmdInsertEje.Parameters.Add("?", OleDbType.Integer)

                                            Dim contEje As Integer = 0

                                            While readerEJE.Read()
                                                Dim vAnio As Object = If(readerEJE.IsDBNull(0), Date.Today.Year, Convert.ToInt32(readerEJE.GetValue(0)))

                                                cmdCheckEje.Parameters(0).Value = vAnio

                                                If Convert.ToInt32(cmdCheckEje.ExecuteScalar()) = 0 Then
                                                    cmdInsertEje.Parameters(0).Value = vAnio
                                                    cmdInsertEje.ExecuteNonQuery()
                                                    contEje += 1
                                                End If
                                            End While
                                        End Using
                                    End Using
                                End Using
                            End Using

                            ' 🌟 CERROJO DE ORO: Recogemos errores generales y cerramos las conexiones principales
                        Catch ex As Exception
                            MsgBox(rmse.GetString("ErrorTransferenciaApuntes") & ": " & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Error"))
                        End Try
                    End Using ' 🔒 Cierra biológico definitivo de connDestino
                End Using ' 🔒 Cierra biológico definitivo de connClon

                ' =========================================================================
                ' 🚀 FASE C: PURGA DE LIMPIEZA TOTAL Y NUEVO CARTEL ÚNICO
                ' =========================================================================
                Try
                    If File.Exists(RutaClonMigrada) Then File.Delete(RutaClonMigrada)
                Catch
                End Try
                ' El gran cartel de la victoria comercial premium final
                MsgBox(rmse.GetString("ImportaciónRelacionalCompletada"), MsgBoxStyle.Information, rmse.GetString("ActualizacionCompletada"))
                ' =========================================================================
                ' 🌟 APAGAMOS LA REDONDA QUE GIRA: Libertad al ratón pase lo que pase
                ' =========================================================================
            End If
        Me.Cursor = Cursors.Default
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmSeleccionEjercicio)
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
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmPreferencias)
        ' Llamamos al formulario de manera modal.
        frmPreferencias.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmPreferencias.Dispose()
        Me.TsLabelFormulario.Text = rmse.GetString("MsgEspera")
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
                MsgBox(resManager.GetString("ErrorVaciarApuntes") & ": " & ex.Message, MsgBoxStyle.Critical)
            End Try

            ' B. Eliminar Registros Apuntes Periódicos
            cmdMdb1cr.CommandText = "DELETE FROM apuper WHERE EjercicioAPP = ?"
            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.Add("@eje", OleDbType.Integer).Value = Convert.ToInt32(vAñoEjercicio)
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(rmse.GetString("ApuntesPeriodicosVaciado"), MsgBoxStyle.Information)
            Catch ex As Exception
                MsgBox(resManager.GetString("ErrorVaciarPeriodicos") & ": " & ex.Message, MsgBoxStyle.Critical)
            End Try

            ' C. Eliminar Registros Presupuestos
            cmdMdb1cr.CommandText = "DELETE FROM presupuesto WHERE EjercicioPRE = ?"
            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.Add("@eje", OleDbType.Integer).Value = Convert.ToInt32(vAñoEjercicio)
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(rmse.GetString("PresupuestosVaciado"), MsgBoxStyle.Information)
            Catch ex As Exception
                MsgBox(resManager.GetString("ErrorVaciarPresupuesto") & ": " & ex.Message, MsgBoxStyle.Critical)
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
        ' =========================================================================
        ' 🚀 CONFIGURACIÓN DE RUTA SEGURA COMPATIBLE CON MICROSOFT STORE (MSIX)
        ' =========================================================================
        ' Creamos una carpeta de Backups dócil y libre de derechos dentro de "Mis Documentos"
        Dim carpetaBackupSegura As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ContaHogar_Backups")

        Try
            If Not Directory.Exists(carpetaBackupSegura) Then
                Directory.CreateDirectory(carpetaBackupSegura)
            End If
        Catch ex As Exception
            ' Cortafuegos por si acaso, si falla cae directamente al directorio raíz de Documentos
            carpetaBackupSegura = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        End Try

        ' Preparamos el nombre cronológico del archivo de copia de seguridad
        Dim NombreBaseDatos As String = "ContaHogar3.0" & "[" & Now.ToString("ddMMyyyy") & "]" & "[" & Now.ToString("HHmmss") & "]" & ".mdb"
        Dim DataBaseFile As String = vRuta

        ' Configuramos el objeto de diálogo oficial de Windows de forma elástica
        backup.InitialDirectory = carpetaBackupSegura
        backup.Title = "Backup BD Access - ContaHogar 3.0"
        backup.CheckFileExists = False
        backup.CheckPathExists = True
        backup.DefaultExt = "mdb"
        backup.FileName = NombreBaseDatos
        backup.Filter = "Access (ContaHogar*.mdb)|ContaHogar*.mdb|All files (*.*)|*.*"
        backup.RestoreDirectory = True

        If backup.ShowDialog = Windows.Forms.DialogResult.OK Then
            Try
                ' 🚀 LA JUGADA MAESTRA: Copiamos la base de datos a la ruta EXACTA elegida por el usuario
                ' Al ser una acción explícita en el SaveFileDialog, Windows otorga inmunidad total de escritura
                Dim FileDestinoReal As String = backup.FileName

                FileCopy(DataBaseFile, FileDestinoReal)
                MessageBox.Show(rmse.GetString("BackupOk"), "BACKUP", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MsgBox(resManager.GetString("ErrorCrearCopiaSeguridad") & ": " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub

    Private Sub BtnRestaurarCopia_Click(sender As Object, e As EventArgs) Handles BtnRestaurarCopia.Click
        RestaurarCopiaDeSeguridadToolStripMenuItem.PerformClick()
    End Sub

    Private Sub RestaurarCopiaDeSeguridadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestaurarCopiaDeSeguridadToolStripMenuItem.Click
        Dim respuesta As MsgBoxResult = ConfirmarAccionTraducida(rmse.GetString("PreguntaBackup"), rmse.GetString("RestaurarBD"))
        If respuesta = vbYes Then

            ' =========================================================================
            ' 🚀 CONFIGURACIÓN DE RUTA SEGURA COMPATIBLE CON MICROSOFT STORE (MSIX)
            ' =========================================================================
            ' Apuntamos a la misma carpeta dócil y libre de derechos dentro de "Mis Documentos"
            Dim carpetaBackupSegura As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ContaHogar_Backups")

            ' Cortafuegos por si la carpeta aún no existiera en el perfil del usuario
            If Not Directory.Exists(carpetaBackupSegura) Then
                carpetaBackupSegura = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            End If

            Dim RestoreFile As String = vRuta ' Tu ruta oficial en Mis Documentos configurada en el Load

            ' Configuramos el objeto de diálogo de forma elástica e inmune a bloqueos
            restore.InitialDirectory = carpetaBackupSegura
            restore.Title = rmse.GetString("RestaurarBD")
            restore.CheckFileExists = True ' 🚀 OBLIGATORIO: El archivo debe existir de verdad para poder restaurarlo
            restore.CheckPathExists = True
            restore.DefaultExt = "mdb"
            restore.Filter = "Access (ContaHogar*.mdb)|ContaHogar*.mdb|All files (*.*)|*.*"
            restore.RestoreDirectory = True

            If restore.ShowDialog = Windows.Forms.DialogResult.OK Then
                Try
                    ' 🚀 LA JUGADA MAESTRA: Machacamos tu archivo de producción vRuta en Mis Documentos
                    ' usando el archivo legítimo seleccionado por el usuario en la ventana
                    FileCopy(restore.FileName, RestoreFile)
                    MessageBox.Show(rmse.GetString("RestaurarOk"), rmse.GetString("Restaurar"), MessageBoxButtons.OK, MessageBoxIcon.Information)

                    ' Avisamos de forma dócil y cerramos para asentar los hilos de la base de datos
                    MsgBox(resManager.GetString("CerrarApp"), vbInformation)
                    Me.Close()

                Catch ex As Exception
                    MsgBox(rmse.GetString("ErrorCriticoRestauracion") & ": " & ex.Message, MsgBoxStyle.Critical)
                End Try
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