Imports Microsoft.Win32
Imports System.IO
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.ServiceProcess

Public Class frmAnalizarPC
    'Inherits Responsive
    'Implements IResponsive



    Private PreviousClientSize As Size = Me.ClientSize
    Private url As String = "https://uc26fd291bd9001af0f34ab586c3.dl.dropboxusercontent.com/cd/0/get/CDVaKs_i1P6-W-Zqp82I9sHA9jfeUvz8Q021Yj2wECrxZ3eQgaYvIpXzJZ6hURBOgL-rR3f1zVhBxvPM79UoGrrwnZklgZH2JY7st3VSyZolDzVigWA-pTDmgBX1QAWlCJ61zPh6iBsdr3reKl3iIVy9QsUVy5VFyqL1IGlzQfT1tg/file?_download_id=18035308363901886217408120256598173350645871471758631148858686661&_notify_domain=www.dropbox.com&dl=1"
    'Private Shared ReadOnly serviceName As String = aSetVars(1)




    Private Sub frmAnalizarPC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim drives() As DriveInfo = DriveInfo.GetDrives()

        For Each drive As DriveInfo In drives
            CBoxDiscos.Items.Add(drive.Name)
        Next

        If CBoxDiscos.Items.Count > 0 Then
            CBoxDiscos.SelectedIndex = 0
        End If

    End Sub

    Private Sub frmAnalizarPC_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        e.Handled = True
        e.SuppressKeyPress = True
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        Else
            e.Handled = False
            e.SuppressKeyPress = False
        End If
    End Sub



    '----------------------------Metodos----------------------------------------------

#Region "Metodos"

    '------------------------------Sistema Operativo-------------------------------------------------
#Region "Sistema Operativo"

    Public Function IsWindows10ProOrAbove() As Boolean
        Dim osVersion As Version = Environment.OSVersion.Version
        Dim osName As String = ObtenerNombreSistemaOperativo()

        If (osName.Contains("7")) Or (osName.Contains("Home")) Then
            Return False

        Else
            Return True

        End If

    End Function


    Public Function ObtenerNombreSistemaOperativo() As String
        Dim osName As String = Environment.OSVersion.VersionString
        Return osName
    End Function

#End Region


    '------------------------------Procesador-------------------------------------------------
#Region "Procesador"


    Public Function ObtenerNombreProcesador() As String
        Dim processorName As String = "Desconocido"

        Try
            Dim regKey As RegistryKey = Registry.LocalMachine.OpenSubKey("HARDWARE\DESCRIPTION\System\CentralProcessor\0")
            If regKey IsNot Nothing Then
                processorName = CStr(regKey.GetValue("ProcessorNameString"))
                regKey.Close()
            End If
        Catch ex As Exception
            ' Manejo de excepciones si ocurre algún problema al acceder al registro
        End Try

        Return processorName
    End Function


    Protected Function ObtenerHzTotales() As Integer
        Try
            Dim category As PerformanceCounterCategory = New PerformanceCounterCategory("Processor Information")
            Dim counters() As String = category.GetInstanceNames()

            If counters.Length > 0 Then
                Dim clockSpeedCounter As New PerformanceCounter("Processor Information", "Processor Frequency", counters(0))
                Dim clockSpeed As Integer = CInt(clockSpeedCounter.RawValue)
                clockSpeedCounter.Dispose()
                Return clockSpeed
            End If
        Catch ex As Exception
            ' Manejo de excepciones si ocurre algún problema al acceder a los contadores de rendimiento
        End Try

        Return 0
    End Function


    'Protected Function ObtenerPorcentajeCPUenUso() As Double
    '    Dim cpuCounter As New PerformanceCounter("Processor", "% Processor Time", "_Total")
    '    cpuCounter.NextValue() ' La primera llamada a NextValue() suele devolver un valor incorrecto, así que lo descartamos.
    '    System.Threading.Thread.Sleep(1000) ' Esperamos un segundo para obtener una lectura más precisa.
    '    Dim cpuUsage As Double = cpuCounter.NextValue()
    '    cpuCounter.Close()
    '    Return cpuUsage

    'End Function


    Protected Function ObtenerPorcentajeCPUenUso() As Double
        Dim cpuCounter As New PerformanceCounter("Processor", "% Processor Time", "_Total")
        cpuCounter.NextValue() ' La primera llamada a NextValue() suele devolver un valor incorrecto, así que lo descartamos.
        System.Threading.Thread.Sleep(1000) ' Esperamos un segundo para obtener una lectura más precisa.
        Dim cpuUsage As Double = cpuCounter.NextValue()
        cpuCounter.Close()

        ' Redondear el resultado a dos decimales
        cpuUsage = Math.Round(cpuUsage, 2)

        Return cpuUsage
    End Function



    Public Function ObtenerHzDisponibles() As Double
        Dim P_Total As Integer = ObtenerHzTotales()
        Dim P_EnUso As Double = ObtenerPorcentajeCPUenUso() * 100
        Dim P_HzRestante As Double = P_Total - P_EnUso

        Return P_HzRestante
    End Function


#End Region


    '------------------------------Memoria-------------------------------------------------
#Region "Memoria"

    Public Function ObtenerMermoriaTotal() As Double
        Dim totalMemory As Double = CDbl(My.Computer.Info.TotalPhysicalMemory) / (1024 * 1024 * 1024)
        Return totalMemory
    End Function

    Public Function ObtnerMemoriaDisponible() As Double
        Dim MemoriaDisponible As Double = CDbl(My.Computer.Info.AvailablePhysicalMemory) / (1024 * 1024 * 1024)
        Return MemoriaDisponible
    End Function

    Public Function ObtenerPorcentajeMemoriaEnUso() As Double
        Try
            Dim MemoriaTotal As Double = ObtenerMermoriaTotal()
            Dim MemoriaDisponible As Double = ObtnerMemoriaDisponible()

            Dim PorcentajeEnUso As Double = ((MemoriaTotal - MemoriaDisponible) / MemoriaTotal) * 100

            Return Math.Round(PorcentajeEnUso, 2)
        Catch ex As Exception

            MessageBox.Show("Se produjo un error al obtener el porcentaje de memoria en uso: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return 0
        End Try
    End Function


#End Region


#Region "Velocidad Internet"
    Public Function CalcularVelocidadDescargaMB(url As String) As Double
        Try
            ' Crear un WebClient para descargar el archivo de prueba
            Using webClient As New WebClient()
                Dim startTime As DateTime = DateTime.Now
                Dim archivoTemporal As String = "tempfile.tmp"

                ' Descargar el archivo de prueba
                webClient.DownloadFile(url, archivoTemporal)

                Dim endTime As DateTime = DateTime.Now
                Dim tiempoTranscurrido As TimeSpan = endTime - startTime

                ' Calcular la velocidad de descarga en bytes por segundo
                Dim tamanoArchivo As Long = New System.IO.FileInfo(archivoTemporal).Length
                Dim velocidadDescargaBytesPorSegundo As Double = tamanoArchivo / tiempoTranscurrido.TotalSeconds

                ' Convertir la velocidad de descarga a megabytes por segundo
                Dim velocidadDescargaMBPorSegundo As Double = velocidadDescargaBytesPorSegundo / (1024.0 * 1024.0)

                ' Eliminar el archivo temporal
                System.IO.File.Delete(archivoTemporal)

                ' Devolver la velocidad de descarga en megabytes por segundo como un valor Double
                Return velocidadDescargaMBPorSegundo
            End Using
        Catch ex As Exception
            Throw New Exception("Error al calcular la velocidad de descarga: " & ex.Message)
        End Try
    End Function

#End Region


    '***************************Ethernet o Wifi*****************************


#Region "Conexion"
    Public Function TipoConexionRed() As String
        Dim interfaces As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()

        For Each netInterface As NetworkInterface In interfaces
            If netInterface.OperationalStatus = OperationalStatus.Up Then
                If netInterface.NetworkInterfaceType = NetworkInterfaceType.Ethernet Then

                    Return "Ethernet"
                ElseIf netInterface.NetworkInterfaceType = NetworkInterfaceType.Wireless80211 Then
                    Return "WiFi"
                End If
            End If
        Next

        Return "Desconocido"
    End Function
#End Region


#Region "Disco"
    Public Sub EspacioDisco()
        Dim selectedDrive As String = CBoxDiscos.SelectedItem.ToString()

        Dim driveInfo As New DriveInfo(selectedDrive)
        Dim EspacioLibre As Long = driveInfo.AvailableFreeSpace
        Dim EspacioGB As Double = EspacioLibre / (1024.0 * 1024.0 * 1024.0)

        TextBox5.Text = EspacioGB.ToString("N2") & " GB"

        If EspacioGB > 15 Then
            TextBox5.BackColor = Color.Green
            TextBox5.ForeColor = Color.White
        ElseIf EspacioGB < 15 And EspacioGB > 7 Then
            TextBox5.BackColor = Color.Yellow
            TextBox5.ForeColor = Color.Black
        Else
            TextBox5.BackColor = Color.Red
            TextBox5.ForeColor = Color.White
        End If
    End Sub
#End Region


#End Region




    '****************************Botones***************************************

#Region "Botones"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Procesador()

        Memoria()

        Pantalla()

        SistemaOperativo()


        Internet()

        EspacioDisco()

        'CalcularVelocidadDescargaMB(url)

    End Sub


    'Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
    '    Dim serviceController As New ServiceController(serviceName)

    '    Try
    '        ' Detener el servicio de SQL Server si está en ejecución
    '        If serviceController.Status = ServiceControllerStatus.Running Then
    '            serviceController.Stop()
    '            serviceController.WaitForStatus(ServiceControllerStatus.Stopped, New TimeSpan(0, 0, 30)) ' Esperar hasta que se detenga
    '        End If

    '        ' Iniciar el servicio de SQL Server
    '        serviceController.Start()
    '        serviceController.WaitForStatus(ServiceControllerStatus.Running, New TimeSpan(0, 0, 30)) ' Esperar hasta que se inicie

    '        MessageBox.Show("El servidor SQL Server se ha reiniciado con éxito.", "Aceptar", MessageBoxButtons.OK)
    '    Catch ex As Exception
    '        MessageBox.Show("Se produjo un error al reiniciar el SQL SERVER: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    Finally
    '        serviceController.Close()
    '    End Try
    'End Sub

#End Region



    '****************************Validaciones***************************************



#Region "Validaciones"

    '******************* Procesador ******************************************

#Region "Procesador"
    Private Sub Procesador()
        Try

            Dim CpuDisponible = ObtenerPorcentajeCPUenUso()

            TextBox1.Text = CpuDisponible.ToString & " %"

            If CpuDisponible < 40 Then
                TextBox1.BackColor = Color.Green
                TextBox1.ForeColor = Color.White
            ElseIf CpuDisponible > 40 And CpuDisponible < 70 Then
                TextBox1.BackColor = Color.Yellow
                TextBox1.ForeColor = Color.Black
            Else
                TextBox1.BackColor = Color.Red
                TextBox1.ForeColor = Color.White
            End If

        Catch ex As Exception
            MessageBox.Show("Se produjo un error al obtener la información del procesador: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region



    '*****************  Memoria *********************************************

#Region "Memoria"

    Private Sub Memoria()
        Try
            Dim MemoriaTotal As Double = ObtenerMermoriaTotal()
            Dim MemoriaDisponible As Double = ObtnerMemoriaDisponible()
            Dim MemoriaPorcentaje As Double = ObtenerPorcentajeMemoriaEnUso()

            TextBox3.Text = MemoriaPorcentaje.ToString("N2") & "%"

            If MemoriaPorcentaje < 40 Then
                TextBox3.BackColor = Color.Green
                TextBox3.ForeColor = Color.White
            ElseIf MemoriaPorcentaje > 40 And MemoriaPorcentaje < 70 Then
                TextBox3.BackColor = Color.Yellow
                TextBox3.ForeColor = Color.Black
            Else
                TextBox3.BackColor = Color.Red
                TextBox3.ForeColor = Color.Black
            End If
        Catch ex As Exception
            MessageBox.Show("Se produjo un error al obtener la información de la memoria: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region




    '*************Pantalla*******************

#Region "Pantalla"
    Private Sub Pantalla()
        Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
        Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height

        Dim resolucion = screenWidth.ToString & " x " & screenHeight.ToString
        TextBox7.Text = resolucion

        If screenHeight > 1024 Then  ' Alto
            TextBox7.BackColor = Color.Green
            TextBox7.ForeColor = Color.White
        Else
            TextBox7.BackColor = Color.Red
            TextBox7.ForeColor = Color.White
        End If

        If screenWidth > 720 Then  'Ancho
            TextBox7.BackColor = Color.Green
            TextBox7.ForeColor = Color.White
        Else
            TextBox7.BackColor = Color.Red
            TextBox7.ForeColor = Color.White
        End If
    End Sub
#End Region





    'Sistema Operativo **************************************************************

#Region "S Operativo"

    Private Sub SistemaOperativo()
        Try
            TextBox9.Text = My.Computer.Info.OSFullName.ToString

            If IsWindows10ProOrAbove() = True Then
                TextBox9.BackColor = Color.Green
                TextBox9.ForeColor = Color.White
            Else
                TextBox9.BackColor = Color.Red
                TextBox9.ForeColor = Color.White
            End If
        Catch ex As Exception
            MessageBox.Show("Se produjo un error al obtener la información del sistema operativo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region





#Region "Internet"

    Private Sub Internet()
        Dim Internet = TipoConexionRed()



        Try
            TextBox11.Text = Internet.ToString()

            If Internet = "Ethernet" Then
                TextBox11.BackColor = Color.Green
                TextBox11.ForeColor = Color.White
            Else
                TextBox11.BackColor = Color.Red
                TextBox11.ForeColor = Color.White
            End If
        Catch ex As Exception
            MessageBox.Show("Se produjo un error al obtener la información del sistema operativo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region



#Region "Velocidad Descarga"

    Private Sub VelocidadDescarga()
        Dim velocidadDescargaMB As Double = CalcularVelocidadDescargaMB(url)
        Try
            TextBox11.Text = velocidadDescargaMB.ToString()

            If velocidadDescargaMB = "   " Then
                TextBox13.BackColor = Color.Green
                TextBox13.ForeColor = Color.White
            Else
                TextBox13.BackColor = Color.Red
                TextBox13.ForeColor = Color.White
            End If
        Catch ex As Exception
            MessageBox.Show("Se produjo un error al obtener la información del sistema operativo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region






#End Region


    'Private Sub btnReq_Click(sender As Object, e As EventArgs) Handles btnReq.Click
    '    frmRequisitos.Show()
    'End Sub

    'Private Sub ayuda_Click(sender As Object, e As EventArgs) Handles ayuda.Click
    '    oComun.ayuda(935)
    'End Sub


End Class