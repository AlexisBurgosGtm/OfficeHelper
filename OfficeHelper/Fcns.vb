Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Text
Imports DevExpress.XtraBars.Docking2010.Customization
Imports DevExpress.XtraBars.Docking2010.Views.WindowsUI
Imports Microsoft.Win32
'importaciones para el envio de correo
Imports System
Imports System.Collections
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Mime
Imports DevExpress.XtraEditors.Camera
Imports System.IO
Imports System.Xml
Imports Security
Imports System.Speech.Synthesis
Imports System.ComponentModel
'**********************************

Module Fcns


#Region " ** BITACORA ** "
    Public Sub InsertarBitacora(ByVal Usuario As String, ByVal CodModulo As Integer, ByVal Accion As String, ByVal Importe As Double)
        Dim Modulo As String
        Select Case CodModulo
            Case 0 'login
                Modulo = "Inicio de Sesión"
            Case 1 'inicio
                Modulo = "Menú principal"
            Case 2 'pedidos
                Modulo = "Pedidos"
            Case 3 'ventas
                Modulo = "Facturación"
            Case 4 'compras
                Modulo = "Compras"
            Case 5 'corte caja
                Modulo = "Corte de caja"
            Case 6 'gastos
                Modulo = "Gastos"
            Case 7 'configuraciones
                Modulo = "Configuraciones"
            Case 8 'reportes
                Modulo = "Reportes"
            Case 9 'mantenimientos
                Modulo = "Mantenimientos"
            Case 10 'vendedores
                Modulo = "Vendedores"
            Case 11 'clientes
                Modulo = "Clientes"
            Case 12 'proveedores
                Modulo = "Proveedores"
            Case 13 'productos
                Modulo = "Productos"
            Case 14 'inventarios
                Modulo = "Inventarios"
            Case 15 'listados
                Modulo = "Listados"
            Case 16 'cotizaciones
                Modulo = "Cotizaciones"
            Case 17 'devoluciones
                Modulo = "Devoluciones"
        End Select
        Try
            Call AbrirConexionSql()
            Dim cmd As New SqlCommand("INSERT INTO BITACORA (ANIO,MES,DIA,FECHA,HORA,MINUTO,USUARIO,MODULO,ACCION,IMPORTE) VALUES (@ANIO,@MES,@DIA,@FECHA,@HORA,@MINUTO,@USUARIO,@MODULO,@ACCION,@IMPORTE)", cn)
            cmd.Parameters.AddWithValue("@ANIO", Today.Date.Year)
            cmd.Parameters.AddWithValue("@MES", Today.Date.Month)
            cmd.Parameters.AddWithValue("@DIA", Today.Date.Day)
            cmd.Parameters.AddWithValue("@FECHA", Today.Date)
            cmd.Parameters.AddWithValue("@HORA", Now.Hour)
            cmd.Parameters.AddWithValue("@MINUTO", Now.Minute)
            cmd.Parameters.AddWithValue("@USUARIO", Usuario)
            cmd.Parameters.AddWithValue("@MODULO", Modulo)
            cmd.Parameters.AddWithValue("@ACCION", Accion)
            cmd.Parameters.AddWithValue("@IMPORTE", Importe)
            cmd.ExecuteNonQuery()
            cmd.Dispose()

        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region " ** NUMEROS A LETRAS ** "
    Public Function Letras(ByVal numero As String) As String
        '********Declara variables de tipo cadena************
        Dim palabras, entero, dec, flag As String

        '********Declara variables de tipo entero***********
        Dim num, x, y As Integer

        flag = "N"

        '**********Número Negativo***********
        If Mid(numero, 1, 1) = "-" Then
            numero = Mid(numero, 2, numero.ToString.Length - 1).ToString
            palabras = "menos "
        End If

        '**********Si tiene ceros a la izquierda*************
        For x = 1 To numero.ToString.Length
            If Mid(numero, 1, 1) = "0" Then
                numero = Trim(Mid(numero, 2, numero.ToString.Length).ToString)
                If Trim(numero.ToString.Length) = 0 Then palabras = ""
            Else
                Exit For
            End If
        Next

        '*********Dividir parte entera y decimal************
        For y = 1 To Len(numero)
            If Mid(numero, y, 1) = "." Then
                flag = "S"
            Else
                If flag = "N" Then
                    entero = entero + Mid(numero, y, 1)
                Else
                    dec = dec + Mid(numero, y, 1)
                End If
            End If
        Next y

        If Len(dec) = 1 Then dec = dec & "0"

        '**********proceso de conversión***********
        flag = "N"

        If Val(numero) <= 999999999 Then
            For y = Len(entero) To 1 Step -1
                num = Len(entero) - (y - 1)
                Select Case y
                    Case 3, 6, 9
                        '**********Asigna las palabras para las centenas***********
                        Select Case Mid(entero, num, 1)
                            Case "1"
                                If Mid(entero, num + 1, 1) = "0" And Mid(entero, num + 2, 1) = "0" Then
                                    palabras = palabras & "cien "
                                Else
                                    palabras = palabras & "ciento "
                                End If
                            Case "2"
                                palabras = palabras & "doscientos "
                            Case "3"
                                palabras = palabras & "trescientos "
                            Case "4"
                                palabras = palabras & "cuatrocientos "
                            Case "5"
                                palabras = palabras & "quinientos "
                            Case "6"
                                palabras = palabras & "seiscientos "
                            Case "7"
                                palabras = palabras & "setecientos "
                            Case "8"
                                palabras = palabras & "ochocientos "
                            Case "9"
                                palabras = palabras & "novecientos "
                        End Select
                    Case 2, 5, 8
                        '*********Asigna las palabras para las decenas************
                        Select Case Mid(entero, num, 1)
                            Case "1"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    flag = "S"
                                    palabras = palabras & "diez "
                                End If
                                If Mid(entero, num + 1, 1) = "1" Then
                                    flag = "S"
                                    palabras = palabras & "once "
                                End If
                                If Mid(entero, num + 1, 1) = "2" Then
                                    flag = "S"
                                    palabras = palabras & "doce "
                                End If
                                If Mid(entero, num + 1, 1) = "3" Then
                                    flag = "S"
                                    palabras = palabras & "trece "
                                End If
                                If Mid(entero, num + 1, 1) = "4" Then
                                    flag = "S"
                                    palabras = palabras & "catorce "
                                End If
                                If Mid(entero, num + 1, 1) = "5" Then
                                    flag = "S"
                                    palabras = palabras & "quince "
                                End If
                                If Mid(entero, num + 1, 1) > "5" Then
                                    flag = "N"
                                    palabras = palabras & "dieci"
                                End If
                            Case "2"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "veinte "
                                    flag = "S"
                                Else
                                    palabras = palabras & "veinti"
                                    flag = "N"
                                End If
                            Case "3"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "treinta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "treinta y "
                                    flag = "N"
                                End If
                            Case "4"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "cuarenta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "cuarenta y "
                                    flag = "N"
                                End If
                            Case "5"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "cincuenta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "cincuenta y "
                                    flag = "N"
                                End If
                            Case "6"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "sesenta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "sesenta y "
                                    flag = "N"
                                End If
                            Case "7"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "setenta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "setenta y "
                                    flag = "N"
                                End If
                            Case "8"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "ochenta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "ochenta y "
                                    flag = "N"
                                End If
                            Case "9"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "noventa "
                                    flag = "S"
                                Else
                                    palabras = palabras & "noventa y "
                                    flag = "N"
                                End If
                        End Select
                    Case 1, 4, 7
                        '*********Asigna las palabras para las unidades*********
                        Select Case Mid(entero, num, 1)
                            Case "1"
                                If flag = "N" Then
                                    If y = 1 Then
                                        palabras = palabras & "uno "
                                    Else
                                        palabras = palabras & "un "
                                    End If
                                End If
                            Case "2"
                                If flag = "N" Then palabras = palabras & "dos "
                            Case "3"
                                If flag = "N" Then palabras = palabras & "tres "
                            Case "4"
                                If flag = "N" Then palabras = palabras & "cuatro "
                            Case "5"
                                If flag = "N" Then palabras = palabras & "cinco "
                            Case "6"
                                If flag = "N" Then palabras = palabras & "seis "
                            Case "7"
                                If flag = "N" Then palabras = palabras & "siete "
                            Case "8"
                                If flag = "N" Then palabras = palabras & "ocho "
                            Case "9"
                                If flag = "N" Then palabras = palabras & "nueve "
                        End Select
                End Select

                '***********Asigna la palabra mil***************
                If y = 4 Then
                    If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or
                    (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And
                    Len(entero) <= 6) Then palabras = palabras & "mil "
                End If

                '**********Asigna la palabra millón*************
                If y = 7 Then
                    If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                        palabras = palabras & " millón "
                    Else
                        palabras = palabras & " millones "
                    End If
                End If
            Next y

            '**********Une la parte entera y la parte decimal*************
            If dec <> "" Then
                Letras = palabras & "con " & dec & "/100"
            Else
                Letras = palabras
            End If
        Else
            Letras = ""
        End If
    End Function
#End Region


    Public Sub HablarTexto(ByVal texto As String)
        Dim Hablar As New SpeechSynthesizer
        Hablar.SpeakAsync(texto)
        Hablar.Dispose()
    End Sub


    Public cn As New SqlConnection

    Public Sub AbrirConexionSql()
        If cn.State = 0 Then
            cn.Open()
        End If
    End Sub


    'OBTIENE LA IMAGEN DEL CAMPO BINARIO DE SQL
    Public Function ObtenerImgSql(ByVal DrField As Object) As Image
        Try
            Dim imageData As Byte()

            If DrField Is DBNull.Value Then
                Return Nothing
            Else
                imageData = DrField
                Using ms As New MemoryStream(imageData, 0, imageData.Length)
                    ms.Write(imageData, 0, imageData.Length)
                    Return Image.FromStream(ms, True)
                End Using
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function DiaSemanaPago(ByVal fecha As Date) As Integer
        Return Weekday(fecha, 2)
    End Function

    Public Function CompaniaCel(ByVal Telefono As String) As String
        Try
            Dim PrimerosDigitos As Integer

            PrimerosDigitos = Integer.Parse(Mid(Telefono, 1, 4))

            If cn.State = 0 Then
                cn.Open()
            End If

            Dim dr As SqlDataReader
            Dim cmd As New SqlCommand("SELECT Empresa FROM MSysTemp WHERE PrimerosDigitos=" & PrimerosDigitos, cn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                Return dr(0).ToString
                'Select Case dr(0).ToString
                'Case "CLARO"
                '         'Return My.Resources.LogoClaro

                'Case "TIGO"
                ' Return My.Resources.LogoTigo
                'Case "MOVISTAR"
                ' Return My.Resources.LogoMovistar
                'End Select
            Else
                'Return Nothing
                Return "SN"
            End If

            dr.Close()
            cmd.Dispose()
            dr = Nothing
            cn.Close()
        Catch ex As Exception
        End Try
    End Function


    Public Sub Aviso(ByVal Titulo As String, ByVal Mensaje As String, ByVal Owner As Form)
        Dim action As New FlyoutAction()
        action.Caption = Titulo
        action.Description = Mensaje
        action.Commands.Add(FlyoutCommand.OK)

        FlyoutDialog.Show(Owner, action)

    End Sub

    Public Function Confirmacion(ByVal Mensaje As String, ByVal Owner As Form) As Boolean
        Dim action As New FlyoutAction()
        action.Caption = "Confirme"
        action.Description = Mensaje
        action.Commands.Add(FlyoutCommand.OK)
        action.Commands.Add(FlyoutCommand.Cancel)

        '  action.Image = Image.FromFile("")
        Dim x As Integer
        x = FlyoutDialog.Show(Owner, action)

        If x = DialogResult.OK Then
            Return True
        Else
            Return False
        End If

    End Function



    Public Sub CargarConexion()
        Dim servidor As String, UserDb As String, UserPass As String, db As String
        Dim MyServidor, MyUsuario, MyPass As String
        Dim MyPort As Integer

        Try
            Const fic As String = "C:\AresPOS\Cn"
            Dim sr As New System.IO.StreamReader(fic)
            db = sr.ReadLine()
            servidor = sr.ReadLine()
            UserDb = sr.ReadLine()
            UserPass = sr.ReadLine()
            MyServidor = sr.ReadLine()
            MyUsuario = sr.ReadLine()
            MyPass = sr.ReadLine()
            MyPort = CType(sr.ReadLine(), Integer)
            'GlobalSocketIp = sr.ReadLine()
            'GlobalSocketPort = CType(sr.ReadLine(), Integer)
            sr.Close()

            Dim ps As New Seguridad("razors1805")
            Dim Pass As String = ps.DecryptData(UserPass)
            Dim User As String = ps.DecryptData(UserDb)

            cn = New SqlConnection("Data Source=" & servidor & ";Initial Catalog=" & db & ";User ID=" & User & ";Password=" & Pass & ";MultipleActiveResultSets=True")

        Catch ex As Exception
            MessageBox.Show("UPS !! El periodo de prueba ha terminado, debes solicitar una versión completa al correo SistemaAresPos@gmail.com o al número +502-5725-5092 (Whatsapp)")
            End
        End Try
        ' CnMysql = New MySql.Data.MySqlClient.MySqlConnection("Server=localhost;Database=ares;Uid=iEx;Pwd=iEx;Port=3306") '" & MyUsuario & ";Pwd=" & MyPass & ";Port=" & MyPort & "")
    End Sub

    'customize
    Public GlobalSkin As String, TipoPOS As String



    Public Sub CrearPathAccess() 'ByVal RutaCarpeta As String) 'crea la Ubicación de Confianza para la app access
        Try
            Dim RutaCarpeta As String = "C:\iExPrestamos\"
            Dim KeyPathOffice2010 As String = "SOFTWARE\Microsoft\Office\14.0\Access\Security\Trusted Locations\iExperts"
            Dim KeyPathOffice2013 As String = "SOFTWARE\Microsoft\Office\15.0\Access\Security\Trusted Locations\iExperts"
            Dim ValueName As String = "Path"

            Registry.CurrentUser.CreateSubKey(KeyPathOffice2010)
            Registry.CurrentUser.CreateSubKey(KeyPathOffice2013)

            Dim key2010 As RegistryKey = Registry.CurrentUser.OpenSubKey(KeyPathOffice2010, True) ' True indica que se abre para escritura
            key2010.SetValue(ValueName, RutaCarpeta, RegistryValueKind.String)
            Dim key2013 As RegistryKey = Registry.CurrentUser.OpenSubKey(KeyPathOffice2013, True) ' True indica que se abre para escritura
            key2013.SetValue(ValueName, RutaCarpeta, RegistryValueKind.String)
            MessageBox.Show("Path creado exitosamente")

        Catch ex As Exception
            MessageBox.Show("No se creó el Path - " & ex.ToString)
        End Try

    End Sub

    Public Sub EnviarGmail(ByVal Subject As String, ByVal Body As String, ByVal Destino As String)

        Try

            Dim MMessage As System.Net.Mail.MailMessage = New MailMessage
            MMessage.To.Add(Destino)
            MMessage.From = New MailAddress(GmailEmisor, GmailEmisor, System.Text.Encoding.UTF8)
            MMessage.Subject = Subject
            MMessage.SubjectEncoding = System.Text.Encoding.UTF8
            MMessage.Body = Body
            MMessage.BodyEncoding = System.Text.Encoding.UTF8
            MMessage.IsBodyHtml = False

            Dim sClient As New SmtpClient()
            sClient.Credentials = New System.Net.NetworkCredential(GmailEmisor, GmailEmisorPass)
            sClient.Host = "smtp.gmail.com"
            sClient.Port = 587

            sClient.EnableSsl = True


            sClient.Send(MMessage)
        Catch ex As System.Net.Mail.SmtpException
            '   MsgBox(ex.ToString)
        End Try

    End Sub
    Public Sub EnviarGmailAdjunto(ByVal Subject As String, ByVal Body As String, ByVal Destino As String, ByVal RutaAdjunto As String)

        Try

            Dim MMessage As System.Net.Mail.MailMessage = New MailMessage
            MMessage.To.Add(Destino)
            MMessage.From = New MailAddress(GmailEmisor, GmailEmisor, System.Text.Encoding.UTF8)
            MMessage.Subject = Subject
            MMessage.SubjectEncoding = System.Text.Encoding.UTF8
            MMessage.Body = Body
            MMessage.BodyEncoding = System.Text.Encoding.UTF8
            MMessage.IsBodyHtml = False
            Dim archivo As Net.Mail.Attachment = New Net.Mail.Attachment(RutaAdjunto)
            MMessage.Attachments.Add(archivo)

            Dim sClient As New SmtpClient()
            sClient.Credentials = New System.Net.NetworkCredential(GmailEmisor, GmailEmisorPass)
            sClient.Host = "smtp.gmail.com"
            sClient.Port = 587

            sClient.EnableSsl = True


            sClient.Send(MMessage)
        Catch ex As System.Net.Mail.SmtpException
            '   MsgBox(ex.ToString)
        End Try

    End Sub

    Public Function EnviarGmailAttach(ByVal Subject As String, ByVal Body As String, ByVal Destino As String, ByVal RutaAdjunto As String) As Boolean

        Try

            Dim MMessage As System.Net.Mail.MailMessage = New MailMessage
            MMessage.To.Add(Destino)
            MMessage.From = New MailAddress(GmailEmisor, GmailEmisor, System.Text.Encoding.UTF8)
            MMessage.Subject = Subject
            MMessage.SubjectEncoding = System.Text.Encoding.UTF8
            MMessage.Body = Body
            MMessage.BodyEncoding = System.Text.Encoding.UTF8
            MMessage.IsBodyHtml = False
            Dim archivo As Net.Mail.Attachment = New Net.Mail.Attachment(RutaAdjunto)
            MMessage.Attachments.Add(archivo)

            Dim sClient As New SmtpClient()
            sClient.Credentials = New System.Net.NetworkCredential(GmailEmisor, GmailEmisorPass)
            sClient.Host = "smtp.gmail.com"
            sClient.Port = 587

            sClient.EnableSsl = True

            sClient.Send(MMessage)

            Return True
        Catch ex As System.Net.Mail.SmtpException
            '   MsgBox(ex.ToString)
            Return False
        End Try

    End Function

    'variables globales para correo y pass del gmail
    Public GmailEmisor As String
    Public GmailEmisorPass As String


    Public Sub EnviarGmail2(ByVal Subject As String, ByVal Body As String)

        Dim MMessage As System.Net.Mail.MailMessage = New MailMessage
        MMessage.To.Add("ralexmailreu@gmail.com")
        MMessage.From = New MailAddress(GmailEmisor, GmailEmisor, System.Text.Encoding.UTF8)
        MMessage.Subject = Subject
        MMessage.SubjectEncoding = System.Text.Encoding.UTF8
        MMessage.Body = Body
        MMessage.BodyEncoding = System.Text.Encoding.UTF8
        MMessage.IsBodyHtml = False

        Dim sClient As New SmtpClient()
        sClient.Credentials = New System.Net.NetworkCredential(GmailEmisor, GmailEmisorPass)
        sClient.Host = "smtp.gmail.com"
        sClient.Port = 587

        sClient.EnableSsl = True

        Try
            sClient.Send(MMessage)
        Catch ex As System.Net.Mail.SmtpException
            '   MsgBox(ex.ToString)
        End Try

    End Sub


#Region " ** LICENCIA DE USO ** "
    Public Function AlexisKey() As Boolean
        Dim MesVence, AnioVence As Integer

        'establecer el mes y año para el demo
        AnioVence = 2018
        MesVence = 2

        Try

            Dim readValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\VB and VBA Program Settings\AresPOS\Datos PC", "PcId", Nothing)
            If readValue.ToString = "Ares" Then
                Return True
            End If
        Catch ex As Exception

            'si el mes es febrero, retorna falso,
            'de lo contrario va a retornar verdadero
            If Today.Date.Year <= AnioVence Then

                If Today.Date.Month >= MesVence Then
                    Return False
                Else
                    Return True
                End If

                Return False

            End If
        End Try

    End Function

    Public Sub CrearValue() 'crea un registro en el registro de win para compararlo desde la app access

        Try
            Dim KeyPath As String = "Software\VB and VBA Program Settings\AresPOS\Datos PC"
            Dim ValueName As String = "PcId"

            Registry.CurrentUser.CreateSubKey(KeyPath)

            Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey(KeyPath, True) ' True indica que se abre para escritura
            key.SetValue(ValueName, "Ares", RegistryValueKind.String)
            MessageBox.Show("Key agregada")
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

#End Region


End Module
