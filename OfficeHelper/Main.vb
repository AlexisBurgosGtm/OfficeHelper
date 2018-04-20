Imports System.Speech.Recognition
Imports System.Threading
Imports DevExpress.XtraSplashScreen

Public Class Main

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.txtFecha.DateTime = Today.Date


        Dim stfecha As String
        stfecha = Me.txtFecha.DateTime.Day.ToString & "/" & Me.txtFecha.DateTime.Month.ToString & "/" & Me.txtFecha.DateTime.Year.ToString
        Me.txtMotivo.Text = "Daily Distribuidora Popular Reu - " & stfecha


        Call CargarMails()

        Me.TimerVoice.Enabled = True

    End Sub


#Region " ** RECONOCIMIENTO DE VOZ ** " 'BORRAR SI NO SIRVE


    Public activado As Boolean
    Public ESTANDO As Integer

    Dim _recognizer As SpeechRecognitionEngine = Nothing
    Dim manualResetEvent As ManualResetEvent = Nothing

    Dim REC As New SpeechRecognitionEngine
    Dim Mensaje As New Speech.Synthesis.SpeechSynthesizer
    Dim PALABRA As String
    Dim CONTADOR As Integer = 0
    Dim MIARRAY(CONTADOR) As String

    Private Sub SwitchVoice_Toggled(sender As Object, e As EventArgs) Handles SwitchVoice.Toggled
        Me.TimerVoice.Enabled = True
    End Sub


    Private Sub NORECONOCE()
        Beep()
    End Sub
    Private Sub DETECTA()
    End Sub
    Private Sub RECONOCE(ByVal sender As Object, ByVal e As SpeechRecognizedEventArgs)
        Dim result As RecognitionResult
        result = e.Result
        Dim word As String
        word = result.Text

        Dim RESULTADO As RecognitionResult
        RESULTADO = e.Result
        Dim PALABRA As String
        PALABRA = RESULTADO.Text

        Select Case PALABRA
            Case "Alessa, abre el control de rutas"
                Process.Start("E:\COMPARTIDA\NUEVO CONTROL DE RUTAS.xlsx")

            Case "Alessa, ¿qué hora es?"
                Mensaje.SpeakAsync("Son las" & Hour(Now).ToString & "horas, y" & Minute(Now) & "minutos")

            Case "Alessa, navegar al envio de correos"
                Me.NavFrameMain.SelectedPage = NP_Send

            Case "Alessa, navegar al inicio"
                Me.NavFrameMain.SelectedPage = NP_Inicio

            Case "Alessa, abre el sistema de ventas"
                Process.Start("C:\SHORCUTS\VENTAS.LNK")

            Case "Alessa, abre facturación"
                Process.Start("C:\SHORCUTS\APP.LNK")



        End Select

        'End If
    End Sub
    Private Sub TimerVoice_Tick(sender As Object, e As EventArgs) Handles TimerVoice.Tick
        If Me.SwitchVoice.IsOn = True Then

            Dim Vocabulario As New GrammarBuilder
            Vocabulario.Append(New Choices("Alessa, abre el control de rutas",
                                           "Alessa, ¿qué hora es?",
                                           "Alessa, navegar al envio de correos",
                                           "Alessa, navegar al inicio",
                                           "Alessa, abre el sistema de ventas",
                                           "Alessa, abre facturación"
                                                                               ))
            Try

                REC.LoadGrammar(New Grammar(Vocabulario))
                REC.SetInputToDefaultAudioDevice()
                REC.RecognizeAsync(RecognizeMode.Multiple)
                AddHandler REC.SpeechRecognized, AddressOf RECONOCE
                AddHandler REC.SpeechRecognitionRejected, AddressOf NORECONOCE
                AddHandler REC.SpeechDetected, AddressOf DETECTA

            Catch ex As Exception

            End Try
        End If

        TimerVoice.Enabled = False
    End Sub

#End Region

    Private Sub navBtn_Update_ElementClick(sender As Object, e As DevExpress.XtraBars.Navigation.NavElementEventArgs) Handles navBtn_Update.ElementClick
        Me.NavFrameMain.SelectedPage = NP_Update
    End Sub

    Private Sub btn_updateAtras_Click(sender As Object, e As EventArgs) Handles btn_updateAtras.Click
        Me.NavFrameMain.SelectedPage = NP_Inicio
    End Sub


    Private Sub btnEnviar_Click(sender As Object, e As EventArgs) Handles btnEnviar.Click
        If Confirmacion("¿Enviar?", Me) = True Then
            SplashScreenManager.ShowForm(Me, GetType(GlobalWaitForm), True, True)

            If EnviarGmailAttach(Me.txtMotivo.Text, "Saludos,", ObtenerCorreos, Me.txtPath.Text) = True Then
                SplashScreenManager.CloseForm()
                Call Aviso("Importante", "Se ha enviado el correo", Me)
            End If

            Try
                SplashScreenManager.CloseForm()
            Catch ex As Exception

            End Try


        End If
    End Sub

    Private Function ObtenerCorreos() As String
        Dim stmail As String

        If Me.txtMail.Text <> "" Then
            stmail = Me.txtMail.Text
        Else
            stmail = ""
        End If

        If Me.txtMail1.Text <> "" Then
            stmail = stmail & "," & Me.txtMail1.Text
        Else
            stmail = ""
        End If

        If Me.txtMail2.Text <> "" Then
            stmail = stmail & "," & Me.txtMail2.Text
        End If

        If Me.txtMail3.Text <> "" Then
            stmail = stmail & "," & Me.txtMail3.Text
        End If


        If Me.txtMail4.Text <> "" Then
            stmail = stmail & "," & Me.txtMail4.Text
        End If

        Return stmail
    End Function


    Private Sub CargarMails()

        Try
            Const fic As String = "C:\OfficceHelper\Mails.txt"
            Dim sr As New System.IO.StreamReader(fic)
            GmailEmisor = sr.ReadLine()
            GmailEmisorPass = sr.ReadLine()
            Me.txtPath.Text = sr.ReadLine()
            Me.txtMail.Text = sr.ReadLine()
            Me.txtMail1.Text = sr.ReadLine()
            Me.txtMail2.Text = sr.ReadLine()
            Me.txtMail3.Text = sr.ReadLine()
            Me.txtMail4.Text = sr.ReadLine()

            sr.Close()

            If Me.txtMail.Text = "SN" Then Me.txtMail.Text = ""
            If Me.txtMail1.Text = "SN" Then Me.txtMail1.Text = ""
            If Me.txtMail2.Text = "SN" Then Me.txtMail2.Text = ""
            If Me.txtMail3.Text = "SN" Then Me.txtMail3.Text = ""
            If Me.txtMail4.Text = "SN" Then Me.txtMail4.Text = ""

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            End
        End Try

    End Sub

    Private Sub navbtn_clientes_ElementClick(sender As Object, e As DevExpress.XtraBars.Navigation.NavElementEventArgs) Handles navbtn_clientes.ElementClick
        Me.NavFrameMain.SelectedPage = NP_Clientes
        Dim objClientes As New ClassGeneral
        Me.GridClientes.DataSource = Nothing
        Me.GridClientes.DataSource = objClientes.tblClientes
    End Sub

    Private Sub navbtn_daily_ElementClick(sender As Object, e As DevExpress.XtraBars.Navigation.NavElementEventArgs) Handles navbtn_daily.ElementClick
        Me.NavFrameMain.SelectedPage = NP_Send

    End Sub

    Private Sub btnSendHome_Click(sender As Object, e As EventArgs) Handles btnSendHome.Click
        Me.NavFrameMain.SelectedPage = NP_Inicio
    End Sub

    Private Sub btnClientesHome_Click(sender As Object, e As EventArgs) Handles btnClientesHome.Click
        Me.NavFrameMain.SelectedPage = NP_Inicio
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Try
            Me.GridClientes.ExportToHtml("C:\OfficceHelper\tabla.txt")
            MsgBox("Done!!")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub btnUpdate_exportar_Click(sender As Object, e As EventArgs) Handles btnUpdate_exportar.Click
        Me.GridUpdate.ExportToXlsx("C:\OfficceHelper\Update.xlsx")
    End Sub


End Class