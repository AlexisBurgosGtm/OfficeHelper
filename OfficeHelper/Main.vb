Imports DevExpress.XtraSplashScreen

Public Class Main

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.txtFecha.DateTime = Today.Date


        Dim stfecha As String
        stfecha = Me.txtFecha.DateTime.Day.ToString & "/" & Me.txtFecha.DateTime.Month.ToString & "/" & Me.txtFecha.DateTime.Year.ToString
        Me.txtMotivo.Text = "Daily Distribuidora Popular Reu - " & stfecha


        Call CargarMails()


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
End Class