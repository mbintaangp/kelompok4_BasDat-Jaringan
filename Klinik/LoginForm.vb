Imports System.Data.SqlClient

Public Class LoginForm
    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        Dim connectionString As String = ("Data Source=LAPTOP-13U1C95E; " &
                "user id = userol; password=123456; Integrated Security=True; " &
                "database=Klinik")
        Dim query As String = "SELECT COUNT(*) FROM Users WHERE username=@username AND password=@password"

        Using connection As New SqlConnection(connectionString)
            connection.Open()

            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@username", txtUsername.Text.Trim())
                command.Parameters.AddWithValue("@password", txtPassword.Text.Trim())

                Dim count As Integer = Convert.ToInt32(command.ExecuteScalar())

                If count > 0 Then
                    MessageBox.Show("Login successful!")

                    ' Lakukan pengalihan ke Form1 atau halaman berikutnya di sini
                    Dim form1 As New Form1()
                    form1.Show()
                    Me.Hide()
                Else
                    MessageBox.Show("Invalid username or password!")
                End If
            End Using
        End Using
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub
End Class
