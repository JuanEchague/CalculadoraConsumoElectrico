Public Class Resultado
    Private Sub Resultado_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = Emsa.gbl_WattsConsumidos


        If (gbl_WattsConsumidos >= 100) Then
            If (gbl_WattsConsumidos >= 150) Then
                If (gbl_WattsConsumidos >= 200) Then
                    TextBox2.Text = Emsa.CalcularPrimeros100(100)
                    TextBox3.Text = Emsa.CalcularSegundos50(50)
                    TextBox4.Text = Emsa.CalcularTerceros50(50)
                    TextBox5.Text = Emsa.CalcularSuperior(gbl_WattsConsumidos - 200)
                    TextBox6.Text = "$" & Emsa.CalcularTotal(TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text)

                Else
                    TextBox2.Text = Emsa.CalcularPrimeros100(100)
                    TextBox3.Text = Emsa.CalcularSegundos50(50)
                    TextBox4.Text = Emsa.CalcularTerceros50(gbl_WattsConsumidos - 150)
                    TextBox5.Text = "0"
                    TextBox6.Text = "$" & Emsa.CalcularTotal(TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text)

                End If
            Else
                TextBox2.Text = Emsa.CalcularPrimeros100(100)
                TextBox3.Text = Emsa.CalcularSegundos50(gbl_WattsConsumidos - 100)
                TextBox4.Text = "0"
                TextBox5.Text = "0"
                TextBox6.Text = "$" & Emsa.CalcularTotal(TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text)

            End If
        Else

            TextBox2.Text = Emsa.CalcularPrimeros100(gbl_WattsConsumidos)
            TextBox3.Text = "0"
            TextBox4.Text = "0"
            TextBox5.Text = "0"
            TextBox6.Text = "$" & Emsa.CalcularTotal(TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text)

        End If

    End Sub

    Private Sub Volver_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Ingreso.Show()
        Hide()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
    End Sub
End Class
