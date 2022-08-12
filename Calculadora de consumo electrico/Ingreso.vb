Public Class Ingreso

    Private Sub Calcular_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Emsa.gbl_WattsConsumidos = Convert.ToDouble(TextBox1.Text)
        Resultado.Show()
        Hide()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Validacion.SoloNumeros(e)
    End Sub
End Class