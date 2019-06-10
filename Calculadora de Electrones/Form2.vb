Public Class Form2
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If numero = 1 Then
            lbl_1.Text = 1
            lbl_2.Text = 0
            lbl_3.Text = 0
            lbl_4.Text = 0
            lbl_5.Text = 0
            lbl_6.Text = 0
        End If
        If numero = 0 Then
            lbl_1.Text = 0
            lbl_2.Text = 0
            lbl_3.Text = 0
            lbl_4.Text = 0
            lbl_5.Text = 0
            lbl_6.Text = 0
        End If
        If numero > 1 And numero <= 12 Then
            Modulo.descomponer()
            If resto = 0 Then
                Modulo.pares()
            End If
            If resto = 1 Then
                Modulo.impares()
            End If
        End If
    End Sub
End Class