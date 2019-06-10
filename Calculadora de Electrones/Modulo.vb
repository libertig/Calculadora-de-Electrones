Module Modulo
    Public dividendo As Integer
    Public divisor As Integer
    Public cociente As Integer
    Public resto As Integer
    Public numero As Integer

    Public Sub descomponer()
        dividendo = Form1.txt_numero.Text
        divisor = 2
        cociente = Trim(Math.Floor(dividendo / divisor))
        resto = Trim(dividendo Mod divisor)
    End Sub

    Public Sub pares()
        If resto = 0 Then
            If cociente = 1 Then
                Form2.lbl_1.Text = 2
                Form2.lbl_2.Text = 0
                Form2.lbl_3.Text = 0
                Form2.lbl_4.Text = 0
                Form2.lbl_5.Text = 0
                Form2.lbl_6.Text = 0
            End If
            If cociente = 2 Then
                Form2.lbl_1.Text = 2
                Form2.lbl_2.Text = 2
                Form2.lbl_3.Text = 0
                Form2.lbl_4.Text = 0
                Form2.lbl_5.Text = 0
                Form2.lbl_6.Text = 0
            End If
            If cociente = 3 Then
                Form2.lbl_1.Text = 2
                Form2.lbl_2.Text = 2
                Form2.lbl_3.Text = 2
                Form2.lbl_4.Text = 0
                Form2.lbl_5.Text = 0
                Form2.lbl_6.Text = 0
            End If
            If cociente = 4 Then
                Form2.lbl_1.Text = 2
                Form2.lbl_2.Text = 2
                Form2.lbl_3.Text = 2
                Form2.lbl_4.Text = 2
                Form2.lbl_5.Text = 0
                Form2.lbl_6.Text = 0
            End If
            If cociente = 5 Then
                Form2.lbl_1.Text = 2
                Form2.lbl_2.Text = 2
                Form2.lbl_3.Text = 2
                Form2.lbl_4.Text = 2
                Form2.lbl_5.Text = 2
                Form2.lbl_6.Text = 0
            End If
            If cociente = 6 Then
                Form2.lbl_1.Text = 2
                Form2.lbl_2.Text = 2
                Form2.lbl_3.Text = 2
                Form2.lbl_4.Text = 2
                Form2.lbl_5.Text = 2
                Form2.lbl_6.Text = 2
            End If
        End If
    End Sub

    Public Sub impares()
        If resto = 1 Then
            If cociente = 1 Then
                Form2.lbl_1.Text = 2
                Form2.lbl_2.Text = 1
                Form2.lbl_3.Text = 0
                Form2.lbl_4.Text = 0
                Form2.lbl_5.Text = 0
                Form2.lbl_6.Text = 0
            End If
            If cociente = 2 Then
                Form2.lbl_1.Text = 2
                Form2.lbl_2.Text = 2
                Form2.lbl_3.Text = 1
                Form2.lbl_4.Text = 0
                Form2.lbl_5.Text = 0
                Form2.lbl_6.Text = 0
            End If
            If cociente = 3 Then
                Form2.lbl_1.Text = 2
                Form2.lbl_2.Text = 2
                Form2.lbl_3.Text = 2
                Form2.lbl_4.Text = 1
                Form2.lbl_5.Text = 0
                Form2.lbl_6.Text = 0
            End If
            If cociente = 4 Then
                Form2.lbl_1.Text = 2
                Form2.lbl_2.Text = 2
                Form2.lbl_3.Text = 2
                Form2.lbl_4.Text = 2
                Form2.lbl_5.Text = 1
                Form2.lbl_6.Text = 0
            End If
        End If
    End Sub
End Module
