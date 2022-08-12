Module Emsa
    Public gbl_WattsConsumidos As Double
    Public gbl_superior As Double
    Public gbl_Primeros100 As Double
    Public gbl_Segundos50 As Double
    Public gbl_Terceros50 As Double
    Public Function fgbl_Primeros100() As Double
        Dim resultado As Double
        resultado = 0.1
        Return resultado
    End Function

    Public Function fgbl_Segundos50() As Double
        Dim resultado As Double
        resultado = 0.25
        Return resultado
    End Function
    Public Function fgbl_Terceros50() As Double
        Dim resultado As Double
        resultado = 0.5
        Return resultado
    End Function

    Public Function fgbl_superior()
        Dim resultado As Double
        resultado = 0.85
        Return resultado
    End Function


    Public Function CalcularPrimeros100(consumo As Double) As Double
        Dim resultado As Double
        resultado = consumo * fgbl_Primeros100()
        Return resultado

    End Function
    Public Function CalcularSegundos50(consumo As Double) As Double
        Dim resultado As Double
        resultado = consumo * fgbl_Segundos50()
        Return resultado

    End Function
    Public Function CalcularTerceros50(consumo As Double) As Double
        Dim resultado As Double
        resultado = consumo * fgbl_Terceros50()

        Return resultado

    End Function
    Public Function CalcularSuperior(consumo As Double) As Double
        Dim resultado As Double
        resultado = consumo * fgbl_superior()
        Return resultado
    End Function
    Public Function CalcularTotal(a As Double, b As Double, c As Double, d As Double) As Double
        Dim resultado As Double
        resultado = a + b + c + d
        Return resultado


    End Function
End Module
