Option Strict Off
Module ImpConversionofBooltoSingle
    Function Main() As Integer
        Dim b As Boolean = True
        Dim a As Single = b
        If a <> -1 Then
            System.Console.WriteLine("Implicit Conversion of Bool(Tru):return 1 to Single has Failed. Expected -1 got " & a)
        End If
    End Function
End Module
