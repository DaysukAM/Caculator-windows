Public Class Form1
    'Définission de result Afin d'afficher le résultat et s'en servir pour le bouton Ans (Answer)
    Dim result As Double

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub CalcBox_TextChanged(sender As Object, e As EventArgs)

    End Sub
    Private Sub ResultBox_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        CalcBox.Text = CalcBox.Text + "1"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        CalcBox.Text = CalcBox.Text + "2"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        CalcBox.Text = CalcBox.Text + "3"
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        CalcBox.Text = CalcBox.Text + "4"
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        CalcBox.Text = CalcBox.Text + "5"
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        CalcBox.Text = CalcBox.Text + "6"
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        CalcBox.Text = CalcBox.Text + "7"
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        CalcBox.Text = CalcBox.Text + "8"
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        CalcBox.Text = CalcBox.Text + "9"
    End Sub

    Private Sub Button0_Click(sender As Object, e As EventArgs) Handles Button0.Click
        CalcBox.Text = CalcBox.Text + "0"
    End Sub

    Private Sub ButtonPoint_Click(sender As Object, e As EventArgs) Handles ButtonPoint.Click
        CalcBox.Text = CalcBox.Text + ","
    End Sub

    'Bouton Ans (Answer) qui permet à l'utilisateur d'écrire la réponse dans le calcul (Arrondi au millième près)
    Private Sub ButtonAns_Click(sender As Object, e As EventArgs) Handles ButtonAns.Click
        If result = vbEmpty Then
        Else
            result = Math.Round(result, 3)
            CalcBox.Text = CalcBox.Text + CStr(result)
        End If
    End Sub

    Private Sub ButtonPlus_Click(sender As Object, e As EventArgs) Handles ButtonPlus.Click
        CalcBox.Text = CalcBox.Text + "+"
    End Sub

    Private Sub ButtonMin_Click(sender As Object, e As EventArgs) Handles ButtonMin.Click
        CalcBox.Text = CalcBox.Text + "-"
    End Sub

    Private Sub ButtonMult_Click(sender As Object, e As EventArgs) Handles ButtonMult.Click
        CalcBox.Text = CalcBox.Text + "x"
    End Sub

    Private Sub ButtonDiv_Click(sender As Object, e As EventArgs) Handles ButtonDiv.Click
        CalcBox.Text = CalcBox.Text + "/"
    End Sub

    Private Sub ButtonResult_Click(sender As Object, e As EventArgs) Handles ButtonResult.Click

        If CStr(Calcul(CalcBox.Text)) = "Error" Then
        Else
            result = Calcul(CalcBox.Text)
        End If
        ResultBox.Text = Calcul(CalcBox.Text)
        CalcBox.Text = ""
    End Sub

    Private Sub ButtonAllClear_Click(sender As Object, e As EventArgs) Handles ButtonAllClear.Click
        CalcBox.Text = ""
    End Sub

    Private Sub ButtonClear_Click(sender As Object, e As EventArgs) Handles ButtonClear.Click
        If CalcBox.Text.Length > 0 Then
            CalcBox.Text = CalcBox.Text.Remove(CalcBox.Text.Length - 1, 1)
        End If
    End Sub

    'Fonction de Calcul et de vérification d'erreur
    Private Function Calcul(CalcString As String)
        '1-Algorythme de vérificationet séparation
        Dim CalcArray() As Char = CalcString
        Dim nb(2) As String
        'y est un compteur pour le nombre
        Dim y As Integer = 0
        'x est un compteur pour l'opérateur
        Dim x As Integer = 0
        Dim op(1) As String
        'verification que le champ n'est pas vide
        If CalcString = "" Then
            Return 0
        End If
        'verification premier et dernier caractère
        If IsNumeric(CalcArray(0)) And IsNumeric(CalcArray(CalcArray.Length - 1)) Then
        Else
            Return "Error"
        End If
        'Boucle de parcours de la chaine de caractère 

        For i As Integer = 0 To CalcArray.Length - 1

            'verification du nombre
            If IsNumeric(CalcArray(i)) Then
                nb(y) = nb(y) & CalcArray(i)
            Else
                'système de verification des , pour nombre décimal
                If CalcArray(i) = "," Then
                    nb(y) = nb(y) & CalcArray(i)
                    'verification du caractère juste après la ,
                    If i + 1 >= CalcArray.Length Then
                        Return "Error"
                    ElseIf IsNumeric(CalcArray(i + 1)) Then
                    Else
                        Return "Error"
                    End If
                Else
                    'verifiction du 1er opérateur
                    Select Case CalcArray(i)
                        Case "+"
                            op(x) = "+"
                        Case "-"
                            op(x) = "-"
                        Case "x"
                            op(x) = "x"
                        Case "/"
                            op(x) = "/"
                    End Select
                    'verification du caractère juste après l'opérateur
                    If i + 1 >= CalcArray.Length Then
                        Return "Error"
                    ElseIf IsNumeric(CalcArray(i + 1)) Then
                        x = x + 1
                    Else
                        Return "Error"
                    End If
                    y = y + 1
                    nb(y) = ""

                End If
            End If

        Next
        'Vérification pour cas de nombre à plusieurs virgules
        For i As Integer = 0 To nb.Length - 1
            If (Len(nb(i)) - Len(Replace(nb(i), ",", "", , , 0))) / Len(",") > 1 Then
                Return "Error"
            End If
        Next
        '2-Algorythme de calcul en fonction de l'opérateur
        'Redefinition des 3 nombres en Double (nbTemp ets le resultat du premier calcul)
        Dim nb0, nb1, nb2, nbTemp As Double
        nb0 = CDbl(nb(0))
        nb1 = CDbl(nb(1))
        nb2 = CDbl(nb(2))
        'verification qu'il y a bien un opérateur
        If op(0) = "" Then
            Return nb0
        End If
        'verification de la priorité opératoire
        If op(1) = "x" Or op(1) = "/" Then
            'Calcul quand le 2eme opérateur est prioritaire
            Select Case op(1)
                Case "x"
                    'gestion du cas exeptionnel de la division en premier
                    If op(0) = "/" Then
                        nbTemp = nb0 / nb1
                        result = nbTemp * nb2
                        Return result
                    Else
                        nbTemp = nb1 * nb2
                    End If

                Case "/"
                    'gestion du cas exeptionnel des doubles divisions
                    If op(0) = "/" Then
                        nbTemp = nb0 / nb1
                        result = nbTemp / nb2
                        Return result
                    Else
                        nbTemp = nb1 / nb2
                    End If
            End Select
            Select Case op(0)
                Case "+"
                    result = nb0 + nbTemp
                    Return result
                Case "-"
                    result = nb0 - nbTemp
                    Return result
                Case "x"
                    result = nb0 * nbTemp
                    Return result
            End Select
        Else
            'Calcul quand le 1eme opérateur est prioritaire (ou quand il n'y a pas de 2eme opérateur)
            Select Case op(0)
                Case "+"
                    nbTemp = nb0 + nb1
                Case "-"
                    nbTemp = nb0 - nb1
                Case "x"
                    nbTemp = nb0 * nb1
                Case "/"
                    nbTemp = nb0 / nb1

            End Select
            Select Case op(1)
                Case "+"
                    result = nbTemp + nb2
                    Return result
                Case "-"
                    result = nbTemp - nb2
                    Return result
            End Select

        End If
        'si il n'y a pas de 2eme opérateur, la fonction renvoie nbTemp car il correspond au resultat du calcul simple à un opérateur
        Return nbTemp
    End Function
End Class
