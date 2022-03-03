Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles BtnExit.Click
        End
    End Sub

    Private Sub BtnStart_Click(sender As Object, e As EventArgs) Handles BtnStart.Click
        Dim Count, Number As Integer
        Dim Correct As Boolean
        Dim Guess As String

        Randomize()
        Number = Int(Rnd() * 100) + 1 'Selects Random Number
        Correct = False
        Count = 1

        Do While Count < 11 And Correct = False 'Gives user 10 guesses
            Guess = InputBox("Enter Guess", "Attempt " & Count & " of 10")
            If Val(Guess) = Number Then 'If guess is correct set correct to true
                Correct = True
            Else
                If Val(Guess) < Number Then
                    MsgBox("Too Low", vbExclamation, "Attempt " & Count & " of 10")
                ElseIf Guess > Number Then
                    MsgBox("Too High", vbExclamation, "Attempt " & Count & " of 10")
                End If

                Count = Count + 1
            End If
        Loop
        If Count < 11 Then
            MsgBox("YAY You guessed it, have a coffee!. The correct guess was:" & Number)
        ElseIf Count > 10 Then
            MsgBox("Oops, You ran out of guesses!. The correct guess was:" & Number)
        End If
    End Sub
End Class
