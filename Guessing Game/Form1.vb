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
        'Dim disableErrorHandlerMsgs As Boolean 'haha, im learning by looking at the code

        Randomize()
        Number = Int(Rnd() * 100) + 1 'Selects Random Number
        Correct = False
        Count = 1
        Dim disableErrorHandlerMsgs As Boolean = True 'this code does not disable all error handler stuff, and only hides the "an error occured".
        Dim isDebugMode As Boolean = True 'i can make this one line, less compiler optimization!
        Dim exitTyped As Boolean = False

        If Val(isDebugMode) = True Then
            MsgBox("DEBUG: the answer is " & Number)
        End If
ErrorHandler:
        If Val(disableErrorHandlerMsgs) = False Then
            MsgBox("An error occurred, did you type a non-number? (error details): " & Err.Description & " and DEH is " & disableErrorHandlerMsgs)
            Err.Clear()
        End If

        Do While Count < 11 And Correct = False 'Gives user 10 guesses
            disableErrorHandlerMsgs = False 'Reenable error messages incase of an exception
            On Error GoTo ErrorHandler 'attribution: chatgpt :skull:

            Guess = InputBox("Enter Guess", "Attempt " & Count & " of 10. type 'exit' to quit the game.") 'i cant get "type 'exit' to exit" kind of thing.

            If Guess = "exit" Then
                exitTyped = True
                GoTo ExitGame
            End If

            'overly complicated number checker that took me 2 hours to implent along with the label ErrorHandler
            If IsNumeric(Guess) And Int(Guess) = Guess Then
                If isDebugMode = True Then
                    MsgBox("is a number")
                End If
            Else
                disableErrorHandlerMsgs = True 'reenable because we're going to the label... could have made a better approach to this issue of it still resuming but it works and i will never touch it ever ever again.
                MsgBox("Invalid input! Please enter (whole) number!")
                GoTo ErrorHandler
            End If

            If Val(Guess) < 1 Then 'mitigate negative numbers
                disableErrorHandlerMsgs = True
                MsgBox("Please enter a number higher than 1")
                GoTo ErrorHandler
            End If
            If Val(Guess) > 99 Then
                disableErrorHandlerMsgs = True
                MsgBox("Please enter a number lower than 100")
                GoTo ErrorHandler
            End If

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
            If Count = 1 Then
                MsgBox("Horray! you guessed it first try, have a coffee! The correct guess was: " & Number)
            Else
                MsgBox("YAY You guessed it at count " & Count & ", have a coffee! The correct guess was: " & Number)
            End If
        ElseIf Count > 10 Then
                MsgBox("Oops, You ran out of guesses! The correct guess was: " & Number & " and your last guess was: " & Guess) 'no there will be no nullpointer at runtime lmao
        End If
ExitGame:
        If Val(exitTyped) = True Then
            MsgBox("You forfeited, that's okay! The correct guess was: " & Number)
        End If
    End Sub
End Class
