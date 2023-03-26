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
        Dim disableErrorHandlerMsgs As Boolean 'haha, im learning by looking at the code

        Randomize()
        Number = Int(Rnd() * 100) + 1 'Selects Random Number
        Correct = False
        Count = 1
        disableErrorHandlerMsgs = True 'NOTICE: this code does not disable all error handler stuff, and only hides the "an error occured".
        '^ on first pass, ignore it because we're passing through (i could make this a lot better but who cares)

ErrorHandler:
        If Val(disableErrorHandlerMsgs) = False Then
            MsgBox("An error occurred, did you type a non-number? (error details): " & Err.Description & " and DEH is " & disableErrorHandlerMsgs)
            Err.Clear()
        Else 'its disabled
        End If

        'could add a questionbox thingy whatever that asks "how much guesses you want, 10 is default" mode? maybe i'll work on that next year

        Do While Count < 11 And Correct = False 'Gives user 10 guesses
            disableErrorHandlerMsgs = False 'Reenable error messages incase of an exception
            On Error GoTo ErrorHandler 'attribution: chatgpt :skull:

            Guess = InputBox("Enter Guess", "Attempt " & Count & " of 10")
            'overly complicated number checker that took me 2 hours to implent along with the label ErrorHandler
            If IsNumeric(Guess) Then
                MsgBox("is a number")
                'do not put an "exit do" here, this will completely skip all checks at checking if the number is low or not and stuff, which is how it added a hour to developing like 20 lines of code
            Else
                disableErrorHandlerMsgs = True 'reenable because we're going to the label
                'could have made a better approach to this issue of it still resuming but it works and i will never touch it ever ever again.
                MsgBox("Invalid input! Please enter (whole) number!")
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
            MsgBox("YAY You guessed it, have a coffee! The correct guess was: " & Number & " debug: your inputted 'guess' value is " & Guess) 'no, there will be no null-ref exception on runtime i think
        ElseIf Count > 10 Then
            MsgBox("Oops, You ran out of guesses! The correct guess was: " & Number & " debug: your inputted 'guess' value is " & Guess)
        End If
    End Sub
End Class
