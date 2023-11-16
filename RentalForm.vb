'Zachary Christensen
'RCET 2265
'Fall 2023
'Car Rental
'https://github.com/Minidude140/CarRental-ZBC.git

Option Explicit On
Option Strict On
Option Compare Binary

'TODO
'[]fix order of textbox response
Public Class RentalForm
    'Custom Methods

    ''' <summary>
    ''' Checks that each text field has something entered
    ''' </summary>
    Sub ValidateUserInput()
        Dim isValid As Boolean = True
        Dim errorMessage As String
        'checks if each text box is empty
        For Each Item As TextBox In CustomerInputGroupBox.Controls.OfType(Of TextBox)
            If isValid = True Then
                Item.Focus()
            End If
            If Item.Text = "" Then
                isValid = False
                errorMessage &= Replace($"{Item.Name.ToString} is required{vbCrLf}", "TextBox", "")
            End If
        Next
        'Message the user if an error has occurred
        If errorMessage <> "" Then
            MsgBox(errorMessage)
        End If
    End Sub
    'Event Handlers
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Me.Close()
    End Sub
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        ValidateUserInput()
    End Sub
End Class
