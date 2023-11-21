'Zachary Christensen
'RCET 2265
'Fall 2023
'Car Rental
'https://github.com/Minidude140/CarRental-ZBC.git

Option Explicit On
Option Strict On
Option Compare Binary

'TODO
'[~]fix order of text box response
'[]Finish user input validation **See Below**
'{}checkOdometer *beginning Odometer should not be greater than ending odemeter
'{}CheckDays *day greater than 0 no more than 45

'[]Calculations:
'[]Daily Charge is 15$ per day
'[]mileage: first 200mi free, 201-500 12cents/mi, > 500mi 10cents/mi
'{}all calc done in miles convert if Kilometer radio button is selected
'{}1 Kilometer = 0.62 miles
'{}if reading in kilometers return in kilometers
'[]AAA members receive 5% discount
'[]senoir citizens receive 3% discount
'{}both discounts can be used as once

'[]Display:
'{}distance traveled in given units
'{}total millage charge as currency
'{}total daily charge as currency
'{}total discount as currency
'{}total charges as currency

'[]Summary:
'{}only display if at least one rental has been completed
'{}display total # of customers
'{}display total distance in miles
'{}display total charges
'{}clear form **do not clear summary totals**

'[]Set Defaults and Clear
'[]Add Close program confirmation box

Public Class RentalForm
    'Custom Methods

    ''' <summary>
    ''' Checks that each text field has something entered
    ''' </summary>
    Function ValidateUserInput() As Boolean
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
        Return isValid
    End Function
    'Event Handlers
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Me.Close()
    End Sub
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        If ValidateUserInput() Then
            'all text boxes full start to check content of boxes
        End If

    End Sub
End Class
