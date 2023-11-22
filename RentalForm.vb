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
'[~]Finish user input validation **See Below**
'{~}checkOdometer *beginning Odometer should not be greater than ending odometer
'{~}CheckDays *day greater than 0 no more than 45

'[]Calculations:
'[~]Daily Charge is 15$ per day
'[~]mileage: first 200mi free, 201-500 12cents/mi, > 500mi 10cents/mi
'{~}all calc done in miles convert if Kilometer radio button is selected
'{~}1 Kilometer = 0.62 miles
'[]AAA members receive 5% discount
'[]senior citizens receive 3% discount
'{}both discounts can be used at once

'[]Display:
'{~}distance traveled in miles
'{~}total mileage charge as currency
'{~}total daily charge as currency
'{}total discount as currency
'{}total charges as currency

'[]Summary:
'{}only display if at least one rental has been completed
'{}display total # of customers
'{}display total distance in miles
'{}display total charges
'{}clear form **do not clear summary totals**

'[~]Set Defaults and Clear
'[]Add Close program confirmation box

Public Class RentalForm
    'Custom Methods
    ''' <summary>
    ''' Clears all Input and Output text boxes, Un-checks discounts, checks Miles button
    ''' </summary>
    Sub SetDefaults()
        'Clear input text boxes
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DaysTextBox.Text = ""
        'Clear output text boxes
        TotalMilesTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        DayChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalChargeTextBox.Text = ""
        'Clear discount check boxes
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
        'Select Miles RadioButton
        MilesradioButton.Checked = True
    End Sub

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

    ''' <summary>
    ''' Checks if Beginning and End odometer readings are numbers and end larger than beginning
    ''' </summary>
    ''' <returns></returns>
    Function CheckOdemeter() As Boolean
        Dim isValid As Boolean = True
        Dim errorMessage As String = ""
        Dim beginOdometer As Integer
        Dim endOdometer As Integer
        'try to convert BeginOdometerTextBox contents to integer
        Try
            beginOdometer = CInt(BeginOdometerTextBox.Text)
        Catch ex As Exception
            'beginning odometer not a number
            errorMessage = "Beginning Odometer Reading Must be a Number" & vbCrLf
            BeginOdometerTextBox.Focus()
            BeginOdometerTextBox.Text = ""
        End Try
        'Try to convert EndOdometerTextBox content to integer
        Try
            endOdometer = CInt(EndOdometerTextBox.Text)
        Catch ex As Exception
            'ending odometer not a number
            errorMessage &= "Ending Odometer Reading Must be a Number" & vbCrLf
            EndOdometerTextBox.Focus()
            EndOdometerTextBox.Text = ""
            isValid = False
        End Try
        'Check if begin odometer is larger than end if exception was not already raised
        If isValid Then
            If beginOdometer > endOdometer Then
                isValid = False
                errorMessage = "Ending Odometer Reading must be larger than Beginning Odometer Reading"
                BeginOdometerTextBox.Focus()
                'if error message is not empty report message to user
            End If
        End If
        If errorMessage <> "" Then
            MsgBox(errorMessage)
        End If
        Return isValid
    End Function

    ''' <summary>
    ''' Checks if number of days is between 1 and 45 days
    ''' </summary>
    ''' <returns></returns>
    Function CheckDays() As Boolean
        Dim isValid As Boolean = True
        Dim errorMessage As String = ""
        Dim numberOfDays As Integer
        'try to convert DaysTextBox contents to integer
        Try
            numberOfDays = CInt(DaysTextBox.Text)
        Catch ex As Exception
            'number of days not a number
            isValid = False
            DaysTextBox.Text = ""
            DaysTextBox.Focus()
            errorMessage = "Number of Days must be a Number"
        End Try
        'Check if number of days is in acceptable range (1-45) if exception was not already raised
        If isValid Then
            Select Case numberOfDays
                Case 1 To 45
                    'number of days is in acceptable range
                Case Else
                    'number of days is not in acceptable range
                    isValid = False
                    DaysTextBox.Text = ""
                    DaysTextBox.Focus()
                    errorMessage = "Number of Days must be more than 0 and no more than 45"
            End Select
        End If
        'if errorMessage is not empty report message to user
        If errorMessage <> "" Then
            MsgBox(errorMessage)
        End If
        Return isValid
    End Function

    ''' <summary>
    ''' Returns Total cost over a given number days at $15/Day
    ''' </summary>
    ''' <returns></returns>
    Function CalculateDaysCharge(numberOfDays As Integer) As Integer
        Dim daysCharge As Integer
        daysCharge = numberOfDays * 15
        Return daysCharge
    End Function

    ''' <summary>
    ''' Returns the difference of endPoint to startPoint 
    ''' </summary>
    ''' <param name="startPoint"></param>
    ''' <param name="endPoint"></param>
    ''' <returns></returns>
    Function CalculateDistanceTraveled(startPoint As Double, endPoint As Double) As Double
        Dim distance As Double
        distance = System.Math.Round(endPoint - startPoint, 2, MidpointRounding.ToEven)
        Return distance
    End Function

    ''' <summary>
    ''' Returns the mileage charge for a given distance
    ''' </summary>
    ''' <param name="distance"></param>
    ''' <returns></returns>
    Function CalculateMileageCharge(distance As Double) As Double
        Dim mileageCharge As Double
        'calculate charge based on distance here
        Select Case distance
            Case 0 To 200
                'free
                mileageCharge = 0
            Case 201 To 500
                '12 cents per mile
                mileageCharge = System.Math.Round(distance * 0.12, 2, MidpointRounding.ToEven)
            Case > 500
                '10 cents per mile
                mileageCharge = System.Math.Round(distance * 0.1, 2, MidpointRounding.ToEven)
        End Select
        Return mileageCharge
    End Function

    ''' <summary>
    ''' Calculates all charges based on text box info.  Should only Run if input already validated 
    ''' </summary>
    Sub CalculateAllCharges()
        Dim daysCharge As Integer
        Dim startOdometer As Double = CDbl(Me.BeginOdometerTextBox.Text)
        Dim endOdometer As Double = CDbl(Me.EndOdometerTextBox.Text)
        Dim distanceDriven As Double
        Dim mileageCharge As Double
        Const kilometerToMilesRatio As Double = 0.62
        If KilometersradioButton.Checked = True Then
            'convert to miles
            startOdometer = startOdometer / kilometerToMilesRatio
            endOdometer = endOdometer / kilometerToMilesRatio
        End If
        'already in miles or now converted
        'calculate days charge and update output text box
        daysCharge = CalculateDaysCharge(CInt(Me.DaysTextBox.Text))
        DayChargeTextBox.Text = FormatCurrency(daysCharge)
        'Determine distance driven
        distanceDriven = CalculateDistanceTraveled(startOdometer, endOdometer)
        TotalMilesTextBox.Text = CStr(distanceDriven) & " mi"
        'Calculate millage charge and update output text box
        mileageCharge = CalculateMileageCharge(distanceDriven)
        MileageChargeTextBox.Text = FormatCurrency(mileageCharge)
    End Sub

    'Event Handlers
    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        SetDefaults()
    End Sub
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Me.Close()
    End Sub
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        If ValidateUserInput() Then
            'all text boxes full start to check content of boxes
            'Check odometer readings
            If CheckOdemeter() Then
                'Check days
                If CheckDays() Then
                    'Run calculations
                    CalculateAllCharges()
                End If
            End If
        End If

    End Sub
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        SetDefaults()
    End Sub
End Class
