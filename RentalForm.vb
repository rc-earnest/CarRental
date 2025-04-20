'Rudy Earnest
'Car Rental Form
'RCET 2265
'Spring 2025
'https://github.com/rc-earnest/CarRental.git
Option Explicit On
Option Strict On
Option Compare Binary
Public Class RentalForm


    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        Dim errorMessage As String = ""
        Dim focusControl As Control = Nothing
        Dim beginningOdometer As Decimal
        Dim endingOdometer As Decimal
        Dim numberOfDays As Integer
        Dim distanceDriven As Decimal
        Dim mileageCharge As Decimal
        Dim chargeableMiles As Decimal
        Dim dailyCharge As Decimal
        Dim totalCharge As Decimal
        Dim discountAmount As Integer


        ' Validate Customer Information
        If String.IsNullOrWhiteSpace(NameTextBox.Text) Then
            errorMessage &= "Customer Name cannot be blank." & vbCrLf
            If focusControl Is Nothing Then focusControl = NameTextBox
        End If

        If String.IsNullOrWhiteSpace(AddressTextBox.Text) Then
            errorMessage &= "Address cannot be blank." & vbCrLf
            If focusControl Is Nothing Then focusControl = AddressTextBox
        End If

        If String.IsNullOrWhiteSpace(CityTextBox.Text) Then
            errorMessage &= "City cannot be blank." & vbCrLf
            If focusControl Is Nothing Then focusControl = CityTextBox
        End If

        If String.IsNullOrWhiteSpace(StateTextBox.Text) Then
            errorMessage &= "State cannot be blank." & vbCrLf
            If focusControl Is Nothing Then focusControl = StateTextBox
        End If

        If String.IsNullOrWhiteSpace(ZipCodeTextBox.Text) Then
            errorMessage &= "Zip Code Name cannot be blank." & vbCrLf
            If focusControl Is Nothing Then focusControl = ZipCodeTextBox
        End If

        ' Validate Odometer Readings
        If Not Decimal.TryParse(BeginOdometerTextBox.Text, beginningOdometer) Then
            errorMessage &= "Beginning Odometer must be a valid number." & vbCrLf
            BeginOdometerTextBox.Clear()
            If focusControl Is Nothing Then focusControl = BeginOdometerTextBox
        End If

        If Not Decimal.TryParse(EndOdometerTextBox.Text, endingOdometer) Then
            errorMessage &= "Ending Odometer must be a valid number." & vbCrLf
            EndOdometerTextBox.Clear()
            If focusControl Is Nothing Then focusControl = EndOdometerTextBox
        End If

        If Not String.IsNullOrWhiteSpace(BeginOdometerTextBox.Text) AndAlso Not String.IsNullOrWhiteSpace(EndOdometerTextBox.Text) AndAlso beginningOdometer >= endingOdometer Then
            errorMessage &= "Beginning Odometer must be less than Ending Odometer." & vbCrLf
            BeginOdometerTextBox.Clear()
            EndOdometerTextBox.Clear()
            If focusControl Is Nothing Then focusControl = BeginOdometerTextBox
        End If

        ' Validate Number of Days
        If Not Integer.TryParse(DaysTextBox.Text, numberOfDays) Then
            errorMessage &= "Number of Days must be a valid whole number." & vbCrLf
            DaysTextBox.Clear()
            If focusControl Is Nothing Then focusControl = DaysTextBox
        ElseIf numberOfDays <= 0 Then
            errorMessage &= "Number of Days must be greater than zero." & vbCrLf
            DaysTextBox.Clear()
            If focusControl Is Nothing Then focusControl = DaysTextBox
        ElseIf numberOfDays > 45 Then
            errorMessage &= "Number of Days cannot be greater than 45." & vbCrLf
            DaysTextBox.Clear()
            If focusControl Is Nothing Then focusControl = DaysTextBox
        End If

        If errorMessage <> "" Then
            MessageBox.Show(errorMessage, "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If Not IsNothing(focusControl) Then
                focusControl.Focus()
            End If
            Return
        End If

        'Number Crunching
        If KilometersradioButton.Checked Then
            distanceDriven = CInt((endingOdometer - beginningOdometer) * 0.62)
        Else
            distanceDriven = CInt(endingOdometer - beginningOdometer)
        End If

        If distanceDriven > 200 Then
            chargeableMiles = CInt(distanceDriven - 200)
            If chargeableMiles <= (500 - 200) Then
                mileageCharge = CInt(chargeableMiles * 0.12)
            Else
                mileageCharge = CInt((500 - 200) * 0.12 + (chargeableMiles - (500 - 200)) * 0.1)
            End If
        End If

        dailyCharge = CInt(numberOfDays * 15)
        totalCharge = dailyCharge + mileageCharge


        If AAAcheckbox.Checked Then
            discountAmount += CInt(totalCharge * 0.05)
        End If

        If Seniorcheckbox.Checked Then
            discountAmount += CInt(totalCharge * 0.03)
        End If

        totalCharge -= discountAmount


        'Display Output
        TotalMilesTextBox.Text = $"{distanceDriven:N2} mi"
        MileageChargeTextBox.Text = $"{mileageCharge:C}"
        DayChargeTextBox.Text = $"{dailyCharge:C}"
        TotalDiscountTextBox.Text = $"{discountAmount:C}"
        TotalChargeTextBox.Text = $"{totalCharge:C}"
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        ' Clear all text boxes
        NameTextBox.Clear()
        AddressTextBox.Clear()
        CityTextBox.Clear()
        StateTextBox.Clear()
        ZipCodeTextBox.Clear()
        BeginOdometerTextBox.Clear()
        EndOdometerTextBox.Clear()
        DaysTextBox.Clear()
        TotalMilesTextBox.Clear()
        MileageChargeTextBox.Clear()
        DayChargeTextBox.Clear()
        TotalDiscountTextBox.Clear()
        TotalChargeTextBox.Clear()

        ' Clear discount check boxes
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False

        ' Select the miles radio button
        MilesradioButton.Checked = True
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to exit?", "Exit Program", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Me.Close()
        End If
    End Sub
End Class
