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
        Dim beginningOdometer As Decimal
        If Not Decimal.TryParse(BeginOdometerTextBox.Text, beginningOdometer) Then
            errorMessage &= "Beginning Odometer must be a valid number." & vbCrLf
            BeginOdometerTextBox.Clear()
            If focusControl Is Nothing Then focusControl = BeginOdometerTextBox
        End If

        Dim endingOdometer As Decimal
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
        Dim numberOfDays As Integer
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
            Application.Exit()
        End If
    End Sub
End Class
