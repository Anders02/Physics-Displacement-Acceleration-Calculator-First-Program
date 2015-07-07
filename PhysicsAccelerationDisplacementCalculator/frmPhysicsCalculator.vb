'Purpose:       The purpose of the Physics Acceleration and Displacement Calculator 
'               is to calculate acceleration, Displacement and time from user input values.
'Developer:     Anders Blair
'Date:          7/1/2015



Option Strict On
Public Class frmPhysicsCalculator
    'Declare Class variables
    Dim dblAcceleration As Double
    Dim dblTime As Double
    Dim dblDisplacement As Double


    Private Sub frmPhysicsCalculator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'clears dummy values
        lblResult.Text = ""
        lblResult2.Text = ""
        txtAcceleration.Focus()
        MsgBox("Note: This program assumes there is no initial velocity.")
    End Sub

    Private Sub radAcceleration_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radAcceleration.CheckedChanged
        'Clears textboxes and labels. Disables the acceleration textbox.
        'Sets focus on txtDisplacement.
        txtAcceleration.Enabled = False
        txtDisplacement.Enabled = True
        txtTime.Enabled = True
        txtDisplacement.Clear()
        txtAcceleration.Clear()
        txtTime.Clear()
        lblResult.Text = ""
        lblResult2.Text = ""
        txtDisplacement.Focus()
    End Sub

    Private Sub radDisplacement_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radDisplacement.CheckedChanged
        'clears textboxes and labels. Disables displacement textbox.
        'Sets focus to acceleration textbox.
        txtAcceleration.Enabled = True
        txtDisplacement.Enabled = False
        txtTime.Enabled = True
        txtDisplacement.Clear()
        txtDisplacement.Clear()
        txtAcceleration.Clear()
        txtTime.Clear()
        lblResult.Text = ""
        lblResult2.Text = ""
        txtAcceleration.Focus()
    End Sub

    Private Sub radTime_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radTime.CheckedChanged
        'clears textboxes and labels.
        'Sets focus to displacement textbox.
        txtAcceleration.Enabled = True
        txtDisplacement.Enabled = True
        txtTime.Enabled = False
        txtDisplacement.Clear()
        txtAcceleration.Clear()
        txtTime.Clear()
        lblResult.Text = ""
        lblResult2.Text = ""
        txtDisplacement.Focus()
    End Sub
    Private Function CheckIfDisplacementValid() As Boolean
        'Check if the data that the user entered for the displacement can be saved as double and is over 0.
        Dim blnIsValid As Boolean = False
        'See if data entered is valid using try-catch statements.
        Try
            dblDisplacement = Convert.ToDouble(txtDisplacement.Text)
            If dblDisplacement > 0 Then
                blnIsValid = True
            Else
                MsgBox("Please enter a number greater than 0 for the displacement.")
            End If
        Catch exception As FormatException
            txtDisplacement.Text = ""
            txtDisplacement.Focus()
            MsgBox("Please enter a number.")
        Catch exception As OverflowException
            txtDisplacement.Text = ""
            txtDisplacement.Focus()
            MsgBox("Displacement entered is too high.")
        Catch exception As SystemException
            txtDisplacement.Text = ""
            txtDisplacement.Focus()
            MsgBox("Unexpected error occured converting the displacement. Please try again.")
        End Try

        Return blnIsValid

    End Function
    Private Function CheckIfAccelerationValid() As Boolean
        'Check if the data that the user entered for the acceleration can be saved as double and is over 0.
        Dim blnIsValid As Boolean = False
        'See if data entered is valid using try-catch statements.
        Try
            dblAcceleration = Convert.ToDouble(txtAcceleration.Text)
            If dblAcceleration > 0 Then
                blnIsValid = True
            Else
                MsgBox("Please enter a number greater than 0 for the acceleration.")
            End If
        Catch exception As FormatException
            txtAcceleration.Text = ""
            txtAcceleration.Focus()
            MsgBox("Please enter a number for the acceleration.")
        Catch exception As OverflowException
            txtAcceleration.Text = ""
            txtAcceleration.Focus()
            MsgBox("Accerleration entered is too high.")
        Catch exception As SystemException
            txtAcceleration.Text = ""
            txtAcceleration.Focus()
            MsgBox("Unexpected error occured converting the acceleration. Please try again.")
        End Try

        Return blnIsValid

    End Function
    Private Function CheckIfTimeValid() As Boolean
        'Check if the data that the user entered for the time can be saved as double and is over 0.
        Dim blnIsValid As Boolean = False
        'See if data entered is valid using try-catch statements.
        Try
            dblTime = Convert.ToDouble(txtTime.Text)
            If dblTime > 0 Then
                blnIsValid = True
            Else
                MsgBox("Please enter a number greater than 0 for the time.")
            End If
        Catch exception As FormatException
            txtTime.Text = ""
            txtTime.Focus()
            MsgBox("Please enter a number for the time.")
        Catch exception As OverflowException
            txtTime.Text = ""
            txtTime.Focus()
            MsgBox("Time entered is too high.")
        Catch exception As SystemException
            txtTime.Text = ""
            txtTime.Focus()
            MsgBox("Unexpected error occured converting the time. Please try again.")
        End Try

        Return blnIsValid

    End Function



    Private Sub btnCalculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalculate.Click

        'Declare variable containing result of calculation.
        Dim dblResult As Double

        If radDisplacement.Checked Then
            If CheckIfAccelerationValid() AndAlso CheckIfTimeValid() Then
                'Calculates Displacement using the formula d = 0.5*acceleration*time^2
                txtDisplacement.Enabled = False
                'Save user input:
                dblAcceleration = Convert.ToDouble(txtAcceleration.Text)
                dblTime = Convert.ToDouble(txtTime.Text)
                'Processing:
                dblResult = 0.5 * dblAcceleration * dblTime ^ 2

                'Result output:
                lblResult.Text = "The displacement is:"
                lblResult2.Text = dblResult.ToString("N5") & " m"
                lblResult.Visible = True
                lblResult2.Visible = True
            End If

        ElseIf radAcceleration.Checked Then
            If CheckIfDisplacementValid() AndAlso CheckIfTimeValid() Then
                'Calculates acceleration using the formula a = (2*Displacement/Time^2)
                txtAcceleration.Enabled = False
                'Save user input:
                dblDisplacement = Convert.ToDouble(txtDisplacement.Text)
                dblTime = Convert.ToDouble(txtTime.Text)
                'processing:
                dblResult = (2 * dblDisplacement / (dblTime ^ 2))

                'Results:
                lblResult.Text = "The acceleration is:"
                lblResult2.Text = dblResult.ToString("N5") & " m/s"
                lblResult.Visible = True
                lblResult2.Visible = True

            End If
        ElseIf radTime.Checked Then
            If CheckIfDisplacementValid() AndAlso CheckIfAccelerationValid() Then
                'Calculates time using the formula Square root of (Displacement/0.5*Acceleration)
                txtTime.Enabled = False
                'Save user input
                dblDisplacement = Convert.ToDouble(txtDisplacement.Text)
                dblAcceleration = Convert.ToDouble(txtAcceleration.Text)
                'Processing
                Dim dblQuotient As Double
                dblQuotient = (dblDisplacement / (0.5 * dblAcceleration))
                dblResult = dblQuotient ^ (1 / 2)
                'results
                lblResult.Text = "The time is:"
                lblResult2.Text = dblResult.ToString("N5") & " s"
                lblResult.Visible = True
                lblResult2.Visible = True
            End If
     
        End If


    End Sub
    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        'Clears the program of values. Sets rbtnAcceleration as focus.

        txtAcceleration.Text = ""
        txtDisplacement.Text = ""
        txtTime.Text = ""
        lblResult.Text = ""
        lblResult2.Text = ""
        lblResult.Visible = False
        lblResult2.Visible = False
        radDisplacement.Focus()
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        'Closes the program.
        Close()
    End Sub

End Class
