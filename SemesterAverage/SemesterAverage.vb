''Author: Trevor Andrus
''Date: 2018-02-17
''Name: SemesterAverage.vb
''Purpose: This form allows the user to enter 6 grades and shows them the letter grades for each as well as the semester average

Option Strict On
Public Class frmSemesterAverage

    'Creates variables
    Private Const NUM_COURSES As Integer = 6
    Private inputValidated As Double
    Private letterGrade As String
    Dim GradesInput(6) As TextBox
    Dim Grades(6) As Double
    Private semesterTotal As Double
    Private mark1Valid, mark2Valid, mark3Valid, mark4Valid, mark5Valid, mark6Valid As Boolean


    ''These 6 Sub's for the 6 input textboxes determine if the input is valid.
    ''If it is we call the lette grade function and display that as well as add it into our array and set its valid boolean to true
    ''If it isn't we display and error message and set focus to the text box
    Private Sub tbMark1_LostFocus(sender As Object, e As EventArgs) Handles tbMark1.Leave
        If ValidateInput(tbMark1.Text) = True Then
            tbLetterGrade1.Text = CalculateLetterGrade(inputValidated)
            GradesInput(0) = tbMark1
            mark1Valid = True
        Else
            tbMessages.Text = tbMessages.Text + "ERROR: Please make sure what you input in Course 1 is a number from 0 - 100." + vbNewLine
            tbMark1.Focus()
        End If
        inputValidated = Nothing
    End Sub

    Private Sub tbMark2_LostFocus(sender As Object, e As EventArgs) Handles tbMark2.Leave
        If ValidateInput(tbMark2.Text) = True Then
            tbLetterGrade2.Text = CalculateLetterGrade(inputValidated)
            GradesInput(1) = tbMark2
            mark2Valid = True
        Else
            tbMessages.Text = tbMessages.Text + "ERROR: Please make sure what you input in Course 2 is a number from 0 - 100." + vbNewLine
            tbMark2.Focus()
        End If
        inputValidated = Nothing
    End Sub

    Private Sub tbMark3_LostFocus(sender As Object, e As EventArgs) Handles tbMark3.Leave
        If ValidateInput(tbMark3.Text) = True Then
            tbLetterGrade3.Text = CalculateLetterGrade(inputValidated)
            GradesInput(2) = tbMark3
            mark3Valid = True
        Else
            tbMessages.Text = tbMessages.Text + "ERROR: Please make sure what you input in Course 3 is a number from 0 - 100." + vbNewLine
            tbMark3.Focus()
        End If
        inputValidated = Nothing
    End Sub
    Private Sub tbMark4_LostFocus(sender As Object, e As EventArgs) Handles tbMark4.Leave
        If ValidateInput(tbMark4.Text) = True Then
            tbLetterGrade4.Text = CalculateLetterGrade(inputValidated)
            GradesInput(3) = tbMark4
            mark4Valid = True
        Else
            tbMessages.Text = tbMessages.Text + "ERROR: Please make sure what you input in Course 4 is a number from 0 - 100." + vbNewLine
            tbMark4.Focus()
        End If
        inputValidated = Nothing
    End Sub

    Private Sub tbMark5_LostFocus(sender As Object, e As EventArgs) Handles tbMark5.Leave
        If ValidateInput(tbMark5.Text) = True Then
            tbLetterGrade5.Text = CalculateLetterGrade(inputValidated)
            GradesInput(4) = tbMark5
            mark5Valid = True
        Else
            tbMessages.Text = tbMessages.Text + "ERROR: Please make sure what you input in Course 5 is a number from 0 - 100." + vbNewLine
            tbMark5.Focus()
        End If
        inputValidated = Nothing
    End Sub
    Private Sub tbMark6_LostFocus(sender As Object, e As EventArgs) Handles tbMark6.Leave
        If ValidateInput(tbMark6.Text) = True Then
            tbLetterGrade6.Text = CalculateLetterGrade(inputValidated)
            GradesInput(5) = tbMark6
            mark6Valid = True
        Else
            tbMessages.Text = tbMessages.Text + "ERROR: Please make sure what you input in Course 6 is a number from 0 - 100." + vbNewLine
            tbMark6.Focus()
        End If
        inputValidated = Nothing
    End Sub





    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Closes form
        Close()
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        ''Clears textboxes and labels as well as clearing all variables necessary for further operation
        tbMark1.Clear()
        tbMark2.Clear()
        tbMark3.Clear()
        tbMark4.Clear()
        tbMark5.Clear()
        tbMark6.Clear()
        tbLetterGrade1.Clear()
        tbLetterGrade2.Clear()
        tbLetterGrade3.Clear()
        tbLetterGrade4.Clear()
        tbLetterGrade5.Clear()
        tbLetterGrade6.Clear()
        tbSemester.Clear()
        tbSemesterGrade.Clear()
        tbMessages.Clear()
        inputValidated = Nothing
        mark1Valid = Nothing
        mark2Valid = Nothing
        mark3Valid = Nothing
        mark4Valid = Nothing
        mark5Valid = Nothing
        mark6Valid = Nothing
        semesterTotal = 0
        tbMark1.ReadOnly = False
        tbMark2.ReadOnly = False
        tbMark3.ReadOnly = False
        tbMark4.ReadOnly = False
        tbMark5.ReadOnly = False
        tbMark6.ReadOnly = False
        btnCalculate.Enabled = True
        For i As Integer = 0 To 5 Step 1
            Grades(i) = Nothing
            GradesInput(i) = Nothing
        Next

    End Sub

    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        If mark1Valid = True And mark2Valid = True And mark3Valid = True And mark4Valid = True And mark5Valid = True And mark6Valid = True Then
            ''Loops through the array converting to doubles and adding to a total to be used fro average calculation
            For i As Integer = 0 To 5 Step 1
                Grades(i) = Convert.ToDouble(GradesInput(i).Text)
                semesterTotal += Grades(i)
            Next
            ''Calculates the average and letter grade and displays them in the appropriate controls
            semesterTotal = semesterTotal / NUM_COURSES
            semesterTotal = Math.Round(semesterTotal, 2)
            tbSemester.Text = semesterTotal.ToString
            tbSemesterGrade.Text = CalculateLetterGrade(semesterTotal)
            ''Disables input text boxes and caluclate button until the user presses the reste button
            tbMark1.ReadOnly = True
            tbMark2.ReadOnly = True
            tbMark3.ReadOnly = True
            tbMark4.ReadOnly = True
            tbMark5.ReadOnly = True
            tbMark6.ReadOnly = True
            btnCalculate.Enabled = False
        End If
    End Sub

    Function ValidateInput(input As String) As Boolean
        ''Determines if input is a double and within the range of 0 - 100 if so returns true otherwise returns false
        If Double.TryParse(input, inputValidated) And inputValidated >= 0 And inputValidated <= 100 Then
            Return True
        Else
            Return False
        End If
    End Function

    Function CalculateLetterGrade(grade As Double) As String
        ''Finds which grade range the given value is and then returns it
        If grade >= 90 Then
            letterGrade = "A+"
        ElseIf grade >= 85 Then
            letterGrade = "A"
        ElseIf grade >= 80 Then
            letterGrade = "A-"
        ElseIf grade >= 77 Then
            letterGrade = "B+"
        ElseIf grade >= 73 Then
            letterGrade = "B"
        ElseIf grade >= 70 Then
            letterGrade = "B-"
        ElseIf grade >= 67 Then
            letterGrade = "C+"
        ElseIf grade >= 63 Then
            letterGrade = "C"
        ElseIf grade >= 60 Then
            letterGrade = "C-"
        ElseIf grade >= 57 Then
            letterGrade = "D+"
        ElseIf grade >= 53 Then
            letterGrade = "D"
        ElseIf grade >= 50 Then
            letterGrade = "D-"
        Else
            letterGrade = "F"
        End If
        Return letterGrade
    End Function
End Class
