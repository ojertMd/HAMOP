Imports System.Numerics
Public Class Form1
    Dim hamop As New Hamop
    Dim num1 As New Hamop.VLnum
    Dim Itst As Integer = 0   '0
    Dim reversed As Boolean = False
    Dim rnd As New Random()
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        hamop.MaximumLength = CInt(TxtBoxMaxLength.Text)
        hamop.DigitGroupLength = CInt(TxtBoxDigitGroupLength.Text)
        Me.Label8.Text = ""
        Me.lblResult.Text = ""
        Me.lblResultCPU.Text = ""
        Me.BtnePx.Enabled = False
        Me.Btn10pX.Enabled = False
        Me.BtnEXP.Enabled = False
        Me.BtnVLnumXinv.Enabled = False
        Me.BtnX2.Enabled = False
        Me.ChkboxTrimTzeros.Enabled = False
        Me.ChkBoxCorrectToMaxlen.Enabled = False
        Me.ChkBxAuto.Enabled = False
        Me.ChkBxStep.Enabled = False
        Me.btnTestResult.Enabled = False
        Me.RdoBtnAccept.Enabled = False
        Me.RdoBtnReject.Enabled = False
    End Sub

    Private Sub TxtBoxMaxLength_LostFocus(sender As Object, e As EventArgs) Handles TxtBoxMaxLength.LostFocus
        If TxtBoxMaxLength.Text = "" Or TxtBoxMaxLength.Text = Nothing Or IsNumeric(TxtBoxMaxLength.Text) = False Or
            TxtBoxMaxLength.Text(0) = "-" Then

            MsgBox("Proper Value required", vbOKOnly, "Enter a positive number for Maximum No. of Digits")
            TxtBoxMaxLength.Focus()
        Else
            If CInt(TxtBoxMaxLength.Text) > 3000 Then
                Label8.Text = "Wait Constants are being evaluated"
                Me.Label8.Update()
            End If
            hamop.MaximumLength = CInt(TxtBoxMaxLength.Text)
            If CInt(TxtBoxMaxLength.Text) > 3000 Then
                Label8.Text = "Constants are evaluated"
                Me.Label8.Update()
            End If

        End If
    End Sub
    Private Sub TxtBoxDigitGroupLength_LostFocus(sender As Object, e As EventArgs) Handles TxtBoxDigitGroupLength.LostFocus
        If TxtBoxDigitGroupLength.Text = "" Or TxtBoxDigitGroupLength.Text = Nothing Or IsNumeric(TxtBoxDigitGroupLength.Text) = False Or
            TxtBoxDigitGroupLength.Text(0) = "-" Then

            MsgBox("Proper Value required", vbOKOnly, "Enter a positive number for Digit Group length")
            TxtBoxDigitGroupLength.Focus()
        Else
            hamop.DigitGroupLength = CInt(TxtBoxDigitGroupLength.Text)
        End If
    End Sub

    Private Sub BtnPlus_Click(sender As Object, e As EventArgs) Handles BtnPlus.Click
        Dim Bn1, Bn2 As BigInteger
        Dim num1, num2 As New Hamop.VLnum
        lblResultComp.Text = "Compare Results"
        'Try        'try catch block should be used to catch exceptions thrown inside the class like div by zero or overflow etc
        num1.Value = TxtBoxNum1.Text
        num2.Value = TxtBoxNum2.Text
        TxtBoxResult.Text = (num1 + num2).Value
        'Catch ex As Exception

        'MsgBox(ex, vbOKOnly, "Error")
        'Exit Sub LINE IS NECESSARY TO STOP FURTHER EXECUTION OF CODE, ORELSE IT WILL EXECUTE TxtBoxResult.Text = num1.Value
        ' AND IT WILL THROW ANOTHER EXCEPTION IN VALUE.GET PROPERTY BECAUSE strH BECOMES ""
        'Exit Sub
        'End Try
        If TxtBoxMaxLength.Text < 29 Then
            Try
                If IsNumeric(TxtBoxNum1.Text) And IsNumeric(TxtBoxNum2.Text) Then   'IF NOT USED EXCEPTION WILL OCCUR WHEN TEXT BOX IS "" OR NOTHING
                    TxtBoxResultCPU.Text = CDec(TxtBoxNum1.Text) + CDec(TxtBoxNum2.Text)
                End If
            Catch ex As Exception
                If TxtBoxNum1.Text.IndexOf(".") > -1 Or TxtBoxNum2.Text.IndexOf(".") > -1 Then
                Else
                    If IsNumeric(TxtBoxNum1.Text) And IsNumeric(TxtBoxNum2.Text) Then   'IF NOT USED EXCEPTION WILL OCCUR WHEN TEXT BOX IS "" OR NOTHING
                        Bn1 = BigInteger.Parse(TxtBoxNum1.Text)
                        Bn2 = BigInteger.Parse(TxtBoxNum2.Text)
                        TxtBoxResultCPU.Text = (Bn1 + Bn2).ToString
                    End If
                End If
            End Try
        Else
            If TxtBoxNum1.Text.IndexOf(".") > -1 Or TxtBoxNum2.Text.IndexOf(".") > -1 Then
            Else
                If IsNumeric(TxtBoxNum1.Text) And IsNumeric(TxtBoxNum2.Text) Then   'IF NOT USED EXCEPTION WILL OCCUR WHEN TEXT BOX IS "" OR NOTHING
                    Bn1 = BigInteger.Parse(TxtBoxNum1.Text)
                    Bn2 = BigInteger.Parse(TxtBoxNum2.Text)
                    TxtBoxResultCPU.Text = (Bn1 + Bn2).ToString
                End If
            End If
        End If

        Call ResultsChecker()
    End Sub

    Private Sub BtnMinus_Click(sender As Object, e As EventArgs) Handles BtnMinus.Click
        Dim Bn1, Bn2 As BigInteger
        Dim num1, num2 As New Hamop.VLnum
        lblResultComp.Text = "Compare Results"

        num1.Value = TxtBoxNum1.Text
        num2.Value = TxtBoxNum2.Text
        TxtBoxResult.Text = (num1 - num2).Value
        If TxtBoxMaxLength.Text < 29 Then
            Try
                If IsNumeric(TxtBoxNum1.Text) And IsNumeric(TxtBoxNum2.Text) Then   'IF NOT USED EXCEPTION WILL OCCUR WHEN TEXT BOX IS "" OR NOTHING
                    TxtBoxResultCPU.Text = CDec(TxtBoxNum1.Text) - CDec(TxtBoxNum2.Text)
                End If
            Catch ex As Exception
                If TxtBoxNum1.Text.IndexOf(".") > -1 Or TxtBoxNum2.Text.IndexOf(".") > -1 Then
                Else
                    If IsNumeric(TxtBoxNum1.Text) And IsNumeric(TxtBoxNum2.Text) Then   'IF NOT USED EXCEPTION WILL OCCUR WHEN TEXT BOX IS "" OR NOTHING
                        Bn1 = BigInteger.Parse(TxtBoxNum1.Text)
                        Bn2 = BigInteger.Parse(TxtBoxNum2.Text)
                        TxtBoxResultCPU.Text = (Bn1 - Bn2).ToString
                    End If
                End If
            End Try
        Else
            If TxtBoxNum1.Text.IndexOf(".") > -1 Or TxtBoxNum2.Text.IndexOf(".") > -1 Then
            Else
                If IsNumeric(TxtBoxNum1.Text) And IsNumeric(TxtBoxNum2.Text) Then   'IF NOT USED EXCEPTION WILL OCCUR WHEN TEXT BOX IS "" OR NOTHING
                    Bn1 = BigInteger.Parse(TxtBoxNum1.Text)
                    Bn2 = BigInteger.Parse(TxtBoxNum2.Text)
                    TxtBoxResultCPU.Text = (Bn1 - Bn2).ToString
                End If
            End If
        End If
        Call ResultsChecker()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim VLNt As New VLNumTests
        Dim finish As Boolean = False
        LblRegex.Text = ""
        If TxtBoxTestNo.Text = 1 Then
            VLNt.testVLNum(Itst, finish)
        ElseIf TxtBoxTestNo.Text = 2 Then
            VLNt.testVLNum1(Itst, finish)
        ElseIf TxtBoxTestNo.Text = 3 Then
            VLNt.IntGen(Itst, finish, hamop.MaximumLength)
        ElseIf TxtBoxTestNo.Text = 4 Then
            VLNt.DblGen(Itst, finish, hamop.MaximumLength)
        ElseIf TxtBoxTestNo.Text = 5 Then
            VLNt.IntDblGen(Itst, finish, hamop.MaximumLength)
        ElseIf TxtBoxTestNo.Text = 6 Then
            VLNt.operatorTest(Itst, finish)
        ElseIf TxtBoxTestNo.Text = 7 Then
            VLNt.TESTvalsOPR(Itst, finish, hamop.MaximumLength)
        End If

        Itst = Itst + 1
        If finish = True Then
            MessageBox.Show("Tests completed successfully")
            Itst = 0
        End If
    End Sub
    Private Sub BtnInto_Click(sender As Object, e As EventArgs) Handles BtnInto.Click
        Dim Bn1, Bn2 As BigInteger
        Dim num1, num2 As New Hamop.VLnum
        lblResultComp.Text = "Compare Results"

        num1.Value = TxtBoxNum1.Text
        num2.Value = TxtBoxNum2.Text
        TxtBoxResult.Text = (num1 * num2).Value
        If TxtBoxMaxLength.Text < 29 Then
            Try
                If IsNumeric(TxtBoxNum1.Text) And IsNumeric(TxtBoxNum2.Text) Then   'IF NOT USED EXCEPTION WILL OCCUR WHEN TEXT BOX IS "" OR NOTHING
                    TxtBoxResultCPU.Text = CDec(TxtBoxNum1.Text) * CDec(TxtBoxNum2.Text)
                End If
            Catch ex As Exception
                If TxtBoxNum1.Text.IndexOf(".") > -1 Or TxtBoxNum2.Text.IndexOf(".") > -1 Then
                Else
                    If IsNumeric(TxtBoxNum1.Text) And IsNumeric(TxtBoxNum2.Text) Then   'IF NOT USED EXCEPTION WILL OCCUR WHEN TEXT BOX IS "" OR NOTHING
                        Bn1 = BigInteger.Parse(TxtBoxNum1.Text)
                        Bn2 = BigInteger.Parse(TxtBoxNum2.Text)
                        TxtBoxResultCPU.Text = (Bn1 * Bn2).ToString
                    End If
                End If
            End Try
        Else
            If TxtBoxNum1.Text.IndexOf(".") > -1 Or TxtBoxNum2.Text.IndexOf(".") > -1 Then
            Else
                If IsNumeric(TxtBoxNum1.Text) And IsNumeric(TxtBoxNum2.Text) Then   'IF NOT USED EXCEPTION WILL OCCUR WHEN TEXT BOX IS "" OR NOTHING
                    Bn1 = BigInteger.Parse(TxtBoxNum1.Text)
                    Bn2 = BigInteger.Parse(TxtBoxNum2.Text)
                    TxtBoxResultCPU.Text = (Bn1 * Bn2).ToString
                End If
            End If
        End If
        Call ResultsChecker()

    End Sub
    Private Sub BtnDivide_Click(sender As Object, e As EventArgs) Handles BtnDivide.Click
        Dim Bn1, Bn2 As BigInteger
        Dim num1, num2 As New Hamop.VLnum
        lblResultComp.Text = "Compare Results"

        num1.Value = TxtBoxNum1.Text
        num2.Value = TxtBoxNum2.Text
        TxtBoxResult.Text = (num1 / num2).Value
        If TxtBoxMaxLength.Text < 29 Then
            Try
                If IsNumeric(TxtBoxNum1.Text) And IsNumeric(TxtBoxNum2.Text) Then   'IF NOT USED EXCEPTION WILL OCCUR WHEN TEXT BOX IS "" OR NOTHING
                    TxtBoxResultCPU.Text = CDec(TxtBoxNum1.Text) / CDec(TxtBoxNum2.Text)
                End If
            Catch ex As Exception
                If TxtBoxNum1.Text.IndexOf(".") > -1 Or TxtBoxNum2.Text.IndexOf(".") > -1 Then
                Else
                    If IsNumeric(TxtBoxNum1.Text) And IsNumeric(TxtBoxNum2.Text) Then   'IF NOT USED EXCEPTION WILL OCCUR WHEN TEXT BOX IS "" OR NOTHING
                        Bn1 = BigInteger.Parse(TxtBoxNum1.Text)
                        Bn2 = BigInteger.Parse(TxtBoxNum2.Text)
                        TxtBoxResultCPU.Text = (Bn1 / Bn2).ToString
                    End If
                End If
            End Try
        Else
            If TxtBoxNum1.Text.IndexOf(".") > -1 Or TxtBoxNum2.Text.IndexOf(".") > -1 Then
            Else
                If IsNumeric(TxtBoxNum1.Text) And IsNumeric(TxtBoxNum2.Text) Then   'IF NOT USED EXCEPTION WILL OCCUR WHEN TEXT BOX IS "" OR NOTHING
                    Bn1 = BigInteger.Parse(TxtBoxNum1.Text)
                    Bn2 = BigInteger.Parse(TxtBoxNum2.Text)
                    TxtBoxResultCPU.Text = (Bn1 / Bn2).ToString
                End If
            End If
        End If

        Call ResultsChecker()

    End Sub

    Private Sub TxtBoxNum1_TextChanged(sender As Object, e As EventArgs) Handles TxtBoxNum1.TextChanged
        Dim ps As Integer
        If TxtBoxNum1.Text = "" Or TxtBoxNum1.Text = Nothing Then
            TxtBoxNum1Len.Text = 0
        Else

            If TxtBoxNum1.Text(0) = "+" Or TxtBoxNum1.Text(0) = "-" Then    'THIS SECTION WILL PREVENT ENTERING MORE DIGITS THAN -MaxLength
                ps += 1
            End If
            If TxtBoxNum1.Text.IndexOf(".") > -1 Then
                ps += 1
            End If
            If TxtBoxNum1.Text.Length > hamop.MaximumLength + ps Then
                TxtBoxNum1.Text = TxtBoxNum1.Text.Substring(0, hamop.MaximumLength + ps)
                MsgBox("You entered " & hamop.MaximumLength & " digits to enter more increase Maximum No. of Digits", vbOKOnly, "Warning")
            End If
            TxtBoxNum1Len.Text = TxtBoxNum1.Text.Length - ps
        End If
    End Sub

    Private Sub TxtBoxNum2_TextChanged(sender As Object, e As EventArgs) Handles TxtBoxNum2.TextChanged
        Dim ps As Integer
        If TxtBoxNum2.Text = "" Or TxtBoxNum2.Text = Nothing Then
            TxtBoxNum2Len.Text = 0
        Else

            If TxtBoxNum2.Text(0) = "+" Or TxtBoxNum2.Text(0) = "-" Then    'THIS SECTION WILL PREVENT ENTERING MORE DIGITS THAN -MaxLength
                ps += 1
            End If
            If TxtBoxNum2.Text.IndexOf(".") > -1 Then
                ps += 1
            End If
            If TxtBoxNum2.Text.Length > hamop.MaximumLength + ps Then
                TxtBoxNum2.Text = TxtBoxNum2.Text.Substring(0, hamop.MaximumLength + ps)
                MsgBox("You entered " & hamop.MaximumLength & " digits to enter more increase Maximum No. of Digits", vbOKOnly, "Warning")
            End If
            TxtBoxNum2Len.Text = TxtBoxNum2.Text.Length - ps
        End If
    End Sub

    Private Sub TxtBoxResult_TextChanged(sender As Object, e As EventArgs) Handles TxtBoxResult.TextChanged
        Dim ps As Integer
        If TxtBoxResult.Text = "" Or TxtBoxResult.Text = Nothing Then
            TxtBoxResultLen.Text = 0
        Else

            If TxtBoxResult.Text(0) = "+" Or TxtBoxResult.Text(0) = "-" Then    'THIS WILL GIVE YOU NO. OF DIGITS ONLY . & + OR - WILL BE EXCLUDED
                ps += 1
            End If
            If TxtBoxResult.Text.IndexOf(".") > -1 Then
                ps += 1
            End If
            TxtBoxResultLen.Text = TxtBoxResult.Text.Length - ps
        End If
    End Sub

    Private Sub TxtBoxResultCPU_TextChanged(sender As Object, e As EventArgs) Handles TxtBoxResultCPU.TextChanged
        Dim ps As Integer
        If TxtBoxResultCPU.Text = "" Or TxtBoxResultCPU.Text = Nothing Then
            TxtBoxResultCPULen.Text = 0
        Else

            If TxtBoxResultCPU.Text(0) = "+" Or TxtBoxResultCPU.Text(0) = "-" Then    'THIS WILL GIVE YOU NO. OF DIGITS ONLY . & + OR - WILL BE EXCLUDED
                ps += 1
            End If
            If TxtBoxResultCPU.Text.IndexOf(".") > -1 Then
                ps += 1
            End If
            TxtBoxResultCPULen.Text = TxtBoxResultCPU.Text.Length - ps
        End If
    End Sub

    Private Sub BtnSecondaryOperators_Click(sender As Object, e As EventArgs) Handles BtnSecondaryOperators.Click
        Call Equalto()      'uncomment and use as required  (call only one sub at a time)
        'Call NE()
        'Call GT()
        'Call LT()
        'Call GTeq()
        'Call LTeq()
    End Sub
    Private Sub Equalto()
        Dim NUM1 As New HAMOP.VLNum
        Dim NUM2 As New HAMOP.VLNum
        NUM1.Value = TxtBoxNum1.Text
        NUM2.Value = TxtBoxNum2.Text
        If NUM1 = NUM2 Then
            TxtBoxResult.Text = "Num1 = Num2 is TRUE"
        Else
            TxtBoxResult.Text = "Num1 = Num2 is FALSE"
        End If
        If Val(TxtBoxNum1.Text) = Val(TxtBoxNum2.Text) Then
            TxtBoxResultCPU.Text = "Num1 = Num2 is TRUE"
        Else
            TxtBoxResultCPU.Text = "Num1 = Num2 is FALSE"
        End If
    End Sub
    Private Sub NE()
        Dim NUM1 As New HAMOP.VLNum
        Dim NUM2 As New HAMOP.VLNum
        NUM1.Value = TxtBoxNum1.Text
        NUM2.Value = TxtBoxNum2.Text
        If NUM1 <> NUM2 Then
            TxtBoxResult.Text = "Num1 <> Num2 is TRUE"
        Else
            TxtBoxResult.Text = "Num1 <> Num2 is FALSE"
        End If
        If Val(TxtBoxNum1.Text) <> Val(TxtBoxNum2.Text) Then
            TxtBoxResultCPU.Text = "Num1 <> Num2 is TRUE"
        Else
            TxtBoxResultCPU.Text = "Num1 <> Num2 is FALSE"
        End If
    End Sub
    Private Sub GT()
        Dim NUM1 As New HAMOP.VLNum
        Dim NUM2 As New HAMOP.VLNum
        NUM1.Value = TxtBoxNum1.Text
        NUM2.Value = TxtBoxNum2.Text
        If NUM1 > NUM2 Then
            TxtBoxResult.Text = "Num1 > Num2 is TRUE"
        Else
            TxtBoxResult.Text = "Num1 > Num2 is FALSE"
        End If
        If Val(TxtBoxNum1.Text) > Val(TxtBoxNum2.Text) Then
            TxtBoxResultCPU.Text = "Num1 > Num2 is TRUE"
        Else
            TxtBoxResultCPU.Text = "Num1 > Num2 is FALSE"
        End If
    End Sub
    Private Sub LT()
        Dim NUM1 As New HAMOP.VLNum
        Dim NUM2 As New HAMOP.VLNum
        NUM1.Value = TxtBoxNum1.Text
        NUM2.Value = TxtBoxNum2.Text
        If NUM1 < NUM2 Then
            TxtBoxResult.Text = "Num1 < Num2 is TRUE"
        Else
            TxtBoxResult.Text = "Num1 < Num2 is FALSE"
        End If
        If Val(TxtBoxNum1.Text) < Val(TxtBoxNum2.Text) Then
            TxtBoxResultCPU.Text = "Num1 < Num2 is TRUE"
        Else
            TxtBoxResultCPU.Text = "Num1 < Num2 is FALSE"
        End If
    End Sub
    Private Sub GTeq()
        Dim NUM1 As New HAMOP.VLNum
        Dim NUM2 As New HAMOP.VLNum
        NUM1.Value = TxtBoxNum1.Text
        NUM2.Value = TxtBoxNum2.Text
        If NUM1 >= NUM2 Then
            TxtBoxResult.Text = "Num1 >= Num2 is TRUE"
        Else
            TxtBoxResult.Text = "Num1 >= Num2 is FALSE"
        End If
        If Val(TxtBoxNum1.Text) >= Val(TxtBoxNum2.Text) Then
            TxtBoxResultCPU.Text = "Num1 >= Num2 is TRUE"
        Else
            TxtBoxResultCPU.Text = "Num1 >= Num2 is FALSE"
        End If
    End Sub
    Private Sub LTeq()
        Dim NUM1 As New HAMOP.VLNum
        Dim NUM2 As New HAMOP.VLNum
        NUM1.Value = TxtBoxNum1.Text
        NUM2.Value = TxtBoxNum2.Text
        If NUM1 <= NUM2 Then
            TxtBoxResult.Text = "Num1 <= Num2 is TRUE"
        Else
            TxtBoxResult.Text = "Num1 <= Num2 is FALSE"
        End If
        If Val(TxtBoxNum1.Text) <= Val(TxtBoxNum2.Text) Then
            TxtBoxResultCPU.Text = "Num1 <= Num2 is TRUE"
        Else
            TxtBoxResultCPU.Text = "Num1 <= Num2 is FALSE"
        End If
    End Sub

    Private Sub Btne_Click(sender As Object, e As EventArgs) Handles Btne.Click
        Dim eVal, Num1, sum As New Hamop.VLnum
        Dim StartTime, EndTime As New DateTime
        Dim RunTime As New TimeSpan
        StartTime = DateTime.Now

        eVal = hamop.VLnMath.CALCe()
        TxtBoxResult.Text = eVal.Value
        TxtBoxResultCPU.Text = Math.E
        'Call ResultsChecker(1)

        EndTime = DateTime.Now
        RunTime = EndTime - StartTime
        TxtBoxNum2.Text = "Run Time " & RunTime.TotalMilliseconds & " Milliseconds  ---- " & RunTime.TotalSeconds & " Seconds"
    End Sub

    Private Sub BtnALoge_Click(sender As Object, e As EventArgs) Handles BtnALoge.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double
        N1 = TxtBoxNum1.Text
        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.ALoge(Num1)
        TxtBoxResult.Text = sum.Value
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Exp(N1)
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28)
            TxtBoxResultCPU.Text = Math.Exp(N1)
        End If
        Call ResultsChecker(1)
    End Sub
    Private Sub BtnLoge_Click(sender As Object, e As EventArgs) Handles BtnLoge.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double
        If TxtBoxNum1.Text(0) = "-" Then TxtBoxNum1.Text = TxtBoxNum1.Text.Substring(1)
        N1 = TxtBoxNum1.Text
        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.Loge(Num1)
        TxtBoxResult.Text = sum.Value
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Log(N1)
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28)
            TxtBoxResultCPU.Text = Math.Log(N1)
        End If

        Call ResultsChecker(1)
    End Sub
    Private Sub ResultsChecker(Optional ByVal decrim As Integer = 0)
        Dim Eindex, srCPUlen, srLen, sLen, sLenMax As Integer
        Dim s, s2, srCPU As String
        Dim Numl2 As New Hamop.VLnum
        Dim sr As Decimal
        lblResultCPU.Text = ""
        lblResult.Text = ""
        If reversed = True Then
            reversed = False
        Else
            Label8.Text = ""
        End If
        Eindex = TxtBoxResultCPU.Text.IndexOf("E")
        srLen = TxtBoxResult.Text.Length
        srCPUlen = TxtBoxResultCPU.Text.Length
        srCPU = TxtBoxResultCPU.Text
        If Eindex > -1 Then
            sr = TxtBoxResultCPU.Text
            srCPUlen = sr.ToString.Length
            TxtBoxResultCPU.Text = sr
        End If
        If srCPUlen > srLen Then
            sLen = srLen
            sLenMax = srCPUlen
        Else
            sLen = srCPUlen
            sLenMax = srLen
        End If
        For I As Integer = 0 To sLen - 1
            If TxtBoxResultCPU.Text(I) = TxtBoxResult.Text(I) Then
                If I = (sLen - 1) Then
                    If sLen = sLenMax Then
                        lblResultComp.Text = "Result  and  CPU Result are Equal"
                    Else
                        lblResultComp.Text = I + 1 & " Digits of Result  and  CPU Result are Equal  Check for ROUNDING and expanxion of 'E'  !!!!!!!!!"
                        s = TxtBoxResultCPU.Text.Substring(0, I + 1)
                        s2 = TxtBoxResultCPU.Text.Substring(I + 1)
                        If Eindex > -1 Then
                            lblResultCPU.Text = s & "  " & s2 & "      " & srCPU
                        Else
                            lblResultCPU.Text = s & "  " & s2
                        End If

                        s = TxtBoxResult.Text.Substring(0, I + 1)
                        s2 = TxtBoxResult.Text.Substring(I + 1)
                        If Eindex > -1 Then
                            lblResult.Text = s & "  " & s2
                        Else
                            lblResult.Text = s & "  " & s2
                        End If
                        TxtBoxResultCPU.Text = srCPU
                    End If
                End If
            Else
                If I = 0 Then
                    lblResultComp.Text = "Result  and  CPU Result are Not  Equal     ??????????"
                Else
                    lblResultComp.Text = I & " Digits of Result  and  CPU Result are Equal  Check for ROUNDING and expanxion of 'E'    !!!!!!!!!"
                    s = TxtBoxResultCPU.Text.Substring(0, I)
                    s2 = TxtBoxResultCPU.Text.Substring(I)
                    If Eindex > -1 Then
                        lblResultCPU.Text = s & "  " & s2 & "      " & srCPU
                    Else
                        lblResultCPU.Text = s & "  " & s2
                    End If

                    s = TxtBoxResult.Text.Substring(0, I)
                    s2 = TxtBoxResult.Text.Substring(I)
                    lblResult.Text = s & "  " & s2
                    TxtBoxResultCPU.Text = srCPU
                End If
                Exit For
            End If
        Next
    End Sub

    Private Sub BtnREVERSE_Click(sender As Object, e As EventArgs) Handles BtnREVERSE.Click
        Label8.Text = TxtBoxNum1.Text
        TxtBoxNum1.Text = TxtBoxResult.Text
        reversed = True
    End Sub

    Private Sub BtnALog10_Click(sender As Object, e As EventArgs) Handles BtnALog10.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double
        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.ALog10(Num1)
        TxtBoxResult.Text = sum.Value
        N1 = TxtBoxNum1.Text
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Pow(10, N1)
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28)
            TxtBoxResultCPU.Text = Math.Pow(10, N1)
        End If
        Call ResultsChecker(1)
    End Sub
    Private Sub BtnLog10_Click(sender As Object, e As EventArgs) Handles BtnLog10.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double

        If TxtBoxNum1.Text(0) = "-" Then TxtBoxNum1.Text = TxtBoxNum1.Text.Substring(1)
        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.Log10(Num1)
        TxtBoxResult.Text = sum.Value
        N1 = TxtBoxNum1.Text
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Log10(N1)
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28)
            TxtBoxResultCPU.Text = Math.Log10(N1)
        End If
        Call ResultsChecker(1)
    End Sub
    Private Sub BtnxPy_Click(sender As Object, e As EventArgs) Handles BtnxPy.Click
        Dim sum, Num1, Num2 As New Hamop.VLnum
        Dim N1, N2 As Double
        Num1.Value = TxtBoxNum1.Text
        Num2.Value = TxtBoxNum2.Text
        sum = Num1 ^ Num2
        TxtBoxResult.Text = sum.Value
        N1 = TxtBoxNum1.Text
        N2 = TxtBoxNum2.Text
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Pow(N1, N2)
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28)
            TxtBoxResultCPU.Text = Math.Pow(N1, N2)
        End If
        Call ResultsChecker(1)
    End Sub
    Private Sub BtnSQRT_Click(sender As Object, e As EventArgs) Handles BtnSQRT.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double
        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.Sqrt(Num1)
        TxtBoxResult.Text = sum.Value
        N1 = TxtBoxNum1.Text
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Sqrt(N1)
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28)
            TxtBoxResultCPU.Text = Math.Sqrt(N1)
        End If
        Call ResultsChecker(1)
    End Sub

    Private Sub Btn_SinR_Click(sender As Object, e As EventArgs) Handles Btn_SinR.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double
        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.SinR(Num1)
        TxtBoxResult.Text = sum.Value
        N1 = TxtBoxNum1.Text
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Sin(N1)
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28)
            TxtBoxResultCPU.Text = Math.Sin(N1)
        End If
        Call ResultsChecker(1)
    End Sub
    Private Sub Btn_CosR_Click(sender As Object, e As EventArgs) Handles Btn_CosR.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double

        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.CosR(Num1)
        TxtBoxResult.Text = sum.Value
        N1 = TxtBoxNum1.Text
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Cos(N1)
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28)
            TxtBoxResultCPU.Text = Math.Cos(N1)
        End If
        Call ResultsChecker(1)
    End Sub
    Private Sub Btn_TanR_Click(sender As Object, e As EventArgs) Handles Btn_TanR.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double

        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.TanR(Num1)
        TxtBoxResult.Text = sum.Value
        N1 = TxtBoxNum1.Text
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Tan(N1)
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28)
            TxtBoxResultCPU.Text = Math.Tan(N1)
        End If
        Call ResultsChecker(1)
    End Sub
    Private Sub Btn_ASinR_Click(sender As Object, e As EventArgs) Handles Btn_ASinR.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double

        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.ASinR(Num1)
        TxtBoxResult.Text = sum.Value
        N1 = TxtBoxNum1.Text
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Asin(N1)
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28)
            TxtBoxResultCPU.Text = Math.Asin(N1)
        End If
        Call ResultsChecker(1)
    End Sub
    Private Sub Btn_ACosR_Click(sender As Object, e As EventArgs) Handles Btn_ACosR.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double

        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.ACosR(Num1)
        TxtBoxResult.Text = sum.Value
        N1 = TxtBoxNum1.Text
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Acos(N1)
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28)
            TxtBoxResultCPU.Text = Math.Acos(N1)
        End If
        Call ResultsChecker(1)
    End Sub
    Private Sub Btn_ATanR_Click(sender As Object, e As EventArgs) Handles Btn_ATanR.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double

        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.ATanR(Num1)
        TxtBoxResult.Text = sum.Value
        N1 = TxtBoxNum1.Text
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Atan(N1)
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28)
            TxtBoxResultCPU.Text = Math.Atan(N1)
        End If
        Call ResultsChecker(1)
    End Sub

    Private Sub Btn_SinD_Click(sender As Object, e As EventArgs) Handles Btn_SinD.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double

        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.SinD(Num1)
        TxtBoxResult.Text = sum.Value
        N1 = TxtBoxNum1.Text * Math.PI / 180
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Sin(N1)
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28) * Math.PI / 180
            TxtBoxResultCPU.Text = Math.Sin(N1)
        End If
        Call ResultsChecker(1)
    End Sub
    Private Sub Btn_CosD_Click(sender As Object, e As EventArgs) Handles Btn_CosD.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double

        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.CosD(Num1)
        TxtBoxResult.Text = sum.Value
        N1 = TxtBoxNum1.Text * Math.PI / 180
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Cos(N1)
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28) * Math.PI / 180
            TxtBoxResultCPU.Text = Math.Cos(N1)
        End If
        Call ResultsChecker(1)
    End Sub
    Private Sub Btn_TanD_Click(sender As Object, e As EventArgs) Handles Btn_TanD.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double

        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.TanD(Num1)
        TxtBoxResult.Text = sum.Value
        N1 = TxtBoxNum1.Text * Math.PI / 180
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Tan(N1)
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28) * Math.PI / 180
            TxtBoxResultCPU.Text = Math.Tan(N1)
        End If
        Call ResultsChecker(1)
    End Sub

    Private Sub Btn_ASinD_Click(sender As Object, e As EventArgs) Handles Btn_ASinD.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double

        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.ASinD(Num1)
        TxtBoxResult.Text = sum.Value
        N1 = TxtBoxNum1.Text
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Asin(N1) * 180 / Math.PI
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28)
            TxtBoxResultCPU.Text = Math.Asin(N1) * 180 / Math.PI
        End If
        Call ResultsChecker(1)
    End Sub
    Private Sub Btn_ACosD_Click(sender As Object, e As EventArgs) Handles Btn_ACosD.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double

        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.ACosD(Num1)
        TxtBoxResult.Text = sum.Value
        N1 = TxtBoxNum1.Text
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Acos(N1) * 180 / Math.PI
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28)
            TxtBoxResultCPU.Text = Math.Acos(N1) * 180 / Math.PI
        End If
        Call ResultsChecker(1)
    End Sub
    Private Sub Btn_ATanD_Click(sender As Object, e As EventArgs) Handles Btn_ATanD.Click
        Dim sum, Num1 As New Hamop.VLnum
        Dim N1 As Double

        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.ATanD(Num1)
        TxtBoxResult.Text = sum.Value
        N1 = TxtBoxNum1.Text
        If TxtBoxNum1.Text.Length < 29 Then
            TxtBoxResultCPU.Text = Math.Atan(N1) * 180 / Math.PI
        ElseIf (TxtBoxNum1.Text.IndexOf(".")) < 28 Then
            N1 = TxtBoxNum1.Text.Substring(0, 28)
            TxtBoxResultCPU.Text = Math.Atan(N1) * 180 / Math.PI
        End If
        Call ResultsChecker(1)
    End Sub

    Private Sub BtnAnsToN1_Click(sender As Object, e As EventArgs) Handles BtnAnsToN1.Click
        Call FloorH()
        'Call ceilingH()
    End Sub
    Private Sub FloorH()
        Dim sum, Num1 As New Hamop.VLnum
        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.Floor(Num1)
        TxtBoxResult.Text = sum.Value
        TxtBoxResultCPU.Text = Math.Floor(CDbl(TxtBoxNum1.Text))
    End Sub
    Private Sub ceilingH()
        Dim sum, Num1 As New Hamop.VLnum
        Num1.Value = TxtBoxNum1.Text
        sum = hamop.VLnMath.Ceiling(Num1)
        TxtBoxResult.Text = sum.Value
        TxtBoxResultCPU.Text = Math.Ceiling(CDbl(TxtBoxNum1.Text))
    End Sub

    Private Sub BtnAnsToN2_Click(sender As Object, e As EventArgs) Handles BtnAnsToN2.Click
        'THIS IS FOR MODULUS OPERATOR 
        Dim l1, l2 As Double
        Dim Bn1, Bn2 As BigInteger
        Dim num1, num2 As New Hamop.VLnum
        lblResultComp.Text = "Compare Results"
        l1 = TxtBoxNum1.Text
        l2 = TxtBoxNum2.Text


        num1.Value = TxtBoxNum1.Text
        num2.Value = TxtBoxNum2.Text
        TxtBoxResult.Text = (num1 Mod num2).Value
        If TxtBoxMaxLength.Text < 29 Then
            Try
                If IsNumeric(TxtBoxNum1.Text) And IsNumeric(TxtBoxNum2.Text) Then   'IF NOT USED EXCEPTION WILL OCCUR WHEN TEXT BOX IS "" OR NOTHING
                    TxtBoxResultCPU.Text = (l1 Mod l2).ToString
                End If
            Catch ex As Exception
                If TxtBoxNum1.Text.IndexOf(".") > -1 Or TxtBoxNum2.Text.IndexOf(".") > -1 Then
                Else
                    If IsNumeric(TxtBoxNum1.Text) And IsNumeric(TxtBoxNum2.Text) Then   'IF NOT USED EXCEPTION WILL OCCUR WHEN TEXT BOX IS "" OR NOTHING
                        Bn1 = BigInteger.Parse(TxtBoxNum1.Text)
                        Bn2 = BigInteger.Parse(TxtBoxNum2.Text)
                        TxtBoxResultCPU.Text = (Bn1 Mod Bn2).ToString
                    End If
                End If
            End Try
        Else
            If TxtBoxNum1.Text.IndexOf(".") > -1 Or TxtBoxNum2.Text.IndexOf(".") > -1 Then
            Else
                If IsNumeric(TxtBoxNum1.Text) And IsNumeric(TxtBoxNum2.Text) Then   'IF NOT USED EXCEPTION WILL OCCUR WHEN TEXT BOX IS "" OR NOTHING
                    Bn1 = BigInteger.Parse(TxtBoxNum1.Text)
                    Bn2 = BigInteger.Parse(TxtBoxNum2.Text)
                    TxtBoxResultCPU.Text = (Bn1 Mod Bn2).ToString
                End If
            End If
        End If
        Call ResultsChecker()
    End Sub

    Private Sub BtnX3_Click(sender As Object, e As EventArgs) Handles BtnX3.Click
        Dim num1, num2, sm As New Hamop.VLnum       '(10/3)*3 this sub is to check for rounding (ie to see whether it will be 10 or 9.999999...)
        lblResultComp.Text = "Compare Results"

        num1.Value = TxtBoxNum1.Text
        num2.Value = TxtBoxNum2.Text
        sm = (num1 / num2)
        TxtBoxResult.Text = (sm * num2).Value
        TxtBoxResultCPU.Text = ""
        Call ResultsChecker()
    End Sub

    Private Sub BtnPI_Click(sender As Object, e As EventArgs) Handles BtnPI.Click
        Dim num1, num2, num4, PI As New Hamop.VLnum
        Dim StartTime, EndTime As New DateTime
        Dim RunTime As New TimeSpan
        StartTime = DateTime.Now

        hamop.MaximumLength += 10
        num4.Value = "+4"
        num1.Value = "+1"
        num2 = hamop.VLnMath.ATanRforPI(num1)
        PI = num2 * num4
        TxtBoxResultCPU.Text = Math.PI
        hamop.MaximumLength -= 10
        TxtBoxResult.Text = PI.Value

        EndTime = DateTime.Now
        RunTime = EndTime - StartTime
        TxtBoxNum2.Text = "Run Time " & RunTime.TotalMilliseconds & " Milliseconds  ---- " & RunTime.TotalSeconds & " Seconds"
    End Sub
    Private Sub BtnFn_e_Click(sender As Object, e As EventArgs) Handles BtnFn_e.Click
        Dim streCalc, streChatGPT, stt, st1, st2 As String      'this sub can be removed its only for testing
        Dim i, f, g As Integer
        'VARIABLE INITIALISATION
        streCalc = ""
        streChatGPT = ""
        'USE THE BELOW LINES TO COMPARE THE VALS OF e after commenting out the call statements below that 
        'which can be used to compare the vals of PI
        'Call streVal(streCalc)
        'Call streValChatGpt(streChatGPT)
        
        'USE THE BELOW LINES TO COMPARE THE VALS OF PI
        Call strPIValCalc(streCalc)
        Call strPIValChatGPT(streChatGPT)


       

        f = streCalc.Length
        g = streChatGPT.Length
        
        For i = 0 To streCalc.Length - 1
            If streCalc(i) <> streChatGPT(i) Then
                st1 = streCalc.Substring(i)
                st2 = streChatGPT.Substring(i)
                stt = "Mismatch at " & i.ToString
                MsgBox(stt, vbOKOnly, "Mismatch")
                Exit Sub
            End If
        Next
        TxtBoxResult.Text = "All the digits upto " & i & " are correct"

    End Sub
    Private Sub streVal(ByRef strEvalCalc As String) '10100 digits   run time 8392 seconds
        strEvalCalc =
        "2.718281828459045235360287471352662497757247093699959574966967627724076630353547594571382178525166427427466391932003059921817413596629043572" &
"90033429526059563073813232862794349076323382988075319525101901157383418793070215408914993488416750924476146066808226480016847741185374234544" &
"24371075390777449920695517027618386062613313845830007520449338265602976067371132007093287091274437470472306969772093101416928368190255151086" &
"57463772111252389784425056953696770785449969967946864454905987931636889230098793127736178215424999229576351482208269895193668033182528869398" &
"49646510582093923982948879332036250944311730123819706841614039701983767932068328237646480429531180232878250981945581530175671736133206981125" &
"09961818815930416903515988885193458072738667385894228792284998920868058257492796104841984443634632449684875602336248270419786232090021609902" &
"35304369941849146314093431738143640546253152096183690888707016768396424378140592714563549061303107208510383750510115747704171898610687396965" &
"52126715468895703503540212340784981933432106817012100562788023519303322474501585390473041995777709350366041699732972508868769664035557071622" &
"68447162560798826517871341951246652010305921236677194325278675398558944896970964097545918569563802363701621120477427228364896134225164450781" &
"82442352948636372141740238893441247963574370263755294448337998016125492278509257782562092622648326277933386566481627725164019105900491644998" &
"28931505660472580277863186415519565324425869829469593080191529872117255634754639644791014590409058629849679128740687050489585867174798546677" &
"57573205681288459205413340539220001137863009455606881667400169842055804033637953764520304024322566135278369511778838638744396625322498506549" &
"95886234281899707733276171783928034946501434558897071942586398772754710962953741521115136835062752602326484728703920764310059584116612054529" &
"70302364725492966693811513732275364509888903136020572481765851180630364428123149655070475102544650117272115551948668508003685322818315219600" &
"37356252794495158284188294787610852639813955990067376482922443752871846245780361929819713991475644882626039033814418232625150974827987779964" &
"37308997038886778227138360577297882412561190717663946507063304527954661855096666185664709711344474016070462621568071748187784437143698821855" &
"96709591025968620023537185887485696522000503117343920732113908032936344797273559552773490717837934216370120500545132638354400018632399149070" &
"54797780566978533580489669062951194324730995876552368128590413832411607226029983305353708761389396391779574540161372236187893652605381558415" &
"87186925538606164779834025435128439612946035291332594279490433729908573158029095863138268329147711639633709240031689458636060645845925126994" &
"65572483918656420975268508230754425459937691704197778008536273094171016343490769642372229435236612557250881477922315197477806056967253801718" &
"07763603462459278778465850656050780844211529697521890874019660906651803516501792504619501366585436632712549639908549144200014574760819302212" &
"06602433009641270489439039717719518069908699860663658323227870937650226014929101151717763594460202324930028040186772391028809786660565118326" &
"00436885088171572386698422422010249505518816948032210025154264946398128736776589276881635983124778865201411741109136011649950766290779436460" &
"05851941998560162647907615321038727557126992518275687989302761761146162549356495903798045838182323368612016243736569846703785853305275833337" &
"93990752166069238053369887956513728559388349989470741618155012539706464817194670834819721448889879067650379590366967249499254527903372963616" &
"26589760394985767413973594410237443297093554779826296145914429364514286171585873397467918975712119561873857836447584484235555810500256114923" &
"91518893099463428413936080383091662818811503715284967059741625628236092168075150177725387402564253470879089137291722828611515915683725241630" &
"77225440633787593105982676094420326192428531701878177296023541306067213604600038966109364709514141718577701418060644363681546444005331608778" &
"31431744408119494229755993140118886833148328027065538330046932901157441475631399972217038046170928945790962716622607407187499753592127560844" &
"14737823303270330168237193648002173285734935947564334129943024850235732214597843282641421684878721673367010615094243456984401873312810107945" &
"12722373788612605816566805371439612788873252737389039289050686532413806279602593038772769778379286840932536588073398845721874602100531148335" &
"13238500478271693762180049047955979592905916554705057775143081751126989851884087185640260353055837378324229241856256442550226721559802740126" &
"17971928047139600689163828665277009752767069777036439260224372841840883251848770472638440379530166905465937461619323840363893131364327137688" &
"84102681121989127522305625675625470172508634976536728860596675274086862740791285657699631378975303466061666980421826772456053066077389962421" &
"83408598820718646826232150802882863597468396543588566855037731312965879758105012149162076567699506597153447634703208532156036748286083786568" &
"03073062657633469774295634643716709397193060876963495328846833613038829431040800296873869117066666146800015121143442256023874474325250769387" &
"07777519329994213727721125884360871583483562696166198057252661220679754062106208064988291845439530152998209250300549825704339055357016865312" &
"05264956148572492573862069174036952135337325316663454665885972866594511364413703313936721185695539521084584072443238355860631068069649248512" &
"32632699514603596037297253198368423363904632136710116192821711150282801604488058802382031981493096369596735832742024988245684941273860566491" &
"35252670604623445054922758115170931492187959271800194096886698683703730220047531433818109270803001720593553052070070607223399946399057131158" &
"70996357773590271962850611465148375262095653467132900259943976631145459026858989791158370934193704411551219201171648805669459381311838437656" &
"20627846310490346293950029458341164824114969758326011800731699437393506966295712410273239138741754923071862454543222039552735295240245903805" &
"74450289224688628533654221381572213116328811205214648980518009202471939171055539011394331668151582884368760696110250517100739276238555338627" &
"25535388309606716446623709226468096712540618695021431762116681400975952814939072226011126811531083873176173232352636058381731510345957365382" &
"23534992935822836851007810884634349983518404451704270189381994243410090575376257767571118090088164183319201962623416288166521374717325477727" &
"78348877436651882875215668571950637193656539038944936642176400312152787022236646363575550356557694888654950027085392361710550213114741374410" &
"61344455441921013361729962856948991933691847294785807291560885103967819594298331864807560836795514966364489655929481878517840387733262470519" &
"45050419847742014183947731202815886845707290544057510601285258056594703046836344592652552137008068752009593453607316226118728173928074623094" &
"68536782310609792159936001994623799343421068781349734695924646975250624695861690917857397659519939299399556754271465491045686070209901260681" &
"87049841780791739240719459963230602547079017745275131868099822847308607665368668555164677029113368275631072233467261137054907953658345386371" &
"96235856312618387156774118738527722922594743373785695538456246801013905727871016512966636764451872465653730402443684140814488732957847348490" &
"00301947788802046032466084287535184836495919508288832320652212810419044804724794929134228495197002260131043006241071797150279343326340799596" &
"05314460532304885289729176598760166678119379323724538572096075822771784833616135826128962261181294559274627671377944875867536575448614076119" &
"31125958512655759734573015333642630767985443385761715333462325270572005303988289499034259566232975782488735029259166825894456894655992658454" &
"76269452878051650172067478541788798227680653665064191097343452887833862172615626958265447820567298775642632532159429441803994321700009054265" &
"07630955884658951717091476074371368933194690909819045012903070995662266203031826493657336984195557769637876249188528656866076005660256054457" &
"11337286840205574416030837052312242587223438854123179481388550075689381124935386318635287083799845692619981794523364087429591180747453419551" &
"42035172618420084550917084568236820089773945584267921427347756087964427920270831215015640634134161716644806981548376449157390012121704154787" &
"25919989438253649505147713793991472052195290793961376211072384942906163576045962312535060685376514231153496656837151166042207963944666211632" &
"55157729070978473156278277598788136491951257483328793771571459091064841642678309949723674420175862269402159407924480541255360431317992696739" &
"15754241929660731239376354213923061787675395871143610408940996608947141834069836299367536262154524729846421375289107988438130609555262272083" &
"75186298370667872244301957937937860721072542772890717328548743743557819665117166183308811291202452040486822000723440350254482028342541878846" &
"53602591506445271657700044521097735585897622655484941621714989532383421600114062950718490427789258552743035221396835679018076406042138307308" &
"77446017084268827226117718084266433365178000217190344923426426629226145600433738386833555534345300426481847398921562708609565062934040526494" &
"32442614456659212912256488935696550091543064261342526684725949143142393988454324863274618428466559853323122104662598901417121034460842716166" &
"19001257195870793217569698544013397622096749454185407118446433946990162698351607848924514058940946395267807354579700307051163682519487701189" &
"76400282764841416058720618418529718915401968825328930914966534575357142731848201638464483249903788606900807270932767312758196656394114896171" &
"68329804551397295066876047409154204284299935410258291135022416907694316685742425225090269390348148564513030699251995904363840284292674125734" &
"22447765584177886171737265462085498294498946787350929581652632072258992368768457017823038096567883112289305809140572610865884845873101658151" &
"16753332767488701482916741970151255978257270740643180860142814902414678047232759768426963393577354293018673943971638861176420900406866339885" &
"68416810038723892144831760701166845038872123643670433140911557332801829779887365909166596124020217785588548761761619893707943800566633648843" &
"65089144805571039765214696027662583599051987042300179465536788567430285974600143785483237068701190078499404930918919181649327259774030074879" &
"681484882342932023012"

    End Sub
    Private Sub streValChatGpt(ByRef strEvalChatGPT As String)   '10500 digits
        strEvalChatGPT =
"2.718281828459045235360287471352662497757247093699959574966967627724076630353547594571382178525166427427466391932003059921817413596629043572" &
"90033429526059563073813232862794349076323382988075319525101901157383418793070215408914993488416750924476146066808226480016847741185374234544" &
"24371075390777449920695517027618386062613313845830007520449338265602976067371132007093287091274437470472306969772093101416928368190255151086" &
"57463772111252389784425056953696770785449969967946864454905987931636889230098793127736178215424999229576351482208269895193668033182528869398" &
"49646510582093923982948879332036250944311730123819706841614039701983767932068328237646480429531180232878250981945581530175671736133206981125" &
"09961818815930416903515988885193458072738667385894228792284998920868058257492796104841984443634632449684875602336248270419786232090021609902" &
"35304369941849146314093431738143640546253152096183690888707016768396424378140592714563549061303107208510383750510115747704171898610687396965" &
"52126715468895703503540212340784981933432106817012100562788023519303322474501585390473041995777709350366041699732972508868769664035557071622" &
"68447162560798826517871341951246652010305921236677194325278675398558944896970964097545918569563802363701621120477427228364896134225164450781" &
"82442352948636372141740238893441247963574370263755294448337998016125492278509257782562092622648326277933386566481627725164019105900491644998" &
"28931505660472580277863186415519565324425869829469593080191529872117255634754639644791014590409058629849679128740687050489585867174798546677" &
"57573205681288459205413340539220001137863009455606881667400169842055804033637953764520304024322566135278369511778838638744396625322498506549" &
"95886234281899707733276171783928034946501434558897071942586398772754710962953741521115136835062752602326484728703920764310059584116612054529" &
"70302364725492966693811513732275364509888903136020572481765851180630364428123149655070475102544650117272115551948668508003685322818315219600" &
"37356252794495158284188294787610852639813955990067376482922443752871846245780361929819713991475644882626039033814418232625150974827987779964" &
"37308997038886778227138360577297882412561190717663946507063304527954661855096666185664709711344474016070462621568071748187784437143698821855" &
"96709591025968620023537185887485696522000503117343920732113908032936344797273559552773490717837934216370120500545132638354400018632399149070" &
"54797780566978533580489669062951194324730995876552368128590413832411607226029983305353708761389396391779574540161372236187893652605381558415" &
"87186925538606164779834025435128439612946035291332594279490433729908573158029095863138268329147711639633709240031689458636060645845925126994" &
"65572483918656420975268508230754425459937691704197778008536273094171016343490769642372229435236612557250881477922315197477806056967253801718" &
"07763603462459278778465850656050780844211529697521890874019660906651803516501792504619501366585436632712549639908549144200014574760819302212" &
"06602433009641270489439039717719518069908699860663658323227870937650226014929101151717763594460202324930028040186772391028809786660565118326" &
"00436885088171572386698422422010249505518816948032210025154264946398128736776589276881635983124778865201411741109136011649950766290779436460" &
"05851941998560162647907615321038727557126992518275687989302761761146162549356495903798045838182323368612016243736569846703785853305275833337" &
"93990752166069238053369887956513728559388349989470741618155012539706464817194670834819721448889879067650379590366967249499254527903372963616" &
"26589760394985767413973594410237443297093554779826296145914429364514286171585873397467918975712119561873857836447584484235555810500256114923" &
"91518893099463428413936080383091662818811503715284967059741625628236092168075150177725387402564253470879089137291722828611515915683725241630" &
"77225440633787593105982676094420326192428531701878177296023541306067213604600038966109364709514141718577701418060644363681546444005331608778" &
"31431744408119494229755993140118886833148328027065538330046932901157441475631399972217038046170928945790962716622607407187499753592127560844" &
"14737823303270330168237193648002173285734935947564334129943024850235732214597843282641421684878721673367010615094243456984401873312810107945" &
"12722373788612605816566805371439612788873252737389039289050686532413806279602593038772769778379286840932536588073398845721874602100531148335" &
"13238500478271693762180049047955979592905916554705057775143081751126989851884087185640260353055837378324229241856256442550226721559802740126" &
"17971928047139600689163828665277009752767069777036439260224372841840883251848770472638440379530166905465937461619323840363893131364327137688" &
"84102681121989127522305625675625470172508634976536728860596675274086862740791285657699631378975303466061666980421826772456053066077389962421" &
"83408598820718646826232150802882863597468396543588566855037731312965879758105012149162076567699506597153447634703208532156036748286083786568" &
"03073062657633469774295634643716709397193060876963495328846833613038829431040800296873869117066666146800015121143442256023874474325250769387" &
"07777519329994213727721125884360871583483562696166198057252661220679754062106208064988291845439530152998209250300549825704339055357016865312" &
"05264956148572492573862069174036952135337325316663454665885972866594511364413703313936721185695539521084584072443238355860631068069649248512" &
"32632699514603596037297253198368423363904632136710116192821711150282801604488058802382031981493096369596735832742024988245684941273860566491" &
"35252670604623445054922758115170931492187959271800194096886698683703730220047531433818109270803001720593553052070070607223399946399057131158" &
"70996357773590271962850611465148375262095653467132900259943976631145459026858989791158370934193704411551219201171648805669459381311838437656" &
"20627846310490346293950029458341164824114969758326011800731699437393506966295712410273239138741754923071862454543222039552735295240245903805" &
"74450289224688628533654221381572213116328811205214648980518009202471939171055539011394331668151582884368760696110250517100739276238555338627" &
"25535388309606716446623709226468096712540618695021431762116681400975952814939072226011126811531083873176173232352636058381731510345957365382" &
"23534992935822836851007810884634349983518404451704270189381994243410090575376257767571118090088164183319201962623416288166521374717325477727" &
"78348877436651882875215668571950637193656539038944936642176400312152787022236646363575550356557694888654950027085392361710550213114741374410" &
"61344455441921013361729962856948991933691847294785807291560885103967819594298331864807560836795514966364489655929481878517840387733262470519" &
"45050419847742014183947731202815886845707290544057510601285258056594703046836344592652552137008068752009593453607316226118728173928074623094" &
"68536782310609792159936001994623799343421068781349734695924646975250624695861690917857397659519939299399556754271465491045686070209901260681" &
"87049841780791739240719459963230602547079017745275131868099822847308607665368668555164677029113368275631072233467261137054907953658345386371" &
"96235856312618387156774118738527722922594743373785695538456246801013905727871016512966636764451872465653730402443684140814488732957847348490" &
"00301947788802046032466084287535184836495919508288832320652212810419044804724794929134228495197002260131043006241071797150279343326340799596" &
"05314460532304885289729176598760166678119379323724538572096075822771784833616135826128962261181294559274627671377944875867536575448614076119" &
"31125958512655759734573015333642630767985443385761715333462325270572005303988289499034259566232975782488735029259166825894456894655992658454" &
"76269452878051650172067478541788798227680653665064191097343452887833862172615626958265447820567298775642632532159429441803994321700009054265" &
"07630955884658951717091476074371368933194690909819045012903070995662266203031826493657336984195557769637876249188528656866076005660256054457" &
"11337286840205574416030837052312242587223438854123179481388550075689381124935386318635287083799845692619981794523364087429591180747453419551" &
"42035172618420084550917084568236820089773945584267921427347756087964427920270831215015640634134161716644806981548376449157390012121704154787" &
"25919989438253649505147713793991472052195290793961376211072384942906163576045962312535060685376514231153496656837151166042207963944666211632" &
"55157729070978473156278277598788136491951257483328793771571459091064841642678309949723674420175862269402159407924480541255360431317992696739" &
"15754241929660731239376354213923061787675395871143610408940996608947141834069836299367536262154524729846421375289107988438130609555262272083" &
"75186298370667872244301957937937860721072542772890717328548743743557819665117166183308811291202452040486822000723440350254482028342541878846" &
"53602591506445271657700044521097735585897622655484941621714989532383421600114062950718490427789258552743035221396835679018076406042138307308" &
"77446017084268827226117718084266433365178000217190344923426426629226145600433738386833555534345300426481847398921562708609565062934040526494" &
"32442614456659212912256488935696550091543064261342526684725949143142393988454324863274618428466559853323122104662598901417121034460842716166" &
"19001257195870793217569698544013397622096749454185407118446433946990162698351607848924514058940946395267807354579700307051163682519487701189" &
"76400282764841416058720618418529718915401968825328930914966534575357142731848201638464483249903788606900807270932767312758196656394114896171" &
"68329804551397295066876047409154204284299935410258291135022416907694316685742425225090269390348148564513030699251995904363840284292674125734" &
"22447765584177886171737265462085498294498946787350929581652632072258992368768457017823038096567883112289305809140572610865884845873101658151" &
"16753332767488701482916741970151255978257270740643180860142814902414678047232759768426963393577354293018673943971638861176420900406866339885" &
"68416810038723892144831760701166845038872123643670433140911557332801829779887365909166596124020217785588548761761619893707943800566633648843" &
"65089144805571039765214696027662583599051987042300179465536788567430285974600143785483237068701190078499404930918919181649327259774030074879" &
"68148488234293202301212803232746039221968752834051690697419425761467397811071546418627336909158497318501118396048253351874843892317729261354" &
"30249325628963713619772854566229244616444972845978677115741256703078718851093363444801496752406185365695320741705334867827548278154155619669" &
"11055101472799040386897220465550833170782394808785990501947563108984124144672821865459971596639015641941751820935932616316888380132758752601" &
"46"

    End Sub
    Private Sub strPIValCalc(ByRef strPIvalCal As String)    '3001 digits    Run Time 4747 seconds
        strPIvalCal =
"3.141592653589793238462643383279502884197169399375105820974944592307816406286208998628034825342117067982148086513282306647093844609550582" &
"23172535940812848111745028410270193852110555964462294895493038196442881097566593344612847564823378678316527120190914564856692346034861045" &
"4326648213393607260249141273724587006606315588174881520920962829254091715364367892590360011330530548820466521384146951941511609433057270" &
"3657595919530921861173819326117931051185480744623799627495673518857527248912279381830119491298336733624406566430860213949463952247371907" &
"0217986094370277053921717629317675238467481846766940513200056812714526356082778577134275778960917363717872146844090122495343014654958537" &
"1050792279689258923542019956112129021960864034418159813629774771309960518707211349999998372978049951059731732816096318595024459455346908" &
"3026425223082533446850352619311881710100031378387528865875332083814206171776691473035982534904287554687311595628638823537875937519577818" &
"5778053217122680661300192787661119590921642019893809525720106548586327886593615338182796823030195203530185296899577362259941389124972177" &
"5283479131515574857242454150695950829533116861727855889075098381754637464939319255060400927701671139009848824012858361603563707660104710" &
"1819429555961989467678374494482553797747268471040475346462080466842590694912933136770289891521047521620569660240580381501935112533824300" &
"3558764024749647326391419927260426992279678235478163600934172164121992458631503028618297455570674983850549458858692699569092721079750930" &
"2955321165344987202755960236480665499119881834797753566369807426542527862551818417574672890977772793800081647060016145249192173217214772" &
"3501414419735685481613611573525521334757418494684385233239073941433345477624168625189835694855620992192221842725502542568876717904946016" &
"5346680498862723279178608578438382796797668145410095388378636095068006422512520511739298489608412848862694560424196528502221066118630674" &
"4278622039194945047123713786960956364371917287467764657573962413890865832645995813390478027590099465764078951269468398352595709825822620" &
"5224894077267194782684826014769909026401363944374553050682034962524517493996514314298091906592509372216964615157098583874105978859597729" &
"7549893016175392846813826868386894277415599185592524595395943104997252468084598727364469584865383673622262609912460805124388439045124413" &
"6549762780797715691435997700129616089441694868555848406353422072225828488648158456028506016842739452267467678895252138522549954666727823" &
"9864565961163548862305774564980355936345681743241125150760694794510965960940252288797108931456691368672287489405601015033086179286809208" &
"7476091782493858900971490967598526136554978189312978482168299894872265880485756401427047755513237964145152374623436454285844479526586782" &
"1051141354735739523113427166102135969536231442952484937187110145765403590279934403742007310578539062198387447808478489683321445713868751" &
"9435064302184531910484810053706146806749192781911979399520614196634287544406437451237181921799983910159195618146751426912397489409071864" &
"9423196156794520809514655022"
    End Sub
    Private Sub strPIValChatGPT(ByRef strPIvalChatGPT As String) '10500 digits
        strPIvalChatGPT =
"3.141592653589793238462643383279502884197169399375105820974944592307816406286208998628034825342117067982148086513282306647093844609550582" &
"23172535940812848111745028410270193852110555964462294895493038196442881097566593344612847564823378678316527120190914564856692346034861045" &
"43266482133936072602491412737245870066063155881748815209209628292540917153643678925903600113305305488204665213841469519415116094330572703" &
"65759591953092186117381932611793105118548074462379962749567351885752724891227938183011949129833673362440656643086021394946395224737190702" &
"17986094370277053921717629317675238467481846766940513200056812714526356082778577134275778960917363717872146844090122495343014654958537105" &
"07922796892589235420199561121290219608640344181598136297747713099605187072113499999983729780499510597317328160963185950244594553469083026" &
"42522308253344685035261931188171010003137838752886587533208381420617177669147303598253490428755468731159562863882353787593751957781857780" &
"53217122680661300192787661119590921642019893809525720106548586327886593615338182796823030195203530185296899577362259941389124972177528347" &
"91315155748572424541506959508295331168617278558890750983817546374649393192550604009277016711390098488240128583616035637076601047101819429" &
"55596198946767837449448255379774726847104047534646208046684259069491293313677028989152104752162056966024058038150193511253382430035587640" &
"24749647326391419927260426992279678235478163600934172164121992458631503028618297455570674983850549458858692699569092721079750930295532116" &
"53449872027559602364806654991198818347977535663698074265425278625518184175746728909777727938000816470600161452491921732172147723501414419" &
"73568548161361157352552133475741849468438523323907394143334547762416862518983569485562099219222184272550254256887671790494601653466804988" &
"62723279178608578438382796797668145410095388378636095068006422512520511739298489608412848862694560424196528502221066118630674427862203919" &
"49450471237137869609563643719172874677646575739624138908658326459958133904780275900994657640789512694683983525957098258226205224894077267" &
"19478268482601476990902640136394437455305068203496252451749399651431429809190659250937221696461515709858387410597885959772975498930161753" &
"92846813826868386894277415599185592524595395943104997252468084598727364469584865383673622262609912460805124388439045124413654976278079771" &
"56914359977001296160894416948685558484063534220722258284886481584560285060168427394522674676788952521385225499546667278239864565961163548" &
"86230577456498035593634568174324112515076069479451096596094025228879710893145669136867228748940560101503308617928680920874760917824938589" &
"00971490967598526136554978189312978482168299894872265880485756401427047755513237964145152374623436454285844479526586782105114135473573952" &
"31134271661021359695362314429524849371871101457654035902799344037420073105785390621983874478084784896833214457138687519435064302184531910" &
"48481005370614680674919278191197939952061419663428754440643745123718192179998391015919561814675142691239748940907186494231961567945208095" &
"14655022523160388193014209376213785595663893778708303906979207734672218256259966150142150306803844773454920260541466592520149744285073251" &
"86660021324340881907104863317346496514539057962685610055081066587969981635747363840525714591028970641401109712062804390397595156771577004" &
"20337869936007230558763176359421873125147120532928191826186125867321579198414848829164470609575270695722091756711672291098169091528017350" &
"67127485832228718352093539657251210835791513698820914442100675103346711031412671113699086585163983150197016515116851714376576183515565088" &
"49099898599823873455283316355076479185358932261854896321329330898570642046752590709154814165498594616371802709819943099244889575712828905" &
"92323326097299712084433573265489382391193259746366730583604142813883032038249037589852437441702913276561809377344403070746921120191302033" &
"03801976211011004492932151608424448596376698389522868478312355265821314495768572624334418930396864262434107732269780280731891544110104468" &
"23252716201052652272111660396665573092547110557853763466820653109896526918620564769312570586356620185581007293606598764861179104533488503" &
"46113657686753249441668039626579787718556084552965412665408530614344431858676975145661406800700237877659134401712749470420562230538994561" &
"31407112700040785473326993908145466464588079727082668306343285878569830523580893306575740679545716377525420211495576158140025012622859413" &
"02164715509792592309907965473761255176567513575178296664547791745011299614890304639947132962107340437518957359614589019389713111790429782" &
"85647503203198691514028708085990480109412147221317947647772622414254854540332157185306142288137585043063321751829798662237172159160771669" &
"25474873898665494945011465406284336639379003976926567214638530673609657120918076383271664162748888007869256029022847210403172118608204190" &
"00422966171196377921337575114959501566049631862947265473642523081770367515906735023507283540567040386743513622224771589150495309844489333" &
"09634087807693259939780541934144737744184263129860809988868741326047215695162396586457302163159819319516735381297416772947867242292465436" &
"68009806769282382806899640048243540370141631496589794092432378969070697794223625082216889573837986230015937764716512289357860158816175578" &
"29735233446042815126272037343146531977774160319906655418763979293344195215413418994854447345673831624993419131814809277771038638773431772" &
"07545654532207770921201905166096280490926360197598828161332316663652861932668633606273567630354477628035045077723554710585954870279081435" &
"62401451718062464362679456127531813407833033625423278394497538243720583531147711992606381334677687969597030983391307710987040859133746414" &
"42822772634659470474587847787201927715280731767907707157213444730605700733492436931138350493163128404251219256517980694113528013147013047" &
"81643788518529092854520116583934196562134914341595625865865570552690496520985803385072242648293972858478316305777756068887644624824685792" &
"60395352773480304802900587607582510474709164396136267604492562742042083208566119062545433721315359584506877246029016187667952406163425225" &
"77195429162991930645537799140373404328752628889639958794757291746426357455254079091451357111369410911939325191076020825202618798531887705" &
"84297259167781314969900901921169717372784768472686084900337702424291651300500516832336435038951702989392233451722013812806965011784408745" &
"19601212285993716231301711444846409038906449544400619869075485160263275052983491874078668088183385102283345085048608250393021332197155184" &
"30635455007668282949304137765527939751754613953984683393638304746119966538581538420568533862186725233402830871123282789212507712629463229" &
"56398989893582116745627010218356462201349671518819097303811980049734072396103685406643193950979019069963955245300545058068550195673022921" &
"91393391856803449039820595510022635353619204199474553859381023439554495977837790237421617271117236434354394782218185286240851400666044332" &
"58885698670543154706965747458550332323342107301545940516553790686627333799585115625784322988273723198987571415957811196358330059408730681" &
"21602876496286744604774649159950549737425626901049037781986835938146574126804925648798556145372347867330390468838343634655379498641927056" &
"38729317487233208376011230299113679386270894387993620162951541337142489283072201269014754668476535761647737946752004907571555278196536213" &
"23926406160136358155907422020203187277605277219005561484255518792530343513984425322341576233610642506390497500865627109535919465897514131" &
"03482276930624743536325691607815478181152843667957061108615331504452127473924544945423682886061340841486377670096120715124914043027253860" &
"76482363414334623518975766452164137679690314950191085759844239198629164219399490723623464684411739403265918404437805133389452574239950829" &
"65912285085558215725031071257012668302402929525220118726767562204154205161841634847565169998116141010029960783869092916030288400269104140" &
"79288621507842451670908700069928212066041837180653556725253256753286129104248776182582976515795984703562226293486003415872298053498965022" &
"62917487882027342092222453398562647669149055628425039127577102840279980663658254889264880254566101729670266407655904290994568150652653053" &
"71829412703369313785178609040708667114965583434347693385781711386455873678123014587687126603489139095620099393610310291616152881384379099" &
"04231747336394804575931493140529763475748119356709110137751721008031559024853090669203767192203322909433467685142214477379393751703443661" &
"99104033751117354719185504644902636551281622882446257591633303910722538374218214088350865739177150968288747826569959957449066175834413752" &
"23970968340800535598491754173818839994469748676265516582765848358845314277568790029095170283529716344562129640435231176006651012412006597" &
"55851276178583829204197484423608007193045761893234922927965019875187212726750798125547095890455635792122103334669749923563025494780249011" &
"41952123828153091140790738602515227429958180724716259166854513331239480494707911915326734302824418604142636395480004480026704962482017928" &
"96476697583183271314251702969234889627668440323260927524960357996469256504936818360900323809293459588970695365349406034021665443755890045" &
"63288225054525564056448246515187547119621844396582533754388569094113031509526179378002974120766514793942590298969594699556576121865619673" &
"37862362561252163208628692221032748892186543648022967807057656151446320469279068212073883778142335628236089632080682224680122482611771858" &
"96381409183903673672220888321513755600372798394004152970028783076670944474560134556417254370906979396122571429894671543578468788614445812" &
"31459357198492252847160504922124247014121478057345510500801908699603302763478708108175450119307141223390866393833952942578690507643100638" &
"35198343893415961318543475464955697810382930971646514384070070736041123735998434522516105070270562352660127648483084076118301305279320542" &
"74628654036036745328651057065874882256981579367897669742205750596834408697350201410206723585020072452256326513410559240190274216248439140" &
"35998953539459094407046912091409387001264560016237428802109276457931065792295524988727584610126483699989225695968815920560010165525637567" &
"85667227966198857827948488558343975187445455129656344348039664205579829368043522027709842942325330225763418070394769941597915945300697521" &
"48293366555661567873640053666564165473217043903521329543529169414599041608753201868379370234888689479151071637852902345292440773659495630" &
"51007421087142613497459561513849871375704710178795731042296906667021449863746459528082436944578977233004876476524133907592043401963403911" &
"473202338071509522201068256342747164602433544005152126693249341967397704159568375355516673"
    End Sub


End Class