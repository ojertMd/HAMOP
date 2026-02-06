Public Class VLNumTests
    Dim hamop As New Hamop
    Dim num1 As New Hamop.VLnum
    Dim rnd As New Random()
    Dim vArray(,) As String
    Public Sub testVLNum(ByRef I As Integer, ByRef finish As Boolean)
        'THIS SUB IS USED TO TEST THE FORMAT OF VLnum
        'BELOW ARRAY FIRST NUMBER IS THE VAL & NEXT IS MAXLENGTH IT SHOULD BE ENTERED IN THIS SEQUENCE
        Dim strArray(,) As String = {{"+0", "10"}, {"0", "10"}, {"+12", "10"}, {"-12", "10"}, {"12", "10"}, {"+0.5", "10"}, {"-0.5", "10"}, _
                                    {"0.5", "10"}, {"+2.5", "10"}, {"-2.5", "10"}, {"+2.5E+10", "20"}, {"+2.5E-10", "20"}, {"2.5E10", "20"}, _
                                    {"2.5E+2", "10"}, {"2.5E-2", "10"}, {"2.5E2", "10"}, {"0.2546589", "20"}, {"12.256985658745465", "20"}, _
                                    {"23.6584", "10"}, {"23.6584000000125", "28"}, {"0.000258965E4", "28"}, {"0.0562E1", "28"}}

        If I <= strArray.GetUpperBound(0) Then
            Form1.TxtBoxMaxLength.Text = strArray(I, 1)
            Form1.TxtBoxMaxLength.Focus()       'THIS LINE AND THE LINE BELOW NEXT IS REQUIRED ORELSE _MaxLength WILL NOT BE UPDATED
            Form1.TxtBoxNum1.Text = strArray(I, 0)
            Form1.Button1.Focus()            'THIS LINE IS REQUIRED ORELSE _MaxLength WILL NOT BE UPDATED
            num1.Value = strArray(I, 0)
            Form1.TxtBoxResult.Text = num1.Value
            If I = strArray.GetUpperBound(0) Then
                finish = True
            End If
        End If
    End Sub
    Public Sub testVLNum1(ByRef I As Integer, ByRef finish As Boolean)
        'THIS SUB IS USED TO TEST THE FORMAT OF VLnum
        'BELOW ARRAY FIRST NUMBER IS THE VAL & NEXT IS MAXLENGTH IT SHOULD BE ENTERED IN THIS SEQUENCE THIS IS FOR TESTING REGEX
        'BELOW ARRAY IS FOR TESTING NUMGEN StrCHECKING & Trimming WHEN THIS TEST IS RUN Throw New 'InvalidOperationException' IN StrCHECKING
        'SHOULD BE COMMENTED OUT ORELSE ALL THE TESTS WILL NOT RUN
        Dim strArray(,) As String = {{"+12  36   89   8745", "28"}, {"-123.456", "10"}, {"123.456", "10"}, {"123E+3", "10"}, {"123.456E+3", "10"}, _
                                    {"123.456E-3", "10"}, {"123.456E3", "10"}, {"123.456E+5", "10"}, {"123.456E-5", "10"}, _
                                    {"000123.456", "28"}, {"123.456000", "28"}, {"++123.456", "28"}, {"-123.456E3E2", "28"}, {"123.45.6000", "28"}, _
                                     {"000000.456", "28"}, {"123.456E003", "28"}, {"000a123.456", "28"}, {"12563214785965", "10"}, _
                                     {"1256.32541258965874", "10"}}

        If I <= strArray.GetUpperBound(0) Then
            Form1.TxtBoxMaxLength.Text = strArray(I, 1)
            Form1.TxtBoxMaxLength.Focus()
            Form1.TxtBoxNum1.Text = strArray(I, 0)
            Form1.Button1.Focus()
            num1.Value = strArray(I, 0)
            Form1.TxtBoxResult.Text = num1.Value
            If I = strArray.GetUpperBound(0) Then
                finish = True
            End If
        End If

    End Sub
    Public Sub operatorTest(ByRef I As Integer, ByRef finish As Boolean)
        'BELOW ARRAY FIRST NUMBER IS THE NUM1 SECOND NO. IS NUM2 & NEXT IS MAXLENGTH & LAST IS DigitGroupLength IT SHOULD BE ENTERED IN THIS SEQUENCE
        Dim strArray(,) As String = {{"123", "456", "10", "1"}, {".0749970", "128", "15", "1"}, _
                                     {"999999.6666666666666666666666666666", "0.3333333333333333333333999999999999", "28", "1"}, _
                                     {"123.456789", "123456.789", "15", "1"}, {"123456.789", "123.456789", "15", "1"}, {"123.456789", "651.325625", "10", "1"}, _
                                     {"123.45678912000003456789987", "123.45678912000003456789987", "28", "5"}, _
                                     {"12345678.91200000000003456789987", "1234567891.200000000003456789987", "32", "4"}, _
                                     {"823.456789", "199456.789", "15", "10"}, {"823.456789", "999456.789", "15", "10"}, _
                                     {"105401.3020166", "163069.3", "28", "10"}, {"125401.3020166", "123069.3", "28", "10"}, _
                                     {"1.254013020166", "1.230693040166", "28", "10"}, {"105401.3020166", "163069.3020166", "28", "10"}}


        If I <= strArray.GetUpperBound(0) Then
            Form1.TxtBoxDigitGroupLength.Text = strArray(I, 3)
            Form1.TxtBoxDigitGroupLength.Focus()
            Form1.TxtBoxMaxLength.Text = strArray(I, 2)
            Form1.TxtBoxMaxLength.Focus()
            Form1.TxtBoxNum1.Text = strArray(I, 0)
            Form1.Button1.Focus()
            Form1.TxtBoxNum2.Text = strArray(I, 1)
            If I = strArray.GetUpperBound(0) Then
                finish = True
            End If
        End If
        Form1.lblResultComp.Text = ""
        Form1.TxtBoxResult.Text = ""
        Form1.TxtBoxResultCPU.Text = ""
    End Sub
    Public Sub IntGen(ByRef I As Integer, ByRef finish As Boolean, ByVal MaxLn As Integer)
        Dim size1, size2, sgn As Integer
        Dim sign As String = ""
        Dim numb1 As String = ""
        Dim numb2 As String = ""
        Dim numb As String = ""
        size1 = rnd.Next(MaxLn + 1)
        sgn = rnd.Next(100)
        If ((sgn Mod 2) = 0) Then
            sign = "+"
        Else
            sign = "-"
        End If

        For j = 0 To 10
            numb1 = rnd.Next(Integer.MaxValue).ToString
            numb = numb & numb1
            If numb.Length > size1 Then
                numb = numb.Substring(0, size1)
                Exit For
            End If
        Next
        If numb = "" Then numb = "0" : sign = "+"
        Form1.TxtBoxNum1.Text = sign & numb
        numb = ""
        size2 = rnd.Next(MaxLn + 1)
        sgn = rnd.Next(100)
        If ((sgn Mod 2) = 0) Then
            sign = "+"
        Else
            sign = "-"
        End If

        For k = 0 To 10
            numb2 = rnd.Next(Integer.MaxValue).ToString
            numb = numb & numb2
            If numb.Length > size2 Then
                numb = numb.Substring(0, size2)
                Exit For
            End If
        Next
        If numb = "" Then numb = "0" : sign = "+"
        Form1.TxtBoxNum2.Text = sign & numb

        If I = 10 Then
            finish = True
        End If
        Form1.lblResultComp.Text = ""
        Form1.TxtBoxResult.Text = ""
        Form1.TxtBoxResultCPU.Text = ""
    End Sub
    Public Sub DblGen(ByRef I As Integer, ByRef finish As Boolean, ByVal MaxLn As Integer)
        Dim size1, size2, IDselect, sgn As Integer
        Dim sltID As Boolean
        Dim sign As String = ""
        Dim numb1 As String = ""
        Dim numb2 As String = ""
        Dim numb As String = ""
        size1 = rnd.Next(MaxLn + 1)
        IDselect = rnd.Next(10)
        If IDselect < 5 Then
            sltID = True
        Else
            sltID = False
        End If
        sgn = rnd.Next(100)
        If ((sgn Mod 2) = 0) Then
            sign = "+"
        Else
            sign = "-"
        End If

        For j = 0 To 10
            numb1 = rnd.Next(Integer.MaxValue).ToString
            numb = numb & numb1
            If numb.Length > size1 Then
                numb = numb.Substring(0, size1)
                Exit For
            End If
        Next
        If numb = "" Then
            Form1.TxtBoxNum1.Text = "0"
        Else
            If sltID = True Then
                Form1.TxtBoxNum1.Text = sign & "0." & numb
            Else
                Form1.TxtBoxNum1.Text = sign & numb
            End If
        End If
        'If numb = "" Then numb = "0"
        'Form1.TxtBoxNum1.Text = "0." & numb
        'Form1.TxtBoxNum1.Text = numb
        numb = ""
        size2 = rnd.Next(MaxLn + 1)
        sgn = rnd.Next(100)
        If ((sgn Mod 2) = 0) Then
            sign = "+"
        Else
            sign = "-"
        End If

        For k = 0 To 10
            numb2 = rnd.Next(Integer.MaxValue).ToString
            numb = numb & numb2
            If numb.Length > size2 Then
                numb = numb.Substring(0, size2)
                Exit For
            End If
        Next
        If numb = "" Then
            Form1.TxtBoxNum2.Text = "0"
        Else
            Form1.TxtBoxNum2.Text = sign & "0." & numb
        End If
        'If numb = "" Then numb = "0"
        'Form1.TxtBoxNum2.Text = "0." & numb

        If I = 10 Then
            finish = True
        End If
        Form1.lblResultComp.Text = ""
        Form1.TxtBoxResult.Text = ""
        Form1.TxtBoxResultCPU.Text = ""
    End Sub
    Public Sub IntDblGen(ByRef I As Integer, ByRef finish As Boolean, ByVal MaxLn As Integer)
        Dim size1, size2, Intlen, sgn As Integer
        Dim sign As String = ""
        Dim numb1 As String = ""
        Dim numb2 As String = ""
        Dim numb As String = ""
        size1 = rnd.Next(MaxLn + 1)
        Intlen = rnd.Next(size1)
        sgn = rnd.Next(100)
        If ((sgn Mod 2) = 0) Then
            sign = "+"
        Else
            sign = "-"
        End If

        For j = 0 To 10
            numb1 = rnd.Next(Integer.MaxValue).ToString
            numb = numb & numb1
            If numb.Length > size1 Then
                numb = numb.Substring(0, size1)
                Exit For
            End If
        Next
        numb = numb.Substring(0, Intlen) & "." & numb.Substring(Intlen)
        If numb = "." Then
            numb = "0"
        Else
            numb = sign & numb
        End If

        Form1.TxtBoxNum1.Text = numb
        numb = ""
        size2 = rnd.Next(MaxLn + 1)
        Intlen = rnd.Next(size2)
        sgn = rnd.Next(100)
        If ((sgn Mod 2) = 0) Then
            sign = "+"
        Else
            sign = "-"
        End If


        For k = 0 To 10
            numb2 = rnd.Next(Integer.MaxValue).ToString
            numb = numb & numb2
            If numb.Length > size2 Then
                numb = numb.Substring(0, size2)
                Exit For
            End If
        Next
        numb = numb.Substring(0, Intlen) & "." & numb.Substring(Intlen)
        If numb = "." Then
            numb = "0"
        Else
            numb = sign & numb
        End If
        Form1.TxtBoxNum2.Text = numb

        If I = 10 Then
            finish = True
        End If
        Form1.lblResultComp.Text = ""
        Form1.TxtBoxResult.Text = ""
        Form1.TxtBoxResultCPU.Text = ""
    End Sub

    Public Sub TESTvalsOPR(ByRef I As Integer, ByRef finish As Boolean, ByVal MaxLn As Integer)
        Dim k As String
        Dim data As New List(Of String())
        data.Add(New String() {"105401.3020166", "163069.3", "27", "10"})   'this vals will make the leading digits of subtraction zeros with Addone
        data.Add(New String() {"125401.3020166", "123069.3", "28", "10"})   'this vals will make the leading digits of subtraction zeros with compli
        data.Add(New String() {"1.254013020166", "1.230693040166", "28", "10"}) 'this vals will make the trailing digits of subtraction zeros
        data.Add(New String() {"105401.3020166", "163069.3020166", "28", "10"}) 'this vals will make the trailing digits of subtraction zeros
        data.Add(New String() {"+128668331620582274", "-69894913811487853", "28", "10"})
        data.Add(New String() {".0749970", "128", "28", "10"})       'This will make leading zeros in ip in multiplication
        data.Add(New String() {".0749975", "128.36", "28", "10"})    'This will make trailing zeros in dp in multiplication
        data.Add(New String() {"36", "12", "28", "10"})            'In division ans is 3.00 why the last 2 zeros
        data.Add(New String() {"360", "12", "15", "10"})              'In division the last digit is rounded wrong when the next digit is 0 why 
        data.Add(New String() {"75869545", "45845", "28", "10"})        'division
        data.Add(New String() {"18", "9", "28", "10"})                  'division
        data.Add(New String() {"0.012", "0.003", "28", "10"})           'division
        data.Add(New String() {"1.1000000110000001100000011", "1", "28", "10"}) 'division
        data.Add(New String() {"45", "45874596854", "28", "10"})                'division
        data.Add(New String() {"4582", ".00002548", "28", "10"})                'division
        data.Add(New String() {"+1.9", "+0.1", "28", "10"})                     'division
        data.Add(New String() {"99999", "77777", "5", "1"})                     'addition
        data.Add(New String() {"99999", "7777777", "7", "1"})                   'addition
        data.Add(New String() {"99999", "7777777", "5", "1"})                   'addition
        data.Add(New String() {"9.9999", "1.1111", "5", "1"})                   'addition
        data.Add(New String() {"9.8889", "1.1111", "5", "1"})                   'addition
        data.Add(New String() {"61.800", "5", "28", "10"})                      'division
        'data.Add(New String() {"", "", "", ""})
        'data.Add(New String() {"", "", "", ""})
        'data.Add(New String() {"", "", "", ""})
        'data.Add(New String() {"", "", "", ""})
        'data.Add(New String() {})

        If I <= data.Count - 1 Then
            k = data(I)(1)  'assign the vals appropriately here

            Form1.TxtBoxDigitGroupLength.Text = data(I)(3)
            Form1.TxtBoxDigitGroupLength.Focus()    'THIS LINE IS REQUIRED ORELSE DigitGroupLength WILL NOT BE UPDATED
            Form1.TxtBoxMaxLength.Text = data(I)(2)
            Form1.TxtBoxMaxLength.Focus()       'THIS LINE AND THE LINE BELOW NEXT IS REQUIRED ORELSE _MaxLength WILL NOT BE UPDATED
            Form1.TxtBoxNum1.Text = data(I)(0)
            Form1.Button1.Focus()            'THIS LINE IS REQUIRED ORELSE _MaxLength WILL NOT BE UPDATED
            Form1.TxtBoxNum2.Text = data(I)(1)
            If (I = data.Count - 1) Then
                finish = True
            End If
        End If
        Form1.lblResultComp.Text = ""
        Form1.TxtBoxResult.Text = ""
        Form1.TxtBoxResultCPU.Text = ""
    End Sub
End Class
