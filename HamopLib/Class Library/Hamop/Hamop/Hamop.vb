Imports System.Text.RegularExpressions
Public Class Hamop
    Inherits System.Object
    Private Shared _MaxLength As Integer = 38
    Private Shared _MaxLengthIO As Integer = 28
    Private Shared _MxLenModifier As Integer = 10
    Private Shared ConstantsLen As Integer = 3000
    Private Shared DGLen As Integer = 10
    Private Shared singleDgtAdd As Boolean = False
    Private Shared roundIT As Boolean = True
    Private Shared MTable(9, 9) As Integer
    Private Shared MTableU(9, 9) As Integer
    Private Shared MTableT(9, 9) As Integer


    Public Sub New()        'Initialiser of Hamop
        For I = 0 To 9
            For j = 0 To 9
                MTable(I, j) = I * j
                MTableU(I, j) = MTable(I, j) Mod 10
                MTableT(I, j) = MTable(I, j) \ 10
            Next
        Next
    End Sub

#Region "VLnums"
    Public Class VLnum
        Private Sign As Char
        Private DPval As String
        Private IPval As String
        Private DPlen As Integer
        Private IPlen As Integer
        Private ReturnedFrm As String
        Public Sub New()    'Initialiser of VLnum
            IPval = "0"     'WHEN AN OBJECT IS INSTANTIATED ITS VALUE IS SET TO "+0" 
            IPlen = 1
            DPval = ""
            DPlen = 0
            Sign = "+"
            ReturnedFrm = "o"   '"o" stands for operater
        End Sub
#Region "VLnum Properties"
        Public Property Value As String     'HERE RETURNED VALUE IS SIZED TO _MaxLengthIO & ROUNDED
            Get
                Dim StrH As String
                StrH = ""
                If Hamop.roundIT = True And (IPlen + DPlen > _MaxLengthIO) Then
                    Call Hamop.Round(IPval, IPlen, DPval, DPlen, Sign, _MaxLengthIO)
                Else
                    Call Hamop.StrArranger(IPval, IPlen, DPval, DPlen, Sign)  '_MaxLengthIO is used internally
                End If
                If ReturnedFrm = "f" Then Call VLnMath.FunctionsDpValCorrecter(Me) '_MaxLengthIO is not used in FunctionsDpValCorrecter instead
                Call Hamop.VG(StrH, Me.IPval, Me.DPval, Me.IPlen, Me.DPlen, Me.Sign)    ' _MaxLengthHere var is used in FunctionsDpValCorrecter
                Return StrH
            End Get
            Set(ByVal value As String)
                Call Hamop.VSetter(value, IPval, DPval, IPlen, DPlen, Sign)
            End Set
        End Property
#End Region
        Public Shared Operator +(ByVal vlnum1 As VLnum, ByVal vlnum2 As VLnum) As VLnum
            Dim vlnumR As New Hamop.VLnum
            Call Hamop.adder(vlnum1.IPval, vlnum1.IPlen, vlnum1.DPval, vlnum1.DPlen, vlnum1.Sign, _
                             vlnum2.IPval, vlnum2.IPlen, vlnum2.DPval, vlnum2.DPlen, vlnum2.Sign, _
                             vlnumR.IPval, vlnumR.IPlen, vlnumR.DPval, vlnumR.DPlen, vlnumR.Sign)
            Return vlnumR
        End Operator
        Public Shared Operator -(ByVal vlnum1 As VLnum) As VLnum
            Dim vlnumR As New Hamop.VLnum
            Call Hamop.Negater(vlnum1.IPval, vlnum1.IPlen, vlnum1.DPval, vlnum1.DPlen, vlnum1.Sign, _
                               vlnumR.IPval, vlnumR.IPlen, vlnumR.DPval, vlnumR.DPlen, vlnumR.Sign)
            Return vlnumR
        End Operator
        Public Shared Operator -(ByVal vlnum1 As VLnum, ByVal vlnum2 As VLnum) As VLnum
            Dim vlnumR As New Hamop.VLnum
            vlnumR = vlnum1 + (-vlnum2)
            Return vlnumR
        End Operator
        Public Shared Operator *(ByVal vlnum1 As VLnum, ByVal vlnum2 As VLnum) As VLnum
            Dim vlnumR As New Hamop.VLnum
            Call Hamop.multiplier(vlnum1.IPval, vlnum1.IPlen, vlnum1.DPval, vlnum1.DPlen, vlnum1.Sign, _
                                  vlnum2.IPval, vlnum2.IPlen, vlnum2.DPval, vlnum2.DPlen, vlnum2.Sign, _
                                  vlnumR.IPval, vlnumR.IPlen, vlnumR.DPval, vlnumR.DPlen, vlnumR.Sign)
            Return vlnumR
        End Operator
        Public Shared Operator /(ByVal Vlnum1 As VLnum, ByVal Vlnum2 As VLnum) As VLnum
            Dim vlnumR As New Hamop.VLnum
            Call Hamop.divider(Vlnum1.IPval, Vlnum1.IPlen, Vlnum1.DPval, Vlnum1.DPlen, Vlnum1.Sign, _
                               Vlnum2.IPval, Vlnum2.IPlen, Vlnum2.DPval, Vlnum2.DPlen, Vlnum2.Sign, _
                               vlnumR.IPval, vlnumR.IPlen, vlnumR.DPval, vlnumR.DPlen, vlnumR.Sign)
            Return vlnumR
        End Operator

        Public Shared Operator =(ByVal VlNum1 As VLnum, ByVal VlNum2 As VLnum) As Boolean
            Return Hamop.EQ(VlNum1, VlNum2)
        End Operator
        Public Shared Operator <>(ByVal VlNum1 As VLnum, ByVal VlNum2 As VLnum) As Boolean
            Return Hamop.NEQ(VlNum1, VlNum2)
        End Operator
        Public Shared Operator >(ByVal VlNum1 As VLnum, ByVal VlNum2 As VLnum) As Boolean
            Return Hamop.GT(VlNum1, VlNum2)
        End Operator
        Public Shared Operator >=(ByVal VlNum1 As VLnum, ByVal VlNum2 As VLnum) As Boolean
            Return Hamop.GE(VlNum1, VlNum2)
        End Operator
        Public Shared Operator <(ByVal VlNum1 As VLnum, ByVal VlNum2 As VLnum) As Boolean
            Return Hamop.LT(VlNum1, VlNum2)
        End Operator
        Public Shared Operator <=(ByVal VlNum1 As VLnum, ByVal VlNum2 As VLnum) As Boolean
            Return Hamop.LE(VlNum1, VlNum2)
        End Operator
        Public Shared Operator ^(ByVal X As VLnum, ByVal Y As VLnum) As VLnum
            Return VLnMath.xTOthePOWERy(X, Y)
        End Operator
        Public Shared Operator Mod(ByVal VlNum1 As VLnum, ByVal VlNum2 As VLnum) As VLnum
            Dim VlnumRet As New Hamop.VLnum
            Call Hamop.Remainder(VlNum1.IPval, VlNum1.IPlen, VlNum1.DPval, VlNum1.DPlen, VlNum1.Sign, _
                                 VlNum2.IPval, VlNum2.IPlen, VlNum2.DPval, VlNum2.DPlen, VlNum2.Sign, _
                                 VlnumRet.IPval, VlnumRet.IPlen, VlnumRet.DPval, VlnumRet.DPlen, VlnumRet.Sign)
            Return VlnumRet
        End Operator
        Public Shared Sub ValUpdater(ByRef vlnum1 As VLnum, Optional ByVal signh As String = "N", Optional ByVal IPv As String = "N", _
                                     Optional ByRef DPv As String = "N", Optional ByRef ReturnedfrmH As String = "N")
            If IPv <> "N" Then
                If IPv = "" Then
                    vlnum1.IPval = "0"
                    vlnum1.IPlen = 1
                Else
                    vlnum1.IPval = IPv
                    vlnum1.IPlen = IPv.Length
                End If
            End If
            If DPv <> "N" Then
                vlnum1.DPval = DPv
                vlnum1.DPlen = DPv.Length
            End If
            If signh <> "N" Then vlnum1.Sign = signh
            If ReturnedfrmH <> "N" Then vlnum1.ReturnedFrm = ReturnedfrmH
        End Sub
        Public Shared Sub ValRetriever(ByRef vlnum1 As VLnum, ByRef IPvalH As String, ByRef DPvalH As String, ByRef IPlenH As Integer, _
                                        ByRef DPlenH As Integer, ByRef SignH As Char)
            IPvalH = vlnum1.IPval
            IPlenH = vlnum1.IPlen
            DPvalH = vlnum1.DPval
            DPlenH = vlnum1.DPlen
            SignH = vlnum1.Sign
        End Sub
        Public Shared Sub RoundIt(ByRef vlnum1 As VLnum)
            Call Round(vlnum1.IPval, vlnum1.IPlen, vlnum1.DPval, vlnum1.DPlen, vlnum1.Sign, _MaxLength)
        End Sub
    End Class
#End Region


#Region "HAMOP Library class"
#Region "Hamop Properties"
    Public Property MaximumLength As Integer
        'HERE VALUE OF _MaxLength & _MaxLengthIO ARE SET. _MaxLength IS THE PRECISION AT AT WHICH CALCULATIONS ARE DONE INTERNALLY
        '_MaxLengthIO IS THE PRECISION AT WHICH VALUES ARE ACCEPTED AND RETURNED & IS LESS THAN _MaxLength BY _MxLenModifier
        Set(value As Integer)
            If value > 0 And value < 20 Then
                _MaxLength = 20 + _MxLenModifier
                _MaxLengthIO = 20
            ElseIf value > 19 Then
                _MaxLength = value + _MxLenModifier
                _MaxLengthIO = value
                'IF _MaxLength IS > 3000 INTERNAL CONSTANTS WILL BE UPDATED.
                If _MaxLength > ConstantsLen Then Call VLnMath.UpdateConstants()
            Else
                MsgBox("Enter a Proper value for Maximum Length", vbOKOnly, "Maximum Length Should be Positive & Greater Than Zero")
            End If
        End Set
        Get
            Return _MaxLengthIO
        End Get
    End Property
    Public Property DigitGroupLength As Integer
        'HERE THE NO. OF DIGITS THAT ARE PROCESSED TOGETHER IN ADDITION AND SUBTRACTION IS SET. IT CAN BE FROM 1 TO 18
        Set(value As Integer)
            If value = 1 Then
                singleDgtAdd = True
            ElseIf value > 1 And value < 19 Then
                singleDgtAdd = False
                DGLen = value
            Else
                MsgBox("Enter a Proper value for Digit Group Length", vbOKOnly, "Digit Group Length Should be +ve & > 0 & < 19")
            End If
        End Set
        Get
            Return DGLen
        End Get
    End Property

#End Region
    Private Shared Sub VSetter(ByRef _Value As String, ByRef IPvalH As String, ByRef DPvalH As String, ByRef IPlenH As Integer, ByRef DPlenH As Integer, _
                               ByRef SignH As Char)
        Dim regmatch As Boolean   'THE BELOW RegEx CAN BE USED FOR STRING VALIDATION.
        Dim pattern As String = "^(?:(?:\+?0(?:\.[0-9]*[1-9])?|-(?:0\.[0-9]*[1-9]|[1-9][0-9]*(?:\.[0-9]*[1-9])?)|\+?[1-9][0-9]*(?:\.[0-9]*[1-9])?)(?:[eE][-+]?[1-9][0-9]*)?)$"
        regmatch = Regex.IsMatch(_Value.Trim(), pattern)
        'If regmatch = True Then
        '    Form1.LblRegex.Text = "Accepted"    'REMOVE THIS IF BLOCK BEFORE CONVERTING TO DLL
        'Else
        '    Form1.LblRegex.Text = "Rejected"
        'End If
        If _Value Is Nothing Or String.IsNullOrWhiteSpace(_Value) Then
            MsgBox("Value supplied is not a valid number", vbOKOnly, "VLNum value is not set.  it is zero!")
            Exit Sub
        Else
            Call Hamop.NumGenerator(_Value, IPvalH, DPvalH, IPlenH, DPlenH, SignH)
        End If

    End Sub
    Private Shared Sub VG(ByRef _Str As String, ByRef IPvalH As String, ByRef DPvalH As String, ByRef IPlenH As Integer, ByRef DPlenH As Integer, _
                               ByRef SignH As Char)
        If (IPvalH = "0") And (DPlenH = 0 Or DPvalH = "0") Then 'DPval can be "" ie DPlen = 0
            _Str = "+0"
        ElseIf DPlenH = 0 Or DPvalH = "0" Then
            _Str = SignH & IPvalH
        ElseIf IPvalH = "0" Then
            _Str = SignH & "0" & "." & DPvalH
        Else
            _Str = SignH & IPvalH & "." & DPvalH
        End If
        If _Str(0) = "+" Then _Str = _Str.Substring(1) 'IF THE NO IS +VE SIGN IS NOT RETURNED 
    End Sub
    Private Shared Sub NumGenerator(ByRef _Value As String, ByRef IPvalH As String, ByRef DPvalH As String, ByRef IPlenH As Integer, ByRef DPlenH As Integer, _
                                   ByRef SignH As Char)
        Dim StrChk() As String
        Dim Fchar, Esign, VLNSign As Char
        Dim Estring, Istring, Dstring, IDstring, ZeroStr, _ValNew As String
        Dim DpIndex, ELocation, LdZeros, TrZeros, IDlen As Integer
        Dim DpIndexP1, EXPval As Integer
        'VARIABLE INITIALISATION
        _ValNew = ""

        'BELOW FOR LOOP ELIMINATES THE SPACES IN THE STRING
        StrChk = _Value.Split(" ")
        For I = 0 To StrChk.GetUpperBound(0)
            If StrChk(I) = Nothing Or StrChk(I) = "" Then
            Else
                _ValNew = _ValNew & StrChk(I)
            End If
        Next
        _Value = _ValNew

        'SEPARATE THE SIGN
        Fchar = _Value(0)
        If (Fchar = "-") Or (Fchar = "+") Then
            VLNSign = Fchar
            _Value = _Value.Substring(1)
        Else
            VLNSign = "+"
        End If
        DpIndex = _Value.IndexOf(".")
        ELocation = _Value.IndexOf("E")

        'SPLIT THE _Value STRING INTO ISTRING,DSTRING,ESTRING
        If DpIndex > -1 Then
            Istring = _Value.Substring(0, DpIndex)
            DpIndexP1 = DpIndex + 1
            If ELocation > -1 Then
                Dstring = _Value.Substring(DpIndexP1, ELocation - DpIndexP1)
                Estring = _Value.Substring(ELocation + 1)
            Else
                Dstring = _Value.Substring(DpIndexP1)
                Estring = ""
            End If
        Else
            If ELocation > -1 Then
                Istring = _Value.Substring(0, ELocation)
                Dstring = ""
                Estring = _Value.Substring(ELocation + 1)
            Else
                Istring = _Value.Substring(0)
                Dstring = ""
                Estring = ""
            End If
        End If

        'SEPARATE THE SIGN OF ESTRING
        If Estring <> "" Then
            Esign = Estring(0)
            If Esign = "-" Or Esign = "+" Then
                Estring = Estring.Substring(1)
            Else
                Esign = "+"
            End If
        End If
        If (Istring <> "") Then Call StrCHECKING(Istring, 0, Istring.Length - 1, 1) 'removes leading zeros
        If (Dstring <> "") Then Call StrCHECKING(Dstring, Dstring.Length - 1, 0, -1) 'removes trailing zeros
        If (Estring <> "") Then Call StrCHECKING(Estring, 0, Estring.Length - 1, 1) 'removes leading zeros
        'IF Estring IS PRESENT THEN ADJUSTS THE LOCATION OF DECIMAL POINT BELOW
        DpIndex = Istring.Length ' TO CORECT THE DpIndex IF THE Istring CONTAINED LEADING ZEROS
        If ELocation > 0 Then
            EXPval = CInt(Val(Esign & Estring))
            DpIndex = DpIndex + EXPval
            IDstring = Istring & Dstring
            IDlen = IDstring.Length
            If DpIndex >= 0 And DpIndex <= IDlen Then
                Istring = IDstring.Substring(0, DpIndex)
                Dstring = IDstring.Substring(DpIndex)
            End If
            If DpIndex < 0 Then     'WHICH MEANS NO Istring IT SHOULD NOT BE SET TO "" BUT TO "0"     
                LdZeros = Math.Abs(Istring.Length + EXPval)
                ZeroStr = New String("0", LdZeros)
                Istring = "0"                           'IT SHOULD NOT BE SET TO "" BUT TO "0"
                Dstring = ZeroStr & IDstring            'Call StrCHECKING WHEN THE DP IS SHIFTED TO LEFT (INITIALLY DSTRING ="" ) LAST DIGITS MAY BE ZEROS
            End If
            If DpIndex > IDlen Then
                TrZeros = Math.Abs(EXPval - Dstring.Length)
                ZeroStr = New String("0", TrZeros)
                Istring = IDstring & ZeroStr
                Dstring = ""
            End If
        End If
        Call Trimmer(Istring, Dstring)
        If (Istring <> "") Then Call StrCHECKING(Istring, 0, Istring.Length - 1, 1) 'TO REMOVE LEADING ZEROS AFTER TRIMMING
        If (Dstring <> "") Then Call StrCHECKING(Dstring, Dstring.Length - 1, 0, -1) 'TO REMOVE TRAILING ZEROS AFTER TRIMMING

        Call SetVLN(Istring, Dstring, VLNSign, IPvalH, IPlenH, DPvalH, DPlenH, SignH)
    End Sub
    Private Shared Sub StrCHECKING(ByRef StrH As String, ByRef BiginI As Integer, ByRef EndI As Integer, ByRef StepI As Integer)
        Dim V, ZerosCount, LdZchkd, LdTrZeros As Integer
        'CHECKS THE STRING FOR ANY THING OTHER THAN DIGITS IN THE FOR LOOP BELOW 
        For I = BiginI To EndI Step StepI
            V = Asc(StrH(I))
            If V > 47 And V < 58 Then        'digits 0 to 9  asc 48 to 57
                If V = 48 Then
                    ZerosCount += 1
                Else
                    If LdZchkd = 0 Then
                        LdZchkd = 1
                        LdTrZeros = ZerosCount
                    End If
                End If
            Else
                MsgBox("The String Contains Illegal Characters   " & StrH, vbOKOnly, "VLNum String is invalid")
                Throw New InvalidOperationException("An error occured in the DLL")
            End If
        Next
        'LEADING ZEROS ARE REMOVED BELOW    'IT BECOMES TRAILING ZEROS WHEN DSTRING IS CHECKED BECAUSE CHECKING IS FROM LAST TO FIRST IN THIS CASE
        If ZerosCount = StrH.Length Then
            StrH = "0"
        Else
            If LdTrZeros > 0 Then
                If StepI = 1 Then
                    StrH = StrH.Substring(LdTrZeros)
                Else
                    StrH = StrH.Substring(0, StrH.Length - LdTrZeros)
                End If
            End If
        End If

    End Sub
    Private Shared Sub Trimmer(ByRef Istring As String, ByRef Dstring As String)
        'IF THE NO. OF DIGITS IS MORE THAN _MaxLengthIO THEN IT IS TRIMMED TO _MaxLengthIO BELOW 
        Dim ExsLen, IStrLen As Integer
        Dim FullStr, IDstr As String
        ExsLen = Istring.Length + Dstring.Length - _MaxLengthIO   'EVEN IF Istring = "0" IT IS ALSO ADDED IN CALCULATING ExsLen.
        If ExsLen > 0 Then
            IDstr = Istring & Dstring
            IStrLen = Istring.Length
            FullStr = IDstr.Substring(0, _MaxLengthIO)
            If FullStr.Length > IStrLen Then
                Istring = FullStr.Substring(0, IStrLen)
                Dstring = FullStr.Substring(IStrLen)
            Else
                Istring = FullStr
                Dstring = ""
            End If
        End If
    End Sub
    Private Shared Sub CallerStrCHK(ByRef IPS As String, ByRef DPS As String)
        If (IPS(0) = "0" And IPS.Length > 1) Then Call StrCHECKING(IPS, 0, IPS.Length - 1, 1)
        If (DPS <> "") Then
            If (DPS(DPS.Length - 1) = "0") Then
                Call StrCHECKING(DPS, DPS.Length - 1, 0, -1)
            End If
        End If
    End Sub
    Private Shared Sub StrArranger(ByRef IPval As String, ByRef IPLen As Integer, ByRef DPval As String, ByRef DPLen As Integer, _
                                    ByRef SignR As String)
        Dim IDstr As String
        IDstr = IPval & DPval
        If IPLen > _MaxLengthIO Then
            'throw exception
            MsgBox("OVERFLOW OCCURED", vbOKOnly, "Overflow occured in the Result")
            Throw New OverflowException("Over flow Occured !!!")
        Else
            If IDstr.Length > _MaxLengthIO Then
                DPval = IDstr.Substring(IPLen, (_MaxLengthIO + 1 - IPLen))
            Else
                DPval = IDstr.Substring(IPLen)
            End If
            DPLen = DPval.Length
            If (DPval <> "") Then
                If (DPval(DPLen - 1) = "0") Then
                    Call StrCHECKING(DPval, DPLen - 1, 0, -1)
                End If
            End If
            DPLen = DPval.Length
        End If
    End Sub
    Private Shared Sub SimpleStrArranger(ByRef IPval As String, ByRef IPLen As Integer, ByRef DPval As String, ByRef DPLen As Integer, _
                                ByRef SignR As String)
        'THIS SUB WILL LIMIT IPLen + DPLen AROUND _MaxLength
        If (IPLen <= _MaxLength) And ((IPLen + DPLen) > (_MaxLength + 2)) Then
            DPval = DPval.Substring(0, (_MaxLength + 2 - IPLen))
            DPLen = DPval.Length
        End If
    End Sub
#End Region

#Region "BasicOperators +,-,x,/"

    Private Shared Sub adder(ByVal IPval1 As String, ByVal IPlen1 As Integer, ByVal DPval1 As String, ByVal DPlen1 As Integer, ByVal Sign1 As Char, _
                             ByVal IPval2 As String, ByVal IPlen2 As Integer, ByVal DPval2 As String, ByVal DPlen2 As Integer, ByVal Sign2 As Char, _
                             ByRef IPvalR As String, ByRef IPlenR As Integer, ByRef DPvalR As String, ByRef DPlenR As Integer, ByRef SignR As Char)
        Dim IPLdiff, DPLdiff, N1len, maxIPlen As Integer
        Dim Izeros, Dzeros, IP1, IP2, DP1, DP2, IPS, DPS, N1, N2, Ns As String
        'VARIABLE INITIALISATION
        IPS = ""
        DPS = ""
        Ns = ""

        IPLdiff = IPlen1 - IPlen2
        DPLdiff = DPlen1 - DPlen2
        Izeros = New String("0", Math.Abs(IPLdiff))
        Dzeros = New String("0", Math.Abs(DPLdiff))
        'STRING LENGTH EQUALISATION
        If IPlen1 > IPlen2 Then
            IP1 = IPval1
            IP2 = Izeros & IPval2
        ElseIf IPlen1 < IPlen2 Then
            IP1 = Izeros & IPval1
            IP2 = IPval2
        Else
            IP1 = IPval1
            IP2 = IPval2
        End If
        If DPlen1 > DPlen2 Then
            DP1 = DPval1
            DP2 = DPval2 & Dzeros
        ElseIf DPlen1 < DPlen2 Then
            DP1 = DPval1 & Dzeros
            DP2 = DPval2
        Else
            DP1 = DPval1
            DP2 = DPval2
        End If
        N1 = IP1 & DP1
        N2 = IP2 & DP2
        N1len = N1.Length
        maxIPlen = IP1.Length

        If Sign1 = Sign2 Then
            'ADDITION
            If singleDgtAdd = True Then
                Call SingleDigitAdd(N1len, N1, N2, Ns)
            Else
                Call MultiDigitAdd(N1len, N1, N2, Ns)
            End If
            If Ns.Length > N1len Then maxIPlen += 1
            IPS = Ns.Substring(0, maxIPlen)
            DPS = Ns.Substring(maxIPlen)
            Call CallerStrCHK(IPS, DPS)
            Call SetVLN(IPS, DPS, Sign1, IPvalR, IPlenR, DPvalR, DPlenR, SignR)
        Else
            'SUBTRACTION
            Call Compli(N2)
            If singleDgtAdd = True Then
                Call SingleDigitAdd(N1len, N1, N2, Ns)
            Else
                Call MultiDigitAdd(N1len, N1, N2, Ns)
            End If
            If Ns.Length > N1len Then
                Call ADDone(Ns)     '
                IPS = Ns.Substring(0, maxIPlen)
                DPS = Ns.Substring(maxIPlen)
                Call CallerStrCHK(IPS, DPS)
                Call SetVLN(IPS, DPS, Sign1, IPvalR, IPlenR, DPvalR, DPlenR, SignR)
            Else
                Call Compli(Ns)
                IPS = Ns.Substring(0, maxIPlen)
                DPS = Ns.Substring(maxIPlen)
                Call CallerStrCHK(IPS, DPS)
                Call SetVLN(IPS, DPS, Sign2, IPvalR, IPlenR, DPvalR, DPlenR, SignR)
            End If
        End If
        Call SimpleStrArranger(IPvalR, IPlenR, DPvalR, DPlenR, SignR)
    End Sub
    Private Shared Sub multiplier(ByVal IPval1 As String, ByVal IPlen1 As Integer, ByVal DPval1 As String, ByVal DPlen1 As Integer, ByVal Sign1 As Char, _
                                  ByVal IPval2 As String, ByVal IPlen2 As Integer, ByVal DPval2 As String, ByVal DPlen2 As Integer, ByVal Sign2 As Char, _
                                  ByRef IPvalR As String, ByRef IPlenR As Integer, ByRef DPvalR As String, ByRef DPlenR As Integer, ByRef SignR As Char)
        Dim SignP As Char
        Dim Ls, Ss, ZeroStr, SumStr, Dgt, Prod, IPpr, DPpr As String
        Dim LsLen, SsLen, TotLen1, TotLen2, lenLS, Sum, SumStrLen, SumStrLenM1, Carry, ProdLen, DPlen, ProdLeNmDPlen As Integer
        Dim Kles, A, B, X, Y As Integer
        Prod = ""
        'VARIABLE INITIALISATION
        IPpr = ""
        DPpr = ""

        'RETURNS ZERO IF ANY ONE NO. IS ZERO
        If (IPval1 = "0" And DPval1 = "") Or (IPval2 = "0" And DPval2 = "") Then
            IPpr = "0"
            DPpr = ""
            SignP = "+"
            Call SetVLN(IPpr, DPpr, SignP, IPvalR, IPlenR, DPvalR, DPlenR, SignR)
            Exit Sub
        End If
        'SELECTS LONGER STRING, AS LS AND ADDS Ss NO.OF ZEROS INFRONT OF IT  (eg: 000Ls X Ss)
        TotLen1 = IPlen1 + DPlen1
        TotLen2 = IPlen2 + DPlen2
        If TotLen1 > TotLen2 Then
            Ls = IPval1 & DPval1
            LsLen = TotLen1
            Ss = IPval2 & DPval2
            SsLen = TotLen2
        Else
            Ls = IPval2 & DPval2
            LsLen = TotLen2
            Ss = IPval1 & DPval1
            SsLen = TotLen1
        End If
        ZeroStr = New String("0", SsLen)
        Ls = ZeroStr & Ls
        lenLS = LsLen + SsLen - 1
        'MULTIPLICATION
        For K = lenLS To 1 Step -1
            A = K
            If K < SsLen Then A = SsLen : Kles = SsLen - K
            For L = SsLen - 1 To 0 Step -1
                B = L - Kles
                X = CInt(Val(Ls(A)))
                Y = CInt(Val(Ss(B)))
                Sum += MTable(X, Y)
                A += 1
                If A > lenLS Or B < 1 Then Exit For
            Next
            Sum += Carry
            SumStr = Sum.ToString
            Sum = 0
            Carry = 0
            SumStrLen = SumStr.Length
            If SumStrLen > 1 Then
                If K = 1 Then
                    Dgt = SumStr
                Else
                    SumStrLenM1 = SumStrLen - 1
                    Dgt = SumStr.Substring(SumStrLenM1)
                    Carry = CInt(Val(SumStr.Substring(0, SumStrLenM1)))
                End If
            Else
                Dgt = SumStr
                Carry = 0
            End If
            Prod = Dgt & Prod
        Next
        'SEPARATES IP AND DP
        ProdLen = Prod.Length
        DPlen = DPlen1 + DPlen2
        ProdLeNmDPlen = ProdLen - DPlen
        IPpr = Prod.Substring(0, ProdLeNmDPlen)
        DPpr = Prod.Substring(ProdLeNmDPlen)

        If (IPpr(0) = "0" And ProdLeNmDPlen > 1) Then
            Call StrCHECKING(IPpr, 0, ProdLeNmDPlen - 1, 1)
            ProdLeNmDPlen = IPpr.Length
        End If


        'SETS SIGN
        If Sign1 = Sign2 Then
            SignP = "+"
        Else
            SignP = "-"
        End If

        Call SetVLN(IPpr, DPpr, SignP, IPvalR, IPlenR, DPvalR, DPlenR, SignR)
        Call SimpleStrArranger(IPvalR, IPlenR, DPvalR, DPlenR, SignR)
    End Sub
    Private Shared Sub divider(ByVal IPval1 As String, ByVal IPlen1 As Integer, ByVal DPval1 As String, ByVal DPlen1 As Integer, ByVal Sign1 As Char, _
                               ByVal IPval2 As String, ByVal IPlen2 As Integer, ByVal DPval2 As String, ByVal DPlen2 As Integer, ByVal Sign2 As Char, _
                               ByRef IPvalR As String, ByRef IPlenR As Integer, ByRef DPvalR As String, ByRef DPlenR As Integer, ByRef SignR As Char)
        '75869545/45845 ie. A LARGER INTEGER IS DIVIDED BY A SMALLER INTEGER. DECIMAL POINT (DP) WILL BE AT (NO.OF DGTS IN NUMER - NO.OF DGTS IN DENOMI LESS ONE)
        'IF THE FIRST DGT OF DIVIDENT IS LESS THAN FIRST DGT OG DIVISOR THEN 0 IS TAKEN AS THE FIRST DIGIT OF QUOTIENT, TAKING IT THIS WAY WE WILL GET THE DP AT
        'THE ABOVE FIXED POSITION (OTHERWISE IT IS DIFFICULT TO LOCATE DP) THE REST (OTHER CASES LIKE SINGLE DIGIT DIVISOR, NOS. WITH ONLY DECIMAL PART ETC) ARE 
        'EXTENSITIONS OF THE ABOVE)
        Dim N1IPlen, N2Len, DsD1, DsD2, PDiv, Qdgt, ii, NTP, UTP, Wno, DsDC, DsDCp1, DsDV, DsDVnxt, QdgtV, QIPlen, QDPlen, QdgtZcount, YesExit As Integer
        Dim N1, N2, QdgtStr, QIPstr, QDPstr As String
        Dim SignP As Char
        Dim PDivAll(_MaxLength + 10) As Integer
        'VARIABLE INITIALISATION
        QDPstr = ""

        If (IPval2 = "0") And (DPlen2 = 0 Or DPval2 = "0") Then
            MsgBox("Division By Zero is Undefined", vbOKOnly, "Division By Zero Error")     'comment it later
            Throw New DivideByZeroException("An error occured in the DLL")
            'if you want to return infinity or NaN it is possible.
        End If
        If (IPval1 = "0") And (DPlen1 = 0 Or DPval1 = "0") Then
            IPvalR = "0" : IPlenR = 1 : DPvalR = "" : DPlenR = 0 : SignR = "+"
            Return
        End If

        N1IPlen = IPlen1 + DPlen2       'WHEN Vlnum2.DPlen IS ADDED THE DENOMINATER BECOMES INTEGER AND THEN LINE '*** BELOW CALCULATES THE DPLOCATION
        N1 = IPval1 & DPval1
        N2 = IPval2 & DPval2
        'IF N2 CONTAINS ONLY DECIMALPART ie. (IPVAL=0) THEN THE LEADING ZEROS ARE REMOVED IN THE IF BLOCK BELOW
        If IPval2 = "0" Then
            StrCHECKING(N2, 0, N2.Length - 1, 1)
        End If
        'IF THE DIVISOR (N2) IS SINGLE DIGIT A "0" IS ADDED AT THE END OF N2 TO MAKE IT A TWO DIGIT NO. AND THE NUMERATOR IPlen IS INCREASED BY 1
        If N2.Length = 1 Then N2 = N2 & "0" : N1IPlen += 1
        N2Len = N2.Length
        DsD1 = CInt(Val(N2(0)))   'val returns double  & cint converts to int with out rounding as dp will not be there
        DsD2 = CInt(Val(N2(1)))
        PDiv = CInt(Val(N1(0)))
        Qdgt = Math.Floor(PDiv / DsD1)
        QdgtStr = Qdgt.ToString
        PDivAll(1) = PDiv
        ii = 1
        'LINE BELOW CALCULATES THE LOCATION OF DP
        QIPlen = N1IPlen - (N2Len - 1)  '                                 '***
        YesExit = Math.Max(QIPlen, N1.Length)

        'DIVISION   'pdiv,ntp,wno,utp,qdgt
        Do While ii < _MaxLength + 10
            UTP = 0
            'CALCULATES NTP AND Wno            
            NTP = MTable(DsD1, Qdgt) + MTableT(DsD2, Qdgt)
            If ii >= N1.Length Then
                Wno = CInt(Val((PDiv - NTP).ToString & "0"))
            Else
                Wno = CInt(Val((PDiv - NTP).ToString & N1(ii)))
            End If
            DsDC = 1
            DsDCp1 = 2
            'THE FOR LOOP BELOW CALCULATES UTP
            For j = ii To 1 Step -1
                DsDV = CInt(Val(N2(DsDC)))
                QdgtV = CInt(Val(QdgtStr(j - 1)))
                If DsDCp1 < N2Len Then
                    DsDVnxt = CInt(Val(N2(DsDCp1)))
                Else
                    DsDVnxt = 0
                End If
                UTP += MTableU(DsDV, QdgtV) + MTableT(DsDVnxt, QdgtV)
                If DsDCp1 = N2Len Then Exit For
                DsDC = DsDCp1
                DsDCp1 += 1
            Next
            'IF REQUIRED Qdgt AND QdgtStr ARE CORRECTED BELOW
            If Wno - UTP < 0 Then
                If Qdgt = 0 Then
                    For j = ii To 1 Step -1
                        If CInt(Val(QdgtStr(j - 1))) > 0 Then
                            ii = j
                            Qdgt = CInt(Val(QdgtStr(j - 1))) - 1
                            QdgtStr = QdgtStr.Substring(0, j - 1) & Qdgt.ToString
                            PDiv = PDivAll(j)
                            Exit For
                        End If
                    Next
                Else
                    Qdgt -= 1
                    QdgtStr = QdgtStr.Substring(0, ii - 1) & Qdgt.ToString
                End If
                Continue Do
            End If
            'THE IF BLOCK BELOW DETERMINES WHEN TO STOP CALCULATION (IF THE CALCULATION IS IN THE DECIMAL REGION AND THE NO. OF 
            'CONTINUOUS ZEROS >= N2len AND ALL DIGITS OF DIVIDENT IS USED UP THEN STOP CALC) (PDiv,NTP,Wno,UTP,Qdgt ALL WILL BE ZEROS AFTERWARDS)
            If ii > QIPlen Then
                If Qdgt = 0 Then
                    QdgtZcount += 1
                Else
                    QdgtZcount = 0
                End If
                If QdgtZcount >= N2Len And ii >= YesExit Then
                    If Wno - UTP = 0 Then Exit Do
                End If
            End If
            'NEXT PDiv AND Qdgt ARE CALCULATED
            PDiv = Wno - UTP
            Qdgt = Math.Floor(PDiv / DsD1)
            If Qdgt > 9 Then Qdgt = 9
            QdgtStr = QdgtStr & Qdgt.ToString
            PDivAll(ii + 1) = PDiv
            ii += 1
        Loop
        'THE IF BLOCK BELOW ADDS THE PRECEEDING ZEROS IN QUOTIENT WHEN DENOM IS MUCH LARGER THAN NUMERATOR  (45/45874596854)
        If QIPlen < 1 Then
            QdgtStr = New String("0", -(QIPlen - 1)) & QdgtStr
            QIPlen = 1
        End If
        'THE IF BLOCK BELOW ADDS THE TRAILING ZEROS IN QUOTIENT WHEN DENOM IS MUCH SMALLER THAN NUMER (4582/.00002548)
        If QIPlen > QdgtStr.Length Then
            QdgtStr = QdgtStr & New String("0", QIPlen - QdgtStr.Length)
        End If
        QIPstr = QdgtStr.Substring(0, QIPlen)
        QDPstr = QdgtStr.Substring(QIPlen)
        QDPlen = QDPstr.Length

        If (QIPstr(0) = "0" And QIPlen > 1) Then
            Call StrCHECKING(QIPstr, 0, QIPlen - 1, 1)
            QIPlen = QIPstr.Length
        End If
        'SETS SIGN
        If Sign1 = Sign2 Then
            SignP = "+"
        Else
            SignP = "-"
        End If
        Call SetVLN(QIPstr, QDPstr, SignP, IPvalR, IPlenR, DPvalR, DPlenR, SignR)
        Call SimpleStrArranger(IPvalR, IPlenR, DPvalR, DPlenR, SignR)
    End Sub
    Private Shared Sub Remainder(ByVal IPval1 As String, ByVal IPlen1 As Integer, ByVal DPval1 As String, ByVal DPlen1 As Integer, ByVal Sign1 As Char, _
                                 ByVal IPval2 As String, ByVal IPlen2 As Integer, ByVal DPval2 As String, ByVal DPlen2 As Integer, ByVal Sign2 As Char, _
                                 ByRef IPvalR As String, ByRef IPlenR As Integer, ByRef DPvalR As String, ByRef DPlenR As Integer, ByRef SignR As Char)
        Dim VlnumRet, VlnumD, Vlnum1, Vlnum2 As New Hamop.VLnum
        Dim IPv, DPv, SignVln As String
        Dim IPL, DPL As Integer
        'VARIABLE INITIALISATION
        IPv = ""
        DPv = ""
        SignVln = ""

        Call VLnum.ValUpdater(Vlnum1, "+", IPval1, DPval1)
        Call VLnum.ValUpdater(Vlnum2, "+", IPval2, DPval2)
        VlnumD = Vlnum1 / Vlnum2
        Call VLnum.ValRetriever(VlnumD, IPv, DPv, IPL, DPL, SignVln)
        Call VLnum.ValUpdater(VlnumD, SignVln, IPv, "")
        VlnumRet = Vlnum1 - (Vlnum2 * VlnumD)
        Call VLnum.ValRetriever(VlnumRet, IPvalR, DPvalR, IPlenR, DPlenR, SignR)
        SignR = Sign1
    End Sub

#End Region

#Region "Operater Helper SUBs"
    Private Shared Sub SetVLN(ByRef IPs As String, ByRef DPs As String, ByRef Sign As Char, _
                              ByRef IPvalR As String, ByRef IPlenR As String, ByRef DPvalR As String, ByRef DPlenR As String, _
                              ByRef SignR As Char)
        SignR = Sign
        If IPs = "" Or IPs = "0" Then
            IPvalR = "0"
            IPlenR = 1
        Else
            IPvalR = IPs
            IPlenR = IPs.Length
        End If
        If DPs = "" Or DPs = "0" Then
            DPvalR = ""
            DPlenR = 0
        Else
            DPvalR = DPs
            DPlenR = DPs.Length
        End If
        If IPvalR = "0" And DPvalR = "" And SignR = "-" Then SignR = "+" 'THIS LINE IS NECESSARY FOR COMPARISON OPERATORS
    End Sub
    Private Shared Sub SingleDigitAdd(ByVal Istart As Integer, ByVal val1 As String, ByVal val2 As String, ByRef valR As String)
        Dim sum, carry As Integer
        For I = Istart - 1 To 0 Step -1
            sum = CInt(Val(val1(I))) + CInt(Val(val2(I))) + carry
            carry = 0
            If sum > 9 Then
                sum = sum - 10
                carry = 1
            End If
            valR = sum.ToString & valR
        Next
        If carry = 1 Then valR = carry & valR
    End Sub
    Private Shared Sub MultiDigitAdd(ByVal Istart As Integer, ByVal val1 As String, ByVal val2 As String, ByRef valR As String)
        Dim strLn, sumStr, Izeros As String
        Dim DGLenH, strBegin, carry As Integer
        Dim sum As Long
        DGLenH = DGLen

        'LONG INT CAN HANDLE 18 DIGITS WHICH IS LESS THAN 64 BITS.
        '9223372036854775807(LONG INT MAX SIGNED) IS 63 BIT SO UPTO 18 DGTS CAN BE USED
        For I = Istart To 0 Step -DGLen 'A SUB STRING OF LENGTH DGLen IS SELECTED IN EACH STEP 
            strBegin = I - DGLen
            If strBegin < 1 Then
                DGLenH = strBegin + DGLen
                strBegin = 0
            End If
            sum = CLng(val1.Substring(strBegin, DGLenH)) + CLng(val2.Substring(strBegin, DGLenH)) + carry  
            carry = 0
            sumStr = sum.ToString
            strLn = sumStr.Length
            If strLn > DGLenH Then      'TO SEPERATE THE CARRY
                sumStr = sumStr.Substring(1)
                carry = 1
            ElseIf strLn < DGLenH Then   'ie WHEN THE LEADING DIGITS OF THE SUBSTRING ARE ZEROS
                Izeros = New String("0", (DGLenH - strLn))
                sumStr = Izeros & sumStr
            End If
            valR = sumStr & valR
            If strBegin = 0 And DGLenH = DGLen Then Exit For
        Next
        If carry = 1 Then valR = carry & valR
    End Sub
    Private Shared Sub Round(ByRef IPvalR As String, ByRef IPlenR As Integer, ByRef DPvalR As String, ByRef DPlenR As Integer, ByRef SignR As Char, _
                             ByVal _MaxLengthH As Integer)
        Dim IDstr, IDS, IDSbal As String
        Dim IDlen, FdgtDiscard, sum, carry, flag As Integer
        'variable initialisation
        IDS = ""

        IDstr = IPvalR & DPvalR
        IDlen = IPlenR + DPlenR
        If IPlenR > _MaxLengthH Then
            MsgBox("OVERFLOW OCCURED", vbOKOnly, "Overflow occured in the Result")
            Throw New OverflowException("Over flow Occured !!!")
        ElseIf (IDlen > _MaxLengthH) Then
            FdgtDiscard = CInt(Val(IDstr(_MaxLengthH)))
            IDstr = IDstr.Substring(0, _MaxLengthH)
            If FdgtDiscard > 4 Then
                flag = 1
                carry = 1
                For I = IDstr.Length - 1 To 0 Step -1
                    sum = CInt(Val(IDstr(I))) + carry   'CARRY IS ADDED TO THE IDstr WHEN CARRY BECOMES "0" THE STRING GETS COPIED
                    carry = 0
                    If sum > 9 Then
                        sum = sum - 10
                        carry = 1
                    Else
                        'HERE WHEN CARRY BECOMES 0 THE REMAINING STRING GETS COPIED AND CONCATENATED & 
                        'flag BECOMES 0 ie YOU NEED NOT INCREASE IPLenR 
                        IDS = sum.ToString & IDS
                        IDSbal = IDstr.Substring(0, I)
                        IDS = IDSbal & IDS
                        flag = 0
                        Exit For
                    End If
                    IDS = sum.ToString & IDS
                Next
                If carry = 1 Then IDS = carry.ToString & IDS 'WHEN ALL DIGITS WERE 9S THE LAST CARRY SHOULD BE ADDED HERE
            Else
                IDS = IDstr
            End If
        Else
            IDS = IDstr
        End If
        If flag = 1 Then IPlenR = IPlenR + 1
        IPvalR = IDS.Substring(0, IPlenR)   'IPLenR may increase by 1  (when 999.99999 is rounded i part becomes 1000 & no. becomes 1000.0000 )
        DPvalR = IDS.Substring(IPlenR)
        DPlenR = DPvalR.Length
        'NEED TRIMMING LAST ZEROS AFTER DP 
        If (DPvalR <> "") Then
            If (DPvalR(DPvalR.Length - 1) = "0") Then
                Call StrCHECKING(DPvalR, DPvalR.Length - 1, 0, -1)
                DPlenR = DPvalR.Length      'IF DPvalR CONTAINED TRAILING ZEROS DPlenR SHOULD BE UPDATED
            End If
        End If
    End Sub
    Private Shared Sub Negater(ByRef IPvalIN As String, ByRef IPlenIN As String, ByRef DPvalIN As String, ByRef DPlenIN As String, ByRef SignIN As Char, _
                               ByRef IPvalOUT As String, ByRef IPlenOUT As String, ByRef DPvalOUT As String, ByRef DPlenOUT As String, ByRef signOUT As Char)
        IPvalOUT = IPvalIN
        IPlenOUT = IPlenIN
        DPvalOUT = DPvalIN
        DPlenOUT = DPlenIN
        If SignIN = "+" Then
            signOUT = "-"
        Else
            signOUT = "+"
        End If

    End Sub
    Private Shared Sub Compli(ByRef n2s As String)
        Dim J As Integer
        Dim n2c As String
        'variable initialisation
        n2c = ""

        For I As Integer = 0 To n2s.Length - 1
            J = 9 - CInt(Val(n2s(I)))
            n2c = n2c & J.ToString
        Next
        n2s = n2c
    End Sub
    Private Shared Sub ADDone(ByRef nS As String)
        Dim nSH, nSFinal As String            'THIS SUB IS USED IN SUBTRACTION
        Dim J, Carry, nSHlen As Integer
        'VARIABLE INITIALISATION
        nSFinal = ""

        Carry = 1
        nSH = nS.Substring(1)
        nSHlen = nSH.Length
        For I As Integer = nSHlen - 1 To 0 Step -1
            J = CInt(Val(nSH(I))) + Carry
            Carry = 0
            If J > 9 Then
                Carry = 1
                nSFinal = "0" & nSFinal
            Else
                nSFinal = J.ToString & nSFinal
            End If
            If Carry = 0 Then
                nSFinal = nSH.Substring(0, I) & nSFinal
                Exit For
            End If
        Next
        nS = nSFinal
    End Sub

#End Region
#Region "Comparison or Relational Operators =,<>,>,<,>=,<="
    Private Shared Function EQ(ByVal VlNum1 As VLnum, ByVal VlNum2 As VLnum) As Boolean
        If VlNum1.Value = VlNum2.Value Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Shared Function NEQ(ByVal VlNum1 As VLnum, ByVal VlNum2 As VLnum) As Boolean
        If VlNum1.Value <> VlNum2.Value Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Shared Function GT(ByVal VlNum1 As VLnum, ByVal VlNum2 As VLnum) As Boolean
        Dim Diff As New Hamop.VLnum
        Diff = VlNum2 - VlNum1
        If Diff.Value(0) = "-" Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Shared Function LT(ByVal VlNum1 As VLnum, ByVal VlNum2 As VLnum) As Boolean
        Dim Diff As New Hamop.VLnum
        Diff = VlNum1 - VlNum2
        If Diff.Value(0) = "-" Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Shared Function GE(ByVal VlNum1 As VLnum, ByVal VlNum2 As VLnum) As Boolean
        Dim Diff As New Hamop.VLnum
        Diff = VlNum1 - VlNum2
        If Diff.Value(0) = "-" Then
            Return False
        Else
            Return True
        End If
    End Function
    Private Shared Function LE(ByVal VlNum1 As VLnum, ByVal VlNum2 As VLnum) As Boolean
        Dim Diff As New Hamop.VLnum
        Diff = VlNum2 - VlNum1
        If Diff.Value(0) = "-" Then
            Return False
        Else
            Return True
        End If
    End Function
#End Region

#Region "VLnMath"
    Public Class VLnMath
        Public Sub New()
            'nothing here at the moment
        End Sub

#Region "Math Properties"

        Public Shared ReadOnly Property e As VLnum
            Get
                Dim vlnumRet As New Hamop.VLnum
                Dim DPv As String
                DPv = cVlNums.Constante.Substring(2, _MaxLength)
                Call VLnum.ValUpdater(vlnumRet, "+", "2", DPv)
                If Hamop.roundIT = True Then Call VLnum.RoundIt(vlnumRet)
                Return vlnumRet
            End Get
        End Property
        Public Shared ReadOnly Property PI As VLnum
            Get
                Dim VlNumRet As New VLnum
                Dim DPv As String
                'TAKE CARE NOT TO CHANGE ANY THING IN cVlNums.ConstantPI STRING
                DPv = cVlNums.ConstantPI.Substring(2, _MaxLength)   'LENGTH OF SUBSTRING IS INCREASED BY 1 FOR ROUNDING
                Call VLnum.ValUpdater(VlNumRet, "+", "3", DPv)
                If Hamop.roundIT = True Then Call VLnum.RoundIt(VlNumRet)
                Return VlNumRet
            End Get
        End Property
        Public Shared ReadOnly Property RtoD As VLnum
            Get
                Dim VlNumRet As New VLnum
                Dim DPv As String
                DPv = cVlNums.ConstantRtoD.Substring(3, _MaxLength)   'LENGTH OF SUBSTRING IS INCREASED BY 1 FOR ROUNDING
                Call VLnum.ValUpdater(VlNumRet, "+", "57", DPv)
                If Hamop.roundIT = True Then Call VLnum.RoundIt(VlNumRet)
                Return VlNumRet
            End Get
        End Property
        Public Shared ReadOnly Property DtoR As VLnum
            Get
                Dim VlNumRet As New VLnum
                Dim DPv As String
                DPv = cVlNums.ConstantDtoR.Substring(2, _MaxLength)   'LENGTH OF SUBSTRING IS INCREASED BY 1 FOR ROUNDING
                Call VLnum.ValUpdater(VlNumRet, "+", "0", DPv)
                If Hamop.roundIT = True Then Call VLnum.RoundIt(VlNumRet)
                Return VlNumRet
            End Get
        End Property
#End Region
        Public Shared Function Abs(ByVal X As VLnum) As VLnum
            Dim vlnumR As New VLnum
            Dim IPv, DPv, SgnH As String
            Dim IPl, DPl As Integer
            'Variable Initialiser
            IPv = ""
            DPv = ""
            IPl = 0
            DPl = 0
            SgnH = ""
            Call VLnum.ValRetriever(X, IPv, DPv, IPl, DPl, SgnH)
            Call VLnum.ValUpdater(vlnumR, "+", IPv, DPv)
            Return vlnumR
        End Function

        Private Shared Function HalfPI() As VLnum
            Dim VlNumRet As New VLnum
            Dim DPv As String
            DPv = cVlNums.ConstantHalfPI.Substring(2, _MaxLength)   'LENGTH OF SUBSTRING IS INCREASED BY 1 FOR ROUNDING
            Call VLnum.ValUpdater(VlNumRet, "+", "1", DPv)
            If Hamop.roundIT = True Then Call VLnum.RoundIt(VlNumRet)
            HalfPI = VlNumRet
        End Function
        Private Shared Function Tan22_5H() As VLnum
            Dim VlNumRet As New VLnum
            Dim DPv As String
            DPv = cVlNums.Tan22_5H.Substring(2, _MaxLength)   'LENGTH OF SUBSTRING IS INCREASED BY 1 FOR ROUNDING
            Call VLnum.ValUpdater(VlNumRet, "+", "0", DPv)
            If Hamop.roundIT = True Then Call VLnum.RoundIt(VlNumRet)
            Tan22_5H = VlNumRet
        End Function
        Private Shared Function LogTENe() As VLnum
            Dim vlnumRet As New Hamop.VLnum
            Dim DPv As String
            DPv = cVlNums.ConstantLogTENe.Substring(2, _MaxLength)   'LENGTH OF SUBSTRING IS INCREASED BY 1 FOR ROUNDING
            Call VLnum.ValUpdater(vlnumRet, "+", "2", DPv)
            If Hamop.roundIT = True Then Call VLnum.RoundIt(vlnumRet)
            Return vlnumRet
        End Function
#Region "Logarithmic Functions"

        Public Shared Function CALCe() As VLnum
            'e^X = 1 + x^1/1! + x^2/2! + x^3/3! +---   x = 1
            Dim sum, Term, Numi, VlNumRet As New VLnum
            Dim i As Integer = 1
            VLnum.ValUpdater(Term, "+", "1", "")
            VLnum.ValUpdater(sum, "+", "1", "")
            Do While Term.Value > "0"
                VLnum.ValUpdater(Numi, "+", i.ToString, "")
                Term = Term * (cVlNums.One / Numi)
                sum = sum + Term
                i += 1
            Loop
            Return sum
        End Function
        Public Shared Function ALoge(ByVal X As VLnum) As VLnum
            'e^X = 1 + x^1/1! + x^2/2! + x^3/3! +---    The series is convergent irrespective of the val of x, -- if x=1 then this Fn will return e 
            Dim sum, Term, Numi, VlNumRet, XH, VlnMF, VlnMlen As New VLnum
            Dim i As Integer
            i = 1

            VLnum.ValUpdater(Term, "+", "1", "")
            VLnum.ValUpdater(sum, "+", "1", "")
            Do While Term.Value <> "0"
                VLnum.ValUpdater(Numi, "+", i.ToString, "")
                Term = Term * (X / Numi)
                sum = sum + Term
                i += 1
            Loop
            VLnum.ValUpdater(sum, "N", "N", "N", "f")
            Return sum
        End Function
        Public Shared Function Loge(ByVal VlNumX As VLnum) As VLnum
            Dim VlNumXorInv, VlNumePow, VlNumR, VlNumXPart, VlNumRet, VlNumBelow, VlNumAbove As New VLnum
            Dim sum, Numi, Term, numiM1, MF, SumH As New VLnum
            Dim I, J, SignXPart, SignH As Integer
            Dim Js, JsMax As Integer
            Dim PowerIndicaterUP, PowerIndicaterDN, PIndicater As New VLnum
            Dim lnSign As String = "+"
            Dim VlnumePows(10) As Hamop.VLnum
            Dim IPv, DPv, SignVln As String
            Dim IPL, DPL As Integer
            'VARIABLE INITIALISATION
            IPv = ""
            DPv = ""
            SignVln = ""

            Select Case VlNumX
                Case Is <= cVlNums.Zero         'ELIMINATES ALL -VE NOS.
                    MsgBox(" Invalid Input ", vbOKOnly, "Loge Error")
                    Throw New OverflowException("An error occured in the DLL")
                Case Is < cVlNums.One           'NOS. < 1 ARE INVERTED SO IT BECOMES > 1, lnSign = "-"
                    VlNumXorInv = cVlNums.One / VlNumX
                    lnSign = "-"
                Case Is = cVlNums.One
                    Return cVlNums.Zero
                Case Else
                    VlNumXorInv = VlNumX
            End Select
            VlNumePow = VLnMath.e
            Call VLnum.ValUpdater(PowerIndicaterUP, "+", "1", )
            VlnumePows(0) = cVlNums.One
            VlnumePows(1) = VlNumePow
            I = 1
            Do While VlNumXorInv > VlNumePow
                I += 1
                VlNumePow = VlNumePow * VlNumePow
                If VlnumePows.GetUpperBound(0) < I Then
                    ReDim Preserve VlnumePows(I + 4)
                End If
                VlnumePows(I) = VlNumePow
                PowerIndicaterUP = PowerIndicaterUP * cVlNums.Two
            Loop
            PIndicater = PowerIndicaterUP * cVlNums.Half
            Call VLnum.ValRetriever(PIndicater, IPv, DPv, IPL, DPL, SignVln)
            Call VLnum.ValUpdater(PowerIndicaterUP, , IPv, )
            Call VLnum.ValUpdater(PowerIndicaterDN, SignVln, IPv, )
            VlNumePow = VlnumePows(I - 1)
            'IN THE FOR LOOP BELOW e^X WHICH IS JUST BELOW  VlNumXorInv IS FOUND (ie e^X+1 IS ABOVE VlNumXorInv)
            For J = I - 2 To 0 Step -1
                VlNumR = VlNumePow * VlnumePows(J)
                PIndicater = PowerIndicaterDN * cVlNums.Half
                Call VLnum.ValRetriever(PIndicater, IPv, DPv, IPL, DPL, SignVln)
                Call VLnum.ValUpdater(PowerIndicaterDN, , IPv, )
                If VlNumXorInv > VlNumR Then
                    VlNumePow = VlNumR
                    PowerIndicaterUP += PowerIndicaterDN
                Else
                End If
            Next
            'VlNumXPart IS CALCULATED IN THE IF BLOCK BELOW. VlNumXPart SHOULD BE < .5,  e^X  OR  e^X+1 IS USED
            VlNumBelow = VlNumXorInv / VlNumePow
            If VlNumBelow < cVlNums.a3BY2 Then
                VlNumXPart = VlNumBelow - cVlNums.One
                SignXPart = -1
            Else
                VlNumePow *= VlnumePows(1)
                PowerIndicaterUP = PowerIndicaterUP + cVlNums.One
                VlNumAbove = VlNumXorInv / VlNumePow
                VlNumXPart = cVlNums.One - VlNumAbove
                SignXPart = +1
            End If

            'FOR FASTER CONVERGENCE VALUE OF X SHOULD BE LESS THAN .5 (THE SERIES CONVERGES FOR VALS -1<X<=1)
            'Loge(1+X) = X -  X^2/2 +  X^3/3 - X^4/4 +-----
            'Loge(1-X) = -(X +  X^2/2 +  X^3/3 + X^4/4 +-----)
            SignH = 1
            Term = VlNumXPart
            sum = VlNumXPart
            Js = 0
            JsMax = -1
            I = 1
            Do While Term.Value > "0"
                I += 1
                SignH = SignH * SignXPart
                VLnum.ValUpdater(Numi, "+", I.ToString, "")
                numiM1 = Numi - cVlNums.One
                MF = numiM1 / Numi
                Term = Term * (VlNumXPart * MF)
                If SignH = +1 Then
                    sum = sum + Term
                Else
                    sum = sum - Term
                End If
            Loop
            If SignXPart = +1 Then sum = -sum
            SumH = PowerIndicaterUP + sum
            If lnSign = "-" Then VLnum.ValUpdater(SumH, lnSign, , )
            VLnum.ValUpdater(SumH, "N", "N", "N", "f")
            Return SumH
        End Function
        Public Shared Function ALog10(ByVal X As VLnum) As VLnum
            Dim VlNumR, VlnumRett As New VLnum
            Dim Xstr As String
            Dim IPv, DPv, SignH As String
            Dim IPl, DPl As Integer
            'VARIABLE INITIALISATION
            IPv = ""
            DPv = ""
            SignH = ""

            Call VLnum.ValRetriever(X, IPv, DPv, IPl, DPl, SignH)

            If DPv = "" Then
                'X IS AN INTEGER. THE IF BLOCK BELOW EVALUATES THE ANTI LOGARITHM (10^x)
                If SignH = "-" Then
                    Xstr = New String("0", CInt(Val(IPv)) - 1)
                    Xstr = Xstr & "1"
                    Call VLnum.ValUpdater(VlNumR, "+", "0", Xstr)
                Else
                    Xstr = New String("0", CInt(Val(IPv)))
                    Xstr = "1" & Xstr
                    Call VLnum.ValUpdater(VlNumR, "+", Xstr, "")
                End If
                Return VlNumR
            Else
                'X IS A DECIMAL VALUE. THE BELOW LINE EVALUATES THE ANTILOGARITHM (10^2.3)
                VlNumR = ALoge(X * LogTENe())
            End If
            Return VlNumR
        End Function
        Public Shared Function Log10(ByVal X As VLnum) As VLnum
            Dim VlNumR As New VLnum
            VlNumR = (Loge(X) / LogTENe())
            VLnum.ValUpdater(VlNumR, "N", "N", "N", "f")
            Return VlNumR
        End Function
        Public Shared Function xTOthePOWERy(ByVal X As VLnum, ByVal Y As VLnum) As VLnum
            Dim VlnumR, Sum As New VLnum
            Dim IPv, DPv, SignVln As String
            Dim IPL, DPL As Integer
            'VARIABLE INITIALISATION
            IPv = ""
            DPv = ""
            SignVln = ""

            Select Case Y.Value
                Case "-1"
                    Return (cVlNums.One / X)
                Case "0"
                    Return cVlNums.One
                Case "1"
                    Return X
            End Select
            Call VLnum.ValRetriever(Y, IPv, DPv, IPL, DPL, SignVln)
            If DPL = 0 Then
                Return xTOthePOWERintY(X, Y)
            End If
            VlnumR = Y * Loge(X)
            Sum = ALoge(VlnumR)
            Return Sum
        End Function
        Public Shared Function Sqrt(ByVal x As VLnum) As VLnum
            Dim vlnumR, sum As New VLnum
            Select Case x
                Case Is < cVlNums.Zero
                    MsgBox("invalid input,  number is negative   " & x.Value, vbOKOnly, "error in Sqrt( )")
                    Throw New Exception("-ve arguement in sqrt function")
                Case cVlNums.Zero
                    Return cVlNums.Zero
                Case Else
                    vlnumR = cVlNums.Half * Loge(x)
                    sum = ALoge(vlnumR)
                    Return sum
            End Select
        End Function
        Public Shared Function Floor(ByVal VLnum As VLnum) As VLnum
            Dim VLnumRet As New VLnum
            Dim IPv, DPv, SignVln As String
            Dim IPL, DPL As Integer
            'VARIABLE INITIALISATION
            IPv = ""
            DPv = ""
            SignVln = ""

            Call VLnum.ValRetriever(VLnum, IPv, DPv, IPL, DPL, SignVln)
            If SignVln = "-" Then
                Call VLnum.ValUpdater(VLnumRet, "-", IPv, )
                VLnumRet = VLnumRet + cVlNums.MOne
            Else
                Call VLnum.ValUpdater(VLnumRet, "+", IPv, )
            End If
            Return VLnumRet
        End Function
        Public Shared Function Ceiling(ByVal VLnum As VLnum) As VLnum
            Dim VLnumRet As New VLnum
            Dim IPv, DPv, SignVln As String
            Dim IPL, DPL As Integer
            'VARIABLE INITIALISATION
            IPv = ""
            DPv = ""
            SignVln = ""

            Call VLnum.ValRetriever(VLnum, IPv, DPv, IPL, DPL, SignVln)
            If SignVln = "-" Then
                Call VLnum.ValUpdater(VLnumRet, "-", IPv, )
            Else
                Call VLnum.ValUpdater(VLnumRet, "+", IPv, )
                VLnumRet = VLnumRet + cVlNums.One
            End If
            Return VLnumRet
        End Function
        Private Shared Function xTOthePOWERintY(ByVal X As VLnum, ByVal Y As VLnum) As VLnum
            Dim YH, XH, PowerIndicaterUP, PowerIndicaterDN, PIndicater, VlNumR As New VLnum
            Dim VlnumXPows(50) As Hamop.VLnum
            Dim SignNeg As Boolean = False
            Dim I As Integer
            Dim IPv, yIPv, DPv, SignVln As String
            Dim IPL, DPL As Integer
            'VARIABLE INITIALISATION
            IPv = ""
            yIPv = ""
            DPv = ""
            SignVln = ""
            Call VLnum.ValUpdater(PowerIndicaterUP, "+", "1", )
            VlnumXPows(1) = X
            Call VLnum.ValRetriever(Y, yIPv, DPv, IPL, DPL, SignVln)
            If SignVln = "-" Then
                SignVln = "+"
                SignNeg = True
                Call VLnum.ValUpdater(Y, SignVln, yIPv, )
            End If

            I = 1
            Do While Y > PowerIndicaterUP
                I += 1
                X = X * X
                PowerIndicaterUP = PowerIndicaterUP * cVlNums.Two
                If VlnumXPows.GetUpperBound(0) < I Then
                    ReDim Preserve VlnumXPows(I + 4)
                End If
                VlnumXPows(I) = X
                If Y = PowerIndicaterUP Then
                    If SignNeg Then
                        Call VLnum.ValUpdater(Y, "-", yIPv, )
                        Return cVlNums.One / X
                    Else
                        Return X
                    End If
                End If
            Loop

            PIndicater = PowerIndicaterUP * cVlNums.Half
            Call VLnum.ValRetriever(PIndicater, IPv, DPv, IPL, DPL, SignVln)
            Call VLnum.ValUpdater(PowerIndicaterUP, , IPv, )
            Call VLnum.ValUpdater(PowerIndicaterDN, SignVln, IPv, )
            X = VlnumXPows(I - 1)

            For J = I - 2 To 1 Step -1
                PIndicater = PowerIndicaterDN * cVlNums.Half
                Call VLnum.ValRetriever(PIndicater, IPv, DPv, IPL, DPL, SignVln)
                Call VLnum.ValUpdater(PowerIndicaterDN, , IPv, )

                If Y >= (PowerIndicaterUP + PowerIndicaterDN) Then
                    X = X * VlnumXPows(J)
                    PowerIndicaterUP = PowerIndicaterUP + PowerIndicaterDN
                End If
                If Y = PowerIndicaterUP Then Exit For
            Next
            If SignNeg Then
                X = cVlNums.One / X
                Call VLnum.ValUpdater(Y, "-", yIPv, )
            End If
            Return X
        End Function

        Public Shared Sub FunctionsDpValCorrecter(ByRef Sum As VLnum)
            'THIS BLOCK REMOVES THE TRAILING ZEROS IN THE DPVAL (OR CORRECTS THE REPEATING TRAILING 9 IN THE DPVAL)
            Dim IPv, DPv, SignVln As String
            Dim IPL, DPL As Integer
            Dim mxstrt, mxend, mxlen, _MaxLengthHere As Integer
            'VARIABLE INITIALISATION
            IPv = ""
            DPv = ""
            SignVln = ""

            Call VLnum.ValRetriever(Sum, IPv, DPv, IPL, DPL, SignVln)
            ' FINDS THE NO. OF TRAILING 0S OF THE STRING
            Call DPvsZeroNineCounter(DPv, mxstrt, mxend, mxlen, DPL, "0")
            If mxlen > 2 Then
                DPv = DPv.Substring(0, mxend + 1)
                Call VLnum.ValUpdater(Sum, SignVln, IPv, DPv)
            Else
                mxlen = 0
                ' FINDS THE NO. OF TRAILING 9S OF THE STRING
                Call DPvsZeroNineCounter(DPv, mxstrt, mxend, mxlen, DPL, "9")
                If mxlen > 2 Then
                    _MaxLengthHere = IPL + mxend + 1
                    Call Round(IPv, IPL, DPv, DPL, SignVln, _MaxLengthHere)
                    Call VLnum.ValUpdater(Sum, SignVln, IPv, DPv)
                End If
            End If
        End Sub

        Private Shared Sub DPvsZeroNineCounter(ByRef DPvH As String, ByRef mxstrtH As Integer, ByRef mxendH As Integer, _
                                               ByRef mxlenH As Integer, ByRef DPLH As Integer, ByRef ZN As String)
            Dim strtv, endv, strtb As Integer
            'VARIABLE INITIALISATION
            strtb = 1
            For J = DPLH - 1 To 0 Step -1
                If DPvH(J) = ZN Then
                    If strtb = 1 Then
                        strtb = 0
                        strtv = J
                        endv = -1
                    End If
                Else
                    If strtb = 0 Then
                        endv = J
                        strtb = 2
                        If (strtv - endv) > mxlenH Then
                            mxstrtH = strtv
                            mxendH = endv
                            mxlenH = strtv - endv
                        End If
                    End If
                End If
            Next
            If strtb = 0 And endv = -1 Then
                endv = 0
                If (strtv - endv) > mxlenH Then
                    mxstrtH = strtv
                    mxendH = endv
                    mxlenH = strtv - endv
                End If
            End If
        End Sub
#End Region
#Region "Trigonometric Functions"
        Public Shared Function SinR(ByVal X As VLnum) As VLnum
            'SinR(X) = X -  X^3/3! +  X^5/5! - X^7/7! +-----        
            Dim sum, Num2i, Term, XSquare, Num2iP1, MF, VlNumXRslt, VlNumXDiff, Xc, Xt As New VLnum
            Dim SignH, i As Integer
            'VARIABLE INITIALISATION
            Xc = X
            Xt = X
            XSquare = Xc * Xc
            Term = Xc
            sum = Xc
            SignH = 1

            i = 1
            Do While (Term.Value <> "0")
                SignH = -SignH '* (-1)
                VLnum.ValUpdater(Num2i, "+", (i * 2).ToString, "")
                VLnum.ValUpdater(Num2iP1, "+", (i * 2 + 1).ToString, "")
                MF = Num2i * Num2iP1
                Term = Term * (XSquare / MF)
                If SignH = +1 Then
                    sum = sum + Term
                Else
                    sum = sum - Term
                End If
                i += 1
            Loop

            VLnum.ValUpdater(sum, "N", "N", "N", "f")
            Return sum
        End Function
        Public Shared Function CosR(ByVal X As VLnum) As VLnum
            'CosR(X) = 1 -  X^2/2! +  X^4/4! - X^6/6! +-----
            Dim sum, Num2i, Term, XSquare, Num2iM1, MF, VlNumXRslt, VlNumXDiff, Xc, Xt As New VLnum
            Dim SignH, i As Integer
            'VARIABLE INITIALISATION
            Xc = X
            Xt = X
            XSquare = Xc * Xc
            Term = cVlNums.One
            sum = cVlNums.One
            SignH = 1
            i = 1
            Do While (Term.Value <> 0)
                SignH = -SignH '* (-1)
                VLnum.ValUpdater(Num2i, "+", (i * 2).ToString, "")
                VLnum.ValUpdater(Num2iM1, "+", (i * 2 - 1).ToString, "")
                MF = Num2i * Num2iM1
                Term = Term * (XSquare / MF)
                If SignH = +1 Then
                    sum = sum + Term
                Else
                    sum = sum - Term
                End If
                i += 1
            Loop
            VLnum.ValUpdater(sum, "N", "N", "N", "f")
            Return sum
        End Function
        Public Shared Function TanR(ByVal X As VLnum) As VLnum
            Dim VlnumRet, XcosR, Xt, Sum As New VLnum
            XcosR = CosR(X)
            If XcosR.Value <> "0" Then
                VlnumRet = SinR(X) / XcosR
            Else
                MsgBox("Invalid Input  'Tan(90)'", vbOKOnly, "Invalid Input")
                Throw New DivideByZeroException("An error occured in the DLL")
            End If
            VLnum.ValUpdater(VlnumRet, "N", "N", "N", "f")
            Return VlnumRet
        End Function
        Public Shared Function ASinR(ByVal X As VLnum) As VLnum
            'ASinR(X) = 2*AtanR(X/(1+SQRT(1-X^2))
            Dim VlNumXDiff, VlNumXRslt, XH, XHh, VlNumRet, Term As New VLnum
            Dim SignXNeg As Boolean = False
            Dim IPv, DPv, SignVln As String
            Dim IPL, DPL As Integer
            'VARIABLE INITIALISATION
            IPv = ""
            DPv = ""
            SignVln = ""
            Call VLnum.ValRetriever(X, IPv, DPv, IPL, DPL, SignVln)

            If X > cVlNums.One Or X < cVlNums.MOne Then
                MsgBox("Invalid Input  ' ASin(X) ' X Should be Between -1 & 1 ", vbOKOnly, "Invalid Input")
                Throw New ArgumentOutOfRangeException("Invalid input")
            End If
            If SignVln = "-" Then
                SignVln = "+"
                SignXNeg = True
                Call VLnum.ValUpdater(X, "+", "N", "N")
            End If
            If X.Value = "0" Then
                Call VLnum.ValUpdater(VlNumRet, "+", "0", "N")
                Return VlNumRet
            ElseIf X.Value = "1" Then
                VlNumRet = HalfPI()
                If SignXNeg Then
                    Call VLnum.ValUpdater(VlNumRet, "-", "N", "N")
                    Call VLnum.ValUpdater(X, "-", "N", "N")
                End If
                Return VlNumRet
            End If
            XH = X
            XHh = Sqrt(cVlNums.One - (XH * XH))
            XHh = cVlNums.One + XHh
            XH = XH / XHh
            VlNumRet = cVlNums.Two * ATanR(XH)
            VLnum.ValUpdater(VlNumRet, "N", "N", "N", "f")
            If SignXNeg Then
                Call VLnum.ValUpdater(VlNumRet, "-", "N", "N")
                Call VLnum.ValUpdater(X, "-", "N", "N")
            End If
            Return VlNumRet
        End Function
        Public Shared Function ACosR(ByVal X As VLnum) As VLnum
            'ACosR(X) = 2*AtanR((SQRT(1-X^2))/(1+X))
            Dim VlNumXDiff, VlNumXRslt, XH, XHh, VlNumRet, Term As New VLnum
            Dim SignXNeg As Boolean = False
            Dim IPv, DPv, SignVln As String
            Dim IPL, DPL As Integer
            'VARIABLE INITIALISATION
            IPv = ""
            DPv = ""
            SignVln = ""
            Call VLnum.ValRetriever(X, IPv, DPv, IPL, DPL, SignVln)

            If X > cVlNums.One Or X < cVlNums.MOne Then
                MsgBox("Invalid Input  ' ACos(X) ' X Should be Between -1 & 1 ", vbOKOnly, "Invalid Input")
                Throw New ArgumentOutOfRangeException("Invalid input")
            End If
            If SignVln = "-" Then
                SignVln = "+"
                SignXNeg = True
                Call VLnum.ValUpdater(X, "+", "N", "N")
            End If
            If X.Value = "0" Then
                VlNumRet = HalfPI()
                Return VlNumRet
            ElseIf X.Value = "1" Then
                If SignXNeg Then
                    VlNumRet = VLnMath.PI
                Else
                    Call VLnum.ValUpdater(VlNumRet, "+", "0", "N")
                End If
                Return VlNumRet
            End If
            XH = X
            XHh = Sqrt(cVlNums.One - (XH * XH))
            XH = XHh / (cVlNums.One + XH)
            VlNumRet = cVlNums.Two * ATanR(XH)
            VLnum.ValUpdater(VlNumRet, "N", "N", "N", "f")

            If SignXNeg Then
                VlNumRet = VLnMath.PI - VlNumRet
                Call VLnum.ValUpdater(X, "-", "N", "N")
            End If
            Return VlNumRet
        End Function
        Public Shared Function ATanR(ByVal X As VLnum) As VLnum
            'ATanR(X) = X -  X^3/3 +  X^5/5 - X^7/7 +-----      'Evaluation is only upto PI/8 (=22.5deg) then use eqn for double angle above 45deg tan(A) = 1/Tan(90-A)
            'arctan(x) = 2*arctan(x/(1+sqrt(1+x^2)) if x > .4142 (angle > 22.5 deg & < 45 deg) use this eqn to get angle/2 evaluate the series then multiply the result with 2
            Dim sum, Term, MinVal, XSquare, Num2iP1, Num2iM1, MF, VlNumXDiff, VlNumXRslt, XH, XHh, Termh As New VLnum
            Dim SignH, i, Twoi As Integer
            Dim SignXNeg As Boolean = False
            Dim XHgtOne As Boolean = False
            Dim XHgt22_5H As Boolean = False
            Dim IPv, DPv, SignVln As String
            Dim IPL, DPL As Integer
            'VARIABLE INITIALISATION
            IPv = ""
            DPv = ""
            SignVln = ""
            Call VLnum.ValRetriever(X, IPv, DPv, IPL, DPL, SignVln)
            If SignVln = "-" Then
                SignVln = "+"
                SignXNeg = True
                Call VLnum.ValUpdater(X, "+", "N", "N")
            End If

            If X.Value = "0" Then
                Call VLnum.ValUpdater(sum, "+", "0", "N")
                Return sum
            ElseIf X.Value = "1" Then
                sum = HalfPI() / cVlNums.Two
                If SignXNeg = True Then
                    Call VLnum.ValUpdater(sum, "-", "N", "N")
                    Call VLnum.ValUpdater(X, "-", "N", "N")
                End If
                Return sum
            End If
            XH = X
            If XH > cVlNums.One Then
                XH = cVlNums.One / XH
                XHgtOne = True
            End If
            If XH > Tan22_5H() Then
                XHh = Sqrt(cVlNums.One + (XH * XH))
                XHh = cVlNums.One + XHh
                XH = XH / XHh
                XHgt22_5H = True
            End If

            XSquare = XH * XH
            Term = XH
            sum = XH
            SignH = 1
            Termh = XH
            i = 1
            Do While (Term.Value <> "0")
                SignH = SignH * (-1)
                Twoi = 2 * i
                Call VLnum.ValUpdater(Num2iP1, "+", (Twoi + 1).ToString, "N")
                Call VLnum.ValUpdater(Num2iM1, "+", (Twoi - 1).ToString, "N")

                MF = Num2iM1 / Num2iP1
                Term = Term * XSquare * MF
                If SignH = +1 Then
                    sum = sum + Term
                Else
                    sum = sum - Term
                End If
                i += 1
            Loop

            If XHgt22_5H = True Then
                sum = cVlNums.Two * sum
            End If
            If XHgtOne = True Then
                sum = HalfPI() - sum
            End If
            If SignXNeg = True Then
                Call VLnum.ValUpdater(sum, "-", "N", "N")
                Call VLnum.ValUpdater(X, "-", "N", "N")
            End If
            Return sum
        End Function
        Public Shared Function ATanRforPI(ByVal X As VLnum) As VLnum
            'ATanR(X) = X -  X^3/3 +  X^5/5 - X^7/7 +-----          'Evaluation is only upto PI/8 (=22.5deg) then use eqn for double angle above 45deg tan(A) = 1/Tan(90-A)
            'arctan(x) = 2*arctan(x/(1+sqrt(1+x^2)) if x > .4142 (angle > 22.5 deg & < 45 deg) use this eqn to get angle/2 evaluate the series then multiply the result with 2
            Dim sum, Term, MinVal, XSquare, Num2iP1, Num2iM1, MF, VlNumXDiff, VlNumXRslt, XH, XHh, Termh, Tan22_5Hg As New VLnum
            Dim SignH, i, Twoi As Integer
            Dim SignXNeg As Boolean = False
            Dim XHgtOne As Boolean = False
            Dim XHgt22_5H As Boolean = False
            Dim IPv, DPv, SignVln As String
            Dim IPL, DPL As Integer
            'VARIABLE INITIALISATION
            IPv = ""
            DPv = ""
            SignVln = ""
            Call VLnum.ValRetriever(X, IPv, DPv, IPL, DPL, SignVln)
            If SignVln = "-" Then
                SignVln = "+"
                SignXNeg = True
                Call VLnum.ValUpdater(X, "+", "N", "N")
            End If

            If X.Value = "0" Then
                Call VLnum.ValUpdater(sum, "+", "0", "N")
                Return sum
            End If
            XH = X
            If XH > cVlNums.One Then
                XH = cVlNums.One / XH
                XHgtOne = True
            End If
            'BELOW LINE CAN BE USED AS THIS SUB IS USED ONLY TO EVALUATE VALUE OF PI ie X WILL ALWAYS BE 1
            Tan22_5Hg.Value = "+.414213562373095048801688724209697656188775556357"
            If XH > Tan22_5Hg Then
                XHh = Sqrt(cVlNums.One + (XH * XH))
                XHh = cVlNums.One + XHh
                XH = XH / XHh
                XHgt22_5H = True
            End If
            XSquare = XH * XH
            Term = XH
            sum = XH
            SignH = 1
            i = 1
            Do While (Term.Value <> "0")
                SignH = SignH * (-1)
                Twoi = 2 * i
                Call VLnum.ValUpdater(Num2iP1, "+", (Twoi + 1).ToString, "N")
                Call VLnum.ValUpdater(Num2iM1, "+", (Twoi - 1).ToString, "N")

                MF = Num2iM1 / Num2iP1
                Term = Term * XSquare * MF
                If SignH = +1 Then
                    sum = sum + Term
                Else
                    sum = sum - Term
                End If
                i += 1
            Loop

            If XHgt22_5H = True Then
                sum = cVlNums.Two * sum
            End If
            If XHgtOne = True Then
                sum = HalfPI() - sum
            End If
            If SignXNeg = True Then
                Call VLnum.ValUpdater(sum, "-", "N", "N")
                Call VLnum.ValUpdater(X, "-", "N", "N")
            End If
            Return sum
        End Function
        Public Shared Function SinD(ByVal X As VLnum) As VLnum
            Return SinR(X * VLnMath.DtoR)
        End Function
        Public Shared Function CosD(ByVal X As VLnum) As VLnum
            Return CosR(X * VLnMath.DtoR)
        End Function
        Public Shared Function TanD(ByVal X As VLnum) As VLnum
            Return TanR(X * VLnMath.DtoR)
        End Function
        Public Shared Function ASinD(ByVal X As VLnum) As VLnum
            Dim VlNumRet As New VLnum
            VlNumRet = ASinR(X) * VLnMath.RtoD
            VLnum.ValUpdater(VlNumRet, "N", "N", "N", "f")
            Return VlNumRet
        End Function
        Public Shared Function ACosD(ByVal X As VLnum) As VLnum
            Dim VlNumRet As New VLnum
            VlNumRet = ACosR(X) * VLnMath.RtoD
            VLnum.ValUpdater(VlNumRet, "N", "N", "N", "f")
            Return VlNumRet
        End Function
        Public Shared Function ATanD(ByVal X As VLnum) As VLnum
            Dim VlNumRet As New VLnum
            VlNumRet = ATanR(X) * VLnMath.RtoD
            VLnum.ValUpdater(VlNumRet, "N", "N", "N", "f")
            Return VlNumRet
        End Function

        Public Shared Sub UpdateConstants()
            Dim VlNumRet, VlNumPI As New VLnum
            Dim VlNum1 As New VLnum
            Dim IPv, DPv, SignVln As String
            Dim IPL, DPL As Integer
            IPv = ""
            DPv = ""
            SignVln = ""

            Dim elen, pilen, Hpilen, RtoDlen, DtoRlen, logtenElen, tan225hlen As Integer
            _MaxLength = _MaxLength + 20
            _MaxLengthIO = _MaxLengthIO + 20

            If _MaxLength > 10100 Then
                VlNumRet = ALoge(cVlNums.One)
                Call VLnum.ValRetriever(VlNumRet, IPv, DPv, IPL, DPL, SignVln)
                cVlNums.Constante = IPv & "." & DPv
                elen = cVlNums.Constante.Length
            End If

            If _MaxLength > 3000 Then
                Call VLnum.ValUpdater(VlNum1, "+", "4", "N")
                VlNumPI = VlNum1 * ATanRforPI(cVlNums.One)                   'PI should be calculated after constante is calculated 
                Call VLnum.ValRetriever(VlNumPI, IPv, DPv, IPL, DPL, SignVln)
                cVlNums.ConstantPI = IPv & "." & DPv
                pilen = cVlNums.ConstantPI.Length
                'Else
                'VlNumPI = VLnMath.PI
                'End If


                VlNum1 = VlNumPI * cVlNums.Half
                Call VLnum.ValRetriever(VlNum1, IPv, DPv, IPL, DPL, SignVln)
                cVlNums.ConstantHalfPI = IPv & "." & DPv
                Hpilen = cVlNums.ConstantHalfPI.Length

                Call VLnum.ValUpdater(VlNum1, "+", "180", "", "N")
                VlNumRet = VlNum1 / VlNumPI
                Call VLnum.ValRetriever(VlNumRet, IPv, DPv, IPL, DPL, SignVln)
                cVlNums.ConstantRtoD = IPv & "." & DPv
                RtoDlen = cVlNums.ConstantRtoD.Length

                VlNumRet = VlNumPI / VlNum1
                Call VLnum.ValRetriever(VlNumRet, IPv, DPv, IPL, DPL, SignVln)
                cVlNums.ConstantDtoR = IPv & "." & DPv
                DtoRlen = cVlNums.ConstantDtoR.Length

                Call VLnum.ValUpdater(VlNum1, "+", "10", "", "N")
                VlNumRet = Loge(VlNum1)
                Call VLnum.ValRetriever(VlNumRet, IPv, DPv, IPL, DPL, SignVln)
                cVlNums.ConstantLogTENe = IPv & "." & DPv
                logtenElen = cVlNums.ConstantLogTENe.Length

                Call VLnum.ValUpdater(VlNum1, "+", "22", "5")   '22.5
                VlNumRet = TanD(VlNum1)
                Call VLnum.ValRetriever(VlNumRet, IPv, DPv, IPL, DPL, SignVln)
                cVlNums.Tan22_5H = IPv & "." & DPv
                tan225hlen = cVlNums.Tan22_5H.Length
            End If

            _MaxLength = _MaxLength - 20
            _MaxLengthIO = _MaxLengthIO - 20

            ConstantsLen = _MaxLength
        End Sub
#End Region

    End Class
#End Region
    Private Class cVlNums
        Public Shared Zero As New VLnum
        Public Shared Half As New VLnum
        Public Shared One As New VLnum
        Public Shared MOne As New VLnum
        Public Shared a3BY2 As New VLnum
        Public Shared Two As New VLnum

        Public Shared ConstantRtoD As String
        Public Shared ConstantDtoR As String
        Public Shared Constante As String
        Public Shared ConstantLogTENe As String
        Public Shared ConstantPI As String
        'Public Shared ConstantTwoPI As String
        Public Shared ConstantHalfPI As String
        'Public Shared ConstantA3by2PI As String
        Public Shared Tan22_5H As String

        Shared Sub New()
            Zero.Value = "+0"
            Half.Value = "+0.5"
            One.Value = "+1"
            MOne.Value = "-1"
            a3BY2.Value = "+1.5"
            Two.Value = "+2"

            'UpdateConstants EVALUATES THESE VALS WITH OUT SIGN. THE SUBSTRINGS IN PROPERTIES AND FUNCTIONS WILL NOT BE CORRECT IF SIGN IS ADDED HERE
            'Initially only 100 digits of constants were calculated & used
            'ConstantRtoD = "57.2957795130823208767981548141051703324054724665643215491602438612028471483215526324409689958511109456844248"
            'ConstantDtoR = "0.017453292519943295769236907684886127134428718885417254560971914401710091146034494436822415696345094822123045"
            Call strRtoD(ConstantRtoD)
            Call strDtoR(ConstantDtoR)
            Call streVal(Constante)  'reads it from the sub
            Call strLogTENe(ConstantLogTENe)
            'ConstantLogTENe = "2.30258509299404568401799145468436420760110148862877297603332790096757260967735248023599720508959829834196778"
            'ConstantPI = "3.14159265358979323846264338327950288419716939937510582097494459230781640628620899862803482534211706798214808"
            Call strPIValCalc(ConstantPI)  'reads it from the sub
            Call strHalfPI(ConstantHalfPI)
            Call strTan22_5H(Tan22_5H)

            'ConstantHalfPI = "1.57079632679489661923132169163975144209858469968755291048747229615390820314310449931401741267105853399107404"
            'Tan22_5H = "0.4142135623730950488016887242096980785696718753769480731766797379907324784621070388503875343276"
        End Sub
        Private Shared Sub streVal(ByRef strEvalCalc As String) '10100 digits   run time 8392 seconds
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
        Private Shared Sub streValChatGpt(ByRef strEvalChatGPT As String)   '10500 digits
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
        Private Shared Sub strPIValCalc(ByRef strPIvalCal As String)    '3001 digits    Run Time 4747 seconds
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
        Private Shared Sub strPIValChatGPT(ByRef strPIvalChatGPT As String) '10500 digits
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
        Private Shared Sub strRtoD(ByRef strCRtoD As String) '3000 digits
            strCRtoD =
            "57.29577951308232087679815481410517033240547246656432154916024386120284714832155263244096899585111094418622338163286489328144826" &
            "46012483150360682678634119421225263880974672679263079887028931107679382614426382631582096104604870205064442596568411201719120577" &
            "38566280431284962624203376187937297623870790340315980719624089522045186205459923396314841906966220115126609691801514787637366923" &
            "16410712677403851469016549959419251571198647943521066162438903520230675617779675711331568350620573131336015650134889801878870991" &
            "77764391811593169200139029797682608293230553397026181660490929593282083154995798031955967007118252058466439231799858456719168439" &
            "91775413165295995305640627904496724872253434072473833065185879008820107193548552060485006842647349075605874388567532932178246602" &
            "12423313273427212944530891667616714672023231374957844879817729067269084978013518866274395505591858069303343806736897172694639833" &
            "88546519913267567748257649341692338111382433230532082241985164541395979472567140094041784696151560535821579204602135452019010741" &
            "62102935129261609789844003532238204473276282720909209180171145108905646530208200979123059094023542240418735583708743755384573586" &
            "32910219947278955982189157952380660051799046457828270494545077457614758622997378794309175064277403984832908061531374559794227515" &
            "56017881484918518940459609395139971984530643555626981802696960336478401854845389817269860386482453584697394669284938744103484193" &
            "38670826807976466071601804527844317125267416517826117277993857283514535275841382648354898368971954748747942507185719709447566116" &
            "44011772072636212722247298867746192765759580493338310158760887887009555037264576645556336897953908727599340763884639572313598707" &
            "97922613653859299456731829760185434493762364931628853725016524805627090830946056659456251070994031619085592872592047590851137444" &
            "96282219561728074051789812133099369970324357691395511633716219273108331509936077800662768516345100374220246196814323220495858888" &
            "64098604391498374772365248457401038160517683907375005860137149840535964250749278164464419012173255104442039773277474132023275126" &
            "21566758423962311946218475494856854789875795711269838899487183081334908351570192389306646409405968180138439601736690834103752618" &
            "78587938980076436281892960841778275647074485954160680355133364287619561679420056757716493862578546227081881795684966308392742344" &
            "71335197811813583937460891027231008533858483010246375562087255297063182528131931233736528329059984251906223655691271999031242440" &
            "25074213094377890275209970389790613322899350442217973566000933435216018915815555845879646566881177158624270449272860629813820671" &
            "93161017863158253947123135104153039915542996143494810231621952631198538198082143167177922061432944306758841812979138068828858121" &
            "44451227941600787951613897488065914624794792282070691117404841375862132239321947247305245100749634622695917221984653725972335866" &
            "74999095350791219753901965011720360565813584299729041079852894678263082373920251053942296358829615424822523456226572403772799160" &
            "492308254608300540571862909907528526672386676206949322120165990848160"
        End Sub
        Private Shared Sub strDtoR(ByRef strCDtoR As String) '3000 digits
            strCDtoR =
            "0.017453292519943295769236907684886127134428718885417254560971914401710091146034494436822415696345094822123044925073790592483854" &
            "69227528101239847421893404711731916824501501076956169755358123860530516878869127117208703296358960264249018770435091817334393969" &
            "80475940192241589469684813789632978181124952292984699278144795310454160084495609046069671761964687105143908889518362808267803695" &
            "63245260844119508941294762613143108844183845478429899625621072806214155969235444237497596399365292916062377434350066384054631518" &
            "68022587023936678552747997347076217056766589413168205855120653496209306880374899148705225073333648959525146422682103206301532105" &
            "33842979843262303802272290275190563697199187280599571093847717974556664228451612331159113023231100757209709517220028817067297222" &
            "22213183211388616998509626756090658861246996974149490570236235045851914916862566284378727833350765770849369930740046563447873209" &
            "27304057554585272460419704850644201591045752104218751087656587655851206237114785001071042561775505120233443854497365111703047701" &
            "82159218675187879331568350108446405658498277542979033300771736095654182415507306420825402358563927552823862950936762658660504172" &
            "13231970208138551773639224449598342617438943604577849213120019798375889470612121905308867771926487985830268085443192926928355819" &
            "63692337803713477260828496285376127216195613751201142758902254465638996395140768016686437577915275818479952332929224594401553767" &
            "97487868667185651202289995810350835015899054141983708324361416366032607055316162622822083850164184509185832622375331124248925861" &
            "06221565748876408702054485703014043680843435653192627172098737743333786928111200806939956517873415401945230233186491934229784207" &
            "51417851930967694149132512957726330079630304312453812510546427491978944010678991262527919031604262105830334251926002771459573773" &
            "21449210243545998220926747450052993543686719482225790284733617632943605338007138257052533568998071390122814510350374682145668844" &
            "16361372909521038671979802066207152598692542077568966060365736922198963280433486611663698689327507052602213069976165699014558458" &
            "27448737066376815823667487217168133409108020808502815668609029139860777584128572116217032917187345386923064283254659672554381088" &
            "76276416607231208551581563237149103830154119777325329180330775523947220695815602554848535816436036324263123479227729144891735771" &
            "35502506911869720904337765087174644431667386756052453860380865824479741233734587936048045324755713922315792996957041487104973622" &
            "99179194419259293235548031442286863825683654142498908663130315652402284730597816377472831442274458460488728273969816187148460416" &
            "07808894527961589884926717826375606545829921611650828315332514520308321216284988045649055497151258822492086889681693197507354424" &
            "52508465256857580793658026640365932339173007526309664017296817589456310941867952460847138539928389698696686612666302243011151725" &
            "43632812332437471004710272046289692063260417746392461232473995026933891872563711486066265661776633306700788701904863578135413957" &
            "621217877776883897733089897041745939577638300503992496795534204414004"
        End Sub
        Private Shared Sub strLogTENe(ByRef strCLogTENe As String) '3000 digits
            strCLogTENe =
            "2.302585092994045684017991454684364207601101488628772976033327900967572609677352480235997205089598298341967784042286248633409525" &
            "46508280675666628736909878168948290720832555468084379989482623319852839350530896537773262884616336622228769821988674654366747440" &
            "42432743651550489343149393914796194044002221051017141748003688084012647080685567743216228355220114804663715659121373450747856947" &
            "68346361679210180644507064800027750268491674655058685693567342067058113642922455440575892572420824131469568901675894025677631135" &
            "69192920333765871416602301057030896345720754403708474699401682692828084811842893148485249486448719278096762712757753970276686059" &
            "52496716674183485704422507197965004714951050492214776567636938662976979522110718264549734772662425709429322582798502585509785265" &
            "38320760672631716430950599508780752371033310119785754733154142180842754386359177811705430982748238504564801909561029929182431823" &
            "75253577097505395651876975103749708886921802051893395072385392051446341972652872869651108625714921988499787488737713456862091670" &
            "58498078280597511938544450099781311469159346662410718466923101075984383191912922307925037472986509290098803919417026544168163357" &
            "27555703151596113564846546190897042819763365836983716328982174407366009162177850541779276367731145041782137660111010731042397832" &
            "52189489881759792179866639431952393685591644711824675324563091252877833096360426298215304087456092776072664135478757661626292656" &
            "82987049579549139549180492090694385807900327630179415031178668620924085379498612649334793548717374516758095370882810674524401058" &
            "92444976479686075120275724181874989395971643105518848195288330746699317814634930000321200327765654130472621883970596794457943468" &
            "34321839530441484480370130575367426215367557981477045803141363779323629156012818533649846694226146520645994207291711937060244492" &
            "93580370077189810973625332245483669885055282859661928050984471751985036666808749704969822732202448233430971691111368135884186965" &
            "49323714996941979687803008850408979618598756579894836445212043698216415292987811742973332588607915912510967187510929248475023930" &
            "57266544627620092306879151813580347770129559364629841236649702335517458619556477246185771736936840467657704787431978057385327181" &
            "09338834963388130699455693993461010907456160333122479493604553618491233330637047517248712763791409243983318101647378233796922656" &
            "37682071706935846394531616949411701841938119405416449466111274712819705817783293841742231409930022911502362192186723337268385688" &
            "27353337192510341293070563254442661142976538830182238409102619858288843358745596045300454837078905257847316628370195339223104752" &
            "75649981192287427897137157132283196410034221242100821806795252766898581809561192083917607210809199234615169525990994737827806481" &
            "28058792731993893453415320185969711021407542282796298237068941764740642225757212455392526179373652434440560595336591539160312524" &
            "48014931323457245387952438903683923645050788173135971123814532370150841349112232439092768172474960795579915136398288105828574053" &
            "800065337165555301419633224191808762101820491949265148387765960691546"
        End Sub
        Private Shared Sub strHalfPI(ByRef strCHalfPI As String) '3000 digits
            strCHalfPI =
            "1.570796326794896619231321691639751442098584699687552910487472296153908203143104499314017412671058533991074043256641153323546922" &
            "30477529111586267970406424055872514205135096926055277982231147447746519098221440548783296672306423782411689339158263560095457282" &
            "42834617301743052271633241066968036301245706368622935033031577940874407604604814146270458576821839462951800056652652744102332606" &
            "92073475970755804716528635182879795976546093058690966305896552559274037231189981374783675942876362445613969091505974564916836681" &
            "22032832154301069747319761236859535108993047185138526960858814658837619233740923383470256600028406357263178041389288567137889480" &
            "45868185893607342204506124767150732747926855253961398446294617710099780560645109804320172090799068148873856549802593536056749999" &
            "99186489024975529865866408048159297512229727673454151321261154126672342517630965594085505001568919376443293766604190710308588834" &
            "57365179912674521437773436557978143194117689379687597889092889026608561340330650096393830559795460821009946904762860053274293163" &
            "94329680766909139841151509760176509264844978868112997069456248608876417395657577874286212270753479754147665584308639279445375491" &
            "90877318732469659627530200463850835569504924412006429180801781853830052355090971477798099473383918724724127689887363423552023767" &
            "32310402334212953474564665683851449457605237608102848301202901907509675562669121501779382012374823663195709963630213496139839117" &
            "73908180046708608206099622931575151430914872778533749192527472942934634978454636053987546514776605826724936013779801182403327495" &
            "59940917398876783184903713271263931275909208787336445488886396900040823530008072624596086608607386175070720986784274080680578676" &
            "27606673787092473421926166195369707166727388120843125949178474278104960961109213627512712844383589524730082673340249431361639589" &
            "30428921919139839883407270504769418931804753400321125626025586964924480420642443134728021209826425111053305931533721393110195974" &
            "72523561856893480478182185958643733882328786981206945432916322997906695239013795049732882039475634734199176297854912911310261244" &
            "70386335973913424130073849545132006819721872765253410174812622587469982571571490459532962546861084823075785492919370529894297988" &
            "64877494650808769642340691343419344713870779959279626229769797155249862623404229936368223479243269183681113130495623040256219421" &
            "95225622068274881390398857845717998850064808044720847434277924203176711036112914244324079228014253008421369726133733839447626069" &
            "26127497733336391199322829805817744311528872824901779681728408716205625753803473972554829804701261443985544657283456843361437447" &
            "02800507516543089643404604373804589124692945048574548379926306827748909465648924108414994743613294024287820071352387775661898207" &
            "25761873117182271429222397632933910525570677367869761556713583051067984768115721476242468593555072882701795139967201871003655289" &
            "26953109919372390423924484166072285693437597175321510922659552424050268530734033745963909559896997603070983171437722032187256185" &
            "90960899991955079597809073375713456198744704535932471159807839726040"
        End Sub
        Private Shared Sub strTan22_5H(ByRef strCTan22_5H As String) '3000 digits
            strCTan22_5H =
            "0.414213562373095048801688724209698078569671875376948073176679737990732478462107038850387534327641572735013846230912297024924836" &
            "05585073721264412149709993583141322266592750559275579995050115278206057147010955997160597027453459686201472851741864088919860955" &
            "23292304843087143214508397626036279952514079896872533965463318088296406206220346864419002173044630698557500423278723288078142120" &
            "56735909227205932163532401389802828774920018846707106655940065814838042675399900731459834806555004396029687114790533668310406303" &
            "91340540639888101930560831957126517360403135183048672992708428230633442792711542299940045342737969698382939763357879944453945472" &
            "29350033068400555325134835484043544099824533685379152774130030881297470550393669876088853898855089334443941370890050750891864866" &
            "18284419056574666177001064897675533960747772899043188884881773120086620140927181456917176137117379699238596753716689953583136255" &
            "95231750563373565789768472924653510092527134805989019849573984923527461213805583043913307584157228676377098264936515003205771665" &
            "02409332797052361795228151110564474297183506569502283172266367418817685080421641690732796001025641331811545651562294739030610141" &
            "58516753131531833962784840333504047404905928058365878181361218632829524749673155965784526237397128043010079764301300433367172879" &
            "98362388116661651613883166569575277921422569309427144664881101809920895857511333485405955690805113874538368709940832763686202733" &
            "73714039563658982067470454754816565365030404132398386941569205354924486180493675580989601740958638948026733830367430764576250806" &
            "05315786365242165509090125520882172726869812736594126387973796745797063352485636494290365872008009274178763584525567265487783505" &
            "70371248456521059787893478539554966483779819447689945267763351285417063804673800275700555228658114821049007790698844731413876435" &
            "33352697216460459845249428331011441733778642648996836717798973990518442460034804062693867333732684272494960639522262861209220307" &
            "64624031711697218311196410471861370797501212251757625274606921497610539169322948180964755229177271108777569627899967459512059253" &
            "84552463881128508567508749455233862027794253316151856082986095211717035138465687901160187922545498359719631058596056839734591191" &
            "30901502532027872629593785875146734075174689080028353506549095049702066416105404535787533940192567307445900484241785954242784901" &
            "94079463901516981913340972056195864756279950054690695240818284331326049159196097191507715940408766048006337224313315096799079462" &
            "10192904531575371551112744346139591987502358820852211921181964779406134338660947475272119676433408019873699886525023899946921818" &
            "72797348031508965791303813805023510761166335242194962419281959289395416867965098388938277630979204463082061912654289784022252265" &
            "97799442497893293060119444688873512051603790567695612328455096407414412409590478337870227745355438609120142344860106565043822157" &
            "54704033078580302992208016261631602939952736157157415778913909667931370066736328245593478157856726808549237232366463262378443666" &
            "504364488128206310094544871646589310530620379995750414850625430665974"
        End Sub
    End Class

End Class
