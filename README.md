Hamop – High Accuracy Math Operations (VB.NET)

Hamop is a mathematical library written in VB.NET (Visual Studio 2013) targeting .NET Framework 4.5, designed for arbitrary-precision calculations on integers and decimals.

While VB.NET supports arbitrary-size integer operations using BigInteger, high-precision decimal arithmetic is not supported natively.
Hamop fills this gap by providing very large number (VLnum) support for both integers and decimals.

Key Concepts

Hamop → High Accuracy Math OPerations

VLnum → Very Large Number (arbitrary-precision number type)

Features
Arithmetic Operations

Addition (+)

Subtraction (-)

Multiplication (*)

Division (/)

Logarithmic Functions

Logarithm base e

Logarithm base 10

Trigonometric Functions

Sin, Cos, Tan

Inverse trigonometric functions

Supports radians and degrees

Relational Operations

<, <=, >, >=, <>, =

Math Functions

Floor

Ceiling

Mod

Power (^)

Absolute value (Abs)

Requirements

Windows

Visual Studio 2013 or later

.NET Framework 4.5

Using the Library (Hamop.dll)
Basic Usage

Add Hamop.dll to your project, add a reference, and import the namespace.

Important:
Set the precision (Maximum Length, default = 28) before performing calculations.

Dim hamop As New Hamop()

Creating Numbers
Dim N1, N2, N3, A, B, R As New Hamop.VLnum


All numbers must be supplied as strings:

N1.Value = "12"
N2.Value = "40.26"
B.Value = TextBoxInput.Text

Performing Calculations
A = N1 * N2
N3 = N1 / N2
B = N1 - N3

Getting the Result
TextBoxResult.Text = B.Value   ' Returns a string

Mathematical Functions (VLnMath)

All advanced math functions are implemented in the VLnMath subclass.

Trigonometric Examples
R = hamop.VLnMath.SinR(N1)   ' Sine (radians)
R = hamop.VLnMath.SinD(N1)   ' Sine (degrees)

R = hamop.VLnMath.ASinR(N1)  ' Inverse sine (radians)
R = hamop.VLnMath.ASinD(N1)  ' Inverse sine (degrees)


Inverse trigonometric sine functions require the input value to be between –1 and 1.

Other logarithmic and mathematical functions can be used similarly, much like the built-in Math library in VB.NET.

User Interface (UI)

⚠️ The UI is NOT intended to be used as a calculator.

It exists only for development and testing of the library.

UI Purpose

Validate results

Compare Hamop results with VB.NET Math library

Test very large numbers easily

UI Features

Input fields:

First Number

Second Number

Result

Result CPU (VB.NET Math result)

Digit counters showing number of digits (excluding sign and decimal point)

Automatic trimming if input exceeds Maximum Length

Validation

If Maximum Length < 29, results are also calculated using VB.NET Math

Result comparison labels indicate:

Matching digits

Non-matching digits

Reverse Operation Testing

A REVERSE button is provided for validation:

Addition ↔ Subtraction

Multiplication ↔ Division

Sin() ↔ ASin()

Loge() ↔ ALoge()

⚠️ Note:

For multiplication/division tests, multiply first, then divide
(division results may be rounded and cannot always restore the original value)

Automated Testing (VLnum Tests)

A “Test VLnum” button allows automated test cases.

Input values:

3 → Integer tests

4 → Decimal tests

5 → Mixed integer & decimal tests

6, 7 → Predefined large numbers

Random numbers are used for large-number testing, as manual input is impractical.

Project Structure

Forms Application

UI and testing code

Class Library

Core Hamop implementation

Compiled output:

bin/Debug/Hamop.dll

Constants & High Precision Usage (>3000 digits)

Hamop uses pre-computed constants (including π and e).

Important Notes

For precision above 3000 digits, constants must be recalculated

Automatic recalculation can take hours

Constants must be updated in the order specified in:

VLnMath.UpdateConstants()


If you only need basic arithmetic (+, −, ×, ÷) beyond 3000 digits, you may comment out:

If _MaxLength > ConstantsLen Then Call VLnMath.UpdateConstants()


Constants are stored in the constants folder.

Build Notes

Leave the Root Namespace empty when rebuilding Hamop.dll

License

(Choose MIT or your preferred license)
