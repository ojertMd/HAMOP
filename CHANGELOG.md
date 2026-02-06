# Changelog

All notable changes to **Hamop (High Accuracy Math OPerations)** will be documented in this file.

Hamop provides arbitrary-precision integer and decimal mathematics for VB.NET,
beyond the limits of native data types.

---

## [1.0.0] – Initial Public Release

### Added
- `VLnum` (Very Large Number) data type for arbitrary-size integers and decimals
- Support for numbers supplied and returned as strings
- Basic arithmetic operations:
  - Addition (+)
  - Subtraction (−)
  - Multiplication (×)
  - Division (÷)
- Relational operations:
  - `<`, `<=`, `>`, `>=`, `=`, `<>`
- High-precision mathematical functions:
  - `Floor`
  - `Ceil`
  - `Mod`
  - Power (`^`)
  - Absolute value (`Abs`)
- Logarithmic functions:
  - Natural logarithm (base e)
  - Logarithm base 10
- Trigonometric functions (via `VLnMath`):
  - `Sin`, `Cos`, `Tan` (Radians and Degrees)
  - Inverse trigonometric functions (`ASin`, `ACos`, `ATan`)
- Conversion utilities:
  - Radians to Degrees
  - Degrees to Radians
- Pre-computed mathematical constants:
  - π (Pi)
  - e (Euler’s number)
  - Degree/Radian conversion constants
- Configurable precision via `MaximumLength`
- Development and testing UI (Windows Forms application)
  - Result validation against VB.NET `Math` library (CPU result)
  - Reverse-operation testing for validation
  - Random and predefined test value generation

### Architecture
- Core functionality implemented in **Hamop Class Library**
- Windows Forms application included for development and validation
- Constants stored separately and used by `VLnMath`
- Constants recalculation supported for precision beyond 3000 digits

### Notes
- Target framework: **.NET Framework 4.5**
- Developed using **Visual Studio 2013**
- For precision above 3000 digits:
  - Mathematical constants must be recalculated and stored manually
  - Automatic recalculation is available but may take several hours
- For basic arithmetic only at very high precision, constant updates can be disabled

---

