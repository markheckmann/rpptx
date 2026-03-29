# Create a Length value in English Metric Units (EMU)

Base constructor for length values. All length values are stored as
integer EMU values. Use convenience constructors `Inches()`, `Cm()`,
`Mm()`, `Pt()`, `Emu()`, and `Centipoints()` for common units.

## Usage

``` r
Length(emu)

Inches(inches)

Cm(cm)

Mm(mm)

Pt(pt)

Emu(emu)

Centipoints(centipoints)
```

## Arguments

- emu:

  Integer EMU value.

- inches:

  Numeric length in inches.

- cm:

  Numeric length in centimeters.

- mm:

  Numeric length in millimeters.

- pt:

  Numeric length in points.

- centipoints:

  Integer length in hundredths of a point (1/7200 inch).

## Value

An integer with class `"Length"`.
