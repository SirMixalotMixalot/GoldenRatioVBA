# GoldenRatioVBA
A VBA script that creates a plot of the golden ratio in Microsoft Word.

# Why?
idk, seemed like a fun idea
# Installation
1. Download `golden_ratio.docm`
2. [Unblock the file](https://support.microsoft.com/en-gb/topic/a-potentially-dangerous-macro-has-been-blocked-0952faa0-37e7-4316-b61d-5b5ed6024216)  (trust me)
3. Open it in word
4. Click on the Macros section under view
5. Click the macro and hit run!

And if that doesn't work, here's the source code because sometimes you just can't be bothered:
```VBA
Sub GoldenRain()
'
' Makes the spiral thing. Very funS
'

Const segments = 512
Const PI As Double = 3.14159265358979
Dim poly_points(0 To segments, 1 To 2) As Single
Dim radius As Double
Dim xorigin As Long
Dim yorigin As Long

Dim angle As Double
Dim spins As Long

xorigin = 120
yorigin = 120

spins = 4
Dim scale_factor As Double

scale_factor = 1 / 100
For counter = 0 To segments
    
    angle = (2 * spins) / segments * counter * PI
    radius = (1.358456 ^ angle) * scale_factor
    '          = (golden ratio)^2/PI
    poly_points(counter, 1) = radius * Cos(angle + PI) + xorigin 'Flip it
    poly_points(counter, 2) = radius * -Sin(angle + PI) + yorigin 'And Dip it
    
Next counter

With ActiveDocument.Shapes.AddPolyline(poly_points)
    .Line.ForeColor.RGB = &HFFD700 'This isn't enough to change the colour ¯\_(._.)_/¯
    .Line.Visible = 1
    .Line.Weight = 2
    .Fill.Visible = 0
    .Select
End With


End Sub
```
