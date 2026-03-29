# Return appropriate shape proxy for a shape XML element

Dispatches on element tag and placeholder context. p:sp with p:ph →
SlidePlaceholder / LayoutPlaceholder / MasterPlaceholder; p:sp without
p:ph → Shape; p:pic → Picture; p:cxnSp → Connector; p:grpSp →
GroupShape; p:graphicFrame → GraphicFrame.

## Usage

``` r
shape_factory(shape_elm, parent)
```

## Arguments

- shape_elm:

  A BaseShapeElement (or subclass) R6 wrapper.

- parent:

  The parent ProvidesPart object (e.g. a Slide).

## Value

An R6 shape proxy.
