# Open Packaging Convention package

A new instance is constructed by calling `OpcPackage$open()` with a path
to a package file (.pptx).

## Methods

### Public methods

- [`OpcPackage$new()`](#method-OpcPackage-new)

- [`OpcPackage$drop_rel()`](#method-OpcPackage-drop_rel)

- [`OpcPackage$iter_parts()`](#method-OpcPackage-iter_parts)

- [`OpcPackage$iter_rels()`](#method-OpcPackage-iter_rels)

- [`OpcPackage$next_partname()`](#method-OpcPackage-next_partname)

- [`OpcPackage$part_related_by()`](#method-OpcPackage-part_related_by)

- [`OpcPackage$relate_to()`](#method-OpcPackage-relate_to)

- [`OpcPackage$related_part()`](#method-OpcPackage-related_part)

- [`OpcPackage$target_ref()`](#method-OpcPackage-target_ref)

- [`OpcPackage$save()`](#method-OpcPackage-save)

- [`OpcPackage$clone()`](#method-OpcPackage-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    OpcPackage$new(pkg_file)

------------------------------------------------------------------------

### Method `drop_rel()`

#### Usage

    OpcPackage$drop_rel(rId)

------------------------------------------------------------------------

### Method `iter_parts()`

#### Usage

    OpcPackage$iter_parts()

------------------------------------------------------------------------

### Method `iter_rels()`

#### Usage

    OpcPackage$iter_rels()

------------------------------------------------------------------------

### Method `next_partname()`

#### Usage

    OpcPackage$next_partname(tmpl)

------------------------------------------------------------------------

### Method `part_related_by()`

#### Usage

    OpcPackage$part_related_by(reltype)

------------------------------------------------------------------------

### Method `relate_to()`

#### Usage

    OpcPackage$relate_to(target, reltype, is_external = FALSE)

------------------------------------------------------------------------

### Method `related_part()`

#### Usage

    OpcPackage$related_part(rId)

------------------------------------------------------------------------

### Method `target_ref()`

#### Usage

    OpcPackage$target_ref(rId)

------------------------------------------------------------------------

### Method [`save()`](https://rdrr.io/r/base/save.html)

#### Usage

    OpcPackage$save(pkg_file)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    OpcPackage$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
