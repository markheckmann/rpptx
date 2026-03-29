# Base class for package parts

Provides common properties and methods for all part types.

## Methods

### Public methods

- [`Part$new()`](#method-Part-new)

- [`Part$load_rels_from_xml()`](#method-Part-load_rels_from_xml)

- [`Part$has_rels()`](#method-Part-has_rels)

- [`Part$rels_xml_bytes()`](#method-Part-rels_xml_bytes)

- [`Part$relate_to()`](#method-Part-relate_to)

- [`Part$part_related_by()`](#method-Part-part_related_by)

- [`Part$related_part()`](#method-Part-related_part)

- [`Part$target_ref()`](#method-Part-target_ref)

- [`Part$drop_rel()`](#method-Part-drop_rel)

- [`Part$clone()`](#method-Part-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    Part$new(partname, content_type, package, blob = NULL)

------------------------------------------------------------------------

### Method `load_rels_from_xml()`

#### Usage

    Part$load_rels_from_xml(xml_rels, parts)

------------------------------------------------------------------------

### Method `has_rels()`

#### Usage

    Part$has_rels()

------------------------------------------------------------------------

### Method `rels_xml_bytes()`

#### Usage

    Part$rels_xml_bytes()

------------------------------------------------------------------------

### Method `relate_to()`

#### Usage

    Part$relate_to(target, reltype, is_external = FALSE)

------------------------------------------------------------------------

### Method `part_related_by()`

#### Usage

    Part$part_related_by(reltype)

------------------------------------------------------------------------

### Method `related_part()`

#### Usage

    Part$related_part(rId)

------------------------------------------------------------------------

### Method `target_ref()`

#### Usage

    Part$target_ref(rId)

------------------------------------------------------------------------

### Method `drop_rel()`

#### Usage

    Part$drop_rel(rId)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    Part$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
