# Collection of relationships from a part or package to other parts

Keyed by rId. Supports dict-like access.

## Methods

### Public methods

- [`Relationships$new()`](#method-Relationships-new)

- [`Relationships$get()`](#method-Relationships-get)

- [`Relationships$contains()`](#method-Relationships-contains)

- [`Relationships$values()`](#method-Relationships-values)

- [`Relationships$keys()`](#method-Relationships-keys)

- [`Relationships$get_or_add()`](#method-Relationships-get_or_add)

- [`Relationships$get_or_add_ext_rel()`](#method-Relationships-get_or_add_ext_rel)

- [`Relationships$load_from_xml()`](#method-Relationships-load_from_xml)

- [`Relationships$part_with_reltype()`](#method-Relationships-part_with_reltype)

- [`Relationships$pop()`](#method-Relationships-pop)

- [`Relationships$xml_bytes()`](#method-Relationships-xml_bytes)

- [`Relationships$clone()`](#method-Relationships-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    Relationships$new(base_uri)

------------------------------------------------------------------------

### Method [`get()`](https://rdrr.io/r/base/get.html)

#### Usage

    Relationships$get(rId)

------------------------------------------------------------------------

### Method `contains()`

#### Usage

    Relationships$contains(rId)

------------------------------------------------------------------------

### Method `values()`

#### Usage

    Relationships$values()

------------------------------------------------------------------------

### Method `keys()`

#### Usage

    Relationships$keys()

------------------------------------------------------------------------

### Method `get_or_add()`

#### Usage

    Relationships$get_or_add(reltype, target_part)

------------------------------------------------------------------------

### Method `get_or_add_ext_rel()`

#### Usage

    Relationships$get_or_add_ext_rel(reltype, target_ref)

------------------------------------------------------------------------

### Method `load_from_xml()`

#### Usage

    Relationships$load_from_xml(base_uri, xml_rels, parts)

------------------------------------------------------------------------

### Method `part_with_reltype()`

#### Usage

    Relationships$part_with_reltype(reltype)

------------------------------------------------------------------------

### Method `pop()`

#### Usage

    Relationships$pop(rId)

------------------------------------------------------------------------

### Method `xml_bytes()`

#### Usage

    Relationships$xml_bytes()

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    Relationships$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
