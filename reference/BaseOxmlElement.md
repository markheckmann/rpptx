# Base class for all custom XML element classes

Wraps an xml2 xml_node and provides standardized methods for child
element access, insertion, removal, and attribute manipulation.

## Methods

### Public methods

- [`BaseOxmlElement$new()`](#method-BaseOxmlElement-new)

- [`BaseOxmlElement$get_node()`](#method-BaseOxmlElement-get_node)

- [`BaseOxmlElement$find()`](#method-BaseOxmlElement-find)

- [`BaseOxmlElement$findall()`](#method-BaseOxmlElement-findall)

- [`BaseOxmlElement$get_attr()`](#method-BaseOxmlElement-get_attr)

- [`BaseOxmlElement$set_attr()`](#method-BaseOxmlElement-set_attr)

- [`BaseOxmlElement$remove_attr()`](#method-BaseOxmlElement-remove_attr)

- [`BaseOxmlElement$get_attrs()`](#method-BaseOxmlElement-get_attrs)

- [`BaseOxmlElement$append_child()`](#method-BaseOxmlElement-append_child)

- [`BaseOxmlElement$remove_child()`](#method-BaseOxmlElement-remove_child)

- [`BaseOxmlElement$first_child_found_in()`](#method-BaseOxmlElement-first_child_found_in)

- [`BaseOxmlElement$insert_element_before()`](#method-BaseOxmlElement-insert_element_before)

- [`BaseOxmlElement$remove_all()`](#method-BaseOxmlElement-remove_all)

- [`BaseOxmlElement$xpath()`](#method-BaseOxmlElement-xpath)

- [`BaseOxmlElement$to_xml()`](#method-BaseOxmlElement-to_xml)

- [`BaseOxmlElement$clone()`](#method-BaseOxmlElement-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    BaseOxmlElement$new(node)

------------------------------------------------------------------------

### Method `get_node()`

#### Usage

    BaseOxmlElement$get_node()

------------------------------------------------------------------------

### Method [`find()`](https://rdrr.io/r/utils/apropos.html)

#### Usage

    BaseOxmlElement$find(clark_name)

------------------------------------------------------------------------

### Method `findall()`

#### Usage

    BaseOxmlElement$findall(clark_name)

------------------------------------------------------------------------

### Method `get_attr()`

#### Usage

    BaseOxmlElement$get_attr(attr_name)

------------------------------------------------------------------------

### Method `set_attr()`

#### Usage

    BaseOxmlElement$set_attr(attr_name, value)

------------------------------------------------------------------------

### Method `remove_attr()`

#### Usage

    BaseOxmlElement$remove_attr(attr_name)

------------------------------------------------------------------------

### Method `get_attrs()`

#### Usage

    BaseOxmlElement$get_attrs()

------------------------------------------------------------------------

### Method `append_child()`

#### Usage

    BaseOxmlElement$append_child(child)

------------------------------------------------------------------------

### Method `remove_child()`

#### Usage

    BaseOxmlElement$remove_child(child)

------------------------------------------------------------------------

### Method `first_child_found_in()`

#### Usage

    BaseOxmlElement$first_child_found_in(...)

------------------------------------------------------------------------

### Method `insert_element_before()`

#### Usage

    BaseOxmlElement$insert_element_before(elm, ...)

------------------------------------------------------------------------

### Method `remove_all()`

#### Usage

    BaseOxmlElement$remove_all(...)

------------------------------------------------------------------------

### Method `xpath()`

#### Usage

    BaseOxmlElement$xpath(xpath_str)

------------------------------------------------------------------------

### Method `to_xml()`

#### Usage

    BaseOxmlElement$to_xml()

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    BaseOxmlElement$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
