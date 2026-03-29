# Read an OPC package (ZIP file)

Provides dict-like access to the binary blobs of each part in the
package.

## Methods

### Public methods

- [`PackageReader$new()`](#method-PackageReader-new)

- [`PackageReader$contains()`](#method-PackageReader-contains)

- [`PackageReader$get_blob()`](#method-PackageReader-get_blob)

- [`PackageReader$rels_xml_for()`](#method-PackageReader-rels_xml_for)

- [`PackageReader$clone()`](#method-PackageReader-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    PackageReader$new(pkg_file)

------------------------------------------------------------------------

### Method `contains()`

#### Usage

    PackageReader$contains(pack_uri)

------------------------------------------------------------------------

### Method `get_blob()`

#### Usage

    PackageReader$get_blob(pack_uri)

------------------------------------------------------------------------

### Method `rels_xml_for()`

#### Usage

    PackageReader$rels_xml_for(partname)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    PackageReader$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
