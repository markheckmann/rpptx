# Create a lazy active binding function for R6 classes

Returns a function suitable for use in R6 `active` list that evaluates
`fn` on first access and caches the result in private storage.

## Usage

``` r
lazy_active_binding(fn, cache_field)
```

## Arguments

- fn:

  A function taking `self` that computes the value.

- cache_field:

  Name of the private field to use for caching.

## Value

A function suitable for R6 active bindings.
