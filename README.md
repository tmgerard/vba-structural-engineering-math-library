# VBA Structural Engineering Math Library

This library is intended to provide mathmatical tools for practicing structural engineers that like to leverage VBA in spreadsheets for the analysis and design os structures.

## Using the Library

To use the library, import all class and standard modules into a VBA project. For users without the [Rubberduck VBA](https://github.com/rubberduck-vba/Rubberduck) add-in, or who do not need the unit tests, do not import the modules that start with ***Test***.

## Features

The features listed are currently implemented.

### Analytic (Coordinate) Geometry
- 2D
    - Factory2D - Factory methods for the creation of two-dimensional objects
    - Line2D - Line given a base Point2D origin directed by a Vector2D object
    - Point2D - Cartesian coordinate (x, y) point
    - PointPolar - Polar (r, $\theta$) point
    - Polygon2D - Polygonal shape defined by Point2D vertices
    - Segment2D - Line segment between two Point2D objects
    - Vector2D - Vector from origin of cartesian coordinate axis (u, v)
- 3D
    - Point3D - Cartesian coordinate (x, y, z) point
    - Vector3D - Vector from origin of cartesian coordinate axis (u, v, w)

### Linear Algebra
- CholeskySolver - Linear equation solver using cholesky decomposition for a symmetric, positive definite matrix
- LinearAlgebraFactory - Factory methods for the creation of linear algebra objects
- Matrix - Represents a two-dimensional matrix with a given number of rows and columns
- MatrixOperations - Provides operations, such as matrix multiplication and matrix transposes, where a new matrix is returned
- Vector - Represents an row or column vector of a given length

### Utilities
- Arrays - Provides functions used in conjunction with VBA arrays
- Doubles - Provides functions used with double precision floating point values, primarily for defning comparisons between such values
- Math2 - Basic mathmatical functions not provided in VBA's built-in Math library, that are commonly implemented in other languages.

## Testing

Unit tests can be run with the [Rubberduck VBA](https://github.com/rubberduck-vba/Rubberduck) add-in for excel.