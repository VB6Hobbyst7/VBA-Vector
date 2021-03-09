# VBA-Vector
Modern way to use arrays in VBA.
VBA-Vector is a custom implementation of 1D array to make process of coding more comfortable.

Getting Started
===============
1. Download the [latest release](https://github.com/vadmitriev/VBA-Vector/releases/), unzip, and import ClassVector.cls into your VBA project.
2. Initialize new variable of ```ClassVector``` and convert standart array (String, Variant, Double, etc.) into this class like it shown below.
```vba
Dim vector  As New ClassVector
Dim arr     As Variant

arr = Array(1, 2, 3)

vector.Convert(arr)
```
Now you can use comfortable syntax, supports such methods as `Add()`, `Count()`, `Delete()`, `Sort()` and other.
 
More information you can find in [documentation](https://github.com/vadmitriev/VBA-Vector/wiki/).

Examples
===============
Initialization of ClassVector item and adding new elements:
```vba
Dim vector      As New ClassVector
Dim arr1        As Variant
dim arr2        As Variant
dim arr3        As Variant

arr1 = Array(1, 2, 3)
arr2 = Array("test1", "test2", "test3")
arr3 = Array(1.11, 2.22, 3.14)

vector.Convert(arr1)

vector.Add(arr2, arr3)
```
Check information about data inside vector:
```vba
' Count items in vector:
vector.count                ' Return 9

' Check vector is empty:
vector.isEmpty              ' Return False

' Summarize number values inside vector:
vector.SumValues            ' Return 12.47

' Check existance particular element in vector:
vector.Exist(3.14)          ' Return True

' Print all values in vector in one line with separator:
vector.toString(", ")       ' Return "1, 2, 3, test1, test2, test3, 1.11, 2.22, 3.14"
```
Adding new elements after/before need index in array:
```vba
vector.insertAfter(1, "after1")
vector.insertBefore(3, "before3")
```

Get access to elements
```vba
' Get one item by index:
vector.item(1)

' Get few items by index:
vector.GetValuesByIndex(1, 2)

' Get only unique values:
vector.GetUniqueValues()

' Get items by required type:
vector.FilterType("string")

' Get values between two indexes:
vector.Slice(0, 2)
```


More information about methods you can find in [documentation](https://github.com/vadmitriev/VBA-Vector/wiki/).
