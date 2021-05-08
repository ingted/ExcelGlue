namespace ExcelGlue

module XlXl =
    open System
    open ExcelDna.Integration

    [<ExcelFunction(Category="XL", Description="Equality of 2 xl-values (\'variants\').")>]
    let x_eq 
        ([<ExcelArgument(Description= "Value 1.")>] val1: obj) 
        ([<ExcelArgument(Description= "Value 2.")>] val2: obj) 
        : obj  =

        (val1 = val2) |> box

