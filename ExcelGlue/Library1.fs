namespace ExcelGlue

type Class1() = 
    member this.X = "F#"

module FOO =
    open ExcelDna.Integration

    [<ExcelFunction(Description="My first .NET function")>]
    let HelloDna name =
        "Hello " + name