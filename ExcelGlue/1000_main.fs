namespace ExcelGlue

[<RequireQualifiedAccess>]
module Cast =
    open System
    open ExcelDna.Integration

    /// Indicates which xl-values are to be converted to special values (e.g. Double.NaN in 0D, [||] in 1D) :
    ///    - if OnlyErrorNA, only ExcelErrorNA values are converted.
    ///    - if AllErrors, all Excel error values are converted.
    ///    - if AllNonNumeric, any non-numeric xl-value are.
    type EdgeCaseConversion = | OnlyErrorNA | AllErrors | AllNonNumeric | AllNonNumericAndNA | AllNonNumericAndErrors | NoConversion

    [<RequireQualifiedAccess>]
    module Cast =
        /// Casts an obj to generic type or fails.
        let fail<'a> (msg: string option) (o: obj) : 'a =
            match o with
            | :? 'a as v -> v
            | _ -> failwith (msg |> Option.defaultValue "Cast failed.")

        /// Casts an obj to generic type with typed default value.
        let def<'a> (defValue: 'a) (o: obj) : 'a =
            match o with
            | :? 'a as v -> v
            | _ -> defValue

        /// Casts an obj to an option on generic type with typed default value.
        let tryDV<'a> (defValue: 'a option) (o: obj) : 'a option =
            match o with
            | :? 'a as v -> Some v
            | _ -> defValue

        let map<'a> (mapping: obj -> 'a) (o: obj) : 'a = mapping o // TODO remove?

        /// Replaces an obj with typed default value if it isn't of given generic type 'a.
        let defO<'a> (defValue: obj) (o: obj) : obj =
            match o with
            | :? 'a -> o
            | _ -> defValue

    [<RequireQualifiedAccess>]
    module Bool =
        /// Casts an xl-value to bool or fails.
        let fail (msg: string option) (xlVal: obj) = Cast.fail<bool> msg xlVal

        /// Casts an xl-value to bool with a default value.
        let def (defValue: bool) (xlVal: obj) = Cast.def<bool> defValue xlVal

        /// Casts an xl-value to a bool option type with a default value.
        let tryDV (defValue: bool option) (xlVal: obj) = Cast.tryDV<bool> defValue xlVal

        /// Replaces an xl-value with a bool default value if it isn't a boxed bool type (e.g. box true).
        let defO (defValue: bool) (xlVal: obj) = Cast.def<bool> defValue xlVal

    [<RequireQualifiedAccess>]
    module Stg =
        /// Casts an xl-value to string or fails.
        let fail (msg: string option) (xlVal: obj) = Cast.fail<string> msg xlVal

        /// Casts an xl-value to string with a default value.
        let def (defValue: string) (xlVal: obj) = Cast.def<string> defValue xlVal

        /// Casts an xl-value to a string option type with a default value.
        let tryDV (defValue: string option) (xlVal: obj) = Cast.tryDV<string> defValue xlVal

        /// Replaces an xl-value with a string default value if it isn't a boxed string type (e.g. box true).
        let defO (defValue: string) (xlVal: obj) = Cast.def<string> defValue xlVal

    [<RequireQualifiedAccess>]
    module Dbl =
        /// Casts an xl-value to double or fails.
        let fail (msg: string option) (xlVal: obj) = Cast.fail<double> msg xlVal

        /// Casts an xl-value to double with a default value.
        let def (defValue: double) (xlVal: obj) = Cast.def<double> defValue xlVal

        /// Casts an xl-value to a double option type with a default value.
        let tryDV (defValue: double option) (xlVal: obj) = Cast.tryDV<double> defValue xlVal

        /// Replaces an xl-value with a double default value if it isn't a boxed double type (e.g. box true).
        let defO (defValue: double) (xlVal: obj) = Cast.def<double> defValue xlVal

    [<RequireQualifiedAccess>]
    module Nan =
        /// Converts an xl-value to boxed Double.NaN in some cases.
        let nanify (toNaN: EdgeCaseConversion) (xlVal: obj) : obj = 
            match xlVal, toNaN with
            | :? double, _ -> xlVal
            | :? ExcelError as xlerr, OnlyErrorNA when xlerr = ExcelError.ExcelErrorNA -> box Double.NaN
            | :? ExcelError as xlerr, AllNonNumericAndNA when xlerr = ExcelError.ExcelErrorNA -> box Double.NaN
            | :? ExcelError, AllErrors -> box Double.NaN
            | :? ExcelError, AllNonNumericAndNA -> box Double.NaN
            | :? ExcelError, AllNonNumericAndErrors -> box Double.NaN
            | _, AllNonNumeric -> box Double.NaN
            | _, AllNonNumericAndNA -> box Double.NaN
            | _, AllNonNumericAndErrors -> box Double.NaN
            | _ -> xlVal

        /// Casts an xl-value to double or fails, with some values potentially cast to Double.NaN.
        let fail (toNaN: EdgeCaseConversion) (msg: string option) (xlVal: obj) = 
            nanify toNaN xlVal |> Cast.fail<double> msg

        /// Casts an xl-value to double with a default value, with some values potentially cast to Double.NaN.
        let def (toNaN: EdgeCaseConversion) (defValue: double) (xlVal: obj) = 
            nanify toNaN xlVal |> Cast.def<double> defValue

        /// Casts an xl-value to a double option type with a default value, with some values potentially cast to Double.NaN.
        let tryDV (toNaN: EdgeCaseConversion) (defValue: double option) (xlVal: obj) = 
            nanify toNaN xlVal |> Cast.tryDV<double> defValue

        /// Replaces an xl-value with a double default value if it isn't a boxed double type (e.g. box true), with some values potentially cast to Double.NaN.
        let defO (toNaN: EdgeCaseConversion) (defValue: double) (xlVal: obj) = 
            nanify toNaN xlVal |> Cast.def<double> defValue

        /// Converts a boxed Double.NaN into an ExcelErrorNA.
        let ofNaN (xlVal: obj) : obj =
            match xlVal with
            | :? double as d -> if Double.IsNaN d then ExcelError.ExcelErrorNA |> box else box d
            | _ -> xlVal

    [<RequireQualifiedAccess>]
    module Intg =
        /// Casts an xl-value to int or fails.
        let fail (msg: string option) (xlVal: obj) = Cast.fail<int> msg xlVal

        /// Casts an xl-value to int with a default value.
        let def (defValue: int) (xlVal: obj) = Cast.def<int> defValue xlVal

        /// Casts an xl-value to a int option type with a default value.
        let tryDV (defValue: int option) (xlVal: obj) = Cast.tryDV<int> defValue xlVal

        /// Replaces an xl-value with a int default value if it isn't a boxed int type (e.g. box true).
        let defO (defValue: int) (xlVal: obj) = Cast.def<int> defValue xlVal

    [<RequireQualifiedAccess>]
    module Missing =
        /// Casts an obj to generic type with typed default value.
        /// If missing, also returns the default value.
        let def<'a> (defValue: 'a) (xlVal: obj) : 'a =
            match xlVal with
            | :? ExcelMissing -> defValue
            | _ -> Cast.def<'a> defValue xlVal

        /// Replaces an xl-value with an obj default value if missing.
        /// Otherwise passes the xl-value through.
        let defO (defValue: obj) (o: obj) : obj =
            match o with
            | :? ExcelMissing -> defValue
            | _ -> o

        /// Replaces an xl-value with None if missing.
        /// Otherwise passes the xl-value through.
        let tryO (o: obj) : obj option =
            match o with
            | :? ExcelMissing -> None
            | _ -> Some o

        /// Applies a map to an xl-value, and replaces missing values with a typed default value.
        let map<'a> (defValue: 'a) (mapping: obj -> 'a) (xlVal: obj) : 'a =
            match xlVal with
            | :? ExcelMissing -> defValue
            | _ -> mapping xlVal

        /// Applies a map to an xl-value, but returns None for missing values.
        let tryMap<'a> (mapping: obj -> 'a option) (xlVal: obj) : 'a option =
            match xlVal with
            | :? ExcelMissing -> None
            | _ -> mapping xlVal

    [<RequireQualifiedAccess>]
    module D1 =
        /// Converts a 1-row or 1-column xl-range into a 1D array.
        /// In case of non-1D xl-range, returns the 1st row if rowWiseDef and the 1st column otherwise.
        let of2D (rowWiseDef: bool) (o2D: obj[,]) : obj[] = 
            // column-wise slice
            if Array2D.length2 o2D = 1 then
                o2D.[*, Array2D.base2 o2D]
            // row-wise slice
            elif rowWiseDef || (Array2D.length1 o2D = 1) then
                o2D.[Array2D.base1 o2D, *]
            // column-wise slice as default
            else 
                o2D.[*, Array2D.base2 o2D]
        
        /// Converts an xl-range to a 1D array option.
        /// Use to1D over try1D when the obj argument is an xl-value.
        let to1D (rowWiseDef: bool) (xlVal: obj) : obj[] =
            match xlVal with
            | :? (obj[,]) as o2D -> of2D rowWiseDef o2D
            | :? (obj[]) as o1D -> o1D
            | o0D -> [| o0D |]
        
        /// Converts an xl-range to a 1D array option.
        /// Use try1D over to1D when the obj argument is not an xl-value.
        let try1D (rowWiseDef: bool) (o: obj) : obj[] option =
            match o with
            | :? (obj[,]) as o2D -> of2D rowWiseDef o2D |> Some
            | :? (obj[]) as o1D -> Some o1D
            | _ -> None

        // -----------------------------------
        // -- Convenience functions
        // -----------------------------------

        /// Returns a default value instead of an empty array. 
        let ofEmpty<'a> (defValue: obj) (o1d: 'a[]) : obj[] =
            if o1d |> Array.isEmpty then
                [| defValue |]
            else
                o1d |> Array.map box

        /// Returns a default value instead of an empty array.
        let ofEmptyO (defValue: obj) (o1d: obj[]) : obj[] =
            if o1d |> Array.isEmpty then
                [| defValue |]
            else
                o1d
            


        let _end = "here"

    let private isOptionalType (typeLabel: string) : bool = typeLabel.IndexOf("#") >= 0
    let private prepString (typeLabel: string) =
        typeLabel.Replace(" ", "").Replace(":", "").Replace("#", "").ToUpper()

    /// DU representing xl-value types.
    type Variant = | BOOL | BOOLOPT | STRING | STRINGOPT | DOUBLE | DOUBLEOPT | DOUBLENAN | DOUBLENANOPT | INT | INTOPT | DATE | DATEOPT | VAR | VAROPT | OBJ with
        static member ofLabel (typeLabel: string) : Variant = 
            let isoption = isOptionalType typeLabel
            match prepString typeLabel with
            | "B" | "BOOL" | "BOOLEAN" -> if isoption then BOOLOPT else BOOL
            | "S" | "STR" | "STG" | "STRG" | "STRING" -> if isoption then STRINGOPT else STRING
            | "D" | "DBL" | "DOUBLE" -> if isoption then DOUBLEOPT else DOUBLE
            | "DNAN" | "DBLNAN" | "DOUBLENAN" -> if isoption then DOUBLENANOPT else DOUBLENAN
            | "I" | "INT" | "INTEGER" -> if isoption then INTOPT else INT
            | "DTE" | "DATE" -> if isoption then DATEOPT else DATE
            | "V" | "VAR" -> if isoption then VAROPT else VAR
            | _ -> OBJ

        member this.toLabel : String = 
            match this with
            | BOOL -> "BOOL"
            | BOOLOPT -> "#BOOL"
            | STRING -> "STRING"
            | STRINGOPT -> "#STRING"
            | DOUBLE -> "DOUBLE"
            | DOUBLEOPT -> "#DOUBLE"
            | DOUBLENAN -> "DOUBLENAN"
            | DOUBLENANOPT -> "#DOUBLENAN"
            | INT -> "INT"
            | INTOPT -> "#INT"
            | DATE -> "DATE"
            | DATEOPT -> "#DATE"
            | VAR -> "VAR"
            | VAROPT -> "#VAR"
            | OBJ-> "PBJ"

        member this.toType : Type = 
            match this with
            | BOOL -> typeof<bool>
            | BOOLOPT -> typeof<bool option>
            | STRING -> typeof<string>
            | STRINGOPT -> typeof<string option>
            | DOUBLE -> typeof<double>
            | DOUBLEOPT -> typeof<double option>
            | DOUBLENAN -> typeof<double>
            | DOUBLENANOPT -> typeof<double option>
            | INT -> typeof<int>
            | INTOPT -> typeof<int option>
            | DATE -> typeof<DateTime>
            | DATEOPT -> typeof<DateTime option>
            | VAR -> typeof<obj>
            | VAROPT -> typeof<obj option>
            | OBJ-> typeof<obj>


    [<RequireQualifiedAccess>]
    module Variant =
        let x = 2

module Cast_XL =
    open System
    open ExcelDna.Integration

    [<ExcelFunction(Category="XL", Description="Cast.")>]
    let cast_o1d
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowWiseDef = Cast.Bool.def false rowWiseDirection

        // result
        Cast.D1.to1D rowWiseDef range









