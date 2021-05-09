namespace ExcelGlue

[<RequireQualifiedAccess>]
module Cast0D =
    open System
    open ExcelDna.Integration

    /// Indicates which xl-values are to be converted to special values (e.g. Double.NaN in 0D, [||] in 1D) :
    ///    - if OnlyErrorNA, only ExcelErrorNA values are converted.
    ///    - if AllErrors, all Excel error values are converted.
    ///    - if AllNonNumeric, any non-numeric xl-value are.
    type EdgeCaseConversion = | OnlyErrorNA | AllErrors | AllNonNumeric | AllNonNumericAndNA | AllNonNumericAndErrors | NoConversion

    // ----------------
    // -- 0D functions
    // ----------------

    /// Casts an obj to generic type or fails.
    let fail<'a> (msg: string option) (o: obj) : 'a =
        match o with
        | :? 'a as v -> v
        | _ -> failwith (msg |> Option.defaultValue "Cast failed.")

    /// Casts an obj to generic type with typed default-value.
    let def<'a> (defValue: 'a) (o: obj) : 'a =
        match o with
        | :? 'a as v -> v
        | _ -> defValue

    /// Casts an obj to an option on generic type with typed default-value.
    let tryDV<'a> (defValue: 'a option) (o: obj) : 'a option =
        match o with
        | :? 'a as v -> Some v
        | _ -> defValue

    let map<'a> (mapping: obj -> 'a) (o: obj) : 'a = mapping o // TODO remove?

    /// Replaces an obj with typed default-value if it isn't of given generic type 'a.
    let defO<'a> (defValue: obj) (o: obj) : obj =
        match o with
        | :? 'a -> o
        | _ -> defValue

    // ----------------
    // -- 1D functions
    // ----------------

    /// Slices xl-ranges into 1D arrays.
    /// 1-row or 1-column xl-ranges are directly converted to their corresponding 1D array.
    /// For 2D xl-range, the 1st row is returned if rowWiseDef is true, and the 1st column otherwise.
    let slice2D (rowWiseDef: bool) (o2D: obj[,]) : obj[] = 
        // column-wise slice
        if Array2D.length2 o2D = 1 then
            o2D.[*, Array2D.base2 o2D]
        // row-wise slice
        elif rowWiseDef || (Array2D.length1 o2D = 1) then
            o2D.[Array2D.base1 o2D, *]
        // column-wise slice as default
        else 
            o2D.[*, Array2D.base2 o2D]
        
    /// Converts an xl-value to a 1D array.
    /// (Use to1D rather than try1D when the obj argument is an xl-value).
    let to1D (rowWiseDef: bool) (xlVal: obj) : obj[] =
        match xlVal with
        | :? (obj[,]) as o2D -> slice2D rowWiseDef o2D
        | :? (obj[]) as o1D -> o1D
        | o0D -> [| o0D |]
        
    /// Converts an obj to a 1D array option.
    /// (Use try1D rather than to1D when the obj argument is not an xl-value).
    let try1D (rowWiseDef: bool) (o: obj) : obj[] option =
        match o with
        | :? (obj[,]) as o2D -> slice2D rowWiseDef o2D |> Some
        | :? (obj[]) as o1D -> Some o1D
        | _ -> None

    [<RequireQualifiedAccess>]
    module Bool =
        /// Casts an xl-value to bool or fails.
        let fail (msg: string option) (xlVal: obj) = fail<bool> msg xlVal

        /// Casts an xl-value to bool with a default-value.
        let def (defValue: bool) (xlVal: obj) = def<bool> defValue xlVal

        /// Casts an xl-value to a bool option type with a default-value.
        let tryDV (defValue: bool option) (xlVal: obj) = tryDV<bool> defValue xlVal

        /// Replaces an xl-value with a (boxed bool) default-value if it isn't a (boxed bool) type (e.g. box true).
        let defO (defValue: bool) (xlVal: obj) = defO<bool> defValue xlVal

    [<RequireQualifiedAccess>]
    module Stg =
        /// Casts an xl-value to string or fails.
        let fail (msg: string option) (xlVal: obj) = fail<string> msg xlVal

        /// Casts an xl-value to string with a default-value.
        let def (defValue: string) (xlVal: obj) = def<string> defValue xlVal

        /// Casts an xl-value to a string option type with a default-value.
        let tryDV (defValue: string option) (xlVal: obj) = tryDV<string> defValue xlVal

        /// Replaces an xl-value with a (boxed string) default-value if it isn't a (boxed string) type (e.g. box "foo").
        let defO (defValue: string) (xlVal: obj) = defO<string> defValue xlVal

    [<RequireQualifiedAccess>]
    module Dbl =
        /// Casts an xl-value to double or fails.
        let fail (msg: string option) (xlVal: obj) = fail<double> msg xlVal

        /// Casts an xl-value to double with a default-value.
        let def (defValue: double) (xlVal: obj) = def<double> defValue xlVal

        /// Casts an xl-value to a double option type with a default-value.
        let tryDV (defValue: double option) (xlVal: obj) = tryDV<double> defValue xlVal

        /// Replaces an xl-value with a (boxed double) default-value if it isn't a (boxed double) type (e.g. box 1.0).
        let defO (defValue: double) (xlVal: obj) = defO<double> defValue xlVal  // TODO FIXME defValue type is obj not double

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
            nanify toNaN xlVal |> fail<double> msg

        /// Casts an xl-value to double with a default-value, with some values potentially cast to Double.NaN.
        let def (toNaN: EdgeCaseConversion) (defValue: double) (xlVal: obj) = 
            nanify toNaN xlVal |> def<double> defValue

        /// Casts an xl-value to a double option type with a default-value, with some values potentially cast to Double.NaN.
        let tryDV (toNaN: EdgeCaseConversion) (defValue: double option) (xlVal: obj) = 
            nanify toNaN xlVal |> tryDV<double> defValue

        /// Replaces an xl-value with a double default-value if it isn't a (boxed double) type (e.g. box 1.0), with some values potentially cast to Double.NaN.
        let defO (toNaN: EdgeCaseConversion) (defValue: double) (xlVal: obj) = 
            nanify toNaN xlVal |> defO<double> defValue

        /// Converts a boxed Double.NaN into an ExcelErrorNA.
        let ofNaN (xlVal: obj) : obj =
            match xlVal with
            | :? double as d -> if Double.IsNaN d then ExcelError.ExcelErrorNA |> box else box d
            | _ -> xlVal

    [<RequireQualifiedAccess>]
    module Intg =
        let ofDouble (d: double) : int option = 
            let floor = Math.Floor d
            if d = floor then (int) floor |> Some else None

        // for functions matching on xlVal type below, no need to test xlVal as a int as numeric xl-values are doubles, not int.

        /// Casts an xl-value to int or fails.
        let fail (msg: string option) (xlVal: obj) = // fail<int> msg xlVal
            match xlVal with
            | :? double as d -> match ofDouble d with | Some i -> i | None -> failwith (msg |> Option.defaultValue "Cast failed.")
            | _ -> failwith (msg |> Option.defaultValue "Cast failed.")

        /// Casts an xl-value to int with a default-value.
        let def (defValue: int) (xlVal: obj) =
            match xlVal with
            | :? double as d -> ofDouble d |> Option.defaultValue defValue
            | _ -> defValue

        /// Casts an xl-value to a int option type with a default-value.
        let tryDV (defValue: int option) (xlVal: obj) =
            match xlVal with
            | :? double as d -> match ofDouble d with | None -> defValue | Some i -> Some i
            | _ -> defValue

        /// Replaces an xl-value with a (boxed int) default-value if it isn't a (boxed int) type (e.g. box 42).
        let defO (defValue: obj) (xlVal: obj) =
            match xlVal with
            | :? double as d -> match ofDouble d with | None -> defValue | Some i -> box i
            | _ -> defValue

    [<RequireQualifiedAccess>]
    module Dte =
        /// Casts an xl-value to a DateTime or fails.
        let fail (msg: string option) (xlVal: obj) =
            match xlVal with
            | :? double as d -> DateTime.FromOADate d
            | _ -> failwith (msg |> Option.defaultValue "Cast failed.")

        /// Casts an xl-value to a DateTime with a default-value.
        let def (defValue: DateTime) (xlVal: obj) =
            match xlVal with
            | :? double as d -> DateTime.FromOADate d
            | _ -> defValue

        /// Casts an xl-value to a DateTime option type with a default-value.
        let tryDV (defValue: DateTime option) (xlVal: obj) =
            match xlVal with
            | :? double as d -> DateTime.FromOADate d |> Some
            | _ -> defValue

        /// Replaces an xl-value with a default-value if it isn't a (boxed DateTime) type (e.g. box 36526.0).
        let defO (defValue: obj) (xlVal: obj) =
            match xlVal with
            | :? double as d -> xlVal
            | _ -> defValue

    [<RequireQualifiedAccess>]
    module Missing =
        /// Casts an obj to generic type with typed default-value.
        /// If missing, also returns the default-value.
        let def<'a> (defValue: 'a) (xlVal: obj) : 'a =
            match xlVal with
            | :? ExcelMissing -> defValue
            | _ -> def<'a> defValue xlVal

        /// Replaces an xl-value with an obj default-value if missing.
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

        /// Applies a map to an xl-value, and replaces missing values with a typed default-value.
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
module Cast1D =
    open System

    /// Converts an obj[] to a 'a[], given a typed default-value for elements which can't be cast to 'a.
    let def<'a> (defValue: 'a) (o1D: obj[]) : 'a[] =
        o1D |> Array.map (Cast0D.def<'a> defValue)

    /// Converts an obj[] to a ('a option)[], given a typed default-value for elements which can't be cast to 'a.
    let defOpt<'a> (defValue: 'a option) (o1D: obj[]) : ('a option)[] =
        o1D |> Array.map (Cast0D.tryDV<'a> defValue)

    /// Converts an obj[] to a 'a[], removing any element which can't be cast to 'a.
    let filter<'a> (o1D: obj[]) : 'a[] =
        o1D |> Array.choose (Cast0D.tryDV<'a> None)

    /// Converts an obj[] to an optional 'a[]. All the elements must match the given type, otherwise defValue array is returned. 
    let tryDV<'a> (defValue: 'a[] option) (o1D: obj[]) : 'a[] option =
        let convert = defOpt None o1D
        match convert |> Array.tryFind Option.isNone with
        | None -> convert |> Array.map Option.get |> Some
        | Some _ -> defValue

    [<RequireQualifiedAccess>]
    module Bool =
        /// Converts an obj[] to a bool[], given a default-value for non-bool elements.
        let def (defValue: bool) (o1D: obj[]) = def defValue o1D

        /// Converts an obj[] to a ('a option)[], given a default-value for non-bool elements.
        let defOpt (defValue: bool option) (o1D: obj[]) = defOpt defValue o1D

        /// Converts an obj[] to a bool[], removing any non-bool element.
        let filter (o1D: obj[]) = filter<bool> o1D

        /// Converts an obj[] to an optional 'a[]. All the elements must be bool, otherwise defValue array is returned. 
        let tryDV (defValue: bool[] option) (o1D: obj[])  = tryDV<bool> defValue o1D

    [<RequireQualifiedAccess>]
    module Stg =
        /// Converts an obj[] to a string[], given a default-value for non-string elements.
        let def (defValue: string) (o1D: obj[]) = def defValue o1D

        /// Converts an obj[] to a ('a option)[], given a default-value for non-string elements.
        let defOpt (defValue: string option) (o1D: obj[]) = defOpt defValue o1D

        /// Converts an obj[] to a string[], removing any non-string element.
        let filter (o1D: obj[]) = filter<string> o1D

        /// Converts an obj[] to an optional 'a[]. All the elements must be string, otherwise defValue array is returned. 
        let tryDV (defValue: string[] option) (o1D: obj[])  = tryDV<string> defValue o1D

    [<RequireQualifiedAccess>]
    module Dbl =
        /// Converts an obj[] to a double[], given a default-value for non-double elements.
        let def (defValue: double) (o1D: obj[]) = def defValue o1D

        /// Converts an obj[] to a ('a option)[], given a default-value for non-double elements.
        let defOpt (defValue: double option) (o1D: obj[]) = defOpt defValue o1D

        /// Converts an obj[] to a double[], removing any non-double element.
        let filter (o1D: obj[]) = filter<double> o1D

        /// Converts an obj[] to an optional 'a[]. All the elements must be double, otherwise defValue array is returned. 
        let tryDV (defValue: double[] option) (o1D: obj[])  = tryDV<double> defValue o1D

    [<RequireQualifiedAccess>]
    module Intg =
        /// Converts an obj[] to a int[], given a default-value for non-int elements.
        let def (defValue: int) (o1D: obj[]) =
            o1D |> Array.map (Cast0D.Intg.def defValue)

        /// Converts an obj[] to a ('a option)[], given a default-value for non-int elements.
        let defOpt (defValue: int option) (o1D: obj[]) =
            o1D |> Array.map (Cast0D.Intg.tryDV defValue)

        /// Converts an obj[] to a int[], removing any non-int element.
        let filter (o1D: obj[]) =
            o1D |> Array.choose (Cast0D.Intg.tryDV None)

        /// Converts an obj[] to an optional 'a[]. All the elements must be int, otherwise defValue array is returned. 
        let tryDV (defValue: int[] option) (o1D: obj[])  =
            let convert = defOpt None o1D
            match convert |> Array.tryFind Option.isNone with
            | None -> convert |> Array.map Option.get |> Some
            | Some _ -> defValue

    [<RequireQualifiedAccess>]
    module Dte =
        /// Converts an obj[] to a DateTime[], given a default-value for non-DateTime elements.
        let def (defValue: DateTime) (o1D: obj[]) =
            o1D |> Array.map (Cast0D.Dte.def defValue)

        /// Converts an obj[] to a ('a option)[], given a default-value for non-DateTime elements.
        let defOpt (defValue: DateTime option) (o1D: obj[]) =
            o1D |> Array.map (Cast0D.Dte.tryDV defValue)

        /// Converts an obj[] to a DateTime[], removing any non-DateTime element.
        let filter (o1D: obj[]) =
            o1D |> Array.choose (Cast0D.Dte.tryDV None)

        /// Converts an obj[] to an optional 'a[]. All the elements must be DateTime, otherwise defValue array is returned. 
        let tryDV (defValue: DateTime[] option) (o1D: obj[])  =
            let convert = defOpt None o1D
            match convert |> Array.tryFind Option.isNone with
            | None -> convert |> Array.map Option.get |> Some
            | Some _ -> defValue

    // -----------------------------------
    // -- Convenience functions
    // -----------------------------------

    /// Returns a default-singleton instead of an empty array. 
    let ofEmpty<'a> (sglValue: obj) (o1d: 'a[]) : obj[] =
        if o1d |> Array.isEmpty then
            [| sglValue |]
        else
            o1d |> Array.map box

    /// Returns a default-singleton instead of an empty array. 
    /// None elements are converted to noneValue.
    let ofEmptyOpt<'a> (noneValue: obj) (sglValue: obj) (o1d: ('a option)[]) : obj[] =
        if o1d |> Array.isEmpty then
            [| sglValue |]
        else
            o1d |> Array.map (fun elem -> match elem with | None -> noneValue | Some e -> box e)
            


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

        member this.toLabel : string = 
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

    [<ExcelFunction(Category="XL", Description="Cast an xl-range to obj[]")>]
    let cast1d_obj
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowWiseDef = Cast0D.Bool.def false rowWiseDirection

        // result
        Cast0D.to1D rowWiseDef range

    [<ExcelFunction(Category="XL", Description="Cast an xl-range to bool[].")>]
    let cast1d_bool
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "[Replacement method for non-bool elements. \"Replace\", \"Optional\" (= replace with None) or \"Filter\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value. Default is FALSE.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowwise = Cast0D.Bool.def false rowWiseDirection
        let replmethod = Cast0D.Stg.def "REPLACE" replaceMethod
        let none = Cast0D.Stg.def "<none>" noneValue
        let empty = Cast0D.Stg.def "<empty>" emptyValue
        let defVal = Cast0D.Bool.def false defaultValue

        // result
        let o1D = Cast0D.to1D rowwise range 
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = Cast1D.Bool.filter o1D
                 a1D |> Cast1D.ofEmpty<bool> empty
        | "O" -> let a1D = Cast1D.Bool.defOpt None o1D
                 a1D |> Cast1D.ofEmptyOpt<bool> none empty
        | _   -> let a1D = Cast1D.Bool.def defVal o1D 
                 a1D |> Cast1D.ofEmpty<bool> empty

    [<ExcelFunction(Category="XL", Description="Cast an xl-range to string[].")>]
    let cast1d_stg
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "[Replacement method for non-string elements. \"Replace\", \"Optional\" (= replace with None) or \"Filter\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value. Default is \"-foo-\".]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowwise = Cast0D.Bool.def false rowWiseDirection
        let replmethod = Cast0D.Stg.def "REPLACE" replaceMethod
        let none = Cast0D.Stg.def "<none>" noneValue
        let empty = Cast0D.Stg.def "<empty>" emptyValue
        let defVal = Cast0D.Stg.def "-foo-" defaultValue

        // result
        let o1D = Cast0D.to1D rowwise range 
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = Cast1D.Stg.filter o1D
                 a1D |> Cast1D.ofEmpty<string> empty
        | "O" -> let a1D = Cast1D.Stg.defOpt None o1D
                 a1D |> Cast1D.ofEmptyOpt<string> none empty
        | _   -> let a1D = Cast1D.Stg.def defVal o1D 
                 a1D |> Cast1D.ofEmpty<string> empty

    [<ExcelFunction(Category="XL", Description="Cast an xl-range to double[].")>]
    let cast1d_dbl
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "[Replacement method for non-double elements. \"Replace\", \"Optional\" (= replace with None) or \"Filter\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value. Default is 0.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowwise = Cast0D.Bool.def false rowWiseDirection
        let replmethod = Cast0D.Stg.def "REPLACE" replaceMethod
        let none = Cast0D.Stg.def "<none>" noneValue
        let empty = Cast0D.Stg.def "<empty>" emptyValue
        let defVal = Cast0D.Dbl.def 0.0 defaultValue

        // result
        let o1D = Cast0D.to1D rowwise range 
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = Cast1D.Dbl.filter o1D
                 a1D |> Cast1D.ofEmpty<double> empty
        | "O" -> let a1D = Cast1D.Dbl.defOpt None o1D
                 a1D |> Cast1D.ofEmptyOpt<double> none empty
        | _   -> let a1D = Cast1D.Dbl.def defVal o1D 
                 a1D |> Cast1D.ofEmpty<double> empty

    [<ExcelFunction(Category="XL", Description="Cast an xl-range to int[].")>]
    let cast1d_int
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "[Replacement method for non-integer elements. \"Replace\", \"Optional\" (= replace with None) or \"Filter\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value. Default is 0.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowwise = Cast0D.Bool.def false rowWiseDirection
        let replmethod = Cast0D.Stg.def "REPLACE" replaceMethod
        let none = Cast0D.Stg.def "<none>" noneValue
        let empty = Cast0D.Stg.def "<empty>" emptyValue
        let defVal = Cast0D.Intg.def 0 defaultValue

        // result
        let o1D = Cast0D.to1D rowwise range 
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = Cast1D.Intg.filter o1D
                 a1D |> Cast1D.ofEmpty<int> empty
        | "O" -> let a1D = Cast1D.Intg.defOpt None o1D
                 a1D |> Cast1D.ofEmptyOpt<int> none empty
        | _   -> let a1D = Cast1D.Intg.def defVal o1D 
                 a1D |> Cast1D.ofEmpty<int> empty
        
    [<ExcelFunction(Category="XL", Description="Cast an xl-range to DateTime[].")>]
    let cast1d_dte
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "[Replacement method for non-date elements. \"Replace\", \"Optional\" (= replace with None) or \"Filter\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value. Default is 1-Jan-2000.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowwise = Cast0D.Bool.def false rowWiseDirection
        let replmethod = Cast0D.Stg.def "REPLACE" replaceMethod
        let none = Cast0D.Stg.def "<none>" noneValue
        let empty = Cast0D.Stg.def "<empty>" emptyValue
        let defVal = Cast0D.Dte.def (DateTime(2000,1,1)) defaultValue

        // result
        let o1D = Cast0D.to1D rowwise range 
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = Cast1D.Dte.filter o1D
                 a1D |> Cast1D.ofEmpty<DateTime> empty
        | "O" -> let a1D = Cast1D.Dte.defOpt None o1D
                 a1D |> Cast1D.ofEmptyOpt<DateTime> none empty
        | _   -> let a1D = Cast1D.Dte.def defVal o1D 
                 a1D |> Cast1D.ofEmpty<DateTime> empty






