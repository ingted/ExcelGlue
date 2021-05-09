﻿namespace ExcelGlue

[<RequireQualifiedAccess>]
module Output =
    open System
    open ExcelDna.Integration

    /// Replacement values to return to Excel instead of Double.NaN, None, and [||].
    type ReplaceValues = { nan: obj; none: obj; empty: obj } with
        static member def : ReplaceValues = { nan = ExcelError.ExcelErrorNA; none = "<none>"; empty = "<empty>" }

    /// Returns an xl 1D-range, or a default-singleton instead of an empty array. 
    /// NaN elements are converted according to replaceValues.
    let range<'a> (replaceValues: ReplaceValues) (a1D: 'a[]) : obj[] =
        if a1D |> Array.isEmpty then
            [| replaceValues.empty |]
        else
            if typeof<'a> = typeof<double> then
                a1D |> Array.map (fun num -> let xlval = box num in if Double.IsNaN(xlval :?> double) then replaceValues.nan else xlval)
            else
                a1D |> Array.map box

    /// Returns an xl 1D-range, or a default-singleton instead of an empty array. 
    /// None and NaN elements are converted according to replaceValues.
    let rangeOpt<'a> (replaceValues: ReplaceValues) (a1D: ('a option)[]) : obj[] =
        if a1D |> Array.isEmpty then
            [| replaceValues.empty |]
        else
            if typeof<'a> = typeof<double> then
                a1D 
                |> Array.map 
                    (fun elem -> 
                        match elem with 
                        | None -> replaceValues.none 
                        | Some num -> let xlval = box num in if Double.IsNaN(xlval :?> double) then replaceValues.nan else xlval
                    )
            else
                a1D |> Array.map (fun elem -> match elem with | None -> replaceValues.none | Some e -> box e)

[<RequireQualifiedAccess>]
module Cast0D =
    open System
    open ExcelDna.Integration

    // let private isOptionalType (typeLabel: string) : bool = typeLabel.IndexOf("#") >= 0
    let private prepString (typeLabel: string) =
        typeLabel.Replace(" ", "").Replace(":", "").Replace("#", "").ToUpper()

    /// DU representing xl-value types.
    type Variant = | BOOL | BOOLOPT | STRING | STRINGOPT | DOUBLE | DOUBLEOPT | DOUBLENAN | DOUBLENANOPT | INT | INTOPT | DATE | DATEOPT | VAR | VAROPT | OBJ with
        static member isOptionalType (typeLabel: string) : bool = 
            typeLabel.IndexOf("#") >= 0

        static member ofLabel (typeLabel: string) : Variant = 
            let isoption = Variant.isOptionalType typeLabel
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
            | OBJ-> "OBJ"

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

        static member labelType (typeLabel: string) : Type = 
            let var = Variant.ofLabel typeLabel
            var.toType

        /// Convenience function. Arbitrary defaults.
        member this.defVal : obj = 
            match this with
            | BOOL -> box false
            | BOOLOPT -> (Option.None : bool option) |> box
            | STRING -> box ""
            | STRINGOPT -> (Option.None : string option) |> box
            | DOUBLE -> box 0.0
            | DOUBLEOPT -> (Option.None : double option) |> box
            | DOUBLENAN -> box 0.0
            | DOUBLENANOPT -> (Option.None : double option) |> box
            | INT -> box 0
            | INTOPT -> (Option.None : int option) |> box
            | DATE -> box (DateTime(2000,1,1))
            | DATEOPT -> (Option.None : DateTime option) |> box
            | VAR -> box ExcelError.ExcelErrorNA
            | VAROPT -> (Option.None : obj option) |> box
            | OBJ-> box ExcelError.ExcelErrorNA

        static member labelDefVal (typeLabel: string) : obj = 
            let var = Variant.ofLabel typeLabel
            var.defVal

    /// Indicates which xl-values are to be converted to special values (e.g. Double.NaN in 0D, [||] in 1D) :
    ///    - if OnlyErrorNA, only ExcelErrorNA values are converted.
    ///    - if AllErrors, all Excel error values are converted.
    ///    - if AllNonNumeric, any non-numeric xl-value are.
    type EdgeCaseConversion = | OnlyErrorNA | AllErrors | AllNonNumeric | AllNonNumericAndNA | AllNonNumericAndErrors | NoConversion with
        static member ofLabel (label: string) : EdgeCaseConversion =
            match label.ToUpper() with
            | "NA" -> OnlyErrorNA
            | "ERR" | "ERROR" -> AllErrors
            | "NN" | "NONNUM" | "NONNUMERIC" -> AllNonNumeric
            | "NNNA" | "NN_NA" | "NN+NA" | "NONNUM_NA" | "NONNUM+NA" | "NONNUMERIC_NA" | "NONNUMERIC+NA" -> AllNonNumericAndNA
            | "NNERR" | "NN_ERR" | "NN+ERR" | "NONNUM_ERR" | "NONNUM+ERR" | "NONNUMERIC_ERROR" | "NONNUMERIC+ERROR" -> AllNonNumericAndErrors
            | _ -> NoConversion

        static member labelGuide : (string*string) [] =
            let labels = [| "NA"; "ERR"; "NN"; "NNNA"; "NNERR"; "NOCONVERSION"; "default" |]
            labels |> Array.map (fun lbl -> (lbl, (EdgeCaseConversion.ofLabel lbl).ToString()))

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
        /// Converts xl-values to boxed Double.NaN in special cases.
        let nanify (toNaN: EdgeCaseConversion) (xlVal: obj) : obj = 
            match xlVal, toNaN with
            | :? double, _ -> xlVal

            | :? ExcelError as xlerr, OnlyErrorNA when xlerr = ExcelError.ExcelErrorNA -> box Double.NaN

            | :? ExcelError as xlerr, AllNonNumericAndNA when xlerr = ExcelError.ExcelErrorNA -> box Double.NaN
            | :? ExcelError, AllNonNumericAndNA -> xlVal
            | _, AllNonNumericAndNA -> box Double.NaN

            | :? ExcelError, AllNonNumeric -> xlVal
            | _, AllNonNumeric -> box Double.NaN

            | :? ExcelError, AllErrors -> box Double.NaN
            | _, AllNonNumericAndErrors -> box Double.NaN

            | _ -> xlVal

        /// Casts an xl-value to double or fails, with some other non-double values potentially cast to Double.NaN.
        let fail (toNaN: EdgeCaseConversion) (msg: string option) (xlVal: obj) = 
            nanify toNaN xlVal |> fail<double> msg

        /// Casts an xl-value to double with a default-value, with some other non-double values potentially cast to Double.NaN.
        let def (toNaN: EdgeCaseConversion) (defValue: double) (xlVal: obj) = 
            nanify toNaN xlVal |> def<double> defValue

        /// Casts an xl-value to a double option type with a default-value, with some other non-double values potentially cast to Double.NaN.
        let tryDV (toNaN: EdgeCaseConversion) (defValue: double option) (xlVal: obj) = 
            nanify toNaN xlVal |> tryDV<double> defValue

        /// Replaces an xl-value with a double default-value if it isn't a (boxed double) type (e.g. box 1.0), with some other non-double values potentially cast to Double.NaN.
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
    module Nan =
        /// Converts an obj[] to a double[], given a default-value for non-double elements.
        let def (toNaN: Cast0D.EdgeCaseConversion) (defValue: double) (o1D: obj[]) =
            o1D |> Array.map (Cast0D.Nan.def toNaN defValue)

        /// Converts an obj[] to a ('a option)[], given a default-value for non-double elements.
        let defOpt (toNaN: Cast0D.EdgeCaseConversion) (defValue: double option) (o1D: obj[]) =
            o1D |> Array.map (Cast0D.Nan.tryDV toNaN defValue)

        /// Converts an obj[] to a DateTime[], removing any non-double element.
        let filter (toNaN: Cast0D.EdgeCaseConversion) (o1D: obj[]) =
            o1D |> Array.choose (Cast0D.Nan.tryDV toNaN None)

        /// Converts an obj[] to an optional 'a[]. All the elements must be double, otherwise defValue array is returned. 
        let tryDV (toNaN: Cast0D.EdgeCaseConversion) (defValue: double[] option) (o1D: obj[])  =
            let convert = defOpt toNaN None o1D
            match convert |> Array.tryFind Option.isNone with
            | None -> convert |> Array.map Option.get |> Some
            | Some _ -> defValue

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
    
    [<RequireQualifiedAccess>]
    module Gen = 
        type Gen =
            /// Converts an xl-value to a 'A[], given a default-value for non-'A elements.
            static member def<'A> (rowWiseDef: bool) (defValue: obj) (typeLabel: string) (xlValue: obj) : 'A[] = 
                let o1D = Cast0D.to1D rowWiseDef xlValue

                match typeLabel |> Cast0D.Variant.ofLabel with
                | Cast0D.BOOL -> 
                    let defval = defValue :?> 'A
                    def<'A> defval o1D
                | Cast0D.STRING -> 
                    let defval = defValue :?> string
                    let a1D = Stg.def defval o1D
                    a1D |> Array.map (fun x -> (box x) :?> 'A)
                | Cast0D.DOUBLE -> 
                    let defval = defValue :?> 'A
                    def<'A> defval o1D
                | Cast0D.INT -> 
                    let defval = defValue :?> double
                    let a1D = Intg.def ((int) defval) o1D
                    a1D |> Array.map (fun x -> (box x) :?> 'A)
                | Cast0D.DATE -> 
                    let defval = defValue :?> double
                    let a1D = Dte.def (DateTime.FromOADate(defval)) o1D
                    a1D |> Array.map (fun x -> (box x) :?> 'A)
                | _ -> [||]
        
            /// Same as Gen.def, but returns an obj[] instead of a 'a[].
            static member defObj<'A> (rowWiseDef: bool) (replaceValues: Output.ReplaceValues) (defValue: obj) (typeLabel: string) (xlValue: obj) : obj[] = 
                let a1D = Gen.def<'A> rowWiseDef defValue typeLabel xlValue
                Output.range<'A> replaceValues a1D
    
            /// Converts an xl-value to a ('A option)[], given a default-value for non-'A elements.
            static member defOpt<'A> (rowWiseDef: bool) (defValue: obj option) (typeLabel: string) (xlValue: obj) : ('A option)[] = 
                let o1D = Cast0D.to1D rowWiseDef xlValue

                match typeLabel |> Cast0D.Variant.ofLabel with
                | Cast0D.BOOL -> 
                    let defval = match defValue with | None -> None | Some dfvl -> dfvl :?> 'A option
                    defOpt<'A> defval o1D
                | Cast0D.STRING -> 
                    let defval = match defValue with | None -> None | Some dfvl -> dfvl :?> string option
                    let a1D = Stg.defOpt defval o1D
                    a1D |> Array.map (fun x -> (box x) :?> 'A option)
                | Cast0D.DOUBLE -> 
                    let defval = match defValue with | None -> None | Some dfvl -> dfvl :?> 'A option
                    defOpt<'A> defval o1D
                | Cast0D.INT -> 
                    let defval = match defValue with | None -> None | Some dfvl -> dfvl :?> double option |> Option.map (int)
                    let a1D = Intg.defOpt defval o1D
                    a1D |> Array.map (fun x -> (box x) :?> 'A option)
                | Cast0D.DATE -> 
                    let defval = match defValue with | None -> None | Some dfvl -> dfvl :?> double option |> Option.map (fun d -> DateTime.FromOADate(d))
                    let a1D = Dte.defOpt defval o1D
                    a1D |> Array.map (fun x -> (box x) :?> 'A option)
                | _ -> [||]

            static member defOptObj<'A> (rowWiseDef: bool) (replaceValues: Output.ReplaceValues) (defValue: obj option) (typeLabel: string) (xlValue: obj) : obj[] = 
                let a1D = Gen.defOpt<'A> rowWiseDef defValue typeLabel xlValue
                Output.rangeOpt<'A> replaceValues a1D

            static member filter<'A> (rowWiseDef: bool) (typeLabel: string) (xlValue: obj) : 'A[] = 
                let o1D = Cast0D.to1D rowWiseDef xlValue

                match typeLabel |> Cast0D.Variant.ofLabel with
                | Cast0D.BOOL -> filter<'A> o1D
                | Cast0D.STRING -> 
                    let a1D = Stg.filter o1D
                    a1D |> Array.map (fun x -> (box x) :?> 'A)
                | Cast0D.DOUBLE -> filter<'A> o1D
                | Cast0D.INT -> 
                    let a1D = Intg.filter o1D
                    a1D |> Array.map (fun x -> (box x) :?> 'A)
                | Cast0D.DATE -> 
                    let a1D = Dte.filter o1D
                    a1D |> Array.map (fun x -> (box x) :?> 'A)
                | _ -> [||]

            /// Same as Gen.filter, but returns an obj[] instead of a 'a[].
            static member filterObj<'A> (rowWiseDef: bool) (replaceValues: Output.ReplaceValues) (typeLabel: string) (xlValue: obj) : obj[] = 
                let a1D = Gen.filter<'A> rowWiseDef typeLabel xlValue
                Output.range<'A> replaceValues a1D

        let private callMethod (methodnm: string) (genType: Type) (args: obj[]) : obj =
            let meth = typeof<Gen>.GetMethod(methodnm)
            let genm = meth.MakeGenericMethod(genType)
            let res  = genm.Invoke(null, args)
            res

        let def (rowWiseDef: bool) (defValue: obj) (typeLabel: string) (xlValue: obj) : obj = 
            let gentype = typeLabel |> Cast0D.Variant.labelType
            let args = [| box rowWiseDef; defValue; box typeLabel; xlValue |]
            let res = callMethod "def" gentype args
            res

        let defObj (rowWiseDef: bool) (replaceValues: Output.ReplaceValues) (defValue: obj) (typeLabel: string) (xlValue: obj) : obj[] = 
            let gentype = typeLabel |> Cast0D.Variant.labelType
            let args = [| box rowWiseDef; box replaceValues; defValue; box typeLabel; xlValue |]
            let res = callMethod "defObj" gentype args
            res :?> obj[]
        
        let defOpt (rowWiseDef: bool) (defValue: obj) (typeLabel: string) (xlValue: obj) : obj = 
            let gentype = typeLabel |> Cast0D.Variant.labelType
            let args = [| box rowWiseDef; defValue; box typeLabel; xlValue |]
            let res = callMethod "defOpt" gentype args
            res

        let defOptObj (rowWiseDef: bool) (replaceValues: Output.ReplaceValues) (defValue: obj) (typeLabel: string) (xlValue: obj) : obj[] = 
            let gentype = typeLabel |> Cast0D.Variant.labelType
            let args = [| box rowWiseDef; box replaceValues; defValue; box typeLabel; xlValue |]
            let res = callMethod "defObjOpt" gentype args
            res :?> obj[]

        let filter (rowWiseDef: bool) (typeLabel: string) (xlValue: obj) : obj = 
            let gentype = typeLabel |> Cast0D.Variant.labelType
            let args = [| box rowWiseDef; box typeLabel; xlValue |]
            let res = callMethod "filter" gentype args
            res

        let filterObj (rowWiseDef: bool) (replaceValues: Output.ReplaceValues) (typeLabel: string) (xlValue: obj) : obj[] = 
            let gentype = typeLabel |> Cast0D.Variant.labelType
            let args = [| box rowWiseDef; box replaceValues; box typeLabel; xlValue |]
            let res = callMethod "filterObj" gentype args
            res :?> obj[]





    let _end = "here"


    


//[<RequireQualifiedAccess>]
//module Gen1D =
//    open System

//    type Gen =
//        static member def<'A> (rowWiseDef: bool) (defValue: obj) (typeLabel: string) (xlValue: obj) : 'A[] = 
//            let o1D = Cast0D.to1D rowWiseDef xlValue

//            match typeLabel |> Cast0D.Variant.ofLabel with
//            | Cast0D.BOOL -> 
//                let defval = defValue :?> 'A
//                Cast1D.def<'A> defval o1D
//            | Cast0D.STRING -> 
//                let defval = defValue :?> string
//                let a1D = Cast1D.Stg.def defval o1D
//                a1D |> Array.map (fun x -> (box x) :?> 'A)
//            | Cast0D.DOUBLE -> 
//                let defval = defValue :?> 'A
//                Cast1D.def<'A> defval o1D
//            | Cast0D.INT -> 
//                let defval = defValue :?> double
//                let a1D = Cast1D.Intg.def ((int) defval) o1D
//                a1D |> Array.map (fun x -> (box x) :?> 'A)
//            | Cast0D.DATE -> 
//                let defval = defValue :?> double
//                let a1D = Cast1D.Dte.def (DateTime.FromOADate(defval)) o1D
//                a1D |> Array.map (fun x -> (box x) :?> 'A)
//            | _ -> [||]
        
//        /// Same as Gen.def, but returns an obj[] instead of a 'a[].
//        static member defObj<'A> (rowWiseDef: bool) (replaceValues: Output.ReplaceValues) (defValue: obj) (typeLabel: string) (xlValue: obj) : obj[] = 
//            let a1D = Gen.def<'A> rowWiseDef defValue typeLabel xlValue
//            Output.range<'A> replaceValues a1D
    
//        static member filter<'A> (rowWiseDef: bool) (typeLabel: string) (xlValue: obj) : 'A[] = 
//            let o1D = Cast0D.to1D rowWiseDef xlValue

//            match typeLabel |> Cast0D.Variant.ofLabel with
//            | Cast0D.BOOL -> Cast1D.filter<'A> o1D
//            | Cast0D.STRING -> 
//                let a1D = Cast1D.Stg.filter o1D
//                a1D |> Array.map (fun x -> (box x) :?> 'A)
//            | Cast0D.DOUBLE -> Cast1D.filter<'A> o1D
//            | Cast0D.INT -> 
//                let a1D = Cast1D.Intg.filter o1D
//                a1D |> Array.map (fun x -> (box x) :?> 'A)
//            | Cast0D.DATE -> 
//                let a1D = Cast1D.Dte.filter o1D
//                a1D |> Array.map (fun x -> (box x) :?> 'A)
//            | _ -> [||]

//        /// Same as Gen.filter, but returns an obj[] instead of a 'a[].
//        static member filterObj<'A> (rowWiseDef: bool) (replaceValues: Output.ReplaceValues) (typeLabel: string) (xlValue: obj) : obj[] = 
//            let a1D = Gen.filter<'A> rowWiseDef typeLabel xlValue
//            Output.range<'A> replaceValues a1D

//    let private callMethod (methodnm: string) (genType: Type) (args: obj[]) : obj =
//        let meth = typeof<Gen>.GetMethod(methodnm)
//        let genm = meth.MakeGenericMethod(genType)
//        let res  = genm.Invoke(null, args)
//        res

//    let def (rowWiseDef: bool) (defValue: obj) (typeLabel: string) (xlValue: obj) : obj = 
//        let gentype = typeLabel |> Cast0D.Variant.labelType
//        let args = [| box rowWiseDef; defValue; box typeLabel; xlValue |]
//        let res = callMethod "def" gentype args
//        res

//    let defObj (rowWiseDef: bool) (replaceValues: Output.ReplaceValues) (defValue: obj) (typeLabel: string) (xlValue: obj) : obj[] = 
//        let gentype = typeLabel |> Cast0D.Variant.labelType
//        let args = [| box rowWiseDef; box replaceValues; defValue; box typeLabel; xlValue |]
//        let res = callMethod "defObj" gentype args
//        res :?> obj[]
        
//    let filter (rowWiseDef: bool) (typeLabel: string) (xlValue: obj) : obj = 
//        let gentype = typeLabel |> Cast0D.Variant.labelType
//        let args = [| box rowWiseDef; box typeLabel; xlValue |]
//        let res = callMethod "filter" gentype args
//        res

//    let filterObj (rowWiseDef: bool) (replaceValues: Output.ReplaceValues) (typeLabel: string) (xlValue: obj) : obj[] = 
//        let gentype = typeLabel |> Cast0D.Variant.labelType
//        let args = [| box rowWiseDef; box replaceValues; box typeLabel; xlValue |]
//        let res = callMethod "filterObj" gentype args
//        res :?> obj[]

module Cast_XL =
    open System
    open ExcelDna.Integration

    [<ExcelFunction(Category="XL", Description="Cast an xl-range to DateTime[].")>]
    let cast_edgeCases ()
        : obj[,]  =

        // result
        let (lbls, dus) = Cast0D.EdgeCaseConversion.labelGuide |> Array.map (fun (lbl, du) -> (box lbl, box du)) |> Array.unzip
        [| lbls; dus |] |> array2D

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
        let rplval = { Output.ReplaceValues.def with none = none; empty = empty }

        // result
        let o1D = Cast0D.to1D rowwise range 
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = Cast1D.Bool.filter o1D
                 a1D |> Output.range<bool> rplval
        | "O" -> let a1D = Cast1D.Bool.defOpt None o1D
                 a1D |> Output.rangeOpt<bool> rplval
        | _   -> let a1D = Cast1D.Bool.def defVal o1D 
                 a1D |> Output.range<bool> rplval

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
        let rplval = { Output.ReplaceValues.def with none = none; empty = empty }

        // result
        let o1D = Cast0D.to1D rowwise range 
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = Cast1D.Stg.filter o1D
                 a1D |> Output.range<string> rplval
        | "O" -> let a1D = Cast1D.Stg.defOpt None o1D
                 a1D |> Output.rangeOpt<string> rplval
        | _   -> let a1D = Cast1D.Stg.def defVal o1D 
                 a1D |> Output.range<string> rplval

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
        let rplval = { Output.ReplaceValues.def with none = none; empty = empty }

        // result
        let o1D = Cast0D.to1D rowwise range 
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = Cast1D.Dbl.filter o1D
                 a1D |> Output.range<double> rplval
        | "O" -> let a1D = Cast1D.Dbl.defOpt None o1D
                 a1D |> Output.rangeOpt<double> rplval
        | _   -> let a1D = Cast1D.Dbl.def defVal o1D 
                 a1D |> Output.range<double> rplval

    [<ExcelFunction(Category="XL", Description="Cast an xl-range to an array of doubles (including NaNs).")>]
    let cast1d_dblNan
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "[Replacement method for non-double elements. \"Replace\", \"Optional\" (= replace with None) or \"Filter\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Edge cases mode. Edge cases values are converted to Double.NaN. E.g. NA, ERR, NN, NNNA, NNERR, NONE. Default is NONE.]")>] edgeCase: obj)
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
        let edgecase = Cast0D.Stg.def "NONE" edgeCase |> Cast0D.EdgeCaseConversion.ofLabel
        let defVal = Cast0D.Dbl.def 0.0 defaultValue
        let rplval = { Output.ReplaceValues.def with none = none; empty = empty }

        // result
        let o1D = Cast0D.to1D rowwise range 
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = Cast1D.Nan.filter edgecase o1D
                 a1D |> Output.range<double> rplval
        | "O" -> let a1D = Cast1D.Nan.defOpt edgecase None o1D
                 a1D |> Output.rangeOpt<double> rplval
        | _   -> let a1D = Cast1D.Nan.def edgecase defVal o1D 
                 a1D |> Output.range<double> rplval

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
        let rplval = { Output.ReplaceValues.def with none = none; empty = empty }

        // result
        let o1D = Cast0D.to1D rowwise range 
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = Cast1D.Intg.filter o1D
                 a1D |> Output.range<int> rplval
        | "O" -> let a1D = Cast1D.Intg.defOpt None o1D
                 a1D |> Output.rangeOpt<int> rplval
        | _   -> let a1D = Cast1D.Intg.def defVal o1D 
                 a1D |> Output.range<int> rplval
        
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
        let rplval = { Output.ReplaceValues.def with none = none; empty = empty }

        // result
        let o1D = Cast0D.to1D rowwise range 
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = Cast1D.Dte.filter o1D
                 a1D |> Output.range<DateTime> rplval
        | "O" -> let a1D = Cast1D.Dte.defOpt None o1D
                 a1D |> Output.rangeOpt<DateTime> rplval
        | _   -> let a1D = Cast1D.Dte.def defVal o1D 
                 a1D |> Output.range<DateTime> rplval

    [<ExcelFunction(Category="XL", Description="Cast an xl-range to a generic type array.")>]
    let cast1d_gen
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "Type label.")>] typeLabel: string)
        ([<ExcelArgument(Description= "[Replacement method for non-date elements. \"Replace\", \"Optional\" (= replace with None) or \"Filter\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value. Must be of the appropriate type. Default \"<default>\" (will fail for non-string types).]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowwise = Cast0D.Bool.def false rowWiseDirection
        let replmethod = Cast0D.Stg.def "REPLACE" replaceMethod
        let none = Cast0D.Stg.def "<none>" noneValue
        let empty = Cast0D.Stg.def "<empty>" emptyValue
        let rplval = { Output.ReplaceValues.def with none = none; empty = empty }
        let isoptional = Cast0D.Variant.isOptionalType typeLabel

        let tonone (b: bool) : bool option = if b then None else None

        let defVal = 
            match isoptional with
            | false -> Cast0D.Missing.defO (Cast0D.Variant.labelDefVal typeLabel) defaultValue
            | true ->  (tonone true) |> box // Cast0D.Missing.defO (Cast0D.Variant.labelDefVal typeLabel) defaultValue
            

        match isoptional, replmethod.ToUpper().Substring(0,1) with
        | _, "F" -> Cast1D.Gen.filterObj rowwise rplval typeLabel range
        | true, _ -> Cast1D.Gen.defOptObj rowwise rplval (Some defVal) typeLabel range
        | false, _ -> Cast1D.Gen.defObj rowwise rplval defVal typeLabel range




