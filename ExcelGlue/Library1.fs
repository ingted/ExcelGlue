namespace ExcelGlue

[<RequireQualifiedAccess>]
module CAST =
    open System
    open ExcelDna.Integration

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
        /// Indicates which xl-values are to be converted to Double.NaN :
        ///    - if OnlyErrorNA, only ExcelErrorNA values are converted to Double.NaN.
        ///    - if AllErrors, all Excel error values are converted to Double.NaN.
        ///    - if AllNonNumeric, any non-numeric xl-value are converted to Double.NaN.
        type NanConversion = | OnlyErrorNA | AllErrors | AllNonNumeric | AllNonNumericAndNA | AllNonNumericAndErrors | NoConversionToNaN

        /// Converts an xl-value to boxed Double.NaN in some cases.
        let nanify (toNaN: NanConversion) (xlVal: obj) : obj = 
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
        let fail (toNaN: NanConversion) (msg: string option) (xlVal: obj) = 
            nanify toNaN xlVal |> Cast.fail<double> msg

        /// Casts an xl-value to double with a default value, with some values potentially cast to Double.NaN.
        let def (toNaN: NanConversion) (defValue: double) (xlVal: obj) = 
            nanify toNaN xlVal |> Cast.def<double> defValue

        /// Casts an xl-value to a double option type with a default value, with some values potentially cast to Double.NaN.
        let tryDV (toNaN: NanConversion) (defValue: double option) (xlVal: obj) = 
            nanify toNaN xlVal |> Cast.tryDV<double> defValue

        /// Replaces an xl-value with a double default value if it isn't a boxed double type (e.g. box true), with some values potentially cast to Double.NaN.
        let defO (toNaN: NanConversion) (defValue: double) (xlVal: obj) = 
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
        /// Replaces an xl-value with a typed default value.
        let def<'a> (defValue: 'a) (xlVal: obj) : 'a =
            match xlVal with
            | :? ExcelMissing -> defValue
            | _ -> Cast.def<'a> defValue xlVal

        /// Replaces an xl-value with an obj default value if missing.
        let defO (defValue: obj) (o: obj) : obj =
            match o with
            | :? ExcelMissing -> defValue
            | _ -> o

        /// Replaces an xl-value with None if missing.
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
        let of2D2 (rowWiseDef: bool) (o2D: obj[,]) : obj[] = 
            // column-wise slice
            if Array2D.length2 o2D = 1 then
                o2D.[*, Array2D.base2 o2D]
            // row-wise slice
            elif rowWiseDef || (Array2D.length1 o2D = 1) then
                o2D.[Array2D.base1 o2D, *]
            // column-wise slice as default
            else 
                o2D.[*, Array2D.base2 o2D]



        let _end = "here"













