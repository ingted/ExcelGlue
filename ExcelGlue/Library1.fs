namespace ExcelGlue

[<RequireQualifiedAccess>]
module CAST =
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
        /// Casts an xl value to bool or fails.
        let fail (msg: string option) (xlVal: obj) = Cast.fail<bool> msg xlVal

        /// Casts an xl value to bool with default value.
        let def (defValue: bool) (xlVal: obj) = Cast.def<bool> defValue xlVal

        /// Casts an xl value to a bool option type with default value.
        let tryDV (defValue: bool option) (xlVal: obj) = Cast.tryDV<bool> defValue xlVal

        /// Replaces an xl value with bool default value if it isn't a boxed bool (e.g. box true).
        let defO (defValue: bool) (xlVal: obj) = Cast.def<bool> defValue xlVal

    [<RequireQualifiedAccess>]
    module Stg =
        /// Casts an xl value to string or fails.
        let fail (msg: string option) (xlVal: obj) = Cast.fail<string> msg xlVal

        /// Casts an xl value to string with default value.
        let def (defValue: string) (xlVal: obj) = Cast.def<string> defValue xlVal

        /// Casts an xl value to a string option type with default value.
        let tryDV (defValue: string option) (xlVal: obj) = Cast.tryDV<string> defValue xlVal

        /// Replaces an xl value with string default value if it isn't a boxed string (e.g. box true).
        let defO (defValue: string) (xlVal: obj) = Cast.def<string> defValue xlVal

    // TODO add double, dates, int, doubleNA?


    [<RequireQualifiedAccess>]
    module Missing =

        /// Replaces an xl value with typed default value.
        let def<'a> (defValue: 'a) (xlVal: obj) : 'a =
            match xlVal with
            | :? ExcelMissing -> defValue
            | _ -> Cast.def<'a> defValue xlVal

        /// Applies a map to an xl value, and replaces missing values with default.
        let map<'a> (defValue: 'a) (mapping: obj -> 'a) (xlVal: obj) : 'a =
            match xlVal with
            | :? ExcelMissing -> defValue
            | _ -> mapping xlVal

        /// Applies a map to an xl value, but returns None for missing values.
        let tryMap<'a> (mapping: obj -> 'a option) (xlVal: obj) : 'a option =
            match xlVal with
            | :? ExcelMissing -> None
            | _ -> mapping xlVal

        /// Replaces an xl value with default value if missing.
        let defO (defValue: obj) (o: obj) : obj =
            match o with
            | :? ExcelMissing -> defValue
            | _ -> o

        /// Replaces an xl value with None if missing.
        let tryO (o: obj) : obj option =
            match o with
            | :? ExcelMissing -> None
            | _ -> Some o


        
        let _end = "here"




















