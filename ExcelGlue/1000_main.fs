//  Copyright (c) cdr021. All rights reserved.
//  ExcelGlue is licensed under the MIT license. See LICENSE.txt for details.

namespace ExcelGlue

open System
open System.IO
open System.Collections.Generic
open ExcelDna.Integration
open System.Runtime.Serialization.Formatters.Binary

/// Class where all "registered" xl-sheet objects are stored. TODO better wording
type Registry() =
    // 2 dictionaries to keep track of objects, Registry objects or R-objects, and corresponding Excel-cells references, where the objects are generated.
    /// "The" Registry: 
    /// Dictionary of string key (mostly guids) -> R-object.
    let reg = Dictionary<string, obj>()
    /// The Excel-cell references dictionary: 
    /// Dictionary of Excel-cell reference -> list of associated R-objects keys.
    let ref = Dictionary<string, string[]>()

    // -----------------------------
    // -- Construction functions
    // -----------------------------

    /// Removes all R-objects, filed under the Excel-cell reference refKey, from the Registry,
    /// then removes the reference from the Excel-cell references dictionary.
    member this.removeReferencedObjects (refKey: string) : unit = 
        if ref.ContainsKey refKey then
            for regKey in ref.Item(refKey) do reg.Remove(regKey) |> ignore
            ref.Remove(refKey) |> ignore

    /// Removes a single R-objects, filed under the Excel-cell reference refKey, from the Registry,
    /// If this was the only R-object filed undert this reference, then removes the reference from the dictionary.
    member this.removeSingleReferencedObjects (refKey: string) (regKey: string) : unit = 
        if ref.ContainsKey refKey then
            let regKeys = ref.Item(refKey)
            if regKeys |> Array.contains regKey then    
                let regKeys' = regKeys |> Array.filter ((<>) regKey)
                reg.Remove(regKey) |> ignore
                ref.Remove(refKey) |> ignore
                if regKeys'.Length > 0 then
                    ref.Add(refKey, regKeys')

    /// Removes all objects and their xl-cell references from the Object Registry.
    member this.clear = 
        reg.Clear()
        ref.Clear()

    /// Adds a (xl-cell reference -> single Registry-guid) dictionary entry.
    member private this.addReference (refKey: string) (regKey: string) = 
        this.removeReferencedObjects refKey
        ref.Add(refKey, [| regKey |])

    /// Adds a a single Registry key to a (possibly already existing) Excel-cell reference.
    member private this.appendRef (refKey: string) (regKey: string) =
        if ref.ContainsKey refKey then
            let regKeys = ref.Item(refKey)
            ref.Remove(refKey) |> ignore
            ref.Add(refKey, Array.append regKeys [| regKey |])
        else
            ref.Add(refKey, [| regKey |])

    /// Adds a R-object to the Registry given an Excel-cell reference, 
    /// and removes from the Registry all existing R-objects filed under the same reference.
    member this.addObj (refKey: string) (regObject: obj) (regKey: string) : string = 
        this.addReference refKey regKey
        reg.Add(regKey, regObject)
        regKey

    /// Creates a unique Registry key,
    /// Adds a R-object to the Registry given a Excel-cell reference, 
    /// and removes from the Registry all existing R-objects filed under the same reference.
    member this.register (refKey: string) (regObject: obj) : string = 
        let regKey = (Guid.NewGuid()).ToString()
        regKey |> this.addObj refKey regObject
        //this.addReference refKey regKey
        //reg.Add(regKey, regObject)
        //regKey

    member this.registerBxd (refKey: string) (regObject: obj) : obj = this.register refKey regObject |> box

    /// Adds a R-object to the Registry given a Excel-cell reference, 
    /// without removing existing R-objects filed under the same reference.
    member this.appendObj (refKey: string) (regObject: obj) (regKey: string) : string = 
        this.appendRef refKey regKey
        reg.Add(regKey, regObject)
        regKey

    /// Creates a unique Registry key,
    /// Adds a R-object to the Registry given a Excel-cell reference, 
    /// without removing existing R-objects filed under the same reference.
    member this.append (refKey: string) (regObject: obj) : string = 
        let regKey = (Guid.NewGuid()).ToString()
        this.appendObj refKey regObject regKey

    /// Given an arbitrary Registry key,
    /// Adds a R-object to the Registry given a xl-cell reference, 
    /// and removes from the Registry all existing R-objects filed under the same reference.
    member this.adhoc (append: bool) (refKey: string) (regKey: string) (regObject: obj) : string =
        if append then
            this.appendObj refKey regObject regKey
        else
            this.addObj refKey regObject regKey

    // -----------------------------
    // -- Inspection functions
    // -----------------------------

    /// Returns the number of R-objects contained in the Registry.
    member this.count = reg.Count

    /// Returns a R-object, given its associated Registry-guid.
    member this.tryGet (regKey: string) : obj option =
        if reg.ContainsKey regKey then
            reg.Item(regKey) |> Some
        else
            None

    /// Returns the (unique by construction) reference (key) associated to a given Registry key.
    member this.findRef (regKey: string) : string option =
        let refKeys = 
            [| for kvp in ref -> kvp |] 
            |> Array.filter (fun kvp -> kvp.Value |> Array.contains regKey)
        refKeys |> Array.tryHead |> Option.map (fun kvp -> kvp.Key)

    /// Returns a R-object's type, given its associated Registry-guid.
    member this.tryType (regKey: string) : Type option =
        regKey |> this.tryGet |> Option.map (fun o -> o.GetType())

    /// Returns a corresponding xl-reference, given a Registry-guid.
    member private this.tryFindRef (regKey: string) : string option = 
        if reg.ContainsKey regKey then
            [| for kvp in ref -> if kvp.Value |> Array.contains regKey then [| kvp.Key |] else [||] |]
            |> Array.concat
            |> Array.head // only reference for a given R-object, by construction.
            |> Some
        else
            None

    /// Given a Registry-guid, if its associated R-object is a 1D array,
    /// returns the array element type and the array.
    member this.tryFind1D (regKey: string) : ((Type[])*obj) option =
        match this.tryGet regKey with
        | None -> None
        | Some regObj ->
            let ty = regObj.GetType()

            if ty.IsArray && (ty.GetArrayRank() = 1) then
                let genty = ty.GetElementType()
                ([| genty |], regObj) 
                |> Some
            else
                None

    /// Given a Registry-guid, if its associated R-object is a 2D array,
    /// returns the array element type and the array.
    member this.tryFind2D (regKey: string) : ((Type[])*obj) option =
        match this.tryGet regKey with
        | None -> None
        | Some regObj ->
            let ty = regObj.GetType()

            if ty.IsArray && (ty.GetArrayRank() = 2) then
                let genty = ty.GetElementType()
                ([| genty |], regObj) 
                |> Some
            else
                None

    /// Checks if 2 R-objects are equal.
    member this.equal (regKey1: string) (regKey2: string) : bool = 
        match this.tryGet regKey1, this.tryGet regKey2 with
        | Some o1, Some o2 -> o1 = o2
        | _ -> false

    /// Returns the Registry's keys.
    member this.keys : string[] = [| for kvp in reg -> kvp.Key |]

    /// Returns the Registry's values.
    member this.values : obj[] = [| for kvp in reg -> kvp.Value |]

    /// Returns the Registry's key-value pairs.
    member this.keyValuePairs : (string*obj)[] = [| for kvp in reg -> kvp.Key, kvp.Value |]

    // -----------------------------
    // -- Extraction functions
    // -----------------------------

    member this.tryExtract<'a> (xlValue: obj) : 'a option =
        match xlValue with
        | :? string as regKey ->
            match this.tryGet regKey with
            | None ->
                if typeof<'a> = typeof<string> then Some (unbox xlValue) else None
            | Some regObj -> 
                match regObj with
                | :? 'a as v -> Some v
                | _ -> None
        | :? 'a as v -> Some v
        | _ -> None

    member this.tryExtractO (xlValue: obj) : obj option =
        match xlValue with
        | :? string as regKey -> this.tryGet regKey
        | _ -> None

    // https://docs.microsoft.com/en-us/dotnet/framework/reflection-and-codedom/how-to-examine-and-instantiate-generic-types-with-reflection
    // https://docs.microsoft.com/en-us/dotnet/api/system.type.getgenerictypedefinition?view=net-5.0
    /// targetGenType is the expected generic type, e.g. targetGenType: typeof<GenMTRX<_>>
    member this.tryExtractGen' (targetGenType: Type) (xlValue: string) : obj option =  // TODO change xlValue into regKey
        if not targetGenType.IsGenericType then
            None
        else
            match this.tryGet xlValue with
            | None -> None
            | Some regObj -> 
                let ty = regObj.GetType()
                let gentydef = ty.GetGenericTypeDefinition()
                let tgttydef = targetGenType.GetGenericTypeDefinition()

                if gentydef = tgttydef then
                    Some regObj
                else
                    None

    /// Same as tryExtractGen', but also return the generic types array.
    /// targetGenType is the expected generic type, e.g. targetGenType: typeof<GenMTRX<_>>
    member this.tryExtractGen (targetType: Type) (regKey: string) : ((Type[])*obj) option =
        this.tryExtractGen' targetType regKey
        |> Option.map (fun o -> ((o.GetType().GetGenericArguments()), o))

    // wording
    member this.tryExtractGen1D (targetType: Type) (xlValue: obj) : ((Type[])*(obj[])) option =
        match xlValue with
        | :? (obj[,]) as o2D ->
            let type_robj_pairs = 
                [| for i in o2D.GetLowerBound(0) .. o2D.GetUpperBound(0) do 
                    yield! 
                        [| for j in o2D.GetLowerBound(1) .. o2D.GetUpperBound(1) -> 
                            match o2D.[i, j] with
                            | :? String as regKey -> this.tryExtractGen targetType regKey
                            | _ -> None
                        |] 
                |]
                |> Array.choose id
            if type_robj_pairs.Length = 0 then
                None
            else
                if Array.TrueForAll(type_robj_pairs, fun (tys, o) -> tys = fst type_robj_pairs.[0]) then
                    (fst type_robj_pairs.[0], type_robj_pairs |> Array.map snd)
                    |> Some
                else
                    None
        | :? string as regKey -> 
            this.tryExtractGen targetType regKey
            |> Option.map (fun (tys, o) -> (tys, [| o |]))
        | _ -> None

    // Returns the type of the first R-Obj found in an Excel range, xlValue.
    member this.trySampleType (strict: bool) (xlValue: obj) : Type option =
        match xlValue with
        | :? (obj[,]) as o2D ->
            let rgtypes = 
                [| for i in o2D.GetLowerBound(0) .. o2D.GetUpperBound(0) do 
                    yield! 
                        [| for j in o2D.GetLowerBound(1) .. o2D.GetUpperBound(1) -> 
                            match o2D.[i, j] with
                            | :? String as regKey -> this.tryType regKey 
                            | _ -> None
                        |] 
                |]
                |> Array.distinct
                |> Array.choose id

            if rgtypes.Length = 0 then
                None
            elif strict && rgtypes.Length > 1 then
                None
            else
                rgtypes |> Array.head |> Some

        | :? (obj[]) as o1D ->
            let rgtypes = 
                o1D 
                |> Array.map (fun o -> match o with | :? String as regKey -> this.tryType regKey | _ -> None)
                |> Array.distinct
                |> Array.choose id

            if rgtypes.Length = 0 then
                None
            elif strict && rgtypes.Length > 1 then
                None
            else
                rgtypes |> Array.head |> Some

        | :? string as regKey -> this.tryType regKey
        | _ -> None

    // -----------------------------
    // -- Excel RefID functions
    // -----------------------------

    member this.excelRef (caller: obj) : string = 
        match caller with
        | :? ExcelReference as ref -> ref.ToString()
        | _ -> ""

    /// Generates a unique key based on the caller xl-cell.
    member this.refID = XlCall.Excel XlCall.xlfCaller |> this.excelRef

    // -----------------------------
    // -- Pretty-print functions
    // -----------------------------

    /// Pretty-prints a R-object, given its associated Registry-guid.
    /// Using F# default formatting.
    member this.tryShow (key: string) : string option =
        this.tryGet key |> Option.map (fun o -> sprintf "%A" o)

    /// Pretty-prints a R-object, given its associated Registry-guid.
    /// Using .Net default formatting.
    member this.tryString (key: string) : string option =
        this.tryGet key |> Option.map (fun o -> o.ToString())

    // -----------------------------
    // -- Miscellaneous
    // -----------------------------
    
    /// Returns a default value. TODO wording 
    /// Unsafe!
    member this.defaultValue<'a> (xlValue: obj) : 'a =
        this.tryExtract<'a> xlValue |> Option.defaultValue (Unchecked.defaultof<'a>)

    /// Saves a R-object to disk.
    member this.ioWriteBin (fpath: string) (regKey: string) : DateTime option =
        match regKey |> this.tryGet with
        | None -> None
        | Some o ->
            use stream = new FileStream(fpath, FileMode.Create)
            (new BinaryFormatter()).Serialize(stream, o)
            DateTime.Now |> Some

    /// Loads a R-object from disk.
    member this.ioLoadBin (fpath: string) (refKey: string) : string =
        use stream = new FileStream(fpath, FileMode.Open)
        let regObj = (new BinaryFormatter()).Deserialize(stream)
        this.register refKey regObj

    /// Renames an existing Registry key to an arbitrary (string) value.
    member this.rename (newRefKey: string) (prevRegKey: string) (newRegKey: string) : string option =
        match this.findRef prevRegKey with
        | None -> None
        | Some prevRefKey ->
            let regObject = reg.Item(prevRegKey)
            this.removeSingleReferencedObjects prevRefKey prevRegKey
            this.adhoc false newRefKey newRegKey regObject
            |> Some

module Registry =
    /// Master registry where all registered objects are held.
    let MRegistry = Registry()

module API = 

    /// F# types for the xl-spreadsheet values we want to capture.
    // Could be extended to more numeric types? E.g. Int64, Decimals...?
    type Variant = | BOOL | BOOLOPT | STRING | STRINGOPT | DOUBLE | DOUBLEOPT | DOUBLENAN | DOUBLENANOPT | INT | INTOPT | DATE | DATEOPT | VAR | VAROPT | OBJ with
        static member isOptionalType (typeTag: string) : bool = 
            typeTag.IndexOf("#") >= 0

        static member ofTag (typeTag: string) : Variant = 
            let isoption = Variant.isOptionalType typeTag
            let prepString = typeTag.Replace(" ", "").Replace(":", "").Replace("#", "").ToUpper()
            match prepString with
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

        static member labelType (removeOptionMark: bool) (typeTag: string) : Type = 
            let var = (if removeOptionMark then typeTag.Replace("#", "") else typeTag) |> Variant.ofTag
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

        static member labelDefVal (typeTag: string) : obj = 
            let var = Variant.ofTag typeTag
            var.defVal
        
        /// Re-types Excel values depending on a type-tag.
        static member rebox (typeTag: string) (xlvalue: obj) : obj = 
            let var = Variant.ofTag typeTag
            match var, xlvalue with
            | DOUBLE, (:? ExcelError as e) when e = ExcelError.ExcelErrorNA -> box Double.NaN
            | DOUBLENAN, (:? ExcelError as e) when e = ExcelError.ExcelErrorNA -> box Double.NaN
            | INT, (:? double as d) -> d |> int |> box
            | DATE, (:? double as d) -> DateTime.FromOADate(d) |> box

            | DOUBLEOPT, (:? ExcelError as e) when e = ExcelError.ExcelErrorNA -> Double.NaN |> Some |> box
            | DOUBLENANOPT, (:? ExcelError as e) when e = ExcelError.ExcelErrorNA -> Double.NaN |> Some |> box
            | INTOPT, (:? double as d) -> d |> int |> Some |> box
            | DATE, (:? double as d) -> DateTime.FromOADate(d) |> Some |> box

            | BOOLOPT, (:? bool as b) -> b |> Some |> box
            | STRINGOPT, (:? string as s) -> s |> Some |> box
            | VAROPT, v -> v |> Some |> box

            | _ -> xlvalue

        /// Convenience function.
        member this.empty1D : obj = 
            match this with
            | BOOL -> ([||] : bool[]) |> box
            | BOOLOPT -> ([||] : bool option[]) |> box
            | STRING -> ([||] : string[]) |> box
            | STRINGOPT -> ([||] : string option[]) |> box
            | DOUBLE -> ([||] : double[]) |> box
            | DOUBLEOPT -> ([||] : double option[]) |> box
            | DOUBLENAN -> ([||] : double[]) |> box
            | DOUBLENANOPT -> ([||] : double option[]) |> box
            | INT -> ([||] : int[]) |> box
            | INTOPT -> ([||] : int option[]) |> box
            | DATE -> ([||] : DateTime[]) |> box
            | DATEOPT -> ([||] : DateTime option[]) |> box
            | VAR -> ([||] : obj[]) |> box
            | VAROPT -> ([||] : obj option[]) |> box
            | OBJ-> ([||] : obj[]) |> box

        // TODO: wording
        static member labelEmpty1D (typeTag: string) : obj = 
            let var = Variant.ofTag typeTag
            var.empty1D

    // FIXME better wording
    /// Describes various convenient sets, "kinds", of xl-spreadsheet values.
    type Kind = | Boolean | Numeric | Textual | NA | Error | Missing | Empty with
        static member def : Set<Kind> = NA |> Set.singleton

        static member hasKind (kinds: Set<Kind>) (xlValue: obj) : bool =
            match xlValue with 
            | :? bool -> kinds |> Set.contains Boolean
            | :? double -> kinds |> Set.contains Numeric
            | :? string -> kinds |> Set.contains Textual
            | :? ExcelError as e when e = ExcelError.ExcelErrorNA -> kinds |> Set.contains NA
            | :? ExcelError -> kinds |> Set.contains Error
            | :? ExcelMissing -> kinds |> Set.contains Missing
            | :? ExcelEmpty -> kinds |> Set.contains Empty
            | _ -> false

        static member ofLbl (singleLabel: string) : Kind[] = 
            match singleLabel.ToUpper() with
            | "B" | "BOOL" | "BOOLEAN" -> [| Boolean |]
            | "S" | "STG" | "STRING" | "TXT" | "TEXT" -> [| Textual |]
            | "D" | "DBL" | "DOUBLE" | "NUM" | "NUMERIC" -> [| Numeric |]
            | "N" | "NA" | "NAN" | "#NA" | "#N/A" -> [| NA |]
            | "E" | "ERR" | "ERROR" -> [| Error |]
            | "M" | "MISS" | "MISSING" -> [| Missing |]
            | "EMPTY" -> [| Empty |]
            | "A" | "ABS" | "ABSENT" -> [| Missing; Empty |]
            | _ -> [||]

        static member all = [| Boolean;  Numeric;  Textual; NA;  Error;  Missing;  Empty |] |> Set.ofArray

        /// Translates a comma separated string into a set of kinds.
        /// '!' as first element takes the complement set.
        ///    - "!NUM" // non-numeric kinds, matches any non-numeric value.
        ///    - "NA" // Nan kind, matches #N/A.
        ///    - "ERR" // Error kinds, matches any Excel error.
        ///    - "ABS" // Absent set, matches Missing and Empty arguments.
        ///    - "TXT" // Textual set, matches any string argument
        ///    - "!TXT,BOOL" // set of all non-string, non-boolean elements.
        ///    ...
        static member ofLabel (label: string) : Set<Kind> =
            let neg = label.StartsWith("!")
            let kinds = label.ToUpper().Replace("!","").Split([| "," |], StringSplitOptions.None) |> Array.collect Kind.ofLbl |> Set.ofArray
            if neg then Set.difference Kind.all kinds else kinds

    /// Functions to handle or process Inputs from Excel.
    module In =
        /// Excel native types conversion (e.g. from obj[,] to obj[]...).
        module Cast =
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
                // column-wise slice as default for fat 2D arrays
                else 
                    o2D.[*, Array2D.base2 o2D]
        
            /// Converts an xl-value to a 2D array.
            /// (Use to2D rather than try2D when the obj argument is an xl-value).
            let to2D (xlVal: obj) : obj[,] =
                match xlVal with
                | :? (obj[,]) as o2D -> o2D
                | :? (obj[]) as o1D -> [| o1D |] |> array2D // FIXME - transpose // THIS CASE SHD NOT OCCUR if xlVal is an excel arg
                | o0D -> Array2D.create 1 1 o0D
        
            /// Converts an obj to a 2D array option.
            /// (Use try2D rather than to2D when the obj argument is not an xl-value).
            let try2D (o: obj) : obj[,] option =
                match o with
                | :? (obj[,]) as o2D -> o2D |> Some
                | :? (obj[]) as o1D -> [| o1D |] |> array2D |> Some // TODO - transpose
                | _ -> None

            /// Converts an xl-value to a 1D array.
            /// (Use to1D rather than try1D when the obj argument is an xl-value).
            let to1D (rowWiseDef: bool) (xlVal: obj) : obj[] =
                match xlVal with
                | :? (obj[,]) as o2D -> slice2D rowWiseDef o2D
                | :? (obj[]) as o1D -> o1D
                //| :? ExcelMissing -> [||] // TODO: add missing and empty cases (empty result)?
                //| :? ExcelEmpty -> [||]
                | o0D -> [| o0D |] 
        
            /// Converts an obj to a 1D array option.
            /// (Use try1D rather than to1D when the obj argument is not an xl-value).
            let try1D (rowWiseDef: bool) (o: obj) : obj[] option =
                match o with
                | :? (obj[,]) as o2D -> slice2D rowWiseDef o2D |> Some
                | :? (obj[]) as o1D -> Some o1D
                | _ -> None
        
        /// Obj input functions.
        module D0 =

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

            // Optional-type default FIXX FIXME wording
            module Opt =
                /// Casts an obj to an option on generic type with typed default-value.
                let def<'a> (defValue: 'a option) (o: obj) : 'a option =
                    match o with
                    | :? 'a as v -> Some v
                    | _ -> defValue

            /// Object substitution, based on type.
            module Obj =
                /// Substitutes an object for another one, if it isn't of the specified generic type 'a.
                let subst<'a> (defValue: obj) (o: obj) : obj =
                    match o with
                    | :? 'a -> o
                    | _ -> defValue                

            [<RequireQualifiedAccess>]
            module Bool =
                /// Casts an xl-value to bool or fails.
                let fail (msg: string option) (xlVal: obj) = fail<bool> msg xlVal

                /// Casts an xl-value to bool with a default-value.
                let def (defValue: bool) (xlVal: obj) = def<bool> defValue xlVal

                // optional-type default FIXX
                module Opt =
                    /// Casts an xl-value to a bool option type with a default-value.
                    let def (defValue: bool option) (xlVal: obj) = Opt.def<bool> defValue xlVal

                /// Object substitution, based on type.
                module Obj =
                    /// Substitutes an object for another one, if it isn't a (boxed) bool (e.g. box false).
                    let subst (defValue: obj) (xlVal: obj) : obj = Obj.subst<bool> defValue xlVal

            [<RequireQualifiedAccess>]
            module Stg =
                /// Casts an xl-value to string or fails.
                let fail (msg: string option) (xlVal: obj) = fail<string> msg xlVal

                /// Casts an xl-value to string with a default-value.
                let def (defValue: string) (xlVal: obj) = def<string> defValue xlVal

                // optional-type default FIXX
                module Opt =
                    /// Casts an xl-value to a string option type with a default-value.
                    let def (defValue: string option) (xlVal: obj) = Opt.def<string> defValue xlVal

                /// Object substitution, based on type.
                module Obj =
                    /// Substitutes an object for another one, if it isn't a (boxed) string (e.g. box "foo").
                    let subst (defValue: obj) (xlVal: obj) = Obj.subst<string> defValue xlVal

            [<RequireQualifiedAccess>]
            module Dbl =
                /// Casts an xl-value to double or fails.
                let fail (msg: string option) (xlVal: obj) = fail<double> msg xlVal

                /// Casts an xl-value to double with a default-value.
                let def (defValue: double) (xlVal: obj) = def<double> defValue xlVal

                // optional-type default FIXX
                module Opt =
                    /// Casts an xl-value to a double option type with a default-value.
                    let def (defValue: double option) (xlVal: obj) = Opt.def<double> defValue xlVal

                /// Object substitution, based on type.
                module Obj =
                    /// Substitutes an object for another one, if it isn't a (boxed) double (e.g. box 1.0).
                    let subst (defValue: obj) (xlVal: obj) = Obj.subst<double> defValue xlVal

            [<RequireQualifiedAccess>]
            module Nan =
                /// Converts a boxed ExcelErrorNA into a Double.NaN.
                let ofNA (xlVal: obj) : obj =
                    match xlVal with
                    | :? ExcelError as err when err = ExcelError.ExcelErrorNA -> Double.NaN |> box
                    | _ -> xlVal

                /// Converts xl-values of the provided kinds to to boxed Double.NaN.
                let nanify (xlKinds: Set<Kind>) (xlVal: obj) : obj = 
                    if xlVal |> Kind.hasKind xlKinds then
                        box Double.NaN
                    else
                        xlVal

                /// Converts a boxed Double.NaN into an ExcelErrorNA. // FIXME - should be OUT?
                let ofNaNTBD (xlVal: obj) : obj =
                    match xlVal with
                    | :? double as d -> if Double.IsNaN d then ExcelError.ExcelErrorNA |> box else box d
                    | _ -> xlVal

                /// Casts an xl-value to double or fails, with some other non-double values potentially cast to Double.NaN.
                let fail (xlKinds: Set<Kind>) (msg: string option) (xlVal: obj) = 
                    nanify xlKinds xlVal |> fail<double> msg

                /// Casts an xl-value to double with a default-value, with some other non-double values potentially cast to Double.NaN. // FIXME - improve text
                let def (xlKinds: Set<Kind>) (defValue: double) (xlVal: obj) = 
                    nanify xlKinds xlVal |> def<double> defValue

                // optional-type default TODO wording
                module Opt =
                    /// Casts an xl-value to a double option type with a default-value, with some other non-double values potentially cast to Double.NaN.
                    let def (xlKinds: Set<Kind>) (defValue: double option) (xlVal: obj) = 
                        nanify xlKinds xlVal |> Opt.def<double> defValue

                /// Object substitution, based on type.
                module Obj =
                    /// Substitutes an object for another one, if it isn't a (boxed) double (e.g. box 1.0).
                    /// Replaces an xl-value with a double default-value if it isn't a (boxed double) type (e.g. box 1.0), with some other non-double values potentially cast to Double.NaN.
                    let subst (xlKinds: Set<Kind>) (defValue: obj) (xlVal: obj) = 
                        nanify xlKinds xlVal |> Obj.subst<double> defValue

                    /// Converts a boxed ExcelErrorNA into a Double.NaN.
                    /// Similar to nanify function, but lighter.
                    let ofNA (xlVal: obj) : obj =
                        match xlVal with
                        | :? ExcelError as err when err = ExcelError.ExcelErrorNA -> Double.NaN |> box
                        | _ -> xlVal

                    /// Converts a boxed Double.NaN into an ExcelErrorNA. // FIXME - should be OUT?
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

                // optional-type default FIXX
                module Opt =
                    /// Casts an xl-value to a int option type with a default-value.
                    let def (defValue: int option) (xlVal: obj) =
                        match xlVal with
                        | :? double as d -> match ofDouble d with | None -> defValue | Some i -> Some i
                        | :? int as i -> Some i
                        | _ -> defValue

                /// Object substitution, based on type.
                module Obj =
                    /// Substitutes an object for another one, if it isn't a (boxed) int (e.g. box 1).
                    let subst (defValue: obj) (xlVal: obj) : obj =
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

                // optional-type default FIXX
                module Opt =
                    /// Casts an xl-value to a DateTime option type with a default-value.
                    let def (defValue: DateTime option) (xlVal: obj) : DateTime option =
                        match xlVal with
                        | :? double as d -> DateTime.FromOADate d |> Some
                        | :? DateTime as dte -> Some dte
                        | _ -> defValue

                /// Object substitution, based on type.
                module Obj =
                    /// Substitutes an object for another one, if it isn't a (boxed) DateTime (e.g. box 36526.0).
                    let subst (defValue: obj) (xlVal: obj) : obj =
                        match xlVal with
                        | :? double as d -> DateTime.FromOADate d |> box
                        | _ -> defValue

            [<RequireQualifiedAccess>]
            module Missing =
                /// Casts an obj to a generic type given a typed default-value.
                /// Replaces ExcelMissing values with the default-value.
                let def<'a> (defValue: 'a) (xlVal: obj) : 'a =
                    match xlVal with
                    | :? ExcelMissing -> defValue
                    | _ -> def<'a> defValue xlVal

                /// Applies a map to an xl-value, and replaces ExcelMissing values with a typed default-value.
                let map<'a> (defValue: 'a) (mapping: obj -> 'a) (xlVal: obj) : 'a =
                    match xlVal with
                    | :? ExcelMissing -> defValue
                    | _ -> mapping xlVal

                /// Applies a map to an xl-value, but returns None for ExcelMissing values.
                let tryMap<'a> (mapping: obj -> 'a option) (xlVal: obj) : 'a option =
                    match xlVal with
                    | :? ExcelMissing -> None
                    | _ -> mapping xlVal

                /// Missing functions for untyped inputs.
                module Obj =
                    /// Substitutes a default-value for ExcelMissing values.
                    /// Otherwise passes the xl-value through.
                    let subst (defValue: obj) (xlVal: obj) : obj =
                        match xlVal with
                        | :? ExcelMissing -> defValue
                        | _ -> xlVal
                    
                    /// Substitutes None for ExcelMissing values.
                    /// Otherwise passes the xl-value through.
                    let tryO (o: obj) : obj option =
                        match o with
                        | :? ExcelMissing -> None
                        | _ -> Some o

            /// Same as the Missing module, for ExcelMissing and ExcelEmpty xl-values.
            [<RequireQualifiedAccess>]
            module Absent =
                /// Casts an obj to a generic type given a typed default-value.
                /// Replaces ExcelMissing or ExcelEmpty values with the default-value.
                let def<'a> (defValue: 'a) (xlVal: obj) : 'a =
                    match xlVal with
                    | :? ExcelEmpty -> defValue
                    | _ -> Missing.def<'a> defValue xlVal

                /// Applies a map to an xl-value, and replaces ExcelMissing or ExcelEmpty values with a typed default-value.
                let map<'a> (defValue: 'a) (mapping: obj -> 'a) (xlVal: obj) : 'a =
                    match xlVal with
                    | :? ExcelEmpty -> defValue
                    | _ -> Missing.map<'a> defValue mapping xlVal

                /// Applies a map to an xl-value, but returns None for ExcelMissing or ExcelEmpty values.
                let tryMap<'a> (mapping: obj -> 'a option) (xlVal: obj) : 'a option =
                    match xlVal with
                    | :? ExcelEmpty -> None
                    | _ -> Missing.tryMap<'a> mapping xlVal

                /// Missing functions for untyped inputs.
                module Obj =
                    /// Substitutes a default-value for ExcelMissing or ExcelEmpty values.
                    /// Otherwise passes the xl-value through.
                    let subst (defValue: obj) (xlVal: obj) : obj =
                        match xlVal with
                        | :? ExcelEmpty -> defValue
                        | _ -> Missing.Obj.subst defValue xlVal
                    
                    /// Substitutes None for ExcelMissing values.
                    /// Otherwise passes the xl-value through.
                    let tryO (xlVal: obj) : obj option =
                        match xlVal with
                        | :? ExcelEmpty -> None
                        | _ -> Missing.Obj.tryO xlVal

            type TagFn =
                /// Returns a default-value compatible with 'A and the typeTag.
                static member defaultValue<'A> (typeTag: string) (xlValue: obj option) : 'A =
                    let defval = Variant.labelDefVal typeTag

                    match xlValue with
                    | None -> defval :?> 'A
                    | Some xlval -> 
                        let dv =
                            if typeTag.ToUpper() = "INT" then 
                                Intg.def (defval :?> int) xlval |> box
                            elif typeTag.ToUpper() = "DATE" then 
                                Dte.def (defval :?> DateTime) xlval |> box
                            else
                                def<'A> (defval :?> 'A) xlval |> box
                        dv :?> 'A

                /// Returns a default-value or None.
                static member defaultValueOpt<'A> (xlValue: obj option) : 'A option =
                    match xlValue with
                    | None -> None 
                    | Some xlval -> 
                        match xlval with
                            | :? 'A as a -> Some a
                            | :? ('A option) as aopt -> aopt
                            | _ -> None

                /// Casts an xl-value to a 'A, with a default-value for when the casting fails.
                static member def<'A> (xlKinds: Set<Kind>) (defValue: obj option) (typeTag: string) (xlValue: obj) : 'A = 

                    match typeTag |> Variant.ofTag with
                    | BOOL -> 
                        let defval = TagFn.defaultValue<'A> typeTag defValue
                        def<'A> defval xlValue
                    | STRING -> 
                        let defval = TagFn.defaultValue<'A> typeTag defValue
                        def<'A> defval xlValue
                    | DOUBLE -> 
                        let defval = TagFn.defaultValue<'A> typeTag defValue
                        def<'A> defval xlValue
                    | DOUBLENAN -> 
                        let defval = TagFn.defaultValue<double> typeTag defValue
                        //let a0D = Nan.def Kind.nonNumericAndNA defval xlValue // TODO: pass xlkinds as argument
                        let a0D = Nan.def xlKinds defval xlValue // TODO: pass xlkinds as argument
                        box a0D :?> 'A
                    | INT -> 
                        let defval = TagFn.defaultValue<int> typeTag defValue |> int
                        let a0D = Intg.def defval xlValue
                        box a0D :?> 'A
                    | DATE -> 
                        let defval = TagFn.defaultValue<DateTime> typeTag defValue
                        let a0D = Dte.def defval xlValue
                        box a0D :?> 'A
                    | _ -> failwith "TO BE IMPLEMENTED WITH OTHER VARIANT TYPES" // TODO: Complete the list
    
                /// Casts an xl-value to a 'A option, with a default-value for when the casting fails.
                /// defValue is None, Some 'a or even Some (Some 'a).
                static member defOpt<'A> (xlKinds: Set<Kind>) (defValue: obj option) (typeTag: string) (xlValue: obj) : 'A option = 
                    match typeTag |> Variant.ofTag with
                    | BOOLOPT -> 
                        let defval : 'A option = TagFn.defaultValueOpt<'A> defValue
                        Opt.def<'A> defval xlValue
                    | STRINGOPT -> 
                        let defval : 'A option = TagFn.defaultValueOpt<'A> defValue
                        Opt.def<'A> defval xlValue
                    | DOUBLEOPT -> 
                        let defval : 'A option = TagFn.defaultValueOpt<'A> defValue
                        Opt.def<'A> defval xlValue
                    | DOUBLENANOPT -> 
                        let defval = TagFn.defaultValueOpt<double> defValue
                        // let a0D = Nan.Opt.def Kind.nonNumericAndNA defval xlValue // TODO: pass xlkinds as argument
                        let a0D = Nan.Opt.def xlKinds defval xlValue // TODO: pass xlkinds as argument
                        box a0D :?> 'A option
                        // Opt.def<'A> defval xlValue
                    | INTOPT -> 
                        let defval = TagFn.defaultValueOpt<double> defValue |> Option.map (int)
                        let a0D = Intg.Opt.def defval xlValue
                        box a0D :?> 'A option
                    | DATEOPT -> 
                        let defval = TagFn.defaultValueOpt<double> defValue |> Option.map (fun d -> DateTime.FromOADate(d))
                        let a0D = Dte.Opt.def defval xlValue
                        box a0D :?> 'A option
                    | _ -> failwith "TO BE IMPLEMENTED WITH OTHER VARIANT TYPES" // TODO FIXME

            [<RequireQualifiedAccess>]
            module Tag = 
                /// Casts an xl-value to a 'a, with a default-value for when the casting fails.
                /// 'a is determined by typeTag.
                let def (xlKinds: Set<Kind>) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype = typeTag |> Variant.labelType true
                    let args : obj[] = [| xlKinds; defValue; typeTag; xlValue |]
                    let res = Toolbox.Generics.invoke<TagFn> "def" [| gentype |] args
                    res

                // optional-type default FIXX
                module Opt =
                    /// Casts an xl-value to a 'a option, with a default-value for when the casting fails.
                    /// 'a is determined by typeTag.
                    let def (xlKinds: Set<Kind>) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                        let gentype = typeTag |> Variant.labelType true
                        let args : obj[] = [| xlKinds; defValue; typeTag; xlValue |]
                        let res = Toolbox.Generics.invoke<TagFn> "defOpt" [| gentype |] args
                        res

                /// For when the type-tag is either optional, e.g. "#string", or not, e.g. "string". TODO wording
                module Any =
                    /// Convenient, single function covering def and Opt.def cases.
                    /// The returned (boxed) value might be either a 'a or a ('a option), depending on wether the type-tag is optional or not.
                    let def (xlKinds: Set<Kind>) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                        let gentype = typeTag |> Variant.labelType true
                        let args : obj[] = [| xlKinds; defValue; typeTag; xlValue |]

                        let res =
                            if typeTag |> Variant.isOptionalType then
                                Toolbox.Generics.invoke<TagFn> "defOpt" [| gentype |] args
                            else
                                Toolbox.Generics.invoke<TagFn> "def" [| gentype |] args
                        res

            // -------------------------
            // -- Convenience functions
            // -------------------------
                
            /// Xl-values tests.
            module Is =
                let missing (xlValue: obj) : bool = 
                    match xlValue with    
                    | :? ExcelMissing -> true
                    | _ -> false

                let empty (xlValue: obj) : bool = 
                    match xlValue with    
                    | :? ExcelEmpty -> true
                    | _ -> false

                let absent (xlValue: obj) : bool = 
                    match xlValue with    
                    | :? ExcelMissing -> true
                    | :? ExcelEmpty -> true
                    | _ -> false

                let error (xlValue: obj) : bool = 
                    match xlValue with    
                    | :? ExcelError -> true
                    | _ -> false

                let blank (xlValue: obj) : bool = 
                    match xlValue with
                    | :? ExcelEmpty -> true
                    | :? string as s -> s.Trim() = ""
                    | _ -> false 

                let blankOrError (xlarg: obj) : bool = 
                    (blank xlarg) || (error xlarg)

        /// Obj[] input functions.  // wording
        module D1 =
            //open Excel
            open type Variant

            // non-optional-type default
            /// Converts an obj[] to a 'a[], given a typed default-value for elements which can't be cast to 'a.
            let def<'a> (defValue: 'a) (o1D: obj[]) : 'a[] =
                o1D |> Array.map (D0.def<'a> defValue)

            // optional-type default
            module Opt =
                /// Converts an obj[] to a ('a option)[], given an optional default-value for elements which can't be cast to 'a.
                let def<'a> (defValue: 'a option) (o1D: obj[]) : ('a option)[] =
                    o1D |> Array.map (D0.Opt.def<'a> defValue)

            /// Converts an obj[] to a 'a[], removing any element which can't be cast to 'a.
            let filter<'a> (o1D: obj[]) : 'a[] =
                o1D |> Array.choose (D0.Opt.def<'a> None)

            /// Converts an obj[] to an optional 'a[]. All the elements must match the given type, otherwise defValue array is returned. 
            let tryDV<'a> (defValue1D: 'a[] option) (o1D: obj[]) : 'a[] option =
                let convert = Opt.def None o1D
                match convert |> Array.tryFind Option.isNone with
                | None -> convert |> Array.map Option.get |> Some
                | Some _ -> defValue1D

            [<RequireQualifiedAccess>]
            module Bool =
                /// Converts an obj[] to a bool[]
                /// Non-bool elements are replaced by the default-value. TODO / FIXME use the same comments for other types
                let def (defValue: bool) (o1D: obj[]) = def defValue o1D

                /// optional-type default
                module Opt =
                    /// Converts an obj[] to a ('a option)[], given a default-value for non-bool elements.
                    /// bool elements, x, are wrapped to (Some x).
                    /// Non-bool elements are replaced by the default-value.
                    let def (defValue: bool option) (o1D: obj[]) = Opt.def defValue o1D

                /// Converts an obj[] to a bool[], removing any non-bool element.
                let filter (o1D: obj[]) = filter<bool> o1D

                /// Converts an obj[] to an optional bool[]. All the elements must be bool, otherwise defValue array is returned. 
                let tryDV (defValue1D: bool[] option) (o1D: obj[])  = tryDV<bool> defValue1D o1D

                /// Converts an obj[] to a bool[]. All the elements must be bool, fails otherwise. 
                let fail (o1D: obj[]) = match tryDV None o1D with | Some a1D -> a1D | None -> failwith "Cannot cast elements to bool."

            [<RequireQualifiedAccess>]
            module Stg =
                /// Converts an obj[] to a string[], given a default-value for non-string elements.
                let def (defValue: string) (o1D: obj[]) = def<string> defValue o1D  // TODO add <string> everywhere!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

                /// optional-type default
                module Opt = 
                    /// Converts an obj[] to a ('a option)[], given a default-value for non-string elements.
                    let def (defValue: string option) (o1D: obj[]) = Opt.def<string> defValue o1D

                /// Converts an obj[] to a string[], removing any non-string element.
                let filter (o1D: obj[]) = filter<string> o1D

                /// Converts an obj[] to an optional string[]. All the elements must be string, otherwise defValue array is returned. 
                let tryDV (defValue1D: string[] option) (o1D: obj[])  = tryDV<string> defValue1D o1D

                /// Converts an obj[] to a string[]. All the elements must be string, fails otherwise. 
                let fail (o1D: obj[]) = match tryDV None o1D with | Some a1D -> a1D | None -> failwith "Cannot cast elements to string."

            /// Similar to Stg, but with an xl-value as primary input.
            /// (Choice was made to make rowWiseDef false. A much more frequent choice.)
            [<RequireQualifiedAccess>]
            module OStg =
                /// Converts an obj[] to a string[], given a default-value for non-string elements.
                let def (defValue: string) (xlValue: obj) = Cast.to1D false xlValue |> Stg.def defValue 

                /// optional-type default
                module Opt = 
                    /// Converts an obj[] to a ('a option)[], given a default-value for non-string elements.
                    let def (defValue: string option) (xlValue: obj) = Cast.to1D false xlValue |> Stg.Opt.def defValue 

                /// Converts an obj[] to a string[], removing any non-string element.
                let filter (xlValue: obj) = Cast.to1D false xlValue |> Stg.filter

                /// Converts an obj[] to an optional 'a[]. All the elements must be string, otherwise defValue array is returned. 
                let tryDV (defValue1D: string[] option) (xlValue: obj)  = Cast.to1D false xlValue |> Stg.tryDV defValue1D

                /// Converts an obj[] to an optional string[]. All the elements must be string, otherwise defValue array is returned. 
                let fail (xlValue: obj)  =  match tryDV None xlValue with | Some a1D -> a1D | None -> failwith "Cannot cast elements to string."

            [<RequireQualifiedAccess>]
            module Dbl =
                /// Converts an obj[] to a double[], given a default-value for non-double elements.
                let def (defValue: double) (o1D: obj[]) = def defValue o1D

                /// optional-type default
                module Opt = 
                    /// Converts an obj[] to a ('a option)[], given a default-value for non-double elements.
                    let def (defValue: double option) (o1D: obj[]) = Opt.def defValue o1D

                /// Converts an obj[] to a double[], removing any non-double element.
                let filter (o1D: obj[]) = filter<double> o1D

                /// Converts an obj[] to an optional 'a[]. All the elements must be double, otherwise defValue array is returned. 
                let tryDV (defValue1D: double[] option) (o1D: obj[]) = tryDV<double> defValue1D o1D

                /// Converts an obj[] to a double[]. All the elements must be double, fails otherwise. 
                let fail (o1D: obj[]) = match tryDV None o1D with | Some a1D -> a1D | None -> failwith "Cannot cast elements to double."

            /// Similar to Dbl, but with an xl-value as primary input.
            /// (Choice was made to make rowWiseDef false. A much more frequent choice.)
            [<RequireQualifiedAccess>]
            module ODbl =
                /// Converts an obj[] to a double[], given a default-value for non-double elements.
                let def (defValue: double) (xlValue: obj) = Cast.to1D false xlValue |> Dbl.def defValue 

                /// optional-type default
                module Opt = 
                    /// Converts an obj[] to a ('a option)[], given a default-value for non-double elements.
                    let def (defValue: double option) (xlValue: obj) = Cast.to1D false xlValue |> Dbl.Opt.def defValue 

                /// Converts an obj[] to a double[], removing any non-double element.
                let filter (xlValue: obj) = Cast.to1D false xlValue |> Dbl.filter

                /// Converts an obj[] to an optional 'a[]. All the elements must be double, otherwise defValue array is returned. 
                let tryDV (defValue1D: double[] option) (xlValue: obj) = Cast.to1D false xlValue |> Dbl.tryDV defValue1D

            [<RequireQualifiedAccess>]
            module Nan =
                /// Converts an obj[] to a double[], given a default-value for non-double elements.
                let def (xlKinds: Set<Kind>) (defValue: double) (o1D: obj[]) =
                    o1D |> Array.map (D0.Nan.def xlKinds defValue)

                /// optional-type default
                module Opt = 
                    /// Converts an obj[] to a ('a option)[], given a default-value for non-double elements.
                    let def (xlKinds: Set<Kind>) (defValue: double option) (o1D: obj[]) =
                        o1D |> Array.map (D0.Nan.Opt.def xlKinds defValue)

                /// Converts an obj[] to a DateTime[], removing any non-double element.
                let filter (xlKinds: Set<Kind>) (o1D: obj[]) =
                    o1D |> Array.choose (D0.Nan.Opt.def xlKinds None)

                /// Converts an obj[] to an optional 'a[]. All the elements must be double, otherwise defValue array is returned. 
                let tryDV (xlKinds: Set<Kind>) (defValue1D: double[] option) (o1D: obj[])  =
                    let convert = Opt.def xlKinds None o1D
                    match convert |> Array.tryFind Option.isNone with
                    | None -> convert |> Array.map Option.get |> Some
                    | Some _ -> defValue1D

            [<RequireQualifiedAccess>]
            module Intg =
                /// Converts an obj[] to a int[], given a default-value for non-int elements.
                let def (defValue: int) (o1D: obj[]) =
                    o1D |> Array.map (D0.Intg.def defValue)

                /// optional-type default
                module Opt = 
                    /// Converts an obj[] to a ('a option)[], given a default-value for non-int elements.
                    let def (defValue: int option) (o1D: obj[]) =
                        o1D |> Array.map (D0.Intg.Opt.def defValue)

                /// Converts an obj[] to a int[], removing any non-int element.
                let filter (o1D: obj[]) =
                    o1D |> Array.choose (D0.Intg.Opt.def None)

                /// Converts an obj[] to an optional 'a[]. All the elements must be int, otherwise defValue array is returned. 
                let tryDV (defValue1D: int[] option) (o1D: obj[])  =
                    let convert = Opt.def None o1D
                    match convert |> Array.tryFind Option.isNone with
                    | None -> convert |> Array.map Option.get |> Some
                    | Some _ -> defValue1D

                /// Converts an obj[] to a int[]. All the elements must be int, fails otherwise. 
                let fail (o1D: obj[]) = match tryDV None o1D with | Some a1D -> a1D | None -> failwith "Cannot cast elements to int."

            [<RequireQualifiedAccess>]
            module Dte =
                /// Converts an obj[] to a DateTime[], given a default-value for non-DateTime elements.
                let def (defValue: DateTime) (o1D: obj[]) =
                    o1D |> Array.map (D0.Dte.def defValue)

                /// optional-type default
                module Opt = 
                    /// Converts an obj[] to a ('a option)[], given a default-value for non-DateTime elements.
                    let def (defValue: DateTime option) (o1D: obj[]) =
                        o1D |> Array.map (D0.Dte.Opt.def defValue)

                /// Converts an obj[] to a DateTime[], removing any non-DateTime element.
                let filter (o1D: obj[]) =
                    o1D |> Array.choose (D0.Dte.Opt.def None)

                /// Converts an obj[] to an optional 'a[]. All the elements must be DateTime, otherwise defValue array is returned. 
                let tryDV (defValue1D: DateTime[] option) (o1D: obj[])  =
                    let convert = Opt.def None o1D
                    match convert |> Array.tryFind Option.isNone with
                    | None -> convert |> Array.map Option.get |> Some
                    | Some _ -> defValue1D

                /// Converts an obj[] to a DateTime[]. All the elements must be DateTime, fails otherwise. 
                let fail (o1D: obj[]) = match tryDV None o1D with | Some a1D -> a1D | None -> failwith "Cannot cast elements to DateTime."
    
            /// Useful functions for casting xl-arrays, given a type-tag (e.g. "int", "date", "double", "string"...)
            /// Use module Gen functions for their untyped versions.
            type TagFn =
                /// Converts an xl-value to a 'A[], given a typed default-value for elements which can't be cast to 'A.
                static member def<'A> (xlKinds: Set<Kind>) (rowWiseDef: bool option) (defValue: obj option) (typeTag: string) (xlValue: obj) : 'A[] = 
                    let o1D = Cast.to1D (rowWiseDef |> Option.defaultValue false) xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOL -> 
                        let defval = D0.TagFn.defaultValue<'A> typeTag defValue
                        def<'A> defval o1D
                    | STRING -> 
                        let defval = D0.TagFn.defaultValue<'A> typeTag defValue
                        def<'A> defval o1D
                    | DOUBLE -> 
                        let defval = D0.TagFn.defaultValue<'A> typeTag defValue
                        def<'A> defval o1D
                    | DOUBLENAN -> 
                        let defval = D0.TagFn.defaultValue<double> typeTag defValue
                        let a1D = Nan.def xlKinds defval o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A)
                    | INT -> 
                        let defval = D0.TagFn.defaultValue<int> typeTag defValue |> int
                        let a1D = Intg.def defval o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A)
                    | DATE -> 
                        let defval = D0.TagFn.defaultValue<DateTime> typeTag defValue
                        let a1D = Dte.def defval o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A)
                    | OBJ ->
                        let defval = 
                            match defValue with 
                            | None -> Unchecked.defaultof<'A> 
                            | Some defv -> Registry.MRegistry.defaultValue<'A> defv
                        let regObjs = o1D |> Array.map Registry.MRegistry.tryExtract<'A>
                        regObjs |> Array.map (Option.defaultValue defval)
                    | _ -> [||]
        
                /// Converts an xl-value to a ('A option)[], given an optional default-value for elements which can't be cast to 'A.
                static member defOpt<'A> (xlKinds: Set<Kind>) (rowWiseDef: bool option) (defValue: obj option) (typeTag: string) (xlValue: obj) : ('A option)[] = 
                    let o1D = Cast.to1D (rowWiseDef |> Option.defaultValue false) xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOLOPT -> 
                        let defval : 'A option = D0.TagFn.defaultValueOpt<'A> defValue
                        Opt.def<'A> defval o1D
                    | STRINGOPT -> 
                        let defval : 'A option = D0.TagFn.defaultValueOpt<'A> defValue
                        Opt.def<'A> defval o1D
                    | DOUBLEOPT -> 
                        let defval : 'A option = D0.TagFn.defaultValueOpt<'A> defValue
                        Opt.def<'A> defval o1D
                    | DOUBLENANOPT -> 
                        let defval = D0.TagFn.defaultValueOpt<double> defValue
                        let a1D = Nan.Opt.def xlKinds defval o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A option)
                    | INTOPT -> 
                        let defval = D0.TagFn.defaultValueOpt<double> defValue |> Option.map (int)
                        let a1D = Intg.Opt.def defval o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A option)
                    | DATEOPT -> 
                        let defval = D0.TagFn.defaultValueOpt<double> defValue |> Option.map (fun d -> DateTime.FromOADate(d))
                        let a1D = Dte.Opt.def defval o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A option)
                    | OBJ -> 
                        let defval = defValue |> Option.map (Registry.MRegistry.tryExtract<'A>) |> Option.flatten
                        let regObjs = o1D |> Array.map Registry.MRegistry.tryExtract<'A>
                        regObjs |> Array.map (fun ao -> match ao with | None -> defval | _ -> ao)
                    | _ -> [||]
                
                static member filter<'A> (xlKinds: Set<Kind>) (rowWiseDef: bool option) (typeTag: string) (xlValue: obj) : 'A[] = 
                    let o1D = Cast.to1D (rowWiseDef |> Option.defaultValue false) xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOL -> filter<'A> o1D
                    | STRING -> 
                        let a1D = Stg.filter o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A)
                    | DOUBLE -> filter<'A> o1D
                    | DOUBLENAN -> 
                        let a1D = Nan.filter xlKinds o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A)
                    | INT -> 
                        let a1D = Intg.filter o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A)
                    | DATE -> 
                        let a1D = Dte.filter o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A)
                    | OBJ -> 
                        let regObjs = o1D |> Array.map Registry.MRegistry.tryExtract<'A>
                        regObjs |> Array.choose id
                    | _ -> [||]

                static member tryDV<'A> (xlKinds: Set<Kind>) (rowWiseDef: bool option) (defValue1D: 'A[] option) (typeTag: string) (xlValue: obj) : 'A[] option = 
                    let o1D = Cast.to1D (rowWiseDef |> Option.defaultValue false) xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOL -> tryDV<'A> defValue1D o1D
                    | STRING -> tryDV<'A> defValue1D o1D
                    | DOUBLE -> tryDV<'A> defValue1D o1D
                    | DOUBLENAN -> 
                        let defval1D = box defValue1D :?> (double[] option)
                        let a1D = Nan.tryDV xlKinds defval1D o1D
                        box a1D :?> 'A[] option
                    | INT -> 
                        let defval1D = box defValue1D :?> (int[] option)
                        let a1D = Intg.tryDV defval1D o1D
                        box a1D :?> 'A[] option
                    | DATE -> 
                        let defval1D = box defValue1D :?> (DateTime[] option)
                        let a1D = Dte.tryDV defval1D o1D
                        box a1D :?> 'A[] option
                    | OBJ ->
                        let regObjs = o1D |> Array.map Registry.MRegistry.tryExtract<'A>
                        match regObjs |> Array.tryFind Option.isNone with
                        | None -> regObjs |> Array.map Option.get |> Some
                        | Some _ -> defValue1D
                    | _ -> None

                static member tryDVBxd<'A> (xlKinds: Set<Kind>) (rowWiseDef: bool option) (defValue1D: 'A[] option) (typeTag: string) (xlValue: obj) : obj[] option = 
                    TagFn.tryDV<'A> xlKinds rowWiseDef defValue1D typeTag xlValue
                    |> Option.map (Array.map box)

                static member tryEmpty<'A> (xlKinds: Set<Kind>) (rowWiseDef: bool option) (typeTag: string) (xlValue: obj) : 'A[] = 
                    let defValue1D : 'A[] = [||]
                    TagFn.tryDV<'A> xlKinds rowWiseDef (Some defValue1D) typeTag xlValue
                    |> Option.get



                ///// TODO: wording
                //static member empty<'A> (typeTag: string) : 'A[] = [||]
                //static member emptyOpt<'A> (typeTag: string) : 'A[] option = Some [||]


            /// Useful functions for casting xl-arrays, given a type-tag (e.g. "int", "date", "double", "string"...)
            /// Use type GenFn functions for their typed versions.
            [<RequireQualifiedAccess>]
            module Tag =
                /// Converts an xl-value to a 'a[], given a type-tag and a compatible default-value for when casting to 'a fails.
                /// The type-tag determines 'a. Only works for non-optional type-tags, e.g. "string".
                let def (xlKinds: Set<Kind>) (rowWiseDef: bool option) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype =
                        if typeTag.ToUpper() = "OBJ" then
                            Registry.MRegistry.trySampleType false xlValue |> Option.get // assumes a type is found. TODO: improve this? (when type not found)
                        else
                            typeTag |> Variant.labelType false
                    let args : obj[] = [| xlKinds; rowWiseDef; defValue; typeTag; xlValue |]
                    let res = Toolbox.Generics.invoke<TagFn> "def" [| gentype |] args
                    res

                module Opt =
                    /// Converts an xl-value to a ('a option)[], given a type-tag and a compatible default-value for when casting to 'a fails.
                    /// The type-tag determines 'a. Only works for optional type-tags, e.g. "#string".
                    let def (xlKinds: Set<Kind>) (rowWiseDef: bool option) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                        // let gentype = typeTag |> Variant.labelType true
                        let gentype =
                            if typeTag.ToUpper() = "#OBJ" then
                                // if provided, and if defValue is a R-obj, then its type is used for gentype.
                                // if not provided, then try to find the first R-obj within xlValue and use its type for gentype.
                                match defValue with
                                | None -> Registry.MRegistry.trySampleType false xlValue |> Option.get // assumes a type is found, will fail otherwise.
                                | Some defval -> Registry.MRegistry.trySampleType false defval |> Option.get
                            else
                                typeTag |> Variant.labelType true

                        let args : obj[] = [| xlKinds; rowWiseDef; defValue; typeTag; xlValue |]
                        let res = Toolbox.Generics.invoke<TagFn> "defOpt" [| gentype |] args
                        res

                /// For when the type-tag is either optional, e.g. "#string", or not, e.g. "string".
                module Any =
                    /// Convenient, single function covering def and Opt.def cases.
                    /// The returned (boxed) array might be either a 'a[] or a ('a option)[], depending on wether the type-tag is optional or not.
                    let def (xlKinds: Set<Kind>) (rowWiseDef: bool option) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                        // let gentype = typeTag |> Variant.labelType true
                        if typeTag |> isOptionalType then
                            Opt.def xlKinds rowWiseDef defValue typeTag xlValue
                        else
                            def xlKinds rowWiseDef defValue typeTag xlValue

                /// TODO: wording here. Mentioning the output is a (boxed) 'a[] where 'a is determined by the type tag
                // TODO explain trySampleType strict
                let filter' (xlKinds: Set<Kind>) (rowWiseDef: bool option) (strict: bool) (typeTag: string) (xlValue: obj) : Type*obj = 
                    let gentype =
                        if typeTag.ToUpper() = "OBJ" then
                            Registry.MRegistry.trySampleType strict xlValue |> Option.get // assumes a type is found. TODO: improve this? (when type not found)
                        else
                            typeTag |> Variant.labelType false
                    let args : obj[] = [| xlKinds; rowWiseDef; typeTag; xlValue |]
                    let res = Toolbox.Generics.invoke<TagFn> "filter" [| gentype |] args
                    gentype, res
                
                /// TODO: wording here. Mentioning the output is a (boxed) 'a[] where 'a is determined by the type tag
                let filter (xlKinds: Set<Kind>) (rowWiseDef: bool option) (typeTag: string) (xlValue: obj) : obj = 
                    filter' xlKinds rowWiseDef false typeTag xlValue |> snd

                let tryDV' (methodName: string) (xlKinds: Set<Kind>) (rowWiseDef: bool option) (defValue1D: obj) (typeTag: string) (xlValue: obj) : Type*obj = 
                    let gentype =
                        if typeTag.ToUpper() = "OBJ" then
                            Registry.MRegistry.trySampleType false xlValue |> Option.get // assumes a type is found. TODO: improve this? (when type not found)
                        else
                            typeTag |> Variant.labelType false
                    let args : obj[] = [| xlKinds; rowWiseDef; defValue1D; typeTag; xlValue |]
                    let res = Toolbox.Generics.invoke<TagFn> methodName [| gentype |] args
                    gentype, res

                /// Given a type-tag, which determines the expected array's element-type, 'a, converts an xl-value to a boxed (Some 'a[]) on success 
                /// or returns defValue1D, a boxed ('a[] option), on failure.
                /// Only works with non-optional type-tags, e.g. "string", but not "#string".
                /// Recap: 
                ///    - On success, returns boxed ('a, boxed ('a[] option)), where 'a is the array's element-type.
                ///    - On failure, returns boxed (obj, defValue1D), where defValue1D is a boxed ('a[] option).
                let tryDV (xlKinds: Set<Kind>) (rowWiseDef: bool option) (defValue1D: obj) (typeTag: string) (xlValue: obj) : Type*obj = 
                    tryDV' "tryDV" xlKinds rowWiseDef defValue1D typeTag xlValue

                /// Similiar to tryDV but with boxed elements.
                /// Given a type-tag, which determines the expected array's element-type, 'a, converts an xl-value to a boxed (Some (boxed 'a)[]) on success 
                /// or returns defValue1D, a boxed ('a[] option), on failure.
                /// Only works with non-optional type-tags, e.g. "string", but not "#string".
                /// Recap: 
                ///    - On success, returns boxed ('a, boxed ((boxed 'a)[] option)), where 'a is the array's element-type.
                ///    - On failure, returns boxed (obj, defValue1D), where defValue1D is a boxed ('a[] option).
                let tryDVBxd (xlKinds: Set<Kind>) (rowWiseDef: bool option) (defValue1D: obj) (typeTag: string) (xlValue: obj) : Type*obj = 
                    tryDV' "tryDVBxd" xlKinds rowWiseDef defValue1D typeTag xlValue

                /// Similiar tryDV but returns an empty array on failure.
                /// Given a type-tag, which determines the expected array's element-type, 'a, converts an xl-value to a boxed ('a[]) on success 
                /// or returns boxed (empty [||]), on failure.
                /// Only works with non-optional type-tags, e.g. "string", but not "#string".
                /// Recap: 
                ///    - On success, returns boxed ('a, boxed ('a[])), where 'a is the array's element-type.
                ///    - On failure, returns boxed (obj, boxed [||]).
                let tryEmpty (xlKinds: Set<Kind>) (rowWiseDef: bool option) (typeTag: string) (xlValue: obj) : obj = // TODO : for consistency, should add the type?
                    let gentype =
                        if typeTag.ToUpper() = "OBJ" then
                            Registry.MRegistry.trySampleType false xlValue |> Option.get // assumes a type is found. TODO: improve this? (when type not found)
                        else
                            typeTag |> Variant.labelType false
                    let args : obj[] = [| xlKinds; rowWiseDef; typeTag; xlValue |]
                    let res = Toolbox.Generics.invoke<TagFn> "tryEmpty" [| gentype |] args
                    res

                // Same as the tryDV' functions above but with unboxing of the results.
                module Try =
                    /// Given a type-tag, which determines the expected array's element-type, 'a, converts an xl-value to an optional 'a[] on success 
                    /// or returns defValue1D, a boxed ('a[] option), on failure.
                    /// Only works with non-optional type-tags, e.g. "string", but not "#string".
                    /// Recap: 
                    ///    - Returns Some ('a, boxed ('a[]) or Some (obj, defValue1D unwrapped), where 'a is the array's element-type and defValue1D is a boxed ('a[] option).
                    ///    - or returns None on failure, and if defValue is (boxed) None.
                    let tryDV (xlKinds: Set<Kind>) (rowWiseDef: bool option) (defValue1D: obj) (typeTag: string) (xlValue: obj) : (Type*obj) option = 
                        let genty, xa1D = tryDV xlKinds rowWiseDef defValue1D typeTag xlValue
                        Toolbox.Option.unwrap xa1D
                        |> Option.map (fun res -> (genty, res))

                    /// Same as tryDV but with boxed elements.
                    /// Given a type-tag, which determines the expected array's element-type, 'a, converts an xl-value to an optional (boxed 'a)[] on success 
                    /// or returns defValue1D, a boxed ('a[] option), on failure.
                    /// Only works with non-optional type-tags, e.g. "string", but not "#string".
                    /// Recap: 
                    ///    - Returns Some ('a, (boxed 'a)[]) or Some (obj, defValue1D unwrapped), where 'a is the array's element-type and defValue1D is a boxed ('a[] option).
                    ///    - or returns None on failure, and if defValue is (boxed) None.
                    let tryDVBxd (xlKinds: Set<Kind>) (rowWiseDef: bool option) (defValue1D: obj) (typeTag: string) (xlValue: obj) : (Type*(obj[])) option = 
                        let genty, xo1D = tryDVBxd xlKinds rowWiseDef defValue1D typeTag xlValue
                        Toolbox.Option.unwrap xo1D
                        |> Option.map (fun res -> (genty, res :?> obj[]))

        /// Obj[] input functions.
        module D2 =
            //open Excel
            open type Variant

            let empty2D<'a> : 'a[,] = [|[||]|] |> array2D
            let isEmpty (a2D: 'a[,]) : bool = a2D |> Array2D.length1 = 0 // is this the right way?
            let singleton<'a> (a: 'a) = Array2D.create 1 1 a

            // non-optional-type default
            /// Converts an obj[,] to a 'a[,], given a typed default-value for elements which can't be cast to 'a.
            let def<'a> (defValue: 'a) (o2D: obj[,]) : 'a[,] =
                o2D |> Array2D.map (D0.def<'a> defValue)

            // optional-type default
            module Opt =
                /// Converts an obj[,] to a ('a option)[,], given an optional default-value for elements which can't be cast to 'a.
                let def<'a> (defValue: 'a option) (o2D: obj[,]) : ('a option)[,] =
                    o2D |> Array2D.map (D0.Opt.def<'a> defValue)

            /// Converts an obj[,] to a 'a[,], removing either row or column where any element can't be cast to 'a.
            let filter<'a> (rowWise: bool) (o2D: obj[,]) : 'a[,] =
                let len1, len2 = o2D |> Array2D.length1, o2D |> Array2D.length2
                
                if rowWise then
                    [| for i in 0 .. (len1 - 1) -> 
                        match D1.tryDV<'a> None o2D.[i,*] with
                        | None -> [||]
                        | Some row1D -> row1D
                    |]
                else
                    // FIXME needs to be transposed !!!!!!!!!!
                    [| for j in 0 .. (len2 - 1) -> 
                        match D1.tryDV<'a> None o2D.[*,j] with
                        | None -> [||]
                        | Some col1D -> col1D
                    |]
                |> Array.filter (fun a1D -> a1D |> Array.isEmpty |> not)
                |> array2D

            /// Converts an obj[,] to an optional 'a[,]. All the elements must match the given type, otherwise defValue array is returned. 
            let tryDV<'a> (defValue2D: 'a[,] option) (o2D: obj[,]) : 'a[,] option =
                let len1 = o2D |> Array2D.length1
                let convert = Opt.def None o2D

                let hasNones = 
                    [| for i in 0 .. (len1 - 1) ->
                        convert.[i,*] |> Array.filter Option.isNone
                    |]
                    |> Array.filter (fun o1D -> o1D |> Array.isEmpty |> not)

                if hasNones |> Array.isEmpty then
                    convert |> Array2D.map Option.get |> Some
                else defValue2D

            [<RequireQualifiedAccess>]
            module Bool =
                /// Converts an obj[,] to a bool[,], given a bool default-value for when casting to bool fails.
                let def (defValue: bool) (o2D: obj[,]) : bool[,] = def<bool> defValue o2D

                // optional-type default
                module Opt =
                    /// Converts an obj[,] to a (bool option)[,], given a bool default-value for when casting to bool fails.
                    let def (defValue: bool option) (o2D: obj[,]) : (bool option)[,] = Opt.def defValue o2D

                /// Converts an obj[,] to a bool[,], removing either row or column where any element isn't a (boxed) bool.
                let filter (rowWise: bool) (o2D: obj[,]) : bool[,] = filter<bool> rowWise o2D

                /// Converts an obj[,] to an optional 'a[,]. All the elements must be bools, otherwise defValue array is returned. 
                let tryDV (defValue2D: bool[,] option) (o2D: obj[,]) : bool[,] option = tryDV<bool> defValue2D o2D

            [<RequireQualifiedAccess>]
            module Stg =
                /// Converts an obj[,] to a bool[,], given a bool default-value for when casting to string fails.
                let def (defValue: string) (o2D: obj[,]) : string[,] = def<string> defValue o2D

                // optional-type default
                module Opt =
                    /// Converts an obj[,] to a (bool option)[,], given a bool default-value for when casting to string fails.
                    let def (defValue: string option) (o2D: obj[,]) : (string option)[,] = Opt.def defValue o2D

                /// Converts an obj[,] to a bool[,], removing either row or column where any element isn't a (boxed) string.
                let filter (rowWise: bool) (o2D: obj[,]) : string[,] = filter<string> rowWise o2D

                /// Converts an obj[,] to an optional 'a[,]. All the elements must be strings, otherwise defValue array is returned. 
                let tryDV (defValue2D: string[,] option) (o2D: obj[,]) : string[,] option = tryDV<string> defValue2D o2D

            [<RequireQualifiedAccess>]
            module Dbl =
                /// Converts an obj[,] to a bool[,], given a bool default-value for when casting to double fails.
                let def (defValue: double) (o2D: obj[,]) : double[,] = def<double> defValue o2D

                // optional-type default
                module Opt =
                    /// Converts an obj[,] to a (bool option)[,], given a bool default-value for when casting to double fails.
                    let def (defValue: double option) (o2D: obj[,]) : (double option)[,] = Opt.def defValue o2D

                /// Converts an obj[,] to a bool[,], removing either row or column where any element isn't a (boxed) string.
                let filter (rowWise: bool) (o2D: obj[,]) : double[,] = filter<double> rowWise o2D

                /// Converts an obj[,] to an optional 'a[,]. All the elements must be doubles, otherwise defValue array is returned. 
                let tryDV (defValue2D: double[,] option) (o2D: obj[,]) : double[,] option = tryDV<double> defValue2D o2D

            [<RequireQualifiedAccess>]
            module Nan = 
                /// Converts an obj[,] to a bool[,], given a bool default-value for when casting to double fails.
                let def (xlKinds: Set<Kind>) (defValue: double) (o2D: obj[,]) : double[,] = 
                    o2D |> Array2D.map (D0.Nan.def xlKinds defValue)

                // optional-type default
                module Opt =
                    /// Converts an obj[,] to a (bool option)[,], given a bool default-value for when casting to double fails.
                    let def (xlKinds: Set<Kind>) (defValue: double option) (o2D: obj[,]) : (double option)[,] = 
                        o2D |> Array2D.map (D0.Nan.Opt.def xlKinds defValue)

                /// Converts an obj[,] to a bool[,], removing either row or column where any element isn't a (boxed) string.
                let filter (xlKinds: Set<Kind>) (rowWise: bool) (o2D: obj[,]) : double[,] = 
                    o2D 
                    |> Array2D.map (D0.Nan.nanify xlKinds)
                    |> filter<double> rowWise

                /// Converts an obj[,] to an optional 'a[,]. All the elements must be doubles, otherwise defValue array is returned. 
                let tryDV (xlKinds: Set<Kind>) (defValue2D: double[,] option) (o2D: obj[,]) : double[,] option = 
                    o2D 
                    |> Array2D.map (D0.Nan.nanify xlKinds)
                    |> tryDV<double> defValue2D

            [<RequireQualifiedAccess>]
            module Intg = 
                /// Converts an obj[,] to a bool[,], given a bool default-value for when casting to double fails.
                let def (defValue: int) (o2D: obj[,]) : int[,] =
                    o2D |> Array2D.map (D0.Intg.def defValue)

                // optional-type default
                module Opt =
                    /// Converts an obj[,] to a (bool option)[,], given a bool default-value for when casting to double fails.
                    let def (defValue: int option) (o2D: obj[,]) : (int option)[,] = 
                        o2D |> Array2D.map (D0.Intg.Opt.def defValue)

                /// Converts an obj[,] to a bool[,], removing either row or column where any element isn't a (boxed) string.
                let filter (rowWise: bool) (o2D: obj[,]) : int[,] = 
                    let len1, len2 = o2D |> Array2D.length1, o2D |> Array2D.length2
                
                    if rowWise then
                        [| for i in 0 .. (len1 - 1) -> 
                            match D1.Intg.tryDV None o2D.[i,*] with
                            | None -> [||]
                            | Some row1D -> row1D
                        |]
                    else
                        // FIXME needs to be transposed !!!!!!!!!!
                        [| for j in 0 .. (len2 - 1) -> 
                            match D1.Intg.tryDV None o2D.[*,j] with
                            | None -> [||]
                            | Some col1D -> col1D
                        |]
                    |> Array.filter (fun a1D -> a1D |> Array.isEmpty |> not)
                    |> array2D

                /// Converts an obj[,] to an optional 'a[,]. All the elements must be doubles, otherwise defValue array is returned. 
                let tryDV (defValue2D: int[,] option) (o2D: obj[,]) : int[,] option = 
                    let len1 = o2D |> Array2D.length1
                    let convert = Opt.def None o2D

                    let hasNones = 
                        [| for i in 0 .. (len1 - 1) ->
                            convert.[i,*] |> Array.filter Option.isNone
                        |]
                        |> Array.filter (fun o1D -> o1D |> Array.isEmpty |> not)

                    if hasNones |> Array.isEmpty then
                        convert |> Array2D.map Option.get |> Some
                    else defValue2D

            module Dte = 
                /// Converts an obj[,] to a bool[,], given a bool default-value for when casting to double fails.
                let def (defValue: DateTime) (o2D: obj[,]) : DateTime[,] =
                    o2D |> Array2D.map (D0.Dte.def defValue)

                // optional-type default
                module Opt =
                    /// Converts an obj[,] to a (bool option)[,], given a bool default-value for when casting to double fails.
                    let def (defValue: DateTime option) (o2D: obj[,]) : (DateTime option)[,] = 
                        o2D |> Array2D.map (D0.Dte.Opt.def defValue)

                /// Converts an obj[,] to a DateTime[,], removing either row or column where any element isn't a (boxed) DateTime.
                let filter (rowWise: bool) (o2D: obj[,]) : DateTime[,] = 
                    let len1, len2 = o2D |> Array2D.length1, o2D |> Array2D.length2
                
                    if rowWise then
                        [| for i in 0 .. (len1 - 1) -> 
                            match D1.Dte.tryDV None o2D.[i,*] with
                            | None -> [||]
                            | Some row1D -> row1D
                        |]
                    else
                        // FIXME needs to be transposed !!!!!!!!!!
                        [| for j in 0 .. (len2 - 1) -> 
                            match D1.Dte.tryDV None o2D.[*,j] with
                            | None -> [||]
                            | Some col1D -> col1D
                        |]
                    |> Array.filter (fun a1D -> a1D |> Array.isEmpty |> not)
                    |> array2D

                /// Converts an obj[,] to an optional DateTime[,]. All the elements must be (boxed) DateTime, otherwise defValue array is returned. // TODO wording
                let tryDV (defValue2D: DateTime[,] option) (o2D: obj[,]) : DateTime[,] option = 
                    let len1 = o2D |> Array2D.length1
                    let convert = Opt.def None o2D

                    let hasNones = 
                        [| for i in 0 .. (len1 - 1) ->
                            convert.[i,*] |> Array.filter Option.isNone
                        |]
                        |> Array.filter (fun o1D -> o1D |> Array.isEmpty |> not)

                    if hasNones |> Array.isEmpty then
                        convert |> Array2D.map Option.get |> Some
                    else defValue2D

            /// Useful functions for casting xl-arrays, given a type-tag (e.g. "int", "date", "double", "string"...)
            /// Use module Gen functions for their untyped versions.
            type TagFn =
                /// Converts an xl-value to a 'A[], given a typed default-value for elements which can't be cast to 'A.
                static member def<'A> (xlKinds: Set<Kind>) (defValue: obj option) (typeTag: string) (xlValue: obj) : 'A [,] = 
                    let o2D = Cast.to2D xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOL -> 
                        let defval = D0.TagFn.defaultValue<'A> typeTag defValue
                        def<'A> defval o2D
                    | STRING -> 
                        let defval = D0.TagFn.defaultValue<'A> typeTag defValue
                        def<'A> defval o2D
                    | DOUBLE -> 
                        let defval = D0.TagFn.defaultValue<'A> typeTag defValue
                        def<'A> defval o2D
                    | DOUBLENAN -> 
                        let defval = D0.TagFn.defaultValue<double> typeTag defValue
                        let a2D = Nan.def xlKinds defval o2D
                        a2D |> Array2D.map (fun x -> (box x) :?> 'A)
                    | INT -> 
                        let defval = D0.TagFn.defaultValue<int> typeTag defValue |> int
                        let a2D = Intg.def defval o2D
                        a2D |> Array2D.map (fun x -> (box x) :?> 'A)
                    | DATE -> 
                        let defval = D0.TagFn.defaultValue<DateTime> typeTag defValue
                        let a2D = Dte.def defval o2D
                        a2D |> Array2D.map (fun x -> (box x) :?> 'A)
                    | _ -> empty2D<'A>
        
                /// Converts an xl-value to a ('A option)[], given an optional default-value for elements which can't be cast to 'A.
                static member defOpt<'A> (xlKinds: Set<Kind>) (defValue: obj option) (typeTag: string) (xlValue: obj) : ('A option)[,] = 
                    let o2D = Cast.to2D xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOLOPT -> 
                        let defval : 'A option = D0.TagFn.defaultValueOpt<'A> defValue
                        Opt.def<'A> defval o2D
                    | STRINGOPT -> 
                        let defval : 'A option = D0.TagFn.defaultValueOpt<'A> defValue
                        Opt.def<'A> defval o2D
                    | DOUBLEOPT -> 
                        let defval : 'A option = D0.TagFn.defaultValueOpt<'A> defValue
                        Opt.def<'A> defval o2D
                    | DOUBLENANOPT -> 
                        let defval = D0.TagFn.defaultValueOpt<double> defValue
                        let a2D = Nan.Opt.def xlKinds defval o2D
                        a2D |> Array2D.map (fun x -> (box x) :?> 'A option)
                    | INTOPT -> 
                        let defval = D0.TagFn.defaultValueOpt<double> defValue |> Option.map (int)
                        let a2D = Intg.Opt.def defval o2D
                        a2D |> Array2D.map (fun x -> (box x) :?> 'A option)
                    | DATEOPT -> 
                        let defval = D0.TagFn.defaultValueOpt<double> defValue |> Option.map (fun d -> DateTime.FromOADate(d))
                        let a2D = Dte.Opt.def defval o2D
                        a2D |> Array2D.map (fun x -> (box x) :?> 'A option)
                    | _ -> empty2D<'A option>
                
                static member filter<'A> (xlKinds: Set<Kind>) (rowWise: bool) (typeTag: string) (xlValue: obj) : 'A[,] = 
                    let o2D = Cast.to2D xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOL -> filter<'A> rowWise o2D
                    | STRING -> filter<'A> rowWise o2D
                    | DOUBLE -> filter<'A> rowWise o2D
                    | DOUBLENAN -> 
                        let a2D = Nan.filter xlKinds rowWise o2D
                        a2D |> Array2D.map (fun x -> (box x) :?> 'A)
                    | INT -> 
                        let a2D = Intg.filter rowWise o2D
                        a2D |> Array2D.map (fun x -> (box x) :?> 'A)
                    | DATE -> 
                        let a2D = Dte.filter rowWise o2D
                        a2D |> Array2D.map (fun x -> (box x) :?> 'A)
                    | _ -> empty2D<'A>

                static member tryDV<'A> (xlKinds: Set<Kind>) (defValue2D: 'A[,] option) (typeTag: string) (xlValue: obj) : 'A [,] option = 
                    let o2D = Cast.to2D xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOL -> tryDV<'A> defValue2D o2D
                    | STRING -> tryDV<'A> defValue2D o2D
                    | DOUBLE -> tryDV<'A> defValue2D o2D
                    | DOUBLENAN ->
                        let defval2D = box defValue2D :?> (double[,] option)
                        let a2D = Nan.tryDV xlKinds defval2D o2D
                        box a2D :?> 'A[,] option
                    | INT -> 
                        let defval2D = box defValue2D :?> (int[,] option)
                        let a2D = Intg.tryDV defval2D o2D
                        box a2D :?> 'A[,] option
                    | DATE -> 
                        let defval2D = box defValue2D :?> (DateTime[,] option)
                        let a2D = Dte.tryDV defval2D o2D
                        box a2D :?> 'A[,] option
                    | _ -> None

            /// Useful functions for casting xl-arrays, given a type-tag (e.g. "int", "date", "double", "string"...)
            /// Use type GenFn functions for their typed versions.
            [<RequireQualifiedAccess>]
            module Tag =
                /// Converts an xl-value to a 'a[], given a type-tag and a compatible default-value for when casting to 'a fails.
                /// The type-tag determines 'a. Only works for non-optional type-tags, e.g. "string".
                let def (xlKinds: Set<Kind>) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype = typeTag |> Variant.labelType true
                    let args : obj[] = [| xlKinds; defValue; typeTag; xlValue |]
                    let res = Toolbox.Generics.invoke<TagFn> "def" [| gentype |] args
                    res

                module Opt =
                    /// Converts an xl-value to a ('a option)[], given a type-tag and a compatible default-value for when casting to 'a fails.
                    /// The type-tag determines 'a. Only works for optional type-tags, e.g. "#string".
                    let def (xlKinds: Set<Kind>) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                        let gentype = typeTag |> Variant.labelType true
                        let args : obj[] = [| xlKinds; defValue; typeTag; xlValue |]
                        let res = Toolbox.Generics.invoke<TagFn> "defOpt" [| gentype |] args
                        res

                /// For when the type-tag is either optional, e.g. "#string", or not, e.g. "string".
                module Any =
                    /// Convenient, single function covering def and Opt.def cases.
                    /// The returned (boxed) array might be either a 'a[] or a ('a option)[], depending on wether the type-tag is optional or not.
                    let def (xlKinds: Set<Kind>) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                        let gentype = typeTag |> Variant.labelType true
                        let args : obj[] = [| xlKinds; defValue; typeTag; xlValue |]

                        let res =
                            if typeTag |> isOptionalType then
                                Toolbox.Generics.invoke<TagFn> "defOpt" [| gentype |] args
                            else
                                Toolbox.Generics.invoke<TagFn> "def" [| gentype |] args
                        res

                let filter (xlKinds: Set<Kind>) (rowWise: bool) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype = typeTag |> Variant.labelType false
                    let args : obj[] = [| xlKinds; rowWise; typeTag; xlValue |]
                    let res = Toolbox.Generics.invoke<TagFn> "filter" [| gentype |] args
                    res

                let tryDV' (xlKinds: Set<Kind>) (defValue2D: obj) (typeTag: string) (xlValue: obj) : Type*obj = 
                    let gentype =
                        if typeTag.ToUpper() = "OBJ" then
                            Registry.MRegistry.trySampleType false xlValue |> Option.get // assumes a type is found. TODO: improve this? (when type not found)
                        else
                            typeTag |> Variant.labelType false
                    let args : obj[] = [| xlKinds; defValue2D; typeTag; xlValue |]
                    let res = Toolbox.Generics.invoke<TagFn> "tryDV" [| gentype |] args
                    gentype, res

                let tryDV (xlKinds: Set<Kind>) (defValue2D: obj) (typeTag: string) (xlValue: obj) : obj = 
                    tryDV' xlKinds defValue2D typeTag xlValue |> snd

                //let tryDVTBD (defValue2D: obj) (typeTag: string) (xlValue: obj) : obj = 
                //    let gentype = typeTag |> Variant.labelType false
                //    let args : obj[] = [| defValue2D; typeTag; xlValue |]
                //    let res = Toolbox.Generics.invoke<TagFn> "tryDV" [| gentype |] args
                //    res

                // FIXME: wording. Same as tryDV' with unboxing
                module Try =
                    let tryDV' (xlKinds: Set<Kind>) (defValue2D: obj) (typeTag: string) (xlValue: obj) : (Type*obj) option = 
                        let genty, xa2D = tryDV' xlKinds defValue2D typeTag xlValue
                        Toolbox.Option.unwrap xa2D
                        |> Option.map (fun res -> (genty, res))

                    let tryDV (xlKinds: Set<Kind>) (defValue2D: obj) (typeTag: string) (xlValue: obj) : obj option = 
                        let xa2D = tryDV xlKinds defValue2D typeTag xlValue
                        let res = Toolbox.Option.unwrap xa2D
                        res

                /// 2D arrays with row-wise (column-wise) typed elements, where elements in a given row (given column) have the same type.
                module Multi = 
                    // TODO : wording
                    /// Converts an obj[,] to an optional 'a[,]. 
                    /// rowWise = true: All elements in a given row must match the correspongind type-tag, for all rows, otherwise defValue2D array is returned. 
                    /// rowWise = false: All elements in a given column must match the correspongind type-tag, for all columns, otherwise defValue2D array is returned. 
                    let tryDV (defValue2D: obj[,] option) (xlKinds: Set<Kind>) (rowWise: bool) (typeTags: string[]) (xlValue: obj) : (Type[]*(obj[,])) option =
                        let o2D = Cast.to2D xlValue
                        //let len1, len2 = o2D |> Array2D.length1, o2D |> Array2D.length2

                        let tyxs' = 
                            if rowWise then
                                [| for i in o2D.GetLowerBound(0) .. o2D.GetUpperBound(0) -> D1.Tag.Try.tryDVBxd xlKinds None (box None) typeTags.[i] o2D.[i,*] |]
                            else
                                [| for j in o2D.GetLowerBound(1) .. o2D.GetUpperBound(1) -> D1.Tag.Try.tryDVBxd xlKinds None (box None) typeTags.[j] o2D.[*,j] |]

                        match tyxs' |> Array.tryFind Option.isNone with
                        | Some _ -> 
                            defValue2D |> Option.map (fun a2D -> ([||], a2D)) // TODO: must extract the types list from defValue2D? ([||] will only work when defValue2D is None)
                        | None -> 
                            let tyxs = tyxs' |> Array.map Option.get
                            let (gentys, xa1Ds) = Array.unzip tyxs
                            (gentys, Toolbox.Array2D.concat2D rowWise xa1Ds)
                            |> Some

    /// Functions to retun values to Excel.
    module Out =
        open type Variant

        /// Substitute output values to Excel.
        ///    - Proxys.empty for empty arrays.
        ///    - Proxys.failed for function failure.
        ///    - Proxys.nan for Double.NaN values.
        ///    - Proxys.none for optional F# None values.
        ///    - Proxys.object for non-primitive types values.
        type Proxys = { empty: obj; failed: obj; nan: obj; none: obj; object: obj } with
            static member def : Proxys = { empty = "<empty>"; failed = box ExcelError.ExcelErrorNA ; nan = ExcelError.ExcelErrorNA; none = "<none>"; object = "<obj>" }

        /// Functions to return single values to Excel.
        module D0 =
            /// Outputs double values to Excel, 
            /// with conversion of Double.NaN values, if any.
            [<RequireQualifiedAccess>]
            module Dbl =
                let out (proxys: Proxys) (d: double) : obj =
                    if Double.IsNaN(d) then
                        proxys.nan
                    else    
                        d |> box

                /// Similar to Dbl.out but with a defValue instead of a Proxys object.
                let outBxd (defValue: obj) (d: double) : obj =
                    if Double.IsNaN(d) then
                        defValue
                    else    
                        d |> box

            /// Outputs primitive types back to Excel.
            // https://docs.microsoft.com/en-us/office/client-developer/excel/data-types-used-by-excel
            [<RequireQualifiedAccess>]
            module Bxd =  // TODO : change name to Var(iant) rather than Primitive?
                /// Returns sensible Excel values for non-optional (boxed) primitive types.
                let out (proxys: Proxys) (o0D: obj) : obj =
                    match o0D with
                    | :? double as d -> Dbl.out proxys d
                    | :? string | :? DateTime | :? int | :? bool -> o0D
                    | _ -> proxys.object

                [<RequireQualifiedAccess>]
                /// Outputs optional primitive types:
                ///    - None will return proxys.none
                ///    - Some x will return (Bxd.out x)
                module Opt = 
                    let out (proxys: Proxys) (o0D: obj option) : obj =
                        match o0D with
                        | None -> proxys.none
                        | Some o0d -> o0d |> out proxys

                [<RequireQualifiedAccess>]
                /// Outputs optional and non-optional primitive types.
                /// Option on primitive types (boxed) will return as follow: 
                ///    - None will return proxys.none
                ///    - Some x will return (Bxd.out x)
                module Any = 
                    let out (proxys: Proxys) (o0D: obj) : obj =
                        o0D |> Toolbox.Option.map proxys.none (out proxys)

            [<RequireQualifiedAccess>]
            module Prm =  // TODO : change name to Var(iant) rather than Prm?
                /// Outputs to Excel:
                ///    - Primitives-type: Returns values directly to Excel.
                ///    - Any other type : Returns ReplaveValues.object.
                let out<'a> (proxys: Proxys) (o0D: obj) : obj =
                    o0D |> Toolbox.Option.map proxys.none (Bxd.Any.out proxys)

            [<RequireQualifiedAccess>]
            /// Outputs primitives types directly to Excel, but stores non-primitive types in the Registry.
            module Reg = 
                /// Outputs sensible values to Excel, depending on o0D input value and type:
                ///    - A primitive-type value is returned without change to Excel.
                ///    - A Double.NaN value is returned as proxys.nan.
                ///    - If unwrapOptions is true :
                ///        - Some (primtive-type-value) is returned as primtive-type-value.
                ///        - A None is returned as proxys.none. 
                ///    - Any other type: the value is stored in the Registry and a Registry key is returned to Excel.
                let out<'a> (append: bool) (unwrapOptions: bool) (refKey: String) (proxys: Proxys) (o0D: obj) : obj =
                    let ty = 
                        if unwrapOptions then
                            typeof<'a> |> Toolbox.Option.uType |> Option.defaultValue typeof<'a>
                        else
                            typeof<'a>

                    if ty |> Toolbox.Type.isPrimitive true then
                        o0D |> Toolbox.Option.map proxys.none (Bxd.Any.out proxys)
                    else
                        let regAdd = if append then Registry.MRegistry.append else Registry.MRegistry.register

                        if unwrapOptions then
                            o0D |> Toolbox.Option.map proxys.none (regAdd refKey >> box)
                        else
                            o0D |> regAdd refKey |> box
                
                // TODO : rewrite this function
                let outO (append: bool) (refKey: String) (proxys: Proxys) (xlValue: obj) : obj =
                    let mapping (o: obj) =
                        if o |> isNull then // protects the `ty = o.GetType()` snippet which fails on None values at runtime (= null values at runtime).
                            proxys.none
                        else
                            let ty = o.GetType()
                            if ty |> Toolbox.Type.isPrimitive false then
                                o |> Toolbox.Option.map proxys.none (Bxd.Any.out proxys)
                            else
                                let regAdd = if append then Registry.MRegistry.append else Registry.MRegistry.register
                                o |> Toolbox.Option.map proxys.none (regAdd refKey >> box)

                    match Registry.MRegistry.tryExtractO xlValue with
                    | None -> proxys.failed
                    | Some regObj -> regObj |> Toolbox.Option.map proxys.none mapping

    // -------------------------
    // -- Convenience functions
    // -------------------------

        // default-output function
        let out<'a> (defOutput: obj) (output: 'a option) = match output with None -> defOutput | Some value -> box value
        let outNA<'a> : 'a option -> obj = out (box ExcelError.ExcelErrorNA)
        let outStg<'a> (defString: string) : 'a option -> obj = out (box defString)
        let outDblTBD<'a> (defNum: double) : 'a option -> obj = out (box defNum)
        let outOptTBD<'a> (defNum: double) : 'a option -> obj = out (box defNum)



        /// Functions to output 1D arrays back to Excel.
        module D1 =
            /// Outputs arrays of (boxed) primitive (possibly optional) type elements back to Excel.
            /// Non primitive-type elements ouput as #VALUE!.
            [<RequireQualifiedAccess>]
            module Bxd =
                /// Returns sensible Excel values for 1D arrays of primitive-type elements.
                ///    - Primitive-type: Returns value directly to Excel.
                ///    - Double.NaN values will be returned as ReplaveValues.nan.
                ///    - Some (primtive-type-value) will be returned as primtive-type-value.
                ///    - None values will be returned as #VALUE!.
                ///    - Any other type: Returns ReplaveValues.object.
                ///    - Empty arrays will return [| proxys.empty |]. (Excel naturally returns empty array values as #VALUE!).
                let out (proxys: Proxys) (o1D: obj[]) : obj[] =
                    if o1D |> Array.isEmpty then
                        [| proxys.empty |]
                    else
                        o1D |> Array.map (D0.Bxd.out proxys)

                /// Case of arrays of optional type elements.
                [<RequireQualifiedAccess>]
                /// Similar to Out.D1.Bxd.out but for arrays of (boxed) optional type elements.
                module Opt =
                    let out (proxys: Proxys) (o1D: (obj option)[]) : obj[] =
                        if o1D |> Array.isEmpty then
                            [| proxys.empty |]
                        else
                            o1D |> Array.map (D0.Bxd.Opt.out proxys)
                
                /// Case of arrays of optional or non-optional type elements.
                [<RequireQualifiedAccess>]
                module Any = 
                    /// Out.D1.Bxd.out and Out.D1.Bxd.Opt.out combined.
                    /// Works both for (boxed) optional and (boxed) non-optional elements.
                    let out (proxys: Proxys) (o1D: obj[]) : obj[] =
                        if o1D |> Array.isEmpty then
                            [| proxys.empty |]
                        else
                            o1D |> Array.map (D0.Bxd.Any.out proxys)

            /// Outputs arrays of primitive (possibly optional) type elements back to Excel.
            /// Non primitive-type elements ouput as Proxys.object.
            [<RequireQualifiedAccess>]
            module Prm = 
                /// Returns sensible Excel values for 1D arrays depending on their element types:
                ///    - Primitive-type: Returns value directly to Excel.
                ///    - Double.NaN values will be returned as ReplaveValues.nan.
                ///    - Some (primtive-type-value) will be returned as primtive-type-value.
                ///    - None values will be returned as ReplaveValues.none.
                ///    - Any other type: Returns ReplaveValues.object.
                ///    - Empty arrays will return [| proxys.empty |]. (Excel naturally returns empty array values as #VALUE!).
                let out<'a> (proxys: Proxys) (a1D: 'a[]) : obj[] =
                    a1D 
                    |> Array.map box
                    |> Bxd.Any.out proxys

            /// Outputs primitive type arrays back to Excel.
            /// Stores non primitive-type elements in the Registry (and output a Registry key for each individual element).
            [<RequireQualifiedAccess>]
            module Reg = 
                /// Returns sensible Excel values for 1D arrays, depending on their elements' values and types:

                ///    - Primitive-type element-values are returned without change to Excel.
                ///    - Double.NaN element-values are returned as proxys.nan.
                ///    - If unwrapOptions is true :
                ///        - Some (primtive-type-value) element-values are returned as primtive-type-value.
                ///        - None element-values are returned as proxys.none.
                ///    - Any other type: Each element values are stored individually in the Registry and for each a Registry key is returned to Excel.
                ///    - Empty arrays will return [| proxys.empty |]. (Excel naturally returns empty array values as #VALUE!).
                let out<'a> (removeReferences: bool) (unwrapOptions: bool) (refKey: String) (proxys: Proxys) (o1D: obj[]) : obj[] =
                    if o1D |> Array.isEmpty then
                        [| proxys.empty |]
                    else
                        if removeReferences then Registry.MRegistry.removeReferencedObjects refKey
                        o1D |> Array.map (D0.Reg.out<'a> true unwrapOptions refKey proxys)

                //[<RequireQualifiedAccess>]
                //module Multi = 
                //    /// Same as Reg.out for an array of obj[] arrays.
                //    /// (references are removed only once instead of for each obj[] array).
                //    let out<'a> (unwrapOptions: bool) (refKey: String) (proxys: Proxys) (o1Ds: obj[][]) : obj[][] =
                //        Registry.MRegistry.removeReferencedObjects refKey
                //        o1Ds |> Array.map (out<'a> false unwrapOptions refKey proxys)

            [<RequireQualifiedAccess>]
            module Unbox = 
                type UnboxFn =
                    static member unbox<'A> (a1D: 'A[]) : obj[] = a1D |> Array.map box
                    static member unboxOpt<'A> (a1D: 'A[] option) : obj[] option = a1D |> Option.map (Array.map box)

                /// "Unboxes" a boxed ('a[]) into a (boxed 'a)[].
                /// In other words, casts a obj into a obj[].
                /// Returns None if the casting fails.
                let o1D (boxedArray: obj) : obj[] option = 
                    let ty = boxedArray.GetType()
                    if not ty.IsArray then
                        None
                    else
                        let res = Toolbox.Generics.invoke<UnboxFn> "unbox" [| ty.GetElementType() |] [| boxedArray |]
                        res :?> obj[]
                        |> Some
                    
                /// Convenience function, similar to o1D, but:
                ///    - Returns [| proxys.failed |] if the unboxing fails.
                ///    - Applies a function to the obj[] after unboxing.
                let apply (proxys: Proxys) (fn: obj[] -> obj[]) (boxedArray: obj) : obj[] = 
                    match boxedArray |> o1D with
                    | None -> [| proxys.failed |]
                    | Some o1d -> fn o1d

                module Opt =
                    /// "Unboxes" a boxed ('a[] option) into a (boxed 'a)[].
                    /// In other words, casts a obj into a obj[].
                    /// Returns None for None inputs OR if the casting fails.
                    let o1D (boxedOptArray: obj) : obj[] option = 
                        match Toolbox.Option.unwrap boxedOptArray with
                        | None -> None
                        | Some boxedArray -> o1D boxedArray

                    /// Convenience function, similar to o1D, but:
                    ///    - Returns [| proxys.failed |] if the input is None or if unboxing fails.
                    ///    - Applies a function to the obj[] after unboxing.
                    let apply (proxys: Proxys) (fn: obj[] -> obj[]) (boxedOptArray: obj) : obj[] = 
                        match boxedOptArray |> o1D with
                        | None -> [| proxys.failed |]
                        | Some o1d -> fn o1d


        /// Functions to output 2D arrays back to Excel.
        module D2 =
            /// Outputs arrays of (boxed) primitive (possibly optional) type elements back to Excel.
            /// Non primitive-type elements ouput as #VALUE!.
            [<RequireQualifiedAccess>]
            module Bxd =
                /// Returns sensible Excel values for 2D arrays of primitive-type elements.
                ///    - Primitive-type: Returns value directly to Excel.
                ///    - Double.NaN values will be returned as ReplaveValues.nan.
                ///    - Some (primtive-type-value) will be returned as primtive-type-value.
                ///    - None values will be returned as #VALUE!.
                ///    - Any other type: Returns ReplaveValues.object.
                ///    - Empty arrays will return a 2D singleton { proxys.empty }. (Excel naturally returns empty array values as #VALUE!).
                let out (proxys: Proxys) (o2D: obj[,]) : obj[,] =
                    if o2D |> In.D2.isEmpty then
                        In.D2.singleton<obj> proxys.empty
                    else
                        o2D |> Array2D.map (D0.Bxd.out proxys)

                /// Case of arrays of optional type elements.
                [<RequireQualifiedAccess>]
                /// Similar to Out.D2.Bxd.out but for arrays of (boxed) optional type elements.
                module Opt =
                    let out (proxys: Proxys) (o2D: (obj option)[,]) : obj[,] =
                        if o2D |> In.D2.isEmpty then
                            In.D2.singleton<obj> proxys.empty
                        else
                            o2D |> Array2D.map (D0.Bxd.Opt.out proxys)
                
                /// Case of arrays of optional or non-optional type elements.
                [<RequireQualifiedAccess>]
                module Any = 
                    /// Out.D2.Bxd.out and Out.D2.Bxd.Opt.out combined.
                    /// Works both for (boxed) optional and (boxed) non-optional elements.
                    let out (proxys: Proxys) (o2D: obj[,]) : obj[,] =
                        if o2D |> In.D2.isEmpty then
                            In.D2.singleton<obj> proxys.empty
                        else
                            o2D |> Array2D.map (D0.Bxd.Any.out proxys)

            /// Outputs 2D arrays of primitive (possibly optional) type elements back to Excel.
            /// Non primitive-type elements ouput as Proxys.object.
            [<RequireQualifiedAccess>]
            module Prm = 
                /// Returns sensible Excel values for 2D arrays depending on their element types:
                ///    - Primitive-type: Returns value directly to Excel.
                ///    - Double.NaN values will be returned as ReplaveValues.nan.
                ///    - Some (primtive-type-value) will be returned as primtive-type-value.
                ///    - None values will be returned as ReplaveValues.none.
                ///    - Any other type: Returns ReplaveValues.object.
                ///    - Empty arrays will return [| proxys.empty |]. (Excel naturally returns empty array values as #VALUE!).
                let out<'a> (proxys: Proxys) (a2D: 'a[,]) : obj[,] =
                    a2D 
                    |> Array2D.map box
                    |> Bxd.Any.out proxys

            /// Outputs 2D arrays of primitive (possibly optional) type elements back to Excel.
            /// Stores non primitive-type elements in the Registry (and output a Registry key for each individual element).
            [<RequireQualifiedAccess>]
            module Reg = 
                /// Returns sensible Excel values for 2D arrays depending on their element types:
                ///    - Primitive-type: Returns value directly to Excel.
                ///    - Double.NaN values will be returned as ReplaveValues.nan.
                ///    - Some (primtive-type-value) will be returned as primtive-type-value.
                ///    - None values will be returned as ReplaveValues.none.
                ///    - Any other type: Each element values are stored individually in the Registry and for each a Registry key is returned to Excel.
                ///    - Empty arrays will return [| proxys.empty |]. (Excel naturally returns empty array values as #VALUE!).
                let out<'a> (unwrapOptions: bool) (refKey: String) (proxys: Proxys) (o2D: obj[,]) : obj[,] =
                    if o2D |> In.D2.isEmpty then
                        In.D2.singleton<obj> proxys.empty
                    else
                        Registry.MRegistry.removeReferencedObjects refKey
                        o2D |> Array2D.map (D0.Reg.out true unwrapOptions refKey proxys)

            [<RequireQualifiedAccess>]
            module Unbox = 
                type UnboxFn =
                    static member unbox<'A> (a2D: 'A[,]) : obj[,] = a2D |> Array2D.map box

                /// "Unboxes" a boxed ('a[,]) into a (boxed 'a)[,].
                /// In other words, casts a obj into a obj[,].
                /// Returns None if the casting fails.
                let o2D (boxedArray: obj) : obj[,] option = 
                    let ty = boxedArray.GetType()
                    if not ty.IsArray then
                        None
                    else
                        let res = Toolbox.Generics.invoke<UnboxFn> "unbox" [| ty.GetElementType() |] [| boxedArray |]
                        res :?> obj[,]
                        |> Some

                /// Convenience function, similar to o1D, but:
                ///    - Returns [| proxys.failed |] if the unboxing fails.
                ///    - Applies a function to the obj[] after unboxing.
                let apply (proxys: Proxys) (fn: obj[,] -> obj[,]) (boxedArray: obj) : obj[,] = 
                    match boxedArray |> o2D with
                    | None -> In.D2.singleton<obj> proxys.failed
                    | Some o2d -> fn o2d

                module Opt =
                    /// "Unboxes" a boxed ('a[] option) into a (boxed 'a)[].
                    /// In other words, casts a obj into a obj[].
                    /// Returns None for None inputs OR if the casting fails.
                    let o2D (boxedOptArray: obj) : obj[,] option = 
                        match Toolbox.Option.unwrap boxedOptArray with
                        | None -> None
                        | Some boxedArray -> o2D boxedArray

                    /// Convenience function, similar to o1D, but:
                    ///    - Returns [| proxys.failed |] if the input is None or if unboxing fails.
                    ///    - Applies a function to the obj[] after unboxing.
                    let apply (proxys: Proxys) (fn: obj[,] -> obj[,]) (boxedOptArray: obj) : obj[,] = 
                        match boxedOptArray |> o2D with
                        | None -> In.D2.singleton<obj> proxys.failed
                        | Some o2d -> fn o2d
    
    /// Extension for textual inputs (e.g. csv files).
    module Text =

        [<RequireQualifiedAccess>]
        module Bool =
            /// Casts a string to a bool option.
            let private tryDV (defValue: bool option) (text: string) : bool option =
                match text.Trim().ToUpper() with
                | "TRUE" -> Some true
                | "FALSE" -> Some false
                | _ -> defValue

            /// Casts a string to bool or fails.
            let fail (msg: string) (text: string) : bool =  match tryDV None text with None -> failwith msg | Some x -> x

            /// Casts a string to bool with a default-value.
            let def (defValue: bool) (text: string) : bool = tryDV None text |> Option.defaultValue defValue

            // optional-type with default.
            module Opt =
                /// Casts a string to a bool option type with a default-value.
                let def (defValue: bool option) (text: string) = tryDV defValue text
        
        [<RequireQualifiedAccess>]
        module Dbl =
            /// Casts a string to a double option.
            /// "NaN" string is cast to Double.NaN. ("nan", "Nan", ... will fail).
            let private tryDV (defValue: double option) (text: string) : double option =
                match System.Double.TryParse(text) with | false, _ -> defValue | true, d -> Some d

            /// Casts a string to double or fails.
            /// "NaN" string is cast to Double.NaN. ("nan", "Nan", ... will fail).
            let fail (msg: string) (text: string) : double =  match tryDV None text with None -> failwith msg | Some x -> x

            /// Casts a string to double with a default-value.
            /// "NaN" string is cast to Double.NaN. ("nan", "Nan", ... will fail).
            let def (defValue: double) (text: string) : double = tryDV None text |> Option.defaultValue defValue

            // optional-type with default.
            module Opt =
                /// Casts a string to a bool option type with a default-value.
                /// "NaN" string is cast to Double.NaN. ("nan", "Nan", ... will fail).
                let def (defValue: double option) (text: string) = tryDV defValue text

        [<RequireQualifiedAccess>]
        module Intg =
            /// Casts a string to an int option.
            let private tryDV (defValue: int option) (text: string) : int option =
                match System.Int32.TryParse(text) with | false, _ -> defValue | true, d -> Some d

            /// Casts a string to int or fails.
            let fail (msg: string) (text: string) : int =  match tryDV None text with None -> failwith msg | Some x -> x

            /// Casts a string to int with a default-value.
            let def (defValue: int) (text: string) : int = tryDV None text |> Option.defaultValue defValue

            // optional-type with default.
            module Opt =
                /// Casts a string to an int option type with a default-value.
                let def (defValue: int option) (text: string) = tryDV defValue text

        [<RequireQualifiedAccess>]
        module Dte =
            open System.Globalization

            let private parse (text: string) = DateTime.TryParse(text)
            let private parseExact (format: string) (text: string) = DateTime.TryParseExact(text, format, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None)

            /// Casts a string to a DateTime option.
            /// dateFormat examples : "yyyyMMdd", "dd/MMM/yyyy", etc...
            let private tryDV (dateFormat: string option) (defValue: DateTime option) (text: string) : DateTime option =
                let parseFn = match dateFormat with | None -> parse | Some format -> parseExact format
                match parseFn text with 
                | false, _ -> defValue
                | true, dte -> Some dte

            /// Casts a string to DateTime or fails.
            let fail (msg: string) (dateFormat: string option) (text: string) : DateTime =  match tryDV dateFormat None text with None -> failwith msg | Some x -> x

            /// Casts a string to DateTime with a default-value.
            let def (dateFormat: string option) (defValue: DateTime) (text: string) : DateTime = tryDV dateFormat None text |> Option.defaultValue defValue

            // optional-type with default.
            module Opt =
                /// Casts a string to a DateTime option type with a default-value.
                let def (dateFormat: string option) (defValue: DateTime option) (text: string) = tryDV dateFormat defValue text

        type TagFn =
            /// Returns a default-value compatible with 'A and the typeTag.
            static member defaultValue<'A> (typeTag: string) (defValue: obj option) : 'A =
                defValue |> Option.defaultValue (Variant.labelDefVal typeTag)
                 :?> 'A

            /// Returns a default-value or None.
            static member defaultValueOpt<'A> (defValue: obj option) : 'A option =
                match defValue with
                | None -> None 
                | Some value -> 
                    match value with
                    | :? 'A as a -> Some a
                    | :? ('A option) as aopt -> aopt
                    | _ -> None

            /// Casts a string to a 'A, with a default-value for when the casting fails.
            static member def<'A> (dateFormat: string option) (defValue: obj option) (typeTag: string) (text: string) : 'A = 

                match typeTag |> Variant.ofTag with
                | BOOL -> 
                    let defval = TagFn.defaultValue<bool> typeTag defValue
                    let a0D = Bool.def defval text
                    box a0D :?> 'A
                | STRING -> 
                    box text :?> 'A
                | DOUBLE -> 
                    let defval = TagFn.defaultValue<double> typeTag defValue
                    let a0D = Dbl.def defval text
                    box a0D :?> 'A
                | DOUBLENAN -> 
                    let defval = TagFn.defaultValue<double> typeTag (Double.NaN |> box |> Some)
                    let a0D = Dbl.def defval text
                    box a0D :?> 'A
                | INT -> 
                    let defval = TagFn.defaultValue<int> typeTag defValue
                    let a0D = Intg.def defval text
                    box a0D :?> 'A
                | DATE -> 
                    let defval = TagFn.defaultValue<DateTime> typeTag defValue
                    let a0D = Dte.def dateFormat defval text
                    box a0D :?> 'A
                | _ -> failwith "TO BE IMPLEMENTED WITH OTHER VARIANT TYPES" // TODO: Complete the list

            /// Casts a string to a 'A option, with a default-value for when the casting fails.
            /// defValue is None, Some 'a or even Some (Some 'a).
            static member defOpt<'A> (dateFormat: string option) (defValue: obj option) (typeTag: string) (text: string) : 'A option = 
                match typeTag |> Variant.ofTag with
                | BOOLOPT -> 
                    let defval = TagFn.defaultValueOpt<bool> defValue
                    let a0D = Bool.Opt.def defval text
                    box a0D :?> 'A option
                | STRINGOPT -> 
                    box text :?> 'A option
                | DOUBLEOPT -> 
                    let defval = TagFn.defaultValueOpt<double> defValue
                    let a0D = Dbl.Opt.def defval text
                    box a0D :?> 'A option
                | DOUBLENANOPT -> 
                    let defval = TagFn.defaultValueOpt<double> (Double.NaN |> box |> Some)
                    let a0D = Dbl.Opt.def defval text
                    box a0D :?> 'A option
                | INTOPT -> 
                    let defval = TagFn.defaultValueOpt<double> defValue |> Option.map (int)
                    let a0D = Intg.Opt.def defval text
                    box a0D :?> 'A option
                | DATEOPT -> 
                    let defval = TagFn.defaultValueOpt<double> defValue |> Option.map (fun d -> DateTime.FromOADate(d))
                    let a0D = Dte.Opt.def dateFormat defval text
                    box a0D :?> 'A option
                | _ -> failwith "TO BE IMPLEMENTED WITH OTHER VARIANT TYPES" // TODO FIXME

        [<RequireQualifiedAccess>]
        module Tag = 
            /// Casts a string to a 'a, with a default-value for when the casting fails.
            /// 'a is determined by typeTag.
            let def (dateFormat: string option) (defValue: obj option) (typeTag: string) (text: string) : obj = 
                let gentype = typeTag |> Variant.labelType true
                let args : obj[] = [| dateFormat; defValue; typeTag; text |]
                let res = Toolbox.Generics.invoke<TagFn> "def" [| gentype |] args
                res

            // optional-type with default.
            module Opt =
                /// Casts a string to a 'a option, with a default-value for when the casting fails.
                /// 'a is determined by typeTag.
                let def (dateFormat: string option) (defValue: obj option) (typeTag: string) (text: string) : obj = 
                    let gentype = typeTag |> Variant.labelType true
                    let args : obj[] = [| dateFormat; defValue; typeTag; text |]
                    let res = Toolbox.Generics.invoke<TagFn> "defOpt" [| gentype |] args
                    res

            /// For when the type-tag is either optional, e.g. "#string", or not, e.g. "string". TODO wording
            module Any =
                /// Convenient, single function covering def and Opt.def cases.
                /// The returned (boxed) value might be either a 'a or a ('a option), depending on wether the type-tag is optional or not.
                let def (dateFormat: string option) (defValue: obj option) (typeTag: string) (text: string) : obj = 
                    let gentype = typeTag |> Variant.labelType true
                    let args : obj[] = [| dateFormat; defValue; typeTag; text |]

                    let res =
                        if typeTag |> Variant.isOptionalType then
                            Toolbox.Generics.invoke<TagFn> "defOpt" [| gentype |] args
                        else
                            Toolbox.Generics.invoke<TagFn> "def" [| gentype |] args
                    res

module Registry_XL =
    open API
    open API.In.D0
    open API.Out
    open Registry

    [<ExcelFunction(Category="Registry", Description="Removes all objects from the Registry.")>]
    let rg_free
        ([<ExcelArgument(Description= "Dependency.")>] dependency: obj)
        ([<ExcelArgument(Description= "[Collect. Default is false.]")>] collect: obj)
        : double = 

        // intermediary stage
        let collect = Bool.def false collect

        // result
        MRegistry.clear
        if collect then GC.Collect()
        let timestamp = DateTime.Now in timestamp.ToOADate()

    [<ExcelFunction(Category="Registry", Description="Removes all objects from the Registry.")>]
    let rg_collect
        ([<ExcelArgument(Description= "Dependency.")>] dependency: obj)
        : double =

        // result
        GC.Collect()
        let timestamp = DateTime.Now in timestamp.ToOADate()

    // -----------------------------------
    // -- Inspection functions
    // -----------------------------------

    [<ExcelFunction(Category="Registry", Description="Returns the Registry's count.")>]
    let rg_count () : double = MRegistry.count |> double

    [<ExcelFunction(Category="Registry", Description="Returns all Registry keys.")>]
    let rg_keys () : obj[] = MRegistry.keys |> Array.map box

    [<ExcelFunction(Category="Registry", Description="Shows the textual representation of a registry object.")>]
    let rg_show 
        ([<ExcelArgument(Description= "Reg. key.")>] regKey: string) 
        : obj =
        
        // result
        MRegistry.tryShow regKey |> outNA

    [<ExcelFunction(Category="Registry", Description="Renames a reg. object.")>]
    let rg_rname
        ([<ExcelArgument(Description= "Previous reg. key.")>] prevRegKey: string)
        ([<ExcelArgument(Description= "New reg. key.")>] newRegKey: string)
        ([<ExcelArgument(Description= "[New ref. key. Default is this cell's reference.]")>] newRefKey: obj)
        : obj =
        
        // caller cell's reference ID
        let rfid = MRegistry.refID
        let newRefKey = In.D0.Stg.def rfid newRefKey

        // result
        match MRegistry.rename newRefKey prevRegKey newRegKey with
        | None -> Proxys.def.failed
        | Some regKey -> box regKey

    [<ExcelFunction(Category="Registry", Description="Returns a registry object's type.")>]
    let rg_type 
        ([<ExcelArgument(Description= "Reg. key.")>] regKey: string)
        ([<ExcelArgument(Description= "[ToString() style. Default is false.]")>] toStringStyle: obj)
        : obj =

        // intermediary stage
        let tostringstyle = Bool.def false toStringStyle

        // result
        MRegistry.tryType regKey |> Option.map (Toolbox.Type.pPrint tostringstyle) |> outNA

    [<ExcelFunction(Category="Registry", Description="Equality of registry objects.")>]
    let rg_eq
        ([<ExcelArgument(Description= "Reg. key1.")>] regKey1: obj)
        ([<ExcelArgument(Description= "Reg. key2.")>] regKey2: obj)
        : obj  =

        match regKey1, regKey2 with
        | (:? string as regkey1), (:? string as regkey2) -> (regkey1 = regkey2) || (MRegistry.equal regkey1 regkey2)
        | _ -> regKey1 = regKey2
        |> box

    [<ExcelFunction(Category="Registry", Description="Returns a registry object's type.")>]
    let rg_unwrap 
        ([<ExcelArgument(Description= "Reg. key.")>] regKey: obj)
        ([<ExcelArgument(Description= "[ToString() style. Default is false.]")>] toStringStyle: obj)
        : obj =

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        API.Out.D0.Reg.outO false rfid Proxys.def regKey

    // -----------------------------------
    // -- Misc.
    // -----------------------------------

    [<ExcelFunction(Category="Registry", Description="Saves a registry object to disk.")>]
    let rg_writeFile
        ([<ExcelArgument(Description= "Reg. key.")>] regKey: string)
        ([<ExcelArgument(Description= "File path.")>] filePath: string)
        : obj =

        // result
        match MRegistry.ioWriteBin filePath regKey with
        | None -> Proxys.def.failed
        | Some dte -> box dte

    [<ExcelFunction(Category="Registry", Description="Reads a registry object from a file.")>]
    let rg_readFile
        ([<ExcelArgument(Description= "File path.")>] filePath: string)
        : obj =

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        MRegistry.ioLoadBin filePath rfid
        |> box

module Cast_XL =
    //open Excel
    open API
//    open type Variant
    open type Out.Proxys

    [<ExcelFunction(Category="TEST", Description="test2d")>]
    let test2D ([<ExcelArgument(Description= "Range.")>] range: obj)
        : obj  =

        let xxx : int =
            match range with
            | :? (obj[,]) as o2D -> 2
            | :? (obj[]) as o1D -> 1
            | o0D -> 0

        // result
        box xxx

    [<ExcelFunction(Category="TEST", Description="test2d")>]
    let test2D2 
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "Rowwise.")>] rowwise: bool)
        : obj[,]  =

        let to2D (xlVal: obj) : obj[,] =
            match xlVal with
            | :? (obj[,]) as o2D -> o2D
            | :? (obj[]) as o1D -> if rowwise then [| o1D |] |> array2D  else Array2D.create 1 1 (box 42)
            | o0D -> Array2D.create 1 1 o0D

        let xxx = to2D range

        // result
        xxx

    [<ExcelFunction(Category="ExcelGlue Info", Description="Returns the xl-kind string labels.")>]
    let cast_xlKinds()
        : obj[,]  =

        // intermediary stage
        let labels = [| "BOOL"; "STG"; "DBL"; "NA"; "ERR"; "MISS"; "EMPTY"; "ABS" |]

        // result        
        Array2D.init (labels.Length) 2 
            (fun i j ->
                if j = 1 then
                    labels.[i]
                else
                    String.Join(",", Kind.ofLbl labels.[i] |> Array.map (fun kind -> kind.ToString()))
            )
        |> Array2D.map box

    //[<ExcelFunction(Category="XL", Description="Cast an xl-range to DateTime[].")>]  // TODO: what is this?
    //let cast_edgeCasesTBD ()
    //    : obj[,]  =

    //    // result
    //    let (lbls, dus) = Kind.labelGuideTBD |> Array.map (fun (lbl, du) -> (box lbl, box du)) |> Array.unzip
    //    [| lbls; dus |] |> array2D

    [<ExcelFunction(Category="XL", Description="Cast an xl-range to obj[]")>]
    let cast1d_obj
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowWiseDef = In.D0.Bool.def false rowWiseDirection

        // result
        In.Cast.to1D rowWiseDef range // FIXME - should not use to1D but another In.D1.x function

    [<ExcelFunction(Category="XL", Description="Cast an xl-range to bool[].")>]
    let cast1d_bool
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "[Replacement method for non-bool elements. \"Replace\", \"Optional\" (= replace with None), \"Filter\" or \"Strict\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value. Default is FALSE.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowwise = In.D0.Bool.def false rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let none = In.D0.Stg.def "<none>" noneValue
        let empty = In.D0.Stg.def "<empty>" emptyValue
        let defVal = In.D0.Bool.def false defaultValue
        let proxys = { def with empty = empty; failed = "<failed>"; none = none }

        // result
        let o1D = In.Cast.to1D rowwise range  // FIXME - should not use to1D but another In.D1.x function
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = In.D1.Bool.filter o1D
                 a1D |> Out.D1.Prm.out<bool> proxys
        | "O" -> let a1D = In.D1.Bool.Opt.def None o1D
                 a1D |> Out.D1.Prm.out<bool option> proxys
        // strict method: either all the array's elements are of type bool, or return None (here the 1-elem array [| "failed" |])
        | "S" -> let a1D = In.D1.Bool.tryDV None o1D
                 match a1D with None -> [| proxys.failed |] | Some a1d -> a1d |> Out.D1.Prm.out<bool> proxys
        | _   -> let a1D = In.D1.Bool.def defVal o1D 
                 a1D |> Out.D1.Prm.out<bool> proxys

    [<ExcelFunction(Category="XL", Description="Cast an xl-range to string[].")>]
    let cast1d_stg
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "[Replacement method for non-string elements. \"Replace\", \"Optional\" (= replace with None), \"Filter\" or \"Strict\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value. Default is \"-foo-\".]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[]  =
        
        // intermediary stage
        let rowwise = In.D0.Bool.def false rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let none = In.D0.Stg.def "<none>" noneValue
        let empty = In.D0.Stg.def "<empty>" emptyValue
        let defVal = In.D0.Stg.def "-foo-" defaultValue
        let proxys = { def with empty = empty; failed = "<failed>"; none = none }

        // result
        // TODO replace with In.D1.OStg.filter range
        let o1D = In.Cast.to1D rowwise range  // FIXME - should not use to1D but another In.D1.x function
        // the type annotations are NOT necessary (but are used here for readability).
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = In.D1.Stg.filter o1D
                 a1D |> Out.D1.Prm.out<string> proxys
        | "O" -> let a1D = In.D1.Stg.Opt.def None o1D
                 a1D |> Out.D1.Prm.out<string option> proxys
        // strict method: either all the array's elements are of type string, or return None (here the 1-elem array [| "failed" |])
        | "S" -> let a1D = In.D1.Stg.tryDV None o1D // FIXME should use In.D1.OStg.tryDV instead
                 match a1D with None -> [| proxys.failed |] | Some a1d -> a1d |> Out.D1.Prm.out<string> proxys
        | _   -> let a1D = In.D1.Stg.def defVal o1D 
                 a1D |> Out.D1.Prm.out<string> proxys

    [<ExcelFunction(Category="XL", Description="Cast an xl-range to double[].")>]
    let cast1d_dbl
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "[Replacement method for non-double elements. \"Replace\", \"Optional\" (= replace with None), \"Filter\" or \"Strict\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value. Default is 0.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowwise = In.D0.Bool.def false rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let none = In.D0.Stg.def "<none>" noneValue
        let empty = In.D0.Stg.def "<empty>" emptyValue
        let defVal = In.D0.Dbl.def 0.0 defaultValue
        let proxys = { def with empty = empty; failed = "<failed>"; none = none }

        // result
        let o1D = In.Cast.to1D rowwise range  // FIXME - should not use to1D but another In.D1.x function
        // the type annotations are NOT necessary (but are used here for readability).
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = In.D1.Dbl.filter o1D
                 a1D |> Out.D1.Prm.out<double> proxys
        | "O" -> let a1D = In.D1.Dbl.Opt.def None o1D
                 a1D |> Out.D1.Prm.out<double option> proxys
        // strict method: either all the array's elements are of type double, or return None (here the 1-elem array [| "failed" |])
        | "S" -> let a1D = In.D1.Dbl.tryDV None o1D
                 match a1D with None -> [| proxys.failed |] | Some a1d -> a1d |> Out.D1.Prm.out<double> proxys
        | _   -> let a1D = In.D1.Dbl.def defVal o1D 
                 a1D |> Out.D1.Prm.out<double> proxys

    [<ExcelFunction(Category="XL", Description="Cast an xl-range to an array of doubles (including NaNs).")>]
    let cast1d_dblNan
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "[Replacement method for non-double elements. \"Replace\", \"Optional\" (= replace with None), \"Filter\" or \"Strict\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Kinds for which values are converted to Double.NaN. E.g. NA, ERR, TXT, !NUM... (comma separated). Default is none.]")>] xlKinds: obj)
        ([<ExcelArgument(Description= "[Default Value. Default is 0.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowwise = In.D0.Bool.def false rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let none = In.D0.Stg.def "<none>" noneValue
        let empty = In.D0.Stg.def "<empty>" emptyValue
        let xlkinds = In.D0.Stg.def "NONE" xlKinds |> Kind.ofLabel

        let defVal = In.D0.Dbl.def 0.0 defaultValue
        let proxys = { def with empty = empty; failed = "<failed>"; none = none }

        // result
        let o1D = In.Cast.to1D rowwise range  // FIXME - should not use to1D but another In.D1.x function
        // the type annotations are NOT necessary (but are used here for readability).
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = In.D1.Nan.filter xlkinds o1D
                 a1D |> Out.D1.Prm.out<double> proxys
        | "O" -> let a1D = In.D1.Nan.Opt.def xlkinds None o1D 
                 a1D |> Out.D1.Prm.out<double option> proxys
        | "S" -> let a1D = In.D1.Nan.tryDV xlkinds None o1D
                 match a1D with None -> [| proxys.failed |] | Some a1d -> a1d |> Out.D1.Prm.out<double> proxys
        | _   -> let a1D = In.D1.Nan.def xlkinds defVal o1D 
                 a1D |> Out.D1.Prm.out<double> proxys

    [<ExcelFunction(Category="XL", Description="Cast an xl-range to int[].")>]
    let cast1d_int
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "[Replacement method for non-integer elements. \"Replace\", \"Optional\" (= replace with None), \"Filter\" or \"Strict\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value. Default is 0.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowwise = In.D0.Bool.def false rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let none = In.D0.Stg.def "<none>" noneValue
        let empty = In.D0.Stg.def "<empty>" emptyValue
        let defVal = In.D0.Intg.def 0 defaultValue
        let proxys = { def with empty = empty; failed = "<failed>"; none = none }

        // result
        let o1D = In.Cast.to1D rowwise range  // FIXME - should not use to1D but another In.D1.x function
        // the type annotations are NOT necessary (but are used here for readability).
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = In.D1.Intg.filter o1D
                 a1D |> Out.D1.Prm.out<int> proxys
        | "O" -> let a1D = In.D1.Intg.Opt.def None o1D
                 a1D |> Out.D1.Prm.out<int option> proxys
        // strict method: either all the array's elements are of type int, or return None (here the 1-elem array [| "failed" |])
        | "S" -> let a1D = In.D1.Intg.tryDV None o1D
                 match a1D with None -> [| proxys.failed |] | Some a1d -> a1d |> Out.D1.Prm.out<int> proxys
        | _   -> let a1D = In.D1.Intg.def defVal o1D
                 a1D |> Out.D1.Prm.out<int> proxys
        
    [<ExcelFunction(Category="XL", Description="Cast an xl-range to DateTime[].")>]
    let cast1d_dte
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "[Replacement method for non-date elements. \"Replace\", \"Optional\" (= replace with None), \"Filter\" or \"Strict\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value. Default is 1-Jan-2000.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowwise = In.D0.Bool.def false rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let none = In.D0.Stg.def "<none>" noneValue
        let empty = In.D0.Stg.def "<empty>" emptyValue
        let defVal = In.D0.Dte.def (DateTime(2000,1,1)) defaultValue
        let proxys = { def with empty = empty; failed = "<failed>"; none = none }

        // result
        let o1D = In.Cast.to1D rowwise range  // FIXME - should not use to1D but another In.D1.x function
        // the type annotations are NOT necessary (but are used here for readability).
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = In.D1.Dte.filter o1D
                 a1D |> Out.D1.Prm.out<DateTime> proxys
        | "O" -> let a1D = In.D1.Dte.Opt.def None o1D
                 a1D |> Out.D1.Prm.out<DateTime option> proxys
        // strict method: either all the array's elements are of type int, or return None (here the 1-elem array [| "failed" |])
        | "S" -> let a1D = In.D1.Dte.tryDV None o1D
                 match a1D with None -> [| proxys.failed |] | Some a1d -> a1d |> Out.D1.Prm.out<DateTime> proxys
        | _   -> let a1D = In.D1.Dte.def defVal o1D 
                 a1D |> Out.D1.Prm.out<DateTime> proxys

    [<ExcelFunction(Category="XL", Description="Cast an xl-value to a generic type.")>]
    let cast_gen
        ([<ExcelArgument(Description= "xl-value.")>] xlValue: obj)
        ([<ExcelArgument(Description= "Type tag: bool, date, double, doubleNaN, string or obj. Add \'#'\ prefix for optional type: #bool, #date, #double, #doubleNaN, #string or #obj")>] typeTag: string)
        //([<ExcelArgument(Description= "[Default Value (only for non-optional types). Must be of the appropriate type. Default \"<default>\" (which will fail for non-string types).]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[Default Value (only for non-optional types, optional types default to None). Must be of the appropriate type, else it will fail.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None indicator. Default is \"<none>\".]")>] noneIndicator: obj)
        ([<ExcelArgument(Description= "[Kinds for which values are converted to Double.NaN. E.g. NA, ERR, TXT, !NUM... (comma separated). Default is none.]")>] xlKinds: obj)
        : obj  =

        // intermediary stage
        let none = In.D0.Stg.def "<none>" noneIndicator
        let proxys = { def with none = none }
        let defVal = In.D0.Missing.Obj.tryO defaultValue
        let xlkinds = In.D0.Stg.def "NONE" xlKinds |> Kind.ofLabel

        // for demo purpose only: takes an Excel cell input,
        // converts it into a (boxed) typed value, then outputs it back to Excel.
        //let res = Out.D0.Gen.defAllCasesObj proxys defVal typeLabel xlValue
        let xa0D = In.D0.Tag.Any.def xlkinds defVal typeTag xlValue
        xa0D |> Out.D0.Prm.out proxys

    [<ExcelFunction(Category="XL", Description="Cast a 1D-slice of an xl-range to a generic type 1D array.")>] // FIXME change wording
    let cast1d_gen
        ([<ExcelArgument(Description= "1D xl-range.")>] range: obj)
        ([<ExcelArgument(Description= "Type tag: bool, date, double, doubleNaN, string or obj. Add \'#'\ prefix for optional type: #bool, #date, #double, #doubleNaN, #string or #obj")>] typeTag: string)
        ([<ExcelArgument(Description= "[Replacement method for non-date elements. \"Replace\", \"Filter\" or \"Strict\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value (only for non-optional types, optional types default to None). Must be of the appropriate type, else it will fail.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "[Kinds for which values are converted to Double.NaN. E.g. NA, ERR, TXT, !NUM... (comma separated). Default is none.]")>] xlKinds: obj)
        ([<ExcelArgument(Description= "[Row-wise slice direction when input is a fat, 2D, range. True or false or none. Default is none.]")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowwise = In.D0.Bool.Opt.def None rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let none = In.D0.Stg.def "<none>" noneValue
        let empty = In.D0.Stg.def "<empty>" emptyValue
        let proxys = { def with empty = empty; failed = "<failed>"; none = none }
        let defVal = In.D0.Missing.Obj.tryO defaultValue
        let xlkinds = In.D0.Stg.def "NONE" xlKinds |> Kind.ofLabel
        
        // for demo purpose only: takes an Excel range input,
        // converts it into a (boxed) typed array, then outputs it back to Excel.
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let xa1D = In.D1.Tag.filter xlkinds rowwise typeTag range
                 xa1D |> (Out.D1.Unbox.apply proxys (Out.D1.Prm.out proxys))

        // strict method: either all the array's elements are of type int, or return None (here the 1-elem array [| proxys.failed |])
        | "S" -> let xa1D = In.D1.Tag.tryDV xlkinds rowwise None typeTag range |> snd
                 xa1D |> (Out.D1.Unbox.Opt.apply proxys (Out.D1.Prm.out proxys))

        | _ -> let xa1D = In.D1.Tag.Any.def xlkinds rowwise defVal typeTag range
               xa1D |> (Out.D1.Unbox.apply proxys (Out.D1.Prm.out proxys))

    [<ExcelFunction(Category="XL", Description="Cast a 2D xl-range to a generic type 2D array.")>]
    let cast2d_gen
        ([<ExcelArgument(Description= "2D xl-range.")>] range: obj)
        ([<ExcelArgument(Description= "Type tag: bool, date, double, doubleNaN, string or obj. Add \'#'\ prefix for optional type: #bool, #date, #double, #doubleNaN, #string or #obj")>] typeTag: string)
        ([<ExcelArgument(Description= "[Replacement method for non-date elements. \"Replace\", \"Filter\" or \"Strict\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value (only for non-optional types, optional types default to None). Must be of the appropriate type, else it will fail.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "[Row wise direction for filtering. Default is true.]")>] rowWiseDirection: obj)
        ([<ExcelArgument(Description= "[Kinds for which values are converted to Double.NaN. E.g. NA, ERR, TXT, !NUM... (comma separated). Default is none.]")>] xlKinds: obj)
        : obj[,]  =

        // intermediary stage
        let rowWise = In.D0.Bool.def true rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let none = In.D0.Stg.def "<none>" noneValue
        let empty = In.D0.Stg.def "<empty>" emptyValue
        let proxys = { def with empty = empty; failed = "<failed>"; none = none }
        let defVal = In.D0.Missing.Obj.tryO defaultValue
        let xlkinds = In.D0.Stg.def "NONE" xlKinds |> Kind.ofLabel
        
        // for demo purpose only: takes an Excel range input,
        // converts it into a (boxed) typed 2D array, then outputs it back to Excel.
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let xa2D = In.D2.Tag.filter xlkinds rowWise typeTag range
                 xa2D |> (Out.D2.Unbox.apply proxys (Out.D2.Prm.out proxys))

        // strict method: either all the array's elements are of type int, or return None (here the 1-elem array [| proxys.failed |])
        | "S" -> let xa2D = In.D2.Tag.tryDV xlkinds None typeTag range
                 xa2D |> (Out.D2.Unbox.Opt.apply proxys (Out.D2.Prm.out proxys))

        | _ -> let xa2D = In.D2.Tag.Any.def xlkinds defVal typeTag range
               xa2D |> (Out.D2.Unbox.apply proxys (Out.D2.Prm.out proxys))

module FSI =
    // https://fsharp.github.io/FSharp.Compiler.Service/interactive.html
    open System.Text
    open FSharp.Compiler.Interactive.Shell

    // -----------------------------
    // -- Initializations
    // -----------------------------

    // Initializes output and input streams
    let sbOut = new StringBuilder()
    let sbErr = new StringBuilder()
    let inStream = new StringReader("")
    let outStream = new StringWriter(sbOut)
    let errStream = new StringWriter(sbErr)

    // Builds command line arguments & start FSI session
    let argv = [| "C:\\fsi.exe" |]
    let allArgs = Array.append argv [| "--noninteractive" |]
    let config = FsiEvaluationSession.GetDefaultConfiguration()

    let mutable session = FsiEvaluationSession.Create(config, allArgs, inStream, outStream, errStream)
    session.EvalInteraction("open System")
    
    [<RequireQualifiedAccess>]
    module Session =
        let set (newSession: FsiEvaluationSession) = 
            session <- newSession
            session.EvalInteraction("open System")

        let reset = 
            session <- FsiEvaluationSession.Create(config, allArgs, inStream, outStream, errStream)
            session.EvalInteraction("open System")

module Fun =
    // https://fsharp.github.io/FSharp.Compiler.Service/interactive.html
    open System.Text
    open FSharp.Compiler.Interactive.Shell

    // -----------------------------
    // -- Initializations
    // -----------------------------

    // Initializes output and input streams
    let sbOut = new StringBuilder()
    let sbErr = new StringBuilder()
    let inStream = new StringReader("")
    let outStream = new StringWriter(sbOut)
    let errStream = new StringWriter(sbErr)

    // Builds command line arguments & start FSI session
    let argv = [| "C:\\fsi.exe" |]
    let allArgs = Array.append argv [| "--noninteractive" |]
    
    let fsiConfig = FsiEvaluationSession.GetDefaultConfiguration()
    //let fsiSession = FsiEvaluationSession.Create(fsiConfig, allArgs, inStream, outStream, errStream)

    let diagnose (prefix: string) (diagnostics: FSharp.Compiler.SourceCodeServices.FSharpDiagnostic[]) : string[] =
        [| for w in diagnostics -> sprintf "[[%d,%d]] %s" w.StartLineAlternate w.StartColumn w.Message |]
        |> Array.map (fun s -> prefix + s)

    // TODO: get binary's location and outputs it to user
    //// let loc = System.Reflection.Assembly.GetExecutingAssembly().Location
    ////fsiSession.EvalInteraction(sprintf "#r @\"%s\"" loc)

    /// Evaluates an expression and returns its type and value, as well as warnings.
    /// E.g. expr = "\ x y -> x + y" or expr = "fun x y -> x + y".
    let evaluate (session: FsiEvaluationSession ) (expression: string) : ((Type*obj) option)*string[] =
        let fullexpr = let e = expression.Trim() in (if e.Substring(0,1) = @"\" then "fun " + e.Substring(1) else e)
        let result, diagnostics = session.EvalExpressionNonThrowing(fullexpr)
        let warnings = diagnose "" diagnostics

        match result with
        | Choice1Of2 (Some fsivalue) ->
            let res = fsivalue.ReflectionValue 
            let ty = res.GetType()
            (ty, res) |> Some, warnings
        | _ -> None, warnings

    /// Convenience functions for building Excel UDFs.
    [<RequireQualifiedAccess>]
    module Reg =
        /// Returns either the result of an evaluation, or its warnings.
        /// Outputs can be registered. 
        let evalExpression (registry: Registry) (failIndic: obj) (refKey: String) (register: bool) (outputWarnings: bool) (expression: string) (session: byref<FsiEvaluationSession>) =
            match outputWarnings, evaluate session expression with
            // warnings requested, boxeds as string[]
            | true, (_, warnings) -> warnings |> registry.registerBxd refKey
            // result requested
            // failure
            | false, (None, _) -> failIndic
            // success
            | false, (Some (_, value), _) ->
                if register then
                    value |> registry.registerBxd refKey
                else
                    value

        /// Returns a (boxed) fsi function given:
        ///    - either an (string) expression and (possibly) a fsi session R-object.
        ///    - or a FSharpFunc R-object.
        /// If both are provided, only the FSharpFunc R-object is active and the expression is ignored.
        let fsiFunction (registry: Registry) (failIndic: obj) (refKey: String) (rgFsiSession: obj) (expression: string option) (rgFSharpFunc: string option) : obj option =
            match rgFSharpFunc with
            | None ->
                match expression with
                | Some expr -> 
                    let mutable fsisession = registry.tryExtract<FsiEvaluationSession> rgFsiSession |> Option.defaultValue FSI.session
                    evalExpression registry failIndic refKey false false expr &fsisession |> Some
                | None -> None
            | Some regKey -> registry.tryExtractO regKey


    // -----------------------------
    // -- Reflection functions
    // -----------------------------

    /// Returns true if ofun is a FSharpFunc object.
    let isFunction (ofun: obj) : bool =
        let TYPE_NAME = "FSharpFunc`"
        (ofun.GetType().BaseType.Name).Substring(0,TYPE_NAME.Length) = TYPE_NAME

    /// Returns the arity of a FsharpFunc object.
    let arity (ofun: obj) : int option =
        if not (isFunction ofun) then
            None
        else
            let gentys = ofun.GetType().BaseType.GetGenericArguments()
            gentys.Length - 1 // the last argument is the type of the function result (not part of the arity).
            |> Some

    /// Returns the input arguments' types of a FsharpFunc object.
    let inputTypes (ofun: obj) : Type[] option =
        if not (isFunction ofun) then
            None
        else
            let argTypes = ofun.GetType().BaseType.GetGenericArguments() 
            if argTypes.Length = 1 then
                Some [||]
            else
                argTypes 
                |> Array.take (argTypes.Length - 1)
                |> Some

    /// Returns the input arguments' types of a FsharpFunc object.
    /// argTypes should not include the output argument's type.
    let compatibleArgTypes (ofun: obj) (argTypes: Type[]) : bool =
        match inputTypes ofun with 
        | None -> false
        | Some allInputTypes ->
            (argTypes.Length <= allInputTypes.Length) && (allInputTypes |> Array.take argTypes.Length = argTypes)

    /// Returns the output type of a FsharpFunc object.
    let outputType (ofun: obj) : Type option =
        if not (isFunction ofun) then
            None
        else
            ofun.GetType().BaseType.GetGenericArguments() 
            |> Array.last
            |> Some

    // F# functions type naming convention : FSharpFunc`(Arity + 1). (+1 for result type). E.g. FSharpFunc`2 for 'a -> 'b, FSharpFunc`3 for 'a -> 'b -> 'c... 
    // Arity is (temporarily?) limited to 5 inputs, as F# treats higher arity functions as nested functions.
    //    - 1 to 5 inputs => no inner function // FSharpFunc`2 to FSharpFunc`6
    //    - 6 to 10 inputs => 1 inner function
    //    - 11 to 15 inputs => 2 inner functions
    //    ...

    let MAX_ARITY_FILTER = 10
    type Filter =
        static member filter1<'A1> (f: 'A1 -> bool) (x1: 'A1) : bool = f x1
        static member filter2<'A1,'A2> (f: 'A1 -> 'A2 -> bool) (x1: 'A1) (x2: 'A2) : bool = f x1 x2
        static member filter3<'A1,'A2,'A3> (f: 'A1 -> 'A2 -> 'A3 -> bool) (x1: 'A1) (x2: 'A2) (x3: 'A3) : bool = f x1 x2 x3
        static member filter4<'A1,'A2,'A3,'A4> (f: 'A1 -> 'A2 -> 'A3 -> 'A4 -> bool) (x1: 'A1) (x2: 'A2) (x3: 'A3) (x4: 'A4) : bool = f x1 x2 x3 x4
        static member filter5<'A1,'A2,'A3,'A4,'A5> (f: 'A1 -> 'A2 -> 'A3 -> 'A4 -> 'A5 -> bool) (x1: 'A1) (x2: 'A2) (x3: 'A3) (x4: 'A4) (x5: 'A5) : bool = f x1 x2 x3 x4 x5
        static member filter6<'A1,'A2,'A3,'A4,'A5,'A6> (f: 'A1 -> 'A2 -> 'A3 -> 'A4 -> 'A5 -> 'A6 -> bool) (x1: 'A1) (x2: 'A2) (x3: 'A3) (x4: 'A4) (x5: 'A5) (x6: 'A6) : bool = f x1 x2 x3 x4 x5 x6
        static member filter7<'A1,'A2,'A3,'A4,'A5,'A6,'A7> (f: 'A1 -> 'A2 -> 'A3 -> 'A4 -> 'A5 -> 'A6 -> 'A7 -> bool) (x1: 'A1) (x2: 'A2) (x3: 'A3) (x4: 'A4) (x5: 'A5) (x6: 'A6) (x7: 'A7) : bool = f x1 x2 x3 x4 x5 x6 x7
        static member filter8<'A1,'A2,'A3,'A4,'A5,'A6,'A7,'A8> (f: 'A1 -> 'A2 -> 'A3 -> 'A4 -> 'A5 -> 'A6 -> 'A7 -> 'A8 -> bool) (x1: 'A1) (x2: 'A2) (x3: 'A3) (x4: 'A4) (x5: 'A5) (x6: 'A6) (x7: 'A7) (x8: 'A8) : bool = f x1 x2 x3 x4 x5 x6 x7 x8
        static member filter9<'A1,'A2,'A3,'A4,'A5,'A6,'A7,'A8,'A9> (f: 'A1 -> 'A2 -> 'A3 -> 'A4 -> 'A5 -> 'A6 -> 'A7 -> 'A8 -> 'A9 -> bool) (x1: 'A1) (x2: 'A2) (x3: 'A3) (x4: 'A4) (x5: 'A5) (x6: 'A6) (x7: 'A7) (x8: 'A8) (x9: 'A9) : bool = f x1 x2 x3 x4 x5 x6 x7 x8 x9
        static member filter10<'A1,'A2,'A3,'A4,'A5,'A6,'A7,'A8,'A9,'A10> (f: 'A1 -> 'A2 -> 'A3 -> 'A4 -> 'A5 -> 'A6 -> 'A7 -> 'A8 -> 'A9 -> 'A10 -> bool) (x1: 'A1) (x2: 'A2) (x3: 'A3) (x4: 'A4) (x5: 'A5) (x6: 'A6) (x7: 'A7) (x8: 'A8) (x9: 'A9) (x10: 'A10) : bool = f x1 x2 x3 x4 x5 x6 x7 x8 x9 x10

    let filter (ofun: obj) (args: obj[]) : bool =
        match arity ofun with
        | None -> false
        | Some rank when rank > MAX_ARITY_FILTER -> false
        | Some rank -> 
            let argTypes = args |> Array.map (fun arg -> arg.GetType())
            let funtys = ofun.GetType().BaseType.GetGenericArguments()
            if (argTypes.Length > rank) || (not (compatibleArgTypes ofun argTypes)) then
                false
            else
                let methodNm = sprintf "filter%d" rank
                let res = Toolbox.Generics.invoke<Filter> methodNm funtys (Array.append [| ofun |] args)
                res :?> bool

    /// Same as Fun.filter but for multi arguments.
    /// Assumes that each argument set is of the right type - otherwise fails.
    let filterMulti (ofun: obj)  (argTypes: Type[]) (argss: obj[][]) : bool[] =
        match arity ofun with
        | None -> Array.create argss.Length false
        | Some rank when rank > MAX_ARITY_FILTER -> Array.create argss.Length false
        | Some rank -> 
            //let funtys = ofun.GetType().BaseType.GetGenericArguments()
            if (argTypes.Length > rank) || (not (compatibleArgTypes ofun argTypes)) then
                Array.create argss.Length false
            else
                let methodNm = sprintf "filter%d" rank
                let filterFun (args: obj[]) = Toolbox.Generics.invoke<Filter> methodNm argTypes (Array.append [| ofun |] args)
                let res = argss |> Array.map (filterFun >> (fun o -> o :?> bool))
                res

    let MAX_ARITY_APPLY = 5
    type Apply =
        static member Apply1<'A1,'B> (f: 'A1 -> 'B) (x1: 'A1) : 'B = f x1
        static member Apply2<'A1,'A2,'B> (f: 'A1 -> 'A2 -> 'B) (x1: 'A1) (x2: 'A2) : 'B = f x1 x2
        static member Apply3<'A1,'A2,'A3,'B> (f: 'A1 -> 'A2 -> 'A3 -> 'B) (x1: 'A1) (x2: 'A2) (x3: 'A3) : 'B = f x1 x2 x3
        static member Apply4<'A1,'A2,'A3,'A4,'B> (f: 'A1 -> 'A2 -> 'A3 -> 'A4 -> 'B) (x1: 'A1) (x2: 'A2) (x3: 'A3) (x4: 'A4) : 'B = f x1 x2 x3 x4
        static member Apply5<'A1,'A2,'A3,'A4,'A5,'B> (f: 'A1 -> 'A2 -> 'A3 -> 'A4 -> 'A5 -> 'B) (x1: 'A1) (x2: 'A2) (x3: 'A3) (x4: 'A4) (x5: 'A5) : 'B = f x1 x2 x3 x4 x5

    let apply (ofun: obj) (args: obj[]) : obj option =
        match arity ofun with
        | None -> None
        | Some rank when rank > MAX_ARITY_APPLY -> None
        | Some rank -> 
            let argTypes = args |> Array.map (fun arg -> arg.GetType())
            let funtys = ofun.GetType().BaseType.GetGenericArguments()
            if (argTypes.Length > rank) || (not (compatibleArgTypes ofun argTypes)) then
                None
            else
                let methodNm = sprintf "Apply%d" rank
                let res = Toolbox.Generics.invoke<Apply> methodNm funtys (Array.append [| ofun |] args)
                Some res

    /// Same as Fun.apply but for multi arguments.
    /// Assumes that each argument set is of the right type - otherwise fails.
    let applyMulti (ofun: obj) (argTypes: Type[]) (argss: obj[][]) : obj[] option =
        match arity ofun with
        | None -> None
        | Some rank when rank > 5 -> None
        | Some rank -> 
            let funtys = ofun.GetType().BaseType.GetGenericArguments()
            if (argTypes.Length > rank) || (not (compatibleArgTypes ofun argTypes)) then
                None
            else
                let methodNm = sprintf "Apply%d" rank
                let applyFun (args: obj[]) = Toolbox.Generics.invoke<Apply> methodNm funtys (Array.append [| ofun |] args)
                let res = argss |> Array.map applyFun
                Some res

    let applyOLD (ofun: obj) (args: obj[]) : obj option = // (Type*obj) option =
        match arity ofun with
        | None -> None
        | Some rank when rank > 5 -> None
        | Some rank -> 
            let argtys = args |> Array.map (fun arg -> arg.GetType())
            let funtys = ofun.GetType().BaseType.GetGenericArguments()
            if (argtys.Length > rank) || (argtys <> (funtys |> Array.take argtys.Length)) then
                None
            else
                let methodNm = sprintf "Apply%d" rank
                let res = Toolbox.Generics.invoke<Apply> methodNm funtys (Array.append [| ofun |] args)
                Some res

module Fun_XL =
    open Registry
    open type API.Out.Proxys
    open FSharp.Compiler.Interactive.Shell

    [<ExcelFunction(Category="Fun", Description="Resets the internal FSI session.")>]
    let fn_sessionReset
        ([<ExcelArgument(Description= "Dependency.")>] dependency: obj)
        : obj =

        //result
        FSI.Session.reset
        box DateTime.Now

    [<ExcelFunction(Category="Fun", Description="Returns an ad hoc FSI session R-object.")>]
    let fn_sessionAdhoc
        ([<ExcelArgument(Description= "Dependency.")>] dependency: obj)
        : obj =

        // caller cell's reference ID
        let rfid = MRegistry.refID
 
        //result
        try
           let fsisession = FsiEvaluationSession.Create(Fun.fsiConfig, Fun.allArgs, Fun.inStream, Fun.outStream, Fun.errStream)
           fsisession.EvalInteraction("open System")
           fsisession |> MRegistry.registerBxd rfid

        with
        | _ -> def.failed

    type private State = | Success of int*(string[]) | Failure of int*(string[]) with
        static member initial = Success (0, [||])

    [<ExcelFunction(Category="Fun", Description="Returns a FSI session R-object.")>]
    let fn_addDirectives
        ([<ExcelArgument(Description= "FSI session R-object, or TRUE for internal session.")>] rgFsiSession: string)
        ([<ExcelArgument(Description= "Directives array. E.g. #r, #I, #load. Default \"#r\".]")>] directives: obj)
        ([<ExcelArgument(Description= "Path or code array. E.g. C:\\...\\ExcelDna.Interop.1.2.3\\SomeDna.dll.")>] pathOrCodes: obj)
        ([<ExcelArgument(Description= "[Output warnings. Default is false.]")>] outputWarnings: obj)
        : obj =

        // intermediary stage
        let useInternalSession = API.In.D0.Bool.def false rgFsiSession
        let directive1D = API.In.D1.OStg.tryDV None directives
        let path1D = API.In.D1.OStg.tryDV None pathOrCodes
        let outputWs = API.In.D0.Bool.def false outputWarnings

        let reference (directive: string, path: string) =
            if directive.Substring(0,1) = "#" then
                let path = "@\"" + path.Trim() + "\""
                directive.Trim() + " " + path
            else
                path.Trim()

        let chain (fsisession: FsiEvaluationSession) (state: State) (reference: string) : State = 
            match state with 
            | Failure _ -> state
            | Success (acc, prevWarnings) ->                
                let result, diagnostics = fsisession.EvalInteractionNonThrowing(reference)
                let newWarnings = 
                    Array.append
                        [| sprintf "== STAGE %d ==" acc |]
                        // [| for w in diagnostics ->  sprintf "   [[%d,%d]] %s" w.StartLineAlternate w.StartColumn w.Message |]
                        (Fun.diagnose "   " diagnostics)
                let warnings = Array.append newWarnings prevWarnings
                match result with
                | Choice1Of2 _ -> Success (acc + 1, warnings)
                | _ -> Failure (acc + 1, warnings)

        // caller cell's reference ID
        let rfid = MRegistry.refID
 
        //result
        match directive1D, path1D, MRegistry.tryExtract<FsiEvaluationSession> rgFsiSession with
        | Some d1D, Some p1D, Some fsisession ->
            let references = Array.zip d1D p1D |> Array.map reference
            let addReferences = references |> Array.fold (chain fsisession) State.initial
            match outputWs, addReferences with
            | true, Failure (_, warnings) -> warnings |> MRegistry.registerBxd rfid
            | true, Success (_, warnings) -> warnings |> MRegistry.registerBxd rfid
            | false, Failure _ -> def.failed
            | false, Success _ -> fsisession |> MRegistry.registerBxd rfid

        | _ -> def.failed

    [<ExcelFunction(Category="Fn", Description="Creates a FSharpFunc R-object.")>]
    let fn_expr
        ([<ExcelArgument(Description= "Multiline expression. E.g. \"\\ x y -> x + y\"")>] expression: obj)
        ([<ExcelArgument(Description= "[Creates a R-object. Default is true.]")>] registerObject: obj)
        ([<ExcelArgument(Description= "[Output warnings. Default is false.]")>] outputWarnings: obj)
        ([<ExcelArgument(Description= "Ad hoc FSI session R-object.")>] rgFsiSession: obj)
        : obj =

        // intermediary stage
        let register = API.In.D0.Bool.def true registerObject
        let outputWs = API.In.D0.Bool.def false outputWarnings
        let expression = API.In.D1.OStg.tryDV None expression |> Option.map (fun exprs -> String.Join("\n", exprs))

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match expression with
        | None -> def.failed
        | Some expr ->
            let mutable fsisession = MRegistry.tryExtract<FsiEvaluationSession> rgFsiSession |> Option.defaultValue FSI.session
            Fun.Reg.evalExpression MRegistry def.failed rfid register outputWs expr &fsisession

    [<ExcelFunction(Category="Fn", Description="Creates a function object.")>]
    let fn_apply
        ([<ExcelArgument(Description= "Multiline expression. E.g. \"\\ x -> 2.0 * x\"")>] expression: obj)        
        ([<ExcelArgument(Description= "Argument 1.")>] argument1: obj)
        ([<ExcelArgument(Description= "Argument 2.")>] argument2: obj)
        ([<ExcelArgument(Description= "Argument 3.")>] argument3: obj)
        ([<ExcelArgument(Description= "Argument 4.")>] argument4: obj)
        ([<ExcelArgument(Description= "Argument 5.")>] argument5: obj)
        ([<ExcelArgument(Description= "[Creates a R-object. Default is true.]")>] registerObject: obj)
        ([<ExcelArgument(Description= "[FSharpFunc R-object. Disable expression if present. Default is none.]")>] rgFSharpFunc: obj)
        ([<ExcelArgument(Description= "[Ad hoc FSI session R-object. Inactive if FSharpFunc object provided. Default is internal session.]")>] rgFsiSession: obj)
        : obj =

        // intermediary stage
        let arg1 = API.In.D0.Absent.Obj.tryO argument1
        let arg2 = if arg1 |> Option.isNone then None else API.In.D0.Absent.Obj.tryO argument2
        let arg3 = if arg2 |> Option.isNone then None else API.In.D0.Absent.Obj.tryO argument3
        let arg4 = if arg3 |> Option.isNone then None else API.In.D0.Absent.Obj.tryO argument4
        let arg5 = if arg4 |> Option.isNone then None else API.In.D0.Absent.Obj.tryO argument5
        let args = [| arg1; arg2; arg3; arg4; arg5 |] |> Array.choose id
        let register = API.In.D0.Bool.def true registerObject

        // intermediary stage
        let expression = API.In.D1.OStg.tryDV None expression |> Option.map (fun exprs -> String.Join("\n", exprs))
        let rgFSharpFunc = API.In.D0.Stg.Opt.def None rgFSharpFunc

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match Fun.Reg.fsiFunction MRegistry def.failed rfid rgFsiSession expression rgFSharpFunc with
        | None -> def.failed
        | Some ofun ->
            match Fun.apply ofun args with
            | None -> def.failed
            | Some res ->
                if register then
                    res |> MRegistry.registerBxd rfid
                else
                    res
            
    [<ExcelFunction(Category="Fn", Description="Creates a function object.")>]
    let fn_applyOLD
        ([<ExcelArgument(Description= "FSharpFunc R-object.")>] rgFSharpFunc: string)
        ([<ExcelArgument(Description= "Argument 1.")>] argument1: obj)
        ([<ExcelArgument(Description= "Argument 2.")>] argument2: obj)
        ([<ExcelArgument(Description= "Argument 3.")>] argument3: obj)
        ([<ExcelArgument(Description= "Argument 4.")>] argument4: obj)
        ([<ExcelArgument(Description= "Argument 5.")>] argument5: obj)
        ([<ExcelArgument(Description= "[Creates a R-object. Default is true.]")>] registerObject: obj)
        : obj =

        // intermediary stage
        let arg1 = API.In.D0.Absent.Obj.tryO argument1
        let arg2 = if arg1 |> Option.isNone then None else API.In.D0.Absent.Obj.tryO argument2
        let arg3 = if arg2 |> Option.isNone then None else API.In.D0.Absent.Obj.tryO argument3
        let arg4 = if arg3 |> Option.isNone then None else API.In.D0.Absent.Obj.tryO argument4
        let arg5 = if arg4 |> Option.isNone then None else API.In.D0.Absent.Obj.tryO argument5
        let args = [| arg1; arg2; arg3; arg4; arg5 |] |> Array.choose id
        let register = API.In.D0.Bool.def true registerObject

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match MRegistry.tryExtractO rgFSharpFunc with
        | None -> def.failed
        | Some ofun ->
            match Fun.apply ofun args with
            | None -> def.failed
            | Some res ->
                if register then
                    res |> MRegistry.registerBxd rfid
                else
                    res

module A1D = 
    open Registry
    open Toolbox.Array
    open Toolbox.Generics
    open API.Out

    // -----------------------------------
    // -- Reflection functions
    // -----------------------------------
    type GenFn =
        static member out<'A> (a1D: 'A[]) (unwrapOptions: bool) (refKey: String) (proxys: Proxys) : obj[] = 
            a1D |> Array.map box |> (API.Out.D1.Reg.out<'A> true unwrapOptions refKey proxys)
            
        static member outObj<'A> (o1D: obj[]) (unwrapOptions: bool) (refKey: String) (proxys: Proxys) : obj[] = 
            let a1D = o1D |> Array.map (fun o -> o :?> 'A)
            a1D |> Array.map box |> (API.Out.D1.Reg.out<'A> true unwrapOptions refKey proxys)

        static member outObjWithRefControl<'A> (o1D: obj[]) (removeReferences: bool) (unwrapOptions: bool) (refKey: String) (proxys: Proxys) : obj[] = 
            let a1D = o1D |> Array.map (fun o -> o :?> 'A)
            a1D |> Array.map box |> (API.Out.D1.Reg.out<'A> removeReferences unwrapOptions refKey proxys)

        static member count<'A> (a1D: 'A[]) : int = a1D |> Array.length

        static member tryElem<'A> (a1D: 'A[]) (unwrapOptions: bool) (refKey: String) (proxys: Proxys) (index: int) : obj = 
            match a1D  |> Array.tryItem index with
            | None -> proxys.failed
            | Some elem -> elem |> API.Out.D0.Reg.out<'A> false unwrapOptions refKey proxys

        static member sub<'A> (a1D: 'A[]) (startIndex: int option) (count: int option) : 'A[] =
            a1D |> sub startIndex count

        static member append2<'A> (a1D1: 'A[]) (a1D2: 'A[]) : 'A[] =
            Array.append a1D1 a1D2

        static member append3<'A> (a1D1: 'A[]) (a1D2: 'A[]) (a1D3: 'A[]) : 'A[] =
            Array.append (Array.append a1D1 a1D2) a1D3

        static member sort<'A when 'A: comparison> (a1D: 'A[]) (descending: bool) : 'A[] = 
            if descending then a1D |> Array.sortDescending else a1D |> Array.sort

        static member map1<'A1,'B> (f: 'A1 -> 'B) (a1D: 'A1[]) : 'B[] = a1D |> Array.map f
        static member map2<'A1,'A2,'B> (f: 'A1 -> 'A2 -> 'B) (a1D1: 'A1[]) (a1D2: 'A2[]) : 'B[] = Array.map2 f a1D1 a1D2

        static member filter<'A1> (f: 'A1 -> bool) (a1D: 'A1[]) : 'A1[] = a1D |> Array.filter f

    // -----------------------------------
    // -- Registry functions
    // -----------------------------------
    module Reg =
        module Out =
            let out (regKey: string) (unwrapOptions: bool) (refKey: String) (proxys: Proxys) : obj[] option =
                let methodNm = "out"
                MRegistry.tryFind1D regKey
                |> Option.map (apply<GenFn> methodNm [||] [| unwrapOptions; refKey; proxys |])
                |> Option.map (fun o -> o :?> obj[])

            let count (xlValue: string) : obj option =  // FIXME: rename arg to regKey
                let methodNm = "count"
                MRegistry.tryFind1D xlValue
                |> Option.map (apply<GenFn> methodNm [||] [||])

            let tryElem (unwrapOptions: bool) (refKey: String) (proxys: Proxys) (index: int) (xlValue: string) : obj option =
                let methodNm = "tryElem"
                MRegistry.tryFind1D xlValue
                |> Option.map (apply<GenFn> methodNm [||] [| unwrapOptions; refKey; proxys; index |])

            let sub (regKey: string) (startIndex: int option) (count: int option) : obj option =
                let methodNm = "sub"
                MRegistry.tryFind1D regKey
                |> Option.map (apply<GenFn> methodNm [||] [| startIndex; count |])

            let private append2' (tys1:Type[], o1:obj) (tys2:Type[], o2:obj) : obj option =
                let methodNm = "append2"
                if tys2 <> tys1 then
                    None
                else
                    applyMulti<GenFn> methodNm [||] [||] tys1 [| o1; o2 |]
                    |> Some

            let append2 (regKey1: string)  (regKey2: string) : obj option =
                match MRegistry.tryFind1D regKey1, MRegistry.tryFind1D regKey2 with
                | None, None -> None
                | Some (_, o1), None -> Some o1
                | None, Some (_, o2) -> Some o2
                | Some (tys1, o1), Some (tys2, o2) -> append2' (tys1, o1) (tys2, o2)

            let append3 (regKey1: string)  (regKey2: string)  (regKey3: string) : obj option =
                let methodNm = "append3"
                match MRegistry.tryFind1D regKey1, MRegistry.tryFind1D regKey2, MRegistry.tryFind1D regKey3 with
                | None, None, None -> None
                | Some (_, o1), None, None -> Some o1
                | None, Some (_, o2), None -> Some o2
                | None, None, Some (_, o3) -> Some o3
                | Some (tys1, o1), Some (tys2, o2), None -> append2' (tys1, o1) (tys2, o2)
                | Some (tys1, o1), None, Some (tys3, o3) -> append2' (tys1, o1) (tys3, o3)
                | None, Some (tys2, o2), Some (tys3, o3) -> append2' (tys2, o2) (tys3, o3)
                | Some (tys1, o1), Some (tys2, o2), Some (tys3, o3) -> 
                    if (tys2 <> tys1) || (tys3 <> tys1) then
                        None
                    else
                        applyMulti<GenFn> methodNm [||] [||] tys1 [| o1; o2; o3 |]
                        |> Some
            
            // TODO : should not be in Out module (Better in Reg module?)
            let sort (descending: bool) (xlValue: string) : obj option =
                let methodNm = "sort"
                MRegistry.tryFind1D xlValue
                |> Option.map (apply<GenFn> methodNm [||] [| descending |])

            let map1 (ofun: obj) (xlValue: string) : obj option =
                match MRegistry.tryFind1D xlValue with
                | None -> None
                | Some (elemtys, xa1D) ->
                    match Fun.arity ofun with                    
                    | Some rank when rank = 1 -> 
                        let elemty = elemtys.[0]
                        let funtys = ofun.GetType().BaseType.GetGenericArguments()
                        if funtys.[0] <> elemty then
                            None
                        else
                            let methodNm = "map1"
                            let res = Toolbox.Generics.invoke<GenFn> methodNm funtys [| ofun; xa1D |]
                            Some res
                    | _ -> None

            let filter (ofun: obj) (xlValue: string) : obj option =
                match MRegistry.tryFind1D xlValue with
                | None -> None
                | Some (elemtys, xa1D) ->
                    match Fun.arity ofun with                    
                    | Some rank when rank = 1 -> 
                        let elemty = elemtys.[0]
                        let funtys = ofun.GetType().BaseType.GetGenericArguments()
                        if (funtys.[0] <> elemty) || (funtys |> Array.last <> typeof<bool>) then
                            None
                        else
                            let methodNm = "filter"
                            let res = Toolbox.Generics.invoke<GenFn> methodNm (funtys |> Array.take 1) [| ofun; xa1D |]
                            Some res
                    | _ -> None

module A1D_XL =
    open Registry
    open API
    open API.Out
    open type Variant
    open type Out.Proxys
    open FSharp.Compiler.Interactive.Shell

    // open API.In.D0

    [<ExcelFunction(Category="Array1D", Description="Cast a 1D-slice of an xl-range to a generic type array.")>]
    let a1_ofRng
        ([<ExcelArgument(Description= "1D xl-range.")>] range: obj)
        ([<ExcelArgument(Description= "Type tag: bool, date, double, doubleNaN, string or obj. Add \'#'\ prefix for optional type.")>] typeTag: string)
        ([<ExcelArgument(Description= "[Replacement method for wrong-type elements. \"[R]eplace\", \"[F]ilter\", \"[S]trict\", \"[E]mptyStrict\". Default is \"Strict\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value (only for non-optional types, optional types default to None).]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        ([<ExcelArgument(Description= "[Failure value. Default is #N/A.]")>] failureValue: obj)
        ([<ExcelArgument(Description= "[Only doubleNaN tag: Kinds for which values are converted to Double.NaN. E.g. NA, ERR, TXT, !NUM... Default is none.]")>] xlKinds: obj)
        : obj  =

        // intermediary stage
        let rowwise = In.D0.Bool.Opt.def None rowWiseDirection
        let replmethod = In.D0.Stg.def "STRICT" replaceMethod
        let defVal = In.D0.Absent.Obj.tryO defaultValue
        let failureVal = In.D0.Missing.Obj.subst Proxys.def.failed failureValue
        let isoptional = isOptionalType typeTag
        let xlkinds = In.D0.Stg.def "NONE" xlKinds |> Kind.ofLabel
        
        // caller cell's reference ID
        let rfid = MRegistry.refID

        // wording
        match (replmethod.ToUpper().Substring(0,1)), isoptional with
        | "F", _ -> 
            let xa1D = In.D1.Tag.filter xlkinds rowwise typeTag range
            let res = xa1D |> MRegistry.register rfid
            box res

        // strict / empty-strict methods: 
        //    - return a 1D array if *all* of the array's elements are of expected type (as determined by typeTag)
        // empty-strict: returns an empty array otherwise.
        // strict: return None otherwise. Here returns failed value.
        | "E", _ -> 
            let xa1D = In.D1.Tag.tryEmpty xlkinds rowwise typeTag range
            let res = xa1D |> MRegistry.register rfid
            box res
        | "S", _ -> 
            match In.D1.Tag.Try.tryDV xlkinds rowwise None typeTag range |> Option.map snd with
            | None -> failureVal
            | Some xa1D -> 
                let res = xa1D |> MRegistry.register rfid
                box res
        | _ ->  let xa1D = In.D1.Tag.Any.def xlkinds rowwise defVal typeTag range
                let res = xa1D |> MRegistry.register rfid
                box res

    [<ExcelFunction(Category="Array1D", Description="Extracts an array out of a R-object.")>]
    let a1_toRng
        ([<ExcelArgument(Description= "1D array R-object.")>] rgA1D: string)
        ([<ExcelArgument(Description= "[None indicator. Default is \"<none>\".]")>] noneIndicator: obj)
        ([<ExcelArgument(Description= "[Empty array indicator. Default is \"<empty>\".]")>] emptyIndicator: obj)
        ([<ExcelArgument(Description= "[Unwrap optional types. Default is true.]")>] unwrapOptions: obj)
        // TODO: add nan indicator?
        : obj[] = 
        
        // intermediary stage
        let none = In.D0.Stg.def "<none>" noneIndicator
        let empty = In.D0.Stg.def "<empty>" emptyIndicator
        let proxys = { def with none = none; empty = empty }
        let unwrapoptions = In.D0.Bool.def true unwrapOptions

        // caller cell's reference ID
        let rfid = MRegistry.refID
        
        // result
        match A1D.Reg.Out.out rgA1D unwrapoptions rfid proxys with
        | None -> [| box ExcelError.ExcelErrorNA |]
        | Some o1D -> o1D

    [<ExcelFunction(Category="Array1D", Description="Returns the size of a R-object array.")>]
    let a1_count
        ([<ExcelArgument(Description= "1D array R-object.")>] rgA1D: string) 
        ([<ExcelArgument(Description= "[None indicator. Default is \"<none>\".]")>] noneIndicator: obj) // FIXME: why this arg?
        ([<ExcelArgument(Description= "[Empty array indicator. Default is \"<empty>\".]")>] emptyIndicator: obj) // FIXME: why this arg?
        : obj = 

        // intermediary stage
        let none = In.D0.Stg.def "<none>" noneIndicator
        let empty = In.D0.Stg.def "<empty>" emptyIndicator
        let proxys = { def with none = none; empty = empty }

        // result
        match A1D.Reg.Out.count rgA1D with
        | None -> proxys.failed  // TODO Unbox.apply?
        | Some o -> o
        
    [<ExcelFunction(Category="Array1D", Description="Returns an element of a R-object array.")>]
    let a1_elem
        ([<ExcelArgument(Description= "1D array R-object.")>] rgA1D: string) 
        ([<ExcelArgument(Description= "[Index. Default is 0.]")>] index: obj)
        ([<ExcelArgument(Description= "[None indicator. Default is \"<none>\".]")>] noneIndicator: obj)
        ([<ExcelArgument(Description= "[Unwrap optional types. Default is true.]")>] unwrapOptions: obj)
        ([<ExcelArgument(Description= "[Failure value. Default is #N/A.]")>] failureValue: obj)
        : obj = 

        // intermediary stage
        let index = In.D0.Intg.def 0 index

        let none = In.D0.Stg.def "<none>" noneIndicator
        let failureVal = In.D0.Missing.Obj.subst Proxys.def.failed failureValue
        let proxys = { def with none = none; failed = failureVal }
        let unwrapoptions = In.D0.Bool.def true unwrapOptions

        // caller cell's reference ID (necessary when the elements are not primitive types)
        let rfid = MRegistry.refID
        
        // result
        match A1D.Reg.Out.tryElem unwrapoptions rfid proxys index rgA1D with
        | None -> proxys.failed  // TODO Unbox.apply?
        | Some xo0D -> xo0D

    [<ExcelFunction(Category="Array1D", Description="Returns the sub-array of a R-object array.")>]
    let a1_sub
        ([<ExcelArgument(Description= "1D array R-object.")>] rgA1D: string) 
        ([<ExcelArgument(Description= "[Sub count. Default is full length.]")>] subCount: obj)
        ([<ExcelArgument(Description= "[Start index. Default is 0.]")>] startIndex: obj)
        : obj = 

        // intermediary stage
        let count = In.D0.Intg.Opt.def None subCount
        let start = In.D0.Intg.Opt.def None startIndex

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match A1D.Reg.Out.sub rgA1D start count with
        | None -> Proxys.def.failed  // TODO Unbox.apply?
        | Some o -> o |> MRegistry.register rfid |> box

    [<ExcelFunction(Category="Array1D", Description="Appends several R-object arrays to each other.")>]
    let a1_append
        ([<ExcelArgument(Description= "1D array1 R-object.")>] rgA1D1: string) 
        ([<ExcelArgument(Description= "1D array2 R-object.")>] rgA1D2: string) 
        ([<ExcelArgument(Description= "1D array2 R-object.")>] rgA1D3: obj) 
        : obj = 

        // intermediary stage
        let rO3 = In.D0.Stg.Opt.def None rgA1D3

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match rO3 with
        | None -> 
            match A1D.Reg.Out.append2 rgA1D1 rgA1D2 with
            | None -> Proxys.def.failed  // TODO Unbox.apply?
            | Some o -> o |> MRegistry.register rfid |> box
        | Some rO3 -> 
            match A1D.Reg.Out.append3 rgA1D1 rgA1D2 rO3 with
            | None -> Proxys.def.failed  // TODO Unbox.apply?
            | Some o -> o |> MRegistry.register rfid |> box

    [<ExcelFunction(Category="Array1D", Description="Sort a R-object array.")>]
    let a1_sort
        ([<ExcelArgument(Description= "1D array R-object.")>] rgA1D: string) 
        ([<ExcelArgument(Description= "[Descending. Default is false.]")>] descending: obj)
        : obj = 

        // intermediary stage
        let desc = In.D0.Bool.def false descending

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match A1D.Reg.Out.sort desc rgA1D with
        | None -> Proxys.def.failed
        | Some xa1D -> xa1D |> MRegistry.registerBxd rfid

    [<ExcelFunction(Category="Array1D", Description="Applies a function to each element of the array.")>]
    let a1_map
        ([<ExcelArgument(Description= "1D array R-object.")>] rgA1D: string) 
        ([<ExcelArgument(Description= "Multiline expression. E.g. \"\\ x -> 2.0 * x\"")>] expression: obj)
        ([<ExcelArgument(Description= "[FSharpFunc R-object. Disable expression if present. Default is none.]")>] rgFSharpFunc: obj)
        ([<ExcelArgument(Description= "[Ad hoc FSI session R-object. Inactive if FSharpFunc object provided. Default is internal session.]")>] rgFsiSession: obj)
        : obj =

        // intermediary stage
        let expression = API.In.D1.OStg.tryDV None expression |> Option.map (fun exprs -> String.Join("\n", exprs))
        let rgFSharpFunc = API.In.D0.Stg.Opt.def None rgFSharpFunc

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match Fun.Reg.fsiFunction MRegistry def.failed rfid rgFsiSession expression rgFSharpFunc with
        | None -> def.failed
        | Some ofun ->
            match A1D.Reg.Out.map1 ofun rgA1D with
            | None -> def.failed
            | Some res -> res |> MRegistry.registerBxd rfid

    [<ExcelFunction(Category="Array1D", Description="Filters the array given a predicate function.")>]
    let a1_filter
        ([<ExcelArgument(Description= "1D array R-object.")>] rgA1D: string) 
        ([<ExcelArgument(Description= "Multiline expression. E.g. \"\\ x -> 2.0 * x\"")>] expression: obj)
        ([<ExcelArgument(Description= "[FSharpFunc R-object. Disable expression if present. Default is none.]")>] rgFSharpFunc: obj)
        ([<ExcelArgument(Description= "[Ad hoc FSI session R-object. Inactive if FSharpFunc object provided. Default is internal session.]")>] rgFsiSession: obj)
        : obj =

        // intermediary stage
        let expression = API.In.D1.OStg.tryDV None expression |> Option.map (fun exprs -> String.Join("\n", exprs))
        let rgFSharpFunc = API.In.D0.Stg.Opt.def None rgFSharpFunc

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match Fun.Reg.fsiFunction MRegistry def.failed rfid rgFsiSession expression rgFSharpFunc with
        | None -> def.failed
        | Some ofun ->
            match A1D.Reg.Out.filter ofun rgA1D with
            | None -> def.failed
            | Some res -> res |> MRegistry.registerBxd rfid

    [<ExcelFunction(Category="Array1D", Description="Applies a function to each element of the array.")>]
    let a1_mapOLD
        ([<ExcelArgument(Description= "1D array R-object.")>] rgA1D: string)
        ([<ExcelArgument(Description= "FSharpFunc R-object.")>] rgFSharpFunc: string)
        : obj =

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match MRegistry.tryExtractO rgFSharpFunc with
        | None -> def.failed
        | Some ofun ->
            match A1D.Reg.Out.map1 ofun rgA1D with
            | None -> def.failed
            | Some res -> res |> MRegistry.registerBxd rfid

module A2D = 
    open Registry
    open Toolbox.Generics
    open Toolbox.Array2D
    open API.Out

    //// -----------------------------
    //// -- Main functions
    //// -----------------------------

    ///// Empty 2D array.
    //let empty2D<'a> : 'a[,] = [|[||]|] |> array2D

    ///// Returns true if the first dimension is 0.
    //let isEmpty (a2D: 'a[,]) : bool = a2D |> Array2D.length1 = 0 // is this the right way?

    ///// Convenience function to create a 2D singleton.
    //let singleton<'a> (a: 'a) = Array2D.create 1 1 a

    //let sub' (a2D : 'a[,]) (rowStartIndex: int) (colStartIndex: int) (rowCount: int) (colCount: int) : 'a[,] =
    //    let rowLen, colLen = a2D |> Array2D.length1, a2D |> Array2D.length2

    //    if (rowStartIndex >= rowLen) || (colStartIndex >= colLen) then
    //        empty2D<'a>
    //    else
    //        let rowstart = max 0 rowStartIndex
    //        let colstart = max 0 colStartIndex
    //        let rowcount = (min (rowLen - rowstart) rowCount) |> max 0
    //        let colcount = (min (colLen - colstart) colCount) |> max 0
    //        a2D.[rowstart..(rowstart + rowcount - 1), colstart..(colstart + colcount - 1)]
    
    //let sub (rowStartIndex: int option) (colStartIndex: int option) (rowCount: int option) (colCount: int option) (a2D : 'a[,]) : 'a[,] =
    //    let rowLen, colLen = a2D |> Array2D.length1, a2D |> Array2D.length2

    //    let rowidx = rowStartIndex |> Option.defaultValue 0
    //    let colidx = colStartIndex |> Option.defaultValue 0
    //    let rowcnt = rowCount |> Option.defaultValue rowLen
    //    let colcnt = colCount |> Option.defaultValue colLen
    //    sub' a2D rowidx colidx rowcnt colcnt
        
    // -----------------------------------
    // -- Reflection functions
    // -----------------------------------
    type GenFn =
        static member out<'A> (a2D: 'A[,]) (unwrapOptions: bool) (refKey: String) (proxys: Proxys) : obj[,] = 
            a2D |> Array2D.map box |> (API.Out.D2.Reg.out<'A> unwrapOptions refKey proxys)
            
        static member rowCount<'A> (a2D: 'A[,]) : int = a2D |> Array2D.length1

        static member colCount<'A> (a2D: 'A[,]) : int = a2D |> Array2D.length2

        static member elem<'A> (a2D: 'A[,]) (rowIndex: int) (colIndex: int) : 'A = a2D.[rowIndex, colIndex]

        static member sub<'A> (a2D: 'A[,]) (rowStartIndex: int option) (colStartIndex: int option) (rowCount: int option) (colCount: int option) : 'A[,] =
            a2D |> sub rowStartIndex colStartIndex rowCount colCount

        static member sort<'A when 'A: comparison> (a2D: 'A[,]) (descending: bool) : 'A[,] = 
            let len1, len2 = a2D |> Array2D.length1, a2D |> Array2D.length2
            if (len1 = 0) || (len2 = 0) then
                a2D
            else
                let a2sort = if descending then Array.sortDescending else Array.sort
                // works only rowwise TODO: transpose
                [| for i in a2D.GetLowerBound(0) .. a2D.GetUpperBound(0) -> a2D.[i,*] |]
                |> a2sort
                |> array2D


        //static member append2<'A> (a1D1: 'A[,]) (a1D2: 'A[,]) : 'A[,] =
        //    Array.append a1D1 a1D2

        //static member append3<'A> (a1D1: 'A[,]) (a1D2: 'A[,]) (a1D3: 'A[,]) : 'A[,] =
        //    Array.append (Array.append a1D1 a1D2) a1D3
        
    // -----------------------------------
    // -- Registry functions
    // -----------------------------------
    module Reg =
        module Out =
            let out (regKey: string) (unwrapOptions: bool) (refKey: String) (proxys: Proxys) : obj[,] option =
                let methodNm = "out"
                MRegistry.tryFind2D regKey
                |> Option.map (apply<GenFn> methodNm [||] [| unwrapOptions; refKey; proxys |])
                |> Option.map (fun o -> o :?> obj[,])

            let rowCount (xlValue: string) : obj option =
                let methodNm = "rowCount"
                MRegistry.tryFind2D xlValue
                |> Option.map (apply<GenFn> methodNm [||] [||])

            let colCount (xlValue: string) : obj option =
                let methodNm = "colCount"
                MRegistry.tryFind2D xlValue
                |> Option.map (apply<GenFn> methodNm [||] [||])

            let elem (rowIndex: int) (colIndex: int) (xlValue: string) : obj option =
                let methodNm = "elem"
                MRegistry.tryFind2D xlValue
                |> Option.map (apply<GenFn> methodNm [||] [| rowIndex; colIndex |])

            let sub (regKey: string) (rowStartIndex: int option) (colStartIndex: int option) (rowCount: int option) (colCount: int option) : obj option =
                let methodNm = "sub"
                MRegistry.tryFind1D regKey
                |> Option.map (apply<GenFn> methodNm [||] [| rowStartIndex; rowStartIndex; rowCount; colCount |])

            let private append2' (tys1:Type[], o1:obj) (tys2:Type[], o2:obj) : obj option =
                let methodNm = "append2"
                if tys2 <> tys1 then
                    None
                else
                    applyMulti<GenFn> methodNm [||] [||] tys1 [| o1; o2 |]
                    |> Some

            let append2 (regKey1: string)  (regKey2: string) : obj option =
                match MRegistry.tryFind1D regKey1, MRegistry.tryFind1D regKey2 with
                | None, None -> None
                | Some (_, o1), None -> Some o1
                | None, Some (_, o2) -> Some o2
                | Some (tys1, o1), Some (tys2, o2) -> append2' (tys1, o1) (tys2, o2)

            let append3 (regKey1: string)  (regKey2: string)  (regKey3: string) : obj option =
                let methodNm = "append3"
                match MRegistry.tryFind1D regKey1, MRegistry.tryFind1D regKey2, MRegistry.tryFind1D regKey3 with
                | None, None, None -> None
                | Some (_, o1), None, None -> Some o1
                | None, Some (_, o2), None -> Some o2
                | None, None, Some (_, o3) -> Some o3
                | Some (tys1, o1), Some (tys2, o2), None -> append2' (tys1, o1) (tys2, o2)
                | Some (tys1, o1), None, Some (tys3, o3) -> append2' (tys1, o1) (tys3, o3)
                | None, Some (tys2, o2), Some (tys3, o3) -> append2' (tys2, o2) (tys3, o3)
                | Some (tys1, o1), Some (tys2, o2), Some (tys3, o3) -> 
                    if (tys2 <> tys1) || (tys3 <> tys1) then
                        None
                    else
                        applyMulti<GenFn> methodNm [||] [||] tys1 [| o1; o2; o3 |]
                        |> Some

            let sort (descending: bool) (xlValue: string) : obj option =
                let methodNm = "sort"
                MRegistry.tryFind2D xlValue
                |> Option.map (apply<GenFn> methodNm [||] [| descending |])

module A2D_XL =
    open Registry
    open API
    open API.Out
    open type Out.Proxys

    [<ExcelFunction(Category="Array2D", Description="Cast a 2D xl-range to a generic type array.")>]
    let a2_ofRng
        ([<ExcelArgument(Description= "2D xl-range.")>] range: obj)
        ([<ExcelArgument(Description= "Type tag: bool, date, double, doubleNaN, string or obj. Add \'#'\ prefix for optional type: #bool, #date, #double, #doubleNaN, #string or #obj")>] typeTag: string)
        ([<ExcelArgument(Description= "[Replacement method for non-date elements. \"Replace\", \"Filter\" or \"Strict\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value (only for non-optional types, optional types default to None). Must be of the appropriate type, else it will fail.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[Row wise direction for filtering. Default is true.]")>] rowWiseDirection: obj)
        ([<ExcelArgument(Description= "[Only doubleNaN tag: Kinds for which values are converted to Double.NaN. E.g. NA, ERR, TXT, !NUM... Default is none.]")>] xlKinds: obj)
        : obj  =

        // intermediary stage
        let rowWise = In.D0.Bool.def true rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let defVal = In.D0.Absent.Obj.tryO defaultValue
        let xlkinds = In.D0.Stg.def "NONE" xlKinds |> Kind.ofLabel
        
        // caller cell's reference ID
        let rfid = MRegistry.refID
        // TODO: strict method?
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let xa2D = (In.D2.Tag.filter xlkinds rowWise typeTag range)
                 let res = xa2D |> MRegistry.register rfid 
                 box res
        | _ -> let res = (In.D2.Tag.Any.def xlkinds defVal typeTag range) |> MRegistry.register rfid
               box res

    [<ExcelFunction(Category="Array2D", Description="Extracts a 2D array out of a R-object.")>]
    let a2_toRng
        ([<ExcelArgument(Description= "2D array R-object.")>] rgA2D: string)
        ([<ExcelArgument(Description= "[None indicator. Default is \"<none>\".]")>] noneIndicator: obj)
        ([<ExcelArgument(Description= "[Empty array indicator. Default is \"<empty>\".]")>] emptyIndicator: obj)
        ([<ExcelArgument(Description= "[Unwrap optional types. Default is true.]")>] unwrapOptions: obj)
        // TODO: add nan indicator?
        : obj[,] = 

        // intermediary stage
        let none = In.D0.Stg.def "<none>" noneIndicator
        let empty = In.D0.Stg.def "<empty>" emptyIndicator
        let proxys = { def with none = none; empty = empty }
        let unwrapoptions = In.D0.Bool.def true unwrapOptions

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match A2D.Reg.Out.out rgA2D unwrapoptions rfid proxys with
        | None -> box ExcelError.ExcelErrorNA |> Toolbox.Array2D.singleton
        | Some o2d -> o2d

    [<ExcelFunction(Category="Array2D", Description="Returns the number of rows of a R-object array.")>]
    let a2_rows
        ([<ExcelArgument(Description= "2D array R-object.")>] rgA2D: string) 
        ([<ExcelArgument(Description= "[None indicator. Default is \"<none>\".]")>] noneIndicator: obj)
        ([<ExcelArgument(Description= "[Empty array indicator. Default is \"<empty>\".]")>] emptyIndicator: obj)
        : obj = 

        // intermediary stage
        let none = In.D0.Stg.def "<none>" noneIndicator
        let empty = In.D0.Stg.def "<empty>" emptyIndicator
        let proxys = { def with none = none; empty = empty }

        // result
        match A2D.Reg.Out.rowCount rgA2D with
        | None -> proxys.failed  // TODO Unbox.apply?
        | Some o -> o

    [<ExcelFunction(Category="Array2D", Description="Returns the number of rows of a R-object array.")>]
    let a2_cols
        ([<ExcelArgument(Description= "2D array R-object.")>] rgA2D: string) 
        ([<ExcelArgument(Description= "[None indicator. Default is \"<none>\".]")>] noneIndicator: obj)
        ([<ExcelArgument(Description= "[Empty array indicator. Default is \"<empty>\".]")>] emptyIndicator: obj)
        : obj = 

        // intermediary stage
        let none = In.D0.Stg.def "<none>" noneIndicator
        let empty = In.D0.Stg.def "<empty>" emptyIndicator
        let proxys = { def with none = none; empty = empty }

        // result
        match A2D.Reg.Out.colCount rgA2D with
        | None -> proxys.failed  // TODO Unbox.apply?
        | Some o -> o

    [<ExcelFunction(Category="Array2D", Description="Returns an element of a R-object array.")>]
    let a2_elem
        ([<ExcelArgument(Description= "2D array R-object.")>] rgA2D: string) 
        ([<ExcelArgument(Description= "[Row index. Default is 0.]")>] rowIndex: obj)
        ([<ExcelArgument(Description= "[Column index. Default is 0.]")>] colIndex: obj)
        ([<ExcelArgument(Description= "[None indicator. Default is \"<none>\".]")>] noneIndicator: obj)
        ([<ExcelArgument(Description= "[Unwrap optional types. Default is true.]")>] unwrapOptions: obj)
        : obj = 

        // intermediary stage
        let rowindex = In.D0.Intg.def 0 rowIndex
        let colindex = In.D0.Intg.def 0 colIndex

        let none = In.D0.Stg.def "<none>" noneIndicator
        let proxys = { def with none = none }
        let unwrapoptions = In.D0.Bool.def true unwrapOptions

        // caller cell's reference ID (necessary when the elements are not primitive types)
        let rfid = MRegistry.refID

        // result
        match A2D.Reg.Out.elem rowindex colindex rgA2D with
        | None -> proxys.failed  // TODO Unbox.apply?
        | Some o -> o |> API.Out.D0.Reg.out false unwrapoptions rfid proxys

    [<ExcelFunction(Category="Array2D", Description="Returns a sub-array of a R-object 2D-array.")>]
    let a2_sub
        ([<ExcelArgument(Description= "2D array R-object.")>] rgA2D: string) 
        ([<ExcelArgument(Description= "[Sub row count. Default is full length.]")>] rowCount: obj)
        ([<ExcelArgument(Description= "[Sub col count. Default is full length.]")>] colCount: obj)
        ([<ExcelArgument(Description= "[Row start index. Default is 0.]")>] rowStartIndex: obj)
        ([<ExcelArgument(Description= "[Col start index. Default is 0.]")>] colStartIndex: obj)
        : obj = 

        // intermediary stage
        let rowcount = In.D0.Intg.Opt.def None rowCount
        let colcount = In.D0.Intg.Opt.def None colCount
        let rowstart = In.D0.Intg.Opt.def None rowStartIndex
        let colstart = In.D0.Intg.Opt.def None colStartIndex

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match A2D.Reg.Out.sub rgA2D rowstart colstart rowcount colcount with
        | None -> Proxys.def.failed  // TODO Unbox.apply?
        | Some o -> o |> MRegistry.register rfid |> box

    [<ExcelFunction(Category="Array1D", Description="Sort a R-object 2D-array.")>]
    let a2_sort
        ([<ExcelArgument(Description= "2D array R-object.")>] rgA2D: string) 
        ([<ExcelArgument(Description= "[Descending. Default is false.]")>] descending: obj)
        : obj = 

        // intermediary stage
        let desc = In.D0.Bool.def false descending

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match A2D.Reg.Out.sort desc rgA2D with
        | None -> Proxys.def.failed
        | Some xa2D -> xa2D |> MRegistry.registerBxd rfid

    [<ExcelFunction(Category="Array2D", Description="Cast a 2D xl-range to a generic type array.")>]
    let a2_ofRngMulti
        ([<ExcelArgument(Description= "2D xl-range.")>] range: obj)
        ([<ExcelArgument(Description= "Type tag: bool, date, double, doubleNaN, string or obj. Add \'#'\ prefix for optional type: #bool, #date, #double, #doubleNaN, #string or #obj")>] typeTag: string)
        ([<ExcelArgument(Description= "[Replacement method for non-date elements. \"Replace\", \"Filter\" or \"Strict\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value (only for non-optional types, optional types default to None). Must be of the appropriate type, else it will fail.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[Row wise direction for filtering. Default is true.]")>] rowWiseDirection: obj)
        ([<ExcelArgument(Description= "[Only doubleNaN tag: Kinds for which values are converted to Double.NaN. E.g. NA, ERR, TXT, !NUM... Default is none.]")>] xlKinds: obj)
        : obj  =

        // intermediary stage
        let rowWise = In.D0.Bool.def true rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let defVal = In.D0.Absent.Obj.tryO defaultValue
        let xlkinds = In.D0.Stg.def "NONE" xlKinds |> Kind.ofLabel
        
        // caller cell's reference ID
        let rfid = MRegistry.refID
        // TODO: strict method?
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let xa2D = (In.D2.Tag.filter xlkinds rowWise typeTag range)
                 let res = xa2D |> MRegistry.register rfid 
                 box res
        | _ -> let res = (In.D2.Tag.Any.def xlkinds defVal typeTag range) |> MRegistry.register rfid
               box res

module MAP = 
    open Registry
    // open Toolbox.Generics
    open Toolbox.Array
    open Microsoft.FSharp.Reflection
    open API.Out

    // -----------------------------------
    // -- Main functions
    // -----------------------------------

    /// -----------------------------------
    /// -- Generic functions
    /// -----------------------------------
    type GenFn =

        // -----------------------------------
        // -- Inspection functions
        // -----------------------------------

        /// Returns the number of kvp in the Map's object.
        static member count<'K,'V when 'K: comparison> (map: Map<'K,'V>) : int =
            map |> Map.count

        /// wording : returns keys 1D array to Excel
        static member keys<'K,'V when 'K: comparison> (map: Map<'K,'V>) (refKey: String) (proxys: Proxys) : obj[] =
            let a1D = [| for kvp in map -> kvp.Key |]
            a1D |> Array.map box |> (API.Out.D1.Reg.out<'K> true false refKey proxys)

        ///// wording : returns values 1D array to Excel
        static member values<'K,'V when 'K: comparison> (map: Map<'K,'V>) (unwrapOptions: bool) (refKey: String) (proxys: Proxys) : obj[] =
            let a1D = [| for kvp in map -> kvp.Value |]
            a1D |> Array.map box |> (API.Out.D1.Reg.out<'V> true unwrapOptions refKey proxys)

        static member cast<'K>  (okey: obj) : obj =
            if typeof<'K> = typeof<int> then
                match API.In.D0.Intg.Opt.def None okey with
                | Some i -> box i
                | None -> okey // probable run-time failure.
            elif typeof<'K> = typeof<DateTime> then
                match API.In.D0.Dte.Opt.def None okey with
                | Some dte -> box dte
                | None -> okey // probable run-time failure.
            elif typeof<'K> = typeof<double> then
                match okey with
                | :? ExcelError as e when e = ExcelError.ExcelErrorNA -> Double.NaN |> box
                | _ -> okey
            else
                okey

        static member find1<'K1,'V when 'K1: comparison> (map: Map<'K1,'V>) (append: bool) (proxys: Proxys) (refKey: String) (okey1: obj) : obj =
            match GenFn.cast<'K1> okey1 with 
            | :? 'K1 as key1 ->
                match map |> Map.tryFind key1 with
                | None -> proxys.failed
                | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> append false refKey proxys)
            | _ -> proxys.failed

        static member find1D1<'K1,'V when 'K1: comparison> (map: Map<'K1,'V>) (proxys: Proxys) (refKey: String) (okeys1: obj[]) : obj[] =
            Registry.MRegistry.removeReferencedObjects refKey
            okeys1 |> Array.map (GenFn.find1 map true proxys refKey)

        /// wording : returns values 1D array to Excel
        static member find2<'K1,'K2,'V when 'K1: comparison and 'K2: comparison> 
            (map: Map<'K1*'K2,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) 
            : obj =
                match GenFn.cast<'K1> okey1, GenFn.cast<'K2> okey2 with 
                | (:? 'K1 as key1), (:? 'K2 as key2) ->
                    match map |> Map.tryFind (key1, key2) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false false refKey proxys)
                | _ -> proxys.failed

        static member find3<'K1,'K2,'K3,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison> 
            (map: Map<'K1*'K2*'K3,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) (okey3: obj) 
            : obj =
                match GenFn.cast<'K1> okey1, GenFn.cast<'K2> okey2, GenFn.cast<'K3> okey3 with 
                | (:? 'K1 as key1), (:? 'K2 as key2), (:? 'K3 as key3) ->
                    match map |> Map.tryFind (key1, key2, key3) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false false refKey proxys)
                | _ -> proxys.failed

        static member find4<'K1,'K2,'K3,'K4,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison> 
            (map: Map<'K1*'K2*'K3*'K4,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) (okey3: obj) (okey4: obj)
            : obj =
                match GenFn.cast<'K1> okey1, GenFn.cast<'K2> okey2, GenFn.cast<'K3> okey3, GenFn.cast<'K4> okey4 with 
                | (:? 'K1 as key1), (:? 'K2 as key2), (:? 'K3 as key3), (:? 'K4 as key4) ->
                    match map |> Map.tryFind (key1, key2, key3, key4) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false false refKey proxys)
                | _ -> proxys.failed

        static member find5<'K1,'K2,'K3,'K4,'K5,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison> 
            (map: Map<'K1*'K2*'K3*'K4*'K5,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) (okey3: obj) (okey4: obj) (okey5: obj)
            : obj =
                match GenFn.cast<'K1> okey1, GenFn.cast<'K2> okey2, GenFn.cast<'K3> okey3, GenFn.cast<'K4> okey4, GenFn.cast<'K5> okey5 with 
                | (:? 'K1 as key1), (:? 'K2 as key2), (:? 'K3 as key3), (:? 'K4 as key4), (:? 'K5 as key5) ->
                    match map |> Map.tryFind (key1, key2, key3, key4, key5) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false false refKey proxys)
                | _ -> proxys.failed

        static member find6<'K1,'K2,'K3,'K4,'K5,'K6,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison> 
            (map: Map<'K1*'K2*'K3*'K4*'K5*'K6,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) (okey3: obj) (okey4: obj) (okey5: obj) (okey6: obj)
            : obj =
                match GenFn.cast<'K1> okey1, GenFn.cast<'K2> okey2, GenFn.cast<'K3> okey3, GenFn.cast<'K4> okey4, GenFn.cast<'K5> okey5, GenFn.cast<'K6> okey6 with 
                | (:? 'K1 as key1), (:? 'K2 as key2), (:? 'K3 as key3), (:? 'K4 as key4), (:? 'K5 as key5), (:? 'K6 as key6) ->
                    match map |> Map.tryFind (key1, key2, key3, key4, key5, key6) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false false refKey proxys)
                | _ -> proxys.failed

        static member find7<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison> 
            (map: Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) (okey3: obj) (okey4: obj) (okey5: obj) (okey6: obj) (okey7: obj)
            : obj =
                match GenFn.cast<'K1> okey1, GenFn.cast<'K2> okey2, GenFn.cast<'K3> okey3, GenFn.cast<'K4> okey4, GenFn.cast<'K5> okey5, GenFn.cast<'K6> okey6, GenFn.cast<'K7> okey7 with 
                | (:? 'K1 as key1), (:? 'K2 as key2), (:? 'K3 as key3), (:? 'K4 as key4), (:? 'K5 as key5), (:? 'K6 as key6), (:? 'K7 as key7) ->
                    match map |> Map.tryFind (key1, key2, key3, key4, key5, key6, key7) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false false refKey proxys)
                | _ -> proxys.failed

        static member find8<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'K8,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison and 'K8: comparison> 
            (map: Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) (okey3: obj) (okey4: obj) (okey5: obj) (okey6: obj) (okey7: obj) (okey8: obj)
            : obj =
                match GenFn.cast<'K1> okey1, GenFn.cast<'K2> okey2, GenFn.cast<'K3> okey3, GenFn.cast<'K4> okey4, GenFn.cast<'K5> okey5, GenFn.cast<'K6> okey6, GenFn.cast<'K7> okey7, GenFn.cast<'K8> okey8 with 
                | (:? 'K1 as key1), (:? 'K2 as key2), (:? 'K3 as key3), (:? 'K4 as key4), (:? 'K5 as key5), (:? 'K6 as key6), (:? 'K7 as key7), (:? 'K8 as key8) ->
                    match map |> Map.tryFind (key1, key2, key3, key4, key5, key6, key7, key8) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false false refKey proxys)
                | _ -> proxys.failed

        static member find9<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'K8,'K9,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison and 'K8: comparison and 'K9: comparison> 
            (map: Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8*'K9,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) (okey3: obj) (okey4: obj) (okey5: obj) (okey6: obj) (okey7: obj) (okey8: obj) (okey9: obj)
            : obj =
                match GenFn.cast<'K1> okey1, GenFn.cast<'K2> okey2, GenFn.cast<'K3> okey3, GenFn.cast<'K4> okey4, GenFn.cast<'K5> okey5, GenFn.cast<'K6> okey6, GenFn.cast<'K7> okey7, GenFn.cast<'K8> okey8, GenFn.cast<'K9> okey9 with 
                | (:? 'K1 as key1), (:? 'K2 as key2), (:? 'K3 as key3), (:? 'K4 as key4), (:? 'K5 as key5), (:? 'K6 as key6), (:? 'K7 as key7), (:? 'K8 as key8), (:? 'K9 as key9) ->
                    match map |> Map.tryFind (key1, key2, key3, key4, key5, key6, key7, key8, key9) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false false refKey proxys)
                | _ -> proxys.failed

        static member find10<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'K8,'K9,'K10,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison and 'K8: comparison and 'K9: comparison and 'K10: comparison> 
            (map: Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8*'K9*'K10,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) (okey3: obj) (okey4: obj) (okey5: obj) (okey6: obj) (okey7: obj) (okey8: obj) (okey9: obj) (okey10: obj)
            : obj =
                match GenFn.cast<'K1> okey1, GenFn.cast<'K2> okey2, GenFn.cast<'K3> okey3, GenFn.cast<'K4> okey4, GenFn.cast<'K5> okey5, GenFn.cast<'K6> okey6, GenFn.cast<'K7> okey7, GenFn.cast<'K8> okey8, GenFn.cast<'K9> okey9, GenFn.cast<'K10> okey10 with 
                | (:? 'K1 as key1), (:? 'K2 as key2), (:? 'K3 as key3), (:? 'K4 as key4), (:? 'K5 as key5), (:? 'K6 as key6), (:? 'K7 as key7), (:? 'K8 as key8), (:? 'K9 as key9), (:? 'K10 as key10) ->
                    match map |> Map.tryFind (key1, key2, key3, key4, key5, key6, key7, key8, key9, key10) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false false refKey proxys)
                | _ -> proxys.failed

        // -----------------------------------
        // -- Construction functions
        // -----------------------------------

        /// Builds a Map<'K1,'V> key-value pairs map.
        static member map1<'K1,'V when 'K1: comparison> (keys1: 'K1[]) (values: 'V[]) 
            : Map<'K1,'V> =
            zip keys1 values |> Map.ofArray

        static member omap1<'K1,'V when 'K1: comparison> (okeys1: obj[]) (ovalues: obj[]) 
            : Map<'K1,'V> =
            let fcast (ok1: obj) (oval: obj) = 
                match ok1, oval with
                | (:? 'K1 as k), (:? 'V as v) -> Some (k, v)
                | _ -> None
            let kvs = Array.map2 fcast okeys1 ovalues |> Array.choose id
            kvs  |> Map.ofArray

        /// Builds a Map<'K1*'K2,'V> key-value pairs map.
        static member map2<'K1,'K2,'V when 'K1: comparison and 'K2: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (values: 'V[]) 
            : Map<'K1*'K2,'V> =
            zip (zip keys1 keys2) values |> Map.ofArray

        static member omap2<'K1,'K2,'V when 'K1: comparison and 'K2: comparison> 
            (okeys1: obj[]) (okeys2: obj[]) (ovalues: obj[]) 
            : Map<'K1*'K2,'V> =
            let fcast (ok1: obj) (ok2: obj) (oval: obj) = 
                match ok1, ok2, oval with
                | (:? 'K1 as k1), (:? 'K2 as k2), (:? 'V as v) -> Some ((k1, k2), v)
                | _ -> None
            let kvs = Array.map3 fcast okeys1 okeys2 ovalues |> Array.choose id
            kvs |> Map.ofArray

        /// Builds a Map<'K1*'K2,'V[]> key-value pairs map.
        /// Values for a given key will be combined into one array.
        static member map2s<'K1,'K2,'V when 'K1: comparison and 'K2: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (values: 'V[]) 
            : Map<'K1*'K2,'V[]> =
            let kvs' = zip (zip keys1 keys2) values
            let kvs = kvs' |> Array.groupBy (fun (k,v) -> k) |> Array.map (fun (k,kvx) -> (k, kvx |> Array.map snd))
            kvs |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3,'V> key-value pairs map.
        static member map3<'K1,'K2,'K3,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3,'V> =
            zip (Toolbox.Array.zip3 keys1 keys2 keys3) values |> Map.ofArray

        static member omap3<'K1,'K2,'K3,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison> 
            (okeys1: obj[]) (okeys2: obj[]) (okeys3: obj[]) (ovalues: obj[]) 
            : Map<'K1*'K2*'K3,'V> =
            let fcast (okvs: obj*obj*obj*obj) =
                let (ok1, ok2, ok3, oval) = okvs
                match ok1, ok2, ok3, oval with
                | (:? 'K1 as k1), (:? 'K2 as k2), (:? 'K3 as k3), (:? 'V as v) -> Some ((k1, k2, k3), v)
                | _ -> None
            let kvs = zip4 okeys1 okeys2 okeys3 ovalues |> Array.choose fcast
            kvs |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3,'V[]> key-value pairs map.
        /// Values for a given key will be combined into one array.
        static member map3s<'K1,'K2,'K3,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3,'V[]> =
            let kvs' = zip (Toolbox.Array.zip3 keys1 keys2 keys3) values
            let kvs = kvs' |> Array.groupBy (fun (k,v) -> k) |> Array.map (fun (k,kvx) -> (k, kvx |> Array.map snd))
            kvs |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4,'V> key-value pairs map.
        static member map4<'K1,'K2,'K3,'K4,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3*'K4,'V> =
            zip (Toolbox.Array.zip4 keys1 keys2 keys3 keys4) values |> Map.ofArray

        static member omap4<'K1,'K2,'K3,'K4,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison> 
            (okeys1: obj[]) (okeys2: obj[]) (okeys3: obj[]) (okeys4: obj[]) (ovalues: obj[]) 
            : Map<'K1*'K2*'K3*'K4,'V> =
            let fcast (okvs: obj*obj*obj*obj*obj) =
                let (ok1, ok2, ok3, ok4, oval) = okvs
                match ok1, ok2, ok3, ok4, oval with
                | (:? 'K1 as k1), (:? 'K2 as k2), (:? 'K3 as k3), (:? 'K4 as k4), (:? 'V as v) -> Some ((k1, k2, k3, k4), v)
                | _ -> None
            let kvs = zip5 okeys1 okeys2 okeys3 okeys4 ovalues |> Array.choose fcast
            kvs |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4,'V[]> key-value pairs map.
        /// Values for a given key will be combined into one array.
        static member map4s<'K1,'K2,'K3,'K4,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3*'K4,'V[]> =
            let kvs' = zip (Toolbox.Array.zip4 keys1 keys2 keys3 keys4) values
            let kvs = kvs' |> Array.groupBy (fun (k,v) -> k) |> Array.map (fun (k,kvx) -> (k, kvx |> Array.map snd))
            kvs |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5,'V> key-value pairs map.
        static member map5<'K1,'K2,'K3,'K4,'K5,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3*'K4*'K5,'V> =
            zip (Toolbox.Array.zip5 keys1 keys2 keys3 keys4 keys5) values |> Map.ofArray

        static member omap5<'K1,'K2,'K3,'K4,'K5,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison> 
            (okeys1: obj[]) (okeys2: obj[]) (okeys3: obj[]) (okeys4: obj[]) (okeys5: obj[]) (ovalues: obj[]) 
            : Map<'K1*'K2*'K3*'K4*'K5,'V> =
            let fcast (okvs: obj*obj*obj*obj*obj*obj) =
                let (ok1, ok2, ok3, ok4, ok5, oval) = okvs
                match ok1, ok2, ok3, ok4, ok5, oval with
                | (:? 'K1 as k1), (:? 'K2 as k2), (:? 'K3 as k3), (:? 'K4 as k4), (:? 'K5 as k5), (:? 'V as v) -> Some ((k1, k2, k3, k4, k5), v)
                | _ -> None
            let kvs = zip6 okeys1 okeys2 okeys3 okeys4 okeys5 ovalues |> Array.choose fcast
            kvs |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5,'V[]> key-value pairs map.
        /// Values for a given key will be combined into one array.
        static member map5s<'K1,'K2,'K3,'K4,'K5,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3*'K4*'K5,'V[]> =
            let kvs' = zip (Toolbox.Array.zip5 keys1 keys2 keys3 keys4 keys5) values
            let kvs = kvs' |> Array.groupBy (fun (k,v) -> k) |> Array.map (fun (k,kvx) -> (k, kvx |> Array.map snd))
            kvs |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5*'K6,'V> key-value pairs map.
        static member map6<'K1,'K2,'K3,'K4,'K5,'K6,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (keys6: 'K6[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3*'K4*'K5*'K6,'V> =
            zip (Toolbox.Array.zip6 keys1 keys2 keys3 keys4 keys5 keys6) values |> Map.ofArray

        static member omap6<'K1,'K2,'K3,'K4,'K5,'K6,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison> 
            (okeys1: obj[]) (okeys2: obj[]) (okeys3: obj[]) (okeys4: obj[]) (okeys5: obj[]) (okeys6: obj[]) (ovalues: obj[]) 
            : Map<'K1*'K2*'K3*'K4*'K5*'K6,'V> =
            let fcast (okvs: obj*obj*obj*obj*obj*obj*obj) =
                let (ok1, ok2, ok3, ok4, ok5, ok6, oval) = okvs
                match ok1, ok2, ok3, ok4, ok5, ok6, oval with
                | (:? 'K1 as k1), (:? 'K2 as k2), (:? 'K3 as k3), (:? 'K4 as k4), (:? 'K5 as k5), (:? 'K6 as k6), (:? 'V as v) -> Some ((k1, k2, k3, k4, k5, k6), v)
                | _ -> None
            let kvs = zip7 okeys1 okeys2 okeys3 okeys4 okeys5 okeys6 ovalues |> Array.choose fcast
            kvs |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5*'K6,'V[]> key-value pairs map.
        /// Values for a given key will be combined into one array.
        static member map6s<'K1,'K2,'K3,'K4,'K5,'K6,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (keys6: 'K6[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3*'K4*'K5*'K6,'V[]> =
            let kvs' = zip (Toolbox.Array.zip6 keys1 keys2 keys3 keys4 keys5 keys6) values
            let kvs = kvs' |> Array.groupBy (fun (k,v) -> k) |> Array.map (fun (k,kvx) -> (k, kvx |> Array.map snd))
            kvs |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7,'V> key-value pairs map.
        static member map7<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (keys6: 'K6[]) (keys7: 'K7[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7,'V> =
            zip (Toolbox.Array.zip7 keys1 keys2 keys3 keys4 keys5 keys6 keys7) values |> Map.ofArray

        static member omap7<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison> 
            (okeys1: obj[]) (okeys2: obj[]) (okeys3: obj[]) (okeys4: obj[]) (okeys5: obj[]) (okeys6: obj[]) (okeys7: obj[]) (ovalues: obj[]) 
            : Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7,'V> =
            let fcast (okvs: obj*obj*obj*obj*obj*obj*obj*obj) =
                let (ok1, ok2, ok3, ok4, ok5, ok6, ok7, oval) = okvs
                match ok1, ok2, ok3, ok4, ok5, ok6, ok7, oval with
                | (:? 'K1 as k1), (:? 'K2 as k2), (:? 'K3 as k3), (:? 'K4 as k4), (:? 'K5 as k5), (:? 'K6 as k6), (:? 'K7 as k7), (:? 'V as v) -> Some ((k1, k2, k3, k4, k5, k6, k7), v)
                | _ -> None
            let kvs = zip8 okeys1 okeys2 okeys3 okeys4 okeys5 okeys6 okeys7 ovalues |> Array.choose fcast
            kvs |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7,'V[]> key-value pairs map.
        /// Values for a given key will be combined into one array.
        static member map7s<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (keys6: 'K6[]) (keys7: 'K7[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7,'V[]> =
            let kvs' = zip (Toolbox.Array.zip7 keys1 keys2 keys3 keys4 keys5 keys6 keys7) values
            let kvs = kvs' |> Array.groupBy (fun (k,v) -> k) |> Array.map (fun (k,kvx) -> (k, kvx |> Array.map snd))
            kvs |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8,'V> key-value pairs map.
        static member map8<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'K8,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison and 'K8: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (keys6: 'K6[]) (keys7: 'K7[]) (keys8: 'K8[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8,'V> =
            zip (Toolbox.Array.zip8 keys1 keys2 keys3 keys4 keys5 keys6 keys7 keys8) values |> Map.ofArray

        static member omap8<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'K8,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison and 'K8: comparison> 
            (okeys1: obj[]) (okeys2: obj[]) (okeys3: obj[]) (okeys4: obj[]) (okeys5: obj[]) (okeys6: obj[]) (okeys7: obj[]) (okeys8: obj[]) (ovalues: obj[]) 
            : Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8,'V> =
            let fcast (okvs: obj*obj*obj*obj*obj*obj*obj*obj*obj) =
                let (ok1, ok2, ok3, ok4, ok5, ok6, ok7, ok8, oval) = okvs
                match ok1, ok2, ok3, ok4, ok5, ok6, ok7, ok8, oval with
                | (:? 'K1 as k1), (:? 'K2 as k2), (:? 'K3 as k3), (:? 'K4 as k4), (:? 'K5 as k5), (:? 'K6 as k6), (:? 'K7 as k7), (:? 'K8 as k8), (:? 'V as v) -> Some ((k1, k2, k3, k4, k5, k6, k7, k8), v)
                | _ -> None
            let kvs = zip9 okeys1 okeys2 okeys3 okeys4 okeys5 okeys6 okeys7 okeys8 ovalues |> Array.choose fcast
            kvs |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8,'V[]> key-value pairs map.
        /// Values for a given key will be combined into one array.
        static member map8s<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'K8,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison and 'K8: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (keys6: 'K6[]) (keys7: 'K7[]) (keys8: 'K8[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8,'V[]> =
            let kvs' = zip (Toolbox.Array.zip8 keys1 keys2 keys3 keys4 keys5 keys6 keys7 keys8) values
            let kvs = kvs' |> Array.groupBy (fun (k,v) -> k) |> Array.map (fun (k,kvx) -> (k, kvx |> Array.map snd))
            kvs |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8*'K9,'V> key-value pairs map.
        static member map9<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'K8,'K9,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison and 'K8: comparison and 'K9: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (keys6: 'K6[]) (keys7: 'K7[]) (keys8: 'K8[]) (keys9: 'K9[]) (values: 'V[])
            : Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8*'K9,'V> =
            zip (Toolbox.Array.zip9 keys1 keys2 keys3 keys4 keys5 keys6 keys7 keys8 keys9) values |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8*'K9,'V[]> key-value pairs map.
        /// Values for a given key will be combined into one array.
        static member map9s<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'K8,'K9,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison and 'K8: comparison and 'K9: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (keys6: 'K6[]) (keys7: 'K7[]) (keys8: 'K8[]) (keys9: 'K9[]) (values: 'V[])
            : Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8*'K9,'V[]> =
            let kvs' = zip (Toolbox.Array.zip9 keys1 keys2 keys3 keys4 keys5 keys6 keys7 keys8 keys9) values
            let kvs = kvs' |> Array.groupBy (fun (k,v) -> k) |> Array.map (fun (k,kvx) -> (k, kvx |> Array.map snd))
            kvs |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8*'K9*'K10,'V> key-value pairs map.
        static member map10<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'K8,'K9,'K10,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison and 'K8: comparison and 'K9: comparison and 'K10: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (keys6: 'K6[]) (keys7: 'K7[]) (keys8: 'K8[]) (keys9: 'K9[]) (keys10: 'K10[]) (values: 'V[])
            : Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8*'K9*'K10,'V> =
            zip (Toolbox.Array.zip10 keys1 keys2 keys3 keys4 keys5 keys6 keys7 keys8 keys9 keys10) values |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8*'K9*'K10,'V[]> key-value pairs map.
        /// Values for a given key will be combined into one array.
        static member map10s<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'K8,'K9,'K10,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison and 'K8: comparison and 'K9: comparison and 'K10: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (keys6: 'K6[]) (keys7: 'K7[]) (keys8: 'K8[]) (keys9: 'K9[]) (keys10: 'K10[]) (values: 'V[])
            : Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8*'K9*'K10,'V[]> =
            let kvs' = zip (Toolbox.Array.zip10 keys1 keys2 keys3 keys4 keys5 keys6 keys7 keys8 keys9 keys10) values
            let kvs = kvs' |> Array.groupBy (fun (k,v) -> k) |> Array.map (fun (k,kvx) -> (k, kvx |> Array.map snd))
            kvs |> Map.ofArray

        /// Returns the number of kvp in the Map's object.
        static member pool<'K,'V when 'K: comparison> (omaps: obj[]) : Map<'K,'V> =
            let cast (omap: obj) = match omap with | :? Map<'K,'V> as map -> Some map | _ -> None
            let maps = omaps |> Array.choose cast
            let kvps = maps |> Array.collect (fun mp -> [| for kvp in mp -> kvp.Key, kvp.Value  |])
            kvps |> Map.ofArray
        
        // BOILER PLATE. ALL CASES ARE NOT IMPLEMENTED YET
        // START HERE

        /// Builds a Map<'KV1*'KH1,'V> key-value pairs map.
        static member mapV1H1<'KV1,'KH1,'V when 'KV1: comparison and 'KH1: comparison> 
            (vKeys1: 'KV1[]) (hKeys1: 'KH1[]) (values: 'V[,])
            : Map<'KV1*'KH1,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            if (vKeys1.Length = len1) && (hKeys1.Length = len2) then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> hKeys1.[j], values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun (hkey, vals1D) -> 
                            let keys = vKeys1 |> Array.map (fun vkey -> (vkey, hkey))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KH1*'KV1,'V> key-value pairs map. 
        /// Similar to mapV1H1 but with HKeys placed first.
        static member mapH1V1<'KV1,'KH1,'V when 'KV1: comparison and 'KH1: comparison> 
            (vKeys1: 'KV1[]) (hKeys1: 'KH1[]) (values: 'V[,])
            : Map<'KH1*'KV1,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            if (vKeys1.Length = len1) && (hKeys1.Length = len2) then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> hKeys1.[j], values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun (hkey, vals1D) -> 
                            let keys = vKeys1 |> Array.map (fun vkey -> (hkey, vkey))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KH1*'KH2,'V> key-value pairs map.
        static member mapV1H2<'KV1,'KH1,'KH2,'V when 'KV1: comparison and 'KH1: comparison and 'KH2: comparison> 
            (vKeys1: 'KV1[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (values: 'V[,])
            : Map<'KV1*'KH1*'KH2,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if (vKeys1.Length = len1) && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2), vals1D) -> 
                            let keys = vKeys1 |> Array.map (fun vkey -> (vkey, hkey1, hkey2))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KH1*'KH2,'V> key-value pairs map.
        /// Similar to mapV1H2 but with HKeys placed first.
        static member mapH2V1<'KV1,'KH1,'KH2,'V when 'KV1: comparison and 'KH1: comparison and 'KH2: comparison> 
            (vKeys1: 'KV1[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (values: 'V[,])
            : Map<'KH1*'KH2*'KV1,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if (vKeys1.Length = len1) && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2), vals1D) -> 
                            let keys = vKeys1 |> Array.map (fun vkey -> (hkey1, hkey2, vkey))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KH1*'KH2*'KH3,'V> key-value pairs map.
        static member mapV1H3<'KV1,'KH1,'KH2,'KH3,'V when 'KV1: comparison and 'KH1: comparison and 'KH2: comparison and 'KH3: comparison> 
            (vKeys1: 'KV1[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (hKeys3: 'KH3[]) (values: 'V[,])
            : Map<'KV1*'KH1*'KH2*'KH3,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length; hKeys3.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if (vKeys1.Length = len1) && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j], hKeys3.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2, hkey3), vals1D) -> 
                            let keys = vKeys1 |> Array.map (fun vkey -> (vkey, hkey1, hkey2, hkey3))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KH1*'KH2*'KH3,'V> key-value pairs map.
        /// Similar to mapV1H3 but with HKeys placed first.
        static member mapH3V1<'KV1,'KH1,'KH2,'KH3,'V when 'KV1: comparison and 'KH1: comparison and 'KH2: comparison and 'KH3: comparison> 
            (vKeys1: 'KV1[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (hKeys3: 'KH3[]) (values: 'V[,])
            : Map<'KH1*'KH2*'KH3*'KV1,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length; hKeys3.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if (vKeys1.Length = len1) && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j], hKeys3.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2, hkey3), vals1D) -> 
                            let keys = vKeys1 |> Array.map (fun vkey -> (hkey1, hkey2, hkey3, vkey))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KH1*'KH2*'KH3*'KH4,'V> key-value pairs map.
        static member mapV1H4<'KV1,'KH1,'KH2,'KH3,'KH4,'V when 'KV1: comparison and 'KH1: comparison and 'KH2: comparison and 'KH3: comparison and 'KH4: comparison> 
            (vKeys1: 'KV1[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (hKeys3: 'KH3[]) (hKeys4: 'KH4[]) (values: 'V[,])
            : Map<'KV1*'KH1*'KH2*'KH3*'KH4,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length; hKeys3.Length; hKeys4.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if (vKeys1.Length = len1) && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j], hKeys3.[j], hKeys4.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2, hkey3, hkey4), vals1D) -> 
                            let keys = vKeys1 |> Array.map (fun vkey -> (vkey, hkey1, hkey2, hkey3, hkey4))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KH1*'KH2*'KH3*'KH4*'KV1,'V> key-value pairs map.
        /// Similar to mapV1H4 but with HKeys placed first.
        static member mapH4V1<'KV1,'KH1,'KH2,'KH3,'KH4,'V when 'KV1: comparison and 'KH1: comparison and 'KH2: comparison and 'KH3: comparison and 'KH4: comparison> 
            (vKeys1: 'KV1[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (hKeys3: 'KH3[]) (hKeys4: 'KH4[]) (values: 'V[,])
            : Map<'KH1*'KH2*'KH3*'KH4*'KV1,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length; hKeys3.Length; hKeys4.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if (vKeys1.Length = len1) && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j], hKeys3.[j], hKeys4.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2, hkey3, hkey4), vals1D) -> 
                            let keys = vKeys1 |> Array.map (fun vkey -> (hkey1, hkey2, hkey3, hkey4, vkey))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KH1*'KH2*'KH3*'KH4*'KH5,'V> key-value pairs map.
        static member mapV1H5<'KV1,'KH1,'KH2,'KH3,'KH4,'KH5,'V when 'KV1: comparison and 'KH1: comparison and 'KH2: comparison and 'KH3: comparison and 'KH4: comparison and 'KH5: comparison> 
            (vKeys1: 'KV1[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (hKeys3: 'KH3[]) (hKeys4: 'KH4[]) (hKeys5: 'KH5[]) (values: 'V[,])
            : Map<'KV1*'KH1*'KH2*'KH3*'KH4*'KH5,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length; hKeys3.Length; hKeys4.Length; hKeys5.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if (vKeys1.Length = len1) && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j], hKeys3.[j], hKeys4.[j], hKeys5.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2, hkey3, hkey4, hkey5), vals1D) -> 
                            let keys = vKeys1 |> Array.map (fun vkey -> (vkey, hkey1, hkey2, hkey3, hkey4, hkey5))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KH1*'KH2*'KH3*'KH4*'KH5,'V> key-value pairs map.
        /// Similar to mapV1H5 but with HKeys placed first.
        static member mapH5V1<'KV1,'KH1,'KH2,'KH3,'KH4,'KH5,'V when 'KV1: comparison and 'KH1: comparison and 'KH2: comparison and 'KH3: comparison and 'KH4: comparison and 'KH5: comparison> 
            (vKeys1: 'KV1[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (hKeys3: 'KH3[]) (hKeys4: 'KH4[]) (hKeys5: 'KH5[]) (values: 'V[,])
            : Map<'KH1*'KH2*'KH3*'KH4*'KH5*'KV1,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length; hKeys3.Length; hKeys4.Length; hKeys5.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if (vKeys1.Length = len1) && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j], hKeys3.[j], hKeys4.[j], hKeys5.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2, hkey3, hkey4, hkey5), vals1D) -> 
                            let keys = vKeys1 |> Array.map (fun vkey -> (hkey1, hkey2, hkey3, hkey4, hkey5, vkey))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KH1*'KH2*'KH3*'KH4*'KH5*'KH6,'V> key-value pairs map.
        static member mapV1H6<'KV1,'KH1,'KH2,'KH3,'KH4,'KH5,'KH6,'V when 'KV1: comparison and 'KH1: comparison and 'KH2: comparison and 'KH3: comparison and 'KH4: comparison and 'KH5: comparison and 'KH6: comparison> 
            (vKeys1: 'KV1[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (hKeys3: 'KH3[]) (hKeys4: 'KH4[]) (hKeys5: 'KH5[]) (hKeys6: 'KH6[]) (values: 'V[,])
            : Map<'KV1*'KH1*'KH2*'KH3*'KH4*'KH5*'KH6,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length; hKeys3.Length; hKeys4.Length; hKeys5.Length; hKeys6.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if (vKeys1.Length = len1) && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j], hKeys3.[j], hKeys4.[j], hKeys5.[j], hKeys6.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2, hkey3, hkey4, hkey5, hkey6), vals1D) -> 
                            let keys = vKeys1 |> Array.map (fun vkey -> (vkey, hkey1, hkey2, hkey3, hkey4, hkey5, hkey6))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KH1*'KH2*'KH3*'KH4*'KH5*'KH6*'KV1,'V> key-value pairs map.
        /// Similar to mapV1H6 but with HKeys placed first.
        static member mapH6V1<'KV1,'KH1,'KH2,'KH3,'KH4,'KH5,'KH6,'V when 'KV1: comparison and 'KH1: comparison and 'KH2: comparison and 'KH3: comparison and 'KH4: comparison and 'KH5: comparison and 'KH6: comparison> 
            (vKeys1: 'KV1[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (hKeys3: 'KH3[]) (hKeys4: 'KH4[]) (hKeys5: 'KH5[]) (hKeys6: 'KH6[]) (values: 'V[,])
            : Map<'KH1*'KH2*'KH3*'KH4*'KH5*'KH6*'KV1,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length; hKeys3.Length; hKeys4.Length; hKeys5.Length; hKeys6.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if (vKeys1.Length = len1) && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j], hKeys3.[j], hKeys4.[j], hKeys5.[j], hKeys6.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2, hkey3, hkey4, hkey5, hkey6), vals1D) -> 
                            let keys = vKeys1 |> Array.map (fun vkey -> (hkey1, hkey2, hkey3, hkey4, hkey5, hkey6, vkey))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KV2*'KH1,'V> key-value pairs map.
        static member mapV2H1<'KV1,'KV2,'KH1,'V when 'KV1: comparison and 'KV2: comparison and 'KH1: comparison> 
            (vKeys1: 'KV1[]) (vKeys2: 'KV2[]) (hKeys1: 'KH1[]) (values: 'V[,])
            : Map<'KV1*'KV2*'KH1,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all vKeys arrays must have the same length as values' number of rows
            let testVLen = let lens = [| vKeys1.Length; vKeys2.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len1

            if testVLen && (hKeys1.Length = len2) then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> hKeys1.[j], values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun (hkey, vals1D) -> 
                            let keys = 
                                Array.zip vKeys1 vKeys2 
                                |> Array.map (fun (vkey1, vkey2) -> (vkey1, vkey2, hkey))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KV2*'KH1*'KH2,'V> key-value pairs map.
        static member mapV2H2<'KV1,'KV2,'KH1,'KH2,'V when 'KV1: comparison and 'KV2: comparison and 'KH1: comparison and 'KH2: comparison> 
            (vKeys1: 'KV1[]) (vKeys2: 'KV2[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (values: 'V[,])
            : Map<'KV1*'KV2*'KH1*'KH2,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all vKeys arrays must have the same length as values' number of rows
            let testVLen = let lens = [| vKeys1.Length; vKeys2.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len1
            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if testVLen && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2), vals1D) -> 
                            let keys = 
                                Array.zip vKeys1 vKeys2 
                                |> Array.map (fun (vkey1, vkey2) -> (vkey1, vkey2, hkey1, hkey2))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KV2*'KH1*'KH2*'KH3,'V> key-value pairs map.
        static member mapV2H3<'KV1,'KV2,'KH1,'KH2,'KH3,'V when 'KV1: comparison and 'KV2: comparison and 'KH1: comparison and 'KH2: comparison and 'KH3: comparison> 
            (vKeys1: 'KV1[]) (vKeys2: 'KV2[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (hKeys3: 'KH3[]) (values: 'V[,])
            : Map<'KV1*'KV2*'KH1*'KH2*'KH3,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all vKeys arrays must have the same length as values' number of rows
            let testVLen = let lens = [| vKeys1.Length; vKeys2.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len1
            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length; hKeys3.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if testVLen && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j], hKeys3.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2, hkey3), vals1D) -> 
                            let keys = 
                                Array.zip vKeys1 vKeys2 
                                |> Array.map (fun (vkey1, vkey2) -> (vkey1, vkey2, hkey1, hkey2, hkey3))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KV2*'KH1*'KH2*'KH3*'KH4,'V> key-value pairs map.
        static member mapV2H4<'KV1,'KV2,'KH1,'KH2,'KH3,'KH4,'V when 'KV1: comparison and 'KV2: comparison and 'KH1: comparison and 'KH2: comparison and 'KH3: comparison and 'KH4: comparison> 
            (vKeys1: 'KV1[]) (vKeys2: 'KV2[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (hKeys3: 'KH3[]) (hKeys4: 'KH4[]) (values: 'V[,])
            : Map<'KV1*'KV2*'KH1*'KH2*'KH3*'KH4,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all vKeys arrays must have the same length as values' number of rows
            let testVLen = let lens = [| vKeys1.Length; vKeys2.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len1
            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length; hKeys3.Length; hKeys4.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if testVLen && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j], hKeys3.[j], hKeys4.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2, hkey3, hkey4), vals1D) -> 
                            let keys = 
                                Array.zip vKeys1 vKeys2 
                                |> Array.map (fun (vkey1, vkey2) -> (vkey1, vkey2, hkey1, hkey2, hkey3, hkey4))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KV2*'KH1*'KH2*'KH3*'KH4*'KH5,'V> key-value pairs map.
        static member mapV2H5<'KV1,'KV2,'KH1,'KH2,'KH3,'KH4,'KH5,'V when 'KV1: comparison and 'KV2: comparison and 'KH1: comparison and 'KH2: comparison and 'KH3: comparison and 'KH4: comparison and 'KH5: comparison> 
            (vKeys1: 'KV1[]) (vKeys2: 'KV2[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (hKeys3: 'KH3[]) (hKeys4: 'KH4[]) (hKeys5: 'KH5[]) (values: 'V[,])
            : Map<'KV1*'KV2*'KH1*'KH2*'KH3*'KH4*'KH5,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all vKeys arrays must have the same length as values' number of rows
            let testVLen = let lens = [| vKeys1.Length; vKeys2.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len1
            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length; hKeys3.Length; hKeys4.Length; hKeys5.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if testVLen && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j], hKeys3.[j], hKeys4.[j], hKeys5.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2, hkey3, hkey4, hkey5), vals1D) -> 
                            let keys = 
                                Array.zip vKeys1 vKeys2 
                                |> Array.map (fun (vkey1, vkey2) -> (vkey1, vkey2, hkey1, hkey2, hkey3, hkey4, hkey5))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KV2*'KH1*'KH2*'KH3*'KH4*'KH5*'KH6,'V> key-value pairs map.
        static member mapV2H6<'KV1,'KV2,'KH1,'KH2,'KH3,'KH4,'KH5,'KH6,'V when 'KV1: comparison and 'KV2: comparison and 'KH1: comparison and 'KH2: comparison and 'KH3: comparison and 'KH4: comparison and 'KH5: comparison and 'KH6: comparison> 
            (vKeys1: 'KV1[]) (vKeys2: 'KV2[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (hKeys3: 'KH3[]) (hKeys4: 'KH4[]) (hKeys5: 'KH5[]) (hKeys6: 'KH6[]) (values: 'V[,])
            : Map<'KV1*'KV2*'KH1*'KH2*'KH3*'KH4*'KH5*'KH6,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all vKeys arrays must have the same length as values' number of rows
            let testVLen = let lens = [| vKeys1.Length; vKeys2.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len1
            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length; hKeys3.Length; hKeys4.Length; hKeys5.Length; hKeys6.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if testVLen && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j], hKeys3.[j], hKeys4.[j], hKeys5.[j], hKeys6.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2, hkey3, hkey4, hkey5, hkey6), vals1D) -> 
                            let keys = 
                                Array.zip vKeys1 vKeys2 
                                |> Array.map (fun (vkey1, vkey2) -> (vkey1, vkey2, hkey1, hkey2, hkey3, hkey4, hkey5, hkey6))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KV2*'KV3*'KH1,'V> key-value pairs map.
        static member mapV3H1<'KV1,'KV2,'KV3,'KH1,'V when 'KV1: comparison and 'KV2: comparison and 'KV3: comparison and 'KH1: comparison> 
            (vKeys1: 'KV1[]) (vKeys2: 'KV2[]) (vKeys3: 'KV3[]) (hKeys1: 'KH1[]) (values: 'V[,])
            : Map<'KV1*'KV2*'KV3*'KH1,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all vKeys arrays must have the same length as values' number of rows
            let testVLen = let lens = [| vKeys1.Length; vKeys2.Length; vKeys3.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len1

            if testVLen && (hKeys1.Length = len2) then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> hKeys1.[j], values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun (hkey, vals1D) -> 
                            let keys = 
                                Array.zip3 vKeys1 vKeys2 vKeys3 
                                |> Array.map (fun (vkey1, vkey2, vkey3) -> (vkey1, vkey2, vkey3, hkey))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KV2*'KV3*'KH1*'KH2,'V> key-value pairs map.
        static member mapV3H2<'KV1,'KV2,'KV3,'KH1,'KH2,'V when 'KV1: comparison and 'KV2: comparison and 'KV3: comparison and 'KH1: comparison and 'KH2: comparison> 
            (vKeys1: 'KV1[]) (vKeys2: 'KV2[]) (vKeys3: 'KV3[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (values: 'V[,])
            : Map<'KV1*'KV2*'KV3*'KH1*'KH2,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all vKeys arrays must have the same length as values' number of rows
            let testVLen = let lens = [| vKeys1.Length; vKeys2.Length; vKeys3.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len1
            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if testVLen && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2), vals1D) -> 
                            let keys = 
                                Array.zip3 vKeys1 vKeys2 vKeys3 
                                |> Array.map (fun (vkey1, vkey2, vkey3) -> (vkey1, vkey2, vkey3, hkey1, hkey2))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KV2*'KV3*'KH1*'KH2*'KH3,'V> key-value pairs map.
        static member mapV3H3<'KV1,'KV2,'KV3,'KH1,'KH2,'KH3,'V when 'KV1: comparison and 'KV2: comparison and 'KV3: comparison and 'KH1: comparison and 'KH2: comparison and 'KH3: comparison> 
            (vKeys1: 'KV1[]) (vKeys2: 'KV2[]) (vKeys3: 'KV3[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (hKeys3: 'KH3[]) (values: 'V[,])
            : Map<'KV1*'KV2*'KV3*'KH1*'KH2*'KH3,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all vKeys arrays must have the same length as values' number of rows
            let testVLen = let lens = [| vKeys1.Length; vKeys2.Length; vKeys3.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len1
            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length; hKeys3.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if testVLen && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j], hKeys3.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2, hkey3), vals1D) -> 
                            let keys = 
                                Array.zip3 vKeys1 vKeys2 vKeys3 
                                |> Array.map (fun (vkey1, vkey2, vkey3) -> (vkey1, vkey2, vkey3, hkey1, hkey2, hkey3))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KV2*'KV3*'KH1*'KH2*'KH3*'KH4,'V> key-value pairs map.
        static member mapV3H4<'KV1,'KV2,'KV3,'KH1,'KH2,'KH3,'KH4,'V when 'KV1: comparison and 'KV2: comparison and 'KV3: comparison and 'KH1: comparison and 'KH2: comparison and 'KH3: comparison and 'KH4: comparison> 
            (vKeys1: 'KV1[]) (vKeys2: 'KV2[]) (vKeys3: 'KV3[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (hKeys3: 'KH3[]) (hKeys4: 'KH4[]) (values: 'V[,])
            : Map<'KV1*'KV2*'KV3*'KH1*'KH2*'KH3*'KH4,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all vKeys arrays must have the same length as values' number of rows
            let testVLen = let lens = [| vKeys1.Length; vKeys2.Length; vKeys3.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len1
            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length; hKeys3.Length; hKeys4.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if testVLen && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j], hKeys3.[j], hKeys4.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2, hkey3, hkey4), vals1D) -> 
                            let keys = 
                                Array.zip3 vKeys1 vKeys2 vKeys3 
                                |> Array.map (fun (vkey1, vkey2, vkey3) -> (vkey1, vkey2, vkey3, hkey1, hkey2, hkey3, hkey4))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KV2*'KV3*'KH1*'KH2*'KH3*'KH4*'KH5,'V> key-value pairs map.
        static member mapV3H5<'KV1,'KV2,'KV3,'KH1,'KH2,'KH3,'KH4,'KH5,'V when 'KV1: comparison and 'KV2: comparison and 'KV3: comparison and 'KH1: comparison and 'KH2: comparison and 'KH3: comparison and 'KH4: comparison and 'KH5: comparison> 
            (vKeys1: 'KV1[]) (vKeys2: 'KV2[]) (vKeys3: 'KV3[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (hKeys3: 'KH3[]) (hKeys4: 'KH4[]) (hKeys5: 'KH5[]) (values: 'V[,])
            : Map<'KV1*'KV2*'KV3*'KH1*'KH2*'KH3*'KH4*'KH5,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all vKeys arrays must have the same length as values' number of rows
            let testVLen = let lens = [| vKeys1.Length; vKeys2.Length; vKeys3.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len1
            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length; hKeys3.Length; hKeys4.Length; hKeys5.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if testVLen && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j], hKeys3.[j], hKeys4.[j], hKeys5.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2, hkey3, hkey4, hkey5), vals1D) -> 
                            let keys = 
                                Array.zip3 vKeys1 vKeys2 vKeys3 
                                |> Array.map (fun (vkey1, vkey2, vkey3) -> (vkey1, vkey2, vkey3, hkey1, hkey2, hkey3, hkey4, hkey5))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        /// Builds a Map<'KV1*'KV2*'KV3*'KH1*'KH2*'KH3*'KH4*'KH5*'KH6,'V> key-value pairs map.
        static member mapV3H6<'KV1,'KV2,'KV3,'KH1,'KH2,'KH3,'KH4,'KH5,'KH6,'V when 'KV1: comparison and 'KV2: comparison and 'KV3: comparison and 'KH1: comparison and 'KH2: comparison and 'KH3: comparison and 'KH4: comparison and 'KH5: comparison and 'KH6: comparison> 
            (vKeys1: 'KV1[]) (vKeys2: 'KV2[]) (vKeys3: 'KV3[]) (hKeys1: 'KH1[]) (hKeys2: 'KH2[]) (hKeys3: 'KH3[]) (hKeys4: 'KH4[]) (hKeys5: 'KH5[]) (hKeys6: 'KH6[]) (values: 'V[,])
            : Map<'KV1*'KV2*'KV3*'KH1*'KH2*'KH3*'KH4*'KH5*'KH6,'V> =
            let len1, len2 = values |> Array2D.length1, values |> Array2D.length2

            // all vKeys arrays must have the same length as values' number of rows
            let testVLen = let lens = [| vKeys1.Length; vKeys2.Length; vKeys3.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len1
            // all hKeys arrays must have the same length as values' number of columns
            let testHLen = let lens = [| hKeys1.Length; hKeys2.Length; hKeys3.Length; hKeys4.Length; hKeys5.Length; hKeys6.Length |] |> Array.distinct in (lens.Length = 1) && lens.[0] = len2

            if testVLen && testHLen then 
                let htranskeys = [| for j in 0 .. (len2 - 1) -> (hKeys1.[j], hKeys2.[j], hKeys3.[j], hKeys4.[j], hKeys5.[j], hKeys6.[j]), values.[*, j] |]
                let kvps = 
                    htranskeys
                    |> Array.collect 
                        (fun ((hkey1, hkey2, hkey3, hkey4, hkey5, hkey6), vals1D) -> 
                            let keys = 
                                Array.zip3 vKeys1 vKeys2 vKeys3 
                                |> Array.map (fun (vkey1, vkey2, vkey3) -> (vkey1, vkey2, vkey3, hkey1, hkey2, hkey3, hkey4, hkey5, hkey6))
                            Array.zip keys vals1D
                        )
                kvps |> Map.ofArray
            else
                Map.empty

        // NOT IMPLEMENTED YET : CASES mapVnHm where n > = 4 

        // BOILER PLATE. ALL CASES ARE NOT IMPLEMENTED YET
        // END HERE

    /// -----------------------------------
    /// -- Reflection functions
    /// -----------------------------------
    module RxFn =
        open Toolbox.Generics
        let genType = typeof<Map<_,_>>

        /// Collection of functions which, loosely speaking, add elements to the Registry.
        module In = 
            let MAX_ARITY_MAPN = 10 // TODO
            /// Builds a Map<'K1*'K2 .. *'KN,'V> key-value pairs map (Currently N <= 10).
            /// Output type: Map<'K1*'K2 .. *'KN,'V>
            let mapN (gtykeys: Type[]) (gtyval: Type) (keys: obj[]) (values: obj) : obj = 
                // if arity > MAX_ARITY_TOMAP then
                let gtys = Array.append gtykeys [| gtyval |]
                let args : obj[] = Array.append keys [| values |]
                let methodnm = sprintf "map%d" gtykeys.Length

                let res = invoke<GenFn> methodnm gtys args
                res

            let omapN (gtykeys: Type[]) (gtyval: Type) (keys: obj[]) (values: obj) : obj = 
                // if arity > MAX_ARITY_TOMAP then
                let gtys = Array.append gtykeys [| gtyval |]
                let args : obj[] = Array.append keys [| values |]
                let methodnm = sprintf "omap%d" gtykeys.Length

                let res = invoke<GenFn> methodnm gtys args
                res

            /// Builds a Map<'KV1*'KV2 .. *'KH1*'KH2..,'V> key-value pairs map (Currently N <= 10).
            /// Output type: Map<'KV1*'KV2 .. *'KH1*'KH2..,'V>
            let map2D (vgtykeys: Type[]) (hgtykeys: Type[]) (gtyval: Type) (vkeys: obj[]) (hkeys: obj[]) (values: obj) : obj = 
                let gtys = Array.append (Array.append vgtykeys hgtykeys) [| gtyval |]
                let args : obj[] = Array.append (Array.append vkeys hkeys) [| values |]
                let methodnm = sprintf "mapV%dH%d" vgtykeys.Length hgtykeys.Length

                let res = invoke<GenFn> methodnm gtys args
                res

            /// Merges several Map<'K*'V> objects together.
            /// Output type: Map<'KV1*'KV2 .. *'KH1*'KH2..,'V>
            let pool (xlValue: obj) : obj option =
                let methodNm = "pool"
                MRegistry.tryExtractGen1D genType xlValue |> Option.map (fun (tys, objs) -> (tys, box objs))
                |> Option.map (apply<GenFn> methodNm [||] [||]) 

        /// Collection of functions which, loosely speaking, output Registry object to Excel.
        module Out =
            /// Output type: int
            let count (regKey: string) : obj option =
                let methodNm = "count"
                MRegistry.tryExtractGen genType regKey
                |> Option.map (apply<GenFn> methodNm [||] [||])

            /// Output type: 'K[]
            let keys (regKey: string) (refKey: String) (proxys: Proxys) : obj[] option =
                let methodNm = "keys"
                MRegistry.tryExtractGen genType regKey
                |> Option.map (apply<GenFn> methodNm [||] [| refKey; proxys |])
                |> Option.map (fun o -> o :?> obj[])

            /// Output type: 'V[]
            let values (regKey: string) (unwrapOptions: bool) (refKey: String) (proxys: Proxys) : obj[] option =
                let methodNm = "values"
                MRegistry.tryExtractGen genType regKey
                |> Option.map (apply<GenFn> methodNm [||] [| unwrapOptions; refKey; proxys |])
                |> Option.map (fun o -> o :?> obj[])

            /// Output type: obj
            let find1 (regKey: string) (proxys: Proxys) (refKey: string) (okey1: obj) : obj option =
                let methodNm = "find1"
                let okey1' = Registry.MRegistry.tryExtractO okey1 |> Option.defaultValue okey1
                MRegistry.tryExtractGen genType regKey
                |> Option.map (apply<GenFn> methodNm [||] [| false; proxys; refKey; okey1' |])

            /// Output type: obj
            let findN (regKey: string) (proxys: Proxys) (refKey: string) (okeys: obj[]) : obj option =
                let args : obj[] = Array.append [| proxys; refKey |] okeys
                let methodNm = sprintf "find%d" okeys.Length
                match MRegistry.tryExtractGen genType regKey with
                | None -> None
                | Some (tys, o) -> 
                    // tys is a [| the map-object's key type, the map object's value type |]
                    if tys.Length <> 2 then
                        failwith ("SHOULD NEVER EVER BE THERE!!")  // TODO remove this tys.length always = 2 here
                    else
                        if not (FSharpType.IsTuple tys.[0]) then
                            None
                        else
                            let elemTys = FSharpType.GetTupleElements(tys.[0])
                            let genTypeRObj = (Array.append elemTys [| tys.[1] |], o)
                            apply<GenFn> methodNm [||] args genTypeRObj
                            |> Some
            
            /// Output type: obj[]
            let find1D1 (regKey: string) (proxys: Proxys) (refKey: string) (okeys1: obj[]) : obj option =
                let methodNm = "find1D1"
                MRegistry.tryExtractGen genType regKey
                |> Option.map (apply<GenFn> methodNm [||] [| proxys; refKey; okeys1 |])

module MAP_XL =
    open Registry
    open API
    open API.Out
    open type Out.Proxys

    [<ExcelFunction(Category="Map", Description="Creates a Map<'Key1*'Key2...,'Val> R-object from several 1D-arrays of keys and one 1D-array of values.")>]
    let map_ofRng
        ([<ExcelArgument(Description= "Key1 type tag: bool, date, double, doubleNaN, string or obj. Add \'#'\ prefix for optional type: #bool, #date, #double, #doubleNaN, #string or #obj.")>] k1TypeTag: string)
        ([<ExcelArgument(Description= "Key2 type tag.")>] k2TypeTag: obj)
        ([<ExcelArgument(Description= "Key3 type tag.")>] k3TypeTag: obj)
        ([<ExcelArgument(Description= "Key4 type tag.")>] k4TypeTag: obj)
        ([<ExcelArgument(Description= "Key5 type tag.")>] k5TypeTag: obj)
        ([<ExcelArgument(Description= "Key6 type tag.")>] k6TypeTag: obj)
        ([<ExcelArgument(Description= "Key7 type tag.")>] k7TypeTag: obj)
        ([<ExcelArgument(Description= "Key8 type tag.")>] k8TypeTag: obj)
        ([<ExcelArgument(Description= "Key9 type tag.")>] k9TypeTag: obj)
        ([<ExcelArgument(Description= "Key10 type tag.")>] k10TypeTag: obj)
        ([<ExcelArgument(Description= "Value type tag.")>] valueTypeTag: string)
        ([<ExcelArgument(Description= "keys1.")>] mapKeys1: obj)
        ([<ExcelArgument(Description= "keys2.")>] mapKeys2: obj)
        ([<ExcelArgument(Description= "keys3.")>] mapKeys3: obj)
        ([<ExcelArgument(Description= "keys4.")>] mapKeys4: obj)
        ([<ExcelArgument(Description= "keys5.")>] mapKeys5: obj)
        ([<ExcelArgument(Description= "keys6.")>] mapKeys6: obj)
        ([<ExcelArgument(Description= "keys7.")>] mapKeys7: obj)
        ([<ExcelArgument(Description= "keys8.")>] mapKeys8: obj)
        ([<ExcelArgument(Description= "keys9.")>] mapKeys9: obj)
        ([<ExcelArgument(Description= "keys10.")>] mapKeys10: obj)
        ([<ExcelArgument(Description= "values.")>] mapValues: obj)
        ([<ExcelArgument(Description= "[Kinds for which values are converted to Double.NaN. E.g. NA, ERR, TXT, !NUM... (comma separated). Default is none.]")>] xlKinds: obj)
        : obj  =

        // intermediary stage
        let ktag2 = In.D0.Stg.Opt.def None k2TypeTag
        let ktag3 = In.D0.Stg.Opt.def None k3TypeTag
        let ktag4 = In.D0.Stg.Opt.def None k4TypeTag
        let ktag5 = In.D0.Stg.Opt.def None k5TypeTag
        let ktag6 = In.D0.Stg.Opt.def None k6TypeTag
        let ktag7 = In.D0.Stg.Opt.def None k7TypeTag
        let ktag8 = In.D0.Stg.Opt.def None k8TypeTag
        let ktag9 = In.D0.Stg.Opt.def None k9TypeTag
        let ktag10 = In.D0.Stg.Opt.def None k10TypeTag
        let xlkinds = In.D0.Stg.def "NONE" xlKinds |> Kind.ofLabel

        // caller cell's reference ID
        let rfid = MRegistry.refID

        let gtykeys_keys_gtyvals_vals =
            match ktag2, ktag3, ktag4, ktag5, ktag6, ktag7, ktag8, ktag9, ktag10 with
            | Some ktg2, Some ktg3, Some ktg4, Some ktg5, Some ktg6, Some ktg7, Some ktg8, Some ktg9, Some ktg10 -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV xlkinds None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg2 mapKeys2
                let trykeys3 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg3 mapKeys3
                let trykeys4 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg4 mapKeys4
                let trykeys5 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg5 mapKeys5
                let trykeys6 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg6 mapKeys6
                let trykeys7 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg7 mapKeys7
                let trykeys8 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg8 mapKeys8
                let trykeys9 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg9 mapKeys9
                let trykeys10 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg10 mapKeys10
                let tryvals =  API.In.D1.Tag.Try.tryDV xlkinds None None valueTypeTag mapValues

                match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5, trykeys6, trykeys7, trykeys8, trykeys9, trykeys10, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5), Some (gtykey6, keys6), Some (gtykey7, keys7), Some (gtykey8, keys8), Some (gtykey9, keys9), Some (gtykey10, keys10), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5; gtykey6; gtykey7; gtykey8; gtykey9; gtykey10 |]
                    let keys = [| keys1; keys2; keys3; keys4; keys5; keys6; keys7; keys8; keys9; keys10 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

            | Some ktg2, Some ktg3, Some ktg4, Some ktg5, Some ktg6, Some ktg7, Some ktg8, Some ktg9, None -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV xlkinds None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg2 mapKeys2
                let trykeys3 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg3 mapKeys3
                let trykeys4 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg4 mapKeys4
                let trykeys5 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg5 mapKeys5
                let trykeys6 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg6 mapKeys6
                let trykeys7 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg7 mapKeys7
                let trykeys8 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg8 mapKeys8
                let trykeys9 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg9 mapKeys9
                let tryvals =  API.In.D1.Tag.Try.tryDV xlkinds None None valueTypeTag mapValues

                match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5, trykeys6, trykeys7, trykeys8, trykeys9, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5), Some (gtykey6, keys6), Some (gtykey7, keys7), Some (gtykey8, keys8), Some (gtykey9, keys9), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5; gtykey6; gtykey7; gtykey8; gtykey9 |]
                    let keys = [| keys1; keys2; keys3; keys4; keys5; keys6; keys7; keys8; keys9 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None
                
            | Some ktg2, Some ktg3, Some ktg4, Some ktg5, Some ktg6, Some ktg7, Some ktg8, None, None -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV xlkinds None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg2 mapKeys2
                let trykeys3 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg3 mapKeys3
                let trykeys4 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg4 mapKeys4
                let trykeys5 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg5 mapKeys5
                let trykeys6 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg6 mapKeys6
                let trykeys7 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg7 mapKeys7
                let trykeys8 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg8 mapKeys8
                let tryvals =  API.In.D1.Tag.Try.tryDV xlkinds None None valueTypeTag mapValues

                match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5, trykeys6, trykeys7, trykeys8, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5), Some (gtykey6, keys6), Some (gtykey7, keys7), Some (gtykey8, keys8), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5; gtykey6; gtykey7; gtykey8 |]
                    let keys = [| keys1; keys2; keys3; keys4; keys5; keys6; keys7; keys8 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

            | Some ktg2, Some ktg3, Some ktg4, Some ktg5, Some ktg6, Some ktg7, None, None, None -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV xlkinds None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg2 mapKeys2
                let trykeys3 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg3 mapKeys3
                let trykeys4 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg4 mapKeys4
                let trykeys5 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg5 mapKeys5
                let trykeys6 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg6 mapKeys6
                let trykeys7 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg7 mapKeys7
                let tryvals =  API.In.D1.Tag.Try.tryDV xlkinds None None valueTypeTag mapValues

                match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5, trykeys6, trykeys7, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5), Some (gtykey6, keys6), Some (gtykey7, keys7), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5; gtykey6; gtykey7 |]
                    let keys = [| keys1; keys2; keys3; keys4; keys5; keys6; keys7 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

            | Some ktg2, Some ktg3, Some ktg4, Some ktg5, Some ktg6, None, None, None, None -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV xlkinds None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg2 mapKeys2
                let trykeys3 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg3 mapKeys3
                let trykeys4 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg4 mapKeys4
                let trykeys5 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg5 mapKeys5
                let trykeys6 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg6 mapKeys6
                let tryvals =  API.In.D1.Tag.Try.tryDV xlkinds None None valueTypeTag mapValues

                match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5, trykeys6, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5), Some (gtykey6, keys6), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5; gtykey6 |]
                    let keys = [| keys1; keys2; keys3; keys4; keys5; keys6 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

            | Some ktg2, Some ktg3, Some ktg4, Some ktg5, None, None, None, None, None -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV xlkinds None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg2 mapKeys2
                let trykeys3 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg3 mapKeys3
                let trykeys4 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg4 mapKeys4
                let trykeys5 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg5 mapKeys5
                let tryvals =  API.In.D1.Tag.Try.tryDV xlkinds None None valueTypeTag mapValues

                match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5 |]
                    let keys = [| keys1; keys2; keys3; keys4; keys5 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

            | Some ktg2, Some ktg3, Some ktg4, None, None, None, None, None, None -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV xlkinds None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg2 mapKeys2
                let trykeys3 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg3 mapKeys3
                let trykeys4 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg4 mapKeys4
                let tryvals =  API.In.D1.Tag.Try.tryDV xlkinds None None valueTypeTag mapValues

                match trykeys1, trykeys2, trykeys3, trykeys4, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4 |]
                    let keys = [| keys1; keys2; keys3; keys4 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

            | Some ktg2, Some ktg3, None, None, None, None, None, None, None -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV xlkinds None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg2 mapKeys2
                let trykeys3 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg3 mapKeys3
                let tryvals =  API.In.D1.Tag.Try.tryDV xlkinds None None valueTypeTag mapValues

                match trykeys1, trykeys2, trykeys3, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2; gtykey3 |]
                    let keys = [| keys1; keys2; keys3 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

            | Some ktg2, None, None, None, None, None, None, None, None -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV xlkinds None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV xlkinds None None ktg2 mapKeys2
                let tryvals =  API.In.D1.Tag.Try.tryDV xlkinds None None valueTypeTag mapValues

                match trykeys1, trykeys2, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2 |]
                    let keys = [| keys1; keys2 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

            | _ -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV xlkinds None None k1TypeTag mapKeys1
                let tryvals =  API.In.D1.Tag.Try.tryDV xlkinds None None valueTypeTag mapValues

                match trykeys1, tryvals with
                | Some (gtykey1, keys1), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1 |]
                    let keys = [| keys1 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

        match gtykeys_keys_gtyvals_vals with
        | None -> Proxys.def.failed
        | Some (gtykeys, keys, gtyval, vals) ->
            let map = MAP.RxFn.In.mapN gtykeys gtyval keys vals
            let res = map |> MRegistry.register rfid
            res |> box

    [<ExcelFunction(Category="Map", Description="Creates a Map<'Key1*'Key2...,'Val> R-object from several 1D-arrays of keys and one 2D-array of values.")>]
    let map_ofRng2D
        ([<ExcelArgument(Description= "VKey1 type tag: bool, date, double, doubleNaN, string or obj. Add \'#'\ prefix for optional type: #bool, #date, #double, #doubleNaN, #string or #obj.")>] vk1Tag: string)
        ([<ExcelArgument(Description= "VKey2 type tag.")>] vk2Tag: obj)
        ([<ExcelArgument(Description= "VKey3 type tag.")>] vk3Tag: obj)
        ([<ExcelArgument(Description= "VKey4 type tag.")>] vk4Tag: obj)
        ([<ExcelArgument(Description= "VKey5 type tag.")>] vk5Tag: obj)
        ([<ExcelArgument(Description= "VKey6 type tag.")>] vk6Tag: obj)
        ([<ExcelArgument(Description= "HKey1 type tag.")>] hk1Tag: string)
        ([<ExcelArgument(Description= "HKey2 type tag.")>] hk2Tag: obj)
        ([<ExcelArgument(Description= "HKey3 type tag.")>] hk3Tag: obj)
        ([<ExcelArgument(Description= "HKey4 type tag.")>] hk4Tag: obj)
        ([<ExcelArgument(Description= "HKey5 type tag.")>] hk5Tag: obj)
        ([<ExcelArgument(Description= "HKey6 type tag.")>] hk6Tag: obj)
        ([<ExcelArgument(Description= "Value type tag.")>] valueTag: string)
        ([<ExcelArgument(Description= "Vkeys1.")>] mapVKeys1: obj)
        ([<ExcelArgument(Description= "Vkeys2.")>] mapVKeys2: obj)
        ([<ExcelArgument(Description= "Vkeys3.")>] mapVKeys3: obj)
        ([<ExcelArgument(Description= "Vkeys4.")>] mapVKeys4: obj)
        ([<ExcelArgument(Description= "Vkeys5.")>] mapVKeys5: obj)
        ([<ExcelArgument(Description= "Vkeys6.")>] mapVKeys6: obj)
        ([<ExcelArgument(Description= "Hkeys1.")>] mapHKeys1: obj)
        ([<ExcelArgument(Description= "Hkeys2.")>] mapHKeys2: obj)
        ([<ExcelArgument(Description= "Hkeys3.")>] mapHKeys3: obj)
        ([<ExcelArgument(Description= "Hkeys4.")>] mapHKeys4: obj)
        ([<ExcelArgument(Description= "Hkeys5.")>] mapHKeys5: obj)
        ([<ExcelArgument(Description= "Hkeys6.")>] mapHKeys6: obj)
        ([<ExcelArgument(Description= "values.")>] mapValues: obj)
        : obj  =

        // intermediary stage
        let vktag1 = vk1Tag
        let vktag2 = In.D0.Stg.Opt.def None vk2Tag
        let vktag3 = In.D0.Stg.Opt.def None vk3Tag
        let vktag4 = In.D0.Stg.Opt.def None vk4Tag
        let vktag5 = In.D0.Stg.Opt.def None vk5Tag
        let vktag6 = In.D0.Stg.Opt.def None vk6Tag
        let hktag1 = hk1Tag
        let hktag2 = In.D0.Stg.Opt.def None hk2Tag
        let hktag3 = In.D0.Stg.Opt.def None hk3Tag
        let hktag4 = In.D0.Stg.Opt.def None hk4Tag
        let hktag5 = In.D0.Stg.Opt.def None hk5Tag
        let hktag6 = In.D0.Stg.Opt.def None hk6Tag

        // caller cell's reference ID
        let rfid = MRegistry.refID

        let tryvals =  API.In.D2.Tag.Try.tryDV' Set.empty None valueTag mapValues

        match tryvals with
        | None -> Proxys.def.failed
        | Some (gtyval, vals) ->
            let vgtykeys_keys =
                match vktag2, vktag3, vktag4, vktag5, vktag6 with
                | Some vktg2, Some vktg3, Some vktg4, Some vktg5, Some vktg6 -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktag1 mapVKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktg2 mapVKeys2
                    let trykeys3 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktg3 mapVKeys3
                    let trykeys4 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktg4 mapVKeys4
                    let trykeys5 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktg5 mapVKeys5
                    let trykeys6 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktg6 mapVKeys6

                    match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5, trykeys6 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5), Some (gtykey6, keys6) -> 
                        let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5; gtykey6 |]
                        let keys = [| keys1; keys2; keys3; keys4; keys5; keys6 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | Some vktg2, Some vktg3, Some vktg4, Some vktg5, None -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktag1 mapVKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktg2 mapVKeys2
                    let trykeys3 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktg3 mapVKeys3
                    let trykeys4 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktg4 mapVKeys4
                    let trykeys5 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktg5 mapVKeys5

                    match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5) -> 
                        let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5 |]
                        let keys = [| keys1; keys2; keys3; keys4; keys5 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | Some vktg2, Some vktg3, Some vktg4, None, None -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktag1 mapVKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktg2 mapVKeys2
                    let trykeys3 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktg3 mapVKeys3
                    let trykeys4 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktg4 mapVKeys4

                    match trykeys1, trykeys2, trykeys3, trykeys4 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4) -> 
                        let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4 |]
                        let keys = [| keys1; keys2; keys3; keys4 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | Some vktg2, Some vktg3, None, None, None -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktag1 mapVKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktg2 mapVKeys2
                    let trykeys3 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktg3 mapVKeys3

                    match trykeys1, trykeys2, trykeys3 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3) -> 
                        let gtykeys = [| gtykey1; gtykey2; gtykey3 |]
                        let keys = [| keys1; keys2; keys3 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | Some vktg2, None, None, None, None -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktag1 mapVKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktg2 mapVKeys2

                    match trykeys1, trykeys2 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2) -> 
                        let gtykeys = [| gtykey1; gtykey2 |]
                        let keys = [| keys1; keys2 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | _ -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV Set.empty None None vktag1 mapVKeys1

                    match trykeys1 with
                    | Some (gtykey1, keys1) -> 
                        let gtykeys = [| gtykey1 |]
                        let keys = [| keys1 |]
                        Some (gtykeys, keys)
                    | _ -> None

            let hgtykeys_keys =
                match hktag2, hktag3, hktag4, hktag5, hktag6 with
                | Some hktg2, Some hktg3, Some hktg4, Some hktg5, Some hktg6 -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktag1 mapHKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktg2 mapHKeys2
                    let trykeys3 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktg3 mapHKeys3
                    let trykeys4 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktg4 mapHKeys4
                    let trykeys5 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktg5 mapHKeys5
                    let trykeys6 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktg6 mapHKeys6

                    match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5, trykeys6 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5), Some (gtykey6, keys6) -> 
                        let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5; gtykey6 |]
                        let keys = [| keys1; keys2; keys3; keys4; keys5; keys6 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | Some hktg2, Some hktg3, Some hktg4, Some hktg5, None -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktag1 mapHKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktg2 mapHKeys2
                    let trykeys3 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktg3 mapHKeys3
                    let trykeys4 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktg4 mapHKeys4
                    let trykeys5 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktg5 mapHKeys5

                    match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5) -> 
                        let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5 |]
                        let keys = [| keys1; keys2; keys3; keys4; keys5 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | Some hktg2, Some hktg3, Some hktg4, None, None -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktag1 mapHKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktg2 mapHKeys2
                    let trykeys3 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktg3 mapHKeys3
                    let trykeys4 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktg4 mapHKeys4

                    match trykeys1, trykeys2, trykeys3, trykeys4 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4) -> 
                        let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4 |]
                        let keys = [| keys1; keys2; keys3; keys4 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | Some hktg2, Some hktg3, None, None, None -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktag1 mapHKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktg2 mapHKeys2
                    let trykeys3 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktg3 mapHKeys3

                    match trykeys1, trykeys2, trykeys3 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3) -> 
                        let gtykeys = [| gtykey1; gtykey2; gtykey3 |]
                        let keys = [| keys1; keys2; keys3 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | Some hktg2, None, None, None, None -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktag1 mapHKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktg2 mapHKeys2

                    match trykeys1, trykeys2 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2) -> 
                        let gtykeys = [| gtykey1; gtykey2 |]
                        let keys = [| keys1; keys2 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | _ -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV Set.empty None None hktag1 mapHKeys1

                    match trykeys1 with
                    | Some (gtykey1, keys1) -> 
                        let gtykeys = [| gtykey1 |]
                        let keys = [| keys1 |]
                        Some (gtykeys, keys)
                    | _ -> None

            match vgtykeys_keys, hgtykeys_keys with
            | None, _ -> Proxys.def.failed
            | _, None -> Proxys.def.failed
            | Some (vgtykeys, vkeys), Some (hgtykeys, hkeys) ->
                let map = MAP.RxFn.In.map2D vgtykeys hgtykeys gtyval vkeys hkeys vals
                let res = map |> MRegistry.register rfid
                res |> box

    [<ExcelFunction(Category="Map", Description="Returns the size of a Map R-object.")>]
    let test_isEmpty
        ([<ExcelArgument(Description= "arg1")>] arg1: obj) 
        ([<ExcelArgument(Description= "arg2")>] arg2: obj) 
        : obj = 

        // result
        let test1 = match arg1 with | :? ExcelMissing -> true | _ -> false
        let test2 = match arg2 with | :? ExcelEmpty -> true | _ -> false

        sprintf "%b <> %b" test1 test2
        |> box

    [<ExcelFunction(Category="Map", Description="Returns the size of a Map R-object.")>]
    let map_count
        ([<ExcelArgument(Description= "Map R-object.")>] rgMap: string) 
        : obj = 

        // result
        match MAP.RxFn.Out.count rgMap with
        | None -> Proxys.def.failed  // TODO Unbox.apply?
        | Some o -> o

    [<ExcelFunction(Category="Map", Description="Returns a R-object map's keys array.")>]
    let map_keys
        ([<ExcelArgument(Description= "Map R-object.")>] rgMap: string)
        ([<ExcelArgument(Description= "[Return R-object. Default is false.]")>] returnRObject: obj)
        : obj[] = 

        // intermediary stage
        let robjoutput = API.In.D0.Bool.def false returnRObject

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match MAP.RxFn.Out.keys rgMap rfid Proxys.def with
        | None -> [| Proxys.def.failed |]  // TODO Unbox.apply?
        | Some o1D -> o1D

    [<ExcelFunction(Category="Map", Description="Returns a R-object map's values array.")>]
    let map_vals
        ([<ExcelArgument(Description= "Map R-object.")>] rgMap: string) 
        : obj[] = 

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match MAP.RxFn.Out.values rgMap true rfid Proxys.def with
        | None -> [| Proxys.def.failed |]  // TODO Unbox.apply?
        | Some o1D -> o1D

    [<ExcelFunction(Category="Map", Description="Returns the value associated to a tuple of keys.")>]
    let map_find
        ([<ExcelArgument(Description= "Map R-object.")>] rgMap: string) 
        ([<ExcelArgument(Description= "Map key1.")>] mapKey1: obj)
        ([<ExcelArgument(Description= "Map key2.")>] mapKey2: obj)
        ([<ExcelArgument(Description= "Map key3.")>] mapKey3: obj)
        ([<ExcelArgument(Description= "Map key4.")>] mapKey4: obj)
        ([<ExcelArgument(Description= "Map key5.")>] mapKey5: obj)
        ([<ExcelArgument(Description= "Map key6.")>] mapKey6: obj)
        ([<ExcelArgument(Description= "Map key7.")>] mapKey7: obj)
        ([<ExcelArgument(Description= "Map key8.")>] mapKey8: obj)
        ([<ExcelArgument(Description= "Map key9.")>] mapKey9: obj)
        ([<ExcelArgument(Description= "Map key10.")>] mapKey10: obj)
        ([<ExcelArgument(Description= "[Failure value. Default is #N/A.]")>] failureValue: obj)
        : obj = 

        // caller cell's reference ID
        let rfid = MRegistry.refID
        
        // intermediary stage
        let mapkey2 = In.D0.Missing.Obj.tryO mapKey2
        let mapkey3 = In.D0.Missing.Obj.tryO mapKey3
        let mapkey4 = In.D0.Missing.Obj.tryO mapKey4
        let mapkey5 = In.D0.Missing.Obj.tryO mapKey5
        let mapkey6 = In.D0.Missing.Obj.tryO mapKey6
        let mapkey7 = In.D0.Missing.Obj.tryO mapKey7
        let mapkey8 = In.D0.Missing.Obj.tryO mapKey8
        let mapkey9 = In.D0.Missing.Obj.tryO mapKey9
        let mapkey10 = In.D0.Missing.Obj.tryO mapKey10
        let failureVal = In.D0.Missing.Obj.subst Proxys.def.failed failureValue
        let proxys = { def with failed = failureVal }

        let okeys =
            // result
            match mapkey2, mapkey3, mapkey4, mapkey5, mapkey6, mapkey7, mapkey8, mapkey9, mapkey10 with
            | Some key2, Some key3, Some key4, Some key5, Some key6, Some key7, Some key8, Some key9, Some key10 -> 
                [| mapKey1; key2; key3; key4; key5; key6; key7; key8; key9; key10 |]
            | Some key2, Some key3, Some key4, Some key5, Some key6, Some key7, Some key8, Some key9, None -> 
                [| mapKey1; key2; key3; key4; key5; key6; key7; key8; key9 |]
            | Some key2, Some key3, Some key4, Some key5, Some key6, Some key7, Some key8, None, None -> 
                [| mapKey1; key2; key3; key4; key5; key6; key7; key8 |]
            | Some key2, Some key3, Some key4, Some key5, Some key6, Some key7, None, None, None -> 
                [| mapKey1; key2; key3; key4; key5; key6; key7 |]
            | Some key2, Some key3, Some key4, Some key5, Some key6, None, None, None, None -> 
                [| mapKey1; key2; key3; key4; key5; key6 |]
            | Some key2, Some key3, Some key4, Some key5, None, None, None, None, None -> 
                [| mapKey1; key2; key3; key4; key5 |]
            | Some key2, Some key3, Some key4, None, None, None, None, None, None -> 
                [| mapKey1; key2; key3; key4 |]
            | Some key2, Some key3, None, None, None, None, None, None, None -> 
                [| mapKey1; key2; key3 |]
            | Some key2, None, None, None, None, None, None, None, None -> 
                [| mapKey1; key2 |]
            | _ -> 
                [| mapKey1 |]

        if okeys.Length = 1 then
            match MAP.RxFn.Out.find1 rgMap proxys rfid okeys.[0] with
            | None -> proxys.failed
            | Some o0D -> o0D
        else
            match MAP.RxFn.Out.findN rgMap proxys rfid okeys with
            | None -> proxys.failed
            | Some o0D -> o0D

    [<ExcelFunction(Category="Map", Description="Returns the value associated to an array of keys.")>]
    let map_find1D
        ([<ExcelArgument(Description= "Map R-object.")>] rgMap: string) 
        ([<ExcelArgument(Description= "Array of keys.")>] keys1D: obj)
        ([<ExcelArgument(Description= "[Failure value. Default is #N/A.]")>] failureValue: obj)
        : obj[] = 

        // intermediary stage
        let okeys = In.Cast.to1D false keys1D
        let failureVal = In.D0.Missing.Obj.subst Proxys.def.failed failureValue
        let proxys = { def with failed = failureVal }

        // caller cell's reference ID
        let rfid = MRegistry.refID

        match MAP.RxFn.Out.find1D1 rgMap proxys rfid okeys with
        | None -> [| proxys.failed |]
        | Some xo1D -> xo1D |> Out.D1.Unbox.apply proxys id

    [<ExcelFunction(Category="Map", Description="Returns the union of many compatible Map<K,V> R-objects.")>]
    let map_pool
        ([<ExcelArgument(Description= "Map R-objects.")>] rgMap1D: obj) 
        : obj = 

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match MAP.RxFn.In.pool rgMap1D with
        | None -> Proxys.def.failed  // TODO Unbox.apply?
        | Some regObjMap -> regObjMap |> MRegistry.register rfid |> box

module Rel =
    open API
    open System.Collections
    open type Out.Proxys




    // a custom compare function is needed because we use F# Set<Elem>, which requires a comparison constraint on its elements.
    [<RequireQualifiedAccess>]
    module Compare =
        // -----------------------------
        // -- Comparaison functions
        // -----------------------------

        // write your custom compare function here.
        /// Used if Structural Equality does not apply.
        let custom (o1: obj) (o2: obj) : int =
            match (o1, o2) with
            //| (:? Foo), (:? Foo) -> 0
            //| (:? Foo), _ -> -1
            //| _ , (:? Foo) -> 1
            | _ -> failwith "System.ArgumentException: Object must be of type IComparable or IStructuralComparable.\nat oCompare (Object value1) (Object value2)"
    
        let compare (o1: obj) (o2: obj) : int =
            match (o1, o2) with
            | (:? IComparable as c1), (:? IComparable as c2) -> 
                try 
                    c1.CompareTo(c2)
                with
                | e -> failwith e.Message
            | (:? IStructuralComparable as c1), (:? IStructuralComparable as c2) -> 
                try 
                    Operators.compare c1 c2
                with
                | e -> failwith e.Message
            | _ ->
                custom o1 o2

    // F# Set requires comparison, and the relation object is defined as Set<Elem>, so custom comparison is necessary for Elem and Field.
    // As the Relation object does not allow two Fields with the same name in its header,
    // we can compare fields with respect to their name only (and ignore their type, which we cannot compare).
    [<CustomEquality; CustomComparison>]
    type Field = { fname: string; ftype: Type } with
        override f1.Equals(o2) =
            match o2 with 
            | :? Field as f2 -> (f1.fname = f2.fname) && (f1.ftype = f2.ftype)
            | _ -> false
        override f.GetHashCode() = hash f

        interface System.IComparable with
            member f1.CompareTo o2 =
                match o2 with 
                | :? Field as f2 -> Operators.compare f1.fname f2.fname |> sign // f1.fname.CompareTo(f2.fname)
                | _ -> invalidArg "o2" "Field: Cannot compare values of different types."

        static member zip (fieldNames: string[]) (fieldTypes: Type[]) : Set<Field> =
            Array.map2 (fun fname ftype -> { fname = fname ; ftype = ftype }) fieldNames fieldTypes
            |> Set.ofArray

        /// Returns all field-names, sorted.
        static member names (fields: Set<Field>) : Set<string> = fields |> Set.map (fun field -> field.fname)

        /// Returns field-types, given field-names.
        /// Returns None if any fieldName is not included in fields.
        /// Preserves fieldNames order.
        static member types (fields: Set<Field>) (fieldNames: string[]) : Type[] option = 
            let mapping = fields |> Set.toArray |> Array.map (fun field -> (field.fname, field.ftype)) |> Map.ofArray
            let types = fieldNames |> Array.choose (fun fname -> mapping |> Map.tryFind fname)
            if types.Length < fieldNames.Length then
                None
            else
                Some types

        /// Returns all fields, which fname is included in keepNames (away = false).
        /// Discards all of keepNames names which are not included in fields's names (away = true). 
        static member project (fields: Set<Field>) (away: bool) (keepNames: Set<string>) : Set<Field> =
            let fnames =
                if away then
                    Set.difference (fields |> Field.names) keepNames
                else
                    Set.intersect (fields |> Field.names) keepNames

            fields |> Set.filter (fun field -> fnames |> Set.contains field.fname)

        /// Returns fields' name and index,
        /// for which the field's fname is included in keepNames (away = false),
        /// or for which the field's fname is not included in keepNames (away = true).
        /// Indexing is done following F# sets natural order.
        static member private indexing (fields: Set<Field>) (away: bool) (keepNames: Set<string>) : (int*string)[] =
            let indexedNames = fields |> Field.names |> Set.toArray |> Array.indexed
            let projectNames = Field.project fields away keepNames |> Field.names
            indexedNames |> Array.filter (fun (idx, fnm) -> projectNames |> Set.contains fnm)

        /// Returns fields' indices,
        /// for which the field's fname is included in keepNames (away = false),
        /// or for which the field's fname is not included in keepNames (away = true).
        /// Indexing is done following F# sets natural order.
        static member indices (fields: Set<Field>) (away: bool) (keepNames: Set<string>) : int[] = 
            Field.indexing fields away keepNames |> Array.map fst

        /// Returns a fields's (name -> index) mapping,
        /// for which the field's fname is included in keepNames (away = false),
        /// or for which the field's fname is not included in keepNames (away = true).
        /// Indexing is done following F# sets natural order.
        static member indexMap (fields: Set<Field>) (away: bool) (keepNames: Set<string>) : Map<string,int> =
            Field.indexing fields away keepNames |> Array.map (fun (x, y) -> (y, x)) |> Map.ofArray

        /// Same as Field.indices, but with same ordering as keepNames.
        /// All of keepNames must be valid field names, otherwise returns None
        static member indicesOrdered (fields: Set<Field>) (keepNames: string[]) : int[] option = 
            let indexmap = Field.indexMap fields false (keepNames |> Set.ofArray)
            let indices = keepNames |> Array.choose (fun name -> indexmap |> Map.tryFind name)
            if indices.Length < keepNames.Length then
                None
            else
                Some indices

        /// Returns true if two fields or more have same name but different type.
        static member incompatible (fields1: Set<Field>) (fields2: Set<Field>) : bool =
            let fNames1 = fields1 |> Field.names
            let fNames2 = fields2 |> Field.names
            let comNames = Set.intersect fNames1 fNames2
            if comNames |> Set.isEmpty then
                false
            else
                let comFields1 = fields1 |> Set.filter (fun field -> comNames |> Set.contains field.fname)
                let comFields2 = fields2 |> Set.filter (fun field -> comNames |> Set.contains field.fname)
                comFields2 <> comFields1

        /// Returns fields common to both fields sets.
        /// Incompatible fields will be excluded.
        static member common (fields1: Set<Field>) (fields2: Set<Field>) : Set<Field> = Set.intersect fields1 fields2

        /// Returns true if the two fields sets are disjoint.
        /// Incompatible sets might return true.
        static member disjoint (fields1: Set<Field>) (fields2: Set<Field>) : bool = Field.common fields1 fields2 |> Set.isEmpty

        /// Returns true if fields2 is included in fields1.
        static member isSuperset (fields1: Set<Field>) (fields2: Set<Field>) : bool = Set.isSuperset fields1 fields2

        /// Returns true if all fnames are all fields' names. 
        static member containsNames (fnames: Set<string>) (fields: Set<Field>) : bool = Set.isSuperset (fields |> Field.names) fnames

    // F# Set requires comparison, and the relation object is defined as Set<Elem>, so custom comparison is necessary for Elem and Field.
    [<CustomEquality; CustomComparison>]
    type Elem = { fname: string; value: obj } with
        override e1.Equals(o2) =
            match o2 with 
            | :? Elem as e2 -> (e1.fname = e2.fname) && (e1.value = e2.value)
            | _ -> false

        override e.GetHashCode() = hash e.value

        interface System.IComparable with
            member e1.CompareTo o2 =
                match o2 with 
                | :? Elem as e2 -> 
                    if e1.fname <> e2.fname then 
                        e1.fname.CompareTo(e2.fname)
                    else 
                        Compare.compare e1.value e2.value // cases when we reach here?
                | _ -> invalidArg "o2" "Elem: Cannot compare values of different types."

        static member zip (fieldNames: string[]) (elems: obj[]) : Elem[] =
            Array.zip fieldNames elems |> Array.sortBy fst |> Array.map (fun (fname, o) -> { fname = fname; value = o })
    
    /// Row is an (unordered) set of Elem-s. Similar to C.J. Date's concept of 'Tuple' in 'An Introduction to Database Systems'.
    type Row = Elem[]

    module Row =
        let item (index: int) (row: Row) : obj = let elem = row.[index] in elem.value
        let sort (row: Row) : Row = row |> Array.sortBy (fun elem -> elem.fname)

        let ofPositions (positions: int[]) (row: Row) : Row = positions |> Array.map (fun i -> row.[i])
        let ofPositions2 (positions1: int[]) (positions2: int[]) (row: Row) : Row = Array.append (ofPositions positions1 row) (ofPositions positions2 row)

        /// Returns a lookup index of (hashed) row-value -> row indexes.
        /// Maps each row index to the indices of all rows which have common mapping image (that is, each row where (row |> mapping) is the same).
        /// Will only works if the rows are sorted by fname.
        let index (mapping: Row -> Row) (rows: Row[]) : Dictionary<int, int[]> =
            let dc = new Dictionary<int, int[]>()
            let addEntry index row = 
                let rowHash = mapping row |> hash
                if dc.ContainsKey(rowHash) then 
                    dc.[rowHash] <- Array.append dc.[rowHash] [| index |]
                else
                    dc.Add(rowHash, [| index |])
            rows |> Array.iteri addEntry
            dc

    /// A relation, Rel, is:
    /// 1. A header, a set of fields.
    /// 2. A body, a set of Row-s, compatible with the header.
    type Rel = { fields: Set<Field>; body: Set<Row> } with
        static member DEE : Rel = { fields = Set.empty; body = [||] |> Set.singleton }
        static member DUM : Rel = { fields = Set.empty; body = Set.empty }

        static member ofRng (xlNanKinds: Set<Kind>) (fieldNames: string[]) (typeTags: string[]) (o2D: obj[,]) : Rel option = 
            let len = fieldNames.Length
            let valid = (typeTags.Length = len) && (o2D |> Array2D.length2 = len) && (fieldNames |> Array.distinct |> Array.length = len)

            if not valid then
                None
            else
                match In.D2.Tag.Multi.tryDV None xlNanKinds false typeTags o2D with
                | None -> None
                | Some (gentys, o2D) ->
                    let fields = Field.zip fieldNames gentys
                    let body : Set<Row> = [| for i in o2D.GetLowerBound(0) .. o2D.GetUpperBound(0) -> Elem.zip fieldNames o2D.[i,*] |] |> Set.ofArray

                    let rel : Rel = { fields = fields; body = body }
                    rel |> Some

        static member ofHead (fieldNames: string[]) (typeTags: string[]) : Rel option = 
            let len = fieldNames.Length
            let ftypes = typeTags |> Array.map (Variant.labelType false)

            let valid = (ftypes.Length = len) && (fieldNames |> Array.distinct |> Array.length = len)

            if not valid then
                None
            else                
                let fields = Field.zip fieldNames (typeTags |> Array.map (Variant.labelType false))
                let rel : Rel = { fields = fields; body = Set.empty }
                rel |> Some

        member r.head = r.fields |> Field.names |> Set.toArray

        member r.header = r.fields |> Field.names |> Set.toArray |> Array.map box

        member r.toRng (showHead: bool) (unwrapOptions: bool) (refKey: String) (proxys: Out.Proxys) : obj[,] =
            // let rows = r.body |> Set.toArray |> Array.map (fun row -> row |> Array.map (fun elem -> elem.value))
            if r = Rel.DEE then
                box "<DEE>" |> Toolbox.Array2D.singleton
            elif r = Rel.DUM then
                box "<DUM>" |> Toolbox.Array2D.singleton
            elif r.card = 0 then
                box proxys.nan |> Toolbox.Array2D.singleton
            elif r.count = 0 then
                if showHead then
                    [| r.head |> Array.map box |] |> array2D
                else
                    box proxys.empty |> Toolbox.Array2D.singleton
            else
                let ftypes = r.fields |> Set.toArray |> Array.map (fun field -> field.ftype)
                let rows = r.body |> Set.toArray |> Array.map (fun row -> row |> Array.map (fun elem -> elem.value))
                let xlColumns = 
                    [| for j in ftypes.GetLowerBound(0) .. ftypes.GetUpperBound(0) -> 
                          let methodNm = "outObjWithRefControl"
                          let removeReferences = (j = ftypes.GetLowerBound(0))
                          let col = [| for i in rows.GetLowerBound(0) .. rows.GetUpperBound(0) -> rows.[i].[j] |]
                          let xlcol = Toolbox.Generics.apply<A1D.GenFn> methodNm [||] [| removeReferences; unwrapOptions; refKey; proxys |] ([| ftypes.[j] |], col)
                          xlcol :?> obj[]
                    |]

                if showHead then
                    let head = r.head
                    let xlColumns' = xlColumns |> Array.mapi (fun i col -> Array.append [| box head.[i] |] col)
                    xlColumns' |> Toolbox.Array2D.concat2D false
                else
                    xlColumns |> Toolbox.Array2D.concat2D false

        /// Returns true if some fields have the same name but have different types.
        static member incompatible (r1: Rel) (r2: Rel) : bool = Field.incompatible r1.fields r2.fields

        /// Returns true if two relations have the same header (fields / attributes).
        static member rmatch (r1: Rel) (r2: Rel) : bool = r1.fields = r2.fields

        /// Returns true if the relations do not share any common field / attribute.
        static member disjoint (r1: Rel) (r2: Rel) : bool = Field.disjoint r1.fields r2.fields

        /// Returns the fields common to both relations.
        static member commonFields (r1: Rel) (r2: Rel) : Set<Field> = Field.common r1.fields r2.fields

        /// Returns true if *all* of r2's fields are included in r1's.
        /// Usage r2 |> isCompatibleWith <| r1
        static member isCompatibleWith (r2: Rel) (r1: Rel) = Set.isSuperset r1.fields r2.fields

        /// Returns the number of fields / attributes.
        member this.card : int = this.fields |> Set.count

        /// Returns the number of rows.
        member this.count : int = this.body |> Set.count

    /// Union operator.
    /// The fields (attributes) of the two relation operands must be equal.
    let union (r1: Rel) (r2: Rel) : Rel option =
        //if Rel.incompatible r1 r2 then
        //    None
        //else
        if Rel.rmatch r1 r2 then
            let body = Set.union r1.body r2.body
            let rel : Rel = { fields = r1.fields; body = body }
            rel |> Some
        else
            None

    /// Intersection operator.
    /// The fields (attributes) of the two relation operands must be equal.
    let inter (r1: Rel) (r2: Rel) : Rel option =
        //if Rel.incompatible r1 r2 then
        //    None
        //else
        if Rel.rmatch r1 r2 then
            let body = Set.intersect r1.body r2.body
            let rel : Rel = { fields = r1.fields; body = body }
            rel |> Some
        else
            None

    /// Difference operator.
    /// The fields (attributes) of the two relation operands must be equal.
    let minus (r1: Rel) (r2: Rel) : Rel option =
        //if Rel.incompatible r1 r2 then
        //    None
        //else
        if Rel.rmatch r1 r2 then
            let body = Set.difference r1.body r2.body
            let rel : Rel = { fields = r1.fields; body = body }
            rel |> Some
        else
            None

    /// Product operator.
    /// The fields (attributes) of the two relation operands must be disjoint.
    let prod (r1: Rel) (r2: Rel) : Rel option =
        if Rel.incompatible r1 r2 then
            None
        elif r1 = Rel.DEE then Some r2
        elif r2 = Rel.DEE then Some r1
        elif r1 = Rel.DUM then Some { r2 with body = Set.empty }
        elif r2 = Rel.DUM then Some { r1 with body = Set.empty }
        elif Rel.disjoint r1 r2 then
            let pairedFields = Set.union r1.fields r2.fields
            let pairedRows : Row[] = Array.allPairs (r1.body |> Set.toArray) (r2.body |> Set.toArray) |> Array.map (fun (a1, a2) -> Array.append a1 a2 |> Array.sort)
            let rel : Rel = { fields = pairedFields; body = pairedRows |> Set.ofArray }
            rel |> Some
        else
            None

    /// Rename operator.
    let rename (r: Rel) (mapOldNamesNewNames: Map<string,string>) : Rel option =
        let rname (old: string) = mapOldNamesNewNames |> Map.tryFind old |> Option.defaultValue old

        let oldNames = r.fields |> Field.names |> Set.toArray
        let newNames = oldNames |> Array.map rname |> Array.distinct

        if newNames.Length <> oldNames.Length then
            None
        else
            let fields = r.fields |> Set.map (fun field -> { field with fname = rname field.fname })
            let body = r.body |> Set.map (fun row -> row |> Array.map (fun elem -> { elem with fname = rname elem.fname }) |> Row.sort)

            let rel : Rel = { fields = fields; body = body }
            rel |> Some

    /// Project operator.
    /// Keeps all of r's fields which name are included in R's field names (away = false).
    /// Discards all of r's fields which name are not included in keepNames (away = true).
    let project (r: Rel) (away: bool) (keepNames: Set<string>) : Rel =
        let fields = Field.project r.fields away keepNames
        let indices = Field.indices r.fields away keepNames

        let body = r.body |> Set.map (fun row -> indices |> Array.map (fun i -> row.[i]))

        let rel : Rel = { fields = fields; body = body }
        rel

    /// Join operator.
    let join (r1: Rel) (r2: Rel) : Rel option = 
        if Rel.incompatible r1 r2 then
            None
        elif Rel.rmatch r1 r2 then
            inter r1 r2
        else
            match Rel.commonFields r1 r2 with
            // disjoint relations
            | x when x |> Set.isEmpty -> prod r1 r2

            // relations which share common fields
            | commonFields ->
                let comNames = commonFields |> Field.names
                let comIdxs1 = Field.indices r1.fields false comNames
                let outIdxs1 = Field.indices r1.fields true comNames
                let comIdxs2 = Field.indices r2.fields false comNames
                let outIdxs2 = Field.indices r2.fields true comNames

                let rows1 = r1.body |> Set.toArray
                let rows2 = r2.body |> Set.toArray
                let map1 = Row.index (Row.ofPositions comIdxs1) rows1 // (hashvalue -> rows1 indices) mapping
                let map2 = Row.index (Row.ofPositions comIdxs2) rows2 // (hashvalue -> rows2 indices) mapping

                let comHash = Set.intersect (map1.Keys |> Seq.cast<int> |> Set.ofSeq) (map2.Keys |> Seq.cast<int> |> Set.ofSeq) |> Set.toArray

                let ofHash (rowHash: int) : Row[] =
                    let rws1 = map1.[rowHash] |> Array.map (fun idx -> let row = rows1.[idx] in Row.ofPositions2 comIdxs1 outIdxs1 row)
                    let rws2 = map2.[rowHash] |> Array.map (fun idx -> let row = rows2.[idx] in Row.ofPositions outIdxs2 row)
                    let pairs = Array.allPairs rws1 rws2
                    let rows = pairs |> Array.map (fun (row1, row2) -> Array.append row1 row2 |> Row.sort)
                    rows

                let body = comHash |> Array.collect ofHash |> Set.ofArray

                let rel : Rel = { fields = Set.union r1.fields r2.fields; body = body }
                rel |> Some

    /// Semi operators. Only valid when there are common fields.
    ///    - difference = true: implements semi-minus.
    ///    - difference = false: implements semi-join.
    let private semi (difference: bool) (r1: Rel) (r2: Rel) : Rel option = 
        match Rel.commonFields r1 r2 with
        // disjoint relations
        | x when x |> Set.isEmpty -> if difference then r1 |> Some else { r1 with body = Set.empty } |> Some

        // relations which share common fields
        | commonFields ->
            let comNames = commonFields |> Field.names
            let comIdxs1 = Field.indices r1.fields false comNames
            let comIdxs2 = Field.indices r2.fields false comNames

            let rows1 = r1.body |> Set.toArray
            let rows2 = r2.body |> Set.toArray
            let map1 = Row.index (Row.ofPositions comIdxs1) rows1 // (hashvalue -> rows1 indices) mapping
            let map2 = Row.index (Row.ofPositions comIdxs2) rows2 // (hashvalue -> rows2 indices) mapping

            let oper = if difference then Set.difference else Set.intersect
            let comHash = oper (map1.Keys |> Seq.cast<int> |> Set.ofSeq) (map2.Keys |> Seq.cast<int> |> Set.ofSeq)|> Set.toArray

            let ofHash (rowHash: int) : Row[] =
                let rows = map1.[rowHash] |> Array.map (fun idx -> rows1.[idx])
                rows

            let body = comHash |> Array.collect ofHash |> Set.ofArray

            let rel : Rel = { r1 with body = body }
            rel |> Some

    /// Semi-join operator.
    /// All rows / tuples of r1 which have a counterpart in r2.
    let semiJoin (r1: Rel) (r2: Rel) : Rel option = 
        if Rel.incompatible r1 r2 then
            None
        elif Rel.rmatch r1 r2 then
            inter r1 r2
        else
            semi false r1 r2

    /// Semi-minus operator.
    /// All rows / tuples of r1 which have no counterpart in r2.
    let semiMinus (r1: Rel) (r2: Rel) : Rel option = 
        if Rel.incompatible r1 r2 then
            None
        elif Rel.rmatch r1 r2 then
            minus r1 r2
        else
            semi true r1 r2

    /// Left outer join operator.
    /// defaultValue maps r2-specific fields (which aren't not r1's) to a default value.
    let leftJoin (defaultValues: Field -> obj) (r1: Rel) (r2: Rel) : Rel option = 
        if Rel.incompatible r1 r2 then
            None
        elif Rel.rmatch r1 r2 then
            None // inter r1 r2 // WRONG! TO: check if it's capture by the below
        else
            match Rel.commonFields r1 r2 with
            // disjoint relations
            | x when x |> Set.isEmpty -> prod r1 r2

            // relations which share common fields
            | commonFields ->
                let comNames = commonFields |> Field.names
                let comIdxs1 = Field.indices r1.fields false comNames
                let outIdxs1 = Field.indices r1.fields true comNames
                let comIdxs2 = Field.indices r2.fields false comNames
                let outIdxs2 = Field.indices r2.fields true comNames

                let rows1 = r1.body |> Set.toArray
                let rows2 = r2.body |> Set.toArray
                let map1 = Row.index (Row.ofPositions comIdxs1) rows1 // (hashvalue -> rows1 indices) mapping
                let map2 = Row.index (Row.ofPositions comIdxs2) rows2 // (hashvalue -> rows2 indices) mapping

                // 1- common rows
                let comHash = Set.intersect (map1.Keys |> Seq.cast<int> |> Set.ofSeq) (map2.Keys |> Seq.cast<int> |> Set.ofSeq)|> Set.toArray

                let ofHash (rowHash: int) : Row[] =
                    let rows1 = map1.[rowHash] |> Array.map (fun idx -> let row = rows1.[idx] in Row.ofPositions2 comIdxs1 outIdxs1 row)
                    let rows2 = map2.[rowHash] |> Array.map (fun idx -> let row = rows2.[idx] in Row.ofPositions outIdxs2 row)
                    let pairs = Array.allPairs rows1 rows2
                    let rows = pairs |> Array.map (fun (row1, row2) -> Array.append row1 row2 |> Row.sort)
                    rows

                let comRows = comHash |> Array.collect ofHash

                // 2- rows in r1 which do not have a counterpart in r2 over their common fields
                let diffHash = Set.difference (map1.Keys |> Seq.cast<int> |> Set.ofSeq) (map2.Keys |> Seq.cast<int> |> Set.ofSeq)|> Set.toArray
                let defRow = Set.difference r2.fields r1.fields |> Set.map (fun field -> { fname = field.fname; value = field |> defaultValues}) |> Set.toArray

                let ofHash (rowHash: int) : Row[] =
                    let rows1 : Row[] = map1.[rowHash] |> Array.map (fun idx -> Array.append rows1.[idx] defRow |> Row.sort)
                    rows1

                let diffRows = diffHash |> Array.collect ofHash

                // union of common and difference rows
                let body = Array.append comRows diffRows |> Set.ofArray

                let rel : Rel = { fields = Set.union r1.fields r2.fields; body = body }
                rel |> Some

    /// Group operator.
    let group (r: Rel) (perNames: Set<string>) (groupName: string) : Rel option =
        let perFields = Field.project r.fields false perNames           
        let grpFields = Field.project r.fields true perNames

        if perFields |> Field.containsNames (groupName |> Set.singleton) then
            None
        else
            let perIdxs = Field.indices r.fields false perNames
            let restIdxs = Field.indices r.fields true perNames

            // grouped relation fields
            let fields = Set.union perFields ({ fname = groupName; ftype = typeof<Rel> } |> Set.singleton)
        
            // grouped relation body
            let grpRel (grpRows: Row[]) : Rel = { fields = grpFields; body = grpRows |> Set.ofArray }
            let grpElem (grpRows: Row[]) : Row = { fname = groupName; value = grpRel grpRows } |> Array.singleton
            let body = 
                r.body 
                |> Set.toArray
                |> Array.map (fun row -> (Row.ofPositions perIdxs row, Row.ofPositions restIdxs row))
                |> Array.groupBy fst
                |> Array.map (fun (perrow, tpldrows) -> Array.append perrow (tpldrows |> Array.map snd |> grpElem) |> Row.sort)
                |> Set.ofArray

            let rel : Rel = { fields = fields; body = body }
            rel |> Some

    /// Ungroup operator.
    let ungroup (r: Rel) (groupName: string) : Rel option =        
        let groupField = { fname = groupName; ftype = typeof<Rel> }

        if r.count = 0 then
            None
        elif r.fields |> Set.contains groupField |> not then
            None
        else
            let groupIdxs = Field.indices r.fields false (groupName |> Set.singleton)
            let restIdxs = Field.indices r.fields true (groupName |> Set.singleton)

            let groupIdx = groupIdxs |> Array.head
            let sampleRel = let elem = r.body |> Set.toArray |> Array.head |> Array.item groupIdx in elem.value :?> Rel

            // ungroup relation fields
            let fields = Set.union sampleRel.fields (Set.difference r.fields (groupField |> Set.singleton))

            // ungroup relation body
            let body = 
                [| for row in r.body do
                    let restrow = Row.ofPositions restIdxs row
                    let rel = let elem = Row.ofPositions groupIdxs row |> Array.head in elem.value :?> Rel
                    yield! [| for relrow in rel.body -> Array.append restrow relrow |> Row.sort |]
                |]
                |> Set.ofArray

            let rel : Rel = { fields = fields; body = body }
            rel |> Some

    /// Summarize operator.
    /// operations = operators, operNames, resultNames, resultTypes
    /// operator should be typeof<operName> -> resultType
    let summarize' (r: Rel) (perRel: Rel) (operations: ((obj[] -> obj)*string*string*Type)[]) : Rel option =
        let operators, operNames, resultNames, resultTypes = operations |> Toolbox.Array.unzip4

        if (perRel.count = 0) && (perRel.card = 0) && (operations.Length = 0) then
            None
        elif resultNames |> Array.distinct |> Array.length <> (resultNames |> Array.length) then
            None
        elif not <| (perRel |> Rel.isCompatibleWith <| r) then
            None
        elif (not <| (r.fields |> Field.containsNames (operNames |> Set.ofArray))) && (perRel.fields |> Field.containsNames (operNames |> Set.ofArray)) && (r.fields |> Field.containsNames (resultNames |> Set.ofArray)) then
            None
        else
            let commonFields = Rel.commonFields r perRel

            // summarize relation fields
            let resultFields : Set<Field> = Field.zip resultNames resultTypes
            let fields = Set.union commonFields resultFields

            let comNames = commonFields |> Field.names
            let rComIdxs = Field.indices r.fields false comNames
            
            let rRows = r.body |> Set.toArray
            let perRows = perRel.body |> Set.toArray

            // map row index -> indexes of all rows which match over comIdxs positions
            let rMap = Row.index (Row.ofPositions rComIdxs) rRows
            let perMap = Row.index id perRows  // TODO: optimization: better to hash directly perRows' rows.
            let comHash = Set.intersect (rMap.Keys |> Seq.cast<int> |> Set.ofSeq) (perMap.Keys |> Seq.cast<int> |> Set.ofSeq)|> Set.toArray

            // map (oper index, operName) -> (operator, operName, resultName, resultType) 
            let operMap = Array.zip (operNames |> Array.indexed) operations |> Map.ofArray
            // map operName -> column (field) index 
            let rOperIndexMap = Field.indexMap r.fields false (operNames |> Set.ofArray)
            let oper (rows: Row[]) (key: (int*string)) : Elem = 
                let _, opName = key
                let operator, _, resName, resType = operMap |> Map.find key
                let colidx = rOperIndexMap |> Map.find opName
                let res = rows |> Array.map (Array.item colidx) |> Array.map (fun elem -> elem.value) |> operator
                let elem : Elem = { fname = resName; value = res }
                elem

            let ofHash (rowHash: int) : Row =
                let hashes = rMap.[rowHash]
                let idx0 = hashes.[0]
                let comRow = Row.ofPositions rComIdxs rRows.[idx0]
                let operRows = hashes |> Array.map (fun idx -> rRows.[idx])
                let resRow : Row = 
                    [| for kvp in operMap -> kvp.Key |]
                    |> Array.map (oper operRows)
                Array.append comRow resRow |> Row.sort

            let body = comHash |> Array.map ofHash |> Set.ofArray

            let rel : Rel = { fields = fields; body = body }
            rel |> Some

    /// Aggregate operators for Summarize.
    type AOper = | COUNT | DISTINCT | SUM | AVG | MIN | MAX | SPAN with
        member this.resultType (operandType: Type) : Type option =
            match this with
            | COUNT -> Some typeof<int>
            | DISTINCT -> Some typeof<int>
            | SPAN -> if operandType = typeof<DateTime> then Some typeof<double> else None
            | _ -> None

        static member ofLabel (label: string) : AOper =
            match label.ToUpper().Trim() with
            | "CNT" | "CNT" | "COUNT" -> COUNT
            | "D" | "DIST" | "DISTINCT" -> DISTINCT
            | "S" | "SUM" -> SUM
            | "A" | "AVG" | "AVERAGE" -> AVG
            | "MIN" | "MINIMUM" -> MIN
            | "MAX" | "MAXIMUM" -> MAX
            | "S" | "SPAN" -> SPAN
            | _ -> COUNT

    /// Boilerplate code
    module AOper =
        module Dbl =
            let mapUnbox (xs: obj[]) : double[] = xs |> Array.map (unbox<double>)
            let count (xs: double[]) = xs.Length
            let distinct (xs: double[]) = xs |> Array.distinct |> Array.length
            let sum (xs: double[]) = Array.sum xs
            let avg (xs: double[]) = Array.average xs
            let min (xs: double[]) = Array.min xs
            let max (xs: double[]) = Array.max xs
            let span (xs: double[]) = (Array.max xs) - Array.min xs
            let aggregate (aoper: AOper) : (obj[] -> obj) =
                match aoper with
                | COUNT -> mapUnbox >> count >> box
                | DISTINCT -> mapUnbox >> distinct >> box
                | SUM -> mapUnbox >> sum >> box
                | AVG -> mapUnbox >> avg >> box
                | MIN -> mapUnbox >> min >> box
                | MAX -> mapUnbox >> max >> box
                | SPAN -> mapUnbox >> span >> box

        module Intg =
            let mapUnbox (xs: obj[]) : int[] = xs |> Array.map (unbox<int>)
            let count (xs: int[]) = xs.Length
            let distinct (xs: int[]) = xs |> Array.distinct |> Array.length
            let sum (xs: int[]) = Array.sum xs
            let avg (xs: int[]) = xs |> Array.map double |> Array.average |> int
            let min (xs: int[]) = Array.min xs
            let max (xs: int[]) = Array.max xs
            let span (xs: int[]) = (Array.max xs) - Array.min xs
            let aggregate (aoper: AOper) : (obj[] -> obj) =
                match aoper with
                | COUNT -> mapUnbox >> count >> box
                | DISTINCT -> mapUnbox >> distinct >> box
                | SUM -> mapUnbox >> sum >> box
                | AVG -> mapUnbox >> avg >> box
                | MIN -> mapUnbox >> min >> box
                | MAX -> mapUnbox >> max >> box
                | SPAN -> mapUnbox >> span >> box

        module Dte =
            let mapUnbox (xs: obj[]) : DateTime[] = xs |> Array.map (unbox<DateTime>)
            let count (xs: DateTime[]) = xs.Length
            let distinct (xs: DateTime[]) = xs |> Array.distinct |> Array.length
            let sum (xs: DateTime[]) = xs |> Array.map (fun date -> date.ToOADate()) |> Array.sum |> (fun d -> DateTime.FromOADate(d)) // probably meaningless
            let avg (xs: DateTime[]) = xs |> Array.map (fun date -> date.ToOADate()) |> Array.average |> (fun d -> DateTime.FromOADate(d))
            let min (xs: DateTime[]) = Array.min xs
            let max (xs: DateTime[]) = Array.max xs
            let span (xs: DateTime[]) = (Array.max xs - Array.min xs).TotalDays |> double
            let aggregate (aoper: AOper) : (obj[] -> obj) =
                match aoper with
                | COUNT -> mapUnbox >> count >> box
                | DISTINCT -> mapUnbox >> distinct >> box
                | SUM -> mapUnbox >> sum >> box
                | AVG -> mapUnbox >> avg >> box
                | MIN -> mapUnbox >> min >> box
                | MAX -> mapUnbox >> max >> box
                | SPAN -> mapUnbox >> span >> box

        module Stg =
            let mapUnbox (xs: obj[]) : string[] = xs |> Array.map (unbox<string>)
            let count (xs: string[]) = xs.Length
            let distinct (xs: string[]) = xs |> Array.distinct |> Array.length
            let sum (xs: string[]) = String.Join(":", xs)
            let avg (xs: string[]) = ""
            let min (xs: string[]) = Array.min xs
            let max (xs: string[]) = Array.max xs
            let span (xs: string[]) = ""
            let aggregate (aoper: AOper) : (obj[] -> obj) =
                match aoper with
                | COUNT -> mapUnbox >> count >> box
                | DISTINCT -> mapUnbox >> distinct >> box
                | SUM -> mapUnbox >> sum >> box
                | AVG -> mapUnbox >> avg >> box
                | MIN -> mapUnbox >> min >> box
                | MAX -> mapUnbox >> max >> box
                | SPAN -> mapUnbox >> span >> box

        let aggregate (aoper: AOper) (operType: Type) : (obj[] -> obj) option =
                match operType with
                | x when x = typeof<double> -> Dbl.aggregate aoper |> Some
                | x when x = typeof<int> -> Intg.aggregate aoper |> Some
                | x when x = typeof<DateTime> -> Dte.aggregate aoper |> Some
                | x when x = typeof<string> -> Stg.aggregate aoper |> Some
                | _ -> None

    /// Summarize operator.
    /// operations = Aggregate operators, operNames, resultNames
    /// operator should be typeof<operName> -> resultType
    let summarizeAOper (r: Rel) (perRel: Rel) (operations: (AOper*string*string)[]) : Rel option =
        let aopers, operNames = operations |> Array.map (fun (aop, opNm, _) -> (aop, opNm)) |> Array.unzip

        match Field.types r.fields operNames with
        | None -> None
        | Some resTypes ->
            let mapTypes = Array.zip operNames resTypes |> Map.ofArray
            let newOperations = 
                operations 
                |> Array.choose
                    (fun (aop, opNm, resNm) ->
                         let operType = mapTypes |> Map.find opNm
                         let resType = aop.resultType operType |> Option.defaultValue (mapTypes |> Map.find opNm)
                         match AOper.aggregate aop operType with
                         | None -> None
                         | Some operFn ->
                             (operFn, opNm, resNm, resType) |> Some
                    )
            summarize' r perRel newOperations

    /// Un-pivot operator.
    let unpivot (r: Rel) (unpivotNames: Set<string>) (resultName: string) (resultValName: string) : Rel option =
        let perFields = Field.project r.fields true unpivotNames           
        let unpivFields = Field.project r.fields false unpivotNames
        let unpivTypes = unpivFields |> Set.toArray |> Array.map (fun field -> field.ftype) |> Array.distinct

        if unpivTypes.Length = 0 then
            Some r
        elif unpivTypes.Length > 1 then
            None
        else
            let unpivType = unpivTypes |> Array.head
            let perIdxs = Field.indices r.fields true unpivotNames
            let unpivIdxs = Field.indices r.fields false unpivotNames

            // unpivot relation fields
            let fields = Set.union perFields ([| { fname = resultName; ftype = typeof<string> }; { fname = resultValName; ftype = unpivType } |] |> Set.ofArray)
        
            // grouped relation body
            let unpivElem (unpivelem: Elem) : Elem[] = [| { fname = resultName; value = unpivelem.fname }; { fname = resultValName; value = unpivelem.value } |]
            let unpivRow (perrow: Row) (unpivrow: Row) : Row[] = unpivrow |> Array.map (fun elem -> Array.append perrow (unpivElem elem) |> Row.sort) 

            let body = 
                r.body 
                |> Set.toArray
                |> Array.map (fun row -> (Row.ofPositions perIdxs row, Row.ofPositions unpivIdxs row))
                |> Array.groupBy fst
                |> Array.collect (fun (perrow, tpldrows) -> (tpldrows |> Array.collect (fun (_, unpivrow) -> unpivRow perrow unpivrow)) )
                |> Set.ofArray

            let rel : Rel = { fields = fields; body = body }
            rel |> Some

    /// Extend operator.
    /// ofun is a FsharpFunc object.
    let extend (r: Rel) (ofun: obj) (argNames: string[]) (resultName: string) : Rel option =
        match Field.types r.fields argNames with
        | None -> None
        | Some argTypes ->
            if not (Fun.compatibleArgTypes ofun argTypes) then
                None
            elif r.fields |> Field.containsNames (resultName |> Set.singleton) then
                None
            else
                match Field.indicesOrdered r.fields argNames with
                | None -> None
                | Some argIdxs ->
                    let outputType = Fun.outputType ofun |> Option.get

                    // extended relation fields
                    let resFields = Set.union r.fields ({ fname = resultName; ftype = outputType } |> Set.singleton)
        
                    // extended relation body
                    let extendRows (rows: Row[]) : Row[] option = 
                        let getArgs (row: Row) = Row.ofPositions argIdxs row |> Array.map (fun elem -> elem.value)
                        let argss = rows |> Array.map getArgs
                        match Fun.applyMulti ofun argTypes argss with
                        | None -> None
                        | Some results ->
                            let resultElems = results |> Array.map (fun res -> { fname = resultName; value = res })
                            let resultRows = 
                                Array.zip rows resultElems 
                                |> Array.map (fun (row, elem) -> Array.append row [| elem |] |> Row.sort)
                            resultRows
                            |> Some
                
                    match r.body |> Set.toArray |> extendRows with
                    | None -> None
                    | Some extendedRows -> 
                        let resBody = extendedRows |> Set.ofArray
                        let rel : Rel = { fields = resFields; body = resBody }
                        rel |> Some

    /// Special case for the Extend operator.
    /// ofun is a FsharpFunc object.
    let extendMulti (r: Rel) (operation: string*(string[])*(obj[] -> obj)) = // (argNames: string[]) (resultName: string) : Rel option =
        

        None


    /// Restrict operator
    /// where the filtering test is limited to equality of the elements for given fields.
    /// names = [foo, foo], values = [x1, x2] 
    ///     => true for rows where foo element is either x1 OR x2 (exclude = false)
    ///     => true for rows where foo element is neither x1 nOR x2 (exclude = true)
    /// names = [foo, bar], values = [x, x] 
    ///     => true for row where foo element is x AND bar element is y (exclude = false)
    ///     => true for row where (foo, bar) is not equal to (x, y) (exclude = true)
    let restrictBasic (r: Rel) (exclude: bool) (names: string[]) (values: obj[]) : Rel option =        
        if (names.Length = 0) && (values.Length = 0) then
            Some r
        elif r.card = 0 then
            None
        elif names.Length <> values.Length then
            None
        else
            // map name -> values associated to the same name
            let mapFilter = Array.zip names values |> Array.groupBy fst |> Array.map (fun (name, pairs) -> (name, pairs |> Array.map snd)) |> Map.ofArray
            let filterNames = [| for kvp in mapFilter -> kvp.Key |]
            
            match Field.indicesOrdered r.fields filterNames with
            | None -> None
            | Some argIdxs ->
                let eqRow (row: Row) = 
                    let tests = Row.ofPositions argIdxs row |> Array.map (fun elem -> let vals = mapFilter |> Map.find elem.fname in vals |> Array.contains elem.value)
                    Array.TrueForAll(tests, fun test -> test)

                let eqRow (row: Row) = 
                    let tests = Row.ofPositions argIdxs row |> Array.map (fun elem -> let vals = mapFilter |> Map.find elem.fname in vals |> Array.contains elem.value)
                    Array.TrueForAll(tests, fun test -> test)

                let nonEqRow (row: Row) = 
                    let tests = Row.ofPositions argIdxs row |> Array.map (fun elem -> let vals = mapFilter |> Map.find elem.fname in vals |> Array.contains elem.value |> not)
                    Array.TrueForAll(tests, fun test -> test)

                let filterRow = if exclude then nonEqRow else eqRow

                let rel : Rel = { r with body = r.body |> Set.filter filterRow }
                rel |> Some

    /// Restrict operator.
    /// ofun is a FsharpFunc object.
    let restrict (r: Rel) (ofun: obj) (argNames: string[]) : Rel option =
        match Field.types r.fields argNames with
        | None -> None
        | Some argTypes ->
            if not (Fun.compatibleArgTypes ofun argTypes) then
                None
            else
                match Field.indicesOrdered r.fields argNames with
                | None -> None
                | Some argIdxs ->
        
                    // filtered relation body
                    let filterRows (rows: Row[]) : Row[] = 
                        let getArgs (row: Row) = Row.ofPositions argIdxs row |> Array.map (fun elem -> elem.value)
                        let argss = rows |> Array.map getArgs
                        Array.zip rows (Fun.filterMulti ofun argTypes argss)
                        |> Array.filter (fun (row, test) -> test)
                        |> Array.map fst
                        
                    let body = r.body |> Set.toArray |> filterRows |> Set.ofArray
                    let rel : Rel = { r with body = body }
                    rel |> Some
    
    let display (r: Rel) (showHead: bool) (unwrapOptions: bool) (refKey: String) (proxys: Out.Proxys) (rowFrom: int option) (rowCount: int option) (sortOn: string[]) (descending: bool) (midColumns: bool) (firstNames: string[]) (lastNames: string[]) : obj[,] =
        if r = Rel.DEE then
            box "<DEE>" |> Toolbox.Array2D.singleton
        elif r = Rel.DUM then
            box "<DUM>" |> Toolbox.Array2D.singleton
        elif r.card = 0 then
            box proxys.nan |> Toolbox.Array2D.singleton
        else
            // header names
            let allNames = r.fields |> Field.names |> Set.toArray
            let fstNames = firstNames |> Array.filter (fun name -> allNames |> Array.contains name)
            let lstNames = lastNames |> Array.filter (fun name -> (allNames |> Array.contains name) || (fstNames |> Array.contains name |> not))
            let midNames = 
                if midColumns then 
                    allNames |> Array.except (Array.append fstNames lstNames)            
                else
                    [||]
                    
            let headerNames = Array.append fstNames (Array.append midNames lastNames)
            let headerBxd = headerNames |> Array.map box
            let hdrIdxs = Field.indicesOrdered r.fields headerNames |> Option.get

            if headerNames.Length = 0 then
                Toolbox.Array2D.empty2D<obj>
            elif r.count = 0 then
                if showHead then
                    [| headerBxd |] |> array2D
                    // [| r.head |> Array.map box |] |> array2D
                else
                    box proxys.empty |> Toolbox.Array2D.singleton
            else
                // filter names
                let sortNames = sortOn |> Array.filter (fun name -> headerNames |> Array.contains name)

                // rows                
                let rowsSorted = 
                    let allRows = r.body |> Set.toArray
                    if sortNames.Length = 0 then
                        allRows
                    else
                        let sortIdxs = Field.indicesOrdered r.fields sortNames |> Option.get
                        let sortBy = if descending then Array.sortByDescending else Array.sortBy
                        allRows |> sortBy (Row.ofPositions sortIdxs)

                let rowFrm = rowFrom |> Option.defaultValue 0
                let rowCnt = rowCount |> Option.defaultValue rowsSorted.Length
                let rowsCropped = rowsSorted.[rowFrm .. (rowFrm + rowCnt - 1)]
                let rowsDisplayColumns = rowsCropped |> Array.map (Row.ofPositions hdrIdxs)
                let rowsBxd = rowsDisplayColumns |> Array.map (fun row -> row  |> Array.map (fun elem -> elem.value))

                if showHead then
                    Array.append [| headerBxd |] rowsBxd |> array2D
                else
                     rowsBxd |> array2D
            
    // -----------------------------
    // -- Conversion functions (from / to relations to other types)
    // -----------------------------

    /// Indicates whether the first and second rows contains a relation field names and types,
    /// Which need to be provided if not present as CSV file rows.
    type CSVFields = 
        | NameFstTypeSnd    /// field names and types are part of the csv file. Names on first row, types of second row.
        | NameSndTypeFst    /// field names and types are part of the csv file. Names on second row, types of first row.
        | NameFst of string[]   /// only field names are part of the csv file, types are provided as user input. Names on first row of csv file.
        | TypeFst of string[]   /// only field typess are part of the csv file, names are provided as user input. Types on first row of csv file.
        | NoHeader of (string[])*(string[]) /// neither field names or types are part of the csv file; they are provided as user input.

    /// (Field Names) * (Field Types) overrides.
    /// Type-tag override is either a type-tag array or a (field name -> type-tag) map.
    type FieldOvrR = (string[] option)*((string[] option)*(Map<string,string> option))
    type CSVFields2 = 
        | NameFstTypeSndX of FieldOvrR /// field names and types are part of the csv file. Names on first row, types of second row.
        | NameSndTypeFstX of FieldOvrR /// field names and types are part of the csv file. Names on second row, types of first row.
        | NameFstX of FieldOvrR        /// only field names are part of the csv file, types are provided as user input. Names on first row of csv file.
        | TypeFstX of FieldOvrR        /// only field typess are part of the csv file, names are provided as user input. Types on first row of csv file.
        | NoHeaderX of FieldOvrR       /// neither field names or types are part of the csv file; they are provided as user input.
        //with
        //    member this.ovrride (fieldOvrR: FieldOvrR) : CSVFields2 =
        //        match this with
        //        | NameFstTypeSndX _ -> NameFstTypeSndX fieldOvrR
        //        | NameSndTypeFstX _ -> NameSndTypeFstX fieldOvrR
        //        | NameFstX _ -> NameFstX fieldOvrR
        //        | TypeFstX _ -> TypeFstX fieldOvrR
        //        | NoHeaderX _ -> NoHeaderX fieldOvrR

    // TO DO, change function to accept external names (when not provided in the file)
    /// Converts a CSV file to a Rel object.
    /// The first or second row of the CSV file must provide the relation field names (*). 
    /// The first or second row of the CSV file can provide the relation field type tags (e.g. "int", "double", "string", "#date" ...) (*).
    /// (*) otherwise they need to be provided to the function
    /// If mapNameType is provided, then field types are derived from field names (the CSV row of types is ignored).
    /// If mapNameDefvals and mapTypeDefvals provide default values derived from field namess or types respectively (default based on field name prevails).
    let ofCSVX (mapTypeDefvals: Map<string,obj>) (mapNameDefvals: Map<string,obj>) (dateFormat: string option) (useVB: bool) (enclosingQuotes: bool) (trim: bool) (sep: string) (csvFields: CSVFields2) (fpath: string) : Rel option = 
        let lines = 
            if useVB then 
                Toolbox.CSV.readLinesVB enclosingQuotes trim sep fpath
            else
                Toolbox.CSV.readLines trim sep fpath
        
        let lineCount = lines |> Seq.length

        if lineCount = 0 then
            None
        else
            let skip n s = if n < lineCount then Seq.skip n s else [||] |> Seq.ofArray

            // determines names, types, body as specified by inputs.
            let (headerNames, headerTypeTags, bodyLines) =
                match csvFields with
                | NameFstTypeSndX (names, (types, mapTypes)) -> 
                    let nms = names |> Option.defaultValue (lines |> Seq.item 0)
                    let tys = 
                        match types, mapTypes with
                        | None, None -> lines |> Seq.item 1
                        | Some ts, None -> ts
                        | _, Some mapNmTy -> nms |> Array.map (fun nm -> mapNmTy |> Map.find nm)
                    (nms, tys, lines |> skip 2)
                | NameSndTypeFstX (names, (types, mapTypes)) -> 
                    let nms = names |> Option.defaultValue (lines |> Seq.item 1)
                    let tys = 
                        match types, mapTypes with
                        | None, None -> lines |> Seq.item 0
                        | Some ts, None -> ts
                        | _, Some mapNmTy -> nms |> Array.map (fun nm -> mapNmTy |> Map.find nm)
                    (nms, tys, lines |> skip 2)

                | NameFstX (names, (types, mapTypes)) -> 
                    let nms = names |> Option.defaultValue (lines |> Seq.item 0)
                    let tys = 
                        match types, mapTypes with
                        | None, None -> failwith "No CSV types provided"
                        | Some ts, None -> ts
                        | _, Some mapNmTy -> nms |> Array.map (fun nm -> mapNmTy |> Map.find nm)
                    (nms, tys, lines |> skip 1)

                | TypeFstX (names, (types, mapTypes)) -> 
                    let nms = 
                        match names with
                        | None -> failwith "No CSV names provided"
                        | Some ns -> ns
                    let tys = 
                        match types, mapTypes with
                        | None, None -> lines |> Seq.item 0
                        | Some ts, None -> ts
                        | _, Some mapNmTy -> nms |> Array.map (fun nm -> mapNmTy |> Map.find nm)
                    (nms, tys, lines |> skip 1)

                | NoHeaderX (names, (types, mapTypes)) -> 
                    let nms = 
                        match names with
                        | None -> failwith "No CSV names provided"
                        | Some ns -> ns
                    let tys = 
                        match types, mapTypes with
                        | None, None -> failwith "No CSV types provided"
                        | Some ts, None -> ts
                        | _, Some mapNmTy -> nms |> Array.map (fun nm -> mapNmTy |> Map.find nm)
                    (nms, tys, lines)

            if (headerNames.Length = 0) || (headerNames.Length <> headerTypeTags.Length) then
                None
            elif bodyLines |> Seq.filter (fun line -> line.Length <> headerNames.Length) |> Seq.isEmpty |> not then
                None
            else
                let headerTypes = headerTypeTags |> Array.map (API.Variant.labelType false)

                if (headerTypes.Length <> headerTypeTags.Length) then
                    None
                else
                    let relFields : Set<Field> = Field.zip headerNames headerTypes

                    let toElem (fname: string, typeTag: string, text: string) : Elem =
                        let defValue = 
                            match mapNameDefvals |> Map.tryFind fname with
                            | Some defval -> Some defval
                            | None -> mapTypeDefvals |> Map.tryFind (typeTag.ToUpper())
                        let res = API.Text.Tag.Any.def dateFormat defValue typeTag text
                        let elem : Elem = { fname = fname; value = res }
                        elem

                    let relBody : Set<Row> = 
                        bodyLines 
                        |> Seq.map (fun line -> Array.zip3 headerNames headerTypeTags line |> Array.map toElem |> Row.sort)
                        |> Set.ofSeq
                
                    let rel : Rel = { fields = relFields; body = relBody }
                    rel |> Some



    // TO DO, change function to accept external names (when not provided in the file)
    /// Converts a CSV file to a Rel object.
    /// The first or second row of the CSV file must provide the relation field names (*). 
    /// The first or second row of the CSV file can provide the relation field type tags (e.g. "int", "double", "string", "#date" ...) (*).
    /// (*) otherwise they need to be provided to the function
    /// If mapNameType is provided, then field types are derived from field names (the CSV row of types is ignored).
    /// If mapNameDefvals and mapTypeDefvals provide default values derived from field namess or types respectively (default based on field name prevails).
    let ofCSV (mapNameTagType: Map<string,string> option) (mapTypeDefvals: Map<string,obj>) (mapNameDefvals: Map<string,obj>) (dateFormat: string option) (useVB: bool) (enclosingQuotes: bool) (trim: bool) (sep: string) (csvFields: CSVFields) (fpath: string) : Rel option = 
        let lines = 
            if useVB then 
                Toolbox.CSV.readLinesVB enclosingQuotes trim sep fpath
            else
                Toolbox.CSV.readLines trim sep fpath
        
        let lineCount = lines |> Seq.length

        if lineCount = 0 then
            None
        else
            let skip n s = if n < lineCount then Seq.skip n s else [||] |> Seq.ofArray

            // determines names, types, body as specified by inputs.
            let (headerNames, headerTypeTags, bodyLines) =
                match mapNameTagType, csvFields with
                // Names from Csv file. (Name -> Type) provided. Types are derived from names. Csv file's type row is ignored.
                | Some mapNmTy, NameFstTypeSnd ->
                    let nms = lines |> Seq.item 0
                    let tys = nms |> Array.map (fun nm -> mapNmTy |> Map.find nm)
                    (nms, tys, lines |> skip 1)

                // Names from Csv file. (Name -> Type) provided. Types are derived from names. Csv file's type row is ignored.
                | Some mapNmTy, NameSndTypeFst ->
                    let nms = lines |> Seq.item 1
                    let tys = nms |> Array.map (fun nm -> mapNmTy |> Map.find nm)
                    (nms, tys, lines |> skip 1)

                // Names from Csv file. (Name -> Type) provided. Types are derived from names.
                | Some mapNmTy, NameFst _ -> 
                    let nms = lines |> Seq.item 0
                    let tys = nms |> Array.map (fun nm -> mapNmTy |> Map.find nm)
                    (nms, tys, lines |> skip 1)

                // Names from inputs. (Name -> Type) provided. Types are derived from names. Csv file's type row is ignored.
                | Some mapNmTy, TypeFst nms -> 
                    let tys = nms |> Array.map (fun nm -> mapNmTy |> Map.find nm)
                    (nms, tys, lines |> skip 1)

                // Names from inputs. (Name -> Type) provided. Types are derived from names.
                | Some mapNmTy, NoHeader (nms, _) -> 
                    let tys = nms |> Array.map (fun nm -> mapNmTy |> Map.find nm)
                    (nms, tys, lines)

                // Names and type from Csv file.
                | None, NameFstTypeSnd ->
                    let nms = lines |> Seq.item 0
                    let tys = lines |> Seq.item 1
                    (nms, tys, lines |> skip 2)

                // Names and type from Csv file.
                | None, NameSndTypeFst ->
                    let tys = lines |> Seq.item 0
                    let nms = lines |> Seq.item 1
                    (nms, tys, lines |> skip 2)

                // Names from Csv file, types from inputs.
                | None, NameFst tys -> 
                    let nms = lines |> Seq.item 0
                    (nms, tys, lines |> skip 1)

                // Names from inputs, types from Csv file.
                | None, TypeFst nms -> 
                    let tys = lines |> Seq.item 0
                    (nms, tys, lines |> skip 1)

                // Names and types from inputs.
                | None, NoHeader (nms, tys) -> (nms, tys, lines)

            if (headerNames.Length = 0) || (headerNames.Length <> headerTypeTags.Length) then
                None
            elif bodyLines |> Seq.filter (fun line -> line.Length <> headerNames.Length) |> Seq.isEmpty |> not then
                None
            else
                let headerTypes = headerTypeTags |> Array.map (API.Variant.labelType false)

                if (headerTypes.Length <> headerTypeTags.Length) then
                    None
                else
                    let relFields : Set<Field> = Field.zip headerNames headerTypes

                    let toElem (fname: string, typeTag: string, text: string) : Elem =
                        let defValue = 
                            match mapNameDefvals |> Map.tryFind fname with
                            | Some defval -> Some defval
                            | None -> mapTypeDefvals |> Map.tryFind (typeTag.ToUpper())
                        let res = API.Text.Tag.Any.def dateFormat defValue typeTag text
                        let elem : Elem = { fname = fname; value = res }
                        elem

                    let relBody : Set<Row> = 
                        bodyLines 
                        |> Seq.map (fun line -> Array.zip3 headerNames headerTypeTags line |> Array.map toElem |> Row.sort)
                        |> Set.ofSeq
                
                    let rel : Rel = { fields = relFields; body = relBody }
                    rel |> Some

    let ofMap1<'key1,'value when 'key1: comparison> 
        (map: Map<'key1,'value>) (fnameKey1: string) (fnameValue: string) : Rel = 
        let fieldNames = [| fnameKey1; fnameValue |]
        let relFields = Field.zip fieldNames [| typeof<'key1>; typeof<'value> |]

        let relBody =
            [| for kvp in map ->
                let k1 = kvp.Key
                let row : Row = Elem.zip fieldNames [| k1; kvp.Value |]
                row
            |] |> Set.ofArray

        { fields = relFields; body = relBody }

    let ofMap2<'key1,'key2,'value when 'key1: comparison and 'key2: comparison> 
        (map: Map<'key1*'key2,'value>) (fnameKey1: string) (fnameKey2: string) (fnameValue: string) : Rel = 
        let fieldNames = [| fnameKey1; fnameKey2; fnameValue |]
        let relFields = Field.zip fieldNames [| typeof<'key1>; typeof<'key2>; typeof<'value> |]

        let relBody =
            [| for kvp in map ->
                let (k1, k2) = kvp.Key
                let row : Row = Elem.zip fieldNames [| k1; k2; kvp.Value |]
                row
            |] |> Set.ofArray

        { fields = relFields; body = relBody }

    let ofMap3<'key1,'key2,'key3,'value when 'key1: comparison and 'key2: comparison and 'key3: comparison> 
        (map: Map<'key1*'key2*'key3,'value>) 
        (fnameKey1: string) (fnameKey2: string) (fnameKey3: string) (fnameValue: string) 
        : Rel = 
        let fieldNames = [| fnameKey1; fnameKey2; fnameKey3; fnameValue |]
        let relFields = Field.zip fieldNames [| typeof<'key1>; typeof<'key2>; typeof<'key3>; typeof<'value> |]

        let relBody =
            [| for kvp in map ->
                let (k1, k2, k3) = kvp.Key
                let row : Row = Elem.zip fieldNames [| k1; k2; k3; kvp.Value |]
                row
            |] |> Set.ofArray

        { fields = relFields; body = relBody }

    let ofMap4<'key1,'key2,'key3,'key4,'value when 'key1: comparison and 'key2: comparison and 'key3: comparison and 'key4: comparison> 
        (map: Map<'key1*'key2*'key3*'key4,'value>) 
        (fnameKey1: string) (fnameKey2: string) (fnameKey3: string) (fnameKey4: string) (fnameValue: string) 
        : Rel = 
        let fieldNames = [| fnameKey1; fnameKey2; fnameKey3; fnameKey4; fnameValue |]
        let relFields = Field.zip fieldNames [| typeof<'key1>; typeof<'key2>; typeof<'key3>; typeof<'key4>; typeof<'value> |]

        let relBody =
            [| for kvp in map ->
                let (k1, k2, k3, k4) = kvp.Key
                let row : Row = Elem.zip fieldNames [| k1; k2; k3; k4; kvp.Value |]
                row
            |] |> Set.ofArray

        { fields = relFields; body = relBody }

    let ofMap5<'key1,'key2,'key3,'key4,'key5,'value when 'key1: comparison and 'key2: comparison and 'key3: comparison and 'key4: comparison and 'key5: comparison> 
        (map: Map<'key1*'key2*'key3*'key4*'key5,'value>) 
        (fnameKey1: string) (fnameKey2: string) (fnameKey3: string) (fnameKey4: string) (fnameKey5: string) (fnameValue: string) 
        : Rel = 
        let fieldNames = [| fnameKey1; fnameKey2; fnameKey3; fnameKey4; fnameKey5; fnameValue |]
        let relFields = Field.zip fieldNames [| typeof<'key1>; typeof<'key2>; typeof<'key3>; typeof<'key4>; typeof<'key5>; typeof<'value> |]

        let relBody =
            [| for kvp in map ->
                let (k1, k2, k3, k4, k5) = kvp.Key
                let row : Row = Elem.zip fieldNames [| k1; k2; k3; k4; k5; kvp.Value |]
                row
            |] |> Set.ofArray

        { fields = relFields; body = relBody }

    let ofMap6<'key1,'key2,'key3,'key4,'key5,'key6,'value when 'key1: comparison and 'key2: comparison and 'key3: comparison and 'key4: comparison and 'key5: comparison and 'key6: comparison> 
        (map: Map<'key1*'key2*'key3*'key4*'key5*'key6,'value>) 
        (fnameKey1: string) (fnameKey2: string) (fnameKey3: string) (fnameKey4: string) (fnameKey5: string) (fnameKey6: string) (fnameValue: string) 
        : Rel = 
        let fieldNames = [| fnameKey1; fnameKey2; fnameKey3; fnameKey4; fnameKey5; fnameKey6; fnameValue |]
        let relFields = Field.zip fieldNames [| typeof<'key1>; typeof<'key2>; typeof<'key3>; typeof<'key4>; typeof<'key5>; typeof<'key6>; typeof<'value> |]

        let relBody =
            [| for kvp in map ->
                let (k1, k2, k3, k4, k5, k6) = kvp.Key
                let row : Row = Elem.zip fieldNames [| k1; k2; k3; k4; k5; k6; kvp.Value |]
                row
            |] |> Set.ofArray

        { fields = relFields; body = relBody }

    let ofMap7<'key1,'key2,'key3,'key4,'key5,'key6,'key7,'value when 'key1: comparison and 'key2: comparison and 'key3: comparison and 'key4: comparison and 'key5: comparison and 'key6: comparison and 'key7: comparison> 
        (map: Map<'key1*'key2*'key3*'key4*'key5*'key6*'key7,'value>) 
        (fnameKey1: string) (fnameKey2: string) (fnameKey3: string) (fnameKey4: string) (fnameKey5: string) (fnameKey6: string) (fnameKey7: string) (fnameValue: string) 
        : Rel = 
        let fieldNames = [| fnameKey1; fnameKey2; fnameKey3; fnameKey4; fnameKey5; fnameKey6; fnameKey7; fnameValue |]
        let relFields = Field.zip fieldNames [| typeof<'key1>; typeof<'key2>; typeof<'key3>; typeof<'key4>; typeof<'key5>; typeof<'key6>; typeof<'key7>; typeof<'value> |]

        let relBody =
            [| for kvp in map ->
                let (k1, k2, k3, k4, k5, k6, k7) = kvp.Key
                let row : Row = Elem.zip fieldNames [| k1; k2; k3; k4; k5; k6; k7; kvp.Value |]
                row
            |] |> Set.ofArray

        { fields = relFields; body = relBody }
    
    module Registry =
        open Registry
        open Toolbox.Generics
        open Microsoft.FSharp.Reflection

        let genType = typeof<Rel>
        let genTypeMap = typeof<Map<_,_>>

        type GenFn =
            static member ofMap1<'K1,'V when 'K1: comparison> 
                (map: Map<'K1,'V>) 
                (fnameKey1: string) (fnameValue: string) 
                : Rel =
                ofMap1 map fnameKey1 fnameValue

            static member ofMap2<'K1,'K2,'V when 'K1: comparison and 'K2: comparison> 
                (map: Map<'K1*'K2,'V>) 
                (fnameKey1: string) (fnameKey2: string) (fnameValue: string) 
                : Rel =
                ofMap2 map fnameKey1 fnameKey2 fnameValue

            static member ofMap3<'K1,'K2,'K3,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison> 
                (map: Map<'K1*'K2*'K3,'V>) 
                (fnameKey1: string) (fnameKey2: string) (fnameKey3: string) (fnameValue: string) 
                : Rel =
                ofMap3 map fnameKey1 fnameKey2 fnameKey3 fnameValue

            static member ofMap4<'K1,'K2,'K3,'K4,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison> 
                (map: Map<'K1*'K2*'K3*'K4,'V>) 
                (fnameKey1: string) (fnameKey2: string) (fnameKey3: string) (fnameKey4: string) (fnameValue: string) 
                : Rel =
                ofMap4 map fnameKey1 fnameKey2 fnameKey3 fnameKey4 fnameValue

            static member ofMap5<'K1,'K2,'K3,'K4,'K5,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison> 
                (map: Map<'K1*'K2*'K3*'K4*'K5,'V>) 
                (fnameKey1: string) (fnameKey2: string) (fnameKey3: string) (fnameKey4: string) (fnameKey5: string) (fnameValue: string) 
                : Rel =
                ofMap5 map fnameKey1 fnameKey2 fnameKey3 fnameKey4 fnameKey5 fnameValue

            static member ofMap6<'K1,'K2,'K3,'K4,'K5,'K6,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison> 
                (map: Map<'K1*'K2*'K3*'K4*'K5*'K6,'V>) 
                (fnameKey1: string) (fnameKey2: string) (fnameKey3: string) (fnameKey4: string) (fnameKey5: string) (fnameKey6: string) (fnameValue: string) 
                : Rel =
                ofMap6 map fnameKey1 fnameKey2 fnameKey3 fnameKey4 fnameKey5 fnameKey6 fnameValue

            static member ofMap7<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison> 
                (map: Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7,'V>) 
                (fnameKey1: string) (fnameKey2: string) (fnameKey3: string) (fnameKey4: string) (fnameKey5: string) (fnameKey6: string) (fnameKey7: string) (fnameValue: string) 
                : Rel =
                ofMap7 map fnameKey1 fnameKey2 fnameKey3 fnameKey4 fnameKey5 fnameKey6 fnameKey7 fnameValue

        module In =
            let MAX_ARITY_TOMAP = 8
            /// Builds a map from a relation R-object and its keys and value columns.
            /// groupValues allows to return several values corresponding to identical keys, otherwise only one value is returned and the others are ignored.
            /// groupValues = true
            ///    - result map type : ('a1, 'a2, ..., 'an) -> 'b[]
            ///    - [| ("foo", 1); "value1" |] and [| ("foo", 1); "value2" |] relation rows will produce a map key-value pair (("foo", 1), [| "value1"; "value2" |])
            /// groupValues = false
            ///    - result map type : ('a1, 'a2, ..., 'an) -> 'b
            ///    - [| ("foo", 1); "value1" |] and [| ("foo", 1); "value2" |] relation rows will produce a map key-value pair (("foo", 1), "value1")
            let toMap (r: Rel) (groupValues: bool) (keysNames: string[]) (valueName: string) : obj option =
                if keysNames.Length = 0 then
                    None
                else
                    let keepNames = Array.append keysNames [| valueName |]
                    match Field.indicesOrdered r.fields keepNames with
                    | None -> None
                    | Some keepIdxs ->
                        let arity = keepIdxs.Length - 1
                        if (arity < 1) || (arity > MAX_ARITY_TOMAP) then
                            None
                        else                                
                            let rows = r.body |> Set.toArray |> Array.map (Row.ofPositions keepIdxs) |> Array.distinct
                            let colkeys = 
                                [| for j in 0 .. arity - 1 -> 
                                    [| for i in 0 .. rows.Length - 1 -> rows.[i] |> Row.item j |]
                                |]
                            let colval = rows |> Array.map (Row.item arity)

                            // let methodNm = sprintf "omap%d%s" arity (if groupValues then "s" else "")
                            let methodTys = let flds = r.fields |> Set.toArray in keepIdxs |> Array.map (fun i -> flds.[i].ftype)
                            let resMap = MAP.RxFn.In.omapN (Array.sub methodTys 0 arity) (methodTys |> Array.last) (colkeys |> Array.map box) colval
                            resMap |> Some


            let toMapOLD (regKey: string) (groupValues: bool) (keysNames: string[]) (valueName: string) : obj option =
                if keysNames.Length = 0 then
                    None
                else
                    match MRegistry.tryExtract<Rel> regKey with
                    | None -> None
                    | Some rel ->
                        let keepNames = Array.append keysNames [| valueName |]
                        match Field.indicesOrdered rel.fields keepNames with
                        | None -> None
                        | Some keepIdxs ->
                            let arity = keepIdxs.Length - 1
                            if (arity < 2) || (arity > MAX_ARITY_TOMAP) then
                                None
                            else                                
                                let rows = rel.body |> Set.toArray |> Array.map (Row.ofPositions keepIdxs) |> Array.distinct
                                let colkeys = 
                                    [| for j in 0 .. keepIdxs.Length - 2 -> 
                                        [| for i in 0 .. rows.Length - 1 -> rows.[i] |> Row.item j |]
                                    |]
                                // let colval = rows |> Array.map (Row.item (keepIdxs |> Array.last))
                                let colval = rows |> Array.map Array.last

                                let methodNm = sprintf "omap%d%s" arity (if groupValues then "s" else "")
                                let methodTys = let flds = rel.fields |> Set.toArray in keepIdxs |> Array.map (fun i -> flds.[i].ftype)
                                let resMap = MAP.RxFn.In.mapN (Array.sub methodTys 0 arity) (methodTys |> Array.last) (colkeys |> Array.map box) colval
                                resMap |> Some

        module Out =
            let MAX_ARITY_OFMAP = 7
            let ofMap (regKey: string) (splitTuple: bool) (fnameKeys: string[]) (fnameValue: string) : obj option = 
                match MRegistry.tryExtractGen genTypeMap regKey with
                | None -> None
                | Some (tys, o) -> 
                    // tys is a [| the map-object's key type; the map-object's value type |], only 2 elements.
                    if not (FSharpType.IsTuple tys.[0] && splitTuple) then
                        if fnameKeys.Length = 0 then
                            None
                        else
                            let methodNm = "ofMap1"
                            let args : obj[] = [| fnameKeys.[0]; box fnameValue |]
                            let genTypeRObj = (tys, o)
                            apply<GenFn> methodNm [||] args genTypeRObj
                            |> Some
                    else
                        let elemTys = FSharpType.GetTupleElements(tys.[0])
                        let arity = elemTys.Length

                        if (arity > fnameKeys.Length) || (arity > MAX_ARITY_OFMAP) then
                            None
                        else
                            let methodNm = sprintf "ofMap%d" arity
                            let args : obj[] = Array.append (fnameKeys |> Array.take arity |> Array.map box) [| box fnameValue |]
                            let genTypeRObj = (Array.append elemTys [| tys.[1] |], o)
                            apply<GenFn> methodNm [||] args genTypeRObj
                            |> Some

        // toCSV, pool


module Rel_XL =
    open Registry
    open API
    open API.Out
    open type Out.Proxys

    let tryExtract<'a> = MRegistry.tryExtract<'a>
    let tryExtractO = MRegistry.tryExtractO

    [<ExcelFunction(Category="Relation", Description="Returns a relation R-object from Excel.")>]
    let rel_ofRng
        ([<ExcelArgument(Description= "Field names.")>] fieldNames: obj[])
        ([<ExcelArgument(Description= "Field type tags: bool, date, double, doubleNaN, string or obj. Add \'#'\ prefix for optional type: #bool, #date, #double, #doubleNaN, #string or #obj")>] typeTags: obj[])
        ([<ExcelArgument(Description= "2D xl-range.")>] range: obj[,])
        ([<ExcelArgument(Description= "[Only for doubleNaN tags: Kinds for which values are converted to Double.NaN. E.g. NA, ERR, TXT, !NUM... Default is \"NA\".]")>] xlKinds: obj)
        : obj  =

        // intermediary stage
        let fieldNames = API.In.D1.Stg.tryDV None fieldNames
        let typeTags = API.In.D1.Stg.tryDV None typeTags
        let xlkinds = In.D0.Stg.def "NA" xlKinds |> Kind.ofLabel

        // caller cell's reference ID
        let rfid = MRegistry.refID
        
        // result
        match fieldNames, typeTags with
        | None, _ -> Proxys.def.failed
        | _, None -> Proxys.def.failed
        | Some fieldnames, Some typetags ->
            match Rel.Rel.ofRng xlkinds fieldnames typetags range with
            | None -> Proxys.def.failed
            | Some rel ->
                rel |> MRegistry.registerBxd rfid 

    [<ExcelFunction(Category="Relation", Description="Returns a (row-wise) empty relation R-object from Excel.")>]
    let rel_ofHead
        ([<ExcelArgument(Description= "Field names.")>] fieldNames: obj[])
        ([<ExcelArgument(Description= "Field type tags: bool, date, double, doubleNaN, string or obj. Add \'#'\ prefix for optional type: #bool, #date, #double, #doubleNaN, #string or #obj")>] typeTags: obj[])
        : obj  =

        // intermediary stage
        let fieldNames = API.In.D1.Stg.tryDV None fieldNames
        let typeTags = API.In.D1.Stg.tryDV None typeTags
        let xlkinds = "NA" |> Kind.ofLabel

        // caller cell's reference ID
        let rfid = MRegistry.refID
        
        // result
        match fieldNames, typeTags with
        | None, _ -> Proxys.def.failed
        | _, None -> Proxys.def.failed
        | Some fieldnames, Some typetags ->
            match Rel.Rel.ofHead fieldnames typetags with
            | None -> Proxys.def.failed
            | Some rel ->
                rel |> MRegistry.registerBxd rfid 

    [<ExcelFunction(Category="Relation", Description="Returns a Relation DEE R-object.")>]
    let rel_DEE () : obj  =
        // caller cell's reference ID
        let rfid = MRegistry.refID

        Rel.Rel.DEE |> MRegistry.registerBxd rfid

    [<ExcelFunction(Category="Relation", Description="Returns a Relation DUM R-object.")>]
    let rel_DUM () : obj  =
        // caller cell's reference ID
        let rfid = MRegistry.refID

        Rel.Rel.DUM |> MRegistry.registerBxd rfid

    [<ExcelFunction(Category="Relation", Description="Extracts a relation.")>]
    let rel_toRng
        ([<ExcelArgument(Description= "Relation R-object.")>] rgRel: string)
        ([<ExcelArgument(Description= "[Show head. Default is true.]")>] showHead: obj)
        ([<ExcelArgument(Description= "[None indicator. Default is \"<none>\".]")>] noneIndicator: obj)
        ([<ExcelArgument(Description= "[Empty array indicator. Default is \"<empty>\".]")>] emptyIndicator: obj)
        ([<ExcelArgument(Description= "[Unwrap optional types. Default is true.]")>] unwrapOptions: obj)
        : obj[,] = 

        // intermediary stage
        let showHead = In.D0.Bool.def true showHead
        let none = In.D0.Stg.def "<none>" noneIndicator
        let empty = In.D0.Stg.def "<empty>" emptyIndicator
        let proxys = { def with none = none; empty = empty }
        let unwrapoptions = In.D0.Bool.def true unwrapOptions

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel with
        | None -> Array2D.create 1 1 Proxys.def.failed
        | Some rel ->
            rel.toRng showHead unwrapoptions rfid proxys

    [<ExcelFunction(Category="Relation", Description="Displays a relation.")>]
    let rel_display
        ([<ExcelArgument(Description= "Relation R-object.")>] rgRel: string)
        ([<ExcelArgument(Description= "Sorting fields.")>] sortingFields: obj)
        ([<ExcelArgument(Description= "[Descending sort. Default if false.]")>] descending: obj)
        ([<ExcelArgument(Description= "[Start row. Default is 0.]")>] startRow: obj)
        ([<ExcelArgument(Description= "[Row count. Default is all rows.]")>] rowCount: obj)
        ([<ExcelArgument(Description= "[Col. ordering. First columns. Default is none.]")>] firstCols: obj)
        ([<ExcelArgument(Description= "[Col. ordering. Last columns. Default is none.]")>] lastCols: obj)
        ([<ExcelArgument(Description= "[Mid columns. Only stated columns appear if true. Default is true.]")>] midCols: obj)
        ([<ExcelArgument(Description= "[Show head. Default is true.]")>] showHead: obj)
        ([<ExcelArgument(Description= "[None indicator. Default is \"<none>\".]")>] noneIndicator: obj)
        ([<ExcelArgument(Description= "[Empty array indicator. Default is \"<empty>\".]")>] emptyIndicator: obj)
        ([<ExcelArgument(Description= "[Unwrap optional types. Default is true.]")>] unwrapOptions: obj)
        : obj[,] = 
        
        // intermediary stage
        let sortNames = API.In.D1.OStg.tryDV None sortingFields |> Option.defaultValue [||]
        let descending = In.D0.Bool.def false descending
        let startRow = In.D0.Intg.Opt.def None startRow
        let rowCount = In.D0.Intg.Opt.def None rowCount
        let firstCols = API.In.D1.OStg.tryDV None firstCols |> Option.defaultValue [||]
        let lastCols = API.In.D1.OStg.tryDV None lastCols |> Option.defaultValue [||]
        let midCols = In.D0.Bool.def true midCols
        let showHead = In.D0.Bool.def true showHead
        let none = In.D0.Stg.def "<none>" noneIndicator
        let empty = In.D0.Stg.def "<empty>" emptyIndicator
        let proxys = { def with none = none; empty = empty }
        let unwrapoptions = In.D0.Bool.def true unwrapOptions

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel with
        | None -> Toolbox.Array2D.singleton Proxys.def.failed
        | Some rel -> 
            let relDspl = Rel.display rel showHead unwrapoptions rfid proxys startRow rowCount sortNames descending midCols firstCols lastCols
            relDspl

    [<ExcelFunction(Category="Relation", Description="Returns a relation number of fields / attributes.")>]
    let rel_card
        ([<ExcelArgument(Description= "Relation R-object.")>] rgRel: string)
        : obj = 

        // result
        match tryExtract<Rel.Rel> rgRel with
        | None -> Proxys.def.failed
        | Some rel -> box rel.card

    [<ExcelFunction(Category="Relation", Description="Returns a relation number of rows / tuples.")>]
    let rel_count
        ([<ExcelArgument(Description= "Relation R-object.")>] rgRel: string)
        : obj = 

        // result
        match tryExtract<Rel.Rel> rgRel with
        | None -> Proxys.def.failed
        | Some rel -> box rel.count

    [<ExcelFunction(Category="Relation", Description="Returns the union of 2 compatible relations.")>]
    let rel_union
        ([<ExcelArgument(Description= "Relation R-object 1.")>] rgRel1: string)
        ([<ExcelArgument(Description= "Relation R-object 2.")>] rgRel2: string)
        : obj = 

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel1, tryExtract<Rel.Rel> rgRel2 with
        | Some rel1, Some rel2 -> 
            match Rel.union rel1 rel2 with
            | None -> Proxys.def.failed
            | Some relUnion -> relUnion |> MRegistry.registerBxd rfid
        | _ -> Proxys.def.failed

    [<ExcelFunction(Category="Relation", Description="Returns the intersection of 2 compatible relations.")>]
    let rel_inter
        ([<ExcelArgument(Description= "Relation R-object 1.")>] rgRel1: string)
        ([<ExcelArgument(Description= "Relation R-object 2.")>] rgRel2: string)
        : obj = 

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel1, tryExtract<Rel.Rel> rgRel2 with
        | Some rel1, Some rel2 -> 
            match Rel.inter rel1 rel2 with
            | None -> Proxys.def.failed
            | Some relInter -> relInter |> MRegistry.registerBxd rfid
        | _ -> Proxys.def.failed

    [<ExcelFunction(Category="Relation", Description="Returns the difference between relation 1 and 2.")>]
    let rel_minus
        ([<ExcelArgument(Description= "Relation R-object 1.")>] rgRel1: string)
        ([<ExcelArgument(Description= "Relation R-object 2.")>] rgRel2: string)
        : obj = 

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel1, tryExtract<Rel.Rel> rgRel2 with
        | Some rel1, Some rel2 -> 
            match Rel.minus rel1 rel2 with
            | None -> Proxys.def.failed
            | Some relMinus -> relMinus |> MRegistry.registerBxd rfid
        | _ -> Proxys.def.failed

    [<ExcelFunction(Category="Relation", Description="Returns the product of 2 disjoint relations.")>]
    let rel_prod
        ([<ExcelArgument(Description= "Relation R-object 1.")>] rgRel1: string)
        ([<ExcelArgument(Description= "Relation R-object 2.")>] rgRel2: string)
        : obj = 

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel1, tryExtract<Rel.Rel> rgRel2 with
        | Some rel1, Some rel2 -> 
            match Rel.prod rel1 rel2 with
            | None -> Proxys.def.failed
            | Some relProd -> relProd |> MRegistry.registerBxd rfid
        | _ -> Proxys.def.failed

    [<ExcelFunction(Category="Relation", Description="Returns the join operation of 2 relations.")>]
    let rel_join
        ([<ExcelArgument(Description= "Relation R-object 1.")>] rgRel1: string)
        ([<ExcelArgument(Description= "Relation R-object 2.")>] rgRel2: string)
        : obj = 

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel1, tryExtract<Rel.Rel> rgRel2 with
        | Some rel1, Some rel2 -> 
            match Rel.join rel1 rel2 with
            | None -> Proxys.def.failed
            | Some relJoin -> relJoin |> MRegistry.registerBxd rfid
        | _ -> Proxys.def.failed

    [<ExcelFunction(Category="Relation", Description="Project a relation onto the given fields.")>]
    let rel_project
        ([<ExcelArgument(Description= "Relation R-object.")>] rgRel: string)
        ([<ExcelArgument(Description= "Project field names.")>] projectNames: obj)
        ([<ExcelArgument(Description= "[Project fields away (and keep the others). Default is false.]")>] projectAway: obj)
        : obj = 

        // intermediary stage
        let projNames = In.D1.OStg.tryDV None projectNames
        let projAway = In.D0.Bool.def false projectAway

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel, projNames with
        | Some rel, Some keepNames -> 
            let relProj = Rel.project rel projAway (keepNames |> Set.ofArray)
            relProj |> MRegistry.registerBxd rfid
        | _ -> Proxys.def.failed

    [<ExcelFunction(Category="Relation", Description="Returns the semi-difference between relation 1 and 2.")>]
    let rel_sminus
        ([<ExcelArgument(Description= "Relation R-object 1.")>] rgRel1: string)
        ([<ExcelArgument(Description= "Relation R-object 2.")>] rgRel2: string)
        : obj = 

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel1, tryExtract<Rel.Rel> rgRel2 with
        | Some rel1, Some rel2 -> 
            match Rel.semiMinus rel1 rel2 with
            | None -> Proxys.def.failed
            | Some relSMinus -> relSMinus |> MRegistry.registerBxd rfid
        | _ -> Proxys.def.failed

    [<ExcelFunction(Category="Relation", Description="Returns the semi-join operation of 2 relations.")>]
    let rel_sjoin
        ([<ExcelArgument(Description= "Relation R-object 1.")>] rgRel1: string)
        ([<ExcelArgument(Description= "Relation R-object 2.")>] rgRel2: string)
        : obj = 

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel1, tryExtract<Rel.Rel> rgRel2 with
        | Some rel1, Some rel2 -> 
            match Rel.semiJoin rel1 rel2 with
            | None -> Proxys.def.failed
            | Some relSJoin -> relSJoin |> MRegistry.registerBxd rfid
        | _ -> Proxys.def.failed

    [<ExcelFunction(Category="Relation", Description="Returns the left (outer) join of 2 relations.")>]
    let rel_ljoin
        ([<ExcelArgument(Description= "Relation R-object 1.")>] rgRel1: string)
        ([<ExcelArgument(Description= "Relation R-object 2.")>] rgRel2: string)
        ([<ExcelArgument(Description= "[Bool default value. Default is 0.]")>] boolDef: obj)
        ([<ExcelArgument(Description= "[Int default value. Default is 0.]")>] intDef: obj)
        ([<ExcelArgument(Description= "[Double default value. Default is 0.0.]")>] doubleDef: obj)
        ([<ExcelArgument(Description= "[Date default value. Default is 1 Jan 2000.]")>] dateDef: obj)
        ([<ExcelArgument(Description= "[String default value. Default is \"\".]")>] stringDef: obj)
        : obj = 
        
        // intermediary stage
        let boolDef = In.D0.Bool.def false boolDef
        let intDef = In.D0.Intg.def 0 intDef
        let doubleDef = In.D0.Dbl.def 0.0 doubleDef // FIXME allow NaN
        let dateDef = In.D0.Dte.def (DateTime(2000,1,1)) dateDef
        let stringDef = In.D0.Stg.def "" stringDef
        let defaultValues (field: Rel.Field) = 
            match field.ftype with
            | x when x = typeof<bool> -> box boolDef
            | x when x = typeof<int> -> box intDef
            | x when x = typeof<double> -> box doubleDef
            | x when x = typeof<DateTime> -> box dateDef
            | x when x = typeof<string> -> box stringDef
            | _ -> failwith "NOT IMPLEMENTED YET"

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel1, tryExtract<Rel.Rel> rgRel2 with
        | Some rel1, Some rel2 -> 
            match Rel.leftJoin defaultValues rel1 rel2 with
            | None -> Proxys.def.failed
            | Some relLJoin -> relLJoin |> MRegistry.registerBxd rfid
        | _ -> Proxys.def.failed


    [<ExcelFunction(Category="Relation", Description="Relation summarize operation.")>]
    let rel_sum
        ([<ExcelArgument(Description= "Relation R-object.")>] rgRel: string)
        ([<ExcelArgument(Description= "Per relation R-object.")>] rgPer: string)  
        ([<ExcelArgument(Description= "Operand field names.")>] operNames: obj)
        ([<ExcelArgument(Description= "[Aggregate operators. COUNT, DISTINCT, SUM, AVG, MIN, MAX, SPAN. Default is COUNT.]")>] aggOperators: obj)
        ([<ExcelArgument(Description= "[Result field names. Default is \"OPERAND_OPERATOR\"")>] resultNames: obj)
        : obj = 

        // intermediary stage
        let operNames = In.D1.OStg.filter operNames
        let aggOperators = In.D1.OStg.filter aggOperators |> Array.map Rel.AOper.ofLabel
        let resultNames = In.D1.OStg.tryDV None resultNames

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel, tryExtract<Rel.Rel> rgPer with
        | None, _ -> Proxys.def.failed
        | _, None -> Proxys.def.failed
        | Some rel, Some relPer -> 
            let defResultNames = Array.map2 (fun nm op -> sprintf "%s %s" nm (op.ToString())) operNames aggOperators
            let resultNames = resultNames |> Option.defaultValue defResultNames
            let operations = Toolbox.Array.zip3 aggOperators operNames resultNames

            match Rel.summarizeAOper rel relPer operations with
            | None -> Proxys.def.failed
            | Some relSumm -> relSumm |> MRegistry.registerBxd rfid

    [<ExcelFunction(Category="Relation", Description="Returns a relation un-pivotted.")>]
    let rel_unpivot
        ([<ExcelArgument(Description= "Relation R-object.")>] rgRel: string)
        ([<ExcelArgument(Description= "Un-pivot field names.")>] unpivotNames: obj)
        ([<ExcelArgument(Description= "Key field name.")>] keyField: string)
        ([<ExcelArgument(Description= "Value field name.")>] valueField: string)
        : obj = 
        
        // intermediary stage
        let unpivotNames = In.D1.OStg.tryDV None unpivotNames

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel, unpivotNames with
        | Some rel, Some unpivNames -> 
            match Rel.unpivot rel (unpivNames |> Set.ofArray) keyField valueField with
            | None -> Proxys.def.failed
            | Some relUpivot -> relUpivot |> MRegistry.registerBxd rfid
        | _ -> Proxys.def.failed

    [<ExcelFunction(Category="Relation", Description="Returns a filtered relation.")>]
    let rel_restrict
        ([<ExcelArgument(Description= "Relation R-object.")>] rgRel: string)
        ([<ExcelArgument(Description= "Filter ('a -> bool) function R-object (max. arity 5).")>] rgFilterFn: string)
        ([<ExcelArgument(Description= "Argument 1")>] argument1: obj)
        ([<ExcelArgument(Description= "Argument 2")>] argument2: obj)
        ([<ExcelArgument(Description= "Argument 3")>] argument3: obj)
        ([<ExcelArgument(Description= "Argument 4")>] argument4: obj)
        ([<ExcelArgument(Description= "Argument 5")>] argument5: obj)
        : obj = 
        
        // intermediary stage
        let argNames = 
            [| argument1; argument2; argument3; argument4; argument5 |]
            |> Array.choose (In.D0.Stg.Opt.def None)

        if argNames.Length = 0 then
            Proxys.def.failed
        else
            // caller cell's reference ID
            let rfid = MRegistry.refID

            // result
            match tryExtract<Rel.Rel> rgRel, tryExtractO rgFilterFn with
            | Some rel, Some ofun -> 
                match Rel.restrict rel ofun argNames with
                | None -> Proxys.def.failed
                | Some relRestrict -> relRestrict |> MRegistry.registerBxd rfid
            | _ -> Proxys.def.failed

    [<ExcelFunction(Category="Relation", Description="Applys a basic filter to a relation.")>]
    let rel_filter
        ([<ExcelArgument(Description= "Relation R-object.")>] rgRel: string)
        ([<ExcelArgument(Description= "Filter field names.")>] filterNames: obj)
        ([<ExcelArgument(Description= "Filter field values.")>] filterValues: obj[])
        ([<ExcelArgument(Description= "[Exclude values. Default is false.]")>] excludeValues: obj) 
        : obj = 
        
        // intermediary stage
        let names = In.D1.OStg.tryDV None filterNames
        let exclude = In.D0.Bool.def false excludeValues

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel, names with
        | None, _ -> Proxys.def.failed
        | _, None -> Proxys.def.failed
        | Some rel, Some nms -> 
            match Rel.restrictBasic rel exclude nms filterValues with
            | None -> Proxys.def.failed
            | Some relRestrict -> relRestrict |> MRegistry.registerBxd rfid
        

    [<ExcelFunction(Category="Relation", Description="Returns an extended relation.")>]
    let rel_extend
        ([<ExcelArgument(Description= "Relation R-object.")>] rgRel: string)
        ([<ExcelArgument(Description= "Filter ('a -> bool) function R-object (max. arity 5).")>] rgFilterFn: string)
        ([<ExcelArgument(Description= "Output name")>] outputName: string)
        ([<ExcelArgument(Description= "Argument 1")>] argument1: obj)
        ([<ExcelArgument(Description= "Argument 2")>] argument2: obj)
        ([<ExcelArgument(Description= "Argument 3")>] argument3: obj)
        ([<ExcelArgument(Description= "Argument 4")>] argument4: obj)
        ([<ExcelArgument(Description= "Argument 5")>] argument5: obj)
        : obj = 
        
        // intermediary stage
        let argNames = 
            [| argument1; argument2; argument3; argument4; argument5 |]
            |> Array.choose (In.D0.Stg.Opt.def None)

        if argNames.Length = 0 then
            Proxys.def.failed
        else
            // caller cell's reference ID
            let rfid = MRegistry.refID

            // result
            match tryExtract<Rel.Rel> rgRel, tryExtractO rgFilterFn with
            | Some rel, Some ofun -> 
                match Rel.extend rel ofun argNames outputName with
                | None -> Proxys.def.failed
                | Some relExtd -> relExtd |> MRegistry.registerBxd rfid
            | _ -> Proxys.def.failed

    [<ExcelFunction(Category="Relation", Description="Groups some fields as a relation field.")>]
    let rel_group
        ([<ExcelArgument(Description= "Relation R-object.")>] rgRel: string)
        ([<ExcelArgument(Description= "Per names.")>] perNames: obj)
        ([<ExcelArgument(Description= "[Grouped attribute name. Default is \"GROUPED\".]")>] groupName: obj)
        : obj = 
        
        // intermediary stage
        let perNames = API.In.D1.OStg.tryDV None perNames
        let groupName = In.D0.Stg.def "GROUPED" groupName

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel, perNames with
        | Some rel, Some perNms -> 
            match Rel.group rel (perNms |> Set.ofArray) groupName with
            | None -> Proxys.def.failed
            | Some relGrpd -> relGrpd |> MRegistry.registerBxd rfid
        | _ -> Proxys.def.failed

    [<ExcelFunction(Category="Relation", Description="Ungroups a relation field.")>]
    let rel_ungroup
        ([<ExcelArgument(Description= "Relation R-object.")>] rgRel: string)
        ([<ExcelArgument(Description= "[Grouped attribute name. Default is \"GROUPED\".]")>] groupName: obj)
        : obj = 
        
        // intermediary stage
        let groupName = In.D0.Stg.def "GROUPED" groupName

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel with
        | Some rel -> 
            match Rel.ungroup rel groupName with
            | None -> Proxys.def.failed
            | Some relUngrpd -> relUngrpd |> MRegistry.registerBxd rfid
        | _ -> Proxys.def.failed


    [<ExcelFunction(Category="Relation", Description="Renames a relation.")>]
    let rel_rname
        ([<ExcelArgument(Description= "Relation R-object.")>] rgRel: string)
        ([<ExcelArgument(Description= "Existing field names.")>] existingNames: obj)
        ([<ExcelArgument(Description= "New field names.")>] newNames: obj)
        : obj = 
        
        // intermediary stage
        let existingNames = In.D1.OStg.tryDV None existingNames
        let newNames = In.D1.OStg.tryDV None newNames

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match tryExtract<Rel.Rel> rgRel, existingNames, newNames with
        | Some rel, Some existingNms, Some newNms -> 
            let mapOldNamesNewNames = Toolbox.Array.zip existingNms newNms |> Map.ofArray
            match Rel.rename rel mapOldNamesNewNames with
            | None -> Proxys.def.failed
            | Some relRName -> relRName |> MRegistry.registerBxd rfid
        | _ -> Proxys.def.failed

    [<ExcelFunction(Category="Relation", Description="Returns the relation field names.")>]
    let rel_head
        ([<ExcelArgument(Description= "Relation R-object.")>] rgRel: string)
        : obj[] = 

        // result
        match tryExtract<Rel.Rel> rgRel with
        | Some rel -> 
            if rel.card = 0 then [| Proxys.def.empty |] else rel.header
        | _ -> [| Proxys.def.failed |]

    [<ExcelFunction(Category="Relation", Description="Extracts a CSV file into a relation.")>]
    let rel_ofCSV
        ([<ExcelArgument(Description= "File path.")>] filePath: string)
        ([<ExcelArgument(Description= "[Field names row index. Default is 1.]")>] fieldNamesRowIndex: obj) 
        ([<ExcelArgument(Description= "[Field types row index. Default is 2.]")>] fieldTypesRowIndex: obj) 
        ([<ExcelArgument(Description= "[Field names override. Default is none.]")>] fieldNamesOvrd: obj) 
        ([<ExcelArgument(Description= "[Field type-tags override (date, double, doubleNaN, #string...). Default is none.]")>] fieldTypeTagsOvrd: obj)
        ([<ExcelArgument(Description= "[Field (name -> type) map R-object. Default is none.]")>] rgMapFNameFType: obj)
        ([<ExcelArgument(Description= "[Separator. Default is \",\".]")>] separator: obj) 
        ([<ExcelArgument(Description= "[Types with default value (#date, int, obj, string...). Default is None.]")>] typesWithDefault: obj)
        ([<ExcelArgument(Description= "[Types default values. Default is None.]")>] typeDefaultValues: obj)
        ([<ExcelArgument(Description= "[Enclosing quotes. Default is false.]")>] enclosingQuotes: obj) 
        ([<ExcelArgument(Description= "[Trim blanks. Default is false.]")>] trimBlanks: obj) 
        ([<ExcelArgument(Description= "[Date format. Default is None.]")>] dateFormat: obj) 
        ([<ExcelArgument(Description= "[Use VB Parser (debugging purpose). Default is true.]")>] useVBParser: obj) 
        : obj = 
        
        // intermediary stage
        let fNamesRowIdx = In.D0.Intg.def 1 fieldNamesRowIndex
        let fTypesRowIdx = In.D0.Intg.def 2 fieldTypesRowIndex
        let csvFields : Rel.CSVFields = 
            match API.In.D1.OStg.tryDV None fieldNamesOvrd, API.In.D1.OStg.tryDV None fieldTypeTagsOvrd with
            | None, None -> if fTypesRowIdx < fNamesRowIdx then Rel.NameSndTypeFst else Rel.NameFstTypeSnd
            | Some fieldNames, None -> Rel.TypeFst fieldNames
            | None, Some fieldTypes -> Rel.NameFst fieldTypes
            | Some fieldNames, Some fieldTypes -> Rel.NoHeader (fieldNames, fieldTypes)

        let mapNameTagType = MRegistry.tryExtract<Map<string,string>> rgMapFNameFType

        let csvFields3 : Rel.CSVFields2 = 
            match In.D0.Intg.Opt.def None fieldNamesRowIndex, In.D0.Intg.Opt.def None fieldTypesRowIndex, API.In.D1.OStg.tryDV None fieldNamesOvrd, API.In.D1.OStg.tryDV None fieldTypeTagsOvrd, mapNameTagType with            
            // when the only inputs are the row indices 
            | None, None, None, None, None -> Rel.NameFstTypeSndX (None, (None, None))
            | Some nmRowIdx, None, None, None, None -> if nmRowIdx > 1 then Rel.NameSndTypeFstX (None, (None, None)) else Rel.NameFstTypeSndX (None, (None, None))
            | None, Some tyRowIdx, None, None, None -> if tyRowIdx < 2 then Rel.NameSndTypeFstX (None, (None, None)) else Rel.NameFstTypeSndX (None, (None, None))
            | Some nmRowIdx, Some tyRowIdx, None, None, None -> if tyRowIdx < nmRowIdx then Rel.NameSndTypeFstX (None, (None, None)) else Rel.NameFstTypeSndX (None, (None, None))
            // when a names-override is provided
            | None, None, Some fieldNames, None, None -> Rel.TypeFstX (Some fieldNames, (None, None))
            | Some nmRowIdx, None, Some fieldNames, None, None -> if nmRowIdx > 1 then Rel.NameSndTypeFstX (Some fieldNames, (None, None)) else Rel.NameFstTypeSndX (Some fieldNames, (None, None))
            | None, Some tyRowIdx, Some fieldNames, None, None -> Rel.TypeFstX (Some fieldNames, (None, None)) // single header can only be on the first row
            | Some nmRowIdx, Some tyRowIdx, Some fieldNames, None, None -> if tyRowIdx < nmRowIdx then Rel.NameSndTypeFstX (Some fieldNames, (None, None)) else Rel.NameFstTypeSndX (Some fieldNames, (None, None))
            // when a types-override is provided
            | None, None, None, Some fieldTypes, None -> Rel.NameFstX (None, (Some fieldTypes, None))
            | Some nmRowIdx, None, None, Some fieldTypes, None -> Rel.NameFstX (None, (Some fieldTypes, None)) // single header can only be on the first row
            | None, Some tyRowIdx, None, Some fieldTypes, None -> if tyRowIdx < 2 then Rel.NameSndTypeFstX (None, (Some fieldTypes, None)) else Rel.NameFstTypeSndX (None, (Some fieldTypes, None))
            | Some nmRowIdx, Some tyRowIdx, None, Some fieldTypes, None -> if tyRowIdx < nmRowIdx then Rel.NameSndTypeFstX (None, (Some fieldTypes, None)) else Rel.NameFstTypeSndX (None, (Some fieldTypes, None))
            // when a map types-override is provided
            | None, None, None, _, Some mapping -> Rel.NameFstX (None, (None, Some mapping))
            | Some nmRowIdx, None, None, _, Some mapping -> Rel.NameFstX (None, (None, Some mapping)) // single header can only be on the first row
            | None, Some tyRowIdx, None, _, Some mapping -> if tyRowIdx < 2 then Rel.NameSndTypeFstX (None, (None, Some mapping)) else Rel.NameFstTypeSndX (None, (None, Some mapping))
            | Some nmRowIdx, Some tyRowIdx, None, _, Some mapping -> if tyRowIdx < nmRowIdx then Rel.NameSndTypeFstX (None, (None, Some mapping)) else Rel.NameFstTypeSndX (None, (None, Some mapping))
            // when both names- and types-override is provided
            | None, None, Some fieldNames, Some fieldTypes, None -> Rel.NoHeaderX (Some fieldNames, (Some fieldTypes, None))
            | Some nmRowIdx, None, Some fieldNames, Some fieldTypes, None -> Rel.NameFstX (Some fieldNames, (Some fieldTypes, None)) // single header can only be on the first row
            | None, Some tyRowIdx, Some fieldNames, Some fieldTypes, None -> Rel.TypeFstX (Some fieldNames, (Some fieldTypes, None)) // single header can only be on the first row
            | Some nmRowIdx, Some tyRowIdx, Some fieldNames, Some fieldTypes, None -> if tyRowIdx < nmRowIdx then Rel.NameSndTypeFstX (Some fieldNames, (Some fieldTypes, None)) else Rel.NameFstTypeSndX (Some fieldNames, (Some fieldTypes, None))
            // when both names- and map types-override is provided
            | None, None, Some fieldNames, _, Some mapping -> Rel.NoHeaderX (Some fieldNames, (None, Some mapping))
            | Some nmRowIdx, None, Some fieldNames, _, Some mapping -> Rel.NameFstX (Some fieldNames, (None, Some mapping)) // single header can only be on the first row
            | None, Some tyRowIdx, Some fieldNames, _, Some mapping -> Rel.TypeFstX (Some fieldNames, (None, Some mapping)) // single header can only be on the first row
            | Some nmRowIdx, Some tyRowIdx, Some fieldNames, _, Some mapping -> if tyRowIdx < nmRowIdx then Rel.NameSndTypeFstX (Some fieldNames, (None, Some mapping)) else Rel.NameFstTypeSndX (Some fieldNames, (None, Some mapping))

        let separator = In.D0.Stg.def "," separator
        let enclosingQuotes = In.D0.Bool.def false enclosingQuotes
        let trimBlanks = In.D0.Bool.def true trimBlanks
        let useVBParser = In.D0.Bool.def true useVBParser
        let dateFormat = In.D0.Stg.Opt.def None dateFormat
        let mapTypeDefvals = 
            match API.In.D1.OStg.tryDV None typesWithDefault with
            | None -> Map.empty
            | Some typeTags ->
                let defvals = In.Cast.to1D false typeDefaultValues
                let map' = Toolbox.Array.zip (typeTags |> Array.map (fun s -> s.ToUpper()))  defvals |> Map.ofArray
                let kvps = [| for kvp in map' -> (kvp.Key, Variant.rebox kvp.Key kvp.Value) |]
                kvps |> Map.ofArray

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        // match Rel.ofCSV mapNameTagType mapTypeDefvals Map.empty dateFormat useVBParser enclosingQuotes trimBlanks separator csvFields filePath with
        match Rel.ofCSVX mapTypeDefvals Map.empty dateFormat useVBParser enclosingQuotes trimBlanks separator csvFields3 filePath with
        | None -> Proxys.def.failed
        | Some relCSV -> relCSV |> MRegistry.registerBxd rfid

    [<ExcelFunction(Category="Relation", Description="Creates a relation from a map R-object.")>]
    let rel_ofMap
        ([<ExcelArgument(Description= "Map R-object.")>] rgMap: string)
        ([<ExcelArgument(Description= "[Key1 field name. Default is KEY1.]")>] fieldNameKey1: obj)  
        ([<ExcelArgument(Description= "[Key2 field name. Default is KEY2.]")>] fieldNameKey2: obj)  
        ([<ExcelArgument(Description= "[Key3 field name. Default is KEY3.]")>] fieldNameKey3: obj)  
        ([<ExcelArgument(Description= "[Key4 field name. Default is KEY4.]")>] fieldNameKey4: obj)  
        ([<ExcelArgument(Description= "[Key5 field name. Default is KEY5.]")>] fieldNameKey5: obj)  
        ([<ExcelArgument(Description= "[Key6 field name. Default is KEY6.]")>] fieldNameKey6: obj)  
        ([<ExcelArgument(Description= "[Key7 field name. Default is KEY7.]")>] fieldNameKey7: obj)  
        ([<ExcelArgument(Description= "[Value field name. Default is VALUE.]")>] fieldNameValue: obj)  
        ([<ExcelArgument(Description= "[Split tupled key. Default is true.]")>] splitTupleKey: obj)  
        : obj = 
        
        // intermediary stage
        let fieldNameKeys = 
            [| fieldNameKey1; fieldNameKey2; fieldNameKey3; fieldNameKey4; fieldNameKey5; fieldNameKey6; fieldNameKey7 |]
            |> Array.mapi (fun i fnamekey -> match In.D0.Stg.Opt.def None fnamekey with | None -> sprintf "KEY%d" i | Some key -> key)

        let fieldNameValue = In.D0.Stg.def "VALUE" fieldNameValue
        let splitTupleKey = In.D0.Bool.def true splitTupleKey

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match Rel.Registry.Out.ofMap rgMap splitTupleKey fieldNameKeys fieldNameValue with
        | None -> Proxys.def.failed
        | Some relofMap -> relofMap |> MRegistry.registerBxd rfid

    [<ExcelFunction(Category="Map", Description="Creates a map R-object from a relation R-object.")>]
    let map_ofRel
        ([<ExcelArgument(Description= "Relation R-object.")>] rgRel: string)
        ([<ExcelArgument(Description= "Key1 field name.")>] key1Name: string)  
        ([<ExcelArgument(Description= "[Key2 field name. Default is none.]")>] key2Name: obj)  
        ([<ExcelArgument(Description= "[Key2 field name. Default is none.]")>] key3Name: obj)  
        ([<ExcelArgument(Description= "[Key2 field name. Default is none.]")>] key4Name: obj)  
        ([<ExcelArgument(Description= "[Key2 field name. Default is none.]")>] key5Name: obj)  
        ([<ExcelArgument(Description= "[Key2 field name. Default is none.]")>] key6Name: obj)  
        ([<ExcelArgument(Description= "[Key2 field name. Default is none.]")>] key7Name: obj)  
        ([<ExcelArgument(Description= "[Key2 field name. Default is none.]")>] key8Name: obj)  
        ([<ExcelArgument(Description= "Value field name.")>] valueName: string)
        ([<ExcelArgument(Description= "[Group results (map type is (keys -> values[]) for true, (keys -> values) for false). Default is false.]")>] groupResults: obj)
        : obj = 

        // intermediary arguments / calculations
        let key2 = In.D0.Stg.Opt.def None key2Name
        let key3 = In.D0.Stg.Opt.def None key3Name
        let key4 = In.D0.Stg.Opt.def None key4Name
        let key5 = In.D0.Stg.Opt.def None key5Name
        let key6 = In.D0.Stg.Opt.def None key6Name
        let key7 = In.D0.Stg.Opt.def None key7Name
        let key8 = In.D0.Stg.Opt.def None key8Name
        let groupResults = In.D0.Bool.def false groupResults
        let keysNames = [| Some key1Name; key2; key3; key4; key5; key6; key7; key8 |] |> Array.choose id

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        //match tryExtract (tryRel None) rgRel with
        match tryExtract<Rel.Rel> rgRel with
        | None -> Out.Proxys.def.failed
        | Some rel ->
            match Rel.Registry.In.toMap rel groupResults keysNames valueName with
            | None -> Out.Proxys.def.failed
            | Some mapofRel -> mapofRel |> MRegistry.registerBxd rfid


/// Simple template for generics
module GenMtrx =
    open type Registry
    open Registry
    open Toolbox.Generics
    open API

    // some generic type
    type GenMTRX<'a>(elem: 'a, size: int) =
        let a2D : 'a[,] = Array2D.create size size elem

        member this.size = a2D |> Array2D.length1
        member this.elem (row: int) (col: int) : 'a = a2D.[row, col]

    let genType = typeof<GenMTRX<_>>

    // some typed creation functions
    let create0D<'a> (size: int) (a0D: 'a) : GenMTRX<'a> = GenMTRX(a0D,  size)
    let create1D<'a> (size: int) (a1D: 'a[]) : GenMTRX<'a> = failwith "NIY"  // [| for i in 0 .. (size - 1) -> a1D |] |> array2D

    // == reflection functions ==
    type GenFn =
        static member mtrx0D<'A> (defValue: obj option) (typeLabel: string) (size: int) (xlValue: obj) : GenMTRX<'A> =
            // let a0Dx = In.D0.Gen.def defValue typeLabel xlValue
            let a0D = In.D0.Tag.Any.def Set.empty defValue typeLabel xlValue :?> 'A
            a0D |> create0D size

        static member mtrx1D<'A> (defValue: obj option) (typeLabel: string) (size: int) (xlValue: obj) : GenMTRX<'A> =
            let a1D = In.D1.TagFn.def Set.empty None defValue typeLabel xlValue
            a1D |> create1D size
        
        static member size<'A> (mtrx: GenMTRX<'A>) : int = mtrx.size
        static member elem<'A> (mtrx: GenMTRX<'A>) (row: int) (col: int) : 'A = mtrx.elem row col

    // == registry functions ==
    module Reg =
        module In =
            let mtrx0D (defValue: obj option) (typeLabel: string) (size: int) (xlValue: obj) : obj = 
                let gentype = typeLabel |> Variant.labelType false
                let args : obj[] = [| defValue; typeLabel; size; xlValue |]
                let res = invoke<GenFn> "mtrx0D" [| gentype |] args
                res

        module Out =
            let size (regKey: string) : obj option =
                let methodNm = "size"
                MRegistry.tryExtractGen genType regKey
                |> Option.map (apply<GenFn> methodNm [||] [||])

            let elem (row: int) (col: int) (regKey: string) : obj option =
                let methodNm = "elem"
                MRegistry.tryExtractGen genType regKey
                |> Option.map (apply<GenFn> methodNm [||] [| row; col |])

/// Simple template
module Mtrx =
    open type Registry
    open Registry

    open Toolbox.Generics
    open API

    type MTRXD = double[,]
    type MTRX<'a> = 'a[,]

    module Test =
        type myType = 
            static member myMember1<'T> (arg1: obj) : 'T[] = [||]
            static member myMember2<'T> (arg2: obj) : 'T[,] = [|[||]|] |> array2D

    // some typed creation functions
    let create0D<'a> (size: int) (a0D: 'a) : MTRX<'a> = Array2D.create size size a0D
    let create1D<'a> (size: int) (a1D: 'a[]) = [| for i in 0 .. (size - 1) -> a1D |] |> array2D

    // == reflection functions ==
    type GenFn =
        static member mtrx0D<'A> (defValue: obj option) (typeLabel: string) (size: int) (xlValue: obj) : MTRX<'A> =
            let a0D = In.D0.TagFn.def Set.empty defValue typeLabel xlValue
            a0D |> create0D size

        static member mtrx1D<'A> (defValue: obj option) (typeLabel: string) (size: int) (xlValue: obj) : MTRX<'A> =
            let a1D = In.D1.TagFn.def Set.empty None defValue typeLabel xlValue
            a1D |> create1D size
        
        static member size<'A> (mtrx: MTRX<'A>) : int = mtrx |> Array2D.length1

    // == registry functions ==
    let mtrxRegOLD (xlValue: string) : obj = 
        match MRegistry.tryType xlValue with
        | None -> box "failed"
        | Some ty -> 
            let isgenty = ty.IsGenericType
            let isgenty = ty.ContainsGenericParameters
            
            let genty = ty.GetGenericTypeDefinition()
            let testTY = typeof<MTRX<_>>.GetGenericTypeDefinition() = genty
            let getGenTy = ty.GetGenericArguments()
            // https://docs.microsoft.com/en-us/dotnet/framework/reflection-and-codedom/how-to-examine-and-instantiate-generic-types-with-reflection
            // https://docs.microsoft.com/en-us/dotnet/api/system.type.getgenerictypedefinition?view=net-5.0
            if testTY then
                match MRegistry.tryExtract xlValue with
                | None -> box "should NOT be here"
                | Some omtrx ->
                    let args : obj[] = [| omtrx |]
                    let res = invoke<GenFn> "size" getGenTy args
                    res
            else
                box "failed"
            //let gentype = typeLabel |> Variant.labelType true
            //let args : obj[] = [| defValue; typeLabel; xlValue |]
            //let res = apply<Gen> "mtrx0D" [| gentype |] args
            //res


    // == xl-values functions ==
    /// 'a is determined by typeLabel.
    let mtrx0D (defValue: obj) (typeLabel: string) (xlValue: obj) : obj = 
        let gentype = typeLabel |> Variant.labelType true
        let args : obj[] = [| defValue; typeLabel; xlValue |]
        let res = invoke<GenFn> "mtrx0D" [| gentype |] args
        res

    /// FIXME
    /// 'a is determined by typeLabel.
    let mtrx1D (defValue: obj) (typeLabel: string) (xlValue: obj) : obj = 
        let gentype = typeLabel |> Variant.labelType true
        let args : obj[] = [| defValue; typeLabel; xlValue |]
        let res = invoke<GenFn> "mtrx1D" [| gentype |] args
        res

module TEST_XL =
    // open type Registry
    open Registry
    // open Excel
    open API
    open API.Out
    open type Variant
    open type API.Out.Proxys
    open GenMtrx

    [<ExcelFunction(Category="EXAMPLE", Description="Creates a Mtrx R-object.")>]
    let mtrxGen_create
        ([<ExcelArgument(Description= "Value.")>] value: obj)
        ([<ExcelArgument(Description= "Size.")>] size: double)
        ([<ExcelArgument(Description= "Type label.")>] typeLabel: string)
        : obj  =

        // intermediary stage
        let mtrxg = Reg.In.mtrx0D None typeLabel ((int) size) value

        // caller cell's reference ID
        let rfid = MRegistry.refID
        
        // result
        let res = mtrxg |> MRegistry.register rfid
        box res

    [<ExcelFunction(Category="EXAMPLE", Description="Returns Mtrx element.")>]
    let mtrxGen_elem
        ([<ExcelArgument(Description= "Matrix R-obj.")>] rgMtrx: string) 
        ([<ExcelArgument(Description= "[Row indice. Default is 0.]")>] row: obj)
        ([<ExcelArgument(Description= "[Col indice. Default is 0.]")>] col: obj)
        : obj = 

        // intermediary stage
        let row = In.D0.Intg.def 0 row
        let col = In.D0.Intg.def 0 col

        // result
        let xxx  = Reg.Out.elem row col rgMtrx
        // xxx |> (Out.cellOptTBD Proxys.def) // TODO FIXME
        xxx |> Out.D0.Bxd.Opt.out Proxys.def

    [<ExcelFunction(Category="EXAMPLE", Description="Returns Mtrx size.")>]
    let mtrxGen_size
        ([<ExcelArgument(Description= "Matrix R-obj.")>] rgMtrx: string) 
        : obj = 

        // result
        Reg.Out.size rgMtrx
        |> Out.outStg "failed"

    [<ExcelFunction(Category="EXAMPLE", Description="Cast a 1D-slice of an xl-range to a generic type array.")>]
    let mtrx_create
        ([<ExcelArgument(Description= "size.")>] size: double)
        ([<ExcelArgument(Description= "value.")>] value: obj)
        ([<ExcelArgument(Description= "Type label.")>] typeLabel: string)
        //([<ExcelArgument(Description= "[Default Value (only for non-optional types). Must be of the appropriate type. Default \"<default>\" (will fail for non-string types).]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[Default Value (only for non-optional types, optional types default to None). Must be of the appropriate type, else it will fail.]")>] defaultValue: obj)
        : obj  =

        // intermediary stage
        let a2D = Array2D.create ((int) size) ((int) size) value
        let isoptional = isOptionalType typeLabel

        // caller cell's reference ID
        let rfid = MRegistry.refID
        //Mtrx.mtrx0D defvalue typeLabel
        let res = a2D |> MRegistry.register rfid
        box res

    [<ExcelFunction(Category="EXAMPLE", Description="Creates an array of a reg. object.")>]
    let mtrx_elem
        ([<ExcelArgument(Description= "Matrix reg. object.")>] rgA1D: string) 
        ([<ExcelArgument(Description= "[Row indice. Default is 0.]")>] row: obj)
        ([<ExcelArgument(Description= "[Col indice. Default is 0.]")>] col: obj)
        : obj = 

        // intermediary stage
        let row = In.D0.Intg.def 0 row
        let col = In.D0.Intg.def 0 col

        // result
        match MRegistry.tryExtract<Mtrx.MTRXD> rgA1D with
        | None -> box "FAILED"
        | Some a2d -> box a2d.[row, col]

    [<ExcelFunction(Category="EXAMPLE", Description="Get Mtrx size.")>]
    let mtrx_size
        ([<ExcelArgument(Description= "Matrix reg. object.")>] rgMtrx: string) 
        : obj = 

        // result
        let xx = Mtrx.mtrxRegOLD rgMtrx
        xx





// Set.indexed + Set.item(i) or at least Set.head
// Set.groupBy
// Set.allPair
// array2D trans
// Set.disjoint
// Set.intersectMany => return empty set if input is empty? (currently throws error)

// io functions
// fun isfunction. fun arity and type expr_funtype...