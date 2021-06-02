namespace ExcelGlue

open System
open System.IO
open System.Collections.Generic
open ExcelDna.Integration
open System.Runtime.Serialization.Formatters.Binary

/// Class where all "registered" xl-sheet objects are stored.
type Registry() =
    // 2 dictionaries to keep track of objects, and xl-cells references where they are located.
    let reg = Dictionary<string, obj>()
    let ref = Dictionary<string, string[]>()

    // -----------------------------------
    // -- Construction functions
    // -----------------------------------

    /// Removes all objects which are filed under the reference refKey.
    member private this.removeReferenceObjects (refKey: string) = 
        if ref.ContainsKey refKey then
            for regKey in ref.Item(refKey) do reg.Remove(regKey) |> ignore
            ref.Remove(refKey) |> ignore

    /// Removes all objects and their xl-cell references from the Object Registry.
    member this.clear = 
        reg.Clear()
        ref.Clear()

    /// Adds a reference -> single registry key 
    member private this.addReference (refKey: string) (regKey: string) = 
        this.removeReferenceObjects refKey
        ref.Add(refKey, [| regKey |])

    /// Adds a a single registry key to a (possibly already existing) reference
    member private this.appendRef (refKey: string) (regKey: string) =
        if ref.ContainsKey refKey then
            let regKeys = ref.Item(refKey)
            ref.Remove(refKey) |> ignore
            ref.Add(refKey, Array.append regKeys [| regKey |])
        else
            ref.Add(refKey, [| regKey |])

    /// Adds regObject to the registry, and delete all previous references to refKey.
    member this.register (refKey: string) (regObject: obj) : string = 
        let regKey = (Guid.NewGuid()).ToString()
        this.addReference refKey regKey
        reg.Add(regKey, regObject)
        regKey

    /// Adds regObject to the registry, under an existing reference refKey.
    member this.append (refKey: string) (regObject: obj) : string = 
        let regKey = (Guid.NewGuid()).ToString()
        this.appendRef refKey regKey
        reg.Add(regKey, regObject)
        regKey

    // -----------------------------------
    // -- Inspection functions
    // -----------------------------------

    // Returns the number of registry objects held.
    member this.count = reg.Count

    /// Returns a registry object, given its regKey.
    member this.tryFind (regKey: string) : obj option =
        if reg.ContainsKey regKey then
            reg.Item(regKey) |> Some
        else
            None

    /// Returns a reference, given a registry key.
    /// There can be only one reference per object.
    member private this.tryFindRef (regKey: string) : string option = 
        if reg.ContainsKey regKey then
            [| for kvp in ref -> if kvp.Value |> Array.contains regKey then [| kvp.Key |] else [||] |]
            |> Array.concat
            |> Array.head
            |> Some
        else
            None

    member this.tryFind1D (regKey: string) : ((Type[])*obj) option =
        match this.tryFind regKey with
        | None -> None
        | Some regObj ->
            let ty = regObj.GetType()

            if ty.IsArray && (ty.GetArrayRank() = 1) then
                let genty = ty.GetElementType()
                ([| genty |], regObj) 
                |> Some
            else
                None

    member this.tryFind2D (regKey: string) : ((Type[])*obj) option =
        match this.tryFind regKey with
        | None -> None
        | Some regObj ->
            let ty = regObj.GetType()

            if ty.IsArray && (ty.GetArrayRank() = 2) then
                let genty = ty.GetElementType()
                ([| genty |], regObj) 
                |> Some
            else
                None

    /// Returns a registry object's type, given its regKey.
    member this.tryType (regKey: string) : Type option =
        regKey |> this.tryFind |> Option.map (fun o -> o.GetType())

    /// Checks equality on 2 registry objects, given their keys.
    member this.equal (regKey1: string) (regKey2: string) : bool = 
        match this.tryFind regKey1, this.tryFind regKey2 with
        | Some o1, Some o2 -> o1 = o2 // should have an equality constraint?
        | _ -> false

    /// Returns the registry keys.
    member this.keys : string[] = [| for kvp in reg -> kvp.Key |]

    /// Returns the registry values.
    member this.values : obj[] = [| for kvp in reg -> kvp.Value |]

    /// Returns the registry key-value pairs.
    member this.keyValuePairs : (string*obj)[] = [| for kvp in reg -> kvp.Key, kvp.Value |]

    // -----------------------------------
    // -- Extraction functions
    // -----------------------------------

    member this.tryExtract<'a> (xlValue: obj) : 'a option =
        match xlValue with
        | :? string as regKey ->
            match this.tryFind regKey with
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
        | :? string as regKey -> this.tryFind regKey
        | _ -> None

    // https://docs.microsoft.com/en-us/dotnet/framework/reflection-and-codedom/how-to-examine-and-instantiate-generic-types-with-reflection
    // https://docs.microsoft.com/en-us/dotnet/api/system.type.getgenerictypedefinition?view=net-5.0
    /// targetGenType is the expected generic type, e.g. targetGenType: typeof<GenMTRX<_>>
    member this.tryExtractGen' (targetGenType: Type) (xlValue: string) : obj option =  // TODO change xlValue into regKey
        if not targetGenType.IsGenericType then
            None
        else
            match this.tryFind xlValue with
            | None -> None
            | Some regObj -> 
                let ty = regObj.GetType()
                let gentydef = ty.GetGenericTypeDefinition()
                let tgttydef = targetGenType.GetGenericTypeDefinition()

                if gentydef = tgttydef then
                    Some regObj
                else
                    None

    /// Same as tryExtractGen, but also return the generic types array.
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
        | :? string as regKey -> this.tryType regKey
        | _ -> None


    // -----------------------------------
    // -- Excel RefID
    // -----------------------------------

    member this.excelRef (caller : obj) : string = 
        match caller with
        | :? ExcelReference as ref -> ref.ToString()
        | _ -> ""

    member this.refID = XlCall.Excel XlCall.xlfCaller |> this.excelRef

    // -----------------------------------
    // -- Pretty-print functions
    // -----------------------------------

    /// Pretty-prints a registry object, given its key.
    member this.tryShow (key: string) : string option =
        this.tryFind key |> Option.map (fun o -> sprintf "%A" o)

    member this.tryString (key: string) : string option =
        this.tryFind key |> Option.map (fun o -> o.ToString())

    /// Pretty-prints a registry object, given its key.
    member this.tryPPrint (key: string) : string option =
        this.tryFind key |> Option.map (fun o -> sprintf "%A" o)

    // -----------------------------------
    // -- Misc.
    // -----------------------------------
    
    /// Unsafe
    member this.defaultValue<'a> (xlValue: obj) : 'a =
        this.tryExtract<'a> xlValue |> Option.defaultValue (Unchecked.defaultof<'a>)

    member this.ioWriteBin (fpath: string) (regKey: string) : DateTime option =
        match regKey |> this.tryFind with
        | None -> None
        | Some o ->
            use stream = new FileStream(fpath, FileMode.Create)
            (new BinaryFormatter()).Serialize(stream, o)
            DateTime.Now |> Some

    member this.ioLoadBin (fpath: string) (refKey: string) : string =
        use stream = new FileStream(fpath, FileMode.Open)
        let regObj = (new BinaryFormatter()).Deserialize(stream)
        this.register refKey regObj

module Registry =
    /// Master registry where all registered objects are held.
    let MRegistry = Registry()

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

/// Excel substitute output values.
///    - Proxys.empty for empty arrays.
///    - Proxys.failed for function failure.
///    - Proxys.nan for Double.NaN values.
///    - Proxys.none for optional F# None values.
///    - Proxys.object for non-primitive types values.
type Proxys = { empty: obj; failed: obj; nan: obj; none: obj; object: obj } with
    static member def : Proxys = { empty = "<empty>"; failed = box ExcelError.ExcelErrorNA ; nan = ExcelError.ExcelErrorNA; none = "<none>"; object = "<obj>" }

module Excel =
    open System

    // FIXME better wording
    /// Indicates which xl-values are to be converted to special values (e.g. Double.NaN in 0D, [||] in 1D) :
    ///    - if OnlyErrorNA, only ExcelErrorNA values are converted.
    ///    - if AllErrors, all Excel error values are converted.
    ///    - if AllNonNumeric, any non-numeric xl-value are.

    /// Describes various convenient sets, "kinds", of xl-spreadsheet values.
    type Kind = | Boolean | Numeric | Textual | NA | Error | Missing | Empty with
        static member nonBoolean = [| Numeric; Textual |] |> Array.sort

        static member nonNumeric = [| Boolean; Textual |] |> Array.sort
        static member nonNumericAndNA = [| Boolean; Textual; NA |] |> Array.sort
        static member nonNumericAndError = [| Boolean; Textual; NA; Error |] |> Array.sort

        static member nonTextual = [| Boolean; Numeric |] |> Array.sort

        static member onlyNA = [| NA |]
        static member anyError = [| NA; Error |] |> Array.sort
        static member none = [||]

        static member ofLabel (label: string) : Kind[] =
            match label.ToUpper() with
            | "NA" -> Kind.onlyNA
            | "ERR" | "ERROR" -> Kind.anyError
            | "NN" | "NONNUM" | "NONNUMERIC" -> Kind.nonNumeric
            | "NNNA" | "NN_NA" | "NN+NA" | "NONNUM_NA" | "NONNUM+NA" | "NONNUMERIC_NA" | "NONNUMERIC+NA" ->  Kind.nonNumericAndNA
            | "NNERR" | "NN_ERR" | "NN+ERR" | "NONNUM_ERR" | "NONNUM+ERR" | "NONNUMERIC_ERROR" | "NONNUMERIC+ERROR" -> Kind.nonNumericAndError
            | _ -> Kind.none

        static member labelGuide : (string*string) [] =  // FIXME do better
            let labels = [| "NA"; "ERR"; "NN"; "NNNA"; "NNERR"; "NONE"; "default" |]
            labels |> Array.map (fun lbl -> (lbl, Kind.ofLabel lbl |> Array.map (fun kinds -> kinds.ToString()) |> (String.concat "|")))

    module Kind =
        /// Matches non-numeric, non-error variants, i.e. Bools and Strings.
        let (|NonNumeric|_|) (xlKinds: Kind[]) (xlVal: obj) : bool option = 
            match xlVal, xlKinds |> Array.sort with            
            | :? string, k when k = Kind.nonNumeric -> Some true
            | :? bool, k when k = Kind.nonNumeric -> Some true
            | _ -> None

        /// Matches non-numeric and #N/A variants, i.e. Bools, Strings and #N/A.
        let (|NonNumericAndNA|_|) (xlKinds: Kind[]) (xlVal: obj) : bool option = 
            match xlVal, xlKinds |> Array.sort with            
            | :? string, k when k = Kind.nonNumericAndNA -> Some true
            | :? bool, k when k = Kind.nonNumericAndNA -> Some true
            | :? ExcelError as xlerr, k  when (xlerr = ExcelError.ExcelErrorNA) && (k = Kind.nonNumericAndNA) -> Some true
            | _ -> None

        /// Matches non-numeric and any error variants, i.e. Bools, Strings and errors.
        let (|NonNumericAndError|_|) (xlKinds: Kind[]) (xlVal: obj) : bool option = 
            match xlVal, xlKinds |> Array.sort with            
            | :? string, k when k = Kind.nonNumericAndError -> Some true
            | :? bool, k when k = Kind.nonNumericAndError -> Some true
            | :? ExcelError, k when k = Kind.nonNumericAndError -> Some true
            | _ -> None

        /// Only matches #N/A variants.
        let (|OnlyNA|_|) (xlKinds: Kind[]) (xlVal: obj) : bool option = 
            match xlVal, xlKinds |> Array.sort with            
            | :? ExcelError as xlerr, k  when (xlerr = ExcelError.ExcelErrorNA) && (k = Kind.onlyNA) -> Some true
            | _ -> None

        /// Matches any error variants, e.g. #N/A, #NUM, #REF...
        let (|AnyError|_|) (xlKinds: Kind[]) (xlVal: obj) : bool option = 
            match xlVal, xlKinds |> Array.sort with            
            | :? ExcelError, k when k = Kind.anyError -> Some true
            | _ -> None

module Useful =
    module Option =   
        /// NONE should always precede SOME in active patterns.
        let (|SOME|_|) : obj -> obj option =
            let tyOpt = typedefof<option<_>>
            fun (a:obj) ->
                let aty = a.GetType()
                let v = aty.GetProperty("Value")
                if aty.IsGenericType && aty.GetGenericTypeDefinition() = tyOpt then
                    if a = null then None else Some(v.GetValue(a, [| |]))
                else None

        /// NONE should always precede SOME in active patterns.
        let (|NONE|_|) : obj -> obj option =
          fun (a:obj) -> if a = null then Some (box "None detected here") else None

        let unbox (o: obj) : obj option =   
            match o with    
            | NONE(_) -> None  // NONE first
            | SOME(x) -> Some x // SOME second
            | _ -> failwith "Not an optional type."
        
        /// Returns none for None values, or map non-optional and Some values.
        /// Useful for mapping optional and non-optional values alike back to Excel.
        let map (none: obj) (mapping: obj -> obj) (o: obj) : obj =   
            match o with    
            | NONE(_) -> none
            | SOME(x) -> mapping x
            | _ -> mapping o
        
        /// Returns the underlying value's type for an optional type.
        /// I.e. typeof<int option> returns typeof<int> |> Some.
        /// None otherwise.
        let uType (aType: Type) : Type option =
            let tyOpt = typedefof<option<_>>
            if aType.IsGenericType && aType.GetGenericTypeDefinition() = tyOpt then
                aType.GenericTypeArguments |> Array.head |> Some
            else
                None

    module Generics =    
        let invoke<'gen> (methodName: string) (methodTypes: Type[]) (methodArguments: obj[]) : obj =
            let meth = typeof<'gen>.GetMethod(methodName)
            let genm = meth.MakeGenericMethod(methodTypes)
            let res  = genm.Invoke(null, methodArguments)
            res

        let apply<'gen> (methodName: string) (otherArgumentsLeft: obj[]) (otherArgumentsRight: obj[]) (genTypeRObj: Type[]*obj) : obj =
            let (gentys, robj) = genTypeRObj
            invoke<'gen> methodName gentys ([| otherArgumentsLeft; [| robj |];  otherArgumentsRight |] |> Array.concat)

        let apply2<'gen,'a> (methodName: string) (otherArgumentsLeft: obj[]) (otherArgumentsRight: obj[]) (genTypeRObj: Type[]*obj) : obj =
            let (gentys, robj) = genTypeRObj
            invoke<'gen> methodName gentys ([| otherArgumentsLeft; [| (robj :?> 'a[]) |> box |];  otherArgumentsRight |] |> Array.concat)

        /// 1 (common) generic-type for 2 generic-arguments.
        /// E.g. myFun<'a> (arg1: 'a) (arg2: 'a) (someExtraArg: ...) = ...
        let applyMulti<'gen> (methodName: string) (otherArgumentsLeft: obj[]) (otherArgumentsRight: obj[]) (genTypes: Type[]) (rObjs: obj[]) : obj =
            invoke<'gen> methodName genTypes ([| otherArgumentsLeft; rObjs;  otherArgumentsRight |] |> Array.concat)

    [<RequireQualifiedAccess>]
    module Type =
        /// Determines whether an object is of an (extended) primitive type.
        let isPrimitive' (includeObject: bool) (o: obj) : bool =
            (includeObject && (o.GetType().Name = "Object")) ||
            match o with
            | :? double -> true
            | :? string -> true
            | :? DateTime -> true
            | :? int -> true
            | :? bool -> true
            | :? Decimal -> true
            | :? Byte -> true
            // ... to be continued.
            | _ -> false

        /// Same as isPrimitive' but via a Type argument.
        let isPrimitive (includeObject: bool) (aType: Type) =
            (includeObject && (aType.Name = "Object")) 
            || aType.IsPrimitive
            || (aType.Name = "String") || (aType.Name = "DateTime") || (aType.Name = "Decimal")

        let pPrint (toStringStyle: bool) (someType: Type) : string =             
            let s = if toStringStyle then someType.ToString() else sprintf "%A" someType
            let pp = 
                s
                 .Replace(someType.Namespace + ".","").Replace("System.", "")
                 .Replace("FSharpOption`1","Option")
                 .Replace("FSharpMap`2","Map")
            pp

module API = 
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
            open type Variant
            open Excel
            open Excel.Kind

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
                /// Converts xl-values to boxed Double.NaN in special cases, depending on xl-value's Kind.
                let nanify (xlKinds: Kind[]) (xlVal: obj) : obj = 
                    match xlVal with
                    | :? double -> xlVal
                    | NonNumeric xlKinds _ -> box Double.NaN
                    | NonNumericAndNA xlKinds _ -> box Double.NaN
                    | NonNumericAndError xlKinds _ -> box Double.NaN
                    | OnlyNA xlKinds _ -> box Double.NaN
                    | AnyError xlKinds _ -> box Double.NaN
                    | _ -> xlVal

                /// Converts a boxed ExcelErrorNA into a Double.NaN.
                /// Similar to nanify function, but simpler.
                let ofNA (xlVal: obj) : obj =
                    match xlVal with
                    | :? ExcelError as err when err = ExcelError.ExcelErrorNA -> Double.NaN |> box
                    | _ -> xlVal

                /// Converts a boxed Double.NaN into an ExcelErrorNA. // FIXME - should be OUT?
                let ofNaN (xlVal: obj) : obj =
                    match xlVal with
                    | :? double as d -> if Double.IsNaN d then ExcelError.ExcelErrorNA |> box else box d
                    | _ -> xlVal

                /// Casts an xl-value to double or fails, with some other non-double values potentially cast to Double.NaN.
                let fail (xlKinds: Kind[]) (msg: string option) (xlVal: obj) = 
                    nanify xlKinds xlVal |> fail<double> msg

                /// Casts an xl-value to double with a default-value, with some other non-double values potentially cast to Double.NaN. // FIXME - improve text
                let def (xlKinds: Kind[]) (defValue: double) (xlVal: obj) = 
                    nanify xlKinds xlVal |> def<double> defValue

                // optional-type default FIXX
                module Opt =
                    /// Casts an xl-value to a double option type with a default-value, with some other non-double values potentially cast to Double.NaN.
                    let def (xlKinds: Kind[]) (defValue: double option) (xlVal: obj) = 
                        nanify xlKinds xlVal |> Opt.def<double> defValue

                /// Object substitution, based on type.
                module Obj =
                    /// Substitutes an object for another one, if it isn't a (boxed) double (e.g. box 1.0).
                    /// Replaces an xl-value with a double default-value if it isn't a (boxed double) type (e.g. box 1.0), with some other non-double values potentially cast to Double.NaN.
                    let subst (xlKinds: Kind[]) (defValue: obj) (xlVal: obj) = 
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
                static member def<'A> (defValue: obj option) (typeTag: string) (xlValue: obj) : 'A = 

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
                        let a0D = Nan.def Kind.nonNumericAndNA defval xlValue // TODO: pass xlkinds as argument
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
                static member defOpt<'A> (defValue: obj option) (typeTag: string) (xlValue: obj) : 'A option = 
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
                        let a0D = Nan.Opt.def Kind.nonNumericAndNA defval xlValue // TODO: pass xlkinds as argument
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
                /// Casts an xl-value to a 'A, with a default-value for when the casting fails.
                /// 'a is determined by typeTag.
                let def (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype = typeTag |> Variant.labelType true
                    let args : obj[] = [| defValue; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<TagFn> "def" [| gentype |] args
                    res

                // optional-type default FIXX
                module Opt =
                    /// Casts an xl-value to a 'a option, with a default-value for when the casting fails.
                    /// 'a is determined by typeTag.
                    let def (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                        let gentype = typeTag |> Variant.labelType true
                        let args : obj[] = [| defValue; typeTag; xlValue |]
                        let res = Useful.Generics.invoke<TagFn> "defOpt" [| gentype |] args
                        res

                /// For when the type-tag is either optional, e.g. "#string", or not, e.g. "string". TODO wording
                module Any =
                    /// Convenient, single function covering def and Opt.def cases.
                    /// The returned (boxed) value might be either a 'a or a ('a option), depending on wether the type-tag is optional or not.
                    let def (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                        let gentype = typeTag |> Variant.labelType true
                        let args : obj[] = [| defValue; typeTag; xlValue |]

                        let res =
                            if typeTag |> isOptionalType then
                                Useful.Generics.invoke<TagFn> "defOpt" [| gentype |] args
                            else
                                Useful.Generics.invoke<TagFn> "def" [| gentype |] args
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


        /// Obj[] input functions.
        module D1 =
            open Excel
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
                /// Converts an obj[] to a bool[], given a default-value for non-bool elements.
                let def (defValue: bool) (o1D: obj[]) = def defValue o1D

                /// optional-type default
                module Opt =
                    /// Converts an obj[] to a ('a option)[], given a default-value for non-bool elements.
                    let def (defValue: bool option) (o1D: obj[]) = Opt.def defValue o1D

                /// Converts an obj[] to a bool[], removing any non-bool element.
                let filter (o1D: obj[]) = filter<bool> o1D

                /// Converts an obj[] to an optional 'a[]. All the elements must be bool, otherwise defValue array is returned. 
                let tryDV (defValue1D: bool[] option) (o1D: obj[])  = tryDV<bool> defValue1D o1D

            [<RequireQualifiedAccess>]
            module Stg =
                /// Converts an obj[] to a string[], given a default-value for non-string elements.
                let def (defValue: string) (o1D: obj[]) = def<string> defValue o1D  // TODO add <string> everywhere!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

                /// optional-type default
                module Opt = 
                    /// Converts an obj[] to a ('a option)[], given a default-value for non-string elements.
                    let def (defValue: string option) (o1D: obj[]) = Opt.def defValue o1D

                /// Converts an obj[] to a string[], removing any non-string element.
                let filter (o1D: obj[]) = filter<string> o1D

                /// Converts an obj[] to an optional 'a[]. All the elements must be string, otherwise defValue array is returned. 
                let tryDV (defValue1D: string[] option) (o1D: obj[])  = tryDV<string> defValue1D o1D

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
                let tryDV (defValue1D: double[] option) (o1D: obj[])  = tryDV<double> defValue1D o1D

            [<RequireQualifiedAccess>]
            module Nan =
                /// Converts an obj[] to a double[], given a default-value for non-double elements.
                let def (xlKinds: Kind[]) (defValue: double) (o1D: obj[]) =
                    o1D |> Array.map (D0.Nan.def xlKinds defValue)

                /// optional-type default
                module Opt = 
                    /// Converts an obj[] to a ('a option)[], given a default-value for non-double elements.
                    let def (xlKinds: Kind[]) (defValue: double option) (o1D: obj[]) =
                        o1D |> Array.map (D0.Nan.Opt.def xlKinds defValue)

                /// Converts an obj[] to a DateTime[], removing any non-double element.
                let filter (xlKinds: Kind[]) (o1D: obj[]) =
                    o1D |> Array.choose (D0.Nan.Opt.def xlKinds None)

                /// Converts an obj[] to an optional 'a[]. All the elements must be double, otherwise defValue array is returned. 
                let tryDV (xlKinds: Kind[]) (defValue1D: double[] option) (o1D: obj[])  =
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
    
            /// Useful functions for casting xl-arrays, given a type-tag (e.g. "int", "date", "double", "string"...)
            /// Use module Gen functions for their untyped versions.
            type TagFn =
                /// Converts an xl-value to a 'A[], given a typed default-value for elements which can't be cast to 'A.
                static member def<'A> (rowWiseDef: bool option) (defValue: obj option) (typeTag: string) (xlValue: obj) : 'A[] = 
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
                        let a1D = Nan.def Kind.nonNumericAndNA defval o1D // TODO: pass xlkinds as argument
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
                static member defOpt<'A> (rowWiseDef: bool option) (defValue: obj option) (typeTag: string) (xlValue: obj) : ('A option)[] = 
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
                        let a1D = Nan.Opt.def Kind.nonNumericAndNA defval o1D
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
                
                static member filter<'A> (rowWiseDef: bool option) (typeTag: string) (xlValue: obj) : 'A[] = 
                    let o1D = Cast.to1D (rowWiseDef |> Option.defaultValue false) xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOL -> filter<'A> o1D
                    | STRING -> 
                        let a1D = Stg.filter o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A)
                    | DOUBLE -> filter<'A> o1D
                    | DOUBLENAN -> 
                        let a1D = Nan.filter Kind.nonNumericAndNA o1D
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

                static member tryDV<'A> (rowWiseDef: bool option) (defValue1D: 'A[] option) (typeTag: string) (xlValue: obj) : 'A[] option = 
                    let o1D = Cast.to1D (rowWiseDef |> Option.defaultValue false) xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOL -> tryDV<'A> defValue1D o1D
                    | STRING -> tryDV<'A> defValue1D o1D
                    | DOUBLE -> tryDV<'A> defValue1D o1D
                    | DOUBLENAN -> 
                        let defval1D = box defValue1D :?> (double[] option)
                        let a1D = Nan.tryDV Kind.nonNumericAndNA defval1D o1D
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

                static member tryEmpty<'A> (rowWiseDef: bool option) (typeTag: string) (xlValue: obj) : 'A[] = 
                    let defValue1D : 'A[] = [||]
                    TagFn.tryDV<'A> rowWiseDef (Some defValue1D) typeTag xlValue
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
                let def (rowWiseDef: bool option) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype =
                        if typeTag.ToUpper() = "OBJ" then
                            Registry.MRegistry.trySampleType false xlValue |> Option.get // assumes a type is found. TODO: improve this? (when type not found)
                        else
                            typeTag |> Variant.labelType false
                    let args : obj[] = [| rowWiseDef; defValue; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<TagFn> "def" [| gentype |] args
                    res

                module Opt =
                    /// Converts an xl-value to a ('a option)[], given a type-tag and a compatible default-value for when casting to 'a fails.
                    /// The type-tag determines 'a. Only works for optional type-tags, e.g. "#string".
                    let def (rowWiseDef: bool option) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
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

                        let args : obj[] = [| rowWiseDef; defValue; typeTag; xlValue |]
                        let res = Useful.Generics.invoke<TagFn> "defOpt" [| gentype |] args
                        res

                /// For when the type-tag is either optional, e.g. "#string", or not, e.g. "string".
                module Any =
                    /// Convenient, single function covering def and Opt.def cases.
                    /// The returned (boxed) array might be either a 'a[] or a ('a option)[], depending on wether the type-tag is optional or not.
                    let def (rowWiseDef: bool option) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                        let gentype = typeTag |> Variant.labelType true
                        if typeTag |> isOptionalType then
                            Opt.def rowWiseDef defValue typeTag xlValue
                        else
                            def rowWiseDef defValue typeTag xlValue


                    //let defOLD (rowWiseDef: bool option) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                    //    let gentype = typeTag |> Variant.labelType true
                    //    let args : obj[] = [| rowWiseDef; defValue; typeTag; xlValue |]

                    //    let res =
                    //        if typeTag |> isOptionalType then
                    //            Useful.Generics.invoke<GenFn> "defOpt" [| gentype |] args
                    //        else
                    //            Useful.Generics.invoke<GenFn> "def" [| gentype |] args
                    //    res

                /// TODO: wording here. Mentioning the output is a (boxed) 'a[] where 'a is determined by the type tag
                // TODO explain trySampleType strict
                let filter' (rowWiseDef: bool option) (strict: bool) (typeTag: string) (xlValue: obj) : Type*obj = 
                    let gentype =
                        if typeTag.ToUpper() = "OBJ" then
                            Registry.MRegistry.trySampleType strict xlValue |> Option.get // assumes a type is found. TODO: improve this? (when type not found)
                        else
                            typeTag |> Variant.labelType false
                    let args : obj[] = [| rowWiseDef; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<TagFn> "filter" [| gentype |] args
                    gentype, res
                
                /// TODO: wording here. Mentioning the output is a (boxed) 'a[] where 'a is determined by the type tag
                let filter (rowWiseDef: bool option) (typeTag: string) (xlValue: obj) : obj = 
                    filter' rowWiseDef false typeTag xlValue |> snd

                // FIXME: wording
                let tryDV' (rowWiseDef: bool option) (defValue1D: obj) (typeTag: string) (xlValue: obj) : Type*obj = 
                    let gentype =
                        if typeTag.ToUpper() = "OBJ" then
                            Registry.MRegistry.trySampleType false xlValue |> Option.get // assumes a type is found. TODO: improve this? (when type not found)
                        else
                            typeTag |> Variant.labelType false
                    let args : obj[] = [| rowWiseDef; defValue1D; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<TagFn> "tryDV" [| gentype |] args
                    gentype, res

                // FIXME: wording
                let tryDV (rowWiseDef: bool option) (defValue1D: obj) (typeTag: string) (xlValue: obj) : obj = 
                    tryDV' rowWiseDef defValue1D typeTag xlValue |> snd

                let tryEmpty (rowWiseDef: bool option) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype =
                        if typeTag.ToUpper() = "OBJ" then
                            Registry.MRegistry.trySampleType false xlValue |> Option.get // assumes a type is found. TODO: improve this? (when type not found)
                        else
                            typeTag |> Variant.labelType false
                    let args : obj[] = [| rowWiseDef; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<TagFn> "tryEmpty" [| gentype |] args
                    res

                // FIXME: wording. Same as tryDV' with unboxing
                module Try =
                    let tryDV' (rowWiseDef: bool option) (defValue1D: obj) (typeTag: string) (xlValue: obj) : (Type*obj) option = 
                        let genty, xa1D = tryDV' rowWiseDef defValue1D typeTag xlValue
                        Useful.Option.unbox xa1D
                        |> Option.map (fun res -> (genty, res))

                    let tryDV (rowWiseDef: bool option) (defValue1D: obj) (typeTag: string) (xlValue: obj) : obj option = 
                        let xa1D = tryDV rowWiseDef defValue1D typeTag xlValue
                        let res = Useful.Option.unbox xa1D
                        res

        /// Obj[] input functions.
        module D2 =
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

            // TODO : ADD ME
            [<RequireQualifiedAccess>]
            module Nan = 
                let x1 = 0

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
                static member def<'A> (defValue: obj option) (typeTag: string) (xlValue: obj) : 'A [,] = 
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
                static member defOpt<'A> (defValue: obj option) (typeTag: string) (xlValue: obj) : ('A option)[,] = 
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
                    | INTOPT -> 
                        let defval = D0.TagFn.defaultValueOpt<double> defValue |> Option.map (int)
                        let a2D = Intg.Opt.def defval o2D
                        a2D |> Array2D.map (fun x -> (box x) :?> 'A option)
                    | DATEOPT -> 
                        let defval = D0.TagFn.defaultValueOpt<double> defValue |> Option.map (fun d -> DateTime.FromOADate(d))
                        let a2D = Dte.Opt.def defval o2D
                        a2D |> Array2D.map (fun x -> (box x) :?> 'A option)
                    | _ -> empty2D<'A option>
                
                static member filter<'A> (rowWise: bool option) (typeTag: string) (xlValue: obj) : 'A[,] = 
                    let o2D = Cast.to2D xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOL -> filter<'A> (rowWise |> Option.defaultValue true) o2D
                    | STRING -> filter<'A> (rowWise |> Option.defaultValue true) o2D
                    | DOUBLE -> filter<'A> (rowWise |> Option.defaultValue true) o2D
                    | INT -> 
                        let a2D = Intg.filter (rowWise |> Option.defaultValue true) o2D
                        a2D |> Array2D.map (fun x -> (box x) :?> 'A)
                    | DATE -> 
                        let a2D = Dte.filter (rowWise |> Option.defaultValue true) o2D
                        a2D |> Array2D.map (fun x -> (box x) :?> 'A)
                    | _ -> empty2D<'A>

                static member tryDV<'A> (defValue2D: 'A[,] option) (typeTag: string) (xlValue: obj) : 'A [,] option = 
                    let o2D = Cast.to2D xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOL -> tryDV<'A> defValue2D o2D
                    | STRING -> tryDV<'A> defValue2D o2D
                    | DOUBLE -> tryDV<'A> defValue2D o2D

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
                let def (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype = typeTag |> Variant.labelType true
                    let args : obj[] = [| defValue; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<TagFn> "def" [| gentype |] args
                    res

                module Opt =
                    /// Converts an xl-value to a ('a option)[], given a type-tag and a compatible default-value for when casting to 'a fails.
                    /// The type-tag determines 'a. Only works for optional type-tags, e.g. "#string".
                    let def (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                        let gentype = typeTag |> Variant.labelType true
                        let args : obj[] = [| defValue; typeTag; xlValue |]
                        let res = Useful.Generics.invoke<TagFn> "defOpt" [| gentype |] args
                        res

                /// For when the type-tag is either optional, e.g. "#string", or not, e.g. "string".
                module Any =
                    /// Convenient, single function covering def and Opt.def cases.
                    /// The returned (boxed) array might be either a 'a[] or a ('a option)[], depending on wether the type-tag is optional or not.
                    let def (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                        let gentype = typeTag |> Variant.labelType true
                        let args : obj[] = [| defValue; typeTag; xlValue |]

                        let res =
                            if typeTag |> isOptionalType then
                                Useful.Generics.invoke<TagFn> "defOpt" [| gentype |] args
                            else
                                Useful.Generics.invoke<TagFn> "def" [| gentype |] args
                        res

                let filter (rowWise: bool option) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype = typeTag |> Variant.labelType false
                    let args : obj[] = [| rowWise; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<TagFn> "filter" [| gentype |] args
                    res

                let tryDV' (defValue2D: obj) (typeTag: string) (xlValue: obj) : Type*obj = 
                    let gentype =
                        if typeTag.ToUpper() = "OBJ" then
                            Registry.MRegistry.trySampleType false xlValue |> Option.get // assumes a type is found. TODO: improve this? (when type not found)
                        else
                            typeTag |> Variant.labelType false
                    let args : obj[] = [| defValue2D; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<TagFn> "tryDV" [| gentype |] args
                    gentype, res

                let tryDV (defValue2D: obj) (typeTag: string) (xlValue: obj) : obj = 
                    tryDV' defValue2D typeTag xlValue |> snd

                let tryDVTBD (defValue2D: obj) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype = typeTag |> Variant.labelType false
                    let args : obj[] = [| defValue2D; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<TagFn> "tryDV" [| gentype |] args
                    res

                // FIXME: wording. Same as tryDV' with unboxing
                module Try =
                    let tryDV' (defValue2D: obj) (typeTag: string) (xlValue: obj) : (Type*obj) option = 
                        let genty, xa2D = tryDV' defValue2D typeTag xlValue
                        Useful.Option.unbox xa2D
                        |> Option.map (fun res -> (genty, res))

                    let tryDV (defValue2D: obj) (typeTag: string) (xlValue: obj) : obj option = 
                        let xa2D = tryDV defValue2D typeTag xlValue
                        let res = Useful.Option.unbox xa2D
                        res

    /// Functions to retun values to Excel.
    module Out =
        open type Variant

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
                        o0D |> Useful.Option.map proxys.none (out proxys)

            [<RequireQualifiedAccess>]
            module Prm =  // TODO : change name to Var(iant) rather than Prm?
                /// Outputs to Excel:
                ///    - Primitives-type: Returns values directly to Excel.
                ///    - Any other type : Returns ReplaveValues.object.
                let out<'a> (proxys: Proxys) (o0D: obj) : obj =
                    o0D |> Useful.Option.map proxys.none (Bxd.Any.out proxys)

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
                let out<'a> (unwrapOptions: bool) (refKey: String) (proxys: Proxys) (o0D: obj) : obj =
                    let ty = 
                        if unwrapOptions then
                            typeof<'a> |> Useful.Option.uType |> Option.defaultValue typeof<'a>
                        else
                            typeof<'a>

                    if ty |> Useful.Type.isPrimitive true then
                        o0D |> Useful.Option.map proxys.none (Bxd.Any.out proxys)
                    else
                        if unwrapOptions then
                            o0D |> Useful.Option.map proxys.none (Registry.MRegistry.append refKey >> box)
                        else
                            o0D |> Registry.MRegistry.append refKey |> box
                
                // TODO : rewrite this function
                let outO (refKey: String) (proxys: Proxys) (xlValue: obj) : obj =
                    let mapping (o: obj) =
                        if o |> isNull then // protects the `ty = o.GetType()` snippet which fails on None values at runtime (= null values at runtime).
                            proxys.none
                        else
                            let ty = o.GetType()
                            if ty |> Useful.Type.isPrimitive false then
                                o |> Useful.Option.map proxys.none (Bxd.Any.out proxys)
                            else
                                o |> Useful.Option.map proxys.none (Registry.MRegistry.append refKey >> box)

                    match Registry.MRegistry.tryExtractO xlValue with
                    | None -> proxys.failed
                    | Some regObj -> regObj |> Useful.Option.map proxys.none mapping

    // -------------------------
    // -- Convenience functions
    // -------------------------

        // default-output function
        let out<'a> (defOutput: obj) (output: 'a option) = match output with None -> defOutput | Some value -> box value
        let outNA<'a> : 'a option -> obj = out (box ExcelError.ExcelErrorNA)
        let outStg<'a> (defString: string) : 'a option -> obj = out (box defString)
        let outDbl<'a> (defNum: double) : 'a option -> obj = out (box defNum)
        let outOptx<'a> (defNum: double) : 'a option -> obj = out (box defNum)

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
                let out<'a> (unwrapOptions: bool) (refKey: String) (proxys: Proxys) (o1D: obj[]) : obj[] =
                    if o1D |> Array.isEmpty then
                        [| proxys.empty |]
                    else
                        o1D |> Array.map (D0.Reg.out<'a> unwrapOptions refKey proxys)

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
                        let res = Useful.Generics.invoke<UnboxFn> "unbox" [| ty.GetElementType() |] [| boxedArray |]
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
                        match Useful.Option.unbox boxedOptArray with
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
                        o2D |> Array2D.map (D0.Reg.out unwrapOptions refKey proxys)

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
                        let res = Useful.Generics.invoke<UnboxFn> "unbox" [| ty.GetElementType() |] [| boxedArray |]
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
                        match Useful.Option.unbox boxedOptArray with
                        | None -> None
                        | Some boxedArray -> o2D boxedArray

                    /// Convenience function, similar to o1D, but:
                    ///    - Returns [| proxys.failed |] if the input is None or if unboxing fails.
                    ///    - Applies a function to the obj[] after unboxing.
                    let apply (proxys: Proxys) (fn: obj[,] -> obj[,]) (boxedOptArray: obj) : obj[,] = 
                        match boxedOptArray |> o2D with
                        | None -> In.D2.singleton<obj> proxys.failed
                        | Some o2d -> fn o2d

module Registry_XL =
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

    [<ExcelFunction(Category="Registry", Description="Returns a registry object's type.")>]
    let rg_type 
        ([<ExcelArgument(Description= "Reg. key.")>] regKey: string)
        ([<ExcelArgument(Description= "[ToString() style. Default is false.]")>] toStringStyle: obj)
        : obj =

        // intermediary stage
        let tostringstyle = Bool.def false toStringStyle

        // result
        MRegistry.tryType regKey |> Option.map (Useful.Type.pPrint tostringstyle) |> outNA

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
        API.Out.D0.Reg.outO rfid Proxys.def regKey
        //MRegistry.tryType regKey |> Option.map (Useful.Type.pPrint tostringstyle) |> outNA

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
    open Excel
    open API
//    open type Variant
    open type Proxys

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

    [<ExcelFunction(Category="XL", Description="Cast an xl-range to DateTime[].")>]
    let cast_edgeCases ()
        : obj[,]  =

        // result
        let (lbls, dus) = Kind.labelGuide |> Array.map (fun (lbl, du) -> (box lbl, box du)) |> Array.unzip
        [| lbls; dus |] |> array2D

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
        let o1D = In.Cast.to1D rowwise range  // FIXME - should not use to1D but another In.D1.x function
        // the type annotations are NOT necessary (but are used here for readability).
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = In.D1.Stg.filter o1D
                 a1D |> Out.D1.Prm.out<string> proxys
        | "O" -> let a1D = In.D1.Stg.Opt.def None o1D
                 a1D |> Out.D1.Prm.out<string option> proxys
        // strict method: either all the array's elements are of type string, or return None (here the 1-elem array [| "failed" |])
        | "S" -> let a1D = In.D1.Stg.tryDV None o1D
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
        ([<ExcelArgument(Description= "[Edge cases mode. Edge cases values are converted to Double.NaN. E.g. NA, ERR, NN, NNNA, NNERR, NONE. Default is NONE.]")>] xlKinds: obj)
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
        //let edgecase = In.D0.Stg.def "NONE" edgeCase |> In.D0.EdgeCaseConversion.ofLabel
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
        : obj  =

        // intermediary stage
        let none = In.D0.Stg.def "<none>" noneIndicator
        let proxys = { def with none = none }
        let defVal = In.D0.Missing.Obj.tryO defaultValue

        // for demo purpose only: takes an Excel cell input,
        // converts it into a (boxed) typed value, then outputs it back to Excel.
        //let res = Out.D0.Gen.defAllCasesObj proxys defVal typeLabel xlValue
        let xa0D = In.D0.Tag.Any.def defVal typeTag xlValue
        xa0D |> Out.D0.Prm.out proxys

    [<ExcelFunction(Category="XL", Description="Cast a 1D-slice of an xl-range to a generic type 1D array.")>] // FIXME change wording
    let cast1d_gen
        ([<ExcelArgument(Description= "1D xl-range.")>] range: obj)
        ([<ExcelArgument(Description= "Type tag: bool, date, double, doubleNaN, string or obj. Add \'#'\ prefix for optional type: #bool, #date, #double, #doubleNaN, #string or #obj")>] typeTag: string)
        ([<ExcelArgument(Description= "[Replacement method for non-date elements. \"Replace\", \"Optional\" (= replace with None), \"Filter\" or \"Strict\". Default is \"Replace\".]")>] replaceMethod: obj)
        // ([<ExcelArgument(Description= "[Default Value (only for non-optional types). Must be of the appropriate type. Default \"<default>\" (which will fail for non-string types).]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[Default Value (only for non-optional types, optional types default to None). Must be of the appropriate type, else it will fail.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "[Row-wise slice direction when input is a fat, 2D, range. True or false or none. Default is none.]")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowwise = In.D0.Bool.Opt.def None rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let none = In.D0.Stg.def "<none>" noneValue
        let empty = In.D0.Stg.def "<empty>" emptyValue
        let proxys = { def with empty = empty; failed = "<failed>"; none = none }
        let defVal = In.D0.Missing.Obj.tryO defaultValue
        
        // for demo purpose only: takes an Excel range input,
        // converts it into a (boxed) typed array, then outputs it back to Excel.
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let xa1D = In.D1.Tag.filter rowwise typeTag range
                 xa1D |> (Out.D1.Unbox.apply proxys (Out.D1.Prm.out proxys))

        // strict method: either all the array's elements are of type int, or return None (here the 1-elem array [| proxys.failed |])
        | "S" -> let xa1D = In.D1.Tag.tryDV rowwise None typeTag range
                 xa1D |> (Out.D1.Unbox.Opt.apply proxys (Out.D1.Prm.out proxys))

        | _ -> let xa1D = In.D1.Tag.Any.def rowwise defVal typeTag range
               xa1D |> (Out.D1.Unbox.apply proxys (Out.D1.Prm.out proxys))

    [<ExcelFunction(Category="XL", Description="Cast a 2D xl-range to a generic type 2D array.")>]
    let cast2d_gen
        ([<ExcelArgument(Description= "2D xl-range.")>] range: obj)
        ([<ExcelArgument(Description= "Type tag: bool, date, double, doubleNaN, string or obj. Add \'#'\ prefix for optional type: #bool, #date, #double, #doubleNaN, #string or #obj")>] typeTag: string)
        ([<ExcelArgument(Description= "[Replacement method for non-date elements. \"Replace\", \"Optional\" (= replace with None), \"Filter\" or \"Strict\". Default is \"Replace\".]")>] replaceMethod: obj)
        // ([<ExcelArgument(Description= "[Default Value (only for non-optional types). Must be of the appropriate type. Default \"<default>\" (which will fail for non-string types).]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[Default Value (only for non-optional types, optional types default to None). Must be of the appropriate type, else it will fail.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "[Row wise direction. Default is none.]")>] rowWiseDirection: obj)
        : obj[,]  =

        // intermediary stage
        let rowwise = In.D0.Bool.Opt.def None rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let none = In.D0.Stg.def "<none>" noneValue
        let empty = In.D0.Stg.def "<empty>" emptyValue
        let proxys = { def with empty = empty; failed = "<failed>"; none = none }
        let defVal = In.D0.Missing.Obj.tryO defaultValue
        
        // for demo purpose only: takes an Excel range input,
        // converts it into a (boxed) typed 2D array, then outputs it back to Excel.
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let xa2D = In.D2.Tag.filter rowwise typeTag range
                 xa2D |> (Out.D2.Unbox.apply proxys (Out.D2.Prm.out proxys))

        // strict method: either all the array's elements are of type int, or return None (here the 1-elem array [| proxys.failed |])
        | "S" -> let xa2D = In.D2.Tag.tryDV None typeTag range
                 xa2D |> (Out.D2.Unbox.Opt.apply proxys (Out.D2.Prm.out proxys))

        | _ -> let xa2D = In.D2.Tag.Any.def defVal typeTag range
               xa2D |> (Out.D2.Unbox.apply proxys (Out.D2.Prm.out proxys))

module A1D = 
    open type Registry
    open Registry
    open Useful.Generics
    open API

    // -----------------------------------
    // -- Basic functions
    // -----------------------------------
    let sub' (xs: 'a[]) (startIndex: int) (subCount: int) : 'a[] =
        if startIndex >= xs.Length then
            [||]
        else
            let start = max 0 startIndex
            let count = (min (xs.Length - startIndex) subCount) |> max 0
            Array.sub xs start count
    
    let sub (startIndex: int option) (count: int option) (xs: 'a[]) : 'a[] =
        match startIndex, count with
        | Some si, Some cnt -> sub' xs si cnt
        | Some si, None -> sub' xs si (xs.Length - si)
        | None, Some cnt -> sub' xs 0 cnt
        | None, None -> xs

    // -----------------------------------
    // -- Zip functions
    // -----------------------------------
    let zip (xs1: 'a1[]) (xs2: 'a2[]) : ('a1*'a2)[] =
        if xs1.Length = 0 || xs2.Length = 0 then 
            [||]
        else
            if xs2.Length > xs1.Length then 
                Array.zip xs1 (Array.sub xs2 0 xs1.Length) 
            else 
                Array.zip (Array.sub xs1 0 xs2.Length) xs2

    let zip3 (xs1: 'a1[]) (xs2: 'a2[]) (xs3: 'a3[]) : ('a1*'a2*'a3)[] =
        if xs1.Length = 0 || xs2.Length = 0 || xs3.Length = 0 then 
            [||]
        else
            let len = Array.min [| xs1.Length; xs2.Length; xs3.Length |]
            let xs1 = Array.sub xs1 0 len
            let xs2 = Array.sub xs2 0 len
            let xs3 = Array.sub xs3 0 len
            Array.zip3 xs1 xs2 xs3

    let zip4 (xs1: 'a1[]) (xs2: 'a2[]) (xs3: 'a3[]) (xs4: 'a4[]) : ('a1*'a2*'a3*'a4)[] =
        if xs1.Length = 0 || xs2.Length = 0 || xs3.Length = 0 || xs4.Length = 0 then 
            [||]
        else
            let len = Array.min [| xs1.Length; xs2.Length; xs3.Length; xs4.Length |]
            let xs1 = Array.sub xs1 0 len
            let xs2 = Array.sub xs2 0 len
            let xs3 = Array.sub xs3 0 len
            let xs4 = Array.sub xs4 0 len
            [| for i in 0 .. (len - 1) -> (xs1.[i], xs2.[i], xs3.[i], xs4.[i]) |]

    let unzip4 (xs : ('a1*'a2*'a3*'a4)[]) : 'a1[]*'a2[]*'a3[]*'a4[] =
        let xs1 = xs |> Array.map (fun x -> let (x1, _ , _ , _ ) = x in x1)
        let xs2 = xs |> Array.map (fun x -> let (_ , x2, _ , _ ) = x in x2)
        let xs3 = xs |> Array.map (fun x -> let (_ , _ , x3, _ ) = x in x3)
        let xs4 = xs |> Array.map (fun x -> let (_ , _ , _ , x4) = x in x4)
        (xs1, xs2, xs3, xs4)

    let zip5 (xs1: 'a1[]) (xs2: 'a2[]) (xs3: 'a3[]) (xs4: 'a4[]) (xs5: 'a5[]) : ('a1*'a2*'a3*'a4*'a5)[] =
        if xs1.Length = 0 || xs2.Length = 0 || xs3.Length = 0 || xs4.Length = 0 || xs5.Length = 0 then 
            [||]
        else
            let len = Array.min [| xs1.Length; xs2.Length; xs3.Length; xs4.Length; xs5.Length |]
            let xs1 = Array.sub xs1 0 len
            let xs2 = Array.sub xs2 0 len
            let xs3 = Array.sub xs3 0 len
            let xs4 = Array.sub xs4 0 len
            let xs5 = Array.sub xs5 0 len
            [| for i in 0 .. (len - 1) -> (xs1.[i], xs2.[i], xs3.[i], xs4.[i], xs5.[i]) |]

    let zip6 (xs1: 'a1[]) (xs2: 'a2[]) (xs3: 'a3[]) (xs4: 'a4[]) (xs5: 'a5[]) (xs6: 'a6[]) : ('a1*'a2*'a3*'a4*'a5*'a6)[] =
        if xs1.Length = 0 || xs2.Length = 0 || xs3.Length = 0 || xs4.Length = 0 || xs5.Length = 0 || xs6.Length = 0 then 
            [||]
        else
            let len = Array.min [| xs1.Length; xs2.Length; xs3.Length; xs4.Length; xs5.Length; xs6.Length |]
            let xs1 = Array.sub xs1 0 len
            let xs2 = Array.sub xs2 0 len
            let xs3 = Array.sub xs3 0 len
            let xs4 = Array.sub xs4 0 len
            let xs5 = Array.sub xs5 0 len
            let xs6 = Array.sub xs6 0 len
            [| for i in 0 .. (len - 1) -> (xs1.[i], xs2.[i], xs3.[i], xs4.[i], xs5.[i], xs6.[i]) |]

    let zip7 (xs1: 'a1[]) (xs2: 'a2[]) (xs3: 'a3[]) (xs4: 'a4[]) (xs5: 'a5[]) (xs6: 'a6[]) (xs7: 'a7[]) : ('a1*'a2*'a3*'a4*'a5*'a6*'a7)[] =
        if xs1.Length = 0 || xs2.Length = 0 || xs3.Length = 0 || xs4.Length = 0 || xs5.Length = 0 || xs6.Length = 0 || xs7.Length = 0 then 
            [||]
        else
            let len = Array.min [| xs1.Length; xs2.Length; xs3.Length; xs4.Length; xs5.Length; xs6.Length; xs7.Length |]
            let xs1 = Array.sub xs1 0 len
            let xs2 = Array.sub xs2 0 len
            let xs3 = Array.sub xs3 0 len
            let xs4 = Array.sub xs4 0 len
            let xs5 = Array.sub xs5 0 len
            let xs6 = Array.sub xs6 0 len
            let xs7 = Array.sub xs7 0 len
            [| for i in 0 .. (len - 1) -> (xs1.[i], xs2.[i], xs3.[i], xs4.[i], xs5.[i], xs6.[i], xs7.[i]) |]

    let zip8 (xs1: 'a1[]) (xs2: 'a2[]) (xs3: 'a3[]) (xs4: 'a4[]) (xs5: 'a5[]) (xs6: 'a6[]) (xs7: 'a7[]) (xs8: 'a8[]) : ('a1*'a2*'a3*'a4*'a5*'a6*'a7*'a8)[] =
        if xs1.Length = 0 || xs2.Length = 0 || xs3.Length = 0 || xs4.Length = 0 || xs5.Length = 0 || xs6.Length = 0 || xs7.Length = 0 || xs8.Length = 0 then 
            [||]
        else
            let len = Array.min [| xs1.Length; xs2.Length; xs3.Length; xs4.Length; xs5.Length; xs6.Length; xs7.Length; xs8.Length |]
            let xs1 = Array.sub xs1 0 len
            let xs2 = Array.sub xs2 0 len
            let xs3 = Array.sub xs3 0 len
            let xs4 = Array.sub xs4 0 len
            let xs5 = Array.sub xs5 0 len
            let xs6 = Array.sub xs6 0 len
            let xs7 = Array.sub xs7 0 len
            let xs8 = Array.sub xs8 0 len
            [| for i in 0 .. (len - 1) -> (xs1.[i], xs2.[i], xs3.[i], xs4.[i], xs5.[i], xs6.[i], xs7.[i], xs8.[i]) |]

    let zip9 (xs1: 'a1[]) (xs2: 'a2[]) (xs3: 'a3[]) (xs4: 'a4[]) (xs5: 'a5[]) (xs6: 'a6[]) (xs7: 'a7[]) (xs8: 'a8[]) (xs9: 'a9[]) : ('a1*'a2*'a3*'a4*'a5*'a6*'a7*'a8*'a9)[] =
        if xs1.Length = 0 || xs2.Length = 0 || xs3.Length = 0 || xs4.Length = 0 || xs5.Length = 0 || xs6.Length = 0 || xs7.Length = 0 || xs8.Length = 0 || xs9.Length = 0 then 
            [||]
        else
            let len = Array.min [| xs1.Length; xs2.Length; xs3.Length; xs4.Length; xs5.Length; xs6.Length; xs7.Length; xs8.Length; xs9.Length |]
            let xs1 = Array.sub xs1 0 len
            let xs2 = Array.sub xs2 0 len
            let xs3 = Array.sub xs3 0 len
            let xs4 = Array.sub xs4 0 len
            let xs5 = Array.sub xs5 0 len
            let xs6 = Array.sub xs6 0 len
            let xs7 = Array.sub xs7 0 len
            let xs8 = Array.sub xs8 0 len
            let xs9 = Array.sub xs9 0 len
            [| for i in 0 .. (len - 1) -> (xs1.[i], xs2.[i], xs3.[i], xs4.[i], xs5.[i], xs6.[i], xs7.[i], xs8.[i], xs9.[i]) |]

    let zip10 (xs1: 'a1[]) (xs2: 'a2[]) (xs3: 'a3[]) (xs4: 'a4[]) (xs5: 'a5[]) (xs6: 'a6[]) (xs7: 'a7[]) (xs8: 'a8[]) (xs9: 'a9[]) (xs10: 'a10[]) : ('a1*'a2*'a3*'a4*'a5*'a6*'a7*'a8*'a9*'a10)[] =
        if xs1.Length = 0 || xs2.Length = 0 || xs3.Length = 0 || xs4.Length = 0 || xs5.Length = 0 || xs6.Length = 0 || xs7.Length = 0 || xs8.Length = 0 || xs9.Length = 0 || xs10.Length = 0 then 
            [||]
        else
            let len = Array.min [| xs1.Length; xs2.Length; xs3.Length; xs4.Length; xs5.Length; xs6.Length; xs7.Length; xs8.Length; xs9.Length; xs10.Length |]
            let xs1 = Array.sub xs1 0 len
            let xs2 = Array.sub xs2 0 len
            let xs3 = Array.sub xs3 0 len
            let xs4 = Array.sub xs4 0 len
            let xs5 = Array.sub xs5 0 len
            let xs6 = Array.sub xs6 0 len
            let xs7 = Array.sub xs7 0 len
            let xs8 = Array.sub xs8 0 len
            let xs9 = Array.sub xs9 0 len
            let xs10 = Array.sub xs10 0 len
            [| for i in 0 .. (len - 1) -> (xs1.[i], xs2.[i], xs3.[i], xs4.[i], xs5.[i], xs6.[i], xs7.[i], xs8.[i], xs9.[i], xs10.[i]) |]

    // -----------------------------------
    // -- Reflection functions
    // -----------------------------------
    type GenFn =
        static member out<'A> (a1D: 'A[]) (unwrapOptions: bool) (refKey: String) (proxys: Proxys) : obj[] = 
            a1D |> Array.map box |> (API.Out.D1.Reg.out<'A> unwrapOptions refKey proxys)
            
        static member count<'A> (a1D: 'A[]) : int = a1D |> Array.length

        static member elem<'A> (a1D: 'A[]) (index: int) : 'A = a1D  |> Array.item index

        static member sub<'A> (a1D: 'A[]) (startIndex: int option) (count: int option) : 'A[] =
            a1D |> sub startIndex count

        static member append2<'A> (a1D1: 'A[]) (a1D2: 'A[]) : 'A[] =
            Array.append a1D1 a1D2

        static member append3<'A> (a1D1: 'A[]) (a1D2: 'A[]) (a1D3: 'A[]) : 'A[] =
            Array.append (Array.append a1D1 a1D2) a1D3

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

            let elem (index: int) (xlValue: string) : obj option =
                let methodNm = "elem"
                MRegistry.tryFind1D xlValue
                |> Option.map (apply<GenFn> methodNm [||] [| index |])

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

module A1D_XL =
    open Registry
    open API
    open type Variant
    open type Proxys

    // open API.In.D0

    [<ExcelFunction(Category="Array1D", Description="Cast a 1D-slice of an xl-range to a generic type array.")>]
    let a1_ofRng
        ([<ExcelArgument(Description= "1D xl-range.")>] range: obj)
        ([<ExcelArgument(Description= "Type tag: bool, date, double, doubleNaN, string or obj. Add \'#'\ prefix for optional type.")>] typeTag: string)
        ([<ExcelArgument(Description= "[Replacement method for wrong-type elements. \"[R]eplace\", \"[F]ilter\", \"[S]trict\", \"[E]mptyStrict\". Default is \"Strict\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value (only for non-optional types, optional types default to None).]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        ([<ExcelArgument(Description= "[Failed value. Default is #N/A.]")>] failedValue: obj)
        : obj  =

        // intermediary stage
        let rowwise = In.D0.Bool.Opt.def None rowWiseDirection
        let replmethod = In.D0.Stg.def "STRICT" replaceMethod
        let defVal = In.D0.Absent.Obj.tryO defaultValue
        let failedval = In.D0.Missing.Obj.subst Proxys.def.failed failedValue
        let isoptional = isOptionalType typeTag
        
        // caller cell's reference ID
        let rfid = MRegistry.refID

        // wording
        match (replmethod.ToUpper().Substring(0,1)), isoptional with
        | "F", _ -> 
            let xa1D = In.D1.Tag.filter rowwise typeTag range
            let res = xa1D |> MRegistry.register rfid
            box res

        // strict / empty-strict methods: 
        //    - return a 1D array if *all* of the array's elements are of expected type (as determined by typeTag)
        // empty-strict: returns an empty array otherwise.
        // strict: return None otherwise. Here returns failed value.
        | "E", _ -> 
            let xa1D = In.D1.Tag.tryEmpty rowwise typeTag range
            let res = xa1D |> MRegistry.register rfid
            box res
        | "S", _ -> 
            match In.D1.Tag.Try.tryDV rowwise None typeTag range with
            | None -> failedval
            | Some xa1D -> 
                let res = xa1D |> MRegistry.register rfid
                box res
        | _ ->  let xa1D = In.D1.Tag.Any.def rowwise defVal typeTag range
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
        : obj = 

        // intermediary stage
        let index = In.D0.Intg.def 0 index

        let none = In.D0.Stg.def "<none>" noneIndicator
        let proxys = { def with none = none }
        let unwrapoptions = In.D0.Bool.def true unwrapOptions

        // caller cell's reference ID (necessary when the elements are not primitive types)
        let rfid = MRegistry.refID
        
        // result
        match A1D.Reg.Out.elem index rgA1D with
        | None -> proxys.failed  // TODO Unbox.apply?
        | Some o -> o |> API.Out.D0.Reg.out unwrapoptions rfid proxys

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

module A2D = 
    open type Registry
    open Registry
    open Useful.Generics
    open API

    // -----------------------------------
    // -- Main functions
    // -----------------------------------

    /// Empty 2D array.
    let empty2D<'a> : 'a[,] = [|[||]|] |> array2D

    /// Returns true if the first dimension is 0.
    let isEmpty (a2D: 'a[,]) : bool = a2D |> Array2D.length1 = 0 // is this the right way?

    /// Convenience function to create a 2D singleton.
    let singleton<'a> (a: 'a) = Array2D.create 1 1 a

    let sub' (a2D : 'a[,]) (rowStartIndex: int) (colStartIndex: int) (rowCount: int) (colCount: int) : 'a[,] =
        let rowLen, colLen = a2D |> Array2D.length1, a2D |> Array2D.length2

        if (rowStartIndex >= rowLen) || (colStartIndex >= colLen) then
            empty2D<'a>
        else
            let rowstart = max 0 rowStartIndex
            let colstart = max 0 colStartIndex
            let rowcount = (min (rowLen - rowstart) rowCount) |> max 0
            let colcount = (min (colLen - colstart) colCount) |> max 0
            a2D.[rowstart..(rowstart + rowcount - 1), colstart..(colstart + colcount - 1)]
    
    let sub (rowStartIndex: int option) (colStartIndex: int option) (rowCount: int option) (colCount: int option) (a2D : 'a[,]) : 'a[,] =
        let rowLen, colLen = a2D |> Array2D.length1, a2D |> Array2D.length2

        let rowidx = rowStartIndex |> Option.defaultValue 0
        let colidx = colStartIndex |> Option.defaultValue 0
        let rowcnt = rowCount |> Option.defaultValue rowLen
        let colcnt = colCount |> Option.defaultValue colLen
        sub' a2D rowidx colidx rowcnt colcnt
        

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

module A2D_XL =
    open Registry
    open API
    open type Proxys

    [<ExcelFunction(Category="Array2D", Description="Cast a 2D xl-range to a generic type array.")>]
    let a2_ofRng
        ([<ExcelArgument(Description= "2D xl-range.")>] range: obj)
        ([<ExcelArgument(Description= "Type tag: bool, date, double, doubleNaN, string or obj. Add \'#'\ prefix for optional type: #bool, #date, #double, #doubleNaN, #string or #obj")>] typeTag: string)
        ([<ExcelArgument(Description= "[Replacement method for non-date elements. \"Replace\", \"Optional\" (= replace with None), \"Filter\" or \"Strict\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value (only for non-optional types, optional types default to None). Must be of the appropriate type, else it will fail.]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[Row wise direction. Default is none.]")>] rowWiseDirection: obj)
        : obj  =

        // intermediary stage
        let rowwise = In.D0.Bool.Opt.def None rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let defVal = In.D0.Absent.Obj.tryO defaultValue
        
        // caller cell's reference ID
        let rfid = MRegistry.refID

        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let ooo = (In.D2.Tag.filter rowwise typeTag range)
                 let res = ooo |> MRegistry.register rfid 
                 res |> box
        | _ -> let res = (In.D2.Tag.Any.def defVal typeTag range) |> MRegistry.register rfid
               res |> box

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
        | None -> box ExcelError.ExcelErrorNA |> A2D.singleton
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
        | Some o -> o |> API.Out.D0.Reg.out unwrapoptions rfid proxys

    [<ExcelFunction(Category="Array2D", Description="Returns a sub-array of a R-object array.")>]
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

module MAP = 
    open Registry
    open Useful.Generics
    open Microsoft.FSharp.Reflection
    // open API

    // -----------------------------------
    // -- Main functions
    // -----------------------------------

    // -----------------------------------
    // -- Reflection functions
    // -----------------------------------
    type GenFn = // TODO change to Map or Refl?

        // -----------------------------------
        // -- Inspection functions
        // -----------------------------------

        /// Returns the number of kvp in the Map's object.
        static member count<'K,'V when 'K: comparison> (map: Map<'K,'V>) : int =
            map |> Map.count

        /// wording : returns keys 1D array to Excel
        static member keys<'K,'V when 'K: comparison> (map: Map<'K,'V>) (refKey: String) (proxys: Proxys) : obj[] =
            let a1D = [| for kvp in map -> kvp.Key |]
            a1D |> Array.map box |> (API.Out.D1.Reg.out<'K> false refKey proxys)

        ///// wording : returns values 1D array to Excel
        static member values<'K,'V when 'K: comparison> (map: Map<'K,'V>) (unwrapOptions: bool) (refKey: String) (proxys: Proxys) : obj[] =
            let a1D = [| for kvp in map -> kvp.Value |]
            a1D |> Array.map box |> (API.Out.D1.Reg.out<'V> unwrapOptions refKey proxys)

        static member find1<'K1,'V when 'K1: comparison> (map: Map<'K1,'V>) (proxys: Proxys) (refKey: String) (okey1: obj) : obj =
            match okey1 with 
            | :? 'K1 as key1 ->
                match map |> Map.tryFind key1 with
                | None -> proxys.failed
                | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false refKey proxys)
            | _ -> proxys.failed

        /// wording : returns values 1D array to Excel
        static member find2<'K1,'K2,'V when 'K1: comparison and 'K2: comparison> 
            (map: Map<'K1*'K2,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) 
            : obj =
                match okey1, okey2 with 
                | (:? 'K1 as key1), (:? 'K2 as key2) ->
                    match map |> Map.tryFind (key1, key2) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false refKey proxys)
                | _ -> proxys.failed

        static member find3<'K1,'K2,'K3,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison> 
            (map: Map<'K1*'K2*'K3,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) (okey3: obj) 
            : obj =
                match okey1, okey2, okey3 with 
                | (:? 'K1 as key1), (:? 'K2 as key2), (:? 'K3 as key3) ->
                    match map |> Map.tryFind (key1, key2, key3) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false refKey proxys)
                | _ -> proxys.failed

        static member find4<'K1,'K2,'K3,'K4,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison> 
            (map: Map<'K1*'K2*'K3*'K4,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) (okey3: obj) (okey4: obj)
            : obj =
                match okey1, okey2, okey3, okey4 with 
                | (:? 'K1 as key1), (:? 'K2 as key2), (:? 'K3 as key3), (:? 'K4 as key4) ->
                    match map |> Map.tryFind (key1, key2, key3, key4) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false refKey proxys)
                | _ -> proxys.failed

        static member find5<'K1,'K2,'K3,'K4,'K5,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison> 
            (map: Map<'K1*'K2*'K3*'K4*'K5,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) (okey3: obj) (okey4: obj) (okey5: obj)
            : obj =
                match okey1, okey2, okey3, okey4, okey5 with 
                | (:? 'K1 as key1), (:? 'K2 as key2), (:? 'K3 as key3), (:? 'K4 as key4), (:? 'K5 as key5) ->
                    match map |> Map.tryFind (key1, key2, key3, key4, key5) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false refKey proxys)
                | _ -> proxys.failed

        static member find6<'K1,'K2,'K3,'K4,'K5,'K6,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison> 
            (map: Map<'K1*'K2*'K3*'K4*'K5*'K6,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) (okey3: obj) (okey4: obj) (okey5: obj) (okey6: obj)
            : obj =
                match okey1, okey2, okey3, okey4, okey5, okey6 with 
                | (:? 'K1 as key1), (:? 'K2 as key2), (:? 'K3 as key3), (:? 'K4 as key4), (:? 'K5 as key5), (:? 'K6 as key6) ->
                    match map |> Map.tryFind (key1, key2, key3, key4, key5, key6) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false refKey proxys)
                | _ -> proxys.failed

        static member find7<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison> 
            (map: Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) (okey3: obj) (okey4: obj) (okey5: obj) (okey6: obj) (okey7: obj)
            : obj =
                match okey1, okey2, okey3, okey4, okey5, okey6, okey7 with 
                | (:? 'K1 as key1), (:? 'K2 as key2), (:? 'K3 as key3), (:? 'K4 as key4), (:? 'K5 as key5), (:? 'K6 as key6), (:? 'K7 as key7) ->
                    match map |> Map.tryFind (key1, key2, key3, key4, key5, key6, key7) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false refKey proxys)
                | _ -> proxys.failed

        static member find8<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'K8,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison and 'K8: comparison> 
            (map: Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) (okey3: obj) (okey4: obj) (okey5: obj) (okey6: obj) (okey7: obj) (okey8: obj)
            : obj =
                match okey1, okey2, okey3, okey4, okey5, okey6, okey7, okey8 with 
                | (:? 'K1 as key1), (:? 'K2 as key2), (:? 'K3 as key3), (:? 'K4 as key4), (:? 'K5 as key5), (:? 'K6 as key6), (:? 'K7 as key7), (:? 'K8 as key8) ->
                    match map |> Map.tryFind (key1, key2, key3, key4, key5, key6, key7, key8) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false refKey proxys)
                | _ -> proxys.failed

        static member find9<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'K8,'K9,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison and 'K8: comparison and 'K9: comparison> 
            (map: Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8*'K9,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) (okey3: obj) (okey4: obj) (okey5: obj) (okey6: obj) (okey7: obj) (okey8: obj) (okey9: obj)
            : obj =
                match okey1, okey2, okey3, okey4, okey5, okey6, okey7, okey8, okey9 with 
                | (:? 'K1 as key1), (:? 'K2 as key2), (:? 'K3 as key3), (:? 'K4 as key4), (:? 'K5 as key5), (:? 'K6 as key6), (:? 'K7 as key7), (:? 'K8 as key8), (:? 'K9 as key9) ->
                    match map |> Map.tryFind (key1, key2, key3, key4, key5, key6, key7, key8, key9) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false refKey proxys)
                | _ -> proxys.failed

        static member find10<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'K8,'K9,'K10,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison and 'K8: comparison and 'K9: comparison and 'K10: comparison> 
            (map: Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8*'K9*'K10,'V>) (proxys: Proxys) (refKey: String) 
            (okey1: obj) (okey2: obj) (okey3: obj) (okey4: obj) (okey5: obj) (okey6: obj) (okey7: obj) (okey8: obj) (okey9: obj) (okey10: obj)
            : obj =
                match okey1, okey2, okey3, okey4, okey5, okey6, okey7, okey8, okey9, okey10 with 
                | (:? 'K1 as key1), (:? 'K2 as key2), (:? 'K3 as key3), (:? 'K4 as key4), (:? 'K5 as key5), (:? 'K6 as key6), (:? 'K7 as key7), (:? 'K8 as key8), (:? 'K9 as key9), (:? 'K10 as key10) ->
                    match map |> Map.tryFind (key1, key2, key3, key4, key5, key6, key7, key8, key9, key10) with
                    | None -> proxys.failed
                    | Some a0D -> a0D |> box |> (API.Out.D0.Reg.out<'V> false refKey proxys)
                | _ -> proxys.failed

        // -----------------------------------
        // -- Construction functions
        // -----------------------------------

        /// Builds a Map<'K1,'V> key-value pairs map.
        static member map1<'K1,'V when 'K1: comparison> (keys1: 'K1[]) (values: 'V[]) 
            : Map<'K1,'V> =
            A1D.zip keys1 values |> Map.ofArray

        /// Builds a Map<'K1*'K2,'V> key-value pairs map.
        static member map2<'K1,'K2,'V when 'K1: comparison and 'K2: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (values: 'V[]) 
            : Map<'K1*'K2,'V> =
            A1D.zip (A1D.zip keys1 keys2) values |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3,'V> key-value pairs map.
        static member map3<'K1,'K2,'K3,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3,'V> =
            A1D.zip (A1D.zip3 keys1 keys2 keys3) values |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4,'V> key-value pairs map.
        static member map4<'K1,'K2,'K3,'K4,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3*'K4,'V> =
            A1D.zip (A1D.zip4 keys1 keys2 keys3 keys4) values |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5,'V> key-value pairs map.
        static member map5<'K1,'K2,'K3,'K4,'K5,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3*'K4*'K5,'V> =
            A1D.zip (A1D.zip5 keys1 keys2 keys3 keys4 keys5) values |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5*'K6,'V> key-value pairs map.
        static member map6<'K1,'K2,'K3,'K4,'K5,'K6,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (keys6: 'K6[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3*'K4*'K5*'K6,'V> =
            A1D.zip (A1D.zip6 keys1 keys2 keys3 keys4 keys5 keys6) values |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7,'V> key-value pairs map.
        static member map7<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (keys6: 'K6[]) (keys7: 'K7[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7,'V> =
            A1D.zip (A1D.zip7 keys1 keys2 keys3 keys4 keys5 keys6 keys7) values |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8,'V> key-value pairs map.
        static member map8<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'K8,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison and 'K8: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (keys6: 'K6[]) (keys7: 'K7[]) (keys8: 'K8[]) (values: 'V[]) 
            : Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8,'V> =
            A1D.zip (A1D.zip8 keys1 keys2 keys3 keys4 keys5 keys6 keys7 keys8) values |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8*'K9,'V> key-value pairs map.
        static member map9<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'K8,'K9,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison and 'K8: comparison and 'K9: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (keys6: 'K6[]) (keys7: 'K7[]) (keys8: 'K8[]) (keys9: 'K9[]) (values: 'V[])
            : Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8*'K9,'V> =
            A1D.zip (A1D.zip9 keys1 keys2 keys3 keys4 keys5 keys6 keys7 keys8 keys9) values |> Map.ofArray

        /// Builds a Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8*'K9*'K10,'V> key-value pairs map.
        static member map10<'K1,'K2,'K3,'K4,'K5,'K6,'K7,'K8,'K9,'K10,'V when 'K1: comparison and 'K2: comparison and 'K3: comparison and 'K4: comparison and 'K5: comparison and 'K6: comparison and 'K7: comparison and 'K8: comparison and 'K9: comparison and 'K10: comparison> 
            (keys1: 'K1[]) (keys2: 'K2[]) (keys3: 'K3[]) (keys4: 'K4[]) (keys5: 'K5[]) (keys6: 'K6[]) (keys7: 'K7[]) (keys8: 'K8[]) (keys9: 'K9[]) (keys10: 'K10[]) (values: 'V[])
            : Map<'K1*'K2*'K3*'K4*'K5*'K6*'K7*'K8*'K9*'K10,'V> =
            A1D.zip (A1D.zip10 keys1 keys2 keys3 keys4 keys5 keys6 keys7 keys8 keys9 keys10) values |> Map.ofArray

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

    module Gen =
        /// wording
        let mapN (gtykeys: Type[]) (gtyval: Type) (keys: obj[]) (values: obj) : obj = 
            let gtys = Array.append gtykeys [| gtyval |]
            let args : obj[] = Array.append keys [| values |]
            let methodnm = sprintf "map%d" gtykeys.Length

            let res = Useful.Generics.invoke<GenFn> methodnm gtys args
            res

        let map2D (vgtykeys: Type[]) (hgtykeys: Type[]) (gtyval: Type) (vkeys: obj[]) (hkeys: obj[]) (values: obj) : obj = 
            let gtys = Array.append (Array.append vgtykeys hgtykeys) [| gtyval |]
            let args : obj[] = Array.append (Array.append vkeys hkeys) [| values |]
            let methodnm = sprintf "mapV%dH%d" vgtykeys.Length hgtykeys.Length

            let res = Useful.Generics.invoke<GenFn> methodnm gtys args
            res

    // -----------------------------------
    // -- Registry functions
    // -----------------------------------
    module Reg =
        let genType = typeof<Map<_,_>>

        module In = 
            let pool (xlValue: obj) : obj option =
                let methodNm = "pool"
                MRegistry.tryExtractGen1D genType xlValue |> Option.map (fun (tys, objs) -> (tys, box objs))
                |> Option.map (apply<GenFn> methodNm [||] [||]) 

        module Out =
            let count (regKey: string) : obj option =
                let methodNm = "count"
                MRegistry.tryExtractGen genType regKey
                |> Option.map (apply<GenFn> methodNm [||] [||])

            let keys (regKey: string) (refKey: String) (proxys: Proxys) : obj[] option =
                let methodNm = "keys"
                MRegistry.tryExtractGen genType regKey
                |> Option.map (apply<GenFn> methodNm [||] [| refKey; proxys |])
                |> Option.map (fun o -> o :?> obj[])

            let values (regKey: string) (unwrapOptions: bool) (refKey: String) (proxys: Proxys) : obj[] option =
                let methodNm = "values"
                MRegistry.tryExtractGen genType regKey
                |> Option.map (apply<GenFn> methodNm [||] [| unwrapOptions; refKey; proxys |])
                |> Option.map (fun o -> o :?> obj[])

            let find1 (regKey: string) (proxys: Proxys) (refKey: string) (okey1: obj) : obj option =
                let methodNm = "find1"
                MRegistry.tryExtractGen genType regKey
                |> Option.map (apply<GenFn> methodNm [||] [| proxys; refKey; okey1 |])

            let findN (regKey: string) (proxys: Proxys) (refKey: string) (okeys: obj[]) : obj option =
                let args : obj[] = Array.append [| proxys; refKey |] okeys
                let methodNm = sprintf "find%d" okeys.Length
                match MRegistry.tryExtractGen genType regKey with
                | None -> None
                | Some (tys, o) -> 
                    // tys should be a [| some tuple type (for the map object's key), some other type (for the map object's value) |]
                    if tys.Length <> 2 then
                        None
                    else
                        let elemTys = FSharpType.GetTupleElements(tys.[0])
                        let genTypeRObj = (Array.append elemTys [| tys.[1] |], o)
                        apply<GenFn> methodNm [||] args genTypeRObj
                        |> Some

module MAP_XL =
    open Registry
    open API
    open type Proxys

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
        ([<ExcelArgument(Description= "Map keys1.")>] mapKeys1: obj)
        ([<ExcelArgument(Description= "Map keys2.")>] mapKeys2: obj)
        ([<ExcelArgument(Description= "Map keys3.")>] mapKeys3: obj)
        ([<ExcelArgument(Description= "Map keys4.")>] mapKeys4: obj)
        ([<ExcelArgument(Description= "Map keys5.")>] mapKeys5: obj)
        ([<ExcelArgument(Description= "Map keys6.")>] mapKeys6: obj)
        ([<ExcelArgument(Description= "Map keys7.")>] mapKeys7: obj)
        ([<ExcelArgument(Description= "Map keys8.")>] mapKeys8: obj)
        ([<ExcelArgument(Description= "Map keys9.")>] mapKeys9: obj)
        ([<ExcelArgument(Description= "Map keys10.")>] mapKeys10: obj)
        ([<ExcelArgument(Description= "Map values.")>] mapValues: obj)
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

        // caller cell's reference ID
        let rfid = MRegistry.refID

        let gtykeys_keys_gtyvals_vals =
            match ktag2, ktag3, ktag4, ktag5, ktag6, ktag7, ktag8, ktag9, ktag10 with
            | Some ktg2, Some ktg3, Some ktg4, Some ktg5, Some ktg6, Some ktg7, Some ktg8, Some ktg9, Some ktg10 -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None ktg2 mapKeys2
                let trykeys3 =  API.In.D1.Tag.Try.tryDV' None None ktg3 mapKeys3
                let trykeys4 =  API.In.D1.Tag.Try.tryDV' None None ktg4 mapKeys4
                let trykeys5 =  API.In.D1.Tag.Try.tryDV' None None ktg5 mapKeys5
                let trykeys6 =  API.In.D1.Tag.Try.tryDV' None None ktg6 mapKeys6
                let trykeys7 =  API.In.D1.Tag.Try.tryDV' None None ktg7 mapKeys7
                let trykeys8 =  API.In.D1.Tag.Try.tryDV' None None ktg8 mapKeys8
                let trykeys9 =  API.In.D1.Tag.Try.tryDV' None None ktg9 mapKeys9
                let trykeys10 =  API.In.D1.Tag.Try.tryDV' None None ktg10 mapKeys10
                let tryvals =  API.In.D1.Tag.Try.tryDV' None None valueTypeTag mapValues

                match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5, trykeys6, trykeys7, trykeys8, trykeys9, trykeys10, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5), Some (gtykey6, keys6), Some (gtykey7, keys7), Some (gtykey8, keys8), Some (gtykey9, keys9), Some (gtykey10, keys10), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5; gtykey6; gtykey7; gtykey8; gtykey9; gtykey10 |]
                    let keys = [| keys1; keys2; keys3; keys4; keys5; keys6; keys7; keys8; keys9; keys10 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

            | Some ktg2, Some ktg3, Some ktg4, Some ktg5, Some ktg6, Some ktg7, Some ktg8, Some ktg9, None -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None ktg2 mapKeys2
                let trykeys3 =  API.In.D1.Tag.Try.tryDV' None None ktg3 mapKeys3
                let trykeys4 =  API.In.D1.Tag.Try.tryDV' None None ktg4 mapKeys4
                let trykeys5 =  API.In.D1.Tag.Try.tryDV' None None ktg5 mapKeys5
                let trykeys6 =  API.In.D1.Tag.Try.tryDV' None None ktg6 mapKeys6
                let trykeys7 =  API.In.D1.Tag.Try.tryDV' None None ktg7 mapKeys7
                let trykeys8 =  API.In.D1.Tag.Try.tryDV' None None ktg8 mapKeys8
                let trykeys9 =  API.In.D1.Tag.Try.tryDV' None None ktg9 mapKeys9
                let tryvals =  API.In.D1.Tag.Try.tryDV' None None valueTypeTag mapValues

                match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5, trykeys6, trykeys7, trykeys8, trykeys9, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5), Some (gtykey6, keys6), Some (gtykey7, keys7), Some (gtykey8, keys8), Some (gtykey9, keys9), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5; gtykey6; gtykey7; gtykey8; gtykey9 |]
                    let keys = [| keys1; keys2; keys3; keys4; keys5; keys6; keys7; keys8; keys9 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None
                
            | Some ktg2, Some ktg3, Some ktg4, Some ktg5, Some ktg6, Some ktg7, Some ktg8, None, None -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None ktg2 mapKeys2
                let trykeys3 =  API.In.D1.Tag.Try.tryDV' None None ktg3 mapKeys3
                let trykeys4 =  API.In.D1.Tag.Try.tryDV' None None ktg4 mapKeys4
                let trykeys5 =  API.In.D1.Tag.Try.tryDV' None None ktg5 mapKeys5
                let trykeys6 =  API.In.D1.Tag.Try.tryDV' None None ktg6 mapKeys6
                let trykeys7 =  API.In.D1.Tag.Try.tryDV' None None ktg7 mapKeys7
                let trykeys8 =  API.In.D1.Tag.Try.tryDV' None None ktg8 mapKeys8
                let tryvals =  API.In.D1.Tag.Try.tryDV' None None valueTypeTag mapValues

                match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5, trykeys6, trykeys7, trykeys8, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5), Some (gtykey6, keys6), Some (gtykey7, keys7), Some (gtykey8, keys8), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5; gtykey6; gtykey7; gtykey8 |]
                    let keys = [| keys1; keys2; keys3; keys4; keys5; keys6; keys7; keys8 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

            | Some ktg2, Some ktg3, Some ktg4, Some ktg5, Some ktg6, Some ktg7, None, None, None -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None ktg2 mapKeys2
                let trykeys3 =  API.In.D1.Tag.Try.tryDV' None None ktg3 mapKeys3
                let trykeys4 =  API.In.D1.Tag.Try.tryDV' None None ktg4 mapKeys4
                let trykeys5 =  API.In.D1.Tag.Try.tryDV' None None ktg5 mapKeys5
                let trykeys6 =  API.In.D1.Tag.Try.tryDV' None None ktg6 mapKeys6
                let trykeys7 =  API.In.D1.Tag.Try.tryDV' None None ktg7 mapKeys7
                let tryvals =  API.In.D1.Tag.Try.tryDV' None None valueTypeTag mapValues

                match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5, trykeys6, trykeys7, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5), Some (gtykey6, keys6), Some (gtykey7, keys7), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5; gtykey6; gtykey7 |]
                    let keys = [| keys1; keys2; keys3; keys4; keys5; keys6; keys7 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

            | Some ktg2, Some ktg3, Some ktg4, Some ktg5, Some ktg6, None, None, None, None -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None ktg2 mapKeys2
                let trykeys3 =  API.In.D1.Tag.Try.tryDV' None None ktg3 mapKeys3
                let trykeys4 =  API.In.D1.Tag.Try.tryDV' None None ktg4 mapKeys4
                let trykeys5 =  API.In.D1.Tag.Try.tryDV' None None ktg5 mapKeys5
                let trykeys6 =  API.In.D1.Tag.Try.tryDV' None None ktg6 mapKeys6
                let tryvals =  API.In.D1.Tag.Try.tryDV' None None valueTypeTag mapValues

                match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5, trykeys6, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5), Some (gtykey6, keys6), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5; gtykey6 |]
                    let keys = [| keys1; keys2; keys3; keys4; keys5; keys6 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

            | Some ktg2, Some ktg3, Some ktg4, Some ktg5, None, None, None, None, None -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None ktg2 mapKeys2
                let trykeys3 =  API.In.D1.Tag.Try.tryDV' None None ktg3 mapKeys3
                let trykeys4 =  API.In.D1.Tag.Try.tryDV' None None ktg4 mapKeys4
                let trykeys5 =  API.In.D1.Tag.Try.tryDV' None None ktg5 mapKeys5
                let tryvals =  API.In.D1.Tag.Try.tryDV' None None valueTypeTag mapValues

                match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5 |]
                    let keys = [| keys1; keys2; keys3; keys4; keys5 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

            | Some ktg2, Some ktg3, Some ktg4, None, None, None, None, None, None -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None ktg2 mapKeys2
                let trykeys3 =  API.In.D1.Tag.Try.tryDV' None None ktg3 mapKeys3
                let trykeys4 =  API.In.D1.Tag.Try.tryDV' None None ktg4 mapKeys4
                let tryvals =  API.In.D1.Tag.Try.tryDV' None None valueTypeTag mapValues

                match trykeys1, trykeys2, trykeys3, trykeys4, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4 |]
                    let keys = [| keys1; keys2; keys3; keys4 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

            | Some ktg2, Some ktg3, None, None, None, None, None, None, None -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None ktg2 mapKeys2
                let trykeys3 =  API.In.D1.Tag.Try.tryDV' None None ktg3 mapKeys3
                let tryvals =  API.In.D1.Tag.Try.tryDV' None None valueTypeTag mapValues

                match trykeys1, trykeys2, trykeys3, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2; gtykey3 |]
                    let keys = [| keys1; keys2; keys3 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

            | Some ktg2, None, None, None, None, None, None, None, None -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None k1TypeTag mapKeys1
                let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None ktg2 mapKeys2
                let tryvals =  API.In.D1.Tag.Try.tryDV' None None valueTypeTag mapValues

                match trykeys1, trykeys2, tryvals with
                | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1; gtykey2 |]
                    let keys = [| keys1; keys2 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

            | _ -> 
                let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None k1TypeTag mapKeys1
                let tryvals =  API.In.D1.Tag.Try.tryDV' None None valueTypeTag mapValues

                match trykeys1, tryvals with
                | Some (gtykey1, keys1), Some (gtyval, vals) -> 
                    let gtykeys = [| gtykey1 |]
                    let keys = [| keys1 |]
                    Some (gtykeys, keys, gtyval, vals)
                | _ -> None

        match gtykeys_keys_gtyvals_vals with
        | None -> Proxys.def.failed
        | Some (gtykeys, keys, gtyval, vals) ->
            let map = MAP.Gen.mapN gtykeys gtyval keys vals
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

        let tryvals =  API.In.D2.Tag.Try.tryDV' None valueTag mapValues

        match tryvals with
        | None -> Proxys.def.failed
        | Some (gtyval, vals) ->
            let vgtykeys_keys =
                match vktag2, vktag3, vktag4, vktag5, vktag6 with
                | Some vktg2, Some vktg3, Some vktg4, Some vktg5, Some vktg6 -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None vktag1 mapVKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None vktg2 mapVKeys2
                    let trykeys3 =  API.In.D1.Tag.Try.tryDV' None None vktg3 mapVKeys3
                    let trykeys4 =  API.In.D1.Tag.Try.tryDV' None None vktg4 mapVKeys4
                    let trykeys5 =  API.In.D1.Tag.Try.tryDV' None None vktg5 mapVKeys5
                    let trykeys6 =  API.In.D1.Tag.Try.tryDV' None None vktg6 mapVKeys6

                    match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5, trykeys6 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5), Some (gtykey6, keys6) -> 
                        let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5; gtykey6 |]
                        let keys = [| keys1; keys2; keys3; keys4; keys5; keys6 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | Some vktg2, Some vktg3, Some vktg4, Some vktg5, None -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None vktag1 mapVKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None vktg2 mapVKeys2
                    let trykeys3 =  API.In.D1.Tag.Try.tryDV' None None vktg3 mapVKeys3
                    let trykeys4 =  API.In.D1.Tag.Try.tryDV' None None vktg4 mapVKeys4
                    let trykeys5 =  API.In.D1.Tag.Try.tryDV' None None vktg5 mapVKeys5

                    match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5) -> 
                        let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5 |]
                        let keys = [| keys1; keys2; keys3; keys4; keys5 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | Some vktg2, Some vktg3, Some vktg4, None, None -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None vktag1 mapVKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None vktg2 mapVKeys2
                    let trykeys3 =  API.In.D1.Tag.Try.tryDV' None None vktg3 mapVKeys3
                    let trykeys4 =  API.In.D1.Tag.Try.tryDV' None None vktg4 mapVKeys4

                    match trykeys1, trykeys2, trykeys3, trykeys4 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4) -> 
                        let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4 |]
                        let keys = [| keys1; keys2; keys3; keys4 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | Some vktg2, Some vktg3, None, None, None -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None vktag1 mapVKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None vktg2 mapVKeys2
                    let trykeys3 =  API.In.D1.Tag.Try.tryDV' None None vktg3 mapVKeys3

                    match trykeys1, trykeys2, trykeys3 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3) -> 
                        let gtykeys = [| gtykey1; gtykey2; gtykey3 |]
                        let keys = [| keys1; keys2; keys3 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | Some vktg2, None, None, None, None -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None vktag1 mapVKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None vktg2 mapVKeys2

                    match trykeys1, trykeys2 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2) -> 
                        let gtykeys = [| gtykey1; gtykey2 |]
                        let keys = [| keys1; keys2 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | _ -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None vktag1 mapVKeys1

                    match trykeys1 with
                    | Some (gtykey1, keys1) -> 
                        let gtykeys = [| gtykey1 |]
                        let keys = [| keys1 |]
                        Some (gtykeys, keys)
                    | _ -> None

            let hgtykeys_keys =
                match hktag2, hktag3, hktag4, hktag5, hktag6 with
                | Some hktg2, Some hktg3, Some hktg4, Some hktg5, Some hktg6 -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None hktag1 mapHKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None hktg2 mapHKeys2
                    let trykeys3 =  API.In.D1.Tag.Try.tryDV' None None hktg3 mapHKeys3
                    let trykeys4 =  API.In.D1.Tag.Try.tryDV' None None hktg4 mapHKeys4
                    let trykeys5 =  API.In.D1.Tag.Try.tryDV' None None hktg5 mapHKeys5
                    let trykeys6 =  API.In.D1.Tag.Try.tryDV' None None hktg6 mapHKeys6

                    match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5, trykeys6 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5), Some (gtykey6, keys6) -> 
                        let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5; gtykey6 |]
                        let keys = [| keys1; keys2; keys3; keys4; keys5; keys6 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | Some hktg2, Some hktg3, Some hktg4, Some hktg5, None -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None hktag1 mapHKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None hktg2 mapHKeys2
                    let trykeys3 =  API.In.D1.Tag.Try.tryDV' None None hktg3 mapHKeys3
                    let trykeys4 =  API.In.D1.Tag.Try.tryDV' None None hktg4 mapHKeys4
                    let trykeys5 =  API.In.D1.Tag.Try.tryDV' None None hktg5 mapHKeys5

                    match trykeys1, trykeys2, trykeys3, trykeys4, trykeys5 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4), Some (gtykey5, keys5) -> 
                        let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4; gtykey5 |]
                        let keys = [| keys1; keys2; keys3; keys4; keys5 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | Some hktg2, Some hktg3, Some hktg4, None, None -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None hktag1 mapHKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None hktg2 mapHKeys2
                    let trykeys3 =  API.In.D1.Tag.Try.tryDV' None None hktg3 mapHKeys3
                    let trykeys4 =  API.In.D1.Tag.Try.tryDV' None None hktg4 mapHKeys4

                    match trykeys1, trykeys2, trykeys3, trykeys4 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3), Some (gtykey4, keys4) -> 
                        let gtykeys = [| gtykey1; gtykey2; gtykey3; gtykey4 |]
                        let keys = [| keys1; keys2; keys3; keys4 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | Some hktg2, Some hktg3, None, None, None -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None hktag1 mapHKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None hktg2 mapHKeys2
                    let trykeys3 =  API.In.D1.Tag.Try.tryDV' None None hktg3 mapHKeys3

                    match trykeys1, trykeys2, trykeys3 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2), Some (gtykey3, keys3) -> 
                        let gtykeys = [| gtykey1; gtykey2; gtykey3 |]
                        let keys = [| keys1; keys2; keys3 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | Some hktg2, None, None, None, None -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None hktag1 mapHKeys1
                    let trykeys2 =  API.In.D1.Tag.Try.tryDV' None None hktg2 mapHKeys2

                    match trykeys1, trykeys2 with
                    | Some (gtykey1, keys1), Some (gtykey2, keys2) -> 
                        let gtykeys = [| gtykey1; gtykey2 |]
                        let keys = [| keys1; keys2 |]
                        Some (gtykeys, keys)
                    | _ -> None

                | _ -> 
                    let trykeys1 =  API.In.D1.Tag.Try.tryDV' None None hktag1 mapHKeys1

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
                let map = MAP.Gen.map2D vgtykeys hgtykeys gtyval vkeys hkeys vals
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
        match MAP.Reg.Out.count rgMap with
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
        match MAP.Reg.Out.keys rgMap rfid Proxys.def with
        | None -> [| Proxys.def.failed |]  // TODO Unbox.apply?
        | Some o1D -> o1D

    [<ExcelFunction(Category="Map", Description="Returns a R-object map's values array.")>]
    let map_vals
        ([<ExcelArgument(Description= "Map R-object.")>] rgMap: string) 
        : obj[] = 

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match MAP.Reg.Out.values rgMap true rfid Proxys.def with
        | None -> [| Proxys.def.failed |]  // TODO Unbox.apply?
        | Some o1D -> o1D

    [<ExcelFunction(Category="Map", Description="Returns a R-object map's values array.")>]
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
        ([<ExcelArgument(Description= "[Failed value. Default is #N/A.]")>] failedValue: obj)
        ([<ExcelArgument(Description= "[Tuppled. Default is true.]")>] tuppled: obj)  
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
        let failedval = In.D0.Missing.Obj.subst Proxys.def.failed failedValue
        let proxys = { def with failed = failedval }

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
            match MAP.Reg.Out.find1 rgMap Proxys.def rfid okeys.[0] with
            | None -> proxys.failed
            | Some o0D -> o0D
        else
            match MAP.Reg.Out.findN rgMap Proxys.def rfid okeys with
            | None -> proxys.failed
            | Some o0D -> o0D

        
        //match mapkey2, mapkey3, mapkey4, mapkey5, mapkey6, mapkey7, mapkey8, mapkey9, mapkey10 with
        //| Some key2, Some key3, Some key4, Some key5, Some key6, Some key7, Some key8, Some key9, Some key10 -> 
        //    box "9 keys case here"
        //| Some key2, Some key3, Some key4, Some key5, Some key6, Some key7, Some key8, Some key9, None -> 
        //    box "9 keys case here"
        //| Some key2, Some key3, Some key4, Some key5, Some key6, Some key7, Some key8, None, None -> 
        //    box "8 keys case here"
        //| Some key2, Some key3, Some key4, Some key5, Some key6, Some key7, None, None, None -> 
        //    box "7 keys case here"
        //| Some key2, Some key3, Some key4, Some key5, Some key6, None, None, None, None -> 
        //    box "6 keys case here"
        //| Some key2, Some key3, Some key4, Some key5, None, None, None, None, None -> 
        //    box "5 keys case here"
        //| Some key2, Some key3, Some key4, None, None, None, None, None, None -> 
        //    box "4 keys case here"
        //| Some key2, Some key3, None, None, None, None, None, None, None -> 
        //    match MAP.Reg.Out.findN rgMap Proxys.def rfid [| mapKey1; key2; key3 |] with
        //    | None -> proxys.failed
        //    | Some o0D -> o0D
        //| Some key2, None, None, None, None, None, None, None, None -> 
        //    match MAP.Reg.Out.findN rgMap Proxys.def rfid [| mapKey1; key2 |] with
        //    | None -> proxys.failed
        //    | Some o0D -> o0D
        //| _ -> 
        //    match MAP.Reg.Out.find1 rgMap Proxys.def rfid mapKey1 with
        //    | None -> proxys.failed
        //    | Some o0D -> o0D


    [<ExcelFunction(Category="Map", Description="Returns the union of many compatible Map<K,V> R-objects.")>]
    let map_pool
        ([<ExcelArgument(Description= "Map R-objects.")>] rgMap1D: obj) 
        : obj = 

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match MAP.Reg.In.pool rgMap1D with
        | None -> Proxys.def.failed  // TODO Unbox.apply?
        | Some regObjMap -> regObjMap |> MRegistry.register rfid |> box








/// Simple template for generics
module GenMtrx =
    open type Registry
    open Registry

    open Useful.Generics
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
            let a0D = In.D0.Tag.Any.def defValue typeLabel xlValue :?> 'A
            a0D |> create0D size

        static member mtrx1D<'A> (defValue: obj option) (typeLabel: string) (size: int) (xlValue: obj) : GenMTRX<'A> =
            let a1D = In.D1.TagFn.def None defValue typeLabel xlValue
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

    open Useful.Generics
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
            let a0D = In.D0.TagFn.def defValue typeLabel xlValue
            a0D |> create0D size

        static member mtrx1D<'A> (defValue: obj option) (typeLabel: string) (size: int) (xlValue: obj) : MTRX<'A> =
            let a1D = In.D1.TagFn.def None defValue typeLabel xlValue
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
    open type Variant
    open type Proxys
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





























