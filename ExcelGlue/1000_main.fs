namespace ExcelGlue

open System
open System.Collections.Generic
open ExcelDna.Integration

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

    ///// Removes object and its xl-cell reference from the Object Registry.
    //member this.remove (key: string) = 
    //    reg.Remove(key) |> ignore
    //    ref.Remove(key)

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

    member this.tryExtractS<'a> (regKeyOrString: string) : 'a option =
        match this.tryFind regKeyOrString with
        | None -> if typeof<'a> = typeof<string> then Some (unbox (box regKeyOrString)) else None
        | Some regObj -> 
            match regObj with
            | :? 'a as v -> Some v
            | _ -> None

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

    member this.tryExtract1D (regKey: string) : ((Type[])*obj) option =
        match this.tryFind regKey with
        | None -> None
        | Some regObj ->
            let ty = regObj.GetType()

            if not ty.IsArray then
                None
            else
                let genty = ty.GetElementType()
                ([| genty |], regObj) 
                |> Some

    // https://docs.microsoft.com/en-us/dotnet/framework/reflection-and-codedom/how-to-examine-and-instantiate-generic-types-with-reflection
    // https://docs.microsoft.com/en-us/dotnet/api/system.type.getgenerictypedefinition?view=net-5.0
    /// targetGenType is the expected generic type, e.g. targetGenType: typeof<GenMTRX<_>>
    member this.tryExtractGen' (targetGenType: Type) (xlValue: string) : obj option = 
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
    member this.tryExtractGen (targetType: Type) (xlValue: string) : ((Type[])*obj) option =
        this.tryExtractGen' targetType xlValue
        |> Option.map (fun o -> ((o.GetType().GetGenericArguments()), o))

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

    //static member labelDefValTY<'a> (typeTag: string) (defValue: obj option) : 'a = 
    //    match defValue with
    //    | None -> Variant.labelDefVal typeTag 
    //    | Some dv -> dv
    //    :?> 'a

/// Replacement values to return to Excel instead of Double.NaN, None, and [||]. TODO : wording, description of each element
type ReplaceValues = { nan: obj; none: obj; empty: obj; object: obj; error: obj } with
    static member def : ReplaceValues = { nan = ExcelError.ExcelErrorNA; none = "<none>"; empty = "<empty>"; object = "<obj>"; error = box ExcelError.ExcelErrorNA }


type MTRXD = double[,]
type MTRX<'a> = 'a[,]

module Test =
    
    type myType = 
        static member myMember1<'T> (arg1: obj) : 'T[] = [||]
        static member myMember2<'T> (arg2: obj) : 'T[,] = [|[||]|] |> array2D


module Registry =
    /// Master registry where all registered objects are held.
    let MRegistry = Registry()

    // -----------------------------------
    // -- Excel RefID
    // -----------------------------------

    let refID (caller : obj) : string = 
        match caller with
        | :? ExcelReference as ref -> ref.ToString()
        | _                        -> ""

//[<RequireQualifiedAccess>]
module Useful =
    //open System

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
            let pp = s.Replace(someType.Namespace + ".","").Replace("System.", "")
            pp

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

module API = 

    /// Functions to handle or process Inputs from Excel.
    module In =

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
            //open System
            //open ExcelDna.Integration
            open Excel
            open Excel.Kind

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
                
            //// ----------------
            //// -- 1D functions
            //// ----------------

            ///// Slices xl-ranges into 1D arrays.
            ///// 1-row or 1-column xl-ranges are directly converted to their corresponding 1D array.
            ///// For 2D xl-range, the 1st row is returned if rowWiseDef is true, and the 1st column otherwise.
            //let slice2D (rowWiseDef: bool) (o2D: obj[,]) : obj[] = 
            //    // column-wise slice
            //    if Array2D.length2 o2D = 1 then
            //        o2D.[*, Array2D.base2 o2D]
            //    // row-wise slice
            //    elif rowWiseDef || (Array2D.length1 o2D = 1) then
            //        o2D.[Array2D.base1 o2D, *]
            //    // column-wise slice as default for fat 2D arrays
            //    else 
            //        o2D.[*, Array2D.base2 o2D]
        
            ///// Converts an xl-value to a 2D array.
            ///// (Use to2D rather than try2D when the obj argument is an xl-value).
            //let to2D (rowWise: bool) (xlVal: obj) : obj[,] =
            //    match xlVal with
            //    | :? (obj[,]) as o2D -> o2D
            //    | :? (obj[]) as o1D -> [| o1D |] |> array2D // FIXME - transpose // THIS CASE SHD NOT OCCUR if xlVal is an excel arg
            //    | o0D -> Array2D.create 1 1 o0D
        
            ///// Converts an obj to a 2D array option.
            ///// (Use try2D rather than to2D when the obj argument is not an xl-value).
            //let try2D (rowWise: bool) (o: obj) : obj[,] option =
            //    match o with
            //    | :? (obj[,]) as o2D -> o2D |> Some
            //    | :? (obj[]) as o1D -> [| o1D |] |> array2D |> Some // FIXME - transpose
            //    | _ -> None

            ///// Converts an xl-value to a 1D array.
            ///// (Use to1D rather than try1D when the obj argument is an xl-value).
            //let to1D (rowWiseDef: bool) (xlVal: obj) : obj[] =
            //    match xlVal with
            //    | :? (obj[,]) as o2D -> slice2D rowWiseDef o2D
            //    | :? (obj[]) as o1D -> o1D
            //    | o0D -> [| o0D |]
        
            ///// Converts an obj to a 1D array option.
            ///// (Use try1D rather than to1D when the obj argument is not an xl-value).
            //let try1D (rowWiseDef: bool) (o: obj) : obj[] option =
            //    match o with
            //    | :? (obj[,]) as o2D -> slice2D rowWiseDef o2D |> Some
            //    | :? (obj[]) as o1D -> Some o1D
            //    | _ -> None

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
                    let def (defValue: DateTime option) (xlVal: obj) =
                        match xlVal with
                        | :? double as d -> DateTime.FromOADate d |> Some
                        | _ -> defValue

                /// Object substitution, based on type.
                module Obj =
                    /// Substitutes an object for another one, if it isn't a (boxed) DateTime (e.g. box 36526.0).
                    let subst (defValue: obj) (xlVal: obj) : obj =
                        match xlVal with
                        | :? double as d -> xlVal
                        | _ -> defValue

            [<RequireQualifiedAccess>]
            module Missing = // TODO add empty case
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

            /// Convenience functions to cast default xl-values, given a type-tag.
            module private DefaultValue =
                /// Returns a typed-value given an optional xl-value and a type-tag.
                /// If present, the xl-value must be compatible with the type-tag.
                /// If no xl-value is provided, returns the type-tag's default value.
                let ofTag<'a> (typeTag: string) (xlValue: obj option) : 'a = 
                    let def = Variant.labelDefVal typeTag
                    match xlValue with
                    | None -> 
                        def :?> 'a
                    | Some xlval -> // FIXME : does not work for empty, only works for missing
                        let dv =
                            if typeTag.ToUpper() = "INT" then 
                                Intg.def (def :?> int) xlval |> box
                                //dv :?> double |> int |> box
                            elif typeTag.ToUpper() = "DATE" then 
                                // dv :?> double |> (fun d -> DateTime.FromOADate(d)) |> box
                                Dte.def (def :?> DateTime) xlval |> box
                            else
                                xlval
                        dv :?> 'a

                /// Returns an optional typed-value given an optional xl-value and a type-tag.
                /// If present, the xl-value must be compatible with type-tag.
                module Opt = 
                    let ofTag<'a> (defValue: obj option) : 'a option = 
                        match defValue with
                        | None -> None
                        | Some o ->
                            match o with
                            | :? 'a as a -> Some a
                            | :? ('a option) as aopt -> aopt
                            | _ -> None

            type GenFn =
                /// Casts an xl-value to a 'A, with a default-value for when the casting fails.
                static member def<'A> (defValue: obj option) (typeTag: string) (xlValue: obj) : 'A = 

                    match typeTag |> Variant.ofTag with
                    | BOOL -> 
                        let defval = DefaultValue.ofTag<'A> typeTag defValue
                        def<'A> defval xlValue
                    | STRING -> 
                        let defval = DefaultValue.ofTag<'A> typeTag defValue
                        def<'A> defval xlValue
                    | DOUBLE -> 
                        let defval = DefaultValue.ofTag<'A> typeTag defValue
                        def<'A> defval xlValue
                    | INT -> 
                        let defval = DefaultValue.ofTag<int> typeTag defValue |> int
                        let a0D = Intg.def defval xlValue
                        box a0D :?> 'A
                    | DATE -> 
                        let defval = DefaultValue.ofTag<DateTime> typeTag defValue
                        let a0D = Dte.def defval xlValue
                        box a0D :?> 'A
                    | _ -> failwith "TO BE IMPLEMENTED WITH OTHER VARIANT TYPES" // TODO FIXME
    
                /// Casts an xl-value to a 'A option, with a default-value for when the casting fails.
                /// defValue is None, Some 'a or even Some (Some 'a).
                static member defOpt<'A> (defValue: obj option) (typeTag: string) (xlValue: obj) : 'A option = 
                    match typeTag |> Variant.ofTag with
                    | BOOLOPT -> 
                        let defval : 'A option = DefaultValue.Opt.ofTag<'A> defValue
                        Opt.def<'A> defval xlValue
                    | STRINGOPT -> 
                        let defval : 'A option = DefaultValue.Opt.ofTag<'A> defValue
                        Opt.def<'A> defval xlValue
                    | DOUBLEOPT -> 
                        let defval : 'A option = DefaultValue.Opt.ofTag<'A> defValue
                        Opt.def<'A> defval xlValue
                    | INTOPT -> 
                        let defval = DefaultValue.Opt.ofTag<double> defValue |> Option.map (int)
                        let a0D = Intg.Opt.def defval xlValue
                        box a0D :?> 'A option
                    | DATEOPT -> 
                        let defval = DefaultValue.Opt.ofTag<double> defValue |> Option.map (fun d -> DateTime.FromOADate(d))
                        let a0D = Dte.Opt.def defval xlValue
                        box a0D :?> 'A option
                    | _ -> failwith "TO BE IMPLEMENTED WITH OTHER VARIANT TYPES" // TODO FIXME


            [<RequireQualifiedAccess>]
            module Gen = 
                /// Casts an xl-value to a 'A, with a default-value for when the casting fails.
                /// 'a is determined by typeTag.
                let def (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype = typeTag |> Variant.labelType true
                    let args : obj[] = [| defValue; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<GenFn> "def" [| gentype |] args
                    res

                // optional-type default FIXX
                module Opt =
                    /// Casts an xl-value to a 'a option, with a default-value for when the casting fails.
                    /// 'a is determined by typeTag.
                    let def (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                        let gentype = typeTag |> Variant.labelType true
                        let args : obj[] = [| defValue; typeTag; xlValue |]
                        let res = Useful.Generics.invoke<GenFn> "defOpt" [| gentype |] args
                        res

                /// For when the type-tag is either optional, e.g. "#string", or not, e.g. "string".
                module Any =
                    /// Convenient, single function covering def and Opt.def cases.
                    /// The returned (boxed) value might be either a 'a or a ('a option), depending on wether the type-tag is optional or not.
                    let def (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                        let gentype = typeTag |> Variant.labelType true
                        let args : obj[] = [| defValue; typeTag; xlValue |]

                        let res =
                            if typeTag |> isOptionalType then
                                Useful.Generics.invoke<GenFn> "defOpt" [| gentype |] args
                            else
                                Useful.Generics.invoke<GenFn> "def" [| gentype |] args
                        res



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
            let tryDV<'a> (defValue: 'a[] option) (o1D: obj[]) : 'a[] option =
                let convert = Opt.def None o1D
                match convert |> Array.tryFind Option.isNone with
                | None -> convert |> Array.map Option.get |> Some
                | Some _ -> defValue

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
                let tryDV (defValue: bool[] option) (o1D: obj[])  = tryDV<bool> defValue o1D

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
                let tryDV (defValue: string[] option) (o1D: obj[])  = tryDV<string> defValue o1D

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
                let tryDV (defValue: double[] option) (o1D: obj[])  = tryDV<double> defValue o1D

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
                let tryDV (xlKinds: Kind[]) (defValue: double[] option) (o1D: obj[])  =
                    let convert = Opt.def xlKinds None o1D
                    match convert |> Array.tryFind Option.isNone with
                    | None -> convert |> Array.map Option.get |> Some
                    | Some _ -> defValue

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
                let tryDV (defValue: int[] option) (o1D: obj[])  =
                    let convert = Opt.def None o1D
                    match convert |> Array.tryFind Option.isNone with
                    | None -> convert |> Array.map Option.get |> Some
                    | Some _ -> defValue

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
                let tryDV (defValue: DateTime[] option) (o1D: obj[])  =
                    let convert = Opt.def None o1D
                    match convert |> Array.tryFind Option.isNone with
                    | None -> convert |> Array.map Option.get |> Some
                    | Some _ -> defValue
    
            /// Convenience functions to cast default xl-values, given a type-tag.
            module private DefaultValue =
                /// Returns a typed-value given an optional xl-value and a type-tag.
                /// If present, the xl-value must be compatible with the type-tag.
                /// If no xl-value is provided, returns the type-tag's default value.
                let ofTag<'a> (typeTag: string) (xlValue: obj option) : 'a = 
                    match xlValue with
                    | None -> Variant.labelDefVal typeTag 
                    | Some xlval -> xlval
                    :?> 'a
                
                /// Returns an optional typed-value given an optional xl-value and a type-tag.
                /// If present, the xl-value must be compatible with type-tag.
                module Opt = 
                    let ofTag<'a> (xlValue: obj option) : 'a option = 
                        match xlValue with
                        | None -> None
                        | Some xlval ->
                            match xlval with
                            | :? 'a as a -> Some a
                            | :? ('a option) as aopt -> aopt
                            | _ -> None

            /// Useful functions for casting xl-arrays, given a type-tag (e.g. "int", "date", "double", "string"...)
            /// Use module Gen functions for their untyped versions.
            type GenFn =
                /// Converts an xl-value to a 'A[], given a typed default-value for elements which can't be cast to 'A.
                static member def<'A> (rowWiseDef: bool option) (defValue: obj option) (typeTag: string) (xlValue: obj) : 'A[] = 
                    let o1D = Cast.to1D (rowWiseDef |> Option.defaultValue false) xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOL -> 
                        let defval = DefaultValue.ofTag<'A> typeTag defValue
                        def<'A> defval o1D
                    | STRING -> 
                        let defval = DefaultValue.ofTag<'A> typeTag defValue
                        def<'A> defval o1D
                    | DOUBLE -> 
                        let defval = DefaultValue.ofTag<'A> typeTag defValue
                        def<'A> defval o1D
                    | INT -> 
                        let defval = DefaultValue.ofTag<double> typeTag defValue |> int
                        let a1D = Intg.def defval o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A)
                    | DATE -> 
                        let defval = DateTime.FromOADate(DefaultValue.ofTag<double> typeTag defValue)
                        let a1D = Dte.def defval o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A)
                    | _ -> [||]
        
                /// Converts an xl-value to a ('A option)[], given an optional default-value for elements which can't be cast to 'A.
                static member defOpt<'A> (rowWiseDef: bool option) (defValue: obj option) (typeTag: string) (xlValue: obj) : ('A option)[] = 
                    let o1D = Cast.to1D (rowWiseDef |> Option.defaultValue false) xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOLOPT -> 
                        let defval : 'A option = DefaultValue.Opt.ofTag<'A> defValue
                        Opt.def<'A> defval o1D
                    | STRINGOPT -> 
                        let defval : 'A option = DefaultValue.Opt.ofTag<'A> defValue
                        Opt.def<'A> defval o1D
                    | DOUBLEOPT -> 
                        let defval : 'A option = DefaultValue.Opt.ofTag<'A> defValue
                        Opt.def<'A> defval o1D
                    | INTOPT -> 
                        let defval = DefaultValue.Opt.ofTag<double> defValue |> Option.map (int)
                        let a1D = Intg.Opt.def defval o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A option)
                    | DATEOPT -> 
                        let defval = DefaultValue.Opt.ofTag<double> defValue |> Option.map (fun d -> DateTime.FromOADate(d))
                        let a1D = Dte.Opt.def defval o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A option)
                    | _ -> [||]
                
                static member filter<'A> (rowWiseDef: bool option) (typeTag: string) (xlValue: obj) : 'A[] = 
                    let o1D = Cast.to1D (rowWiseDef |> Option.defaultValue false) xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOL -> filter<'A> o1D
                    | STRING -> 
                        let a1D = Stg.filter o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A)
                    | DOUBLE -> filter<'A> o1D
                    | INT -> 
                        let a1D = Intg.filter o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A)
                    | DATE -> 
                        let a1D = Dte.filter o1D
                        a1D |> Array.map (fun x -> (box x) :?> 'A)
                    | _ -> [||]

            /// Useful functions for casting xl-arrays, given a type-tag (e.g. "int", "date", "double", "string"...)
            /// Use type GenFn functions for their typed versions.
            [<RequireQualifiedAccess>]
            module Gen =
                /// Converts an xl-value to a 'a[], given a type-tag and a compatible default-value for when casting to 'a fails.
                /// The type-tag determines 'a. Only works for non-optional type-tags, e.g. "string".
                let def (rowWiseDef: bool option) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype = typeTag |> Variant.labelType true
                    let args : obj[] = [| rowWiseDef; defValue; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<GenFn> "def" [| gentype |] args
                    res

                module Opt =
                    /// Converts an xl-value to a ('a option)[], given a type-tag and a compatible default-value for when casting to 'a fails.
                    /// The type-tag determines 'a. Only works for optional type-tags, e.g. "#string".
                    let def (rowWiseDef: bool option) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                        let gentype = typeTag |> Variant.labelType true
                        let args : obj[] = [| rowWiseDef; defValue; typeTag; xlValue |]
                        let res = Useful.Generics.invoke<GenFn> "defOpt" [| gentype |] args
                        res

                /// For when the type-tag is either optional, e.g. "#string", or not, e.g. "string".
                module Any =
                    /// Convenient, single function covering def and Opt.def cases.
                    /// The returned (boxed) array might be either a 'a[] or a ('a option)[], depending on wether the type-tag is optional or not.
                    let def (rowWiseDef: bool option) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                        let gentype = typeTag |> Variant.labelType true
                        let args : obj[] = [| rowWiseDef; defValue; typeTag; xlValue |]

                        let res =
                            if typeTag |> isOptionalType then
                                Useful.Generics.invoke<GenFn> "defOpt" [| gentype |] args
                            else
                                Useful.Generics.invoke<GenFn> "def" [| gentype |] args
                        res
                
                /// TODO: wording here. Mentioning the output is a (boxed) 'a[] where 'a is determined by the type tag
                let filter (rowWiseDef: bool option) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype = typeTag |> Variant.labelType false
                    let args : obj[] = [| rowWiseDef; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<GenFn> "filter" [| gentype |] args
                    res


        /// Obj[] input functions.
        module D2 =
            open Excel
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
            let tryDV<'a> (defValue: 'a[,] option) (o2D: obj[,]) : 'a[,] option =
                let len1 = o2D |> Array2D.length1
                let convert = Opt.def None o2D

                let hasNones = 
                    [| for i in 0 .. (len1 - 1) ->
                        convert.[i,*] |> Array.filter Option.isNone
                    |]
                    |> Array.filter (fun o1D -> o1D |> Array.isEmpty |> not)

                if hasNones |> Array.isEmpty then
                    convert |> Array2D.map Option.get |> Some
                else defValue

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
                let tryDV (defValue: bool[,] option) (o2D: obj[,]) : bool[,] option = tryDV<bool> defValue o2D

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
                let tryDV (defValue: string[,] option) (o2D: obj[,]) : string[,] option = tryDV<string> defValue o2D

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
                let tryDV (defValue: double[,] option) (o2D: obj[,]) : double[,] option = tryDV<double> defValue o2D

            // TODO : ADD ME
            [<RequireQualifiedAccess>]
            module Nan = 
                let x1 = 0

            // TODO : ADD ME
            [<RequireQualifiedAccess>]
            module Intg = 
                let x1 = 0

            // TODO : ADD ME
            [<RequireQualifiedAccess>]
            module Dte = 
                let x1 = 0

            /// Convenience functions to cast default xl-values, given a type-tag.
            module private DefaultValue =
                /// Returns a typed-value given an optional xl-value and a type-tag.
                /// If present, the xl-value must be compatible with the type-tag.
                /// If no xl-value is provided, returns the type-tag's default value.
                let ofTag<'a> (typeTag: string) (xlValue: obj option) : 'a = 
                    match xlValue with
                    | None -> Variant.labelDefVal typeTag 
                    | Some xlval -> xlval
                    :?> 'a
                
                /// Returns an optional typed-value given an optional xl-value and a type-tag.
                /// If present, the xl-value must be compatible with type-tag.
                module Opt = 
                    let ofTag<'a> (xlValue: obj option) : 'a option = 
                        match xlValue with
                        | None -> None
                        | Some xlval ->
                            match xlval with
                            | :? 'a as a -> Some a
                            | :? ('a option) as aopt -> aopt
                            | _ -> None

            /// Useful functions for casting xl-arrays, given a type-tag (e.g. "int", "date", "double", "string"...)
            /// Use module Gen functions for their untyped versions.
            type GenFn =
                /// Converts an xl-value to a 'A[], given a typed default-value for elements which can't be cast to 'A.
                static member def<'A> (defValue: obj option) (typeTag: string) (xlValue: obj) : 'A [,] = 
                    let o2D = Cast.to2D xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOL -> 
                        let defval = DefaultValue.ofTag<'A> typeTag defValue
                        def<'A> defval o2D
                    | STRING -> 
                        let defval = DefaultValue.ofTag<'A> typeTag defValue
                        def<'A> defval o2D
                    | DOUBLE -> 
                        let defval = DefaultValue.ofTag<'A> typeTag defValue
                        def<'A> defval o2D
                        // FIXME RESTORE !!!!!!!!!!!
                    //| INT -> 
                    //    let defval = DefaultValue.ofTag<double> typeTag defValue |> int
                    //    let a2D = Intg.def defval o2D
                    //    a2D |> Array.map (fun x -> (box x) :?> 'A)
                    //| DATE -> 
                    //    let defval = DateTime.FromOADate(DefaultValue.ofTag<double> typeTag defValue)
                    //    let a2D = Dte.def defval o2D
                    //    a2D |> Array.map (fun x -> (box x) :?> 'A)
                    | _ -> empty2D<'A>
        
                /// Converts an xl-value to a ('A option)[], given an optional default-value for elements which can't be cast to 'A.
                static member defOpt<'A> (defValue: obj option) (typeTag: string) (xlValue: obj) : ('A option)[,] = 
                    let o2D = Cast.to2D xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOLOPT -> 
                        let defval : 'A option = DefaultValue.Opt.ofTag<'A> defValue
                        Opt.def<'A> defval o2D
                    | STRINGOPT -> 
                        let defval : 'A option = DefaultValue.Opt.ofTag<'A> defValue
                        Opt.def<'A> defval o2D
                    | DOUBLEOPT -> 
                        let defval : 'A option = DefaultValue.Opt.ofTag<'A> defValue
                        Opt.def<'A> defval o2D
                    | INTOPT -> 
                        empty2D<'A option>
                        //let defval = DefaultValue.Opt.ofTag<double> defValue |> Option.map (int)
                        //let a2D = Intg.Opt.def defval o2D
                        //a2D |> Array.map (fun x -> (box x) :?> 'A option)
                    | DATEOPT -> 
                        empty2D<'A option>
                        //let defval = DefaultValue.Opt.ofTag<double> defValue |> Option.map (fun d -> DateTime.FromOADate(d))
                        //let a2D = Dte.Opt.def defval o2D
                        //a2D |> Array.map (fun x -> (box x) :?> 'A option)
                    | _ -> empty2D<'A option>
                
                static member filter<'A> (rowWise: bool option) (typeTag: string) (xlValue: obj) : 'A[,] = 
                    let o2D = Cast.to2D xlValue

                    match typeTag |> Variant.ofTag with
                    | BOOL -> filter<'A> (rowWise |> Option.defaultValue true) o2D
                    | STRING -> filter<'A> (rowWise |> Option.defaultValue true) o2D
                    | DOUBLE -> filter<'A> (rowWise |> Option.defaultValue true) o2D
                    //| INT -> 
                    //    let a2D = Intg.filter o2D
                    //    a2D |> Array.map (fun x -> (box x) :?> 'A)
                    //| DATE -> 
                    //    let a2D = Dte.filter o2D
                    //    a2D |> Array.map (fun x -> (box x) :?> 'A)
                    | _ -> empty2D<'A>

            /// Useful functions for casting xl-arrays, given a type-tag (e.g. "int", "date", "double", "string"...)
            /// Use type GenFn functions for their typed versions.
            [<RequireQualifiedAccess>]
            module Gen =
                /// Converts an xl-value to a 'a[], given a type-tag and a compatible default-value for when casting to 'a fails.
                /// The type-tag determines 'a. Only works for non-optional type-tags, e.g. "string".
                let def (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype = typeTag |> Variant.labelType true
                    let args : obj[] = [| defValue; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<GenFn> "def" [| gentype |] args
                    res

                module Opt =
                    /// Converts an xl-value to a ('a option)[], given a type-tag and a compatible default-value for when casting to 'a fails.
                    /// The type-tag determines 'a. Only works for optional type-tags, e.g. "#string".
                    let def (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                        let gentype = typeTag |> Variant.labelType true
                        let args : obj[] = [| defValue; typeTag; xlValue |]
                        let res = Useful.Generics.invoke<GenFn> "defOpt" [| gentype |] args
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
                                Useful.Generics.invoke<GenFn> "defOpt" [| gentype |] args
                            else
                                Useful.Generics.invoke<GenFn> "def" [| gentype |] args
                        res

                let filter (rowWise: bool option) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype = typeTag |> Variant.labelType false
                    let args : obj[] = [| rowWise; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<GenFn> "filter" [| gentype |] args
                    res

    /// Functions to retun values to Excel.
    module Out =
        open type Variant

        ///// Returns an xl-Value from a typed value. 
        ///// NaN elements are converted according to replaceValues.
        //let cellTBD<'a> (replaceValues: ReplaceValues) (a0D: 'a) : obj =
        //    let xlval = box a0D
        //    if typeof<'a> = typeof<double> then
        //        if Double.IsNaN(xlval :?> double) then replaceValues.nan else xlval
        //    else
        //        xlval

        ///// Returns an xl-Value from a typed value. 
        ///// None and NaN elements are converted according to replaceValues.
        ///// Some 'a elements will be output as would 'a.
        //let cellOptTBD<'a> (replaceValues: ReplaceValues) (a0D: 'a option) : obj =
        //    match a0D with
        //    | None -> replaceValues.none
        //    | Some a0d -> cellTBD<'a> replaceValues a0d

        ///// Returns an xl 1D-range, or a default-singleton instead of an empty array. 
        ///// NaN elements are converted according to replaceValues.
        //let range1D<'a> (replaceValues: ReplaceValues) (a1D: 'a[]) : obj[] =
        //    if a1D |> Array.isEmpty then
        //        [| replaceValues.empty |]
        //    else
        //        if typeof<'a> = typeof<double> then
        //            a1D |> Array.map (fun num -> let xlval = box num in if Double.IsNaN(xlval :?> double) then replaceValues.nan else xlval)
        //        else
        //            a1D |> Array.map box

        ///// Returns an xl 1D-range, or a default-singleton instead of an empty array. 
        ///// None and NaN elements are converted according to replaceValues.
        //let range1DOpt<'a> (replaceValues: ReplaceValues) (a1D: ('a option)[]) : obj[] =
        //    if a1D |> Array.isEmpty then
        //        [| replaceValues.empty |]
        //    else
        //        if typeof<'a> = typeof<double> then
        //            a1D 
        //            |> Array.map 
        //                (fun elem -> 
        //                    match elem with 
        //                    | None -> replaceValues.none 
        //                    | Some num -> let xlval = box num in if Double.IsNaN(xlval :?> double) then replaceValues.nan else xlval
        //                )
        //        else
        //            a1D |> Array.map (fun elem -> match elem with | None -> replaceValues.none | Some e -> box e)



        /// Functions to return single values to Excel.
        module D0 =
            /// Outputs double values to Excel, 
            /// with conversion of Double.NaN values, if any.
            [<RequireQualifiedAccess>]
            module Dbl =
                let out (replaceValues: ReplaceValues) (d: double) : obj =
                    if Double.IsNaN(d) then
                        replaceValues.nan
                    else    
                        d |> box

            /// Outputs primitive types back to Excel.
            // https://docs.microsoft.com/en-us/office/client-developer/excel/data-types-used-by-excel
            [<RequireQualifiedAccess>]
            module Bxd =  // TODO : change name to Var(iant) rather than Primitive?
                /// Returns sensible Excel values for non-optional (boxed) primitive types.
                let out (replaceValues: ReplaceValues) (o0D: obj) : obj =
                    match o0D with
                    | :? double as d -> Dbl.out replaceValues d
                    | :? string | :? DateTime | :? int | :? bool -> o0D
                    | _ -> replaceValues.object

                [<RequireQualifiedAccess>]
                /// Outputs optional primitive types:
                ///    - None will return replaceValues.none
                ///    - Some x will return (Bxd.out x)
                module Opt = 
                    let out (replaceValues: ReplaceValues) (o0D: obj option) : obj =
                        match o0D with
                        | None -> replaceValues.none
                        | Some o0d -> o0d |> out replaceValues

                [<RequireQualifiedAccess>]
                /// Outputs optional and non-optional primitive types.
                /// Option on primitive types (boxed) will return as follow: 
                ///    - None will return replaceValues.none
                ///    - Some x will return (Bxd.out x)
                module Any = 
                    let out (replaceValues: ReplaceValues) (o0D: obj) : obj =
                        o0D |> Useful.Option.map replaceValues.none (out replaceValues)


            [<RequireQualifiedAccess>]
            module Prm =  // TODO : change name to Var(iant) rather than Prm?
                /// Outputs to Excel:
                ///    - Primitives-type: Returns values directly to Excel.
                ///    - Any other type : Returns ReplaveValues.object.
                let out<'a> (replaceValues: ReplaceValues) (o0D: obj) : obj =
                    o0D |> Useful.Option.map replaceValues.none (Bxd.Any.out replaceValues)

            [<RequireQualifiedAccess>]
            /// Outputs primitives types directly to Excel, but stores non-primitive types in the Registry.
            module Reg = 
                /// Outputs to Excel:
                ///    - Primitives-type: Returns values directly to Excel.
                ///    - R-object : Returns a Registry keys and stores values in the Registry.
                let out<'a> (refKey: String) (replaceValues: ReplaceValues) (o0D: obj) : obj =
                    let ty = typeof<'a> |> Useful.Option.uType |> Option.defaultValue typeof<'a>
                
                    if ty |> Useful.Type.isPrimitive true then
                        o0D |> Useful.Option.map replaceValues.none (Bxd.Any.out replaceValues)
                    else    
                        o0D |> Useful.Option.map replaceValues.none (Registry.MRegistry.append refKey >> box)

            type GenFn =                    
                /// Same as In.D0.Gen.def, but returns an obj instead of a 'A.
                static member defObj<'A> (replaceValues: ReplaceValues) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                    let a0D = In.D0.GenFn.def<'A> defValue typeTag xlValue
                    Bxd.out replaceValues (box a0D)

                /// Same as In.D0.Gen.tryDV, but returns an obj instead of a 'A. // FIXME change name?!
                static member tryDVObj<'A> (replaceValues: ReplaceValues) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                    let a0D = In.D0.GenFn.defOpt<'A> defValue typeTag xlValue
                    // cellOpt<'A> replaceValues a0D
                    Bxd.Any.out replaceValues a0D


            // DO WE NEED THESE? check out cast_gen
            [<RequireQualifiedAccess>]
            module Gen = 
                /// Same as In.D0.Gen.def, but returns an obj instead of a 'a.
                // TODO: for OUT there might not be a need for defValue (?)
                // TODO: xlValue for OUT?!?
                let defObj (replaceValues: ReplaceValues) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype = typeTag |> Variant.labelType true
                    let args : obj[] = [| replaceValues; defValue; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<GenFn> "defObj" [| gentype |] args
                    res

                /// Same as In.D0.Gen.tryDV, but returns an obj instead of a 'a.
                let tryDVObj (replaceValues: ReplaceValues) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype = typeTag |> Variant.labelType true
                    let args : obj[] = [| replaceValues; defValue; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<GenFn> "tryDVObj" [| gentype |] args
                    res

                /// Same as In.D0.Gen.defAllCases but returns a obj instead of 'a.
                let defAllCasesObj (replaceValues: ReplaceValues) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj = 
                    let gentype = typeTag |> Variant.labelType true
                    let args : obj[] = [| replaceValues; defValue; typeTag; xlValue |]

                    let res =
                        if typeTag |> isOptionalType then
                            Useful.Generics.invoke<GenFn> "tryDVObj" [| gentype |] args
                        else
                            Useful.Generics.invoke<GenFn> "defObj" [| gentype |] args
                    res

        /// Functions to return 1D arrays back to Excel.
        module D1 =
            open Excel
            open type Variant

            /// Outputs primitive type arrays back to Excel.// TODO: DEFINE PRIMITIVE TYPES
            /// Non-primitive types will return as #VALUE! 
            [<RequireQualifiedAccess>]
            module Bxd =  // TODO : change name to Var(iant) rather than Primitive?
                /// Returns sensible Excel values for primitive types.  // TODO adapt wording to arrays
                /// Empty arrays will return [| replaceValues.empty |].
                /// (Excel naturally returns empty array values as #VALUE!).
                let out (replaceValues: ReplaceValues) (o1D: obj[]) : obj[] =
                    if o1D |> Array.isEmpty then
                        [| replaceValues.empty |]
                    else
                        o1D |> Array.map (D0.Bxd.out replaceValues)

                [<RequireQualifiedAccess>]
                /// Outputs optional and non-optional primitive types.  // TODO adapt wording to arrays
                module Any = 
                    /// Option on primitive types will return as follow : 
                    ///    - None will return replaceValues.none
                    ///    - Some x will return (Prm.out x)
                    let out (replaceValues: ReplaceValues) (o1D: obj[]) : obj[] =
                        if o1D |> Array.isEmpty then
                            [| replaceValues.empty |]
                        else
                            o1D |> Array.map (D0.Bxd.Any.out replaceValues)

            /// Outputs primitives types directly to Excel, but stores non-primitive types in the Registry.
            [<RequireQualifiedAccess>]
            module Prm = 
                /// Outputs to Excel: FIXME check wording
                ///    - Primitives-type: Returns values directly to Excel.
                ///    - Any other type : Returns ReplaveValues.object.
                let out<'a> (replaceValues: ReplaceValues) (a1D: 'a[]) : obj[] =
                    a1D 
                    |> Array.map box
                    |> Bxd.Any.out replaceValues

            /// Outputs primitives types directly to Excel, but stores non-primitive types in the Registry.
            [<RequireQualifiedAccess>]
            module Reg = 
                /// Outputs to Excel: FIXME check wording
                ///    - Primitives-type: Returns values directly to Excel.
                ///    - R-object : Returns a Registry keys and stores values in the Registry.
                let out<'a> (refKey: String) (replaceValues: ReplaceValues) (o1D: obj[]) : obj[] =
                    if o1D |> Array.isEmpty then
                        [| replaceValues.empty |]
                    else
                        o1D |> Array.map (D0.Reg.out refKey replaceValues)

            type UnboxFn =
                static member unbox<'A> (a1D: 'A[]) : obj[] = a1D |> Array.map box

            [<RequireQualifiedAccess>]
            module Unbox = 
                /// Unboxes a (boxed 'a[]) into a (boxed 'a)[].
                /// In other words, unboxes a obj = box ('a[]) into a obj[].
                /// Returns None if the unboxing fails.
                let o1D (boxedArray: obj) : obj[] option = 
                    let ty = boxedArray.GetType()
                    if not ty.IsArray then
                        None
                    else
                        let res = Useful.Generics.invoke<UnboxFn> "unbox" [| ty.GetElementType() |] [| boxedArray |]
                        res :?> obj[]
                        |> Some

                /// Convenience function, similar to o1D but :
                ///    - returns [| replaceValues.error |] if the unboxing fails.
                ///    - applies a mapping function to each element after the unboxing.
                let apply (replaceValues: ReplaceValues) (fn: obj[] -> obj[]) (boxedArray: obj) : obj[] = 
                    match boxedArray |> o1D with
                    | None -> [| replaceValues.error |]
                    | Some o1D -> fn o1D

            //type GenFnTBD =
            //    static member unboxO1D<'A> (a1D: 'A[]) : obj[] = a1D |> Array.map box

            //    // TODO DO WE NEED THESE FUNCTIONS?

            //    /// Same as Gen.def, but returns an obj[] instead of a 'a[].
            //    static member defObj<'A> (rowWiseDef: bool option) (replaceValues: ReplaceValues) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj[] = 
            //        let a1D = In.D1.GenFn.def<'A> rowWiseDef defValue typeTag xlValue
            //        range1D<'A> replaceValues a1D
                    
            //    // FIXME
            //    static member defOptObj<'A> (rowWiseDef: bool option) (replaceValues: ReplaceValues) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj[] = 
            //        let a1D = In.D1.GenFn.defOpt<'A> rowWiseDef defValue typeTag xlValue
            //        range1DOpt<'A> replaceValues a1D

            //    /// Same as In.D1.Gen.filter, but returns an obj[] instead of a 'a[].
            //    static member filterObj<'A> (rowWiseDef: bool option) (replaceValues: ReplaceValues) (typeTag: string) (xlValue: obj) : obj[] = 
            //        let a1D = In.D1.GenFn.filter<'A> rowWiseDef typeTag xlValue
            //        range1D<'A> replaceValues a1D

            //[<RequireQualifiedAccess>]
            //module GenTBDx =

            //    // TODO DO WE NEED THESE FUNCTIONS? SHOULD THESE FUNCTIONS BE IN IN!!!   FIXX ME
            //    // maybe useful for testing / inspection. E.g. take an xl-range and apply these functions, see the result directly into Excel. (direct from xl to xl...)
            //    // ???
            //    /// Same as In.D1.Gen.def but returns a obj[], rather than a (boxed) 'a[].
            //    let defObj (rowWiseDef: bool option) (replaceValues: ReplaceValues) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj[] = 
            //        let gentype = typeTag |> Variant.labelType true
            //        let args : obj[] = [| rowWiseDef; replaceValues; defValue; typeTag; xlValue |]
            //        let res = Useful.Generics.invoke<GenFnTBD> "defObj" [| gentype |] args
            //        res :?> obj[]

            //    /// Same as In.D1.Gen.defOpt but returns a obj[], rather than a (boxed) ('a option)[].
            //    let defOptObj (rowWiseDef: bool option) (replaceValues: ReplaceValues) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj[] = 
            //        let gentype = typeTag |> Variant.labelType true
            //        let args : obj[] = [| rowWiseDef; replaceValues; defValue; typeTag; xlValue |]
            //        let res = Useful.Generics.invoke<GenFnTBD> "defOptObj" [| gentype |] args
            //        res :?> obj[]

            //    // TODO change name to Any.def
            //    /// Same as In.D1.Gen.defAllCases but returns a obj[], rather than a (boxed) ('a option)[].
            //    let defAllCasesObj (rowWiseDef: bool option) (replaceValues: ReplaceValues) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj[] = 
            //        let gentype = typeTag |> Variant.labelType true
            //        let args : obj[] = [| rowWiseDef; replaceValues; defValue; typeTag; xlValue |]

            //        let res =
            //            if typeTag |> isOptionalType then
            //                Useful.Generics.invoke<GenFnTBD> "defOptObj" [| gentype |] args
            //            else
            //                Useful.Generics.invoke<GenFnTBD> "defObj" [| gentype |] args
            //        res :?> obj[]

            //    /// Same as In.D1.Gen.filter.
            //    let filterObj (rowWiseDef: bool option) (replaceValues: ReplaceValues) (typeTag: string) (xlValue: obj) : obj[] = 
            //        let gentype = typeTag |> Variant.labelType false
            //        let args : obj[] = [| rowWiseDef; replaceValues; typeTag; xlValue |]
            //        //let res = callMethod "filterObj" gentype args
            //        let res = Useful.Generics.invoke<GenFnTBD> "filterObj" [| gentype |] args
            //        res :?> obj[]

        // -----------------------------------
        // -- Convenience functions
        // -----------------------------------

        // default-output function
        let out (defOutput: obj) (output: 'a option) = match output with None -> defOutput | Some value -> box value
        let outNA<'a> : 'a option -> obj = out (box ExcelError.ExcelErrorNA)
        let outStg<'a> (defString: string) : 'a option -> obj = out (box defString)
        let outDbl<'a> (defNum: double) : 'a option -> obj = out (box defNum)

        let outOpt<'a> (defNum: double) : 'a option -> obj = out (box defNum)


        /// Obj[,] output functions.
        module D2 =
            open Excel
            open type Variant

            /// Returns an xl 2D-range, or a default-singleton instead of an empty array. 
            /// NaN elements are converted according to replaceValues.
            let range2D<'a> (replaceValues: ReplaceValues) (a2D: 'a[,]) : obj[,] =
                if a2D |> In.D2.isEmpty then
                    In.D2.singleton replaceValues.empty
                else
                    if typeof<'a> = typeof<double> then
                        a2D |> Array2D.map (fun num -> let xlval = box num in if Double.IsNaN(xlval :?> double) then replaceValues.nan else xlval)
                    else
                        a2D |> Array2D.map box

            /// Returns an xl 1D-range, or a default-singleton instead of an empty array. 
            /// None and NaN elements are converted according to replaceValues.
            let range2DOpt<'a> (replaceValues: ReplaceValues) (a2D: ('a option)[,]) : obj[,] =
                if a2D |> In.D2.isEmpty then
                    In.D2.singleton replaceValues.empty
                else
                    if typeof<'a> = typeof<double> then
                        a2D 
                        |> Array2D.map 
                            (fun elem -> 
                                match elem with 
                                | None -> replaceValues.none 
                                | Some num -> let xlval = box num in if Double.IsNaN(xlval :?> double) then replaceValues.nan else xlval
                            )
                    else
                        a2D |> Array2D.map (fun elem -> match elem with | None -> replaceValues.none | Some e -> box e)

            /// Outputs primitives types directly to Excel, but stores non-primitive types in the Registry.
            [<RequireQualifiedAccess>]
            module Reg = 
                /// Outputs to Excel: FIXME check wording
                ///    - Primitives-type: Returns values directly to Excel.
                ///    - R-object : Returns a Registry keys and stores values in the Registry.
                let out<'a> (refKey: String) (replaceValues: ReplaceValues) (o2D: obj[,]) : obj[,] =
                    if o2D |> Array2D.length1 = 0 then  // TODO : what's the best way to check o2D is empty??????
                        Array2D.create 1 1 replaceValues.empty
                    else
                        o2D |> Array2D.map (D0.Reg.out refKey replaceValues)

            type GenFn =
                // TODO DO WE NEED THESE FUNCTIONS? yes checj cast2d_gen illustration purposes

                /// Same as Gen.def, but returns an obj[] instead of a 'a[].
                static member defObj<'A> (replaceValues: ReplaceValues) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj[,] = 
                    let a2D = In.D2.GenFn.def<'A> defValue typeTag xlValue
                    range2D<'A> replaceValues a2D
                    
                // FIXME
                static member defOptObj<'A> (replaceValues: ReplaceValues) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj[,] = 
                    let a2D = In.D2.GenFn.defOpt<'A> defValue typeTag xlValue
                    range2DOpt<'A> replaceValues a2D

                /// Same as In.D1.Gen.filter, but returns an obj[] instead of a 'a[].
                static member filterObj<'A> (rowWise: bool option) (replaceValues: ReplaceValues) (typeTag: string) (xlValue: obj) : obj[,] = 
                    let a2D = In.D2.GenFn.filter<'A> rowWise typeTag xlValue
                    range2D<'A> replaceValues a2D

            [<RequireQualifiedAccess>]
            module Gen =
                // TODO DO WE NEED THESE FUNCTIONS? SHOULD THESE FUNCTIONS BE IN IN!!!   FIXX ME
                // maybe useful for testing / inspection. E.g. take an xl-range and apply these functions, see the result directly into Excel. (direct from xl to xl...)
                // ???
                /// Same as In.D1.Gen.def but returns a obj[], rather than a (boxed) 'a[].
                let defObj (replaceValues: ReplaceValues) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj[,] = 
                    let gentype = typeTag |> Variant.labelType true
                    let args : obj[] = [| replaceValues; defValue; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<GenFn> "defObj" [| gentype |] args
                    res :?> obj[,]

                /// Same as In.D1.Gen.defOpt but returns a obj[], rather than a (boxed) ('a option)[].
                let defOptObj (replaceValues: ReplaceValues) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj[,] = 
                    let gentype = typeTag |> Variant.labelType true
                    let args : obj[] = [| replaceValues; defValue; typeTag; xlValue |]
                    let res = Useful.Generics.invoke<GenFn> "defOptObj" [| gentype |] args
                    res :?> obj[,]

                // TODO change name to Any.def
                /// Same as In.D1.Gen.defAllCases but returns a obj[], rather than a (boxed) ('a option)[].
                let defAllCasesObj (replaceValues: ReplaceValues) (defValue: obj option) (typeTag: string) (xlValue: obj) : obj[,] = 
                    let gentype = typeTag |> Variant.labelType true
                    let args : obj[] = [| replaceValues; defValue; typeTag; xlValue |]
                    let res =
                        if typeTag |> isOptionalType then
                            Useful.Generics.invoke<GenFn> "defOptObj" [| gentype |] args
                        else
                            Useful.Generics.invoke<GenFn> "defObj" [| gentype |] args
                    res :?> obj[,]

                /// Same as In.D1.Gen.filter.
                let filterObj (rowWise: bool option) (replaceValues: ReplaceValues) (typeTag: string) (xlValue: obj) : obj[,] = 
                    let gentype = typeTag |> Variant.labelType false
                    let args : obj[] = [| rowWise; replaceValues; typeTag; xlValue |]
                    //let res = callMethod "filterObj" gentype args
                    let res = Useful.Generics.invoke<GenFn> "filterObj" [| gentype |] args
                    res :?> obj[,]



module A1D = 
    open type Registry
    open Registry
    open Useful.Generics
    open API

    // -----------------------------------
    // -- Main library
    // -----------------------------------
    let sub' (xs : 'a[]) (startIndex: int) (subCount: int) : 'a[] =
        if startIndex >= xs.Length then
            [||]
        else
            let start = max 0 startIndex
            let count = (min (xs.Length - startIndex) subCount) |> max 0
            Array.sub xs start count
    
    let sub (startIndex: int option) (count: int option) (xs : 'a[]) : 'a[] =
        match startIndex, count with
        | Some si, Some cnt -> sub' xs si cnt
        | Some si, None -> sub' xs si (xs.Length - si)
        | None, Some cnt -> sub' xs 0 cnt
        | None, None -> xs

    // -----------------------------------
    // -- Reflection functions
    // -----------------------------------
    type GenFn =
        static member out<'A> (a1D: 'A[]) (refKey: String) (replaceValues: ReplaceValues) : obj[] = 
            a1D |> Array.map box |> (API.Out.D1.Reg.out<'A> refKey replaceValues)
            
        static member count<'A> (a1D: 'A[]) : int = a1D |> Array.length

        static member elem<'A> (a1D: 'A[]) (index: int) : 'A = a1D |> Array.item index

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
        module In =
            let mtrx0D (defValue: obj option) (typeTag: string) (size: int) (xlValue: obj) : obj = 
                let gentype = typeTag |> Variant.labelType false
                let args : obj[] = [| defValue; typeTag; size; xlValue |]
                let res = invoke<GenFn> "mtrx0D" [| gentype |] args
                res

        module Out =
            let out (regKey: string) (refKey: String) (replaceValues: ReplaceValues) : obj[] option =
                let methodNm = "out"
                MRegistry.tryExtract1D regKey
                |> Option.map (apply<GenFn> methodNm [||] [| refKey; replaceValues |])
                |> Option.map (fun o -> o :?> obj[])

            let count (xlValue: string) : obj option =
                let methodNm = "count"
                MRegistry.tryExtract1D xlValue
                |> Option.map (apply<GenFn> methodNm [||] [||])

            let elem (index: int) (xlValue: string) : obj option =
                let methodNm = "elem"
                MRegistry.tryExtract1D xlValue
                |> Option.map (apply<GenFn> methodNm [||] [| index |])

            let sub (regKey: string) (startIndex: int option) (count: int option) : obj option =
                let methodNm = "sub"
                MRegistry.tryExtract1D regKey
                |> Option.map (apply<GenFn> methodNm [||] [| startIndex; count |])

            let private append2' (tys1:Type[], o1:obj) (tys2:Type[], o2:obj) : obj option =
                let methodNm = "append2"
                if tys2 <> tys1 then
                    None
                else
                    applyMulti<GenFn> methodNm [||] [||] tys1 [| o1; o2 |]
                    |> Some

            let append2 (regKey1: string)  (regKey2: string) : obj option =
                match MRegistry.tryExtract1D regKey1, MRegistry.tryExtract1D regKey2 with
                | None, None -> None
                | Some (_, o1), None -> Some o1
                | None, Some (_, o2) -> Some o2
                | Some (tys1, o1), Some (tys2, o2) -> append2' (tys1, o1) (tys2, o2)

            let append3 (regKey1: string)  (regKey2: string)  (regKey3: string) : obj option =
                let methodNm = "append3"
                match MRegistry.tryExtract1D regKey1, MRegistry.tryExtract1D regKey2, MRegistry.tryExtract1D regKey3 with
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

module Cast_XL =
    open Excel
    open API
    open type Variant
    open type ReplaceValues



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
        ([<ExcelArgument(Description= "[Replacement method for non-bool elements. \"Replace\", \"Optional\" (= replace with None) or \"Filter\". Default is \"Replace\".]")>] replaceMethod: obj)
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
        let rplval = { def with none = none; empty = empty }

        // result
        let o1D = In.Cast.to1D rowwise range  // FIXME - should not use to1D but another In.D1.x function
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = In.D1.Bool.filter o1D
                 a1D |> Out.D1.Prm.out<bool> rplval
        | "O" -> let a1D = In.D1.Bool.Opt.def None o1D
                 a1D |> Out.D1.Prm.out<bool option> rplval
        | _   -> let a1D = In.D1.Bool.def defVal o1D 
                 a1D |> Out.D1.Prm.out<bool> rplval

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
        let rowwise = In.D0.Bool.def false rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let none = In.D0.Stg.def "<none>" noneValue
        let empty = In.D0.Stg.def "<empty>" emptyValue
        let defVal = In.D0.Stg.def "-foo-" defaultValue
        let rplval = { def with none = none; empty = empty }

        // result
        let o1D = In.Cast.to1D rowwise range  // FIXME - should not use to1D but another In.D1.x function
        // the type annotations are NOT necessary (but are used here for readability).
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = In.D1.Stg.filter o1D
                 a1D |> Out.D1.Prm.out<string> rplval
        | "O" -> let a1D = In.D1.Stg.Opt.def None o1D
                 a1D |> Out.D1.Prm.out<string option> rplval
        | _   -> let a1D = In.D1.Stg.def defVal o1D 
                 a1D |> Out.D1.Prm.out<string> rplval

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
        let rowwise = In.D0.Bool.def false rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let none = In.D0.Stg.def "<none>" noneValue
        let empty = In.D0.Stg.def "<empty>" emptyValue
        let defVal = In.D0.Dbl.def 0.0 defaultValue
        let rplval = { def with none = none; empty = empty }

        // result
        let o1D = In.Cast.to1D rowwise range  // FIXME - should not use to1D but another In.D1.x function
        // the type annotations are NOT necessary (but are used here for readability).
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = In.D1.Dbl.filter o1D
                 a1D |> Out.D1.Prm.out<double> rplval
        | "O" -> let a1D = In.D1.Dbl.Opt.def None o1D
                 a1D |> Out.D1.Prm.out<double option> rplval
        | _   -> let a1D = In.D1.Dbl.def defVal o1D 
                 a1D |> Out.D1.Prm.out<double> rplval

    [<ExcelFunction(Category="XL", Description="Cast an xl-range to an array of doubles (including NaNs).")>]
    let cast1d_dblNan
        ([<ExcelArgument(Description= "Range.")>] range: obj)
        ([<ExcelArgument(Description= "[Replacement method for non-double elements. \"Replace\", \"Optional\" (= replace with None) or \"Filter\". Default is \"Replace\".]")>] replaceMethod: obj)
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
        let rplval = { def with none = none; empty = empty }

        // result
        let o1D = In.Cast.to1D rowwise range  // FIXME - should not use to1D but another In.D1.x function
        // the type annotations are NOT necessary (but are used here for readability).
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = In.D1.Nan.filter xlkinds o1D
                 //a1D |> Out.range1D<double> rplval
                 a1D |> Out.D1.Prm.out<double> rplval
        | "O" -> let a1D = In.D1.Nan.Opt.def xlkinds None o1D 
                 //a1D |> Out.range1DOpt<double> rplval
                 a1D |> Out.D1.Prm.out<double option> rplval
        | _   -> let a1D = In.D1.Nan.def xlkinds defVal o1D 
                 //a1D |> Out.range1D<double> rplval
                 a1D |> Out.D1.Prm.out<double> rplval

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
        let rowwise = In.D0.Bool.def false rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let none = In.D0.Stg.def "<none>" noneValue
        let empty = In.D0.Stg.def "<empty>" emptyValue
        let defVal = In.D0.Intg.def 0 defaultValue
        let rplval = { def with none = none; empty = empty }

        // result
        let o1D = In.Cast.to1D rowwise range  // FIXME - should not use to1D but another In.D1.x function
        // the type annotations are NOT necessary (but are used here for readability).
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = In.D1.Intg.filter o1D
                 a1D |> Out.D1.Prm.out<int> rplval
        | "O" -> let a1D = In.D1.Intg.Opt.def None o1D
                 a1D |> Out.D1.Prm.out<int option> rplval
        | _   -> let a1D = In.D1.Intg.def defVal o1D 
                 a1D |> Out.D1.Prm.out<int> rplval
        
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
        let rowwise = In.D0.Bool.def false rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let none = In.D0.Stg.def "<none>" noneValue
        let empty = In.D0.Stg.def "<empty>" emptyValue
        let defVal = In.D0.Dte.def (DateTime(2000,1,1)) defaultValue
        let rplval = { def with none = none; empty = empty }

        // result
        let o1D = In.Cast.to1D rowwise range  // FIXME - should not use to1D but another In.D1.x function
        // the type annotations are NOT necessary (but are used here for readability).
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let a1D = In.D1.Dte.filter o1D
                 a1D |> Out.D1.Prm.out<DateTime> rplval
        | "O" -> let a1D = In.D1.Dte.Opt.def None o1D
                 a1D |> Out.D1.Prm.out<DateTime option> rplval
        | _   -> let a1D = In.D1.Dte.def defVal o1D 
                 a1D |> Out.D1.Prm.out<DateTime> rplval

    [<ExcelFunction(Category="XL", Description="Cast an xl-value to a generic type.")>]
    let cast_gen
        ([<ExcelArgument(Description= "xl-value.")>] xlValue: obj)
        ([<ExcelArgument(Description= "Type label.")>] typeLabel: string)
        ([<ExcelArgument(Description= "[Default Value (only for non-optional types). Must be of the appropriate type. Default \"<default>\" (which will fail for non-string types).]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None indicator. Default is \"<none>\".]")>] noneIndicator: obj)
        : obj  =

        // intermediary stage
        let none = In.D0.Stg.def "<none>" noneIndicator
        let rplval = { def with none = none }
        let defVal = In.D0.Missing.tryO defaultValue

        // result
        let res = Out.D0.Gen.defAllCasesObj rplval defVal typeLabel xlValue
        res

//    [<ExcelFunction(Category="XL", Description="Cast a 1D-slice of an xl-range to a generic type 1D array.")>]
//    let cast1d_gen
//        ([<ExcelArgument(Description= "1D xl-range.")>] range: obj)
//        ([<ExcelArgument(Description= "Type label.")>] typeLabel: string)
//        ([<ExcelArgument(Description= "[Replacement method for non-date elements. \"Replace\", \"Optional\" (= replace with None) or \"Filter\". Default is \"Replace\".]")>] replaceMethod: obj)
//        ([<ExcelArgument(Description= "[Default Value (only for non-optional types). Must be of the appropriate type. Default \"<default>\" (which will fail for non-string types).]")>] defaultValue: obj)
//        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
//        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
//        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
//        : obj[]  =

//        // intermediary stage
//        let rowwise = In.D0.Bool.Opt.def None rowWiseDirection
//        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
//        let none = In.D0.Stg.def "<none>" noneValue
//        let empty = In.D0.Stg.def "<empty>" emptyValue
//        let rplval = { def with none = none; empty = empty }
////        let isoptional = isOptionalType typeLabel
//        let defVal = In.D0.Missing.tryO defaultValue
            
//        match replmethod.ToUpper().Substring(0,1) with
//        | "F" -> Out.D1.Gen.filterObj rowwise rplval typeLabel range
//        | _ -> Out.D1.Gen.defAllCasesObj rowwise rplval defVal typeLabel range

    [<ExcelFunction(Category="XL", Description="Cast a 1D-slice of an xl-range to a generic type 1D array.")>] // FIXME change wording
    let cast1d_gen
        ([<ExcelArgument(Description= "1D xl-range.")>] range: obj)
        ([<ExcelArgument(Description= "Type tag. E.g. bool, date, double, string, obj or #bool, #date,  etc...")>] typeTag: string)
        ([<ExcelArgument(Description= "[Replacement method for non-date elements. \"Replace\", \"Optional\" (= replace with None) or \"Filter\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value (only for non-optional types). Must be of the appropriate type. Default \"<default>\" (which will fail for non-string types).]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[]  =

        // intermediary stage
        let rowwise = In.D0.Bool.Opt.def None rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let none = In.D0.Stg.def "<none>" noneValue
        let empty = In.D0.Stg.def "<empty>" emptyValue
        let rplval = { def with none = none; empty = empty }
        let defVal = In.D0.Missing.tryO defaultValue
        
        // for demo purpose only: take an Excel range input,
        // converts it into a (boxed) typed array, then outputs it back to Excel.
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let xa1D = In.D1.Gen.filter None typeTag range
                 xa1D |> (Out.D1.Unbox.apply rplval (Out.D1.Prm.out rplval))
        | _ -> let xa1D = In.D1.Gen.Any.def rowwise defVal typeTag range
               xa1D |> (Out.D1.Unbox.apply rplval (Out.D1.Prm.out rplval))

    [<ExcelFunction(Category="XL", Description="Cast a 2D xl-range to a generic type 2D array.")>]
    let cast2d_gen
        ([<ExcelArgument(Description= "2D xl-range.")>] range: obj)
        ([<ExcelArgument(Description= "Type label.")>] typeLabel: string)
        ([<ExcelArgument(Description= "[Replacement method for non-date elements. \"Replace\", \"Optional\" (= replace with None) or \"Filter\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value (only for non-optional types). Must be of the appropriate type. Default \"<default>\" (which will fail for non-string types).]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "[None Value. Default is \"<none>\".]")>] noneValue: obj)
        ([<ExcelArgument(Description= "[Empty array value. Default is \"<empty>\".]")>] emptyValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj[,]  =

        // intermediary stage
        let rowwise = In.D0.Bool.Opt.def None rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let none = In.D0.Stg.def "<none>" noneValue
        let empty = In.D0.Stg.def "<empty>" emptyValue
        let rplval = { def with none = none; empty = empty }
        let defVal = In.D0.Missing.tryO defaultValue
            
        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> Out.D2.Gen.filterObj rowwise rplval typeLabel range
        | _ -> Out.D2.Gen.defAllCasesObj rplval defVal typeLabel range


module A1D_XL =
    open type Registry
    open Registry
    open Excel
    open API
    open type Variant
    open type ReplaceValues

    // open API.In.D0

    [<ExcelFunction(Category="XL", Description="Cast a 1D-slice of an xl-range to a generic type array.")>]
    let a1_ofRng
        ([<ExcelArgument(Description= "1D xl-range.")>] range: obj)
        ([<ExcelArgument(Description= "Type label.")>] typeLabel: string)
        ([<ExcelArgument(Description= "[Replacement method for non-date elements. \"Replace\", \"Optional\" (= replace with None) or \"Filter\". Default is \"Replace\".]")>] replaceMethod: obj)
        ([<ExcelArgument(Description= "[Default Value (only for non-optional types). Must be of the appropriate type. Default \"<default>\" (will fail for non-string types).]")>] defaultValue: obj)
        ([<ExcelArgument(Description= "Row wise direction. For 2D ranges only.")>] rowWiseDirection: obj)
        : obj  =

        // intermediary stage
        let rowwise = In.D0.Bool.Opt.def None rowWiseDirection
        let replmethod = In.D0.Stg.def "REPLACE" replaceMethod
        let defVal = In.D0.Missing.tryO defaultValue
        
        // caller cell's reference ID
        let rfid = MRegistry.refID

        match replmethod.ToUpper().Substring(0,1) with
        | "F" -> let ooo = (In.D1.Gen.filter rowwise typeLabel range)
                 let res = ooo |> MRegistry.register rfid 
                 res |> box
        | _ -> let res = (In.D1.Gen.Any.def rowwise defVal typeLabel range)  |> MRegistry.register rfid
               res |> box

    [<ExcelFunction(Category="Array1D", Description="Creates an array of an R-object object.")>]
    let a1_toRng
        ([<ExcelArgument(Description= "1D array R-object.")>] rgA1D: string)
        ([<ExcelArgument(Description= "[None indicator. Default is \"<none>\".]")>] noneIndicator: obj)
        ([<ExcelArgument(Description= "[Empty array indicator. Default is \"<empty>\".]")>] emptyIndicator: obj)
        // TODO: add nan indicator?
        : obj[] = 

        // intermediary stage
        let none = In.D0.Stg.def "<none>" noneIndicator
        let empty = In.D0.Stg.def "<empty>" emptyIndicator
        let rplval = { def with none = none; empty = empty }

        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        match A1D.Reg.Out.out rgA1D rfid rplval with
        | None -> [| box ExcelError.ExcelErrorNA |]
        | Some o1D -> o1D

    [<ExcelFunction(Category="Array1D", Description="Returns the size of an R-object array.")>]
    let a1_count
        ([<ExcelArgument(Description= "1D array R-object.")>] rgA1D: string) 
        ([<ExcelArgument(Description= "[None indicator. Default is \"<none>\".]")>] noneIndicator: obj)
        ([<ExcelArgument(Description= "[Empty array indicator. Default is \"<empty>\".]")>] emptyIndicator: obj)
        : obj = 

        // intermediary stage
        let none = In.D0.Stg.def "<none>" noneIndicator
        let empty = In.D0.Stg.def "<empty>" emptyIndicator
        let rplval = { def with none = none; empty = empty }

        // result
        match A1D.Reg.Out.count rgA1D with
        | None -> box "error"
        | Some o -> o

    [<ExcelFunction(Category="Array1D", Description="Returns the size of an R-object array.")>]
    let a1_elem
        ([<ExcelArgument(Description= "1D array R-object.")>] rgA1D: string) 
        ([<ExcelArgument(Description= "[Index. Default is 0.]")>] index: obj)
        ([<ExcelArgument(Description= "[None indicator. Default is \"<none>\".]")>] noneIndicator: obj)
        : obj = 

        // intermediary stage
        let index = In.D0.Intg.def 0 index

        let none = In.D0.Stg.def "<none>" noneIndicator
        let rplval = { def with none = none }

        // caller cell's reference ID (necessary when the elements are not primitive types)
        let rfid = MRegistry.refID

        // result
        match A1D.Reg.Out.elem index rgA1D with
        | None -> box "error"
        | Some o -> o |> API.Out.D0.Reg.out rfid rplval

    [<ExcelFunction(Category="Array1D", Description="Returns the sub-array of an R-object array.")>]
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
        | None -> box "error"
        | Some o -> o |> MRegistry.register rfid |> box

    [<ExcelFunction(Category="Array1D", Description="Appends several R-object arrays.")>]
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
            | None -> box "error"
            | Some o -> o |> MRegistry.register rfid |> box
        | Some rO3 -> 
            match A1D.Reg.Out.append3 rgA1D1 rgA1D2 rO3 with
            | None -> box "error"
            | Some o -> o |> MRegistry.register rfid |> box

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
            let a0D = In.D0.Gen.Any.def defValue typeLabel xlValue :?> 'A
            a0D |> create0D size

        static member mtrx1D<'A> (defValue: obj option) (typeLabel: string) (size: int) (xlValue: obj) : GenMTRX<'A> =
            let a1D = In.D1.GenFn.def None defValue typeLabel xlValue
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
            let size (xlValue: string) : obj option =
                let methodNm = "size"
                MRegistry.tryExtractGen genType xlValue
                |> Option.map (apply<GenFn> methodNm [||] [||])

            let elem (row: int) (col: int) (xlValue: string) : obj option =
                let methodNm = "elem"
                MRegistry.tryExtractGen genType xlValue
                |> Option.map (apply<GenFn> methodNm [||] [| row; col |])

    // == xl-values functions ==


/// Simple template
module Mtrx =
    open type Registry
    open Registry

    open Useful.Generics
    open API

    // some typed creation functions
    let create0D<'a> (size: int) (a0D: 'a) : MTRX<'a> = Array2D.create size size a0D
    let create1D<'a> (size: int) (a1D: 'a[]) = [| for i in 0 .. (size - 1) -> a1D |] |> array2D

    // == reflection functions ==
    type GenFn =
        static member mtrx0D<'A> (defValue: obj option) (typeLabel: string) (size: int) (xlValue: obj) : MTRX<'A> =
            let a0D = In.D0.GenFn.def defValue typeLabel xlValue
            a0D |> create0D size

        static member mtrx1D<'A> (defValue: obj option) (typeLabel: string) (size: int) (xlValue: obj) : MTRX<'A> =
            let a1D = In.D1.GenFn.def None defValue typeLabel xlValue
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
    open type ReplaceValues
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
        // xxx |> (Out.cellOptTBD ReplaceValues.def) // TODO FIXME
        xxx |> Out.D0.Bxd.Opt.out ReplaceValues.def

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
        ([<ExcelArgument(Description= "[Default Value (only for non-optional types). Must be of the appropriate type. Default \"<default>\" (will fail for non-string types).]")>] defaultValue: obj)
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
        match MRegistry.tryExtract<MTRXD> rgA1D with
        | None -> box "FAILED"
        | Some a2d -> box a2d.[row, col]

    [<ExcelFunction(Category="EXAMPLE", Description="Get Mtrx size.")>]
    let mtrx_size
        ([<ExcelArgument(Description= "Matrix reg. object.")>] rgMtrx: string) 
        : obj = 

        // result
        let xx = Mtrx.mtrxRegOLD rgMtrx
        xx

module Registry_XL =
    open Excel

    open API.In.D0
    open API.Out
    open Registry

    open type Registry
    open ExcelDna.Integration

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



























