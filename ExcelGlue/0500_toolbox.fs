//  Copyright (c) cdr021. All rights reserved.
//  ExcelGlue is licensed under the MIT license. See LICENSE.txt for details.

namespace ExcelGlue

module Toolbox =
    open System
    open System.IO

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
          fun (a: obj) -> if a = null then Some (box "None detected here") else None

        let unwrap (o: obj) : obj option =   
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
        /// Invoke 'gen's methodName member, over methodTypes generic types for methodArguments.
        let invoke<'gen> (methodName: string) (methodTypes: Type[]) (methodArguments: obj[]) : obj =
            let meth = typeof<'gen>.GetMethod(methodName)
            let genm = meth.MakeGenericMethod(methodTypes)
            let res  = genm.Invoke(null, methodArguments)
            res

        /// Invoke 'gen's methodName member, with:
        ///    - generic types : fst genTypeRObj
        ///    - arguments : otherArgumentsLeft; fst genTypeRObj; otherArgumentsRight
        let apply<'gen> (methodName: string) (otherArgumentsLeft: obj[]) (otherArgumentsRight: obj[]) (genTypeRObj: Type[]*obj) : obj =
            let (gentys, robj) = genTypeRObj
            invoke<'gen> methodName gentys ([| otherArgumentsLeft; [| robj |];  otherArgumentsRight |] |> Array.concat)

        let apply2TBD<'gen,'a> (methodName: string) (otherArgumentsLeft: obj[]) (otherArgumentsRight: obj[]) (genTypeRObj: Type[]*obj) : obj =
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

    module Array =
        // -----------------------------
        // -- Basic functions
        // -----------------------------
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

        /// Transposes array of arrays.
        /// Jagged arrays are not allowed (No check is done).
        let transpose<'a> (xs: 'a[][]) : 'a[][] =
            if xs.Length = 0 then
                [||]
            else
                Array.init xs.[0].Length (fun j -> Array.init xs.Length (fun i -> xs.[i].[j]))

        /// Returns a 'chunk' of the input array, xs.
        /// Returns [ emptyValue ] for empty array.
        /// If there are not enough elements to fill a count-length chunks, pads its end with defValues. 
        let view (defValue: 'a) (emptyValue: 'a option) (startIndex: int option) (count: int option) (xs : 'a[]) : 'a[] =
            let subArray = 
                match emptyValue, sub startIndex count xs with
                | Some emptyval, [||] -> [| emptyval |]
                | _, resxs -> resxs
            match count with
            | None -> subArray
            | Some cnt ->
                // pads subArray with defValues 
                if subArray.Length < cnt then
                    Array.append subArray (defValue |> Array.replicate (cnt - subArray.Length))
                else
                    subArray


        // -----------------------------
        // -- Zip functions
        // -----------------------------
        let zip (xs1: 'a1[]) (xs2: 'a2[]) : ('a1*'a2)[] =
            if xs1.Length = 0 || xs2.Length = 0 then 
                [||]
            elif xs2.Length = xs1.Length then
                Array.zip xs1 xs2
            elif xs2.Length > xs1.Length then 
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

        let zipWith (fn: 'a1 -> 'a2 -> 'b) (xs1: 'a1[]) (xs2: 'a2[]) : 'b[] =
            zip xs1 xs2 |> Array.map (fun (x1, x2) -> fn x1 x2)

        let zipWith3 (fn: 'a1 -> 'a2 -> 'a3 -> 'b) (xs1: 'a1[]) (xs2: 'a2[]) (xs3: 'a3[]) : 'b[] =
            zip3 xs1 xs2 xs3 |> Array.map (fun (x1, x2, x3) -> fn x1 x2 x3)

        let zipWith4 (fn: 'a1 -> 'a2 -> 'a3 -> 'a4 -> 'b) (xs1: 'a1[]) (xs2: 'a2[]) (xs3: 'a3[]) (xs4: 'a4[]) : 'b[] =
            zip4 xs1 xs2 xs3 xs4 |> Array.map (fun (x1, x2, x3, x4) -> fn x1 x2 x3 x4)

    module Array2D =
        // -----------------------------
        // -- Basic functions
        // -----------------------------

        /// Empty 2D array.
        let empty2D<'a> : 'a[,] = [||] |> array2D

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

        /// Creates a 2D array from an array of 1D arrays.
        /// All the inner arrays must have the same length.
        let concat2D (rowWise: bool) (a1Ds: ('a[]) []): 'a[,] =
            if a1Ds.Length = 0 then
                empty2D<'a>
            elif a1Ds.[0].Length = 0 then
                empty2D<'a>
            elif rowWise then
                a1Ds |> array2D
            else
                Array2D.init a1Ds.[0].Length a1Ds.Length (fun i j -> a1Ds.[j].[i])
        
        /// Returns a 2D array, horizontal (rowWise = true) or vertical (rowWise = false),
        /// from a 1D array.
        let of1D (rowWise: bool) (a1D: 'a[]) : 'a[,] = a1D |> Array.singleton |> concat2D rowWise
        
        let appendV<'a> (a2Dtop: 'a[,]) (a2Dbot: 'a[,]) : 'a[,] option =
            let len2top = Array2D.length2 a2Dtop
            let len2bot = Array2D.length2 a2Dbot

            if len2top <> len2bot then
                None
            else
                Array.append
                    [| for i in a2Dtop.GetLowerBound(0) .. a2Dtop.GetUpperBound(0) 
                            -> a2Dtop.[i,*]
                    |]

                    [| for i in a2Dbot.GetLowerBound(0) .. a2Dbot.GetUpperBound(0) 
                            -> a2Dbot.[i,*]
                    |]

                |> array2D
                |> Some

    /// CSV files reading functions. 
    module CSV =
        open Microsoft.VisualBasic.FileIO // Reference Microsoft.VisualBasic

        let private readLine (trim: bool) (sep: string) (txtreader: TextReader) : string[] =
            let line = txtreader.ReadLine()
            let words = let ws = line.Split([| sep |], StringSplitOptions.None)
                        if not trim then ws else ws |> Array.map (fun s -> s.Trim())
            words
    
        /// Reads CSV file.
        let readLines (trim: bool) (sep: string) (fpath: string) : seq<string[]> =
            let lines =
                seq { use txtreader = File.OpenText(fpath)
                      while not txtreader.EndOfStream do
                          yield readLine trim sep txtreader 
                    }
            lines

        /// Reads CSV file.
        /// Use over CSV.readLines function, when enclosingQuotes are present in the file.
        let readLinesVB (enclosingQuotes: bool) (trim: bool) (sep: string) (fpath: string) : seq<string[]> =
            let lines = 
                seq { use parser = new TextFieldParser(fpath)
                      parser.TextFieldType <- FieldType.Delimited
                      parser.SetDelimiters([| sep |])
                      parser.TrimWhiteSpace <- trim
                      parser.HasFieldsEnclosedInQuotes <- enclosingQuotes
                      while not parser.EndOfData do 
                          yield parser.ReadFields()
                    }
            lines
