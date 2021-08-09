//  Copyright (c) cdr021. All rights reserved.
//  ExcelGlue is licensed under the MIT license. See LICENSE.txt for details.

namespace ExcelGlue

module XlXl =
    open System
    open API
    open ExcelDna.Integration
    open System.Text.RegularExpressions

    // ----------------------
    // -- Basic 1D functions
    // ----------------------

    [<ExcelFunction(Category="XL", Description="Extracts the nth element of an array.")>]
    let x1_elem
        ([<ExcelArgument(Description= "1D range.")>] range1D: obj[]) 
        ([<ExcelArgument(Description= "Element index.")>] elementIndex: double)
        ([<ExcelArgument(Description= "[Default Value. Default is #N/A.]")>] defaultValue: obj)
        : obj = 

        // intermediary stage
        let defvalue = In.D0.Absent.Obj.subst ExcelError.ExcelErrorNA defaultValue

        // result
        match range1D |> Array.tryItem ((int) elementIndex) with
        | None -> defvalue
        | Some o -> o

    [<ExcelFunction(Category="XL", Description="Creates an (obj[]) array from elements.")>]
    let x1_ofElems 
        ([<ExcelArgument(Description= "Element 1.")>] element1: obj)
        ([<ExcelArgument(Description= "Element 2.")>] element2: obj)
        ([<ExcelArgument(Description= "Element 3.")>] element3: obj)
        ([<ExcelArgument(Description= "Element 4.")>] element4: obj)
        ([<ExcelArgument(Description= "Element 5.")>] element5: obj)
        ([<ExcelArgument(Description= "Element 6.")>] element6: obj)
        ([<ExcelArgument(Description= "Element 7.")>] element7: obj)
        ([<ExcelArgument(Description= "Element 8.")>] element8: obj)
        ([<ExcelArgument(Description= "Element 9.")>] element9: obj)
        ([<ExcelArgument(Description= "Element 10.")>] element10: obj)
        ([<ExcelArgument(Description= "Element 11.")>] element11: obj)
        ([<ExcelArgument(Description= "Element 12.")>] element12: obj)
        : obj[] =

        // result
        [| element1; element2; element3; element4; element5; element6; element7; element8; element9; element10; element11; element12 |]
        |> Array.filter (fun x -> match x with | :? ExcelMissing -> false | :? ExcelEmpty -> false | _ -> true)

    [<ExcelFunction(Category="XL", Description="Returns the concatenation of the input arrays.")>]
    let x1_concat
        ([<ExcelArgument(Description= "1D range 1.")>] range1: obj)
        ([<ExcelArgument(Description= "1D range 2.")>] range2: obj) 
        ([<ExcelArgument(Description= "1D range 3.")>] range3: obj) 
        ([<ExcelArgument(Description= "1D range 4.")>] range4: obj) 
        ([<ExcelArgument(Description= "1D range 5.")>] range5: obj) 
        ([<ExcelArgument(Description= "1D range 6.")>] range6: obj) 
        ([<ExcelArgument(Description= "1D range 7.")>] range7: obj)
        ([<ExcelArgument(Description= "1D range 8.")>] range8: obj) 
        ([<ExcelArgument(Description= "1D range 9.")>] range9: obj) 
        ([<ExcelArgument(Description= "1D range 10.")>] range10: obj) 
        ([<ExcelArgument(Description= "1D range 11.")>] range11: obj) 
        ([<ExcelArgument(Description= "1D range 12.")>] range12: obj)
        : obj[] =

        // result
        [| range1; range2; range3; range4; range5; range6; range7; range8; range9; range10; range11; range12 |]
        |> Array.filter (fun x -> match x with | :? ExcelMissing -> false | :? ExcelEmpty -> false | _ -> true)
        |> Array.collect (In.Cast.to1D false)

    let private sub' (xs : 'a[]) (startIndex: int) (subCount: int) : 'a[] =
        if startIndex >= xs.Length then
            [||]
        else
            let start = max 0 startIndex
            let count = (min (xs.Length - startIndex) subCount) |> max 0
            Array.sub xs start count
    
    let private sub (startIndex: int option) (count: int option) (xs: 'a[]) : 'a[] =
        match startIndex, count with
        | Some si, Some cnt -> sub' xs si cnt
        | Some si, None -> sub' xs si (xs.Length - si)
        | None, Some cnt -> sub' xs 0 cnt
        | None, None -> xs

    // same as sub, but pads the sub-array with default value at the end and outputs a constant length array.
    let private view (padding: 'a) (emptyValue: 'a option) (startIndex: int option) (count: int option) (xs: 'a[]) : 'a[] =
        match sub startIndex count xs with
        | [||] -> match emptyValue with | None -> [||] | Some emptyval -> [| emptyval |]
        | subarray -> 
            if subarray.Length < xs.Length then
                Array.append subarray (padding |> Array.replicate (xs.Length - subarray.Length))        
            else
                subarray

    [<ExcelFunction(Category="XL", Description="Returns a contiguous part of the original array.")>]
    let x1_sub
        ([<ExcelArgument(Description= "1D range.")>] range1D: obj[])
        ([<ExcelArgument(Description= "[Start index. Default is 0.]")>] startIndex: obj)
        ([<ExcelArgument(Description= "[Sub-array length. Default is full length.]")>] length: obj) 
        ([<ExcelArgument(Description= "[Empty array indicator. Default is #N/A.]")>] emptyIndicator: obj)
        : obj[] =

        // intermediary stage
        let count = In.D0.Intg.Opt.def None length
        let startIdx = In.D0.Intg.Opt.def None startIndex
        let emptyIndic = In.D0.Absent.def (box ExcelError.ExcelErrorNA) emptyIndicator |> Array.singleton

        // result
        let res = range1D |> sub startIdx count
        if res.Length = 0 then emptyIndic else res

    [<ExcelFunction(Category="XL", Description="Returns an array repeating n times a single element.")>]
    let x1_repeat
        ([<ExcelArgument(Description= "Element to repeat.")>] element: obj)
        ([<ExcelArgument(Description= "[Repeat count. Default is 1.]")>] count: obj)
        : obj[] =

        // intermediary stage
        let count = In.D0.Intg.def 1 count

        // result
        Array.replicate count element

    [<ExcelFunction(Category="XL", Description="Removes duplicates.")>]
    let x1_nub 
        ([<ExcelArgument(Description= "1D range.")>] range1D: obj[])
        : obj[] =

        // result
        range1D |> Array.distinct

    let private scrap (filterValues: obj[]) (excludeValues: bool) (noBlank: bool) (noError: bool) (o1D: obj[]) : obj[] =
        let predicate o = 
            (excludeValues <> (filterValues |> Array.contains o))
            && (if noBlank then not (In.D0.Is.blank o) else true)
            && (if noError then not (In.D0.Is.error o) else true)

        o1D |> Array.filter predicate

    [<ExcelFunction(Category="XL", Description="Filters values, blanks and/or errors out of the array.")>]
    let x1_scrap
        ([<ExcelArgument(Description= "1D range.")>] range1D: obj[])
        ([<ExcelArgument(Description= "[Filter values. Default is none.]")>] filterValues: obj)
        ([<ExcelArgument(Description= "[Exclude filter values (otherwise include them). Default is true.]")>] excludeValues: obj)
        ([<ExcelArgument(Description= "[Remove blank cells. Default is false.]")>] removeBlankCells: obj)
        ([<ExcelArgument(Description= "[Remove errors. Default is false.]")>] removeErrors: obj) 
        : obj[] =

        // intermediary stage
        let filterVals = In.D0.Absent.map<obj[]> [||] (In.Cast.to1D false) filterValues
        let excludeVals = In.D0.Bool.def true excludeValues
        let noBlank = In.D0.Bool.def false removeBlankCells
        let noError = In.D0.Bool.def false removeErrors

        // result
        scrap filterVals excludeVals noBlank noError range1D

    [<ExcelFunction(Category="XL", Description="Filters error and blank values out of the input array (NEB = [N]o [E]rror, no [B]lank).")>]
    let x1_neb 
        ([<ExcelArgument(Description= "1D range.")>] range1D: obj[])
        ([<ExcelArgument(Description= "[Error and blank replacement value. Default is none.]")>] replacementValue: obj)
        ([<ExcelArgument(Description= "[Empty array indicator. Default is #N/A.]")>] emptyIndicator: obj)
        : obj[] =

        // intermediary stage
        let emptyIndic = In.D0.Absent.def (box ExcelError.ExcelErrorNA) emptyIndicator |> Array.singleton

        // result
        let res =
            match In.D0.Absent.Obj.tryO replacementValue with
            | None -> scrap [||] true true true range1D
            | Some replvalue -> range1D |> Array.map (fun o -> if In.D0.Is.blankOrError o then replvalue else o) 

        if res.Length = 0 then
            emptyIndic
        else
            res

    [<ExcelFunction(Category="XL", Description="Returns a sub-array and pads it with extra values.")>]
    let x1_view
        ([<ExcelArgument(Description= "1D range.")>] range1D: obj[])
        ([<ExcelArgument(Description= "[Start index. Default is 0.]")>] startIndex: obj)
        ([<ExcelArgument(Description= "[Sub-array length. Default is full length.]")>] length: obj)
        ([<ExcelArgument(Description= "[Padding value. Default is #N/A.]")>] paddingValue: obj) 
        ([<ExcelArgument(Description= "[Empty array indicator. Default is #N/A.]")>] emptyIndicator: obj)
        ([<ExcelArgument(Description= "[Error and blank replacement value. Default is none.]")>] replacementValue: obj)
        : obj[] =
         
        // intermediary stage
        let count = In.D0.Intg.Opt.def None length
        let startIdx = In.D0.Intg.Opt.def None startIndex
        let emptyIndic = In.D0.Absent.def (box ExcelError.ExcelErrorNA) emptyIndicator
        let padding = In.D0.Absent.def (box ExcelError.ExcelErrorNA) paddingValue
        let xs = if In.D0.Is.absent replacementValue then range1D else x1_neb range1D replacementValue emptyIndicator

        // result
        view padding (Some emptyIndic) startIdx count xs

    // values1D and filter1D should have the same size.
    // returns values1D's elements such as their corresponding filter1D element is either included in or excluded of filterValues.
    let private filter (filterValues: obj[]) (excludeValues: bool) (noBlank: bool) (noError: bool) (filter1D: obj[]) (values1D: obj[]) : obj[] =
        let predicate o = 
            (excludeValues <> (filterValues |> Array.contains o))
            && (if noBlank then not (In.D0.Is.blank o) else true)
            && (if noError then not (In.D0.Is.error o) else true)

        Array.zip values1D filter1D
        |> Array.filter (snd >> predicate)
        |> Array.unzip
        |> fst

    [<ExcelFunction(Category="XL", Description="Filters a 1D array.")>]
    let x1_filter
        ([<ExcelArgument(Description= "Values 1D range.")>] values1D: obj[])
        ([<ExcelArgument(Description= "Filter 1D range.")>] filter1D: obj[])
        ([<ExcelArgument(Description= "[Filter values. Default is none.]")>] filterValues: obj)
        ([<ExcelArgument(Description= "[Exclude filter values (otherwise include them). Default is true.]")>] excludeValues: obj)
        ([<ExcelArgument(Description= "[Remove blank cells. Default is false.]")>] removeBlankCells: obj)
        ([<ExcelArgument(Description= "[Remove errors. Default is false.]")>] removeErrors: obj) 
        ([<ExcelArgument(Description= "[Empty array indicator. Default is #N/A.]")>] emptyIndicator: obj)
        : obj[] =

        // intermediary stage
        let filterVals = In.D0.Absent.map<obj[]> [||] (In.Cast.to1D false) filterValues
        let excludeVals = In.D0.Bool.def true excludeValues
        let noBlank = In.D0.Bool.def false removeBlankCells
        let noError = In.D0.Bool.def false removeErrors
        let emptyIndic = In.D0.Absent.def (box ExcelError.ExcelErrorNA) emptyIndicator

        // result
        let res = filter filterVals excludeVals noBlank noError filter1D values1D
        if res.Length = 0 then [| emptyIndic |] else res

    // ----------------------
    // -- Basic 2D functions
    // ----------------------

    [<ExcelFunction(Category="XL", Description="Returns a 2D contiguous part of the original array.")>]
    let x2_sub
        ([<ExcelArgument(Description= "2D range.")>] range2D: obj[,])
        ([<ExcelArgument(Description= "[Row start index. Default is 0.]")>] rowStartIndex: obj)
        ([<ExcelArgument(Description= "[Column start index. Default is 0.]")>] colStartIndex: obj)
        ([<ExcelArgument(Description= "[Sub-array row count. Default is all rows.]")>] rowCount: obj) 
        ([<ExcelArgument(Description= "[Sub-array column count. Default is all columns.]")>] colCount: obj) 
        ([<ExcelArgument(Description= "[Empty array indicator. Default is #N/A.]")>] emptyIndicator: obj)
        : obj[,] =

        // intermediary stage
        let (len1, len2) = (Array2D.length1 range2D, Array2D.length2 range2D)
        //let rowStartIdx = let rsi = In.D0.Intg.def 0 rowStartIndex in max 0 (min (len1 - 1) rsi)
        //let colStartIdx = let csi = In.D0.Intg.def 0 colStartIndex in max 0 (min (len2 - 1) csi)

        let rowStartIdx = let rsi = In.D0.Intg.def 0 rowStartIndex in max 0 rsi
        let colStartIdx = let csi = In.D0.Intg.def 0 colStartIndex in max 0 csi

        let rowCount = let rc = In.D0.Intg.def len1 rowCount in max 0 (min (len1 - rowStartIdx) rc)
        let colCount = let cc = In.D0.Intg.def len2 colCount in max 0 (min (len2 - colStartIdx) cc)
        let emptyIndic = In.D0.Absent.def (box ExcelError.ExcelErrorNA) emptyIndicator

        // result
        let res = range2D.[rowStartIdx .. (rowStartIdx + rowCount - 1), colStartIdx .. (colStartIdx + colCount - 1)]
        if (res |> Array2D.length1 = 0) || (res |> Array2D.length2 = 0) then (Array2D.create 1 1 emptyIndic) else res

    let private appendV<'a> (o2Dtop: 'a[,]) (o2Dbot: 'a[,]) : 'a[,] option =
        let (len1top, len2top) = (Array2D.length1 o2Dtop, Array2D.length2 o2Dtop)
        let (len1bot, len2bot) = (Array2D.length1 o2Dbot, Array2D.length2 o2Dbot)

        if len2top <> len2bot then
            None
        else
            Array.append
                [| for i in o2Dtop.GetLowerBound(0) .. o2Dtop.GetUpperBound(0) 
                        -> o2Dtop.[i,*]
                |]

                [| for i in o2Dbot.GetLowerBound(0) .. o2Dbot.GetUpperBound(0) 
                        -> o2Dbot.[i,*]
                |]

            |> array2D
            |> Some

    let private appendH<'a> (o2Dleft: 'a[,]) (o2Dright: 'a[,]) : 'a[,] option =
        let (len1left, len2left) = (Array2D.length1 o2Dleft, Array2D.length2 o2Dleft)
        let (len1right, len2right) = (Array2D.length1 o2Dright, Array2D.length2 o2Dright)

        if len1left <> len1right then
            None
        else
            Array2D.initBased (Array2D.base1 o2Dleft) (Array2D.base2 o2Dleft) len1left (len2left + len2right)
                ( fun i j ->
                    if j <= o2Dleft.GetUpperBound(1) then
                        o2Dleft.[i,j]
                    else
                        let iright = i - o2Dleft.GetLowerBound(0) + o2Dright.GetLowerBound(0)
                        let jright = j - o2Dleft.GetUpperBound(1) - 1 + o2Dright.GetLowerBound(1)
                        o2Dright.[iright,jright]
                )
            |> Some

    [<ExcelFunction(Category="XL", Description="Appends 2 2D arrays either horizontally (column-wise) or vertically (row-wise).")>]
    let x2_append
        ([<ExcelArgument(Description= "2D range 1.")>] range2D1: obj[,])
        ([<ExcelArgument(Description= "2D range 2.")>] range2D2: obj[,])
        ([<ExcelArgument(Description= "[Column-wise. Default is false]")>] colWise: obj)
        : obj[,] =

        // result
        if In.D0.Bool.def false colWise then
            appendH range2D1 range2D2
        else
            appendV range2D1 range2D2
        |> Option.defaultValue (Array2D.create 1 1 (box ExcelError.ExcelErrorNA))

    [<ExcelFunction(Category="XL", Description="Returns the dimensions of a 2D array.")>]
    let x2_size
        ([<ExcelArgument(Description= "2D range.")>] range2D: obj[,])
        ([<ExcelArgument(Description= "[Dimension 1 (row count) or 2 (column count) or 0 for both. Default is 1.]")>] dimension: obj)
        : obj =

        // intermediary stage
        let (len1, len2) = (Array2D.length1 range2D, Array2D.length2 range2D)
        let dim = In.D0.Intg.def 1 dimension

        // result
        if dim = 0 then
            [| len1; len2 |] |> Array.map double |> box
        elif dim = 2 then
            box len2
        else
            box len1


    // --------------------
    // -- Equality functions
    // --------------------

    [<ExcelFunction(Category="XL", Description="Equality of 2 xl-values (\'variants\'), including Excel errors.")>]
    let x0_eq
        ([<ExcelArgument(Description= "Value 1.")>] value1: obj) 
        ([<ExcelArgument(Description= "Value 2.")>] value2: obj) 
        : obj  =

        // result
        (value1 = value2) |> box

    [<ExcelFunction(Category="XL", Description="Numbers equality, within an approximation error.")>]
    let x0_eqNum
        ([<ExcelArgument(Description= "Value 1.")>] value1: obj) 
        ([<ExcelArgument(Description= "Value 2.")>] value2: obj) 
        ([<ExcelArgument(Description= "[Approximation error. Default is none.]")>] aproxError: obj) 
        ([<ExcelArgument(Description= "[Compare numbers only. If true returns #N/A for non-numeric input values. Default is false.]")>] onlyNumbers: obj) 
        : obj  =

        // intermediary stage
        let failedValue = box ExcelError.ExcelErrorNA

        let error = In.D0.Dbl.Opt.def None aproxError
        let onlyNums = In.D0.Bool.def false onlyNumbers

        // result
        match onlyNums, value1, value2 with
        | _, (:? double as val1), (:? double as val2) -> 
            match error with
            | None -> (val1 = val2) |> box
            | Some err -> (Math.Abs (val1 - val2)) < Math.Abs (err) |> box
        | false, _, _ -> (value1 = value2) |> box
        | _ -> failedValue

    // --------------------
    // -- Default functions
    // --------------------

    [<ExcelFunction(Category="XL", Description="Replaces errors with a default value.")>]
    let x0_def
        ([<ExcelArgument(Description= "Value.")>] value: obj) 
        ([<ExcelArgument(Description= "Error substitute.")>] defaultValue: obj) 
        : obj  =

        // result
        match value with
        | :? ExcelError -> defaultValue
        | :? ExcelEmpty -> defaultValue
        | _ -> value

    [<ExcelFunction(Category="XL", Description="Replaces non-numeric values with a default value.")>]
    let x0_defNum
        ([<ExcelArgument(Description= "Value.")>] value: obj) 
        ([<ExcelArgument(Description= "Non-numeric substitute.")>] defaultValue: obj) 
        : obj  =

        // result
        match value with
        | :? double -> value
        | _ -> defaultValue

    [<ExcelFunction(Category="XL", Description="Replaces non-text values with a default value.")>]
    let x0_defTxt
        ([<ExcelArgument(Description= "Value.")>] value: obj) 
        ([<ExcelArgument(Description= "Non-text substitute.")>] defaultValue: obj) 
        : obj  =

        // result
        match value with
        | :? string -> value
        | _ -> defaultValue

    [<ExcelFunction(Category="XL", Description="Returns the Excel type of the input.")>]
    let x1_testType
        ([<ExcelArgument(Description= "Range to test (0D, 1D or 2D).")>] range: obj)
        : obj =

        // result
        match range with
        | :? (obj[,]) -> "2D"
        | :? bool -> "Bool"
        | :? double -> "Double"
        | :? string -> "String"
        | :? ExcelError -> "Error"
        | :? ExcelMissing -> "Missing"
        | :? ExcelEmpty -> "Empty"
        | _ -> "Unknown"
        |> box

    // --------------------
    // -- Numeric functions
    // --------------------

    [<ExcelFunction(Category="XL", Description="Returns [| start..skip..finish |] (F# notation).")>]
    let x1_series
        ([<ExcelArgument(Description= "Series start.")>] first: double)
        ([<ExcelArgument(Description= "Series end.")>] last: double)
        ([<ExcelArgument(Description= "[Series increment. Default is 1.]")>] skip: obj) 
        : obj[] =

        // intermediary stage
        let skip = In.D0.Intg.def 1 skip

        // result
        [| ((int) first) .. skip .. ((int) last) |] |> Array.map box

    [<ExcelFunction(Category="XL", Description="Sum numeric values (non-numeric values are replaced with 0.0).")>]
    let x2_sum
        ([<ExcelArgument(Description= "Range.")>] range: obj[,]) 
        : obj =
        
        // intermediary stage
        let a2D = range |> Array2D.map (In.D0.Dbl.def 0.0)
        let len1 = Array2D.length1 a2D

        // result
        let res = [| for i in 0 .. (len1 - 1) -> a2D.[i,*] |> Array.sum |]  |> Array.sum
        box res

    [<ExcelFunction(Category="XL", Description="Sum absolute values (non-numeric values are replaced with 0.0).")>]
    let x2_sumAbs
        ([<ExcelArgument(Description= "Range.")>] range: obj[,]) 
        : obj =
        
        // intermediary stage
        let a2D = range |> Array2D.map (In.D0.Dbl.def 0.0 >> abs)
        let len1 = Array2D.length1 a2D

        // result
        let res = [| for i in 0 .. (len1 - 1) -> a2D.[i,*] |> Array.sum |]  |> Array.sum
        box res

    [<ExcelFunction(Category="XL", Description="Minimum of numeric values (non-numeric values are ignored).")>]
    let x2_min
        ([<ExcelArgument(Description= "Range.")>] range: obj[,]) 
        : obj =
         
        // intermediary stage
        let a2D = range |> Array2D.map (In.D0.Dbl.def Double.MaxValue)
        let len1 = Array2D.length1 a2D

        // result
        let res = [| for i in 0 .. (len1 - 1) -> a2D.[i,*] |> Array.min |]  |> Array.min
        box res

    [<ExcelFunction(Category="XL", Description="Minimum of absolute values (non-numeric values are ignored).")>]
    let x2_minAbs
        ([<ExcelArgument(Description= "Range.")>] range: obj[,]) 
        : obj =
         
        // intermediary stage
        let a2D = range |> Array2D.map (In.D0.Dbl.def Double.MaxValue >> abs)
        let len1 = Array2D.length1 a2D

        // result
        let res = [| for i in 0 .. (len1 - 1) -> a2D.[i,*] |> Array.min |]  |> Array.min
        box res

    [<ExcelFunction(Category="XL", Description="Maximum of numeric values (non-numeric values are ignored).")>]
    let x2_max
        ([<ExcelArgument(Description= "Range.")>] range: obj[,]) 
        : obj =
         
        // intermediary stage
        let a2D = range |> Array2D.map (In.D0.Dbl.def Double.MinValue)
        let len1 = Array2D.length1 a2D

        // result
        let res = [| for i in 0 .. (len1 - 1) -> a2D.[i,*] |> Array.max |]  |> Array.max
        box res

    [<ExcelFunction(Category="XL", Description="Maximum of absolute values (non-numeric values are ignored).")>]
    let x2_maxAbs
        ([<ExcelArgument(Description= "Range.")>] range: obj[,]) 
        : obj =
         
        // intermediary stage
        let a2D = 
            range |> Array2D.map (In.D0.Dbl.def 0.0 >> abs)
        let len1 = Array2D.length1 a2D

        // result
        let res = [| for i in 0 .. (len1 - 1) -> a2D.[i,*] |> Array.max |]  |> Array.max
        box res

    [<ExcelFunction(Category="XL", Description="Sum-product of numeric values in 1D-ranges (non-numeric values are replaced with 0.0).")>]
    let x1_sumprod
        ([<ExcelArgument(Description= "1D range 1.")>] range1: obj)
        ([<ExcelArgument(Description= "1D range 2.")>] range2: obj) 
        ([<ExcelArgument(Description= "1D range 3.")>] range3: obj) 
        ([<ExcelArgument(Description= "1D range 4.")>] range4: obj) 
        ([<ExcelArgument(Description= "1D range 5.")>] range5: obj) 
        ([<ExcelArgument(Description= "1D range 6.")>] range6: obj) 
        : obj =
        
        // result
        let a1Ds = 
            [| range1; range2; range3; range4; range5; range6  |] 
            |> Array.choose In.D0.Absent.Obj.tryO
            |> Array.map (In.Cast.to1D false)
            |> Array.map (Array.map (In.D0.Dbl.def 0.0))
            |> Array.reduce (fun a1D1 a1D2 -> Array.map2 (*) a1D1 a1D2)

        let res = a1Ds |> Array.sum
        box res

    [<ExcelFunction(Category="XL", Description="Sum-product of absolute numeric values in 1D-ranges (non-numeric values are replaced with 0.0).")>]
    let x1_sumprodAbs
        ([<ExcelArgument(Description= "1D range 1.")>] range1: obj)
        ([<ExcelArgument(Description= "1D range 2.")>] range2: obj) 
        ([<ExcelArgument(Description= "1D range 3.")>] range3: obj) 
        ([<ExcelArgument(Description= "1D range 4.")>] range4: obj) 
        ([<ExcelArgument(Description= "1D range 5.")>] range5: obj) 
        ([<ExcelArgument(Description= "1D range 6.")>] range6: obj) 
        : obj =
        
        // result
        let a1Ds = 
            [| range1; range2; range3; range4; range5; range6 |] 
            |> Array.choose In.D0.Absent.Obj.tryO
            |> Array.map (In.Cast.to1D false)
            |> Array.map (Array.map (In.D0.Dbl.def 0.0))
            |> Array.reduce (fun a1D1 a1D2 -> Array.map2 (fun x1 x2 -> x1 * x2 |> abs) a1D1 a1D2)

        let res = a1Ds |> Array.sum
        box res

    // --------------------
    // -- String functions
    // --------------------

    [<ExcelFunction(Category="XL", Description="Returns Regex matches.")>]
    let x0_rgx
        ([<ExcelArgument(Description= "Regex pattern.")>] regexPattern: string)
        ([<ExcelArgument(Description= "Input string.")>] inputString: string)
        ([<ExcelArgument(Description= "[Group index. Default is 1.]")>] groupIndex: obj) 
        : obj =
        
        // intermediary stage
        let failedValue = box ExcelError.ExcelErrorNA

        let inString = In.D0.Stg.def "" inputString
        let groupIdx = In.D0.Intg.def 1 groupIndex
        let m = Regex.Match(inString, regexPattern)

        // result
        if m.Success then 
            box m.Groups.[groupIdx].Value
        else    
            failedValue

    [<ExcelFunction(Category="XL", Description="Breaks a string into pieces separated by the separator.")>]
    let x1_split
        ([<ExcelArgument(Description= "Input string.")>] inputString: string)
        ([<ExcelArgument(Description= "[Separator string. Default is \",\".]")>] separator: obj)
        ([<ExcelArgument(Description= "[Remove spaces. Default is false.]")>] removeSpaces: obj)
        ([<ExcelArgument(Description= "[Exclude blank elements. Default is false.]")>] excludeBlanks: obj)
        : obj[] =

        // intermediary stage
        let sep   = In.D0.Stg.def "," separator
        let inString = if In.D0.Bool.def false removeSpaces then inputString.Replace(" ", "") else inputString
        let exclBlanks = In.D0.Bool.def false excludeBlanks

        // result
        let split = inString.Split([| sep |], StringSplitOptions.None)
        if exclBlanks then split |> Array.filter ((<>) "") else split
        |> Array.map box

    [<ExcelFunction(Category="XL", Description="Joins an array of strings into a single string.")>]
    let x1_join
        ([<ExcelArgument(Description= "Array of strings.")>] stringArray: obj[])
        ([<ExcelArgument(Description= "[Separator string. Default is \",\".]")>] separator: obj)
        ([<ExcelArgument(Description= "[Remove blank cells. Default is false.]")>] removeBlankCells: obj)
        ([<ExcelArgument(Description= "[Remove errors. Default is false.]")>] removeErrors: obj) 
        : obj =

        // intermediary stage
        let sep = In.D0.Stg.def "," separator
        let noBlank = In.D0.Bool.def false removeBlankCells
        let noError = In.D0.Bool.def false removeErrors
        let strings  = 
            scrap [||] true noBlank noError stringArray
            |> Array.map (string) 

        // result
        String.Join(sep, strings) |> box

    // --------------------
    // --Miscellaneous functions
    // --------------------

    [<ExcelFunction(Category="Relation", Description="Returns a hash value of an Excel value or of a R-object.")>]
    let x0_hash
        ([<ExcelArgument(Description= "Value or R-obj.")>] value: obj)
        : obj = 

        // result
        let ovalue = Registry.MRegistry.tryExtractO value |> Option.defaultValue value
        let res = hash ovalue
        box res

module XlIO =
    open System
    open API
    open ExcelDna.Integration

    // ----------------------
    // -- Basic file functions
    // ----------------------

    [<ExcelFunction(Category="IO", Description="Saves lines of text to disk.")>]
    let io_WriteLines
        ([<ExcelArgument(Description= "File path.")>] filePath: string)
        ([<ExcelArgument(Description= "Lines of text.")>] lines: obj)
        : obj =

        // intermediary stage
        let lines = In.D1.OStg.filter lines

        // result
        System.IO.File.WriteAllLines(filePath, lines)
        let now = DateTime.Now
        box now

    [<ExcelFunction(Category="IO", Description="Loads lines of text to disk.")>]
    let io_ReadLines
        ([<ExcelArgument(Description= "File path.")>] filePath: string)
        : obj =

        // caller cell's reference ID
        let rfid = Registry.MRegistry.refID

        // result
        let lines = System.IO.File.ReadAllLines(filePath)
        lines |> Registry.MRegistry.registerBxd rfid

    [<ExcelFunction(Category="IO", Description="Returns a file's last write time.")>]
    let io_fLastMod
        ([<ExcelArgument(Description= "File path.")>] filePath: string)
        : obj = 

        // intermediary stage
        let fileInfo = System.IO.FileInfo(filePath)
        let lastMod = fileInfo.LastWriteTime

        // result
        box lastMod

    [<ExcelFunction(Category="IO", Description="Returns a file's size.")>]
    let io_fSize
        ([<ExcelArgument(Description= "File path.")>] filePath: string)
        : obj = 

        // intermediary stage
        let fileInfo = System.IO.FileInfo(filePath)
        let size = fileInfo.Length

        // result
        box size

    [<ExcelFunction(Category="IO", Description="Returns true if a file exists.")>]
    let io_fExists
        ([<ExcelArgument(Description= "File path.")>] filePath: string)
        : obj = 

        // intermediary stage
        let fileInfo = System.IO.FileInfo(filePath)
        let exists = fileInfo.Exists

        // result
        box exists

    [<ExcelFunction(Category="IO", Description="Copy a file to a specified location.")>]
    let io_fCopyTo
        ([<ExcelArgument(Description= "File path.")>] filePath: string)
        ([<ExcelArgument(Description= "Destination file path.")>] destinationFilePath: string)
        ([<ExcelArgument(Description= "[Overwrite. Default is false.]")>] overwrite: obj)
        : obj = 

        // intermediary stage
        let fileInfo = System.IO.FileInfo(filePath)
        let overwrite = In.D0.Bool.def false overwrite
        let cpyFileInfo = fileInfo.CopyTo(destinationFilePath, overwrite)
        let now = DateTime.Now

        // result
        box now

    [<ExcelFunction(Category="IO", Description="List files which name matches search pattern.")>]
    let io_enumFiles
        ([<ExcelArgument(Description= "File path.")>] filePath: string)
        ([<ExcelArgument(Description= "[Search pattern. E.g. \"*.text\", \"file?.csv\"... Default is none.]")>] searchPattern: obj)
        ([<ExcelArgument(Description= "[Sort on. (A)lphabetical or (D)ate. Default is None.")>] sortOn: obj)
        ([<ExcelArgument(Description= "[Descending. Default is false.")>] descending: obj)
        ([<ExcelArgument(Description= "[Full name. Default is false.")>] fullName: obj)
        : obj[] =

        // intermediary stage
        let searchPattern = In.D0.Stg.def "" searchPattern        
        let sortOn = 
            match In.D0.Stg.Opt.def None sortOn with 
            | None -> None 
            | Some s -> 
                if s.ToUpper().StartsWith "A" then Some "A"
                elif s.ToUpper().StartsWith "D" then Some "D"
                else None
        let descending = In.D0.Bool.def false descending
        let fullName = In.D0.Bool.def false fullName

        let fileInfo = System.IO.FileInfo(filePath)
        let dirInfo = fileInfo.Directory
        let filesInfo = dirInfo.EnumerateFiles(searchPattern) |> Seq.toArray
        let sorted = 
            match sortOn with       
            | None -> filesInfo
            | Some x -> 
                if x = "A" then
                    let sortFn = if descending then Array.sortByDescending else Array.sortBy
                    filesInfo |> sortFn (fun finfo -> finfo.Name)
                else
                    let sortFn = if descending then Array.sortByDescending else Array.sortBy
                    filesInfo |> sortFn (fun finfo -> finfo.LastWriteTime)
        sorted 
        |> Array.map (fun finfo -> (if fullName then finfo.FullName else finfo.Name) |> box)
        
    [<ExcelFunction(Category="IO", Description="Copy a string to the clipboard.")>]
    let io_toClip
        ([<ExcelArgument(Description= "Text.")>] text: string)
        : obj = 

        // result
        System.Windows.Forms.Clipboard.SetText(text)
        let now = DateTime.Now
        box now
















