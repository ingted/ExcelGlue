namespace ExcelGlue

module XlXl =
    open System
    open API
    open ExcelDna.Integration
    open System.Text.RegularExpressions

    // --------------------
    // -- Test functions
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
        ([<ExcelArgument(Description= "[Only compares numbers. Default is false.]")>] onlyNumbers: obj) 
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

    // --------------------
    // -- Basic functions
    // --------------------

    [<ExcelFunction(Category="XL", Description="Extracts the nth element of an array.")>]
    let x1_elem
        ([<ExcelArgument(Description= "Values (1D array).")>] values: obj[]) 
        ([<ExcelArgument(Description= "Element index.")>] elementIndex: double)
        ([<ExcelArgument(Description= "[Default Value. Default is #N/A.]")>] defaultValue: obj)
        : obj = 

        // intermediary stage
        let defvalue = In.D0.Absent.Obj.subst ExcelError.ExcelErrorNA defaultValue

        // result
        match values |> Array.tryItem ((int) elementIndex) with
        | None -> defvalue
        | Some o -> o

    [<ExcelFunction(Category="XL", Description="Creates an array (obj[]) from elements.")>]
    let x1_ofElems 
        ([<ExcelArgument(Description= "Element 0.")>] elem00: obj)
        ([<ExcelArgument(Description= "Element 1.")>] elem01: obj)
        ([<ExcelArgument(Description= "Element 2.")>] elem02: obj)
        ([<ExcelArgument(Description= "Element 3.")>] elem03: obj)
        ([<ExcelArgument(Description= "Element 4.")>] elem04: obj)
        ([<ExcelArgument(Description= "Element 5.")>] elem05: obj)
        ([<ExcelArgument(Description= "Element 6.")>] elem06: obj)
        ([<ExcelArgument(Description= "Element 7.")>] elem07: obj)
        ([<ExcelArgument(Description= "Element 8.")>] elem08: obj)
        ([<ExcelArgument(Description= "Element 9.")>] elem09: obj)
        ([<ExcelArgument(Description= "Element 10.")>] elem10: obj)
        ([<ExcelArgument(Description= "Element 11.")>] elem11: obj)
        ([<ExcelArgument(Description= "Element 12.")>] elem12: obj)
        : obj[] =

        // result
        [| elem00; elem01; elem02; elem03; elem04; elem05; elem06; elem07; elem08; elem09; elem10; elem11; elem12 |]
        |> Array.filter (fun x -> match x with | :? ExcelMissing -> false | :? ExcelEmpty -> false | _ -> true)

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
    let x1_sum
        ([<ExcelArgument(Description= "Range.")>] range: obj[,]) 
        : obj =
        
        // intermediary stage
        let a2D = range |> Array2D.map (In.D0.Dbl.def 0.0)
        let len1 = Array2D.length1 a2D

        // result
        let res = [| for i in 0 .. (len1 - 1) -> a2D.[i,*] |> Array.sum |]  |> Array.sum
        box res

    [<ExcelFunction(Category="XL", Description="Sum absolute values (non-numeric values are replaced with 0.0).")>]
    let x1_sumAbs
        ([<ExcelArgument(Description= "Range.")>] range: obj[,]) 
        : obj =
        
        // intermediary stage
        let a2D = range |> Array2D.map (In.D0.Dbl.def 0.0 >> abs)
        let len1 = Array2D.length1 a2D

        // result
        let res = [| for i in 0 .. (len1 - 1) -> a2D.[i,*] |> Array.sum |]  |> Array.sum
        box res

    [<ExcelFunction(Category="XL", Description="Minimum of numeric values (non-numeric values are ignored).")>]
    let x1_min
        ([<ExcelArgument(Description= "Range.")>] range: obj[,]) 
        : obj =
         
        // intermediary stage
        let a2D = range |> Array2D.map (In.D0.Dbl.def Double.MaxValue)
        let len1 = Array2D.length1 a2D

        // result
        let res = [| for i in 0 .. (len1 - 1) -> a2D.[i,*] |> Array.min |]  |> Array.min
        box res

    [<ExcelFunction(Category="XL", Description="Minimum of absolute values (non-numeric values are ignored).")>]
    let x1_minAbs
        ([<ExcelArgument(Description= "Range.")>] range: obj[,]) 
        : obj =
         
        // intermediary stage
        let a2D = range |> Array2D.map (In.D0.Dbl.def Double.MaxValue >> abs)
        let len1 = Array2D.length1 a2D

        // result
        let res = [| for i in 0 .. (len1 - 1) -> a2D.[i,*] |> Array.min |]  |> Array.min
        box res

    [<ExcelFunction(Category="XL", Description="Maximum of numeric values (non-numeric values are ignored).")>]
    let x1_max
        ([<ExcelArgument(Description= "Range.")>] range: obj[,]) 
        : obj =
         
        // intermediary stage
        let a2D = range |> Array2D.map (In.D0.Dbl.def Double.MinValue)
        let len1 = Array2D.length1 a2D

        // result
        let res = [| for i in 0 .. (len1 - 1) -> a2D.[i,*] |> Array.max |]  |> Array.max
        box res

    [<ExcelFunction(Category="XL", Description="Maximum of absolute values (non-numeric values are ignored).")>]
    let x1_maxAbs
        ([<ExcelArgument(Description= "Range.")>] range: obj[,]) 
        : obj =
         
        // intermediary stage
        let a2D = range |> Array2D.map (In.D0.Dbl.def Double.MinValue >> abs)
        let len1 = Array2D.length1 a2D

        // result
        let res = [| for i in 0 .. (len1 - 1) -> a2D.[i,*] |> Array.max |]  |> Array.max
        box res

    [<ExcelFunction(Category="XL", Description="Sum-product of numeric values in 1D-ranges (non-numeric values are replaced with 0.0).")>]
    let x1_sumprod
        ([<ExcelArgument(Description= "1D-range 1.")>] range1: obj)
        ([<ExcelArgument(Description= "1D-range 2.")>] range2: obj) 
        ([<ExcelArgument(Description= "1D-range 3.")>] range3: obj) 
        ([<ExcelArgument(Description= "1D-range 4.")>] range4: obj) 
        ([<ExcelArgument(Description= "1D-range 5.")>] range5: obj) 
        ([<ExcelArgument(Description= "1D-range 6.")>] range6: obj) 
        : obj =
        
        // result
        let a1Ds = 
            [| range1; range2; range3; range4; range5; range6; range1;  |] 
            |> Array.map (In.Cast.to1D false)
            |> Array.map (Array.map (In.D0.Dbl.def 0.0))
            |> Array.reduce (fun a1D1 a1D2 -> Array.map2 (*) a1D1 a1D2)

        let res = a1Ds |> Array.sum
        box res

    [<ExcelFunction(Category="XL", Description="Sum-product of absolute numeric values in 1D-ranges (non-numeric values are replaced with 0.0).")>]
    let x1_sumprodAbs
        ([<ExcelArgument(Description= "1D-range 1.")>] range1: obj)
        ([<ExcelArgument(Description= "1D-range 2.")>] range2: obj) 
        ([<ExcelArgument(Description= "1D-range 3.")>] range3: obj) 
        ([<ExcelArgument(Description= "1D-range 4.")>] range4: obj) 
        ([<ExcelArgument(Description= "1D-range 5.")>] range5: obj) 
        ([<ExcelArgument(Description= "1D-range 6.")>] range6: obj) 
        : obj =
        
        // result
        let a1Ds = 
            [| range1; range2; range3; range4; range5; range6; range1;  |] 
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
