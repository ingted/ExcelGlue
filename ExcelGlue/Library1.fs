﻿namespace ExcelGlue

type Class1() = 
    member this.X = "F#"

module FOO =
    open ExcelDna.Integration

    [<ExcelFunction(Description="My first .NET function")>]
    let HelloDna name =
        "Hello " + name

[<RequireQualifiedAccess>]
module CAST =
    let u = 3
    let v = Option.map
    let vc = Option.defaultValue
    
    [<RequireQualifiedAccess>]
    module Cast =
        let fail<'a> (msg: string option) (o: obj) : 'a =
            match o with
            | :? 'a as v -> v
            | _ -> failwith (msg |> Option.defaultValue "Cast failed.")

        let def<'a> (defvalue: 'a) (o: obj) : 'a =
            match o with
            | :? 'a as v -> v
            | _ -> defvalue

        let defO<'a> (defvalue: obj) (o: obj) : obj =
            match o with
            | :? 'a -> o
            | _ -> defvalue

        let tryDV<'a> (defvalue: 'a option) (o: obj) : 'a option =
            match o with
            | :? 'a as v -> Some v
            | _ -> defvalue

    [<RequireQualifiedAccess>]
    module Bool =
        let def (defvalue: bool) (xlval: obj) = Cast.def<bool> defvalue xlval
        let tryDV (defvalue: bool option) (xlval: obj) = Cast.tryDV<bool> defvalue xlval

    [<RequireQualifiedAccess>]
    module Stg =
        let def (defvalue: string) (xlval: obj) = Cast.def<string> defvalue xlval
        let tryDV (defvalue: string option) (xlval: obj) = Cast.tryDV<string> defvalue xlval