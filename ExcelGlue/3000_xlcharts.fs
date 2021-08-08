//  Copyright (c) cdr021. All rights reserved.
//  ExcelGlue is licensed under the MIT license. See LICENSE.txt for details.

namespace ExcelGlue
 
open API
open Toolbox.Array

[<RequireQualifiedAccess>]
module CHART =
    open System
    open System.Drawing
    open Microsoft.Office.Interop.Excel
    open Microsoft.Office.Core

    // PROBLEM ?
    type LegendPosition = | Top | Bottom | Left | Right | LeftOverlay | RightOverlay | Nil with
        static member ofLabel (label: string) : LegendPosition =
            match label.ToUpper() with
            | "TOP" -> Top
            | "BOT" | "BOTTOM" -> Bottom
            | "L" | "LEFT" -> Left
            | "R" | "RIGHT" -> Right
            | "LO" | "LEFTO" | "LEFTOVERLAY" -> LeftOverlay
            | "RO" | "RIGHTO" | "RIGHTOVERLAY" -> RightOverlay
            | _ -> Nil

        member this.toXlLegend : Microsoft.Office.Core.MsoChartElementType =
            match this with
            | Top -> Microsoft.Office.Core.MsoChartElementType.msoElementLegendTop
            | Bottom -> Microsoft.Office.Core.MsoChartElementType.msoElementLegendBottom
            | Left -> Microsoft.Office.Core.MsoChartElementType.msoElementLegendLeft
            | Right -> Microsoft.Office.Core.MsoChartElementType.msoElementLegendRight
            | LeftOverlay -> Microsoft.Office.Core.MsoChartElementType.msoElementLegendLeftOverlay
            | RightOverlay -> Microsoft.Office.Core.MsoChartElementType.msoElementLegendRightOverlay
            | _ -> Microsoft.Office.Core.MsoChartElementType.msoElementLegendNone

    type AutoValue = | Auto | Fixed of double with
        static member ofLabel (label: obj) : AutoValue option =
            match label with
            | :? string as s -> if (s = "A") || (s = "AUTO") then Some Auto  else None
            | :? double as v -> Fixed v |> Some
            | _ -> None

        static member ofDouble (value: double) : AutoValue = Fixed value

    type AxisCat = | X | Y of int with
        static member ofLabel (label: string) : AxisCat =
            match label.ToUpper() with
            | "X" | "XAXIS" -> X
            | "Y" | "Y1" | "YAXIS" | "YAXIS1" -> Y 1
            | "Y2" | "YAXIS2" -> Y 2
            | _ -> X

        member this.toXlAxis (chrt: Chart) : Microsoft.Office.Interop.Excel.Axis =
            match this with
            | X -> chrt.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory,Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary) :?> Microsoft.Office.Interop.Excel.Axis
            | Y 2 -> chrt.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue,Microsoft.Office.Interop.Excel.XlAxisGroup.xlSecondary) :?> Microsoft.Office.Interop.Excel.Axis
            | Y _ -> chrt.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue,Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary) :?> Microsoft.Office.Interop.Excel.Axis
    
    type AxisCrosses = | Auto | Custom of double | Max | Min with
        static member ofLabel (crossesat: double) (label: string) : AxisCrosses =
            match label.ToUpper() with
            | "A" | "AUTO" -> Auto
            | "C" | "CUSTOM" -> Custom crossesat
            | "MAX" -> Max
            | "MIN" -> Min
            | _ -> Auto

        member this.toXlAxisCrosses : Microsoft.Office.Interop.Excel.XlAxisCrosses =
            match this with
            | Auto -> Microsoft.Office.Interop.Excel.XlAxisCrosses.xlAxisCrossesAutomatic
            | Custom _ -> Microsoft.Office.Interop.Excel.XlAxisCrosses.xlAxisCrossesCustom
            | Max -> Microsoft.Office.Interop.Excel.XlAxisCrosses.xlAxisCrossesMaximum
            | Min -> Microsoft.Office.Interop.Excel.XlAxisCrosses.xlAxisCrossesMinimum

    type LineStyle = | Single | StyleMixed | ThickBetweenThin | ThickThin | ThinThick | ThinThin with
        static member ofLabel (label: string) : LineStyle =
            match label.ToUpper() with
            | "SINGLE" -> Single
            | "MIXED" -> StyleMixed
            | "TBT" | "THICKBETWEENTHIN" -> ThickBetweenThin
            | "THICKTHIN" -> ThickThin
            | "THINTHICK" -> ThinThick
            | "THINTHIN" -> ThinThin
            | _ -> Single

        member this.toXlLine : Microsoft.Office.Core.MsoLineStyle =
            match this with
            | Single -> MsoLineStyle.msoLineSingle
            | StyleMixed -> MsoLineStyle.msoLineStyleMixed
            | ThickBetweenThin -> MsoLineStyle.msoLineThickBetweenThin
            | ThickThin -> MsoLineStyle.msoLineThickThin
            | ThinThick -> MsoLineStyle.msoLineThinThick
            | ThinThin -> MsoLineStyle.msoLineThinThin

    type LineDashStyle = | Dash | DashDot | DashDotDot | DashMixed | LongDash | LongDashDot | LongDashDotDot | RoundDot | Solid | Square | SysDash | SysDashDot | SysDot with
        static member ofLabel (label: string) : LineDashStyle =
            match label.ToUpper() with
            | "DASH" -> Dash
            | "DD" | "DASHDOT" -> DashDot
            | "DDD" | "DASHDOTDOT" -> DashDotDot
            | "DASHMIXED" -> DashMixed
            | "LD" | "LDASH" -> LongDash
            | "LDD" | "LDASHDOT" -> LongDashDot
            | "LDDD" | "LDASHDOTDOT" -> LongDashDotDot
            | "RD" | "ROUNDDOT" -> RoundDot
            | "SOLID" -> Solid
            | "SQ" | "SQUARE" -> Square
            | "SD" | "SDASH" -> SysDash
            | "SDD" | "SDASHDOT" -> SysDashDot
            | "SDOT" -> SysDot
            | _ -> Solid

        member this.toXlLineDash : Microsoft.Office.Core.MsoLineDashStyle =
            match this with
            | Dash -> MsoLineDashStyle.msoLineDash
            | DashDot -> MsoLineDashStyle.msoLineDashDot
            | DashDotDot -> MsoLineDashStyle.msoLineDashDotDot
            | DashMixed -> MsoLineDashStyle.msoLineDashStyleMixed
            | LongDash -> MsoLineDashStyle.msoLineLongDash
            | LongDashDot -> MsoLineDashStyle.msoLineLongDashDot
            | LongDashDotDot -> MsoLineDashStyle.msoLineLongDashDotDot
            | RoundDot -> MsoLineDashStyle.msoLineRoundDot
            | Solid -> MsoLineDashStyle.msoLineSolid
            | Square -> MsoLineDashStyle.msoLineSquareDot
            | SysDash -> MsoLineDashStyle.msoLineSysDash
            | SysDashDot -> MsoLineDashStyle.msoLineSysDashDot
            | SysDot -> MsoLineDashStyle.msoLineSysDot
        
    type DataLabels = | ShowLabel | ShowValue | ShowNone with
        static member ofLabel (label: string) : DataLabels =
            match label.ToUpper() with
            | "LBL" | "LABEL" | "SHWLBL" | "SHOWLABEL" -> ShowLabel
            | "VAL" | "VALUE" | "SHWVAL" | "SHOWVALUE" -> ShowValue
            | _ -> ShowNone

        member this.toXlDataLabelsType : Microsoft.Office.Interop.Excel.XlDataLabelsType =
            match this with
            | ShowLabel -> Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowLabel
            | ShowValue -> Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowValue
            | _ -> Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowNone

    type MarkerStyle = | Automatic | Circle | Dash | Diamond | Dot | Plus | Square | Star | Triangle | Nil with
        static member ofLabel (label: string) : MarkerStyle =
            match label.ToUpper() with
            | "AUTO" -> Automatic
            | "CIRCLE" -> Circle
            | "DASH" -> Dash
            | "DIA" | "DIAMOND" -> Diamond
            | "DOT" -> Dot
            | "PLUS" -> Plus
            | "SQ" | "SQUARE" -> Square
            | "STAR" -> Star
            | "TRI" | "TRIANGLE" -> Triangle
            | _ -> Nil

        member this.toXlMarker : Microsoft.Office.Interop.Excel.XlMarkerStyle =
            match this with
            | Automatic -> Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleAutomatic
            | Circle -> Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleCircle
            | Dash -> Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleDash
            | Diamond -> Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleDiamond
            | Dot -> Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleDot
            | Plus -> Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStylePlus
            | Square -> Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleSquare
            | Star -> Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleStar
            | Triangle -> Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleTriangle            
            | Nil -> Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleNone

        static member ofXlMarker (xlmarkerstyle: Microsoft.Office.Interop.Excel.XlMarkerStyle) : MarkerStyle =
            match xlmarkerstyle with
            | Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleAutomatic -> Automatic
            | Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleCircle -> Circle
            | Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleDash -> Dash
            | Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleDiamond -> Diamond
            | Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleDot -> Dot
            | Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStylePlus -> Plus
            | Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleSquare -> Square
            | Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleStar -> Star
            | Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleTriangle -> Triangle            
            | _ -> Nil

    type Marker = { style: MarkerStyle option; size: int option; fgcol: Color option; bgcol: Color option; hollow: bool option } with
        static member def : Marker = { style = Some Nil; size = Some 7; fgcol = None; bgcol = None; hollow = None }
        
    type Line = { style: LineStyle option; dstyle: LineDashStyle option; color: Color option; weight: double option; visible: bool option; smooth: bool option }  with
        static member allnone : Line = { style = None; dstyle = None; color = None; weight = None; visible = None; smooth = None }
        static member def : Line = { style = Some Single; dstyle = Some Solid; color = None; weight = Some 1.5; visible = Some true; smooth = Some true }

    type ErrorBar = { endstyle: Microsoft.Office.Interop.Excel.XlEndStyleCap option; weight: double option; visible: bool option; color: Color option; transparency: double option } with
        static member def : ErrorBar = { endstyle = Some Microsoft.Office.Interop.Excel.XlEndStyleCap.xlNoCap; weight = Some 7.0; visible = Some true; color = Color.FromArgb(142,0,0) |> Some; transparency = Some 0.0 }

    [<RequireQualifiedAccess>]
    module Color =
        open System.Text.RegularExpressions
        //XlRgbColor.rgbLightSteelBlue
        // System.Drawing.Color.Red
        
        type RGB = int*int*int
        type Palette = Color[]

        let toExcel (color: Color) : int = System.Drawing.ColorTranslator.ToOle(color)
        let ofExcel (icolor: int) : Color = System.Drawing.ColorTranslator.FromOle(icolor)
        
        let ofObj (ocolor: obj) : Color option = 
            match In.D0.Intg.Opt.def None ocolor, In.D0.Stg.Opt.def None ocolor with
            | Some icolor, _ -> ofExcel icolor |> Some
            | _, Some scolor -> 
                let rgxpattern = "^([0-9]{1,3})(?:\.|\-|,|;)([0-9]{1,3})(?:\.|\-|,|;)([0-9]{1,3})$"
                let m = Regex.Match(scolor, rgxpattern)
                if m.Success && m.Groups.Count = 4 then 
                    let rgb = [| m.Groups.[1].Value; m.Groups.[2].Value; m.Groups.[3].Value |] |> Array.map (fun txt -> Int32.Parse(txt))
                    Color.FromArgb(rgb.[0],rgb.[1],rgb.[2]) |> Some
                else
                    None
            | _ -> None

        let gray (intensity: double) (color: Color) = 
            let r, g, b = color.R |> int, color.G |> int, color.B |> int
            if b = ([| r; g; b |] |> Array.max) then
                let cmax = b
                let r' = let gap = (cmax - r) |> double in cmax - ((gap * intensity) |> int)
                let g' = let gap = (cmax - g) |> double in cmax - ((gap * intensity) |> int)
                Color.FromArgb(0, r', g', b)
            elif g = ([| r; g; b |] |> Array.max) then
                let cmax = g
                let r' = let gap = (cmax - r) |> double in cmax - ((gap * intensity) |> int)
                let b' = let gap = (cmax - b) |> double in cmax - ((gap * intensity) |> int)
                Color.FromArgb(0, r', g, b')
            else
                let cmax = r
                let g' = let gap = (cmax - g) |> double in cmax - ((gap * intensity) |> int)
                let b' = let gap = (cmax - b) |> double in cmax - ((gap * intensity) |> int)
                Color.FromArgb(0, r, g', b')

        /// 0 intenstity : colorfrom
        /// 1 intenstity : colorto
        let gradient (colorfrom: Color) (colorto: Color) (intensity: double) = 
            let intensity = min 1.0 (max 0.0 intensity) 
            let rfrom, gfrom, bfrom = colorfrom.R |> int, colorfrom.G |> int, colorfrom.B |> int
            let rto, gto, bto = colorto.R |> int, colorto.G |> int, colorto.B |> int
            let rgap, ggap, bgap = rto - rfrom, gto - gfrom, bto - bfrom
            let r = ((double) rfrom + ((double) rgap) * intensity) |> int
            let g = ((double) gfrom + ((double) ggap) * intensity) |> int
            let b = ((double) bfrom + ((double) bgap) * intensity) |> int
            Color.FromArgb(0, r, g, b)

        [<RequireQualifiedAccess>]
        module Palette =
            let private ofRGB (rgbs: RGB[]) = rgbs |> Array.map (fun (r,g,b) -> Color.FromArgb(0, r, g, b))
            
            let cycle (palette: Palette) (index: int) : Color = palette.[index % palette.Length]

            let rotate (palette: Palette) (idxfrom: int) (shift: int) : Color = cycle palette (idxfrom + shift)

            let excel11 : Palette = 
                [| (69, 114, 167); (170, 70, 67); (137, 165, 78); (113, 88, 143); (65, 152, 175); (219, 132, 61); (147, 169, 207); (209, 147, 146); (185, 205, 150); (169, 155, 189); (145, 195, 213) |]
                |> ofRGB


[<RequireQualifiedAccess>]
module XL =
    open System
    open System.Drawing
    open Microsoft.Office.Interop.Excel
    open Microsoft.Office.Core
    
    [<RequireQualifiedAccess>]
    module COM =
        let errNA = new System.Runtime.InteropServices.ErrorWrapper(-2146826246)

    [<RequireQualifiedAccess>]
    module Range =
        let tryDim (range: Range) : (double*double*double*double) option =
            match range.Left, range.Top, range.Width, range.Height with
            | (:? double as l), (:? double as t), (:? double as w), (:? double as h) -> (l, t, w, h) |> Some
            | _ -> None

    [<RequireQualifiedAccess>]
    module NamedRange =
        let exists (wks: Worksheet) (rangename: string) : bool =
            try
                let rng = wks.Range(rangename)
                true
            with
            | e -> false

        let tryRange (wks: Worksheet) (rangename: string) : Range option =
            try
                let nm = wks.Names.Item(rangename) // :?> Name
                let rng = nm.RefersToRange
                Some rng 
            with
            | e -> None

    [<RequireQualifiedAccess>]
    module ChartO =
        let exists (app: Application) (sheetname: string) (chartname: string) : bool =
            try
                let wks   = (app.Sheets.Item sheetname) :?> Worksheet
                let chrtO = wks.ChartObjects(chartname) :?> ChartObject
                true
            with
            | e -> false

        let tryChartObject (wks: Worksheet) (chartname: string) : ChartObject option =
            try
                let chrtO = wks.ChartObjects(chartname) :?> ChartObject
                Some chrtO 
            with
            | e -> None

        let create (wks: Worksheet) (nmdrange: Range option) (chartname: string) : Microsoft.Office.Interop.Excel.ChartObject =
            match tryChartObject wks chartname with
            | None -> ()
            | Some chrtO -> chrtO.Delete() |> ignore

            try
                let chrtOs = wks.ChartObjects() :?> Microsoft.Office.Interop.Excel.ChartObjects
                let (left, top, width, height) = 
                    match nmdrange |> Option.map Range.tryDim |> Option.flatten with
                    | None -> (100.0,200.0,200.0,100.0)  // MAGIC NUMBER
                    | Some dims -> dims
                let newchrtO = chrtOs.Add(left, top, width, height)
                newchrtO.Name <- chartname
                newchrtO.Chart.ChartType <- Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterLines
                newchrtO.Chart.DisplayBlanksAs <- Microsoft.Office.Interop.Excel.XlDisplayBlanksAs.xlNotPlotted
                newchrtO
            with
            | e -> failwith e.Message

        //let ofName (chrtO: ChartObject) (name: string) : Series option = 
        //    try
        //        match chrtO.Chart.SeriesCollection(name) with
        //        | :? Series as srs -> Some srs
        //        | _ -> None
        //    with _ -> None        

    [<RequireQualifiedAccess>]
    module Chart =
        
        [<RequireQualifiedAccess>]
        module ChartArea =
            let clearContents (chrtO: ChartObject) : unit =
                 chrtO.Chart.ChartArea.ClearContents() |> ignore
            
            let set  
                (fillcolor: Color option) (filltransparency: double option)
                (linecolor: Color option) (linetransparency: double option)
                (chrtO: ChartObject)
                : unit =

                let xcarea = chrtO.Chart.ChartArea
                
                match fillcolor with 
                    | None -> () 
                    | Some col -> 
                        xcarea.Format.Fill.Solid()
                        xcarea.Format.Fill.ForeColor.RGB <- col |> CHART.Color.toExcel

                match linecolor with | None -> () | Some col -> xcarea.Format.Line.ForeColor.RGB <- col |> CHART.Color.toExcel
                
                match filltransparency with 
                    | None -> () 
                    | Some trans -> 
                        let trans = min 1.0 (max 0.0 trans)
                        xcarea.Format.Fill.Transparency <- (float32) trans
                
                match linetransparency with 
                    | None -> () 
                    | Some trans -> 
                        let trans = min 1.0 (max 0.0 trans)
                        xcarea.Format.Line.Transparency <- (float32) trans
        
        [<RequireQualifiedAccess>]
        module PlotArea =

            let set  
                (fillcolor: Color option) (filltransparency: double option)
                (linecolor: Color option) (linetransparency: double option)
                (chrtO: ChartObject)
                : unit =

                let xparea = chrtO.Chart.PlotArea
                
                match fillcolor with 
                    | None -> () 
                    | Some col -> 
                        xparea.Format.Fill.Solid()
                        xparea.Format.Fill.ForeColor.RGB <- col |> CHART.Color.toExcel

                match linecolor with | None -> () | Some col -> xparea.Format.Line.ForeColor.RGB <- col |> CHART.Color.toExcel
                
                match filltransparency with 
                    | None -> () 
                    | Some trans -> 
                        let trans = min 1.0 (max 0.0 trans)
                        xparea.Format.Fill.Transparency <- (float32) trans
                
                match linetransparency with 
                    | None -> () 
                    | Some trans -> 
                        let trans = min 1.0 (max 0.0 trans)
                        xparea.Format.Line.Transparency <- (float32) trans

        let set 
            (title: string option)
            (lgdposition: CHART.LegendPosition option) (lgdfontsize: int option) (lgdfontbold: bool option)
            (chrtO: ChartObject) //(name: string)
            : unit = 
            
            let chrt = chrtO.Chart

            // title
            match title with 
            | None -> chrt.HasTitle <- false 
            | Some ttl ->
                chrt.HasTitle <- true 
                chrt.ChartTitle.Text <- ttl
            
            // legend
            let sc = chrt.SeriesCollection() :?> Microsoft.Office.Interop.Excel.SeriesCollection
            if sc.Count > 0 then
                match lgdposition with | None -> () | Some legpos -> chrt.SetElement(legpos.toXlLegend)
            if chrt.HasLegend then
                match lgdfontsize with | None -> () | Some size -> chrt.Legend.Font.Size <- size
                match lgdfontbold with | None -> () | Some flag -> chrt.Legend.Font.Bold <- flag

        // let clearContents (chrtO: ChartObject) : unit = chrtO.Chart.ChartArea.ClearContents() |> ignore
        let newSeries (chrtO: ChartObject) (name: string) : Series = 
            let chrt = chrtO.Chart
            let collec = chrt.SeriesCollection() :?> Microsoft.Office.Interop.Excel.SeriesCollection
            let newsrs = collec.NewSeries()
            newsrs.Name <- name
            newsrs

    [<RequireQualifiedAccess>]
    module Axis =
        let setGrid 
            (minorgrid: bool option) (majorgrid: bool option)
            (chrtO: ChartObject) (axis: CHART.AxisCat) : unit =

            let xaxis = axis.toXlAxis chrtO.Chart
            match minorgrid with | None -> () | Some flag -> xaxis.HasMinorGridlines <- flag
            match majorgrid with | None -> () | Some flag -> xaxis.HasMajorGridlines <- flag
        
        let set 
            (minval: CHART.AutoValue option) (maxval: CHART.AutoValue option)
            (minU: CHART.AutoValue option) (majU: CHART.AutoValue option)
            (crosses: CHART.AxisCrosses option)
            (numfmt: string option)
            (chrtO: ChartObject) (axis: CHART.AxisCat) : unit =

            let xaxis = axis.toXlAxis chrtO.Chart
            match minval with
            | None -> ()
            | Some av -> match av with | CHART.Auto -> xaxis.MinimumScaleIsAuto <- true | CHART.Fixed v -> xaxis.MinimumScale <- v

            match maxval with
            | None -> ()
            | Some av -> match av with | CHART.Auto -> xaxis.MaximumScaleIsAuto <- true | CHART.Fixed v -> xaxis.MaximumScale <- v

            match minU with
            | None -> ()
            | Some av -> match av with | CHART.Auto -> xaxis.MinorUnitIsAuto <- true | CHART.Fixed v -> xaxis.MinorUnit <- v

            match majU with
            | None -> ()
            | Some av -> match av with | CHART.Auto -> xaxis.MajorUnitIsAuto <- true | CHART.Fixed v -> xaxis.MajorUnit <- v

            match numfmt with | None -> () | Some nfmt -> xaxis.TickLabels.NumberFormat <- nfmt

            match crosses with 
                | None -> () 
                | Some (CHART.AxisCrosses.Custom crossat) -> 
                    let crss = CHART.AxisCrosses.Custom crossat
                    xaxis.Crosses <- crss.toXlAxisCrosses
                    xaxis.CrossesAt <- crossat
                | Some crss -> 
                    xaxis.Crosses <- crss.toXlAxisCrosses
            

    [<RequireQualifiedAccess>]
    module Series =
        let parentChart (xsrs: Series) : Chart =
            let chrtgroup = xsrs.Parent :?> ChartGroup
            let chrt = chrtgroup.Parent :?> Chart
            chrt

        let count (xsrs: Series) : int =
            let xs = xsrs.XValues :?> Array
            xs.Length

        let trySeries (chrtO: ChartObject) (name: string) : Series option = 
            try
                match chrtO.Chart.SeriesCollection(name) with
                | :? Series as srs -> Some srs
                | _ -> None
            with _ -> None

        let set 
            (line: CHART.Line option) (marker: CHART.Marker option)
            (xlErrNA: obj) (xvals: double[] option) (yvals: double[] option)
            (ebar: CHART.ErrorBar option) (ebvals: double[] option)
            (datalabels: CHART.DataLabels option) (labelsize: double option) (labelnfmt: string option) (filterLabels: (int -> bool) option)
            (axisgroup: int option)
            (xsrs: Series) 
            : unit = 
            
            let chrt = parentChart xsrs

            // series settings
            match line with 
            | None -> ()
            | Some lin ->
                match lin.smooth with | None -> () | Some flag -> xsrs.Smooth <- flag
                match axisgroup with | None -> () | Some grp -> xsrs.AxisGroup <- if grp > 1 then Microsoft.Office.Interop.Excel.XlAxisGroup.xlSecondary else Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary

                // line settings
                match lin.color with | None -> () | Some col -> xsrs.Format.Line.ForeColor.RGB <- col |> CHART.Color.toExcel
                match lin.weight with | None -> () | Some weight -> xsrs.Format.Line.Weight <- single weight
                match lin.style with | None -> () | Some style -> xsrs.Format.Line.Style <- style.toXlLine
                match lin.dstyle with | None -> () | Some dstyle -> xsrs.Format.Line.DashStyle <- dstyle.toXlLineDash
                match lin.visible with | None -> () | Some flag -> xsrs.Format.Line.Visible <- if flag then MsoTriState.msoTrue else MsoTriState.msoFalse

            // xvalues & yvalues
            match xvals with | None -> () | Some xs -> xsrs.XValues <- xs |> Array.map (Out.D0.Dbl.out { Out.Proxys.def with nan = (COM.errNA |> box) })
            match yvals with | None -> () | Some ys -> xsrs.Values <- ys |> Array.map (Out.D0.Dbl.out { Out.Proxys.def with nan = (COM.errNA |> box) })
            let yyvals = match yvals with None -> [||] | Some ys -> ys |> Array.map (Out.D0.Dbl.out { Out.Proxys.def with nan = (COM.errNA |> box) })

            // markers settings
            match marker with 
            | None -> ()
            | Some mkr ->
                match mkr.style with 
                    | None -> () 
                    | Some sty -> 
                        let msty = sty.toXlMarker
                        xsrs.MarkerStyle <- msty

                let srsmkrsty = xsrs.MarkerStyle |> CHART.MarkerStyle.ofXlMarker

                if (srsmkrsty <> CHART.MarkerStyle.Nil) then
                    let bgcolor = chrt.PlotArea.Format.Fill.ForeColor.RGB
                    if (srsmkrsty <> CHART.MarkerStyle.Plus) then // maybe add more exception here
                        match mkr.bgcol with | None -> () | Some col -> xsrs.MarkerBackgroundColor <- col |> CHART.Color.toExcel
                    else
                        xsrs.MarkerBackgroundColor <- bgcolor
                    match mkr.fgcol with | None -> () | Some col -> xsrs.MarkerForegroundColor <- col |> CHART.Color.toExcel
                    match mkr.size with | None -> () | Some sz -> xsrs.MarkerSize <- sz

            match ebvals with
            | None -> ()
            | Some ebs ->
                if ebs.Length = 0 then
                    ()
                else
                    let ebs' = ebs |> Array.map (fun d -> if Double.IsNaN d then 0.0 else d)
                    let ebs'' = ebs' |> sub (Some 0) ((count xsrs) |> Some)
                    xsrs.ErrorBar(XlErrorBarDirection.xlY, Microsoft.Office.Interop.Excel.XlErrorBarInclude.xlErrorBarIncludePlusValues, Microsoft.Office.Interop.Excel.XlErrorBarType.xlErrorBarTypeCustom, ebs'') |> ignore
            
            match ebar with
            | None -> ()
            | Some eb ->
                if not xsrs.HasErrorBars then
                    ()
                else
                    match eb.endstyle with | None -> () | Some esty -> xsrs.ErrorBars.EndStyle <- esty
                    match eb.weight with | None -> () | Some weight -> xsrs.ErrorBars.Format.Line.Weight <- (float32) weight
                    match eb.visible with | None -> () | Some flag -> xsrs.ErrorBars.Format.Line.Visible <- if flag then MsoTriState.msoTrue else MsoTriState.msoFalse
                    match eb.color with | None -> () | Some col -> xsrs.ErrorBars.Format.Line.ForeColor.RGB <- col |> CHART.Color.toExcel

            // data labels setting
            match datalabels with
            | None -> ()
            | Some dlabels -> xsrs.ApplyDataLabels(dlabels.toXlDataLabelsType) |> ignore

            match labelsize with
            | None -> ()
            | Some sz -> 
                let dlabels = xsrs.DataLabels() :?> Microsoft.Office.Interop.Excel.DataLabels
                dlabels.Font.Size <- sz

            match labelnfmt with
            | None -> ()
            | Some nfmt -> 
                let dlabels = xsrs.DataLabels() :?> Microsoft.Office.Interop.Excel.DataLabels
                dlabels.NumberFormat <- nfmt

            match filterLabels with
            | None -> ()
            | Some filterLbls -> 
                for i in 1 .. count xsrs do
                    if filterLbls i |> not then 
                        let pnt = xsrs.Points(i) :?> Microsoft.Office.Interop.Excel.Point
                        pnt.DataLabel.Delete() |> ignore

        //let get
        //    (showxs: bool)
        //    (xsrs: Series) 
        //    //: unit = 
        //    =

        //    // let xlErrNA : int = -2146826246
        //    let chrt = parentChart xsrs
        //    //let mutable xs : obj = box 0
        //    let mutable ys : obj = box 0

        //    // xvalues & yvalues
        //    //xs <- xsrs.XValues // :?> obj[]
        //    let xs = xsrs.XValues :?> seq<_>
        //    ys <- xsrs.Values // :?> obj[]
        //    let strxs = xs.ToString()
        //    let strys = xs.ToString()
        //    // let xxs = [| for x in xs -> x |]
        //    // let yys = [| for x in xs -> x |]

        //    if showxs then box strxs else box strys

module XL_XLCHART =
    open System
    open Microsoft.Office.Interop.Excel
    open ExcelDna.Integration
    open Registry
    
    let private tryExtract<'a> = MRegistry.tryExtract<'a>

    [<ExcelFunction(Category="Excel Chart", Description="Returns an Excel ChartObject reg. object.")>]
    let xc_ofRng
        ([<ExcelArgument(Description= "Sheet name.]")>] sheetName: string)
        ([<ExcelArgument(Description= "[Chart name. Default is \"Chart 1\".]")>] chartName: obj)
        ([<ExcelArgument(Description= "[Named range's name for size. Should be local. Default is none.]")>] rangeName: obj)
        ([<ExcelArgument(Description= "[Force create. Default is false.]")>] forceCreate: obj)
        : obj = 

        // return on error (sentinel value)
        let snel = ExcelError.ExcelErrorNA |> box

        // intermediary arguments / calculations
        let chartname = In.D0.Stg.def "Chart 1" chartName
        let forcecreate = In.D0.Bool.def false forceCreate
        let nmdrange = In.D0.Stg.Opt.def None rangeName
        
        // caller cell's reference ID
        let rfid = MRegistry.refID

        // result
        try
            let app = ExcelDnaUtil.Application :?> Application
            let wks   = (app.Sheets.Item sheetName) :?> Worksheet
            let nmdrange = In.D0.Stg.Opt.def None rangeName |> Option.map (XL.NamedRange.tryRange wks) |> Option.flatten

            let chrtO =
                if forcecreate then 
                    XL.ChartO.create wks nmdrange chartname |> Some
                else
                    XL.ChartO.tryChartObject wks chartname

            match chrtO with
            | None -> snel
            | Some chrto ->
                chrto |> MRegistry.registerBxd rfid
        with
        | e -> box e.Message

    [<ExcelFunction(Category="Excel Chart", Description="Creates a new series.")>]
    let xc_newSeries
        ([<ExcelArgument(Description= "ChartO reg. obj.")>] rgChrtO: string)
        ([<ExcelArgument(Description= "Series name.")>] seriesName: string)
        : obj = 

        // result
        match tryExtract<ChartObject> rgChrtO with
        | None -> Out.Proxys.def.failed
        | Some chrto -> 
            XL.Chart.newSeries chrto seriesName |> ignore
            let now = DateTime.Now
            box now

    [<ExcelFunction(Category="Excel Chart", Description="Clear chart area contents.")>]
    let xc_clearContents ([<ExcelArgument(Description= "ChartO reg. obj.")>] rgChrtO: string)
        : obj = 

        // result
        match tryExtract<ChartObject> rgChrtO with
        | None -> Out.Proxys.def.failed
        | Some chrto -> 
            XL.Chart.ChartArea.clearContents chrto
            let now = DateTime.Now
            box now

    // REAL PROBLEM ???
    [<ExcelFunction(Category="Excel Chart", Description="Sets a chart's properties.")>]
    let xc_setChart
        ([<ExcelArgument(Description= "ChartO reg. obj.")>] rgChrtO: string)
        ([<ExcelArgument(Description= "[Title. Default is none.]")>] chartTitle: obj)
        ([<ExcelArgument(Description= "[Legend position. Default is none.]")>] legendPosition: obj)
        ([<ExcelArgument(Description= "[Legend font size. Default is none.]")>] legendFontSize: obj)
        ([<ExcelArgument(Description= "[Legend font bold. Default is none.]")>] legendFontBold: obj)
        : obj = 

        // return on error (sentinel value)
        let snel = ExcelError.ExcelErrorNA |> box

        // intermediary arguments / calculations
        let title = In.D0.Stg.Opt.def None chartTitle
        let lgdposition = legendPosition |> In.D0.Missing.tryMap (In.D0.Stg.Opt.def None) |> Option.map CHART.LegendPosition.ofLabel
        // 
        let lgdfontsize = In.D0.Intg.Opt.def None legendFontSize
        let lgdfontbold = In.D0.Bool.Opt.def None legendFontBold
        
        // result
        match tryExtract<ChartObject> rgChrtO with
        | None -> snel
        | Some chrto ->
            chrto |> XL.Chart.set title lgdposition lgdfontsize lgdfontbold
            box "Success."

    [<ExcelFunction(Category="Excel Chart", Description="Sets a chart series' properties.")>]
    let xc_setSeries
        ([<ExcelArgument(Description= "ChartO reg. obj.")>] rgChrtO: string)
        ([<ExcelArgument(Description= "Series name.")>] seriesName: string)
        ([<ExcelArgument(Description= "[Series xvalues. Default is none.]")>] xValue: obj)
        ([<ExcelArgument(Description= "[Series yvalues. Default is none.]")>] yValue: obj)
        ([<ExcelArgument(Description= "[Series color. Default is no effect.]")>] seriesColor: obj)
        ([<ExcelArgument(Description= "[Visible line. Default is no effect.]")>] visibleLine: obj)
        ([<ExcelArgument(Description= "[Smooth line. Default is no effect.]")>] smoothLine: obj)
        ([<ExcelArgument(Description= "[Line Weight. Default is no effect.]")>] lineWeight: obj)
        ([<ExcelArgument(Description= "[Line style. E.g. SINGLE, MIXED, TBT, THICKTHIN, THINTHICK, THINTHIN. Default is no effect.]")>] lineStyle: obj)
        ([<ExcelArgument(Description= "[Line dash style. E.g. DASH, DD, DDD, DASHMIXED, LD, LDD, LDDD, RD, SOLID, SQ, SD, SDD, SDOT. Default is no effect.]")>] lineDashStyle: obj)
        ([<ExcelArgument(Description= "[Marker style. E.g. AUTO, NONE, DASH, DIA, DOT, PLUS, SQUARE, STAR, TRIANGLE. Default is no effect.]")>] mkrStyle: obj)
        ([<ExcelArgument(Description= "[Marker size. Default is no effect.]")>] markerSize: obj)
        ([<ExcelArgument(Description= "[Marker hollow. True or False. Default is no effect.]")>] markerHollow: obj)
        ([<ExcelArgument(Description= "[AxisGroup. Default is no effect.]")>] axisGroup: obj)
        ([<ExcelArgument(Description= "[ErrorBar values. Default is none.]")>] errorBarValues: obj)
        ([<ExcelArgument(Description= "[Show DataLabels. Default is no effect.]")>] showDataLabels: obj)
        ([<ExcelArgument(Description= "[DataLabels size. Default is no effect.]")>] dataLabelSize: obj)
        ([<ExcelArgument(Description= "[DataLabels number format. Default is no effect.]")>] dataLabelNumFormat: obj)
        ([<ExcelArgument(Description= "[DataLabels keep ratio. Default is keep all.]")>] dataLabelKeepRatio: obj)
        : obj = 

        // intermediary stage
        let xvals = xValue |> In.D0.Missing.Obj.tryO |> Option.map (In.D1.ODbl.def Double.NaN)
        let yvals = yValue |> In.D0.Missing.Obj.tryO |> Option.map (In.D1.ODbl.def Double.NaN)
        let ebvals = errorBarValues |> In.D0.Missing.Obj.tryO |> Option.map (In.D1.ODbl.def Double.NaN)

        let linecolor = seriesColor |> In.D0.Missing.tryMap (CHART.Color.ofObj)
        let linevisible = In.D0.Bool.Opt.def None visibleLine
        let linesmooth = In.D0.Bool.Opt.def None smoothLine
        let linestyle = In.D0.Stg.Opt.def None lineStyle |> Option.map CHART.LineStyle.ofLabel
        let linedstyle = In.D0.Stg.Opt.def None lineDashStyle |> Option.map CHART.LineDashStyle.ofLabel
        let lineweight = In.D0.Dbl.Opt.def None lineWeight
        let mkrstyle = In.D0.Stg.Opt.def None mkrStyle |> Option.map CHART.MarkerStyle.ofLabel
        let mkrHollow = In.D0.Bool.Opt.def None markerHollow
        let mkrbcolor = linecolor
        let mkrfcolor = linecolor
        let mkrsize = In.D0.Dbl.Opt.def None markerSize |> Option.map int
        let axisgroup = In.D0.Dbl.Opt.def None axisGroup |> Option.map int
        let datalabels = In.D0.Stg.Opt.def None showDataLabels |> Option.map CHART.DataLabels.ofLabel
        let labelsize = In.D0.Dbl.Opt.def None dataLabelSize
        let labelnfmt = In.D0.Stg.Opt.def None dataLabelNumFormat
        let labelFilter = 
            match In.D0.Intg.Opt.def None dataLabelKeepRatio with
            | Some x when x = 0 -> None
            | Some x when x <= 1 -> None
            | Some x when x > 0 -> 
                let fn = fun i -> i % x = 0
                Some fn
            | _ -> None

        // result
        match tryExtract<ChartObject> rgChrtO with
        | None -> Out.Proxys.def.failed
        | Some chrto -> 
            match seriesName |> XL.Series.trySeries chrto with
            | None -> Out.Proxys.def.failed
            | Some xsrs -> 
                let line : CHART.Line = { style = linestyle; dstyle = linedstyle; color = linecolor; weight = lineweight; visible = linevisible; smooth = linesmooth }
                let marker : CHART.Marker = { style = mkrstyle; size = mkrsize; fgcol = linecolor; bgcol = linecolor; hollow = mkrHollow } // TODO : map lineweight
                xsrs |> XL.Series.set (Some line) (Some marker) (ExcelError.ExcelErrorNA |> box) xvals yvals (Some CHART.ErrorBar.def) ebvals datalabels labelsize labelnfmt labelFilter axisgroup
                let now = DateTime.Now
                box now

    [<ExcelFunction(Category="Excel Chart", Description="Sets a chart axis' properties.")>]
    let xc_setAxis
        ([<ExcelArgument(Description= "ChartO reg. obj.")>] rgChrtO: string)
        ([<ExcelArgument(Description= "[Axis. E.g. X, Y1 or Y2. Default is X.]")>] axis: obj)
        ([<ExcelArgument(Description= "[Minimum value. E.g. \"Auto\" or value. Default is no effect.]")>] minValue: obj)
        ([<ExcelArgument(Description= "[Maximum value. E.g. \"Auto\" or value. Default is no effect.]")>] maxValue: obj)
        ([<ExcelArgument(Description= "[Number format. E.g. \"General\" or \"0.00\". Default is no effect.]")>] numberFormat: obj)
        ([<ExcelArgument(Description= "[Crosses. E.g. \"Auto\", \"Custom\", \"Max\" or value. Default is no effect.]")>] crosses: obj)
        : obj = 

        // return on error (sentinel value)
        let snel = ExcelError.ExcelErrorNA |> box

        // intermediary arguments / calculations
        let axis = In.D0.Stg.def "X" axis |> CHART.AxisCat.ofLabel
        let minval = CHART.AutoValue.ofLabel minValue
        let maxval = CHART.AutoValue.ofLabel maxValue
        let numfmt = In.D0.Stg.Opt.def None numberFormat  // LAST - 1  !!!
        let crosses = 
            match crosses with
            | :? double as crossesat -> CHART.AxisCrosses.ofLabel crossesat "CUSTOM" |> Some
            | :? string as s -> CHART.AxisCrosses.ofLabel 0.0 s |> Some
            | _ -> None

        // result
        match tryExtract<ChartObject> rgChrtO with
        | None -> snel
        | Some chrto -> 
            XL.Axis.set minval maxval None None crosses numfmt chrto axis // LAST
            box "Success."

    [<ExcelFunction(Category="Excel Chart", Description="Sets a chart axis' grids.")>]
    let xc_setGrids
        ([<ExcelArgument(Description= "ChartO reg. obj.")>] rgChrtO: string)
        ([<ExcelArgument(Description= "[Axis. E.g. X, Y1 or Y2. Default is X.]")>] axis: obj)
        ([<ExcelArgument(Description= "[Minor grid. Default is no effect.]")>] minorGrid: obj)
        ([<ExcelArgument(Description= "[Major grid. Default is no effect.]")>] majorGrid: obj)
        : obj = 

        // return on error (sentinel value)
        let snel = ExcelError.ExcelErrorNA |> box

        // intermediary arguments / calculations
        let axis = In.D0.Stg.def "X" axis |> CHART.AxisCat.ofLabel
        let minorgrid = In.D0.Bool.Opt.def None minorGrid
        let majorgrid = In.D0.Bool.Opt.def None majorGrid

        // result
        match tryExtract<ChartObject> rgChrtO with
        | None -> snel
        | Some chrto ->
            XL.Axis.setGrid minorgrid majorgrid chrto axis
            box "Success."

    [<ExcelFunction(Category="Excel Chart", Description="Returns a palette's color.")>]
    let xc_colorExcel
        ([<ExcelArgument(Description= "Palette. TBI.")>] palette: obj)
        ([<ExcelArgument(Description= "Color index.")>] colorIndex: double)
        : obj =

        // result
        CHART.Color.Palette.cycle CHART.Color.Palette.excel11 ((int) colorIndex) |> CHART.Color.toExcel 
        |> box

    [<ExcelFunction(Category="Excel Chart", Description="Sets a chart series' properties.")>]
    let xc_setAreas
        ([<ExcelArgument(Description= "ChartO reg. obj.")>] rgChrtO: string)
        ([<ExcelArgument(Description= "[Chart Area Fill color. Default is no effect.]")>] chartAreaFillColor: obj)
        ([<ExcelArgument(Description= "[Chart Area Fill transparency. 0 for opaque, 1 for transparent. Default is no effect.]")>] chartAreaFillTransparency: obj)
        ([<ExcelArgument(Description= "[Chart Area Line color. Default is no effect.]")>] chartAreaLineColor: obj)
        ([<ExcelArgument(Description= "[Chart Area Line transparency. 0 for opaque, 1 for transparent. Default is no effect.]")>] chartAreaLineTransparency: obj)
        ([<ExcelArgument(Description= "[Plot Area Fill color. Default is no effect.]")>] plotAreaFillColor: obj)
        ([<ExcelArgument(Description= "[Plot Area Fill transparency. 0 for opaque, 1 for transparent. Default is no effect.]")>] plotAreaFillTransparency: obj)
        ([<ExcelArgument(Description= "[Plot Area Line color. Default is no effect.]")>] plotAreaLineColor: obj)
        ([<ExcelArgument(Description= "[Plot Area Line transparency. 0 for opaque, 1 for transparent. Default is no effect.]")>] plotAreaLineTransparency: obj)
        : obj = 

        // return on error (sentinel value)
        let snel = ExcelError.ExcelErrorNA |> box

        // intermediary arguments / calculations
        let cafillcolor = chartAreaFillColor |> In.D0.Missing.tryMap (CHART.Color.ofObj)
        let cafilltrans = In.D0.Dbl.Opt.def None chartAreaFillTransparency
        let calinecolor = chartAreaLineColor |> In.D0.Missing.tryMap (CHART.Color.ofObj)
        let calinetrans = In.D0.Dbl.Opt.def None chartAreaLineTransparency
        let pafillcolor = plotAreaFillColor |> In.D0.Missing.tryMap (CHART.Color.ofObj)
        let pafilltrans = In.D0.Dbl.Opt.def None plotAreaFillTransparency
        let palinecolor = plotAreaLineColor |> In.D0.Missing.tryMap (CHART.Color.ofObj)
        let palinetrans = In.D0.Dbl.Opt.def None plotAreaLineTransparency

        // result
        match tryExtract<ChartObject> rgChrtO with
        | None -> snel
        | Some chrto ->
            chrto |> XL.Chart.ChartArea.set cafillcolor cafilltrans calinecolor calinetrans
            chrto |> XL.Chart.PlotArea.set pafillcolor pafilltrans palinecolor palinetrans
            box "Success."

    [<ExcelFunction(Category="Excel Chart", Description="Sets a chart series' properties.")>]
    let xc_srsDetails
        ([<ExcelArgument(Description= "ChartO reg. obj.")>] rgChrtO: string)
        ([<ExcelArgument(Description= "Series name.")>] seriesName: string)
        ([<ExcelArgument(Description= "[Series detail. Default is Count.]")>] seriesDetail: obj)
        : obj = 

        // return on error (sentinel value)
        let snel = ExcelError.ExcelErrorNA |> box

        // intermediary arguments / calculations
        let detail = In.D0.Stg.def "COUNT" seriesDetail

        // result
        match tryExtract<ChartObject> rgChrtO with
        | None -> snel
        | Some chrto -> 
            match seriesName |> XL.Series.trySeries chrto with
            | None -> snel
            | Some xsrs -> 
                if detail.ToUpper() = "COUNT" then
                    XL.Series.count xsrs
                    |> box
                else
                    snel