/// Generate feature stickies from Azure DevOps features
open FSharp.Data
open PdfSharp.Drawing
open PdfSharp.Pdf
open System.Diagnostics
open PdfSharp.Drawing.Layout

/// Path to CSV file with list of features (obtained from ADO using Excel/CSV export)
[<Literal>]
let CsvFile = @"C:\FeatureList.csv"

let outputPdf = "FeatureStickiers.pdf"

type Items  = CsvProvider<CsvFile, HasHeaders = true, AssumeMissingValues = false>

[<EntryPoint>]
let main argv =
    let items = Items.Load(CsvFile)
    let sorted =
        items.Rows
        |> Seq.sortBy (fun i -> i.``Backlog Priority``.GetValueOrDefault(0))
        |> Seq.filter (fun i -> i.``Work Item Type`` = "Feature")
        //|> Seq.filter (fun i -> i.``Iteration Path`` = @"AreaPath\Milestone")
    printfn "%A" sorted

    use document = new PdfDocument()
    document.Info.Title <- "planningstickies"

    let titleFont = XFont("Verdana", 28.0, XFontStyle.Bold)
    let areaFont = XFont("Verdana", 18.0, XFontStyle.Regular)
    let tagFont = XFont("Verdana", 24.0, XFontStyle.Italic)
    let idFont = XFont("Verdana", 34.0, XFontStyle.Regular)

    let indexFont = XFont("Verdana", 24.0, XFontStyle.Regular)

    sorted
    |> Seq.iteri (fun i item ->
        let page = document.AddPage()
        page.Height <- XUnit 340.0
        let gfx = XGraphics.FromPdfPage(page)

        let pageRect = XRect(0.0, 0.0, float page.Width, float page.Height)

        let outsideMarginX = 25.0
        let outsideMarginY = 15.0

        pageRect.Inflate(-outsideMarginX, -outsideMarginY)

        let s = XSize(10.0,10.0)

        let areaUnder prefix = item.``Area Path``.StartsWith(prefix)
        let c =
            if areaUnder @"\Fuzzing" then
                XKnownColor.Aquamarine
            elif areaUnder @"\Service" then
                XKnownColor.Lavender
            elif areaUnder @"\Web" then
                XKnownColor.BurlyWood
            elif areaUnder @"\Success" then
                XKnownColor.LightBlue
            elif areaUnder @"\General" then
                XKnownColor.LightGreen
            elif areaUnder @"\REST Fuzzing" then
                XKnownColor.DarkSalmon
            else
                XKnownColor.White

        let b = XSolidBrush(XColor.FromKnownColor(c))
        let p = XPen(XColor.FromKnownColor(XKnownColor.BlueViolet), 3.0)
        gfx.DrawRoundedRectangle(p, b, pageRect, s)

        pageRect.Inflate(-10.0,-10.0)

        let tf = XTextFormatter(gfx, Alignment = XParagraphAlignment.Left)
        tf.DrawString(item.Title, titleFont, XBrushes.Black,
             pageRect,
             XStringFormats.TopLeft)

        tf.DrawString(item.``Area Path``, areaFont, XBrushes.Black,
             pageRect + XPoint(0.0, 140.0),
             XStringFormats.TopLeft)

        tf.DrawString(item.Tags, tagFont, XBrushes.Black,
             pageRect + XPoint(0.0, 190.0),
             XStringFormats.TopLeft)

        gfx.DrawString(sprintf "#%d  " item.ID, idFont, XBrushes.Black,
            pageRect,
            XStringFormats.BottomRight)

        gfx.DrawString(sprintf "P%d  " i, indexFont, XBrushes.Black,
            pageRect,
            XStringFormats.BottomLeft)
    )

    document.Save(outputPdf)

    let _ = Process.Start(outputPdf)
    0