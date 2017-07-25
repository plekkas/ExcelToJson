open Newtonsoft.Json
open Microsoft.Office.Interop

type OutputType (_category : string, _tuples : list<string * string>) = 
    member this.category = _category
    member this.tuples = _tuples

let path = @"C:\ExcelPath\Input.xlsx"
let sheets = [ "Sheet1"; "Sheet2"; "Sheet3" ]

let xl = new Excel.ApplicationClass()
let wb = xl.Workbooks.Open(path)

let getSheet sheetName = wb.Worksheets.[sheetName] :?> Excel.Worksheet

let openSheets =
    sheets
    |> List.map (fun s -> getSheet s)

let getCell (cell:Excel.Range, r : int,  c : int) = cell.[r,c] :?> Excel.Range

let getStringTuple (t1 : Excel.Range, t2: Excel.Range) = (string t1.Value2, string t2.Value2)

let tupleList (sheet : Excel.Worksheet) =
    [for r in 1 .. sheet.UsedRange.Rows.Count do
        yield getStringTuple (getCell (sheet.Cells, r, 1), getCell (sheet.Cells, r, 2))]
    
let sheetList =
    openSheets
    |> List.map (fun c -> (OutputType(c.Name, tupleList c)))

let json = JsonConvert.SerializeObject sheetList

[<EntryPoint>]
let main argv = 
    System.IO.File.WriteAllText("output.json", json);
    printfn "%s" json
    xl.Workbooks.Close() 
    0 // return an integer exit code
