// See https://aka.ms/new-console-template for more information
using GrapeCity.Documents.Excel;

Console.WriteLine("ワークシートのデータにアウトラインを追加して小計を計算する");

// 新規ワークブックの作成
var workbook = new GrapeCity.Documents.Excel.Workbook();

// テストデータを読み込み
workbook.Open("test-data.xlsx");

// 使用中の範囲を取得
var worksheet = workbook.Worksheets[0];
var usedrange = worksheet.UsedRange;

// アウトラインを追加、小計を計算
usedrange.Subtotal(1, ConsolidationFunction.Sum, new[] { 6 });

// 列幅を調整
usedrange.AutoFit();

//// グループ情報を取得、グループを折りたたみ
//var groupInfo = worksheet.Outline.RowGroupInfo;
//var rowInfo = new Dictionary<int, int>();

//foreach (var item in groupInfo)
//{
//    if (item.Children != null)
//    {
//        foreach (var childItem in item.Children)
//        {
//            childItem.Collapse();
//        }
//    }
//}

// Excelファイルに保存
workbook.Save("AddSubtotal.xlsx");

// 使用中の範囲を取得
var usedrange1 = worksheet.UsedRange;

// アウトラインをクリア
usedrange1.ClearOutline();

// Excelファイルに保存
workbook.Save("AddSubtotalClearOutline.xlsx");

//// 特定の行をグループ化
//var worksheet = workbook.Worksheets[0];
//worksheet.Range["8:13"].Group();
//worksheet.Range["20:23"].Group();

//// Excelファイルに保存
//workbook.Save("AddRowGroup.xlsx");
