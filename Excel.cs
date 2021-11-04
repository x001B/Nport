using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using static ETF.Helper;
using static ETF.Search;

namespace ETF
{
    public static class Excel
    {
        public static ExcelWorksheet Holdings { get; set; }
        public static ExcelWorksheet Notes { get; set; }
        public static ExcelWorksheet Debts { get; set; }
        public static ExcelWorksheet CurrencyInfos { get; set; }
        public static ExcelWorksheet Borrowings { get; set; }
        public static ExcelPackage Output { get; set; }
        public static int FileCount { get; set; } = -1;
        public static void Template()
        {
            if (FileCount == -1)
            {
                Output = new(CreateFile(FileCount));
                Console.WriteLine("\nExcel workbook created. Starting to write");
            }
            else
            {
                if (FileCount == 0)
                {
                    File.Move(GetFileName(FileCount - 1), GetFileName(FileCount));
                    FileCount += 1;
                }

                Output = new(CreateFile(FileCount));
                
                Console.WriteLine("\nAdditional Excel workbook created. Starting to write");
            }
            foreach (var styleName in new string[] { "Header", "Data", "Url", "Section", "Series" })
            {
                var style = Output.Workbook.Styles.CreateNamedStyle(styleName);
                switch (styleName)
                {
                    case "Header":
                        style.Style.Font.SetFromFont(new("Arial", 11f, FontStyle.Bold));
                        style.Style.Indent = 1;
                        style.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        style.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        break;
                    case "Data":
                        style.Style.Font.SetFromFont(new("Arial", 10f, FontStyle.Regular));
                        style.Style.Numberformat.Format = "###,###,###,###,##0.00000000";
                        break;
                    case "Url":
                        style.Style.Font.SetFromFont(new("Arial", 10f, FontStyle.Italic | FontStyle.Underline));
                        style.Style.Font.Color.SetColor(Color.Blue);
                        break;
                    case "Section":
                        style.Style.Font.SetFromFont(new("Arial", 11f, FontStyle.Regular | FontStyle.Underline));
                        break;
                    case "Series":
                        style.Style.Font.SetFromFont(new("Arial", 10f, FontStyle.Italic | FontStyle.Underline | FontStyle.Bold));
                        style.Style.Font.Color.SetColor(Color.DarkGray);
                        style.Style.Numberformat.Format = "yyyy-MM-dd";
                        style.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        break;
                    default:
                        break;
                }
            }
            Holdings = Output.Workbook.Worksheets.Add("Holdings");
            Borrowings = Output.Workbook.Worksheets.Add("Borrowings");
            Notes = Output.Workbook.Worksheets.Add("Notes");
            Debts = Output.Workbook.Worksheets.Add("Debts");
            CurrencyInfos = Output.Workbook.Worksheets.Add("CurrencyInfos");
            FileCount += 1;
        }
        public static void MakeExcel(List<(string rptDate, string fileDate, FormType formType, string url, 
            System.Xml.Linq.XElement xml, string regName, string seriesName)> filings)
        {
            Template();
            Holdings.FormatSheet(out int holdingsRow, Form.Headers());
            Borrowings.FormatSheet(out int borRow, null);
            var lastCol = Holdings.Dimension.End.Column;

            // Cycle through Nports and collect them so we can return link formulas correct
            // if we don't do this then we have to use formulas which capture entire columns & rows
            List<Nport> nports = new();
            var filingCounter = 0;
            var derivCounter = 0;
            List<object> com = new();
            using ProgressBar pb = new();

            foreach (var (_, fileDate, formType, nurl, x, _, seriesName) in filings)
            {
                pb.Report((double) filingCounter/filings.Count);
                var nport = new Nport(x, nurl, formType, fileDate);
                if (nport.HasDerivative)
                    derivCounter += 1;

                var init = new object[] { nurl, fileDate };
                com.AddRange(init);
                com.AddRange(nport.Form.GenInfo.ExcelValues());
                if (AddFundInfo)
                    com.AddRange(nport.Form.FundInfo.ExcelValues());

                var lentPosition = false;
                for (int i = 0; i < nport.PositionCount; i++)
                {
                    lentPosition = lentPosition || nport.Form.InvOrSecs[i].SecurityLending.LoanByFundCondition.LoanValue != default;
                    Holdings.Row(holdingsRow).StyleName = "Data";
                    Holdings.AddDataByForm(ref holdingsRow, com.Concat(nport.Form.InvOrSecs[i].ExcelValues()), formType);
                }

                com.Clear(); // "from an ipod to a cloud of gibsons"... that's a line in "if I only had a brain" from wiz of oz. what's idispose, alex?
                
                // If ever figure out will need to add who's borrowing back to InvOrSec
                if (lentPosition)
                {
                    ///////////////////////////////////
                    ///
                    //var col = 1;
                    //foreach (var (name, bal, val, loan) in nport.Form.Lendings)
                    //{
                    //    for (int i = 0; i < 4; i++)
                    //        Borrowings.Cells[borRow + i, col].Value = i switch
                    //        {
                    //            0 => name,
                    //            1 => loan,
                    //            2 => bal,
                    //            3 => val
                    //        };

                    //    col += 1;
                    //}
                    //col = 1;
                    //foreach (var b in nport.Form.FundInfo.Borrowers)
                    //{
                    //    Borrowings.Cells[borRow + 4, col].Value = b.Name;
                    //    Borrowings.Cells[borRow + 5, col].Value = b.Value;
                    //    col += 1;
                    //}

                    //borRow += 6;
                    ///////////////////////////////////
                    ///

                    Dictionary<string, decimal> borrowers = new();
                    List<KeyValuePair<string, decimal>> tempBorrowers = new(nport.Form.FundInfo.Borrowers.Select(x 
                        => new KeyValuePair<string, decimal>(x.Name, x.Value)).OrderBy(x => x.Key));

                    foreach (var kv in tempBorrowers)
                        if (borrowers.ContainsKey(kv.Key))
                            borrowers[kv.Key] += kv.Value;
                        else
                            borrowers.Add(kv.Key, kv.Value);

                    List<KeyValuePair<string, decimal>> lentPositions = nport.Form.Lendings.Select(x 
                        => new KeyValuePair<string, decimal>(x.title, x.amt)).OrderBy(x => x.Key).ToList();

                    Dictionary<string, Dictionary<string, decimal>> solve = new(borrowers.Select(x 
                        => new KeyValuePair<string, Dictionary<string, decimal>>(x.Key, new())));

                    foreach (var b in borrowers)
                        foreach (var lp in lentPositions)
                            if (!solve[b.Key].ContainsKey(lp.Key))
                                solve[b.Key].Add(lp.Key, default);

                    // match borrower to position if loanAmt == borrowAmt, then remove that position

                    foreach (var lp in nport.Form.Lendings)
                        if (borrowers.ContainsValue(lp.amt))
                        {
                            var bName = borrowers.First(x => x.Value == lp.amt).Key;
                            solve[bName][lp.title] += lp.amt;

                            // remove only, or first, instance of position where the security is the same as another position
                            lentPositions.RemoveAt(lentPositions.FindIndex(x => x.Key == lp.title));

                            // check if there's any position with same security that's being lent
                            if (!lentPositions.Any(x => x.Key == lp.title))
                                borrowers.Remove(bName);

                            for (int i = 0; i < nport.PositionCount; i++)
                                if (nport.Form.InvOrSecs[i].Title == lp.title)
                                    Holdings.AddValue(holdingsRow - nport.PositionCount + i, lastCol + 1, bName);
                            
                            if (lp.title.CheckAll())
                            {
                                if (borrowers.ContainsValue(lp.amt))
                                    Console.WriteLine($"\n{nport.SeriesName} ({nport.FileDate}): \nMultiple borrowers who borrowed: {lp.amt}" +
                                        $"\n{bName} assumed to borrow {lp.title}");

                                if (lentPositions.Any(x => x.Value == lp.amt))
                                    Console.WriteLine($"\n{nport.SeriesName} ({nport.FileDate}): \nMultiple positions lending: {lp.amt}" +
                                        $"\n{bName} assumed to borrow {lp.title}" +
                                        $"\n{lentPositions.First(x => x.Value == lp.amt).Key} assumed to borrow " +
                                        lentPositions.First(x => x.Value == lp.amt).Value);
                            }
                        }
                }

                x.Document.Root.Remove();
                nports.Add(nport);
                filingCounter += 1;
            }
            pb.Dispose(); // need to dispose here rather than let it self-dispose

            Console.WriteLine("\nAdded main holding sheet with values, need to do derivatives" +
                "\nStarting that now, creating formulas for hyperlinks");

            var noteCol = Form.GetHeaderIndex("Explanatory Notes");
            var debtInstCol = Form.GetHeaderIndex("Debt Instruments");
            var debtCurCol = Form.GetHeaderIndex("Debt Currency Infos");
            var comCols = new string[] { "URL", "% of Portfolio", "Return Link", "Series", "FileDate" };
            Notes.FormatSheet(out int noteRow, new string[] { "NPORT Item", "Note" });
            Debts.FormatSheet(out int debtRow, comCols.Concat(DebtSecRefInst.ExcelHeaders()));
            CurrencyInfos.FormatSheet(out int curRow, comCols.Concat(new string[] { "Conversion Ratio x 1000 Currency Units", "Currency" }));

            for (int i = 1; i < 3; i++)
                Debts.Column(i).Hidden = CurrencyInfos.Column(i).Hidden = true;

            holdingsRow = 2; // reset but need to skip over header row
            filingCounter = 0;
            using var pb2 = new ProgressBar();
            var gs = nports.GroupBy(x => x.Form.GenInfo.SeriesLei.ToString()).ToList();
            foreach (var g in gs)
                foreach (var nport in g.ToArray())
                {
                    pb2.Report((double)filingCounter / filings.Count);
                    var last = nport.Url == g.Last().Url;

                    ExcelWorksheet sheet = null;
                    var derRow = 1;
                    var hidden = new object[] { nport.Url, null };
                    var seriesSection = new object[] { nport.SeriesName, nport.FileDate };
                    foreach (var ios in nport.Form.InvOrSecs)
                    {
                        //if (last)
                        //    Holdings.Row(holdingsRow).Style.Font.Color.SetColor(Color.GreenYellow);

                        if (!ios.IsDerivative())
                        {
                            holdingsRow += 1;
                            continue;
                        }

                        var sname = nport.MakeSheetName();

                        if (Output.Workbook.Worksheets[sname] is null)
                            sheet = Output.Workbook.Worksheets.Add(sname);
                        else if (sheet is null)
                        {
                            sheet = Output.Workbook.Worksheets[sname];
                            derRow = sheet.Dimension.End.Row + 3;

                            using var er = sheet.Cells[derRow, 1];
                            er.StyleName = "Header";
                            er.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            switch (nport.Type)
                            {
                                case FormType.Ammended:
                                    er.Value = "Amended Filing Data";
                                    er.Style.Font.Color.SetColor(Color.Yellow);
                                    break;
                                case FormType.NotTimely:
                                    er.Value = "NotTimely Filing Data";
                                    er.Style.Font.Color.SetColor(Color.Orange);
                                    break;
                                case FormType.Repeat:
                                    er.Value = "Repeat Filing Data";
                                    er.Style.Font.Color.SetColor(Color.Red);
                                    break;
                                case FormType.Repeat | FormType.Regular:
                                    er.Value = "Repeat Regular Filing Data";
                                    er.Style.Font.Color.SetColor(Color.WhiteSmoke);
                                    er.Style.Fill.BackgroundColor.SetColor(Color.Red);
                                    break;
                                case FormType.Repeat | FormType.NotTimely:
                                    er.Value = "Repeat NotTimely Filing Data";
                                    er.Style.Font.Color.SetColor(Color.WhiteSmoke);
                                    er.Style.Fill.BackgroundColor.SetColor(Color.Orange);
                                    break;
                                case FormType.Repeat | FormType.Ammended:
                                    er.Value = "Repeat Ammended Filing Data";
                                    er.Style.Font.Color.SetColor(Color.WhiteSmoke);
                                    er.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                    break;
                                default:
                                    er.Value = $"{nport.Type.ToString("G")} Filing Data";
                                    er.Style.Font.Color.SetColor(Color.Red);
                                    er.Style.Fill.BackgroundColor.SetColor(Color.Black);
                                    break;
                            }
                            derRow += 1;
                        }

                        hidden[1] = ios.Percent;
                        if (ios.DebtSec is not null && ios.DebtSec.DebtSecRefInsts is not null)
                        {//GetReturnLink(holdingsRow, debtCol, "Go To Debt Instrument")

                            Holdings.AddHyperlink(holdingsRow, debtInstCol,
                                $"HYPERLINK(\"#{Debts.Cells[debtRow, 3].FullAddress}\", \"Go To Debt Inst\")");

                            Holdings.AddHyperlink(holdingsRow, debtCurCol,
                                $"HYPERLINK(\"#{CurrencyInfos.Cells[curRow, 3].FullAddress}\", \"Go To Debt Currency Infos\")");

                            foreach (var dsri in ios.DebtSec.DebtSecRefInsts)
                            {
                                Debts.AddValues(debtRow, hidden);

                                Debts.AddHyperlink(debtRow, hidden.Length + 1,
                                    Debts.GetReturnLink(debtRow, debtInstCol, ios.Title.ToString(), true));

                                Debts.AddValues(ref debtRow, seriesSection.Concat(dsri.ExcelValues()), hidden.Length + 2);
                            }

                            foreach (var (ratio, cur) in ios.DebtSec.CurrencyInfos)
                            {
                                CurrencyInfos.AddValues(curRow, hidden);

                                CurrencyInfos.AddHyperlink(curRow, hidden.Length + 1,
                                    CurrencyInfos.GetReturnLink(curRow, debtCurCol, ios.Title.ToString(), true));

                                CurrencyInfos.AddValues(ref curRow, seriesSection.Concat(new object[] { ratio, cur }), hidden.Length + 2);
                            }
                        }
                        if (ios.RepoAgreement is not null)
                        {
                            sheet.AddSection(ref derRow, hidden, seriesSection,
                                ios.RepoAgreement.ExcelHeaders(), ios.RepoAgreement.ExcelValues(), holdingsRow,
                                Form.GetHeaderIndex("Repo Agreement"), "Agreement", ios.Title);

                            if (ios.RepoAgreement.RepoCollaterals is not null) // this should never be null
                            {
                                derRow -= 1;
                                sheet.AddMultiLineSection(ref derRow, "Repo Collaterals",
                                    ios.RepoAgreement.RepoCollaterals);

                                derRow += 2;
                            }
                        }
                        if (ios.DerivativeHolding is not null)
                        {
                            var nestCount = 0;
                            sheet.AddSection(ref derRow, hidden, seriesSection,
                                ios.DerivativeHolding.ExcelHeaders(), ios.DerivativeHolding.ExcelValues(), holdingsRow,
                                Form.GetHeaderIndex("Derivative Holding"), "Derivative", ios.Title);

                            sheet.NestedDerivates(ref derRow, ref nestCount, ios.DerivativeHolding.GetRefInst());
                        }
                        holdingsRow += 1;
                    }
                    if (sheet is not null)
                    {
                        sheet.FormatSheet();
                        sheet.Column(1).Width = 34.5;
                        sheet.Column(2).Width = 34.5;
                        sheet.Column(3).Width = 34.5;
                        sheet.Cells[2, 4, derRow + 3, 50].AutoFitColumns();
                    }
                    if (nport.Form.ExplanatoryNotes is not null && nport.Form.InvOrSecs.Count != 0)
                    {
                        Notes.AddValues(ref noteRow, nport.Form.InvOrSecs.Select(x => x.Percent.ToString()).Prepend(nport.Url));
                        Notes.Row(noteRow - 1).Hidden = true;

                        holdingsRow -= nport.PositionCount;
                        for (int i = 0; i < nport.PositionCount; i++)
                            Holdings.AddHyperlink(ref holdingsRow, noteCol, 
                                $"HYPERLINK(\"#{Notes.Cells[noteRow, 3 + i].FullAddress}\", \"Go To Notes\")");

                        var os = seriesSection.Concat(Notes.GetReturnLinks(noteRow, noteCol, nport.Form.Titles));
                        Notes.AddValues(noteRow, os);
                        Notes.Cells[noteRow, 1, noteRow, 2].StyleName = "Series";
                        Notes.Cells[noteRow, 3, noteRow, os.Count()].StyleName = "Url";
                        Notes.Cells[noteRow, 3, noteRow, os.Count()].Style.Font.Size = 8f;
                        Notes.Cells[noteRow, 1, noteRow, os.Count()].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        noteRow += 1;

                        foreach (var (noteItem, note) in nport.Form.ExplanatoryNotes.OrderBy(x => x.Item1.ToString()))
                            Notes.AddStrings(ref noteRow, new string[] { noteItem.ToString(), note.ToString() }, "Data");

                        noteRow += 2;
                    }
                    filingCounter += 1;
                }
            pb2.Dispose(); // need to dispose here rather than let it self-dispose

            Notes.Column(1).AutoFit();
            Notes.Cells[3, 1].AutoFitColumns(); // this should always be a seriesName
            Notes.Cells[3, 2].AutoFitColumns(); // this should always be a date
            Notes.Cells[1, 3, Notes.Dimension.End.Row + 4, Notes.Dimension.End.Column].AutoFitColumns();
            Holdings.FormatColumns(Form.Headers(), holdingsRow);
            Output.Save();
            Console.WriteLine($"\nWritten to {Output.File}");
        }
        public static void NestedDerivates(this ExcelWorksheet sheet, ref int row, ref int nestCount, RefInst inst = null)
        {
            if (inst is not null)
            {
                nestCount += 1;

                if (inst.Basket is not null)
                {
                    sheet.AddSectionHeader(ref row, "Nested ".Times(nestCount) + "Index Basket");
                    sheet.AddHeader(ref row, Basket.ExcelHeaders());
                    sheet.AddValues(row, inst.Basket.ExcelValues(), "Data");
                    sheet.FormatRow(Basket.ExcelHeaders(), row);
                    row += 2;
                    if (inst.Basket.Components is not null)
                    {
                        sheet.AddHeader(ref row, Component.ExcelHeaders());
                        foreach (var comp in inst.Basket.Components.OrderBy(x => x.Name).ToArray())
                        {
                            sheet.AddValues(row, comp.ExcelValues(), "Data");
                            row += 1;
                        }
                    }
                    else
                        row += 1;

                    row += 1;
                }
                else
                {
                    if (inst.OtherRefInst is not null)
                        sheet.AddSectionHeader(ref row, "Nested ".Times(nestCount) + "Other Ref Instrument");
                    else
                        sheet.AddSectionHeader(ref row, "Nested ".Times(nestCount) + "Derivative");

                    sheet.AddHeader(ref row, inst.ExcelHeaders());
                    sheet.AddValues(row, inst.ExcelValues(), "Data");
                    sheet.FormatRow(inst.ExcelHeaders(), row);
                    row += 2;
                    if (inst.NestedDerivative is not null)
                        sheet.NestedDerivates(ref row, ref nestCount, inst.NestedDerivative.GetRefInst());
                    else
                        row += 2;
                }
            }
        }
        public static string GetRowsToSearch(int col) 
            => Holdings.Cells[Holdings.Dimension.Start.Row, col, Holdings.Dimension.End.Row, col].FullAddress;
    }
    public static class ExcelExtensions
    {
        public static void AddHyperlink(this ExcelWorksheet sheet, int row, int col, string hl)
        {
            sheet.Cells[row, col].Formula = hl;
            sheet.Cells[row, col].StyleName = "Url";
        }
        public static void AddHyperlink(this ExcelWorksheet sheet, ref int row, int col, string hl)
        {
            sheet.AddHyperlink(row, col, hl);
            row += 1;
        }
        public static void AddHyperlinks(this ExcelWorksheet sheet, int row, IEnumerable<string> hls, int col = 1)
        {
            for (int i = 0; i < hls.Count(); i++)
                sheet.AddHyperlink(row, col + i, hls.ElementAt(i));
        }
        public static void AddHyperlinks(this ExcelWorksheet sheet, ref int row, IEnumerable<string> hls, int col = 1)
        {
            sheet.AddHyperlinks(row, hls, col);
            row += 1;
        }
        public static void AddValue(this ExcelWorksheet sheet, int row, int col, object val)
        {
            switch (val)
            {
                case string s:
                    if (s.Contains("HYPERLINK"))
                    {
                        sheet.Cells[row, col].Formula = s;
                        sheet.Cells[row, col].StyleName = "Url";
                    }
                    else if (s.Length == 1 && int.TryParse(s, out int ti))
                        sheet.Cells[row, col].Value = ti;
                    else if (s.Length == 10 && s.Count(x => x == '-') == 2 && DateTime.TryParse(s, out DateTime dt))
                        sheet.Cells[row, col].Value = dt;
                    else if ((s.Contains(".") || !s.StartsWith("0")) && decimal.TryParse(s, out decimal tp))
                        sheet.Cells[row, col].Value = tp;
                    else
                        sheet.Cells[row, col].Value = s;
                    break;
                case decimal d:
                    sheet.Cells[row, col].Value = d;
                    break;
                default:
                    sheet.Cells[row, col].Value = val;
                    break;
            }
        }
        public static void AddString(this ExcelWorksheet sheet, int row, int col, string s) => sheet.Cells[row, col].Value = s;
        public static void AddString(this ExcelWorksheet sheet, ref int row, int col, string s)
        {
            sheet.AddString(row, col, s);
            row += 1;
        }
        public static void AddStrings(this ExcelWorksheet sheet, int row, IEnumerable<string> values, int col = 1)
        {
            for (int i = 0; i < values.Count(); i++)
                sheet.AddString(row, col + i, values.ElementAt(i));
        }
        public static void AddStrings(this ExcelWorksheet sheet, ref int row, IEnumerable<string> values, int col = 1)
        {
            sheet.AddStrings(row, values, col);
            row += 1;
        }
        public static void AddStrings(this ExcelWorksheet sheet, ref int row, IEnumerable<string> values, string style, int col = 1)
        {
            sheet.Cells[row, col, row, col + values.Count() - 1].StyleName = style;
            sheet.AddStrings(ref row, values, col);
        }
        public static void AddValues(this ExcelWorksheet sheet, int row, IEnumerable<object> values, int col = 1)
        {
            for (int i = 0; i < values.Count(); i++)
                sheet.AddValue(row, col + i, values.ElementAt(i));
        }
        public static void AddValues(this ExcelWorksheet sheet, ref int row, IEnumerable<object> values, int col = 1)
        {
            sheet.AddValues(row, values, col);
            row += 1;
        }
        public static void AddValues(this ExcelWorksheet sheet, int row, IEnumerable<object> values, string style)
        {
            sheet.Cells[row, 1, row, values.Count()].StyleName = style;
            sheet.AddValues(row, values);
        }

        public static void AddDataByForm(this ExcelWorksheet sheet, ref int row, IEnumerable<object> values, FormType ft)
        {
            if (ft != FormType.Regular)
            {
                if (ft.HasFlag(FormType.Repeat))
                {
                    sheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Row(row).Style.Fill.BackgroundColor.SetColor(
                        ft.HasFlag(FormType.NotTimely) ? Color.Orange :
                        ft.HasFlag(FormType.Ammended) ? Color.Yellow : Color.Red);
                    sheet.Row(row).Style.Font.Color.SetColor(Color.WhiteSmoke);
                }
                else
                    sheet.Row(row).Style.Font.Color.SetColor(
                        ft.HasFlag(FormType.NotTimely) ? Color.Orange :
                        ft.HasFlag(FormType.Ammended) ? Color.Yellow : Color.Red);
            }
            sheet.AddValues(ref row, values);
        }
        public static string GetReturnLink(this ExcelWorksheet sheet, int row, int retCol, string title, bool reverse = false) 
            => GetHyperlink(
                Form.GetHeaderIndex("URL"), sheet.Cells[reverse ? row : row - 1, 1].Address,
                Form.GetHeaderIndex("% of Portfolio"), sheet.Cells[reverse ? row : row - 1, 2].Address, retCol, title);
        public static IEnumerable<string> GetReturnLinks(this ExcelWorksheet sheet, int row, int hcol, IEnumerable<object> titles)
        {
            var list = new List<string>();
            for (int i = 0; i < titles.Count(); i++)
                list.Add(GetHyperlink(Form.GetHeaderIndex("URL"), sheet.Cells[row - 1, 1].Address, Form.GetHeaderIndex("% of Portfolio"), 
                    sheet.Cells[row - 1, i + 2].Address, hcol, titles.ElementAt(i).ToString()));
            return list;
        }
        public static string GetHyperlink(int returnUrlCol, string sheetUrlCell, int returnPctCol, string sheetPctCell, 
            int retCol, string retString = "Go Back") 
            => $"HYPERLINK(CHAR(35)&CELL(\"address\"," +
            $"INDEX({Excel.GetRowsToSearch(retCol)}," +
            $"SMALL(IF(COUNTIF({Excel.GetRowsToSearch(returnUrlCol)},{sheetUrlCell})," +
            $"MATCH({sheetPctCell},{Excel.GetRowsToSearch(returnPctCol)},0),\"\"),ROWS($A$1:A1)))),\"{retString}\")";
        public static void FormatSheet(this ExcelWorksheet sheet)
        {
            sheet.View.ShowGridLines = false;
            sheet.PrinterSettings.FitToPage = true;
        }
        public static void FormatSheet(this ExcelWorksheet sheet, out int row, IEnumerable<string> cols)
        {
            sheet.FormatSheet();
            row = 1;
            if (cols is not null)
            {
                sheet.AddHeader(ref row, cols);
                sheet.View.FreezePanes(2, 1);
            }
        }
        public static void FormatRange(this ExcelRange er, string col)
        {
            if (col.Contains('%'))
                er.Style.Numberformat.Format = "###,###,###,###,##0.00000000%";
            else if (col.Contains("Fair")) //cols.ElementAt(i).Contains("Days of Holding")
            {
                er.Style.Numberformat.Format = "###";
                er.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            else if (col.Contains("Date"))
            {
                er.Style.Numberformat.Format = "yyyy-MM-dd";
                er.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            else if (col.Contains("/") || col.Contains("Flag") ||
                col.Contains("Payoff") || col.Contains("Country") ||
                col.Contains("Coupon") || col.Contains("Cat") ||
                col.Contains("Cur"))
                er.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }
        public static void FormatColumns(this ExcelWorksheet sheet, IEnumerable<string> cols, int row)
        {
            for (int i = 1; i <= cols.Count(); i++)
                sheet.Cells[2, i, row, i].FormatRange(cols.ElementAt(i - 1));

            var table = sheet.Tables.Add(sheet.Cells[1, 1, row, cols.Count()], "Holdings");
            table.TableBorderStyle.BorderAround();
            table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium22;
            table.ShowRowStripes = true;
            table.Range.AutoFitColumns();
            sheet.Column(1).Hidden = true;
        }
        public static void FormatRow(this ExcelWorksheet sheet, IEnumerable<string> cols, int row)
        {
            for (int i = 1; i <= cols.Count(); i++)
                sheet.Cells[row, i].FormatRange(cols.ElementAt(i - 1));
        }
        public static void AddHeader(this ExcelWorksheet sheet, ref int row, IEnumerable<string> cols)
        {
            sheet.Cells[row, 1, row, cols.Count()].StyleName = "Header";

            for (int i = 1; i <= cols.Count(); i++)
                sheet.Cells[row, i].Value = cols.ElementAt(i - 1);

            row += 1;
        }
        public static void AddSection(this ExcelWorksheet sheet, ref int row, object[] hidden, object[] seriesValues, IEnumerable<string> hs, IEnumerable<object> vs,
            int hrow, int hcol, string s, object title)
        {
            sheet.AddValues(ref row, hidden);
            sheet.Row(row - 1).Hidden = true;
            Excel.Holdings.AddHyperlink(hrow, hcol, $"HYPERLINK(\"#{sheet.Cells[row, 3].FullAddress}\", \"Go To {s}\")");
            sheet.AddValues(row, seriesValues.Append(sheet.GetReturnLink(row, hcol, title.ToString())));
            sheet.Cells[row, 1, row, 2].StyleName = "Series";
            sheet.Cells[row, 3].StyleName = "Url";
            sheet.Cells[row, 3].Style.Font.Size = 8f;
            row += 1;
            //sheet.AddValues(ref row, hs, "Header");
            sheet.AddHeader(ref row, hs);
            sheet.AddValues(row, vs, "Data");
            sheet.FormatRow(hs, row);
            sheet.Cells[row - 1, 1, row, hs.Count()].AutoFitColumns();
            row += 2;
        }
        public static void AddSectionHeader(this ExcelWorksheet sheet, ref int row, string s)
        {
            row += 1;
            sheet.Cells[row, 1].Value = s;
            sheet.Cells[row, 1].StyleName = "Section";
            row += 1;
        }
        public static void AddMultiLineSection<T>(this ExcelWorksheet sheet, ref int row, string sHeader, List<T> list)
            where T : ICheck
        {
            sheet.AddSectionHeader(ref row, sHeader);
            sheet.AddHeader(ref row, list[0].ExcelHeaders());
            foreach (var ri in list)
            {
                sheet.AddValues(row, ri.ExcelValues(), "Data");
                sheet.FormatRow(list[0].ExcelHeaders(), row);
                row += 1;
            }
        }
    }
}
