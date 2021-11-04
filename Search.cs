using Flurl;
using Flurl.Http;
using Flurl.Http.Xml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using static ETF.Helper;
using static ETF.Excel;
using System.IO;

namespace ETF
{
    public partial class Search
    {
        public static void Conquer(RateLimiter rateLimiter, string term, DateTime? curStart = null, DateTime? curEnd = null)
        {
            curStart ??= Start;
            curEnd ??= End;

            var span = (TimeSpan)(curEnd - curStart);
            var mid = ((DateTime)curStart).AddDays(span.Days / 2);
            if (span.Days < 3)
            {
                SearchDates.Add(((DateTime)curStart, mid));
                SearchDates.Add((mid, (DateTime)curEnd));
                return;
            }

            if (!rateLimiter.IsRunning)
                rateLimiter.Start();

            var results = NportUpdates(term, 0, curStart, mid).GetAwaiter().GetResult();
            var num = (int)results.hits.total.value;
            Console.WriteLine("Split search date due to results being >=10,000\n\tUsing:" + ((DateTime)curStart).ToString("yyyy-MM-dd")
                + $" - {mid.ToString("yyyy-MM-dd")} = {num}");
            rateLimiter.Limit();
            if (num == 10_000)
                Conquer(rateLimiter, term, null, mid);
            else
            {
                SearchDates.Add(((DateTime)curStart, mid));

                results = NportUpdates(term, 0, mid).GetAwaiter().GetResult();
                num = (int)results.hits.total.value;
                rateLimiter.Limit();
                if (num == 10_000)
                    Conquer(rateLimiter, term, mid);
                else
                {
                    Console.WriteLine("Last split search date\n\tUsing:" + mid.ToString("yyyy-MM-dd")
                        + $" - {End.ToString("yyyy-MM-dd")} = {num}");
                    SearchDates.Add((mid, End));
                    rateLimiter.Purge();
                    rateLimiter.Stop();
                    if (SearchDates.First().start != Start && SearchDates.Last().end != End)
                        Console.WriteLine("Something in dividing up dates screwed up.\nManually split search dates");

                    return;
                }
            }
        }
        private static void DoSearch(string term, List<string> messages, ref int totalNumberOfResults)
        {
            // Search the SEC directory for any new files
            var results = NportUpdates(term).GetAwaiter().GetResult();

            // How many results from the search
            var numberOfResults = (int)results.hits.total.value;

            if (numberOfResults == 0)
            {
                messages.Add($"No results for \"{term}\" found on EDGAR");
                return;
            }
            else if (numberOfResults == 10_000)
            {
                if ((End - Start).Days < 2)
                    Console.WriteLine("Search term is way too general. Search something specific");
                else
                {
                    SearchDates = new();
                    Conquer(new RateLimiter(), term);
                }
                //var temp = new List<(DateTime start, DateTime end)>();
            }
            else
                totalNumberOfResults += numberOfResults;


            if (SearchDates is not null)
                foreach (var (start, end) in SearchDates)
                {
                    results = NportUpdates(term, 0, start, end).GetAwaiter().GetResult();
                    numberOfResults = (int)results.hits.total.value;
                    totalNumberOfResults += numberOfResults;

                    DoHeaderSearch(term, messages, results, numberOfResults);
                }
            else
                DoHeaderSearch(term, messages, results, numberOfResults);

            PrintMessages(messages);
            Console.WriteLine($"\n{Series.Count} out of {totalNumberOfResults} are useful");
        }
        private static void DoHeaderSearch(string term, List<string> messages, dynamic results, int numberOfResults)
        {
            if (numberOfResults > 100)
            {
                var pages = numberOfResults / 100;
                Console.WriteLine($"{pages + 1} pages of search results for \"{term}\", {numberOfResults} potential filings\nLet's get to whittlin'");
                using var pb = new ProgressBar();
                SearchFilingHeaders(results, messages);

                for (int i = 1; i <= pages; i++)
                {
                    pb.Report((double)i / pages);
                    results = NportUpdates(term, 100 * i).GetAwaiter().GetResult();

                    SearchFilingHeaders(results, messages);
                }
            }
            else
            {
                Console.WriteLine($"1 page of search results for \"{term}\", {numberOfResults} potential filings\nThis'll be quick");
                using var pb = new ProgressBar();
                SearchFilingHeaders(results, messages);
                pb.Report(1);
            }
        }
        private static void SearchFilingHeaders(dynamic results, List<string> messages)
        {
            RateLimiter rateLimiter = new();

            foreach (var item in results.hits.hits)
            {
                var form = GetFormType(((string)item._source.form).Trim());          // form will show NT or /A
                if (((string)item._source.file_type).Contains("EX"))    // root_form never shows EX | form doesn't always show EX
                    continue;

                if (item._source.ciks.Count > 1)                        // yet to experience this
                {
                    Console.WriteLine("more than one CIK");
                    Console.ReadKey();
                }

                var adsh = (string)item._source.adsh;
                headerUrl.AppendPathSegments(
                    new object[] { "Archives", "edgar", "data", item._source.ciks[0], adsh.Replace("-", ""), adsh + "-index-headers.html" });

                rateLimiter.Start();
                var response = headerUrl.GetStringAsync().GetAwaiter().GetResult();
                headerUrl.Url.Reset();
                if (response == "")                                     // yet to experience this
                {
                    rateLimiter.Limit();
                    Console.WriteLine("response is blank");
                    Console.ReadKey();
                    continue;
                }

                var seriesId = "";
                var rptDate = ((string)item._source.period_ending).Trim();
                var fileDate = ((string)item._source.file_date).Trim();
                var url = "";

                var arr = response.Split("\n");
                foreach (var line in arr)
                    if (line.Contains("SERIES-ID"))
                    {
                        seriesId = line.Trim().Remove(0, "<SERIES-ID>".Length);
                        break;
                    }

                XDocument x = null;
                archiveUrl.AppendPathSegments(new object[] { "Archives", "edgar", "data", item._source.ciks[0], adsh.Replace("-", ""), "primary_doc.xml" });
                url = archiveUrl.Url.ToString();

                if (seriesId == "")
                {
                    rateLimiter.Start();
                    try
                    {
                        x = archiveUrl.GetXDocumentAsync().GetAwaiter().GetResult();
                    }
                    catch (FlurlHttpException ex)
                    {
                        messages.Add("Error returned from " + ex.Call.Request.Url);
                    }

                    x.StripNamespace();

                    var gen = x.Element("edgarSubmission").Element("formData").Element("genInfo");
                    seriesId = gen.Element("seriesId")?.Value is not null ? (gen.Element("seriesId")?.Value) : gen.Element("seriesLei").Value;

                    if (seriesId == "N/A")
                    {
                        seriesId = gen.Element("seriesName").Value;
                        if (seriesId == "N/A")
                            seriesId = gen.Element("regName").Value.Acronize() + $"_N_{rptDate}";
                    }

                    rateLimiter.Limit();
                }
                archiveUrl.Url.Reset();

                if (!Series.ContainsKey(seriesId))
                    Series.Add(seriesId, new() { (rptDate, fileDate, form, url) });
                else if (!Series[seriesId].Any(x => x.url == url))
                {
                    var matches = Series[seriesId].Where(x => x.rptDate == rptDate).ToArray();
                    if (!matches.Any() || (matches.Length == 1 && form != matches[0].type))
                        Series[seriesId].Add((rptDate, fileDate, form, url)); // add new report date or, in theory, ammended file.
                    else if (matches.Length > 1)
                    {
                        Series[seriesId].Add((rptDate, fileDate, FormType.Repeat, url));
                        messages.Add($"Report Date has {(matches.Length + 1).ToString()} filings:" +
                            $"\n\tseriesId: {seriesId}\n\trptDate: {rptDate}\n\tfileDate: {fileDate}\n\turl: {url}" +
                            $"\n\tlast form type: {matches.Last().type.ToString("G")}, current: {form.ToString("G")}");
                    }
                    else
                    {
                        Series[seriesId].Add((rptDate, fileDate, FormType.Repeat | form, url));
                        messages.Add($"Repeat Filing ({(FormType.Repeat | matches[0].type).ToString("G")}):" +
                            $"\n\tseriesId: {seriesId}\n\trptDate: {rptDate}\n\tfileDate: {fileDate}\n\turl: {url}");
                    }
                }
                rateLimiter.Limit();
            }
            rateLimiter.Purge();
        }
        public class Form
        {
            public Form(XElement x, string url, out int positionCount, out bool hasDeriv)
            {
                positionCount = 0;
                hasDeriv = false;
                GenInfo = new(x.Element("genInfo"));
                FundInfo = new(x.Element("fundInfo"));
                Titles = new();
                if (x.Element("explntrNotes") is not null) // handle before holdings tacking them on to each holding is easy
                {
                    ExplanatoryNotes = new();
                    foreach (var exp in x.Element("explntrNotes").Descendants())
                        ExplanatoryNotes.Add((exp.Attribute("noteItem").Value, exp.Attribute("note").Value));
                }

                if (x.Element("invstOrSecs")?.Descendants("invstOrSec") is IEnumerable<XElement> holdings)
                {
                    InvOrSecs = new();
                    Lendings = new();
                    var unkLending = false;
                    
                    foreach (var holding in holdings.ToArray())
                    {
                        var positionFlag = false;
                        InvOrSec invOrSec = new(holding);

                        if (invOrSec.SecurityLending.LoanByFundCondition.LoanValue != default)
                        {
                            if (FundInfo.Borrowers is null && !unkLending)
                            {
                                Console.WriteLine($"{GenInfo.SeriesName} on {GenInfo.ReportPeriodFilingEndDate} is lending without borrowers?: {url}");
                                unkLending = true;
                            }

                            Lendings.Add((invOrSec.Name, invOrSec.Balance, invOrSec.Value, 
                                invOrSec.SecurityLending.LoanByFundCondition.LoanValue));
                        }

                        invOrSec.Check(ref positionFlag);

                        if (positionFlag || Testing)
                        {
                            InvOrSecs.Add(invOrSec);
                            Titles.Add(invOrSec.Title is null ? default : invOrSec.Title.ToString());
                            hasDeriv = hasDeriv != false || invOrSec.IsDerivative();
                            positionCount += 1;
                        }
                    }
                }
            }

            public GenInfo GenInfo { get; set; }
            public FundInfo FundInfo { get; set; }
            public List<InvOrSec> InvOrSecs { get; set; }
            public List<(string title, decimal bal, decimal val, decimal amt)> Lendings { get; set; }
            public List<(object, object)> ExplanatoryNotes { get; set; }
            public List<string> Titles { get; set; }
            private static List<string> _headers;
            public static int GetHeaderIndex(string s) => (_headers is null ? Headers() : _headers).IndexOf(s) + 1;
            public static List<string> Headers()
            {
                if (_headers is not null)
                    return _headers;

                _headers = new List<string>();

                _headers.AddRange(InitialHeaders);                  // These are the columns used to filter filings by their header txt file 
                _headers.AddRange(GenInfo.ExcelHeaders());
                if (AddFundInfo)
                    _headers.AddRange(FundInfo.ExcelHeaders());
                _headers.AddRange(InvOrSec.ExcelHeaders());

                return _headers;
            }
        }
        static async Task Main()
        {
            var messages = new List<string>();
            string[] args = null;
            var numberOfResults = 0;

            // Maybe handle other shit.. dunno, args ain't my bag
            if (args is not null)
                Canaries = args;

            Series = new();
            List<(string, string, FormType, string, XDocument)> nportXmls;
            List<(string rptDate, string fileDate, FormType formType, string url,
                    XElement xml, string regName, string seriesName)> sortedNportXmls;
            if (!Testing)
            {
                if (UseMultiSearch)
                {
                    Console.WriteLine($"Searching for these different terms: \"{string.Join("\", \"", SearchTerms)}\"");

                    foreach (var term in SearchTerms)
                        DoSearch(term, messages, ref numberOfResults);
                }
                else
                {
                    Console.WriteLine($"Searching for: \"{SearchTerm}\"");
                    DoSearch(SearchTerm, messages, ref numberOfResults);
                }

                if (numberOfResults == 0)
                {
                    PrintMessages(messages);
                    return;
                }

                // Create a list of XML files from results
                Console.WriteLine("Start pulling filings, this could take some time depending on filing count");
                nportXmls = await RetrieveNportXmls(messages);

                sortedNportXmls = new();
                foreach (var (rptDate, fileDate, formType, url, xml) in nportXmls)
                {
                    xml.StripNamespace();

                    var x = xml.Element("edgarSubmission");
                    var g = x.Element("formData").Element("genInfo");
                    sortedNportXmls.Add((rptDate, fileDate, formType, url, 
                        x, g.Element("regName").Value, g.Element("seriesName").Value));
                }
            }
            else
            {
                DirectoryInfo dir = new(Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.FullName + @"\SampleNports");
                sortedNportXmls = new();
                var formTypes = Enum.GetValues(typeof(FormType)).Cast<FormType>();
                var files = dir.GetFiles().OrderBy(x => x.Name).ToArray();
                for (int i = 0; i < files.Length; i++)
                {
                    var x = XDocument.Load(files[i].Open(FileMode.Open)).StripNamespace().Element("edgarSubmission");
                    var g = x.Element("formData").Element("genInfo");

                    // if more test files get added this may need revised
                    sortedNportXmls.Add((End.SecString(), End.SecString(), 
                        i > 3 ? formTypes.ElementAt(i % 3) | FormType.Repeat : formTypes.ElementAt(i), files[i].Name,
                        x, g.Element("regName").Value, g.Element("seriesName").Value));
                }
            }

            sortedNportXmls = sortedNportXmls
                .OrderBy(x => x.regName)
                .ThenBy(x => x.seriesName)
                .ThenBy(x => ((DateTimeOffset)DateTime.Parse(x.rptDate)).ToUnixTimeSeconds())
                .ThenBy(x => ((DateTimeOffset)DateTime.Parse(x.fileDate)).ToUnixTimeSeconds())
                .ThenBy(x => x.formType)
                .ToList();

            // Lay down the Excel whilst digesting the filings
            MakeExcel(sortedNportXmls);
        }
        private static void PrintMessages(List<string> messages)
        {
            for (int i = 0; i < messages.Count; i++)
            {
                if (i == 0)
                    Console.WriteLine("Comments on found filings below");
                Console.WriteLine(messages[i]);
                if (i + 1 == messages.Count)
                    Console.WriteLine('-'.Times(12) + "EOC");
            }

            messages.Clear();
        }
        private static async Task<List<(string, string, FormType, string, XDocument)>> RetrieveNportXmls(List<string> messages)
        {
            List<(string, string, FormType, string, XDocument)> filings = new();
            var rateLimiter = new RateLimiter();
            using var pb = new ProgressBar();
            for (int i = 0; i < Series.Count; i++)
            {
                pb.Report((double)i / Series.Count);
                foreach (var (rptDate, fileDate, formType, url) in Series.ElementAt(i).Value)
                {
                    archiveUrl.Url = Url.Parse(url);

                    rateLimiter.Start();
                    try
                    {
                        filings.Add((rptDate, fileDate, formType, url, await archiveUrl.GetXDocumentAsync()));
                    }
                    catch (FlurlHttpException ex)
                    {
                        messages.Add($"Error returned from {ex.Call.Request.Url} - Retrying ...");
                        try
                        {
                            filings.Add((rptDate, fileDate, formType, url, await archiveUrl.GetXDocumentAsync()));
                            messages.Add($"Success");
                        }
                        catch (FlurlHttpException exx)
                        {
                            messages.Add($"Failed again for: {exx.Call.Request.Url}");
                        }
                    }

                    archiveUrl.Url.Reset();
                    rateLimiter.Limit();
                }
            }
            return filings;
        }
        public struct Nport
        {
            public Nport(XElement x, string url, FormType type, string fileDate)
            {
                Url = url;
                Header = new(x.Element("headerData"), url);
                Form = new(x.Element("formData"), url, out int positionCount, out bool hasDeriv);
                PositionCount = positionCount;
                HasDerivative = hasDeriv;
                Type = type;
                FileDate = fileDate;
                SeriesName = Form.GenInfo.SeriesName;
            }
            public string Url { get; set; }
            public Header Header { get; set; }
            public Form Form { get; set; }
            public int PositionCount { get; set; }
            public bool HasDerivative { get; set; }
            public FormType Type { get; set; }
            public string FileDate { get; set; }
            public object SeriesName { get; set; }
            public string MakeSheetName()
            {
                var s = $"{Form.GenInfo.SeriesId ?? Form.GenInfo.RegName.ToString().Acronize() + '_' + SeriesName.ToString().Acronize()}";
                if (s.Length >= 20)
                    s = string.Join("", s.Take(20));

                return s + $"_{Form.GenInfo.ReportPeriodFilingEndDate}";
            }
        }
        public class FundInfo : ICheck
        {
            public FundInfo(XElement x)
            {
                TotalAssets = x.Element("totAssets").Value;
                TotalLiabilities = x.Element("totLiabs").Value;
                NetAssets = x.Element("netAssets").Value;
                NonPublicAssets = x.Element("assetsAttrMiscSec").Value;
                AssetsInvested = x.Element("assetsInvested").Value;
                AmountLessOneYrBank = x.Element("amtPayOneYrBanksBorr").Value;
                AmountLessOneYrCtrldComp = x.Element("amtPayOneYrCtrldComp").Value;
                AmountLessOneYrOthAffil = x.Element("amtPayOneYrOthAffil").Value;
                AmountLessOneYrOth = x.Element("amtPayOneYrOther").Value;
                AmountOneYrBank = x.Element("amtPayAftOneYrBanksBorr").Value;
                AmountOneYrCtrldComp = x.Element("amtPayAftOneYrCtrldComp").Value;
                AmountOneYrOthAffil = x.Element("amtPayAftOneYrOthAffil").Value;
                AmountOneYrOth = x.Element("amtPayAftOneYrOther").Value;
                DelayedDelivery = x.Element("delayDeliv").Value;
                StandbyDelivery = x.Element("standByCommit").Value;
                LiquidationPref = x.Element("liquidPref").Value;
                CashNotReportedInPartD = x.Element("cshNotRptdInCorD")?.Value;
                if (x.Element("borrowers") is not null)
                {
                    Borrowers = new();
                    foreach (var b in x.Element("borrowers").Descendants())
                        Borrowers.Add(new(b));
                }

                IsNonCashCollateral = x.Element("isNonCashCollateral")?.Value;
                if (IsNonCashCollateral == "Y")
                {
                    NonCashCollateralInfos = ("Y", new());
                    foreach (var collat in x.Element("aggregateCondition").Element("aggregateInfos").Descendants())
                        NonCashCollateralInfos.CollateralInfos.Add(new(collat));
                }

                Month1Flow = new(x.Element("mon1Flow"));
                Month2Flow = new(x.Element("mon2Flow"));
                Month3Flow = new(x.Element("mon3Flow"));
                MonthFlowsTotal = Month1Flow + Month2Flow + Month3Flow;
                LiquidInvestment = new(x.Element("liquidInvst"));

                var derivTrans = x.Element("derivTrans");
                if (derivTrans is not null)
                {
                    DerivativeTransactions = (derivTrans.Element("classification")?.Value, new());
                    if (DerivativeTransactions.Classification != "N/A" && derivTrans.Descendants() is not null)
                        foreach (var dTran in derivTrans.Descendants().ToArray())
                            DerivativeTransactions.DerivativeTransactions.Add((dTran.Attribute("classification").Value,
                                dTran.Attribute("fundPct").Value));
                }

            }
            public object TotalAssets { get; set; }
            public object TotalLiabilities { get; set; }
            public object NetAssets { get; set; }
            public object NonPublicAssets { get; set; }
            public object AssetsInvested { get; set; }
            public object AmountLessOneYrBank { get; set; }
            public object AmountLessOneYrCtrldComp { get; set; }
            public object AmountLessOneYrOthAffil { get; set; }
            public object AmountLessOneYrOth { get; set; }
            public object AmountOneYrBank { get; set; }
            public object AmountOneYrCtrldComp { get; set; }
            public object AmountOneYrOthAffil { get; set; }
            public object AmountOneYrOth { get; set; }
            public object DelayedDelivery { get; set; }
            public object StandbyDelivery { get; set; }
            public object LiquidationPref { get; set; }
            public object CashNotReportedInPartD { get; set; }
            public List<Borrower> Borrowers { get; set; }
            public string IsNonCashCollateral { get; set; }
            public (string IsNonCashCollateral, List<CollateralInfo> CollateralInfos) NonCashCollateralInfos { get; set; }
            public MonthFlow Month1Flow { get; set; }
            public MonthFlow Month2Flow { get; set; }
            public MonthFlow Month3Flow { get; set; }
            public MonthFlow MonthFlowsTotal { get; set; }
            public LiquidInvestment LiquidInvestment { get; set; }
            public (string Classification, List<(string Classification, string FundPct)> DerivativeTransactions) DerivativeTransactions { get; set; }
            public DerivExposure DerivExposure { get; set; }
            public VarInfo VarInfo { get; set; }
            public List<string> Dts { get; set; }
            public void Check(ref bool positionFlag)
            {
                if (Borrowers is not null && Borrowers.Where(x => x.Name.CheckAll()).Any())
                    positionFlag = true;

                if (NonCashCollateralInfos.CollateralInfos is not null)
                    foreach (var ci in NonCashCollateralInfos.CollateralInfos)
                        ci.Check(ref positionFlag);

                if (DerivativeTransactions.Classification is not null && DerivativeTransactions.Classification.CheckAll())
                    positionFlag = true;

                if (VarInfo is not null)
                    VarInfo.Check(ref positionFlag);
            }
            public IEnumerable<object> ExcelValues()
            {
                Dts = new();
                if (DerivativeTransactions.DerivativeTransactions is not null)
                    foreach (var (c, f) in DerivativeTransactions.DerivativeTransactions)
                        Dts.Add($"{c}: {f}%");

                var tmp = NonCashCollateralInfos.CollateralInfos is null ? null
                    : NonCashCollateralInfos.CollateralInfos.Select(x
                    => $"Amount: {x.Amount}\nCollat Val: {x.Collateral}\nCategory: {x.InvestmentCategory}\nInvCatCond: {x.InvestmentCategoryConditional.Cat}: {x.InvestmentCategoryConditional.Desc}");

                return new object[] { TotalAssets, TotalLiabilities, NetAssets, NonPublicAssets, AssetsInvested,
                    AmountLessOneYrBank, AmountLessOneYrCtrldComp, AmountLessOneYrOthAffil, AmountLessOneYrOth, AmountOneYrBank,
                    AmountOneYrCtrldComp, AmountOneYrOthAffil, AmountOneYrOth, DelayedDelivery, StandbyDelivery, LiquidationPref,
                    CashNotReportedInPartD, IsNonCashCollateral, NewlineValues(tmp), Month1Flow.ToString(), Month2Flow.ToString(), Month3Flow.ToString(),
                    MonthFlowsTotal.ToString()}.Concat(LiquidInvestment.ExcelValues()).Concat(new object[] { DerivativeTransactions.Classification, NewlineValues(Dts) }
                    .Concat(DerivExposure.ExcelValues()).Concat(VarInfo.ExcelValues()));
            }
            public static IEnumerable<string> ExcelHeaders() => new string[] { "Total Assets", "Total Liabilities", "Net Assets", "Non-Public Assets (Part D)", "Assets Invested",
                "Inv. Amt <1y in Bank", "Inv. Amt <1y in Foreign Ctrld Cmpny", "Inv. Amt <1y in Other Affil.", "Inv. Amt <1y in Other", "Inv. Amt >1y in Bank",
                "Inv. Amt >1y in Foreign Ctrld Cmpny", "Inv. Amt >1y in Other Affil.", "Inv. Amt >1y in Other", "Delayed Deliv.", "Standby Commit Deliv.", "Liquidation Pref.",
                "Cash Not Reported In Part D", "Non Cash Collateral (Y/N)", "Non Cash Collateral Infos", "Month 1 Flow", "Month 2 Flow", "Month 3 Flow",
                "Total Month Flow"}.Concat(LiquidInvestment.ExcelHeaders()).Concat(new string[] { "Deriv. Xactions Classication", "Deriv. Xactions" }
                .Concat(DerivExposure.ExcelHeaders()).Concat(VarInfo.ExcelHeaders()));

            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        public class DerivExposure : ICheck
        {
            public DerivExposure(XElement x)
            {
                ExposurePct = x?.Element("derivExposurePct")?.Value;
                CurrencyExposurePct = x?.Element("derivCurrencyExposurePct")?.Value;
                RateExposurePct = x?.Element("derivRateExposurePct")?.Value;
                ExcessiveBusinessDays = x?.Element("noOfBusinessDaysInExcess")?.Value;
            }
            public object ExposurePct { get; set; }
            public object CurrencyExposurePct { get; set; }
            public object RateExposurePct { get; set; }
            public object ExcessiveBusinessDays { get; set; }
            public void Check(ref bool positionFlag)
            {// Nothin to do
            }
            public static IEnumerable<string> ExcelHeaders()
                => new string[] { "Exposure %", "Currency Exposure %", "Rate Exposure %", "Excessive Bussiness Days" };
            public IEnumerable<object> ExcelValues()
                => new object[] { ExposurePct, CurrencyExposurePct, RateExposurePct, ExcessiveBusinessDays };
            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        public class VarInfo : ICheck
        {
            public VarInfo(XElement x)
            {
                MedianDailyVarPct = x.Element("medianDailyVarPct")?.Value;
                if (x?.Element("fundsDesignatedInfo") is not null)
                    FundDesignatedInfo = new(x.Element("fundsDesignatedInfo"));
                else
                    FundDesignatedInfo = new();
                BacktestingResults = x.Element("backtestingResults")?.Value;
            }
            public object MedianDailyVarPct { get; set; }
            public FundDesignatedInfo FundDesignatedInfo { get; set; }
            public object BacktestingResults { get; set; }
            public void Check(ref bool positionFlag) => FundDesignatedInfo.Check(ref positionFlag);
            public static IEnumerable<string> ExcelHeaders()
                => new string[] { "Median Daily Var %", "Backtesting Results" }.Concat(FundDesignatedInfo.ExcelHeaders());
            public IEnumerable<object> ExcelValues()
                => new object[] { MedianDailyVarPct, BacktestingResults }.Concat(FundDesignatedInfo.ExcelValues());
            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        public struct FundDesignatedInfo : ICheck
        {
            public FundDesignatedInfo(XElement x)
            {
                DesignatedIndexName = x.Element("nameDesignatedIndex").Value;
                IndexIdentifier = x.Element("indexIdentifier").Value;
                MedianVarRatioPct = x.Element("medianVarRatioPct").Value;
            }
            public string DesignatedIndexName { get; set; }
            public string IndexIdentifier { get; set; }
            public object MedianVarRatioPct { get; set; }

            public void Check(ref bool positionFlag)
            {
                if (DesignatedIndexName is not null && DesignatedIndexName.CheckAll())
                    positionFlag = true;

                if (IndexIdentifier is not null && IndexIdentifier.CheckAll())
                    positionFlag = true;
            }
            public static IEnumerable<string> ExcelHeaders()
                => new string[] { "Designated Index Name", "Index Identifier", "Median Var Ratio %" };
            public IEnumerable<object> ExcelValues()
                => new object[] { DesignatedIndexName, IndexIdentifier, MedianVarRatioPct };
            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        public class LiquidInvestment : ICheck
        {
            public LiquidInvestment(XElement x)
            {
                HighlyLiquid = x?.Element("highlyLiquidInvst")?.Value;
                DaysOfHolding = x?.Element("daysOfHolding")?.Value;
                IsReportingPeriodChanged = x?.Element("isChangeRepPd")?.Value;
                var liquidCond = x?.Element("liquidInvConditional");
                if (liquidCond is not null)
                {
                    ReportingPeriodChangeConditional = (liquidCond.Element("isChangeRepPd").Value, new());
                    foreach (var rptPdChange in liquidCond.Element("rptPdChanges").Descendants("rptPdChange"))
                        ReportingPeriodChangeConditional.ReportingPeriodChanges.Add(rptPdChange.Value);
                }

            }
            public object HighlyLiquid { get; set; }
            public object DaysOfHolding { get; set; }
            public object IsReportingPeriodChanged { get; set; }
            public (object IsReportingPeriodChanged, List<object> ReportingPeriodChanges) ReportingPeriodChangeConditional { get; set; }
            public void Check(ref bool positionFlag)
            {// nothing to do here
            }
            public IEnumerable<object> ExcelValues() => new object[] {HighlyLiquid, DaysOfHolding, IsReportingPeriodChanged,
                $"{ReportingPeriodChangeConditional.IsReportingPeriodChanged}: {NewlineValues(ReportingPeriodChangeConditional.ReportingPeriodChanges)}"
            };
            public static IEnumerable<string> ExcelHeaders() => new string[] {
                "Highly Liquid", "Days of Holding", "Reporting Period Changed", "Reporting Period Change Conditional"
            };
            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        public struct MonthFlow
        {
            public MonthFlow(XElement x)
            {
                Sales = x.Attribute("sales").Value;
                Reinvestment = x.Attribute("reinvestment").Value;
                Redemption = x.Attribute("redemption").Value;
            }
            public object Sales { get; set; }
            public object Reinvestment { get; set; }
            public object Redemption { get; set; }
            public static MonthFlow operator +(MonthFlow mf, MonthFlow mf2)
            {
                var tmf = new MonthFlow();
                if (mf.Sales.ToString() == "N/A")
                    tmf.Sales = "N/A";
                else
                    tmf.Sales = mf.Sales;

                if (mf2.Sales.ToString() != "N/A")
                    if (tmf.Sales.ToString() != "N/A")
                        tmf.Sales = (tmf.Sales is decimal t ? t : decimal.Parse((string)tmf.Sales))
                            + (mf2.Sales is decimal m ? m : decimal.Parse((string)mf2.Sales));
                    else
                        tmf.Sales = mf2.Sales;

                if (mf.Reinvestment.ToString() == "N/A")
                    tmf.Reinvestment = "N/A";
                else
                    tmf.Reinvestment = mf.Reinvestment;

                if (mf2.Reinvestment.ToString() != "N/A")
                    if (tmf.Reinvestment.ToString() != "N/A")
                        tmf.Reinvestment = (tmf.Reinvestment is decimal t ? t : decimal.Parse((string)tmf.Reinvestment))
                            + (mf2.Reinvestment is decimal m ? m : decimal.Parse((string)mf2.Reinvestment));
                    else
                        tmf.Reinvestment = mf2.Reinvestment;

                if (mf.Redemption.ToString() == "N/A")
                    tmf.Redemption = "N/A";
                else
                    tmf.Redemption = mf.Redemption;

                if (mf2.Redemption.ToString() != "N/A")
                    if (tmf.Redemption.ToString() != "N/A")
                        tmf.Redemption = (tmf.Redemption is decimal t ? t : decimal.Parse((string)tmf.Redemption))
                            + (mf2.Redemption is decimal m ? m : decimal.Parse((string)mf2.Redemption));
                    else
                        tmf.Redemption = mf2.Redemption;

                return tmf;
            }
            public new string ToString() => $"Sales: {Sales}\nReinvestment: {Reinvestment}\nRedemption: {Redemption}";
        }
        public struct Borrower
        {
            public Borrower(XElement x)
            {
                Name = x.Attribute("name").Value;
                Lei = x.Attribute("lei").Value;
                Value = decimal.Parse(x.Attribute("aggrVal").Value);
            }
            public string Name { get; set; }
            public string Lei { get; set; }
            public decimal Value { get; set; }
        }
        public class CollateralInfo : ICheck
        {
            public CollateralInfo(XElement x)
            {
                Amount = x.Attribute("amt").Value;
                Collateral = x.Attribute("collatrl").Value;
                InvestmentCategory = x.Element("invstCat").Value;
                if (InvestmentCategory is null)
                    InvestmentCategoryConditional = (x.Element("invstCatConditional").Attribute("invstCat").Value,
                        x.Element("invstCatConditional").Attribute("otherDesc").Value);
            }
            public string Amount { get; set; }
            public string Collateral { get; set; }
            public object InvestmentCategory { get; set; }
            public (string Cat, string Desc) InvestmentCategoryConditional { get; set; }
            public void Check(ref bool positionFlag)
            {
                if (InvestmentCategoryConditional.Desc is not null && InvestmentCategoryConditional.Desc.CheckAll())
                    positionFlag = true;
            }
            public IEnumerable<object> ExcelValues() => new object[] {
                Amount, Collateral, InvestmentCategory ?? InvestmentCategoryConditional.Cat,
                InvestmentCategoryConditional.Desc ?? InvCats[InvestmentCategory.ToString()]
            };
            public static IEnumerable<string> ExcelHeaders() => new string[] {
                "Amount", "Collateral", "Investment Cat.", "Investment Cat. Desc."
            };
            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        [NportItem("A")]
        public struct GenInfo : ICheck
        {
            public GenInfo(XElement x)
            {
                RegName = x.Element("regName").Value;
                RegCik = x.Element("regCik").Value;
                RegLei = x.Element("regLei").Value;
                SeriesName = x.Element("seriesName").Value;
                SeriesId = x.Element("seriesId")?.Value;
                SeriesLei = x.Element("seriesLei").Value;
                ReportPeriodFiscalEndDate = x.Element("repPdEnd")?.Value;
                ReportPeriodFilingEndDate = x.Element("repPdDate")?.Value;
            }
            [NportItem("A.1.a")]
            public object RegName { get; set; }
            //[NportItem("A.1.b")]
            [NportItem("A.1.c")]
            public object RegCik { get; set; }
            [NportItem("A.1.d")]
            //[NportItem("A.1.e")]
            public object RegLei { get; set; }
            [NportItem("A.2.a")]
            public object SeriesName { get; set; }
            [NportItem("A.2.b")]
            public object SeriesId { get; set; }
            [NportItem("A.2.c")]
            public object SeriesLei { get; set; }
            [NportItem("A.3.a")]
            public object ReportPeriodFiscalEndDate { get; set; }
            [NportItem("A.3.b")]
            public object ReportPeriodFilingEndDate { get; set; }
            //[NportItem("A.4")]
            public IEnumerable<object> ExcelValues() => new object[] { RegName, RegCik, RegLei, SeriesName, SeriesId, SeriesLei,
                    ReportPeriodFiscalEndDate, ReportPeriodFilingEndDate };
            public static IEnumerable<string> ExcelHeaders() => new string[] { "Registerer Name", "Registerer CIK", "Registerer LEI", "Series Name", "Series Id", "Series LEI",
                "Report Period Fiscal End Date", "Report Period Filind End Date" };
            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
            public void Check(ref bool positionFlag)
            {
                if (RegNameCanary is not null && RegName is not null && RegName.ToString().CheckId(RegNameCanary))
                    positionFlag = true;

                if (RegCikCanary is not null && RegCik is not null && RegCik.ToString().CheckId(RegCikCanary))
                    positionFlag = true;

                if (RegLeiCanary is not null && RegLei is not null && RegLei.ToString().CheckId(RegLeiCanary))
                    positionFlag = true;

                if (SeriesNameCanary is not null && SeriesName is not null && SeriesName.ToString().CheckId(SeriesNameCanary))
                    positionFlag = true;

                if (SeriesIdCanary is not null && SeriesId is not null && SeriesId.ToString().CheckId(SeriesIdCanary))
                    positionFlag = true;

                if (SeriesLeiCanary is not null && SeriesLei is not null && SeriesLei.ToString().CheckId(SeriesLeiCanary))
                    positionFlag = true;
            }
        }
        public class Header
        {
            public Header(XElement x, string url)
            {
                SubmissionType = x.Element("submissionType").Value;
                IsConfidential = x.Element("isConfidential").Value;
                AccessionNumber = x.Element("accessionNumber")?.Value;

                if (bool.Parse((string)IsConfidential) == true)
                    Console.WriteLine($"{url} is Confidential");

                FilerInfo = new(x.Element("filerInfo"), url);
            }
            public object SubmissionType { get; set; }
            public object IsConfidential { get; set; }
            public object AccessionNumber { get; set; }
            public FilerInfo FilerInfo { get; set; }
        }
        public class FilerInfo
        {
            public FilerInfo(XElement x, string url)
            {
                var lot = x.Element("liveTestFlag")?.Value;
                if (lot == "TEST")
                    Console.WriteLine($"{url} is {lot}");

                var sci = x.Element("seriesClassInfo");
                if (sci is not null)
                {
                    SeriesId = sci.Element("seriesId")?.Value;
                    var cids = sci.Elements("classId");
                    if (cids is not null)
                    {
                        ClassIds = new();

                        foreach (var cid in cids)
                            ClassIds.Add(cid.Value);
                    }
                }

            }
            public object SeriesId { get; set; }
            public List<object> ClassIds { get; set; }
        }
        public class InvOrSec : SecInfo
        {
            public InvOrSec(XElement x) : base(x)
            {
                PayoffProfile = x.Element("payoffProfile").Value;
                IsRestricted = x.Element("isRestrictedSec").Value;
                FundCategory = x.Element("fundCat")?.Value;
                FundCategories = new(new(), new());
                if (x.Element("fundCats") is not null)
                {
                    foreach (var fcat in x.Element("fundCats").Descendants("fundCat"))
                        FundCategories.FundCategories.Add((fcat.Attribute("category")?.Value, fcat.Attribute("pct").Value));

                    var fundCatCircs = x.Element("fundCats").Descendants("circumstances");
                    if (fundCatCircs is not null)
                        foreach (var circ in fundCatCircs)
                            FundCategories.Circumstances.Add(circ.Value);
                }
                FairValueLevel = x.Element("fairValLevel")?.Value;
                DebtSec = x.Element("debtSec") is not null ?
                    new(x.Element("debtSec")) : null;
                RepoAgreement = x.Element("repurchaseAgrmt") is not null ?
                    new(x.Element("repurchaseAgrmt")) : null;
                DerivativeHolding = x.Element("derivativeInfo") is not null ?
                    new(x.Element("derivativeInfo").Descendants().First()) : null;

                if (x.Element("securityLending") is not null)
                    SecurityLending = new(x.Element("securityLending"));
            }
            public object PayoffProfile { get; set; }
            public object IsRestricted { get; set; }
            public object FundCategory { get; set; }
            public (List<(object FundCategory, object Percent)> FundCategories, List<object> Circumstances) FundCategories { get; set; }
            public object FairValueLevel { get; set; }
            public DebtSec DebtSec { get; set; }
            public RepoAgreement RepoAgreement { get; set; }
            public DerivativeHolding DerivativeHolding { get; set; }
            public SecurityLending SecurityLending { get; set; }
            public List<string> Fcs { get; set; }
            public List<string> Crs { get; set; }
            public bool IsDerivative() => RepoAgreement is not null || DerivativeHolding is not null;
            public new void Check(ref bool positionFlag)
            {
                base.Check(ref positionFlag);
                Fcs = new();
                Crs = new();

                if (FundCategories.FundCategories is not null)
                    foreach (var (cat, per) in FundCategories.FundCategories)
                        Fcs.Add($"{cat}: {per}%");

                if (FundCategories.Circumstances is not null)
                    foreach (var circ in FundCategories.Circumstances)
                        Crs.Add(circ.ToString());

                foreach (var fc in Fcs)
                    if (fc.CheckAll())
                        positionFlag = true;

                foreach (var cr in Crs)
                    if (cr.CheckAll())
                        positionFlag = true;

                if (DebtSec is not null)
                    DebtSec.Check(ref positionFlag);

                if (RepoAgreement is not null)
                    RepoAgreement.Check(ref positionFlag);

                if (DerivativeHolding is not null)
                    DerivativeHolding.Check(ref positionFlag);
            }
            public new IEnumerable<object> ExcelValues()
            {
                var secLending = SecurityLending.ExcelValues().ToList();
                decimal? percent = null;
                if (secLending[2] is not null && (decimal.Parse(Value.ToString()) != 0)) // Loan Value
                    percent = decimal.Parse(secLending[2].ToString()) / decimal.Parse(Value.ToString());

                secLending.Add(percent);
                return base.ExcelValues().Concat(new object[] {
                    PayoffProfile, IsRestricted, FundCategory, NewlineValues(Fcs), NewlineValues(Crs), FairValueLevel,
                    null, null, null }).Concat(secLending) // 3 nulls = RepoAgreement, Derivative, Note
                    .Concat(DebtSec is null ? Array.Empty<object>() : DebtSec.ExcelValues());
            }
            public new static IEnumerable<string> ExcelHeaders() => SecInfo.ExcelHeaders().Concat(new string[] {
                "Payoff Profile", "Restricted (Y/N)", "Fund Cat.", "Fund Categories", "FC Circumstances", "Fair Value Level",
                "Repo Agreement", "Derivative Holding", "Explanatory Notes"}).Concat(SecurityLending.ExcelHeaders()
                .Concat(DebtSec.ExcelHeaders()));
        }
        public class SecInfo : Instrument, ICheck
        {
            public SecInfo(XElement x) : base(x)
            {
                Lei = x.Element("lei")?.Value ?? default;
                Cusip = x.Element("cusip")?.Value ?? default;
                Balance = decimal.TryParse(x.Element("balance")?.Value, out decimal _d) ? _d : default;
                Units = x.Element("units")?.Value;
                UnitDesc = x.Element("descOthUnits")?.Value;
                Currency = x.Element("curCd")?.Value;
                if (Currency is null)
                    CurrencyConditional = (x.Element("currencyConditional").Attribute("curCd").Value,
                        x.Element("currencyConditional").Attribute("exchangeRt").Value);

                Value = decimal.TryParse(x.Element("valUSD")?.Value, out _d) ? _d : default;
                Percent = decimal.TryParse(x.Element("pctVal")?.Value, out _d) ? _d / 100m : default;
                AssetCategory = x.Element("assetCat")?.Value;
                if (AssetCategory is null)
                    AssetConditional = (x.Element("assetConditional").Attribute("assetCat").Value,
                        x.Element("assetConditional").Attribute("desc").Value);
                IssuerCategory = x.Element("issuerCat")?.Value;
                if (IssuerCategory is null)
                    IssuerConditional = (x.Element("issuerConditional").Attribute("issuerCat").Value,
                        x.Element("issuerConditional").Attribute("desc").Value);
                Country = x.Element("invCountry")?.Value;
                OtherCountry = x.Element("invOthCountry")?.Value;
            }
            public string Lei { get; set; }
            public string Cusip { get; set; }
            public decimal Balance { get; set; }
            public object Units { get; set; }
            public object UnitDesc { get; set; }
            public object Currency { get; set; }
            public (object Currency, object ExchangeRate) CurrencyConditional { get; set; }
            public decimal Value { get; set; }
            public decimal Percent { get; set; }
            public object AssetCategory { get; set; }
            public (string AssetCategory, string Desc) AssetConditional { get; set; }
            public object IssuerCategory { get; set; }
            public (string IssuerCategory, string Desc) IssuerConditional { get; set; }
            public object Country { get; set; }
            public object OtherCountry { get; set; }
            public new void Check(ref bool positionFlag)
            {
                base.Check(ref positionFlag);
                if (Lei.CheckId(LeiCanary))
                    positionFlag = true;

                if (Cusip.CheckId(CusipCanary))
                    positionFlag = true;

                if (AssCatCanary is not null && AssetCategory is not null 
                    && AssetCategory.ToString().CheckId(AssCatCanary))
                    positionFlag = true;

                if (IssCatCanary is not null &&
                    IssuerCategory is not null && IssuerCategory.ToString().CheckId(IssCatCanary))
                    positionFlag = true;

                if (UnitDesc is not null && ((string)UnitDesc).CheckAll())
                    positionFlag = true;

                if (AssetConditional.Desc is not null && AssetConditional.Desc.CheckAll())
                    positionFlag = true;

                if (IssuerConditional.Desc is not null && IssuerConditional.Desc.CheckAll())
                    positionFlag = true;
            }
            public new IEnumerable<object> ExcelValues()
            {
                return base.ExcelValues().Concat(new object[] {
                    Lei, Cusip, Balance, Units, UnitDesc ?? UnitDescs[Units.ToString()], Currency ?? CurrencyConditional.Currency, CurrencyConditional.ExchangeRate,
                    Value, Percent, AssetCategory ?? AssetConditional.AssetCategory, AssetConditional.Desc ?? AssCats[AssetCategory.ToString()],
                    IssuerCategory ?? IssuerConditional.IssuerCategory, IssuerConditional.Desc ?? IssCats[IssuerCategory.ToString()], Country, OtherCountry
                });
            }

            public new static IEnumerable<string> ExcelHeaders() => Instrument.ExcelHeaders().Concat(new string[] {
                "LEI", "CUSIP", "Balance", "Units", "Unit Desc.", "Currency", "USD Exchange Rate",
                "Value", "% of Portfolio", "Asset Cat.", "Asset Desc.",
                "Issuer Cat.", "Issuer Desc.", "Country", "Other Country"
            });

            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        public class DebtSecRefInst : Instrument, ICheck
        {
            public DebtSecRefInst(XElement x) : base(x) => Currency = x.Element("curCd").Value;
            public object Currency { get; set; }
            public new IEnumerable<object> ExcelValues() => base.ExcelValues().Append(Currency);
            public new static IEnumerable<string> ExcelHeaders() => Instrument.ExcelHeaders().Append("Currency");
        }
        public class DebtSec : ICheck
        {
            public DebtSec(XElement x)
            {
                MaturityDate = x.Element("maturityDt")?.Value;
                CouponKind = x.Element("couponKind").Value;
                AnnualizedRate = x.Element("annualizedRt").Value;
                IsDefault = x.Element("isDefault").Value;
                AreInterestPaymentsInArrears = x.Element("areIntrstPmntsInArrs").Value;
                IsPaidKind = x.Element("isPaidKind").Value;
                IsMandatoryConvert = x.Element("isMandatoryConvrtbl")?.Value;
                if (IsMandatoryConvert is not null)
                {
                    IsContingentConvert = x.Element("isContngtConvrtbl").Value;
                    DebtSecRefInsts = new();

                    // Descendants doesn't work here to get the nested dbtSecRefInstrument ??? but why ???
                    foreach (var inst in x.Element("dbtSecRefInstruments").Elements("dbtSecRefInstrument"))
                        DebtSecRefInsts.Add(new(inst));

                    CurrencyInfos = new();
                    foreach (var curr in x.Element("currencyInfos").Descendants())
                        CurrencyInfos.Add((curr.Attribute("convRatio").Value, curr.Attribute("curCd").Value));

                    Delta = x.Element("delta").Value;
                }
            }
            public object MaturityDate { get; set; }
            public object CouponKind { get; set; }
            public object AnnualizedRate { get; set; }
            public object IsDefault { get; set; }
            public object AreInterestPaymentsInArrears { get; set; }
            public object IsPaidKind { get; set; }
            public object IsMandatoryConvert { get; set; } // the fields below are dependent on this being here but wastes to many sheets to do conditional splitting
            public object IsContingentConvert { get; set; }
            public List<DebtSecRefInst> DebtSecRefInsts { get; set; }
            public List<(string ConversionRatio, string Currency)> CurrencyInfos { get; set; }
            public object Delta { get; set; }
            public void Check(ref bool positionFlag)
            {
                if (DebtSecRefInsts is not null)
                    foreach (var dsri in DebtSecRefInsts)
                        dsri.Check(ref positionFlag);
            }
            public IEnumerable<object> ExcelValues()
            {
                //if (IsMandatoryConvert is not null)
                    return new object[] {
                        MaturityDate, CouponKind, AnnualizedRate, IsDefault, AreInterestPaymentsInArrears, IsPaidKind,
                        IsMandatoryConvert, IsContingentConvert, Delta
                    };
                //else
                //    return new object[] {
                //        MaturityDate, CouponKind, AnnualizedRate, IsDefault, AreInterestPaymentsInArrears, IsPaidKind
                //    };
            }
            public static IEnumerable<string> ExcelHeaders()
            {
                //if (IsMandatoryConvert is not null)
                    return new string[] {
                        "Maturity Date", "Coupon Kind", "Annualized Rate", "In Default (Y/N)", "Interest Payments Past Due (Y/N)", "Paid in Kind (Y/N)",
                        "Mandatory Convert (Y/N)", "Contingent Convert (Y/N)", "Delta", "Debt Instruments", "Debt Currency Infos"
                    };
                //else
                //    return new string[] { "Maturity Date", "Coupon Kind", "Annualized Rate", "In Default (Y/N)", "Interest Payments in Arrears (Y/N)", "Paid in Kind (Y/N)" };
            }
            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        public class RepoAgreement : ICheck
        {
            public RepoAgreement(XElement x)
            {
                TransactionCategory = x.Element("transCat").Value;
                if (x.Element("clearedCentCparty") is not null)
                    ClearedCentralParty = (x.Element("clearedCentCparty").Attribute("isCleared").Value,
                        x.Element("clearedCentCparty").Attribute("centralCounterparty").Value);
                else
                {
                    NotClearedCentralParty = (x.Element("notClearedCentCparty").Attribute("isCleared").Value, new());
                    foreach (var info in x.Element("notClearedCentCparty").Element("counterpartyInfos").Descendants())
                        NotClearedCentralParty.Counterparties.Add(new(info, true));
                }

                IsTriParty = x.Element("isTriParty").Value;
                RepoRate = x.Element("repurchaseRt").Value;
                MaturityDate = x.Element("maturityDt").Value;
                RepoCollaterals = new();
                foreach (var rc in x.Element("repurchaseCollaterals").Descendants("repurchaseCollateral"))
                    RepoCollaterals.Add(new(rc));
            }
            public object TransactionCategory { get; set; }
            public (string IsCleared, string CentralCounterparty) ClearedCentralParty { get; set; }
            public (string IsCleared, List<Counterparty> Counterparties) NotClearedCentralParty { get; set; }
            public object IsTriParty { get; set; }
            public object RepoRate { get; set; }
            public object MaturityDate { get; set; }
            public List<RepoCollateral> RepoCollaterals { get; set; }
            public void Check(ref bool positionFlag)
            {
                if (RepoCollaterals is not null)
                    foreach (var rc in RepoCollaterals)
                        if (rc.InvestmentCategoryConditional.Desc is not null && rc.InvestmentCategoryConditional.Desc.CheckAll())
                            positionFlag = true;
            }
            public IEnumerable<object> ExcelValues()
            {
                var cps = new List<string>();
                if (NotClearedCentralParty.Counterparties is not null)
                    foreach (var cp in NotClearedCentralParty.Counterparties)
                        cps.Add($"{cp.Name} (lei: {cp.Lei})");

                return new object[] {
                    TransactionCategory, ClearedCentralParty.IsCleared ?? NotClearedCentralParty.IsCleared,
                    ClearedCentralParty.CentralCounterparty ?? NewlineValues(cps), IsTriParty, RepoRate, MaturityDate
                };
            }
            public IEnumerable<string> ExcelHeaders() => new string[] {
                "Transaction Cat.", "Cleared Central Party (Y/N)", "Central Party", "Tri Party (Y/N)", "Repo Rate", "Maturity Date"
            };
        }
        public struct RepoCollateral : ICheck
        {
            public RepoCollateral(XElement x)
            {
                PrincipalAmount = x.Element("principalAmt").Value;
                PrincipalCurrency = x.Element("principalCd").Value;
                CollateralValue = x.Element("collateralVal").Value;
                CollateralCurrency = x.Element("collateralCd").Value;
                InvestmentCategory = x.Element("invstCat")?.Value;
                if (InvestmentCategory is null)
                    InvestmentCategoryConditional = (x.Element("invstCatConditional").Attribute("invstCat").Value,
                        x.Element("invstCatConditional").Attribute("desc").Value);
                else
                    InvestmentCategoryConditional = (null, null);
            }
            public object PrincipalAmount { get; set; }
            public object PrincipalCurrency { get; set; }
            public object CollateralValue { get; set; }
            public object CollateralCurrency { get; set; }
            public object InvestmentCategory { get; set; }
            public (string InvestmentCategory, string Desc) InvestmentCategoryConditional { get; set; }
            public IEnumerable<object> ExcelValues() => new object[] { PrincipalAmount, PrincipalCurrency, CollateralValue, CollateralCurrency,
                InvestmentCategory ?? InvestmentCategoryConditional.InvestmentCategory,
                InvestmentCategoryConditional.Desc ?? RepCats[InvestmentCategory.ToString()]
            };
            public static IEnumerable<string> ExcelHeaders() => new string[] {"Principal Amount", "Principal Currency", "Collateral Value",
                "Collateral Currency", "Investment Category", "Investment Category Desc."
            };
            public void Check(ref bool positionFlag)
            {
                if (InvestmentCategoryConditional.Desc is not null && InvestmentCategoryConditional.Desc.CheckAll())
                    positionFlag = true;
            }
            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        public class OtherRefInst : Instrument, ICheck
        {
            public OtherRefInst(XElement x) : base(x, true) { }
            public new static IEnumerable<string> ExcelHeaders()
                => new string[] { "Issuer Name", "Issuer Title"}.Concat(Identifiers.ExcelHeaders());
            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        public class RefInst : ICheck
        {
            public RefInst(XElement x)
            {
                if (x.Element("nestedDerivInfo") is not null)
                    NestedDerivative = new(x.Element("nestedDerivInfo").Descendants().First());
                else if (x.Element("indexBasketInfo") is not null)
                    Basket = new(x.Element("indexBasketInfo"));
                else if (x.Element("otherRefInst") is not null)
                    OtherRefInst = new(x.Element("otherRefInst"));
            }
            public NestedDerivative NestedDerivative { get; set; }
            public Basket Basket { get; set; }
            public OtherRefInst OtherRefInst { get; set; }
            public void Check(ref bool positionFlag)
            {
                if (NestedDerivative is not null)
                    NestedDerivative.Check(ref positionFlag);
                else if (Basket is not null)
                    Basket.Check(ref positionFlag);
                else if (OtherRefInst is not null)
                    OtherRefInst.Check(ref positionFlag);
            }
            public IEnumerable<object> ExcelValues()
            {
                if (NestedDerivative is not null)
                    return NestedDerivative.ExcelValues();
                else if (Basket is not null)
                    return Basket.ExcelValues();
                else if (OtherRefInst is not null)
                    return OtherRefInst.ExcelValues();
                else
                    return Array.Empty<object>();
            }
            public IEnumerable<string> ExcelHeaders()
            {
                if (NestedDerivative is not null)
                    return NestedDerivative.ExcelHeaders();
                else if (Basket is not null)
                    return Basket.ExcelHeaders();
                else if (OtherRefInst is not null)
                    return OtherRefInst.ExcelHeaders();
                else
                    return Array.Empty<string>();
            }
        }
        public class Record : IRecord, ICheck
        {
            public Record(XElement x, string type)
            {
                if (x.Element($"fixed{type}Desc") is not null)
                    FixedRate = new(x.Element($"fixed{type}Desc"));
                else if (x.Element($"floating{type}Desc") is not null)
                    FloatingRate = new(x.Element($"floating{type}Desc"));
                else if (x.Element($"other{type}Desc") is not null)
                    Other = (x.Element($"other{type}Desc").Attribute("fixedOrFloating").Value, x.Element($"other{type}Desc").Value);
            }
            public FixedRate FixedRate { get; set; }
            public FloatingRate FloatingRate { get; set; }
            public (string FixedOrFloating, object Desc) Other { get; set; }
            public void Check(ref bool positionFlag)
            {
                // Nothing to check for fixed
                if (FloatingRate is not null)
                    FloatingRate.Check(ref positionFlag);
                else if (Other is not (null, null) && Other.Desc.ToString().CheckAll())
                    positionFlag = true;
            }
            public IEnumerable<object> ExcelValues()
            {
                if (FixedRate is not null)
                    return FixedRate.ExcelValues();
                else if (FloatingRate is not null)
                    return FloatingRate.ExcelValues();
                else if (Other is not (null, null))
                    return new object[] { $"{Other.FixedOrFloating}: {Other.Desc}" };
                else
                    return Array.Empty<object>();
            }
            public IEnumerable<string> ExcelHeaders()
            {
                if (FixedRate is not null)
                    return FixedRate.ExcelHeaders();
                else if (FloatingRate is not null)
                    return FloatingRate.ExcelHeaders();
                else if (Other is not (null, null))
                    return new string[] { "Other" };
                else
                    return Array.Empty<string>();
            }
        }
        public class Rate
        {
            public Rate(XElement x)
            {
                FixedOrFloating = x.Attribute("fixedOrFloating")?.Value;
                Currency = x.Attribute("curCd")?.Value;
            }
            public string FixedOrFloating { get; set; }
            public string Currency { get; set; }
            public string Amount { get; set; }
        }
        public class FixedRate : Rate, ICheck
        {
            public FixedRate(XElement x) : base(x)
            {
                FixedRt = x.Attribute("fixedRt")?.Value;
                Amount = x.Attribute("amount")?.Value;
            }
            public string FixedRt { get; set; }
            public void Check(ref bool positionFlag)
            {// nothing to do
            }
            public static IEnumerable<string> ExcelHeaders() => new string[] { "Fixed/Floating", "Fixed Rate", "Amount", "Currency" };
            public IEnumerable<object> ExcelValues() => new object[] { FixedOrFloating, FixedRt, Amount, Currency };
            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        public class FloatingRate : Rate, ICheck
        {
            public FloatingRate(XElement x) : base(x)
            {
                Index = x.Attribute("floatingRtIndex")?.Value;
                Spread = x.Attribute("floatingRtSpread")?.Value;
                Amount = x.Attribute("pmntAmt")?.Value;
                RateResets = new();
                foreach (var rrt in x.Element("rtResetTenors").Descendants("rtResetTenor"))
                    RateResets.Add(new(rrt));
            }
            public string Index { get; set; }
            public string Spread { get; set; }
            public List<RateReset> RateResets { get; set; }
            public void Check(ref bool positionFlag)
            {
                if (Index is not null && Index.CheckAll())
                    positionFlag = true;
            }
            public static IEnumerable<string> ExcelHeaders() => new string[] { "Fixed/Floating", "Floating Rate Index", "Spread", "Payment Amt", "Currency",
                "Payment Due/Rate Resets" };
            public IEnumerable<object> ExcelValues() => new object[] { FixedOrFloating, Index, Spread, Amount, Currency,
                    NewlineValues(RateResets.Select(rrt => $"Due: {rrt.TenorUnits} {rrt.Tenor}, Reset: {rrt.DateUnit} {rrt.Date}")) };
            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        public struct RateReset
        {
            public RateReset(XElement x)
            {
                Date = x.Attribute("resetDt")?.Value;
                DateUnit = x.Attribute("resetDtUnit")?.Value;
                Tenor = x.Attribute("rateTenor")?.Value;
                TenorUnits = x.Attribute("rateTenorUnit")?.Value;
            }
            public string Date { get; set; }
            public string DateUnit { get; set; }
            public string Tenor { get; set; }
            public string TenorUnits { get; set; }
        }
        public class DerBase : ICheck
        {
            public DerBase(XElement x)
            {
                DerCategory = x.Attribute("derivCat").Value;
                Counterparties = new();
                foreach (var cp in x.Elements("counterparties"))
                    Counterparties.Add(new(cp, false));
            }
            public string DerCategory { get; set; }
            public List<Counterparty> Counterparties { get; set; }
            public void Check(ref bool positionFlag) { }// nothing to check unless it's personnel related
            public IEnumerable<string> ExcelHeaders() => new string[] { "Derivative Category", "Counterparties" };
            public IEnumerable<object> ExcelValues() => new object[] { DerCats[DerCategory],
                NewlineValues(Counterparties.Select(cp => $"{cp.Name}, lei: {cp.Lei}")) };
        }
        public class DerBaseRef : DerBase, IRef, ICheck
        {
            public DerBaseRef(XElement x) : base(x)
            {
                if (x.Element("derivAddlInfo") is not null)
                    DerAddlInfo = new(x.Element("derivAddlInfo"));
            }
            public SecInfo DerAddlInfo { get; set; }
            public new void Check(ref bool positionFlag)
            {
                if (DerAddlInfo is not null)
                    DerAddlInfo.Check(ref positionFlag);
            }
            public new IEnumerable<string> ExcelHeaders() => DerAddlInfo is null ? base.ExcelHeaders() :
                base.ExcelHeaders().Concat(SecInfo.ExcelHeaders());
            public new IEnumerable<object> ExcelValues() => DerAddlInfo is null ? base.ExcelValues() :
                base.ExcelValues().Concat(DerAddlInfo.ExcelValues());
        }
        public class DerBaseRef<T> : DerBaseRef, ICheck, IDerivative
            where T : ICheck, IDerivative
        {
            public DerBaseRef(XElement x, T der) : base(x) => Der = der;
            public T Der { get; set; }

            public new void Check(ref bool positionFlag)
            {
                base.Check(ref positionFlag);
                if (Der is not null)
                    Der.Check(ref positionFlag);
            }
            public new IEnumerable<string> ExcelHeaders() => base.ExcelHeaders().Concat(Der.ExcelHeaders());
            public new IEnumerable<object> ExcelValues() => base.ExcelValues().Concat(Der.ExcelValues());
            public RefInst GetRefInst() => Der.GetRefInst();
        }
        public class DerBaseRef<T, U> : DerBaseRef<ForexFwdOrSwap>, ICheck, IDerivative
            where U : ICheck, IDerivative
        {
            public DerBaseRef(XElement x, U der) : base(x, new(x)) => Der2 = der;
            public U Der2 { get; set; }
            public new void Check(ref bool positionFlag)
            {
                base.Check(ref positionFlag);
                if (Der2 is not null)
                    Der2.Check(ref positionFlag);
            }
            public new IEnumerable<string> ExcelHeaders() => base.ExcelHeaders().Concat(Der2.ExcelHeaders());
            public new IEnumerable<object> ExcelValues() => base.ExcelValues().Concat(Der2.ExcelValues());
            public new RefInst GetRefInst() => Der2.GetRefInst();
        }
        public class DerBaseHolding<T> : DerBase, ICheck, IDerivative
            where T : IHolding, ICheck, IDerivative
        {
            public DerBaseHolding(XElement x, T der) : base(x)
            {
                Der = der;
                Der.UnrealizedAppreciation = x.Element("unrealizedAppr")?.Value;
            }
            public T Der { get; set; }
            public new void Check(ref bool positionFlag)
            {
                base.Check(ref positionFlag);
                if (Der is not null)
                    Der.Check(ref positionFlag);
            }
            public new IEnumerable<string> ExcelHeaders() => base.ExcelHeaders().Concat(Der.ExcelHeaders()).Append("Unrealized Appreciation");
            public new IEnumerable<object> ExcelValues() => base.ExcelValues().Concat(Der.ExcelValues()).Append(Der.UnrealizedAppreciation);
            public RefInst GetRefInst() => Der.GetRefInst();
        }
        public class DerBaseHolding<T, U> : DerBaseHolding<ForexFwdOrSwapHolding>, ICheck, IDerivative
            where U : IHolding, ICheck, IDerivative
        {
            public DerBaseHolding(XElement x, U der) : base(x, new(x))
            {
                Der2 = der;
                Der2.UnrealizedAppreciation = x.Element("unrealizedAppr")?.Value;
            }
            public U Der2 { get; set; }
            public new void Check(ref bool positionFlag)
            {
                base.Check(ref positionFlag);
                if (Der2 is not null)
                    Der2.Check(ref positionFlag);
            }
            public new IEnumerable<string> ExcelHeaders() => base.ExcelHeaders().Concat(Der2.ExcelHeaders()).Append("Unrealized Appreciation");
            public new IEnumerable<object> ExcelValues() => base.ExcelValues().Concat(Der2.ExcelValues()).Append(Der2.UnrealizedAppreciation);
            public new RefInst GetRefInst() => Der2.GetRefInst();
        }
        public class Derivative<T, U, V, W, X> : ICheck, IDerivative
            where T : ICheck, IDerivative
            where U : ICheck, IDerivative
            where V : ICheck, IDerivative
            where W : ICheck, IDerivative
            where X : ICheck, IDerivative
        {
            public Derivative(XElement x) => DerCat = x.Attribute("derivCat").Value;
            public T Forward { get; set; }
            public U Future { get; set; }
            public V Swap { get; set; }
            public W OptionSwaptionWarrant { get; set; }
            public X Other { get; set; }
            public string DerCat { get; set; }
            public void Check(ref bool positionFlag)
            {
                if (Forward is not null)
                    Forward.Check(ref positionFlag);
                else if (Future is not null)
                    Future.Check(ref positionFlag);
                else if (Swap is not null)
                    Swap.Check(ref positionFlag);
                else if (OptionSwaptionWarrant is not null)
                    OptionSwaptionWarrant.Check(ref positionFlag);
                else if (Other is not null)
                    Other.Check(ref positionFlag);
            }
            public IEnumerable<object> ExcelValues()
            {
                if (Forward is not null)
                    return Forward.ExcelValues();
                else if (Future is not null)
                    return Future.ExcelValues();
                else if (Swap is not null)
                    return Swap.ExcelValues();
                else if (OptionSwaptionWarrant is not null)
                    return OptionSwaptionWarrant.ExcelValues();
                else if (Other is not null)
                    return Other.ExcelValues();
                else
                    return Array.Empty<object>();
            }
            public IEnumerable<string> ExcelHeaders()
            {
                if (Forward is not null)
                    return Forward.ExcelHeaders();
                else if (Future is not null)
                    return Future.ExcelHeaders();
                else if (Swap is not null)
                    return Swap.ExcelHeaders();
                else if (OptionSwaptionWarrant is not null)
                    return OptionSwaptionWarrant.ExcelHeaders();
                else if (Other is not null)
                    return Other.ExcelHeaders();
                else
                    return Array.Empty<string>();
            }
            public RefInst GetRefInst()
            {
                switch (DerCat)
                {
                    case "FWD":
                        return Forward.GetRefInst();
                    case "FUT":
                        return Future.GetRefInst();
                    case "SWP":
                        return Swap.GetRefInst();
                    case "OPT":
                        return OptionSwaptionWarrant.GetRefInst();
                    case "SWO":
                        goto case "OPT";
                    case "WAR":
                        goto case "OPT";
                    case "OTH":
                        return Other.GetRefInst();
                    default:
                        Console.WriteLine($"New Nested Derivative Category found!!! (someone fucked up filing)");
                        return null;
                }
            }
        }
        public class DerivativeHolding : Derivative<ForwardHolding, FutureHolding, SwapHolding, OptionSwaptionWarrantHolding, OtherHolding>
        {
            public DerivativeHolding(XElement x) : base(x)
            {
                switch (DerCat)
                {
                    case "FWD":
                        Forward = new(x);
                        break;
                    case "FUT":
                        Future = new(x);
                        break;
                    case "SWP":
                        Swap = new(x);
                        break;
                    case "OPT":
                        OptionSwaptionWarrant = new(x);
                        break;
                    case "SWO":
                        goto case "OPT";
                    case "WAR":
                        goto case "OPT";
                    case "OTH":
                        Other = new(x);
                        break;
                    default:
                        Console.WriteLine($"New Derivative Category found!!! (someone fucked up filing)");
                        break;
                }
            }
        }
        public class NestedDerivative : Derivative<ForwardRef, FutureRef, SwapRef, OptionSwaptionWarrantRef, OtherRef>
        {
            public NestedDerivative(XElement x) : base(x)
            {
                switch (DerCat)
                {
                    case "FWD":
                        Forward = new(x);
                        break;
                    case "FUT":
                        Future = new(x);
                        break;
                    case "SWP":
                        Swap = new(x);
                        break;
                    case "OPT":
                        OptionSwaptionWarrant = new(x);
                        break;
                    case "SWO":
                        goto case "OPT";
                    case "WAR":
                        goto case "OPT";
                    case "OTH":
                        Other = new(x);
                        break;
                    default:
                        Console.WriteLine($"New Nested Derivative Category found!!! (someone fucked up filing)");
                        break;
                }
            }
        }
        public class ForexFwdOrSwap : ICheck, IDerivative
        {
            public ForexFwdOrSwap(XElement x)
            {
                AmountSold = x.Element("amtCurSold")?.Value;
                if (!IsDerEmpty())
                {
                    CurrencySold = x.Element("curSold").Value;
                    AmountPurchased = x.Element("amtCurPur").Value;
                    CurrencyPurchased = x.Element("curPur").Value;
                    SettlementDate = x.Element("settlementDt").Value;
                }
            }
            public object AmountSold { get; set; }
            public object CurrencySold { get; set; }
            public object AmountPurchased { get; set; }
            public object CurrencyPurchased { get; set; }
            public object SettlementDate { get; set; }
            public void Check(ref bool positionFlag)
            {// nothing to do here
            }
            public IEnumerable<object> ExcelValues() => IsDerEmpty() ? Array.Empty<object>() :
                new object[] { AmountSold, CurrencySold, AmountPurchased, CurrencyPurchased, SettlementDate };
            public IEnumerable<string> ExcelHeaders() => IsDerEmpty() ? Array.Empty<string>() :
                new string[] { "Amount Sold", "Currency Sold", "Amount Purchased", "Currency Purchased", "Settlement Date" };
            public RefInst GetRefInst() => null;
            public bool IsDerEmpty() => AmountSold is null;
        }
        public class ForexFwdOrSwapHolding : ForexFwdOrSwap, IHolding
        {
            public ForexFwdOrSwapHolding(XElement x) : base(x) => UnrealizedAppreciation = x.Element("unrealizedAppr")?.Value;
            public object UnrealizedAppreciation { get; set; }
            public new IEnumerable<object> ExcelValues() => IsDerEmpty() ? base.ExcelValues() : // this will pass the empty array
                base.ExcelValues().Append(UnrealizedAppreciation);
            public new IEnumerable<string> ExcelHeaders() => IsDerEmpty() ? base.ExcelHeaders() :
                base.ExcelHeaders().Append("Unrealized Appreciation");
        }
        public class NonForexFutOrFwd : ICheck, IDerivative // no difference between NonFxOrFwd_base & Future_base so combined
        {
            public NonForexFutOrFwd(XElement x)
            {
                PayoffProfile = x.Element("payOffProf")?.Value;
                // RefInst is mandatory under Future instruments but not Forwards, so check if null covers both cases
                RefInst = x.Element("descRefInstrmnt") is null ? null : new(x.Element("descRefInstrmnt"));
                ExpirationDate = x.Element("expDate")?.Value;
                NotionalAmount = x.Element("notionalAmt")?.Value;
                Currency = x.Element("curCd")?.Value;
            }
            public object PayoffProfile { get; set; }
            public RefInst RefInst { get; set; }
            public object ExpirationDate { get; set; }
            public object NotionalAmount { get; set; }
            public object Currency { get; set; }
            public void Check(ref bool positionFlag)
            {
                if (RefInst is not null)
                    RefInst.Check(ref positionFlag);
            }
            public IEnumerable<object> ExcelValues()
            {
                if (IsDerEmpty() && RefInst is null)
                    return Array.Empty<object>();

                var list = new List<object>();
                if (!IsDerEmpty())
                    list.AddRange(new object[] { PayoffProfile, ExpirationDate, NotionalAmount, Currency });
                //if (RefInst is not null)
                //    list.AddRange(RefInst.ExcelValues());
                return list;
            }
            public IEnumerable<string> ExcelHeaders()
            {
                if (IsDerEmpty() && RefInst is null)
                    return Array.Empty<string>();

                var list = new List<string>();
                if (!IsDerEmpty())
                    list.AddRange(new string[] { "Payoff Profile", "Exp. Date", "Notional Amount", "Currency" });
                //if (RefInst is not null)
                //    list.AddRange(RefInst.ExcelHeaders());
                return list;
            }
            public RefInst GetRefInst() => RefInst;
            public bool IsDerEmpty() => PayoffProfile is null;
        }
        public class NonForexFutOrFwdHolding : NonForexFutOrFwd, IHolding // only diff between Fut_hldng & NFOF: NFOF only takes 2 insts
        {
            public NonForexFutOrFwdHolding(XElement x) : base(x) => UnrealizedAppreciation = x.Element("unrealizedAppr")?.Value;
            public object UnrealizedAppreciation { get; set; }
            public new IEnumerable<object> ExcelValues() => IsDerEmpty() ? base.ExcelValues() :
                base.ExcelValues().Append(UnrealizedAppreciation);
            public new IEnumerable<string> ExcelHeaders() => IsDerEmpty() ? base.ExcelHeaders() :
                base.ExcelHeaders().Append("Unrealized Appreciation");
        }
        public class NonForexOrSwap : ICheck, IDerivative
        {
            public NonForexOrSwap(XElement x)
            {
                RefInst = x.Element("descRefInstrmnt") is null ? null : new(x.Element("descRefInstrmnt"));
                SwapFlag = x.Element("swapFlag")?.Value;
                Receipts = new(x);
                Payments = new(x);
                TerminationDate = x.Element("terminationDt")?.Value;
                UpfrontPaymentAmount = x.Element("upfrontPmnt")?.Value;
                PaymentCurrency = x.Element("pmntCurCd")?.Value;
                UpfrontReceiptAmount = x.Element("upfrontRcpt")?.Value;
                ReceiptCurrency = x.Element("rcptCurCd")?.Value;
                NotionalAmount = x.Element("notionalAmt")?.Value;
                Currency = x.Element("curCd")?.Value;
            }
            public RefInst RefInst { get; set; }
            public object SwapFlag { get; set; }
            public Receipt Receipts { get; set; }
            public Payment Payments { get; set; }
            public object TerminationDate { get; set; }
            public object UpfrontPaymentAmount { get; set; }
            public object PaymentCurrency { get; set; }
            public object UpfrontReceiptAmount { get; set; }
            public object ReceiptCurrency { get; set; }
            public object NotionalAmount { get; set; }
            public object Currency { get; set; }

            public void Check(ref bool positionFlag)
            {
                if (RefInst is not null)
                    RefInst.Check(ref positionFlag);

                Receipts.Check(ref positionFlag);
                Payments.Check(ref positionFlag);
            }
            public IEnumerable<object> ExcelValues()
            {
                var list = new List<object>() { SwapFlag }; // if this is ever a reference instrument this should not be there
                list.AddRange(Receipts.ExcelValues());
                list.AddRange(Payments.ExcelValues());
                if (!IsDerEmpty())
                    list.AddRange(new object[] { TerminationDate, UpfrontPaymentAmount, PaymentCurrency,
                        UpfrontReceiptAmount, ReceiptCurrency, NotionalAmount, Currency });
                return list;
            }
            public IEnumerable<string> ExcelHeaders()
            {
                var list = new List<string> { "Swap Flag" };
                list.AddRange(Receipts.ExcelHeaders());
                list.AddRange(Payments.ExcelHeaders());
                if (!IsDerEmpty())
                    list.AddRange(new string[] { "Termination Date", "Upfront Payment Amt", "Payment Currency",
                        "Upfront Receipt Amt", "Receipt Currency", "Notional Amount", "Currency"});
                return list;
            }
            public RefInst GetRefInst() => RefInst;
            public bool IsDerEmpty() => SwapFlag is null;
        }
        public class NonForexOrSwapHolding : NonForexOrSwap, IHolding
        {
            public NonForexOrSwapHolding(XElement x) : base(x) => UnrealizedAppreciation = x.Element("unrealizedAppr")?.Value;
            public object UnrealizedAppreciation { get; set; }
            public new IEnumerable<object> ExcelValues() => IsDerEmpty() ? base.ExcelValues() :
                base.ExcelValues().Append(UnrealizedAppreciation);
            public new IEnumerable<string> ExcelHeaders() => IsDerEmpty() ? base.ExcelHeaders() :
                base.ExcelHeaders().Append("Unrealized Appreciation");
        }
        public class OptionSwaptionWarrantDer : ICheck, IDerivative
        {
            public OptionSwaptionWarrantDer(XElement x)
            {
                PutOrCall = x.Element("putOrCall")?.Value;
                WrittenOrPurchased = x.Element("writtenOrPur")?.Value;
                RefInst = new(x.Element("descRefInstrmnt")); // mandatory
                PrincipalAmount = x.Element("principalAmt")?.Value;
                Currency = x.Element("curCd")?.Value;
                ExercisePrice = x.Element("exercisePrice")?.Value;
                ExercisePriceCurrency = x.Element("exercisePriceCurCd")?.Value;
                ExpirationDate = x.Element("expDt")?.Value;
            }
            public object PutOrCall { get; set; }
            public object WrittenOrPurchased { get; set; }
            public RefInst RefInst { get; set; }
            public object PrincipalAmount { get; set; }
            public object Currency { get; set; }
            public object ExercisePrice { get; set; }
            public object ExercisePriceCurrency { get; set; }
            public object ExpirationDate { get; set; }
            public void Check(ref bool positionFlag)
            {
                if (RefInst is not null)
                    RefInst.Check(ref positionFlag);
            }
            public IEnumerable<object> ExcelValues() => new object[] { PutOrCall, WrittenOrPurchased, PrincipalAmount, Currency,
                ExercisePrice, ExercisePriceCurrency, ExpirationDate };//.Concat(RefInst.ExcelValues());
            public IEnumerable<string> ExcelHeaders() => new string[] { "Put/Call", "Written/Purchased", "Principal Amt", "Currency",
                "Exercise Price", "Exercise Price Currency", "Exp. Date" };//.Concat(RefInst.ExcelHeaders());
            public RefInst GetRefInst() => RefInst;
        }
        public class OptionSwaptionWarrantDerHolding : OptionSwaptionWarrantDer, IHolding
        {
            public OptionSwaptionWarrantDerHolding(XElement x) : base(x)
            {
                ShareNo = x.Element("shareNo")?.Value;
                Delta = x.Element("delta")?.Value;
                UnrealizedAppreciation = x.Element("unrealizedAppr")?.Value;
            }
            public object ShareNo { get; set; }
            public object Delta { get; set; }
            public object UnrealizedAppreciation { get; set; }
            public new IEnumerable<object> ExcelValues() => base.ExcelValues().Concat(new object[] { ShareNo, Delta, UnrealizedAppreciation });
            public new IEnumerable<string> ExcelHeaders() => base.ExcelHeaders().Concat(new string[] { "ShareNo", "Delta", "Unrealized Appreciation" });
        }
        public class OtherDer : ICheck, IDerivative
        {
            public OtherDer(XElement x)
            {
                Desc = x.Value;
                RefInst = x.Element("descRefInstrmnt") is null ? null : new(x.Element("descRefInstrmnt"));
                TerminationDate = x.Element("terminationDt")?.Value;
                NotionalAmounts = new();
                var nAmts = x.Element("notionalAmts").Descendants("notionalAmt");
                foreach (var na in nAmts)
                    NotionalAmounts.Add((na.Attribute("amt")?.Value, na.Attribute("curCd")?.Value));
            }
            public object Desc { get; set; }
            public RefInst RefInst { get; set; }
            public object TerminationDate { get; set; }
            public List<(string Amount, string Currency)> NotionalAmounts { get; set; }

            public void Check(ref bool positionFlag)
            {
                if (Desc is not null && Desc.ToString().CheckAll())
                    positionFlag = true;

                if (RefInst is not null)
                    RefInst.Check(ref positionFlag);
            }
            public IEnumerable<object> ExcelValues()
                => new object[] { Desc, TerminationDate,
                    NewlineValues(NotionalAmounts.Select(x => $"{x.Amount} {x.Currency}")) };//.Concat(RefInst.ExcelValues());
            public IEnumerable<string> ExcelHeaders() =>
                new string[] { "Desc", "Termination Date", "Notional Amounts" };//.Concat(RefInst.ExcelHeaders());
            public RefInst GetRefInst() => RefInst;
        }
        public class OtherDerHolding : OtherDer, IHolding
        {
            public OtherDerHolding(XElement x) : base(x)
            {
                Delta = x.Element("delta")?.Value;
                UnrealizedAppreciation = x.Element("unrealizedAppr")?.Value;
            }
            public object Delta { get; set; }
            public object UnrealizedAppreciation { get; set; }
            public new IEnumerable<object> ExcelValues() => base.ExcelValues().Concat(new object[] { Delta, UnrealizedAppreciation });
            public new IEnumerable<string> ExcelHeaders() => base.ExcelHeaders().Concat(new string[] { "Delta", "Unrealized Appreciation" });
        }
        public class Component : ICheck
        {
            public Component(XElement x)
            {
                Name = x.Element("othIndName")?.Value;
                foreach (var d in x.Element("identifiers").Descendants())
                {
                    Id = (d.Name.LocalName == "other" ? d.Attribute("otherDesc").Value : d.Name.LocalName, d.Attribute("value").Value);
                }
                NotionalAmount = x.Element("othIndNotAmt")?.Value;
                Currency = x.Element("othIndCurCd")?.Value;
                Value = x.Element("othIndValue")?.Value;
                IssuerCurrency = x.Element("othIndIssCurCd")?.Value;
            }
            public object Name { get; set; }
            public (string Cat, object Desc) Id { get; set; }
            public object NotionalAmount { get; set; }
            public object Currency { get; set; }
            public object Value { get; set; }
            public object IssuerCurrency { get; set; }
            public void Check(ref bool positionFlag)
            {
                if (Name is not null && Name.ToString().CheckAll())
                    positionFlag = true;

                if (Id is not (null, null) && Id.Desc.ToString().CheckAll())
                    positionFlag = true;
            }
            public IEnumerable<object> ExcelValues() => new object[] { Name,
                Id is (null, null) ? null : $"{Id.Cat}: {Id.Desc}", NotionalAmount, Currency, Value, IssuerCurrency };

            // Could likely make headers using these rather than multiline cell like current, need to figure out solid way of doing so
            public static IEnumerable<string> ExcelHeaders() => new string[] { "Name",
                "Id", "Notional Amount", "Currency", "Value", "Issuer Currency" };
            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        public class Basket : ICheck
        {
            public Basket(XElement x)
            {
                IndexName = x.Element("indexName")?.Value;
                IndexIdentifier = x.Element("indexIdentifier")?.Value;
                Narrative = x.Element("narrativeDesc")?.Value;
                var comps = x.Element("components");
                if (comps is null)
                    return;

                Components = new();
                foreach (var comp in comps.Descendants("component"))
                    Components.Add(new(comp));
            }
            public object IndexName { get; set; }
            public object IndexIdentifier { get; set; }
            public object Narrative { get; set; }
            public List<Component> Components { get; set; }
            public void Check(ref bool positionFlag)
            {
                if (IndexName is not null && IndexName.ToString().CheckAll())
                    positionFlag = true;

                if (IndexIdentifier is not null && IndexIdentifier.ToString().CheckAll())
                    positionFlag = true;

                if (Narrative is not null && Narrative.ToString().CheckAll())
                    positionFlag = true;

                if (Components is not null)
                    foreach (var comp in Components)
                        comp.Check(ref positionFlag);
            }
            public IEnumerable<object> ExcelValues() => new object[] { IndexName, IndexIdentifier, Narrative };
            public static IEnumerable<string> ExcelHeaders() => new string[] { "Index Name", "Index Id", "Narrative" };
            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        public struct SecurityLending : ICheck
        {
            public SecurityLending(XElement x)
            {
                CashCollateralCondition = (null, default);
                NonCashCollateralCondition = (null, default);
                LoanByFundCondition = (null, default);
                IsCashCollateral = x.Element("isCashCollateral")?.Value;
                if (IsCashCollateral is null)
                    CashCollateralCondition = (x.Element("cashCollateralCondition").Attribute("isCashCollateral").Value,
                        decimal.Parse(x.Element("cashCollateralCondition").Attribute("cashCollateralVal").Value));

                IsNonCashCollateral = x.Element("isNonCashCollateral")?.Value;
                if (IsNonCashCollateral is null)
                    NonCashCollateralCondition = (x.Element("nonCashCollateralCondition").Attribute("isNonCashCollateral").Value,
                        decimal.Parse(x.Element("nonCashCollateralCondition").Attribute("nonCashCollateralVal").Value));

                IsLoanByFund = x.Element("isLoanByFund")?.Value;
                if (IsLoanByFund is null)
                    LoanByFundCondition = (x.Element("loanByFundCondition").Attribute("isLoanByFund").Value,
                        decimal.Parse(x.Element("loanByFundCondition").Attribute("loanVal").Value));
            }
            public object IsCashCollateral { get; set; }
            public (string IsCashCollateral, decimal CashCollateralValue) CashCollateralCondition { get; set; }
            public object IsNonCashCollateral { get; set; }
            public (string IsNonCashCollateral, decimal NonCashCollateralValue) NonCashCollateralCondition { get; set; }
            public object IsLoanByFund { get; set; }
            public (string IsLoanByFund, decimal LoanValue) LoanByFundCondition { get; set; }
            public IEnumerable<object> ExcelValues() => new object[] {
                CashCollateralCondition.CashCollateralValue,
                NonCashCollateralCondition.NonCashCollateralValue,
                LoanByFundCondition.LoanValue
            };
            public static IEnumerable<string> ExcelHeaders() => new string[] { "Cash Collateral", "Non Cash Collateral", "Loan By Fund",
                "% of Holding on Loan" };
            void ICheck.Check(ref bool positionFlag) { } // nothin to do since no desc
            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        public class Instrument : ICheck
        {
            public Instrument(XElement x, bool issuer = false)
            {
                if (issuer)
                {
                    Name = x?.Element("issuerName")?.Value ?? default;
                    Title = x?.Element("issueTitle")?.Value ?? default;
                }
                else
                {
                    Name = x.Element("name").Value;
                    Title = x.Element("title").Value;
                }
                Identifiers = new(x);
            }
            public string Name { get; set; }
            public string Title { get; set; }
            public Identifiers Identifiers { get; set; } // this needs to be a list of IDs for OtherRefInst IDs
            public bool IsRelevant { get; set; } = false;
            public void Check(ref bool positionFlag)
            {
                if (Name != default && Name.CheckAll())
                    positionFlag = true;

                if (Title != default && Title.CheckAll())
                    positionFlag = true;

                Identifiers.Check(ref positionFlag);
            }
            public static IEnumerable<string> ExcelHeaders() => new string[] { "Name", "Title" }.Concat(Identifiers.ExcelHeaders());
            public IEnumerable<object> ExcelValues() => new object[] { Name, Title }.Concat(Identifiers.ExcelValues());
            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        public struct Identifiers : ICheck
        {
            public Identifiers(XElement x)
            {
                Cusip = null;
                Isin = null;
                Ticker = null;
                Other = null;
                foreach (var id in x.Element("identifiers").Descendants())
                    switch (id.Name.LocalName)
                    {
                        case "cusip":
                            Cusip = id.Attribute("value").Value ?? default;
                            break;
                        case "isin":
                            Isin = id.Attribute("value").Value ?? default;
                            break;
                        case "ticker":
                            Ticker = id.Attribute("value").Value ?? default;
                            break;
                        default:
                            Other = $"{id.Attribute("otherDesc").Value}: {id.Attribute("value").Value}";
                            break;
                    }
            }
            public string Cusip { get; set; }
            public string Isin { get; set; }
            public string Ticker { get; set; }
            public string Other { get; set; }

            public void Check(ref bool positionFlag)
            {
                if (Cusip.CheckId(CusipCanary))
                    positionFlag = true;

                if (Isin.CheckId(IsinCanary))
                    positionFlag = true;

                if (Ticker.CheckId(TickerCanary))
                    positionFlag = true;

                if (Other.CheckAll())
                    positionFlag = true;
            }

            public static IEnumerable<string> ExcelHeaders() => new string[] { "Inst. Cusip", "Inst. Isin", "Inst. Ticker", "Inst. Other Id" };

            public IEnumerable<object> ExcelValues() => new object[] { Cusip, Isin, Ticker, Other };
            IEnumerable<string> ICheck.ExcelHeaders() => ExcelHeaders();
        }
        #region EssentiallyEmptyClasses
        public class ForwardHolding : DerBaseHolding<ForexFwdOrSwapHolding, NonForexFutOrFwdHolding>
        {
            public ForwardHolding(XElement x) : base(x, new(x))
            {
            }
        }
        public class ForwardRef : DerBaseRef<ForexFwdOrSwap, NonForexFutOrFwd>
        {
            public ForwardRef(XElement x) : base(x, new(x))
            {
            }
        }
        public class FutureHolding : DerBaseHolding<NonForexFutOrFwdHolding>
        {
            public FutureHolding(XElement x) : base(x, new(x))
            {
            }
        }
        public class FutureRef : DerBaseRef<NonForexFutOrFwd>
        {
            public FutureRef(XElement x) : base(x, new(x))
            {
            }
        }
        public class SwapHolding : DerBaseHolding<ForexFwdOrSwapHolding, NonForexOrSwapHolding>
        {
            public SwapHolding(XElement x) : base(x, new(x))
            {
            }
        }
        public class SwapRef : DerBaseRef<ForexFwdOrSwap, NonForexOrSwap>
        {
            public SwapRef(XElement x) : base(x, new(x))
            {
            }
        }
        public class OptionSwaptionWarrantHolding : DerBaseHolding<OptionSwaptionWarrantDerHolding>
        {
            public OptionSwaptionWarrantHolding(XElement x) : base(x, new(x))
            {
            }
        }
        public class OptionSwaptionWarrantRef : DerBaseRef<OptionSwaptionWarrantDer>
        {
            public OptionSwaptionWarrantRef(XElement x) : base(x, new(x))
            {
            }
        }
        public class OtherHolding : DerBaseHolding<OtherDerHolding>
        {
            public OtherHolding(XElement x) : base(x, new(x))
            {
            }
        }
        public class OtherRef : DerBaseRef<OtherDer>
        {
            public OtherRef(XElement x) : base(x, new(x))
            {
            }
        }
        public class Receipt : Record
        {
            public Receipt(XElement x) : base(x, "Rec")
            {
            }
        }
        public class Payment : Record
        {
            public Payment(XElement x) : base(x, "Pmnt")
            {
            }
        }
        public class Counterparty
        {
            public Counterparty(XElement x, bool attribute)
            {
                if (attribute)
                {
                    Name = x.Attribute("name")?.Value;
                    Lei = x.Attribute("lei")?.Value;
                }
                else
                {
                    Name = x.Element("counterpartyName")?.Value;
                    Lei = x.Element("counterpartyLei")?.Value;
                }
            }
            public object Name { get; set; }
            public object Lei { get; set; }
        }
        #endregion
        #region interfaces
        public interface IHolding
        {
            public object UnrealizedAppreciation { get; set; }
        }
        public interface IRef
        {
            public SecInfo DerAddlInfo { get; set; }
        }
        public interface IRecord
        {
            public FixedRate FixedRate { get; set; }
            public FloatingRate FloatingRate { get; set; }
            public (string FixedOrFloating, object Desc) Other { get; set; }
        }
        public interface IDerivative
        {
            public RefInst GetRefInst();
        }
        public interface ICheck
        {
            public void Check(ref bool positionFlag);
            public IEnumerable<object> ExcelValues();
            public IEnumerable<string> ExcelHeaders();
        }
        #endregion
        #region Category Dictionaries
        public static readonly Dictionary<string, string> AssCats = new()
        {
            ["STIV"] = "Short-term investment vehicle",
            ["RA"] = "Repurchase agreement",
            ["EC"] = "Equity-common",
            ["EP"] = "Equity-preferred",
            ["DBT"] = "Debt",
            ["DCO"] = "Derivative-commodity",
            ["DCR"] = "Derivative-credit",
            ["DE"] = "Derivative-equity",
            ["DFE"] = "Derivative-foreign exchange",
            ["DIR"] = "Derivative-interest rate",
            ["DO"] = "Derivatives-other",
            ["SN"] = "Structured note",
            ["LON"] = "Loan",
            ["ABS-MBS"] = "ABS-mortgage backed security",
            ["ABS-APCP"] = "ABS-asset backed commercial paper",
            ["ABS-CBDO"] = "ABS-collaterilized bond/debt obligation",
            ["ABS-O"] = "ABS-other",
            ["COMM"] = "Commodity",
            ["RE"] = "Real estate"
        };
        public static readonly Dictionary<string, string> IssCats = new()
        {
            ["CORP"] = "Corporate",
            ["UST"] = "U.S. Treasury",
            ["USGA"] = "U.S. government agency",
            ["USGSE"] = "U.S. government sponsored entity",
            ["MUN"] = "Municipal",
            ["NUSS"] = "Non-U.S. sovereign",
            ["PF"] = "Private fund",
            ["RF"] = "Registered fund",
        };
        public static readonly Dictionary<string, string> LiqCats = new()
        {
            ["MLI"] = "Moderately liquid investments",
            ["LLI"] = "Less liquid investments",
            ["ILI"] = "Illiquid investments",
        };
        public static readonly Dictionary<string, string> LiqInfos = new()
        {
            ["HLI"] = "Highly Liquid Investments",
            ["MLI"] = "Moderately liquid investments",
            ["LLI"] = "Less liquid investments",
            ["ILI"] = "Illiquid investments",
            ["N/A"] = "Not applicable",
        };
        public static readonly Dictionary<string, string> RepCats = new()
        {
            ["ABS"] = "Asset-backed securities",
            ["ACMO"] = "Agency collateralized mortgage obligations",
            ["ADAS"] = "Agency debentures and agency strips",
            ["AMBS"] = "Agency mortgage-backed securities",
            ["PLCMO"] = "Private label collateralized mortgage obligations",
            ["CDS"] = "Corporate debt securities",
            ["EQT"] = "Equities",
            ["MM"] = "Money market",
            ["UST"] = "U.S. Treasuries (including strips)",
        };
        public static readonly Dictionary<string, string> InvCats = new()
        {
            ["ABS"] = "Asset-backed securities",
            ["ACMO"] = "Agency collateralized mortgage obligations",
            ["ACMBS"] = "Agency debentures and agency strips",
            ["AMBS"] = "Agency mortgage-backed securities",
            ["UST"] = "U.S. Treasuries (including strips)",
            ["N/A"] = "Not applicable",
        };
        public static readonly Dictionary<string, string> DerCats = new()
        {
            ["FWD"] = "Forward",
            ["FUT"] = "Future",
            ["SWP"] = "Swap",
            ["OPT"] = "Option",
            ["SWO"] = "Swaption",
            ["WAR"] = "Warrant",
            ["OTH"] = "Other Derivative"
        };
        public static readonly Dictionary<string, string> UnitDescs = new()
        {
            ["NC"] = "Contracts",
            ["NS"] = "Shares",
            ["PA"] = "Principal Amount"
        };
        #endregion
        [Flags]
        public enum FormType
        {
            Regular = 1,
            NotTimely = 2,
            Ammended = 4,
            Repeat = 8
        }
        public class NportItem : Attribute
        {
            private readonly string _item;
            public override string ToString() => _item;
            public NportItem(string item) { _item = item; }
        }
        private static IFlurlRequest SecSearch = "https://efts.sec.gov"
            .AppendPathSegments(new object[] { "LATEST", "search-index" })
            .WithHeaders(new { Accept = "application/json", User_Agent = "Retail Investor investorrelations@gamestop.com" });
        private static IFlurlRequest headerUrl = "https://www.sec.gov"
            .WithHeaders(new { Accept = "text/html", User_Agent = "Retail Investor investorrelations@gamestop.com", Host = "www.sec.gov" });
        private static IFlurlRequest archiveUrl = "https://www.sec.gov"
            .WithHeaders(new { Accept = "text/xml", User_Agent = "Retail Investor investorrelations@gamestop.com", Host = "www.sec.gov" }); //Accept_Encoding = "gzip, deflate"
        private static readonly string[] NportForms = new string[] {
            "NPORT-P",              // NPORT
            "NPORT-NP",             // non public
            "NPORT-P/A",            // ammended
            "NT NPORT-P",           // Not-Timely NPORT
            "NT NPORT-NP",          // non public
            "NT NPORT-NP/A",        // ammended np
        };
        private static Task<dynamic> NportUpdates(string term, int from = 0, DateTime? start = null, DateTime? end = null) => SecSearch.PostJsonAsync(new
        {
            q = term,
            dateRange = "custom",
            category = "custom",
            startdt = start is DateTime st ? st.SecString() : Start.SecString(),
            enddt = end is DateTime et ? et.SecString() : End.SecString(),
            forms = NportForms,
            from
        }).ReceiveJson();
    }
}
