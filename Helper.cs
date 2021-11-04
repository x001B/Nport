using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;
using System.Xml.Linq;

namespace ETF
{
    public static class Helper
    {
        public static bool Testing { get; set; } = false;
        public static bool UseMultiSearch { get; set; } = true;
        public static string RegNameCanary { get; set; }
        public static string RegCikCanary { get; set; }
        public static string RegLeiCanary { get; set; }
        public static string SeriesNameCanary { get; set; }
        public static string SeriesIdCanary { get; set; }
        public static string SeriesLeiCanary { get; set; }
        public static string IssCatCanary { get; set; }
        public static string AssCatCanary { get; set; }
        public static string CusipCanary { get; set; }// = "36467W109";
        public static string IsinCanary { get; set; }// = "US36467W1099";
        public static string TickerCanary { get; set; }// = "GME";
        public static string LeiCanary { get; set; }// = "549300505KLOET039L77";
        public static string[] IdCanaries { get; set; } = { CusipCanary, IsinCanary, TickerCanary, LeiCanary };
        //public static string[] Canaries { get; set; } = { "SWU0ZM928", "CGCMT", "SPSUCSN", "SPSU" };
        //public static string SearchTerm { get; set; } = @"REPO CGCMT 2016 C3 D REPO CGCMT 2016 C3 D";
        //public static string[] Canaries { get; set; } = { "EVERGRANDE", "3333-HK", "3333.HK", "3333" };
        //public static string SearchTerm { get; set; } = @"Evergrande";
        //public static string[] Canaries { get; set; } = { "ITV" };
        //public static string SearchTerm { get; set; } = @"ITV PLC -britvic";
        public static string[] Canaries { get; set; } = { "GameStop", "GS2C", "A0HGDX", "36467W109", "US36467W1099", "549300505KLOET039L77", "GME ", " GME" };
        public static string SearchTerm { get; set; } = @"GameStop";
        public static string[] SearchTerms { get; set; } = { @"GameStop", "GME", "36467W109", "549300505KLOET039L77" };
        //public static string[] SearchTerms { get; set; } = { "FTCV", "FTCVU", "FTCVW" };
        public static DateTime Start { get; set; } = DateTime.Parse("2019-10-01");//DateTime.Parse("2021-09-20");//
        public static DateTime End { get; set; } = DateTime.Now;
        public static List<(DateTime start, DateTime end)> SearchDates { get; set; }
        public static bool AddFundInfo { get; set; } = false;
        public static string[] InitialHeaders { get; set; } = { "URL", "File Date" };
        public static string SaveFolder { get; set; } = 
            $@"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\ETF-Reports\" + (Testing ? "TEST" : SearchTerm);
        public static string SaveFile { get; set; } = $@"{SaveFolder}\" + (Testing ? "TEST—" : $"{SearchTerm}—{Start.SecString()}~{End.SecString()} ") +
                $"{Convert.ToDateTime(DateTime.Now.ToString()).ToString("yyyy-MM-dd_HH-mm-ss")}.xlsx";
        public static List<(string start, string end)> Months { get; set; } = new();
        public static Dictionary<string, List<(string rptDate, string fileDate, Search.FormType type, string url)>> Series;
        public static bool Check(this string str, string canary)
            => str is not null && str.Contains(canary, StringComparison.InvariantCultureIgnoreCase);
        public static bool CheckId(this string str, string id) => str is not null && id != default && str.Length == id.Length && str == id;
        public static bool CheckIds(this string str) => str is not null && 
            str.Contains(' ') ? str.Split(' ').Any(x => x.CheckIds())
            : IdCanaries.Any(x => str is not null && str.Trim().CheckId(x));
        public static bool CheckAll(this string str, string[] canaries = null) 
            => (canaries ?? Canaries).Any(x => str.Check(x)) || str.CheckIds();
        public static string SecString(this DateTime dt) 
            => dt.ToString(@"yyyy-MM-dd");
        public static string Acronize(this string s) 
            => string.Join("", s.Split(' ').SelectMany(x => x.Split('.')).SelectMany(x => x.Take(1)));
        public static string Times(this string s, int i) => i == 0 ? "" : string.Concat(Enumerable.Repeat(s, i));
        public static string Times(this char c, int i) => i == 0 ? "" : string.Concat(Enumerable.Repeat(c, i));
        public static Search.FormType GetFormType(string s) 
            => s.StartsWith("NT") ? Search.FormType.NotTimely :
            s.EndsWith('A') ? Search.FormType.Ammended :
            s.EndsWith('P') ? Search.FormType.Regular : 
            Search.FormType.Repeat;
        public static string NewlineValues(IEnumerable<object> list, string s = "\n")
        {
            if (list is null)
                return null;
            else
                return string.Join(s, list);
        }
        public static FileInfo CreateFile(int fileCount, string fldr = null)
        {
            // Create folder & file for output
            Directory.CreateDirectory(fldr ?? SaveFolder);
            return new(GetFileName(fileCount));
        }
        public static string GetFileName(int fileCount) =>
            fileCount == -1 ? SaveFile :
            fileCount == 0 ? SaveFile.Replace("—", "—0—") :
            SaveFile.Remove(SaveFile.IndexOf("—") + 1, 2).Replace("—", "—" + fileCount.ToString() + "—");
        public static XDocument StripNamespace(this XDocument x)
        {
            foreach (var node in x.Root.DescendantsAndSelf().ToArray())
                node.Name = node.Name.LocalName;

            return x;
        }
    }
    public class ProgressBar : IDisposable, IProgress<double>
    {
        private const int blockCount = 10;
        private readonly TimeSpan animationInterval = TimeSpan.FromSeconds(1.0 / 8);
        private const string animation = @"|/-\";

        private readonly Timer timer;

        private double currentProgress = 0;
        private string currentText = string.Empty;
        private bool disposed = false;
        private int animationIndex = 0;

        public ProgressBar()
        {
            timer = new Timer(TimerHandler);

            // A progress bar is only for temporary display in a console window.
            // If the console output is redirected to a file, draw nothing.
            // Otherwise, we'll end up with a lot of garbage in the target file.
            if (!Console.IsOutputRedirected)
                ResetTimer();
        }

        public void Report(double value)
        {
            // Make sure value is in [0..1] range
            value = Math.Max(0, Math.Min(1, value));
            Interlocked.Exchange(ref currentProgress, value);
        }

        private void TimerHandler(object state)
        {
            lock (timer)
            {
                if (disposed) return;

                int progressBlockCount = (int)(currentProgress * blockCount);
                int percent = (int)(currentProgress * 100);
                string text = string.Format("[{0}{1}] {2,3}% {3}",
                    new string('■', progressBlockCount), new string('-', blockCount - progressBlockCount),
                    percent,
                    animation[animationIndex++ % animation.Length]);
                UpdateText(text);

                ResetTimer();
            }
        }

        private void UpdateText(string text)
        {
            // Get length of common portion
            int commonPrefixLength = 0;
            int commonLength = Math.Min(currentText.Length, text.Length);
            while (commonPrefixLength < commonLength && text[commonPrefixLength] == currentText[commonPrefixLength])
                commonPrefixLength++;

            // Backtrack to the first differing character
            StringBuilder outputBuilder = new();
            outputBuilder.Append('\b', currentText.Length - commonPrefixLength);

            // Output new suffix
            outputBuilder.Append(text[commonPrefixLength..]);

            // If the new text is shorter than the old one: delete overlapping characters
            int overlapCount = currentText.Length - text.Length;
            if (overlapCount > 0)
            {
                outputBuilder.Append(' ', overlapCount);
                outputBuilder.Append('\b', overlapCount);
            }

            Console.Write(outputBuilder);
            currentText = text;
        }

        private void ResetTimer() => timer.Change(animationInterval, TimeSpan.FromMilliseconds(-1));

        public void Dispose()
        {
            lock (timer)
            {
                disposed = true;
                UpdateText("[■■■DONE■■■]");
            }
        }

    }
    public class RateLimiter : Stopwatch
    {
        public RateLimiter()
        {
            Count = 1;              // start at 1 with assumption we've made a request
            TimeRemaining = 0;      // remaining time we'll use to sleep to avoid excess rate
        }
        public int Count { get; set; }
        public int TimeRemaining { get; set; }
        public void Limit()
        {
            Count += 1;
            if (Count % 10 == 0)
            {
                Stop();
                Purge();
                Restart();
            }
        }
        public void Purge()
        {
            TimeRemaining = 1100 - (int)ElapsedMilliseconds;
            if (TimeRemaining > 0)
                Thread.Sleep(TimeRemaining);
        }
    }
}
