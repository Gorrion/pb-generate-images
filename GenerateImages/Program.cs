using System;
using System.Web;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CefSharp.Internals;
using CefSharp;
using CefSharp.OffScreen;
using System.Threading;
using System.Collections.Generic;
using System.Drawing.Imaging;
using ClosedXML.Excel;
using System.Runtime.InteropServices;
using System.Net;

namespace GenerateImages
{
    class RowInfo
    {
        public int RowIndex;
        public string CellFilesPath { set { _CellFilesPath = HttpUtility.UrlDecode(value); } get { return _CellFilesPath; } }

        private string _CellFilesPath;
        public string[] GetFilesPath()
        {
            return _CellFilesPath.Trim().Split('|');
        }

        public string CellUrls;
        public string[] GetUrls()
        {
            return CellUrls.Trim().Split('|');
        }
    }

    class Program
    {
        // private const string TestUrl = "http://82.209.219.185:3030/";
        private static int PageWidth = 1302;
        private static int PageHeight = 732;

        static System.Timers.Timer RunTimer;
        static List<string> ServersNames = new List<string>
        {
            Properties.Settings.Default.ServerLeftName,
            Properties.Settings.Default.ServerMiddleName,
            Properties.Settings.Default.ServerRightName,
        };

        public static int Main(string[] args)
        {
            var settings = new CefSettings()
            {
                //By default CefSharp will use an in-memory cache, you need to specify a Cache Folder to persist data
                CachePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CefSharp\\Cache"),
                LogSeverity = LogSeverity.Disable
            };

            //Perform dependency check to make sure all relevant resources are in our output directory.
            Cef.Initialize(settings, performDependencyCheck: true, browserProcessHandler: null);
            AppDomain.CurrentDomain.ProcessExit += CurrentDomain_ProcessExit;

            handler = new ConsoleEventDelegate(ConsoleEventCallback);
            SetConsoleCtrlHandler(handler, true);

            RunTimer = new System.Timers.Timer(1000 * 60 * 30);
            RunTimer.Elapsed += RunTimer_Elapsed;
            RunTimer.AutoReset = true;
            RunTimer.Enabled = true;
            RunTimer.Start();
            MainEvent();

            while (true)
            {
                Console.Write("genimg> ");
                string line = Console.ReadLine();
            }
        }

        private static void CurrentDomain_ProcessExit(object sender, EventArgs e)
        {
            try { Cef.Shutdown(); }
            catch { }
        }

        private static void RunTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            MainEvent();
        }

        static void MainEvent()
        {
            try
            {
                Console.WriteLine("Запуск обновления изображений. Время: {0}.", DateTime.Now);

                string filePathExcell = Properties.Settings.Default.SourceFilePath;
                if (string.IsNullOrEmpty(filePathExcell) || !File.Exists(filePathExcell)) { filePathExcell = "ПриложениеWall.xlsx"; }

                var workbook = new XLWorkbook(filePathExcell);
                var ws1 = workbook.Worksheet(1);

                var LstRows = new List<RowInfo>();
                var rows = ws1.Rows().ToList();

                for (var i = rows.Count - 1; i > 1; i--)
                {
                    var cellPath = rows[i].Cell(8).Value;
                    var cellUrl = rows[i].Cell(35).Value;

                    if (cellPath != null && cellUrl != null && !string.IsNullOrEmpty(cellPath.ToString()) && !string.IsNullOrEmpty(cellUrl.ToString()))
                    {
                        var rowInfo = new RowInfo { RowIndex = i, CellFilesPath = cellPath.ToString(), CellUrls = cellUrl.ToString() };
                        LstRows.Add(rowInfo);

                        var urls = rowInfo.GetUrls();
                        var filePaths = rowInfo.GetFilesPath();

                        for (var j = 0; j < urls.Count(); j++)
                        {
                            try
                            {
                                Console.WriteLine("Обработка URL: {0}.", urls[j]);

                                var imgPath = filePaths[j];
                                var imgPathSpl = imgPath.Split('/').Where(x => !string.IsNullOrEmpty(x)).ToArray();
                                var relPath = string.Join("\\", imgPathSpl.Take(imgPathSpl.Length - 1).Skip(2));

                                var clearList = new List<string>();
                                var task = MainAsync("cachePath1", urls[j]);
                                using (var bitmap = task.Result)
                                {
                                    var images = SplitImage(bitmap, PageHeight);
                                    int len = images.Count;

                                    for (int x = images.Count - 1; x >= 0; x--)
                                    {
                                        var bmp = images[x];
                                        using (var b = new Bitmap(bmp.Width, bmp.Height))
                                        {
                                            b.SetResolution(bmp.HorizontalResolution, bmp.VerticalResolution);
                                            using (var g = Graphics.FromImage(b))
                                            {
                                                g.Clear(Color.White);
                                                g.DrawImageUnscaled(bmp, 0, 0);
                                            }
                                            images.RemoveAt(x);
                                            bmp.Dispose();

                                            string filename = (x + 1) + ".jpg";
                                            try
                                            {
                                                foreach (var srv in ServersNames)
                                                {
                                                    if (string.IsNullOrWhiteSpace(srv)) { continue; }
                                                    var srvDirPath = Path.Combine("\\\\" + srv, relPath).ToString();
#if DEBUG
                                                    srvDirPath = Path.Combine("D:\\", relPath).ToString();
#endif
                                                    if (!clearList.Contains(srvDirPath))
                                                    {
                                                        clearList.Add(srvDirPath);
                                                        try
                                                        {
                                                            var srvDir = new DirectoryInfo(srvDirPath);
                                                            if (srvDir.Exists) { srvDir.Empty(); }
                                                            else { Directory.CreateDirectory(srvDirPath); }
                                                        }
                                                        catch (Exception ex) { ConsoleLogRed("Ошибка очистки или создания папки по пути {0}. Ошибка: {1}", srvDirPath, ex.Message); }
                                                    }

                                                    try
                                                    {
                                                        b.Save(Path.Combine(srvDirPath, filename), ImageFormat.Jpeg);
                                                        ConsoleLog("Сохранен файл {0} на сервер {1}", ConsoleColor.DarkGreen, Path.Combine(srvDirPath, filename), srv);
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        ConsoleLogRed("Ошибка сохранения файлов по пути {0}. Ошибка: {1}", Path.Combine(srvDirPath, filename), ex.Message);
                                                    }
                                                }
                                            }
                                            finally { b.Dispose(); }
                                        }
                                    }

                                    var cellNextPage = rows[i].Cell(7).Value;
                                    if (cellNextPage != null && cellNextPage.ToString().StartsWith("7."))
                                    {
                                        try
                                        {
                                            var curValue = ws1.Row(i + 1).Cell(34).Value;
                                            if (curValue == null || curValue.ToString() != len.ToString())
                                            {
                                                ws1.Row(i + 1).Cell(34).Value = len.ToString();
                                                workbook.Save();
                                                RefreshProgrammExcell();
                                            }
                                        }
                                        catch (Exception ex) { ConsoleLogRed("Ошибка сохранентя excel {0}", ex.Message); }
                                    }
                                }
                            }
                            catch (Exception ex) { ConsoleLogRed("Ошибка работы по url {0}. Ошибка {1}", urls[j], ex.Message); }
                        }
                    }
                }
                ConsoleLog("Сохранение завершено...", ConsoleColor.Green);
            }
            catch (Exception ex) { ConsoleLogRed("Ошибка работы программы {0}", ex.Message); }
            return;
        }

        static bool ConsoleEventCallback(int eventType)
        {
            if (eventType == 2)
            {
                try { Cef.Shutdown(); }
                catch { }
                Console.WriteLine("Console window closing, death imminent");
            }
            return false;
        }
        static ConsoleEventDelegate handler;   // Keeps it from getting garbage collected
                                               // Pinvoke
        private delegate bool ConsoleEventDelegate(int eventType);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetConsoleCtrlHandler(ConsoleEventDelegate callback, bool add);

        private static async Task<Bitmap> MainAsync(string cachePath, string url)
        {
            var browserSettings = new BrowserSettings();
            //Reduce rendering speed to one frame per second so it's easier to take screen shots
            browserSettings.WindowlessFrameRate = 1;

            var requestContextSettings = new RequestContextSettings { CachePath = cachePath };
            // RequestContext can be shared between browser instances and allows for custom settings
            // e.g. CachePath
            using (var requestContext = new RequestContext(requestContextSettings))
            using (var browser = new ChromiumWebBrowser(url, browserSettings, requestContext))
            {
                browser.Size = new Size() { Width = PageWidth, Height = PageHeight };

                await LoadPageAsync(browser);
                Thread.Sleep(4000);
                //Check preferences on the CEF UI Thread
                //await Cef.UIThreadTaskFactory.StartNew(delegate
                //{
                //    var preferences = requestContext.GetAllPreferences(true);
                //    //Check do not track status
                //    var doNotTrack = (bool)preferences["enable_do_not_track"];
                //    Debug.WriteLine("DoNotTrack:" + doNotTrack);
                //});

                var onUi = Cef.CurrentlyOnThread(CefThreadIds.TID_UI);

                var task = browser.GetMainFrame().EvaluateScriptAsync("(function() { document.body.style.overflow = 'hidden'; var body = document.body, html = document.documentElement; return  Math.max( body.scrollHeight, body.offsetHeight, html.clientHeight, html.scrollHeight, html.offsetHeight ); })();", null);
                await task.ContinueWith(t =>
                 {
                     if (!t.IsFaulted)
                     {
                         var response = t.Result;
                         double height = 0;
                         if (response.Success && response.Result != null && double.TryParse(response.Result.ToString(), out height))
                         {
                             if (height > PageHeight) { browser.Size = new Size() { Width = PageWidth, Height = (int)height }; }
                         }
                     }
                 });

                var task2 = browser.GetMainFrame().EvaluateScriptAsync("(function() { document.body.style.overflow = 'hidden'; })();", null);
                Bitmap result = null;
                await task2.ContinueWith(t =>
                {
                    if (!t.IsFaulted)
                    {
                        result = browser.ScreenshotAsync(true).Result;//.ContinueWith(DisplayBitmap).Wait();
                        browser.Delete();
                    }
                });

                return result;
            }
        }

        public static Task LoadPageAsync(IWebBrowser browser, string address = null)
        {
            //If using .Net 4.6 then use TaskCreationOptions.RunContinuationsAsynchronously
            //and switch to tcs.TrySetResult below - no need for the custom extension method
            var tcs = new TaskCompletionSource<bool>();
            EventHandler<LoadingStateChangedEventArgs> handler = null;
            handler = (sender, args) =>
            {
                //Wait for while page to finish loading not just the first frame
                if (!args.IsLoading)
                {
                    browser.LoadingStateChanged -= handler;
                    Thread.Sleep(3000);
                    //This is required when using a standard TaskCompletionSource
                    //Extension method found in the CefSharp.Internals namespace
                    tcs.TrySetResultAsync(true);
                }
            };

            browser.LoadingStateChanged += handler;
            if (!string.IsNullOrEmpty(address)) { browser.Load(address); }
            return tcs.Task;
        }

        static IList<Bitmap> SplitImage(Bitmap sourceBitmap, int splitHeight)
        {
            Size dimension = sourceBitmap.Size;
            if (sourceBitmap.Size.Height <= splitHeight) { return new List<Bitmap> { sourceBitmap }; }

            IList<Bitmap> results = new List<Bitmap>() { new Bitmap(sourceBitmap.Size.Width, splitHeight, PixelFormat.Format32bppArgb) };
            var latHeight = 0;
            var emptyPageRowCount = 0;
            for (var y = 0; y < dimension.Height; y++)
            {
                if (latHeight == splitHeight - 1)
                {
                    var emptyRowSequence = 0;
                    for (var j = 0; j < 100; j++)
                    {
                        var emptyInRowPixel = 0;
                        for (var x = 0; x < sourceBitmap.Size.Width; x++)
                        {
                            if (sourceBitmap.GetPixel(x, y - j).ToArgb() <= 0) { emptyInRowPixel++; }
                        }

                        if (emptyInRowPixel == sourceBitmap.Size.Width) { emptyRowSequence++; }
                        else { emptyRowSequence = 0; }

                        if (emptyRowSequence > 12)
                        {
                            for (var j2 = 0; j2 < j; j2++)
                            {
                                for (var x = 0; x < sourceBitmap.Size.Width; x++) 
                                {
                                    results.Last().SetPixel(x, latHeight - j2, Color.FromArgb(0, 255, 255, 255));
                                }
                            }
                            if (dimension.Height - y > 12) { results.Add(new Bitmap(sourceBitmap.Size.Width, splitHeight, PixelFormat.Format32bppArgb)); }
                            latHeight = 0;
                            emptyPageRowCount = 0;
                            y = y - j - 1;
                            break;
                        }
                    }
                }

                var emptyPxInRowCount = 0;
                for (var x = 0; x < sourceBitmap.Size.Width; x++)
                {
                    var p = sourceBitmap.GetPixel(x, y);
                    if (p.ToArgb() <= 0) { emptyPxInRowCount++; }
                    results.Last().SetPixel(x, latHeight, p);
                }
                if (emptyPxInRowCount == sourceBitmap.Size.Width) { emptyPageRowCount++; }
                latHeight++;
            }

            if (emptyPageRowCount == latHeight && results.Count > 1) { results.RemoveAt(results.Count - 1); }
            return results;
        }

        public static void RefreshProgrammExcell()
        {
            string html = string.Empty;
            string url = Properties.Settings.Default.MainProgramHttpUrl + "?action=refresh";
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (Stream stream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(stream))
                {
                    html = reader.ReadToEnd();
                }

                Console.WriteLine("Обновление экселя в программе управления пандорой");
            }
            catch (Exception ex) { ConsoleLogRed("Неудача при обращении к основной программе по http"); }
        }

        //static IList<Bitmap> SplitImage(Bitmap sourceBitmap, int splitHeight)
        //{
        //    //sourceBitmap.

        //    Size dimension = sourceBitmap.Size;
        //    Rectangle sourceRectangle = new Rectangle(0, 0, dimension.Width, splitHeight);
        //    Rectangle targetRectangle = new Rectangle(0, 0, dimension.Width, splitHeight);

        //    IList<Bitmap> results = new List<Bitmap>();

        //    while (sourceRectangle.Top < dimension.Height)
        //    {
        //        Bitmap pageBitmap = new Bitmap(targetRectangle.Size.Width, sourceRectangle.Bottom < dimension.Height ?
        //            targetRectangle.Size.Height
        //            :
        //            dimension.Height - sourceRectangle.Top, PixelFormat.Format32bppArgb);

        //        using (Graphics g = Graphics.FromImage(pageBitmap))
        //        {
        //            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear;
        //            g.DrawImage(sourceBitmap, targetRectangle, sourceRectangle, GraphicsUnit.Pixel);
        //        }
        //        sourceRectangle.Y += sourceRectangle.Height;
        //        results.Add(pageBitmap);
        //    }

        //    return results;
        //}

        //static void Main(string[] args)
        //{
        //    Thread t = new Thread(SaveImage);
        //    t.SetApartmentState(ApartmentState.STA);
        //    t.Start();
        //    t.Join();

        //    Console.ReadLine();
        //}

        ////private static WebBrowser _wb = null;

        //private static void SaveImage()
        //{
        //    var wb = new WebBrowser() { Width = PageWidth, Height = PageHeight };

        //    wb.DocumentCompleted += (s, e) => {
        //       // Thread.Sleep(10000);
        //    };
        //    wb.Navigated += (s, e) => {
        //    };

        //    wb.AllowNavigation = true;
        //    wb.ScrollBarsEnabled = false;
        //    wb.ScriptErrorsSuppressed = true;
        //    wb.Navigate("https://mail.ru");

        //    wb.ProgressChanged += new WebBrowserProgressChangedEventHandler((s,e) => {
        //        int max = (int)Math.Max(e.MaximumProgress, e.CurrentProgress);
        //        int min = (int)Math.Min(e.MaximumProgress, e.CurrentProgress);
        //        if (min.Equals(max))
        //        {
        //            Thread.Sleep(5000);
        //            //Run your code here when page is actually 100% complete
        //            Bitmap bitmap = new Bitmap(wb.Width, wb.Height);
        //            wb.DrawToBitmap(bitmap, new Rectangle(0, 0, wb.Width, wb.Height));
        //            wb.Dispose();
        //            bitmap.Save("img.png", System.Drawing.Imaging.ImageFormat.Png);
        //        }
        //    });
        //    // wb.Navigate("https://yandex.ru/");
        //    while (wb != null && !wb.IsDisposed && wb.ReadyState != WebBrowserReadyState.Complete)
        //        Application.DoEvents();
        //    //while (wb.ReadyState != WebBrowserReadyState.Complete)
        //    //{
        //    //    Thread.Sleep(400);
        //    //}         
        //}
        //async Task SI()
        //{
        //    using (var requestContext = new RequestContext(requestContextSettings))
        //    using (var browser = new ChromiumWebBrowser("https://mail.ru", browserSettings, requestContext))
        //    {
        //        //if (zoomLevel > 1)
        //        //{
        //        //    browser.FrameLoadStart += (s, argsi) =>
        //        //    {
        //        //        var b = (ChromiumWebBrowser)s;
        //        //        if (argsi.IsMainFrame)
        //        //        {
        //        //            b.SetZoomLevel(zoomLevel);
        //        //        }
        //        //    };
        //        //}

        //        await LoadPageAsync(browser); // loads the first page

        //        // once we are on the page, we want to inject html and take a screen shot
        //        for (int i = 0; i < 26; i++)
        //        {
        //            var html = "<div style=\"font-size:300px !important;\">" + GetColNameFromIndex(i) + "</div>";
        //            // inject content
        //            var htmlEncoded = html.Replace("\"", "\\\"")
        //                .Replace(Environment.NewLine, "\\" + Environment.NewLine);

        //            var javascript3 = "(function() { $('#content').html(\"" + htmlEncoded + "\");   return true; })()";
        //            await browser.EvaluateScriptAsync(javascript3);

        //            var bitmap = await browser.ScreenshotAsync();
        //            await browser.EvaluateScriptAsync("(function() {  $('#content').empty(); return true; })()");
        //            SaveBitmap(bitmap);
        //        }
        //    }
        //}


        static object _lockThis = new object();
        public static void ConsoleError(Exception ex)
        {
            lock (_lockThis)
            {
                ConsoleColor currentBackground = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write(ex.Message);
                Console.ForegroundColor = currentBackground;
                Console.WriteLine();
                Console.Write("genimg> ");
            }
        }

        public static void ConsoleLog(string message, params string[] paramsArr)
        {
            lock (_lockThis)
            {
                if (paramsArr.Length > 0) { message = string.Format(message, paramsArr); }
                Console.Write(message);
                Console.WriteLine();
                Console.Write("genimg> ");
            }
        }

        public static void ConsoleLog(string message, ConsoleColor color, params string[] paramsArr)
        {
            lock (_lockThis)
            {
                ConsoleColor currentBackground = Console.ForegroundColor;
                Console.ForegroundColor = color;
                Console.Write(message, paramsArr);
                Console.ForegroundColor = currentBackground;
                Console.WriteLine();
                Console.Write("genimg> ");
            }
        }

        public static void ConsoleLogRed(string message, params string[] paramsArr) { ConsoleLog(message, ConsoleColor.Red, paramsArr); }
    }

    static class FuncExt
    {
        public static void Empty(this DirectoryInfo directory)
        {
            foreach (FileInfo file in directory.GetFiles()) file.Delete();
            foreach (DirectoryInfo subDirectory in directory.GetDirectories()) subDirectory.Delete(true);
        }
    }
}
