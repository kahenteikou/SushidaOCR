using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Windows.Media.Ocr;

namespace SushidaOCR
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            OpenQA.Selenium.Chrome.ChromeOptions choptions = new OpenQA.Selenium.Chrome.ChromeOptions();
            choptions.AddExcludedArgument("enable-automation");
            chromekun = new ChromeDriver(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), choptions);
        }

        CancellationTokenSource tokensourcekun;
        bool task_running = false;
        ChromeDriver chromekun;
        private void OpenBT_Click(object sender, RoutedEventArgs e)
        {
            chromekun.Url = "http://typingx0.net/sushida/play.html?soundless";

        }
        async Task<OcrResult> RunWin10Ocr(Windows.Graphics.Imaging.SoftwareBitmap snap)
        {
            // OCRの準備。言語設定を英語にする
            Windows.Globalization.Language language = new Windows.Globalization.Language("en");
            OcrEngine ocrEngine = OcrEngine.TryCreateFromLanguage(language);

            // OCRをはしらせる
            var ocrResult = await ocrEngine.RecognizeAsync(snap);
            return ocrResult;
        }
        private async void ERunning(CancellationToken token)
        {
            //var bodykun = chromekun.FindElementByTagName("body");
            var bodykun = chromekun.FindElement(By.TagName("body"));
            //var gazouElementkun = chromekun.FindElementById("#canvas");
            var gazouElementkun = chromekun.FindElement(By.Id("#canvas"));
            CroppedBitmap cloppedkun;
            CroppedBitmap cloppedkunii;
            Windows.Graphics.Imaging.SoftwareBitmap cropped_simage;
            OcrResult ocrresult;
            string ocrtext;
            string buf_str = "";
            while (!token.IsCancellationRequested)
            {
                try
                {
                    cloppedkun = GetElementScreenShot(chromekun, gazouElementkun);
                    cloppedkunii = new CroppedBitmap(ConvertCropToBitmap(cloppedkun, new PngBitmapEncoder()), new Int32Rect(266, 310, 219, 51));
                    cropped_simage = await ConvertCropToUWPBitmap(cloppedkunii);
                    ocrresult = await RunWin10Ocr(cropped_simage);
                    ocrtext = ocrresult.Text;
                    if(ocrtext != buf_str && ocrtext != "")
                    {
                        buf_str = ocrtext;
                        bodykun.SendKeys(ocrtext);
                    }
                    Thread.Sleep(700);

                }
                catch (Exception e)
                {
                    /*
                    task_running = false;
                    System.Console.WriteLine(e.Message);
                    Console.WriteLine("Press enter key");
                    Console.ReadLine();
                    return;*/
                }
            }
        }

        private void StartBT_Click(object sender, RoutedEventArgs e)
        {
            if (task_running) return;
            task_running = true;
            tokensourcekun = new CancellationTokenSource();
            var token = tokensourcekun.Token;
            Task.Run(() => ERunning(token));
            

        }
        public static BitmapImage ConvertCropToBitmap(CroppedBitmap cbit,BitmapEncoder encoder)
        {
            var bmpImage=new BitmapImage();
            using(var srcMS=new MemoryStream())
            {
                encoder.Frames.Add(BitmapFrame.Create(cbit));
                encoder.Save(srcMS);
                srcMS.Position = 0;
                using(var destMs = new MemoryStream(srcMS.ToArray()))
                {
                    bmpImage.BeginInit();
                    bmpImage.StreamSource = destMs;
                    bmpImage.CacheOption= BitmapCacheOption.OnLoad;
                    bmpImage.EndInit();
                    bmpImage.Freeze();

                }
            }
            return bmpImage;
        }
        public static async Task<Windows.Graphics.Imaging.SoftwareBitmap> ConvertCropToUWPBitmap(CroppedBitmap cbit)
        {
            var bmpImage = new BitmapImage();
            var encoder = new PngBitmapEncoder();
            Windows.Graphics.Imaging.SoftwareBitmap softwareBitmap;
            using (var srcMS = new MemoryStream())
            {
                encoder.Frames.Add(BitmapFrame.Create(cbit));
                encoder.Save(srcMS);
                srcMS.Position = 0;
                using (var destMs = new MemoryStream(srcMS.ToArray()))
                {
                    using(var randomStream = destMs.AsRandomAccessStream())
                    {
                        Windows.Graphics.Imaging.BitmapDecoder decoder = await Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(randomStream);
                        softwareBitmap = await decoder.GetSoftwareBitmapAsync();
                    }
                }
            }
            return softwareBitmap;
        }
        public static CroppedBitmap GetElementScreenShot(IWebDriver driver,IWebElement element)
        {
            Screenshot scrsh=((ITakesScreenshot)driver).GetScreenshot();
            var imagekunData = new BitmapImage();
            using (var mem = new MemoryStream(scrsh.AsByteArray))
            {
                mem.Position = 0;
                imagekunData.BeginInit();
                imagekunData.CreateOptions = BitmapCreateOptions.PreservePixelFormat;
                imagekunData.CacheOption = BitmapCacheOption.OnLoad;
                imagekunData.UriSource = null;
                imagekunData.StreamSource = mem;
                imagekunData.EndInit();

            }
            imagekunData.Freeze();
            CroppedBitmap croppedkun = new CroppedBitmap(imagekunData,
                new Int32Rect(element.Location.X, element.Location.Y, element.Size.Width, element.Size.Height));
            return croppedkun;
        }

        private void StopBT_Click(object sender, RoutedEventArgs e)
        {
            if (tokensourcekun != null) tokensourcekun.Cancel();

            task_running = false;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            chromekun.Quit();
        }
    }
}
