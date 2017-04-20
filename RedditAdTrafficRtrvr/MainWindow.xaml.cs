using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Windows;
using HtmlAgilityPack;

// ReSharper disable AssignNullToNotNullAttribute
// ReSharper disable ConditionIsAlwaysTrueOrFalse

namespace RedditAdBase
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public CookieContainer Cookies;

        public MainWindow()
        {
            InitializeComponent();

            GeneralProgressBar.Visibility = Visibility.Hidden;
        }

        /// <summary>
        /// Register events to MainWindow status textblock and Log file
        /// </summary>
        /// <param name="type">Log message type (INFO, WARNING, ERROR)</param>
        /// <param name="msg">Message to display</param>
        private void Log(string type, string msg)
        {
            var status = string.Empty;

            switch (type)
            {
                case "INFO":

                    status += $"[INFO] {DateTime.Now.ToString(CultureInfo.CurrentCulture)} - {msg}";

                    break;

                case "ERROR":

                    status += $"[ERROR] {DateTime.Now.ToString(CultureInfo.CurrentCulture)} - {msg}";

                    break;

                case "WARNING":

                    status += $"[WARNING] {DateTime.Now.ToString(CultureInfo.CurrentCulture)} - {msg}";

                    break;
            }

            StatusTextBlock.Text += status + Environment.NewLine;
            StatusTextBlock.ScrollToEnd();

            using (var writer = new StreamWriter("Log.txt", true))
            {
                writer.Write(status + Environment.NewLine);
            }
        }

        /// <summary>
        /// Get a HtmlDocument from the supplied Url
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        private HtmlDocument GetHtmlDocument(string url)
        {
            var doc = new HtmlDocument();

            var docRequest = WebRequest.Create(url) as HttpWebRequest;

            if (docRequest != null)
            {
                docRequest.CookieContainer = Cookies;
                docRequest.Method = "GET";
                docRequest.Accept =
                    "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";

                var docResponse = (HttpWebResponse)docRequest.GetResponse();

                if (docResponse.StatusCode == HttpStatusCode.OK)
                    using (var s = docResponse.GetResponseStream())
                    {
                        using (var sr = new StreamReader(s, Encoding.GetEncoding(name: docResponse.CharacterSet)))
                        {
                            doc.Load(sr);
                        }
                    }
            }

            return doc;
        }
    }
}
