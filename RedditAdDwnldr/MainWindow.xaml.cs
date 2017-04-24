using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using HtmlAgilityPack;
using Microsoft.Win32;
using OfficeOpenXml;
using RedditAdDwnldr.Model;
using OfficeOpenXml.Style;

// ReSharper disable AssignNullToNotNullAttribute
// ReSharper disable ConditionIsAlwaysTrueOrFalse

namespace RedditAdDwnldr
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public CookieContainer Cookies;
        public List<Advertisement> Ads;

        public MainWindow()
        {
            InitializeComponent();

            GeneralProgressBar.Visibility = Visibility.Hidden;
            DownloadAdDataButton.IsEnabled = false;
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
                        using (var sr = new StreamReader(s, Encoding.GetEncoding(docResponse.CharacterSet)))
                        {
                            doc.Load(sr);
                        }
                    }
            }

            return doc;
        }

        private void DestinationFilePickerButton_Click(object sender, RoutedEventArgs e)
        {
            var openDialog = new SaveFileDialog()
            {
                Filter = "Excel files (*.xlsx; *.xlsm) | *.xlsx; *.xlsm"
            };

            if (openDialog.ShowDialog() == true)
            {
                DestinationFileTextBox.Text = openDialog.FileName;
                DownloadAdDataButton.IsEnabled = true;
            }
        }

        private async void DownloadAdDataButton_Click(object sender, RoutedEventArgs e)
        {
            Log("INFO", "Starting ad data gathering process...");
            GeneralProgressBar.Visibility = Visibility.Visible;

            var doc = new HtmlDocument();
            var path = DestinationFileTextBox.Text;
            var stop = false;
            Ads = new List<Advertisement>();

            var task = Task.Factory.StartNew(() =>
            {
                doc = GetHtmlDocument("https://www.reddit.com/promoted/");
            });

            await task;

            do
            {
                var adNodes = doc.DocumentNode.SelectNodes("//div[contains(@class, 'promotedlink')]");
                var counter = 0;

                foreach (var ad in adNodes)
                {
                    var adId = ad.GetAttributeValue("data-fullname", "").Split('_')[1];

                    Log("INFO", $"Downloading ad ID {adId}");

                    task = Task.Factory.StartNew(() =>
                    {
                        Advertisement newAd = new Advertisement
                        {
                            AdvertisementNumber = counter++,
                            Title = ad.SelectSingleNode("//a[@data-event-action='title']").InnerText,
                            Url = $"https://www.reddit.com/promoted/edit_promo/{ad.GetAttributeValue("data-fullname", "").Split('_')[1]}",
                            RedditAdId = ad.GetAttributeValue("data-fullname", "").Split('_')[1],
                            Campaigns = new List<Campaign>()
                        };

                        var adDoc = GetHtmlDocument(newAd.Url);
                        var disabledCommentsChecked = adDoc.DocumentNode
                            .SelectSingleNode("//input[@name='disable_comments']")
                            .GetAttributeValue("checked", "");
                        var sendRepliesChecked = adDoc.DocumentNode
                            .SelectSingleNode("//input[@name='disable_comments']")
                            .GetAttributeValue("checked", "");
                        
                        newAd.DisableComments =
                            !string.IsNullOrEmpty(disabledCommentsChecked) && disabledCommentsChecked == "checked";

                        newAd.SendComments = !string.IsNullOrEmpty(sendRepliesChecked) && sendRepliesChecked == "checked";

                        var campaigns = adDoc.DocumentNode.SelectNodes("//tr[contains(@class, 'campaign-row')]");

                        if (campaigns != null)
                        {
                            foreach (var campaign in campaigns)
                            {
                                var camp = new Campaign
                                {
                                    Target = campaign.GetAttributeValue("data-targeting-type", ""),
                                    TargetDetail = campaign.GetAttributeValue("data-targeting", ""),
                                    Location = campaign.GetAttributeValue("data-country", ""),
                                    Location2 = campaign.GetAttributeValue("data-region", ""),
                                    Platform = campaign.GetAttributeValue("data-platform", ""),
                                    Budget = Convert.ToDecimal(campaign.GetAttributeValue("data-total_budget_dollars", ""), new CultureInfo("en-US")),
                                    BudgetOptionDeliverFast =
                                        Convert.ToBoolean(campaign.GetAttributeValue("data-no_daily_budget", "")),
                                    Start = DateTime.ParseExact(campaign.GetAttributeValue("data-startdate", ""), "MM/dd/yyyy", CultureInfo.InvariantCulture),
                                    End = DateTime.ParseExact(campaign.GetAttributeValue("data-enddate", ""), "MM/dd/yyyy", CultureInfo.InvariantCulture),
                                    OptionExtend =
                                        Convert.ToBoolean(campaign.GetAttributeValue("data-is_auto_extending", "")),
                                    PricingCpm = Convert.ToDecimal(campaign.GetAttributeValue("data-bid_dollars", ""), new CultureInfo("en-US"))
                                };


                                newAd.Campaigns.Add(camp);
                            }
                        }

                        Ads.Add(newAd);
                    });

                    await task;
                }

                if (doc.DocumentNode.SelectSingleNode("//a[contains(text(), 'next')]") != null)
                {
                    Log("INFO", (doc.DocumentNode.SelectSingleNode("//a[contains(text(), 'next')]") != null).ToString());
                    Log("INFO", (stop).ToString());
                    Log("INFO", "Handling pagination");

                    var newUrl =
                        doc.DocumentNode.SelectSingleNode("//a[contains(text(), 'next')]").GetAttributeValue("href", "");

                    await Task.Factory.StartNew(() =>
                    {
                        doc = GetHtmlDocument(newUrl);
                    });
                }
                else
                    stop = true;
            }
            while (stop != true);

            GenerateExcel(path);

            GeneralProgressBar.Visibility = Visibility.Hidden;
            Log("INFO", "Ended URL gathering process...");
        }

        private void GenerateExcel(string path)
        {
            FileInfo fi = new FileInfo(path);

            Log("INFO", "Generating file");

            using (var package = new ExcelPackage(fi))
            {
                ExcelWorksheet adsWorksheet = package.Workbook.Worksheets.Add("advertisements");
                ExcelWorksheet campaignsWorksheet = package.Workbook.Worksheets.Add("campaigns");

                adsWorksheet.Cells[1, 1].Value = "ADVERTISEMENT NUMBER";
                adsWorksheet.Cells[1, 2].Value = "THUMBNAIL NAME";
                adsWorksheet.Cells[1, 3].Value = "TITLE URL";
                adsWorksheet.Cells[1, 4].Value = "OPTION_DISABLECOMMENTS";
                adsWorksheet.Cells[1, 5].Value = "OPTION_SENDCOMMENTS";

                campaignsWorksheet.Cells[1, 1].Value = "WHICH ADVERTISEMENT?";
                campaignsWorksheet.Cells[1, 2].Value = "TARGET";
                campaignsWorksheet.Cells[1, 3].Value = "TARGET_DETAIL";
                campaignsWorksheet.Cells[1, 4].Value = "LOCATION";
                campaignsWorksheet.Cells[1, 5].Value = "LOCATION_2";
                campaignsWorksheet.Cells[1, 6].Value = "PLATFORM";
                campaignsWorksheet.Cells[1, 7].Value = "BUDGET";
                campaignsWorksheet.Cells[1, 8].Value = "BUDGET_OPTION_DELIVERFAST";
                campaignsWorksheet.Cells[1, 9].Value = "START";
                campaignsWorksheet.Cells[1, 10].Value = "END";
                campaignsWorksheet.Cells[1, 11].Value = "OPTION_EXTEND";
                campaignsWorksheet.Cells[1, 12].Value = "PRICINGCPM";

                using (var range = adsWorksheet.Cells[1, 1, 1, 5])
                {
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.Black);
                    range.Style.Font.Color.SetColor(Color.White);
                    range.Style.ShrinkToFit = false;
                }

                using (var range = campaignsWorksheet.Cells[1, 1, 1, 12])
                {
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.Black);
                    range.Style.Font.Color.SetColor(Color.White);
                    range.Style.ShrinkToFit = false;
                }

                var row = 2;
                var campaignRow = 2;

                foreach (var ad in Ads)
                {
                    var column = 1;

                    adsWorksheet.Cells[row, column++].Value = ad.AdvertisementNumber;
                    adsWorksheet.Cells[row, column++].Value = ad.ThumbnailName;
                    adsWorksheet.Cells[row, column++].Value = ad.Title;
                    adsWorksheet.Cells[row, column++].Value = ad.DisableComments ? 1 : 0;
                    adsWorksheet.Cells[row, column].Value = ad.SendComments ? 1 : 0;

                    foreach (var campaign in ad.Campaigns)
                    {
                        var campaignColumn = 1;

                        campaignsWorksheet.Cells[campaignRow, campaignColumn++].Value = ad.AdvertisementNumber;
                        campaignsWorksheet.Cells[campaignRow, campaignColumn++].Value = campaign.Target;
                        campaignsWorksheet.Cells[campaignRow, campaignColumn++].Value = campaign.TargetDetail;
                        campaignsWorksheet.Cells[campaignRow, campaignColumn++].Value = campaign.Location;
                        campaignsWorksheet.Cells[campaignRow, campaignColumn++].Value = campaign.Location2;
                        campaignsWorksheet.Cells[campaignRow, campaignColumn++].Value = campaign.Platform;
                        campaignsWorksheet.Cells[campaignRow, campaignColumn].Style.Numberformat.Format = "0.00";
                        campaignsWorksheet.Cells[campaignRow, campaignColumn++].Value = campaign.Budget;
                        campaignsWorksheet.Cells[campaignRow, campaignColumn++].Value = campaign.BudgetOptionDeliverFast ? 1 : 0;
                        campaignsWorksheet.Cells[campaignRow, campaignColumn].Style.Numberformat.Format = "MM/dd/yyyy";
                        campaignsWorksheet.Cells[campaignRow, campaignColumn++].Value = campaign.Start;
                        campaignsWorksheet.Cells[campaignRow, campaignColumn].Style.Numberformat.Format = "MM/dd/yyyy";
                        campaignsWorksheet.Cells[campaignRow, campaignColumn++].Value = campaign.End;
                        campaignsWorksheet.Cells[campaignRow, campaignColumn++].Value = campaign.OptionExtend ? 1 : 0;
                        campaignsWorksheet.Cells[campaignRow, campaignColumn].Style.Numberformat.Format = "0.00";
                        campaignsWorksheet.Cells[campaignRow, campaignColumn].Value = campaign.PricingCpm;

                        campaignRow++;
                    }

                    row++;
                }

                adsWorksheet.Cells[adsWorksheet.Dimension.Address].AutoFitColumns();
                campaignsWorksheet.Cells[campaignsWorksheet.Dimension.Address].AutoFitColumns();

                package.Save();
            }
        }
    }
}
