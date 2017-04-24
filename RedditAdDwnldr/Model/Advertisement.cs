using System.Collections.Generic;

namespace RedditAdDwnldr.Model
{
    public class Advertisement
    {
        public int AdvertisementNumber { get; set; }
        public string ThumbnailName { get; set; }
        public string Title { get; set; }
        public string Url { get; set; }
        public bool DisableComments { get; set; }
        public bool SendComments { get; set; }
        public string RedditAdId { get; set; }
        public List<Campaign> Campaigns { get; set; }
    }
}