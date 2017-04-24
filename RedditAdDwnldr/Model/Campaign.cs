using System;

namespace RedditAdDwnldr.Model
{
    public class Campaign
    {
        public Advertisement Advertisement { get; set; }
        public string Target { get; set; }
        public string TargetDetail { get; set; }
        public string Location { get; set; }
        public string Location2 { get; set; }
        public string Platform { get; set; }
        public decimal Budget { get; set; }
        public bool BudgetOptionDeliverFast { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public bool OptionExtend { get; set; }
        public decimal PricingCpm { get; set; }
    }
}