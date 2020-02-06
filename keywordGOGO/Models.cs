using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace keywordGOGO
{
    class CustomerLink
    {
        public int ManagerEnable { get; set; }

        public long ManagerCustomerId { get; set; }

        public string ManagerLoginId { get; set; }

        public string ManagerCompanyName { get; set; }

        public int RoleId { get; set; }

        public long ClientCustomerId { get; set; }

        public string ClientLoginId { get; set; }
    }

    class Campaign
    {
        public string NccCampaignId { get; set; }

        public string CampaignTp { get; set; }

        public long CustomerId { get; set; }

        public string Name { get; set; }

        public UserLock UserLock { get; set; }

        public int DeliveryMethod { get; set; }

        public int UseDailyBudget { get; set; }

        public long DailyBudget { get; set; }

        public int UsePeriod { get; set; }

        public DateTime PeriodStartDt { get; set; }

        public DateTime PeriodEndDt { get; set; }
    }

    class RelKwdStat
    {
        public string siteId { get; set; }
        public int biztpId { get; set; }
        public string hintKeywords { get; set; }
        public int intEvent { get; set; }
        public int month { get; set; }
        public int showDetail { get; set; }
        public string relKeyword { get; set; }
        public string monthlyPcQcCnt { get; set; }
        public string monthlyMobileQcCnt { get; set; }
        public string monthlyAvePcClkCnt { get; set; }
        public string monthlyAveMobileClkCnt { get; set; }
        public string monthlyAvePcCtr { get; set; }
        public string monthlyAveMobileCtr { get; set; }
        public string plAvgDepth { get; set; }
        public string compIdx { get; set; }
    }

    enum UserLock
    {
        ENABLED = 0,
        PAUSED = 1
    }
}
