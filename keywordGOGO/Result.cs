using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace keywordGOGO
{
    class KeywordList
    {
        public string Keyword { get; set; }
        public string Kind { get; set; }
    }

    class KeyWordResult
    {
        public string RelKeyword { get; set; } // 연관 키워드
        public int SellPrdQcCnt { get; set; } // 해당 키워드의 전체 상품수
        public string MonthlyPcQcCnt { get; set; } // 월간 PC 검색수
        public string MonthlyMobileQcCnt { get; set; } // 월간 모바일 검색수
        public string MonthlyAvePcClkCnt { get; set; } // 월간  PC 클릭수
        public string MonthlyAveMobileClkCnt { get; set; } // 월간 모바일 클릭수
        public string MonthlyAvePcCtr { get; set; } // 월간 PC 클릭률 
        public string MonthlyAveMobileCtr { get; set; } // 월간 모바일 클릭률
        public string PlAvgDepth { get; set; } // 경쟁정도
        public string CompIdx { get; set; } // 월간노출 광고수

        private List<ShopAPIResult> shopResult = new List<ShopAPIResult>();
        public List<ShopAPIResult> ShopResult
        {
            set { shopResult = value; }
            get { return shopResult; }
        }
    }

    class ShopAPIResult
    {
        public string RelKeyword { get; set; }
        public string Title { get; set; }
        public string Link { get; set; }
        public string Image { get; set; }
        public string Lprice { get; set; }
        public string Hprice { get; set; }
        public string MallName { get; set; }
        public string ProductId { get; set; }
        public string ProductType { get; set; }
        private List<String> titleKeywordList = new List<String>();
        public List<String> TitleKeywordList
        {
            set { titleKeywordList = value; }
            get { return titleKeywordList; }
        }
    }

    class OpenApiDataSetResult
    {
        public int Total { get; set; }

        private List<String> titleKeywordList = new List<String>();
        public List<String> TitleKeywordList
        {
            set { titleKeywordList = value; }
            get { return titleKeywordList; }
        }

        private List<ShopAPIResult> shopAPIResultList = new List<ShopAPIResult>();
        public List<ShopAPIResult> ShopAPIResultList
        {
            set { shopAPIResultList = value; }
            get { return shopAPIResultList; }
        }
    }

    class GridResultData
    {
        public int RefAdTotalQcCnt { get; set; }

        private List<KeywordList> shoppingRefGrid = new List<KeywordList>();
        public List<KeywordList> ShoppingRefGrid
        {
            set { shoppingRefGrid = value; }
            get { return shoppingRefGrid; }
        }

        private List<KeyWordResult> adRefGrid = new List<KeyWordResult>();
        public List<KeyWordResult> AdRefGrid
        {
            set { adRefGrid = value; }
            get { return adRefGrid; }
        }

        private List<KeyWordResult> tagRefGrid = new List<KeyWordResult>();
        public List<KeyWordResult> TagRefGrid
        {
            set { tagRefGrid = value; }
            get { return tagRefGrid; }
        }

        private List<ShopAPIResult> shopAPIResultList = new List<ShopAPIResult>();
        public List<ShopAPIResult> ShopAPIResultList
        {
            set { shopAPIResultList = value; }
            get { return shopAPIResultList; }
        }

        private List<ProductKeyWordList> titleKeywordList = new List<ProductKeyWordList>();
        public List<ProductKeyWordList> TitleKeywordList
        {
            set { titleKeywordList = value; }
            get { return titleKeywordList; }
        }

        public ShopWebResult ShopWebDataResult { get; set; }
    }


    class RankingList
    {
        public string rank { get; set; }
        public string oldrank { get; set; }
        public string pageNo { get; set; }
        public string oldPageNo { get; set; }
        public string productNo { get; set; }
        public string count { get; set; }
        public string mallName { get; set; }
        public string productName { get; set; }
        public string keyword { get; set; }
        public string productUrl { get; set; }
        public string productPrice { get; set; }
        public string categoryName { get; set; }
        public string adprice { get; set; }
        public string oldadprice { get; set; }

    }

    class GridResultData2
    {
        private List<RankingList> nonadRankingRefGrid = new List<RankingList>();
        public List<RankingList> NonAdRankingRefGrid
        {
            set { nonadRankingRefGrid = value; }
            get { return nonadRankingRefGrid; }
        }

        private List<RankingList> adRankingRefGrid = new List<RankingList>();
        public List<RankingList> AdRankingRefGrid
        {
            set { adRankingRefGrid = value; }
            get { return adRankingRefGrid; }
        }
    }


    class ShopWebResult
    {
        public string RelKeyword { get; set; } // 초기 키워드

        public int TotalCount { get; set; }

        public int AdCount { get; set; }

        private List<string> shoppingRefKeyWord = new List<string>();
        public List<string> ShoppingRefKeyWord
        {
            set { shoppingRefKeyWord = value; }
            get { return shoppingRefKeyWord; }
        }

        private List<KeywordList> outTagList = new List<KeywordList>();
        public List<KeywordList> OutTagList
        {
            set { outTagList = value; }
            get { return outTagList; }
        }

        private List<string> mallList = new List<string>();
        public List<string> MallList
        {
            set { mallList = value; }
            get { return mallList; }
        }

        private List<string> productNmList = new List<string>();
        public List<string> ProductNmList
        {
            set { productNmList = value; }
            get { return productNmList; }
        }
    }
    class ProductKeyWordList
    {
        public string value { get; set; }
        public int count { get; set; }
    }

    class InstagramTagWordList
    {
        public string value { get; set; }
        public int count { get; set; }
    }

}
