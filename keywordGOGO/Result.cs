﻿using System;
using System.Collections.Generic;

namespace keywordGOGO
{
    class KeywordList
    {
        public string Keyword { get; set; }
        public string Kind { get; set; }
    }

    class SaleAmountResult
    {
        public string productName { get; set; } // 상품명
        public string adYn { get; set; } // 광고여부
        public string categoryName { get; set; } // 카테고리
        public string openDate { get; set; } // 오픈일
        public string mallName { get; set; } // 몰이름
        public string totalReviewCount { get; set; } // 리뷰수
        public string averageReviewScore { get; set; } // 평점

        public string leadTimeCount1 { get; set; } // 평점
        public string leadTimeCount2 { get; set; } // 평점
        public string leadTimeCount3 { get; set; } // 평점
        public string leadTimeCount4 { get; set; } // 평점
        public string totalleadTimeCount1 { get; set; } // 평점

        //public string cumulationSaleCount { get; set; } // 6개월
        //public string recentSaleCount { get; set; } // 최근 3일
        public string urlLink { get; set; } //url
    }

    class ExcellOutResult
    {
        public string RelKeyword { get; set; } // 연관 키워드
        public string PlAvgDepth { get; set; } // 월간노출 광고수
        public string Kinds { get; set; } // 키워드 종류
        public int Count { get; set; } // 중복갯수
    }
    
    class RelKeyWordResult
    {
        public string RelKeyword { get; set; } // 연관 키워드
        public string MonthlyPcQcCnt { get; set; } // 월간 PC 검색수
        public string MonthlyMobileQcCnt { get; set; } // 월간 모바일 검색수
        public string MonthlyAvePcClkCnt { get; set; } // 월간  PC 클릭수
        public string MonthlyAveMobileClkCnt { get; set; } // 월간 모바일 클릭수
        public string MonthlyAvePcCtr { get; set; } // 월간 PC 클릭률 
        public string MonthlyAveMobileCtr { get; set; } // 월간 모바일 클릭률
        public string PlAvgDepth { get; set; } // 경쟁정도
        public string CompIdx { get; set; } // 월간노출 광고수

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
        public string category { get; set; } // 대표카테고리

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
        public string category1 { get; set; }
        public string category2 { get; set; }
        public string category3 { get; set; }
        public string category4 { get; set; }

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
        public string category { get; set; }

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
        public string similarity { get; set; } //유사성
        public string relevance { get; set; } // 관련성
        public string hitRank { get; set; }
        public string reviewCountSum { get; set; }//리뷰총합
        public string purchaseCnt { get; set; }//구매건수

        public string daysSaleSum7 { get; set; }//daysSaleSum7
        public string adYn { get; set; }


    }

    class CategoryList
    {
        public string productName { get; set; }
        public string CategoryName_1st { get; set; }
        public string CategoryCode_1st { get; set; }
        public string CategoryCnt_1st { get; set; }
        public string CategoryName_2nd { get; set; }
        public string CategoryCode_2nd { get; set; }
        public string CategoryCnt_2nd { get; set; }
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

        private List<RankingList> allRankingRefGrid = new List<RankingList>();
        public List<RankingList> AllRankingRefGrid
        {
            set { allRankingRefGrid = value; }
            get { return allRankingRefGrid; }
        }
    }


    class ShopWebResult
    {
        public string RelKeyword { get; set; } // 초기 키워드

        public int TotalCount { get; set; }

        public int AdCount { get; set; }

        private List<KeywordList> shoppingRefKeyWord = new List<KeywordList>();
        public List<KeywordList> ShoppingRefKeyWord
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

    class CategoryListData
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
