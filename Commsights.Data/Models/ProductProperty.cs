﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Commsights.Data.Models
{
    public partial class ProductProperty : BaseModel
    {
        public int? Source { get; set; }
        public string Day { get; set; }
        public string Week { get; set; }
        public string Month { get; set; }
        public string Quarter { get; set; }
        public string Year { get; set; }
        public int? ArticleTypeID { get; set; }
        public int? CategoryMainID { get; set; }
        public string CategoryMain { get; set; }
        public int? CategorySubID { get; set; }
        public string CategorySub { get; set; }
        public int? CompanyID { get; set; }
        public string CompanyName { get; set; }
        public int? CorpCopyID { get; set; }
        public string CorpCopy { get; set; }
        public decimal? SOECompany { get; set; }
        public int? FeatureCorpID { get; set; }
        public string FeatureCorp { get; set; }
        public int? AssessID { get; set; }
        public int? SegmentID { get; set; }
        public string Segment { get; set; }
        public int? ProductID { get; set; }
        public string ProductName_ProjectName { get; set; }
        public int? SegmentProductID { get; set; }
        public string SegmentProduct { get; set; }
        public decimal? SOEProduct { get; set; }
        public int? FeatureProductID { get; set; }
        public string FeatureProduct { get; set; }
        public int? IndustryID { get; set; }
        public string Industry { get; set; }
        public string GUICode { get; set; }
        public string Code { get; set; }
        public decimal? Point { get; set; }
        public int? KeywordCount { get; set; }
        public bool? IsSummary { get; set; }
        public bool? IsData { get; set; }
        public bool? IsDaily { get; set; }
        public bool? IsWeekly { get; set; }
        public bool? IsMonthly { get; set; }
        public bool? IsQuarterly { get; set; }
        public bool? IsYearly { get; set; }
        public decimal? ROME_Corp_VND { get; set; }
        public decimal? ROME_Corp_USD { get; set; }
        public decimal? ROME_Product_VND { get; set; }
        public decimal? ROME_Product_USD { get; set; }
        public int? TierCommsightsID { get; set; }
        public string TierCommsights { get; set; }
        public int? TierCustomerID { get; set; }
        public string TierCustomer { get; set; }
        public string SpokePersonName { get; set; }
        public string SpokePersonTitle { get; set; }
        public decimal? SpokePersonValue { get; set; }
        public decimal? ToneValue { get; set; }
        public decimal? HeadlineValue { get; set; }
        public decimal? FeatureValue { get; set; }
        public decimal? TierValue { get; set; }
        public decimal? PictureValue { get; set; }
        public decimal? MPS { get; set; }
        public int? SentimentCorpID { get; set; }
        public string SentimentCorp { get; set; }
        public int? SentimentProductID { get; set; }
        public string SentimentProduct { get; set; }
        public int? ReportMonthlyID { get; set; }
        public string Headline { get; set; }
        public string URL { get; set; }
        public string Page { get; set; }
        public string Journalist { get; set; }
        public string HeadlineEngLish { get; set; }
        public string MediaTitle { get; set; }
        public string MediaType { get; set; }
        public string Duration { get; set; }
        public decimal? Advalue { get; set; }
        public DateTime? DatePublish { get; set; }
        public string FileName { get; set; }
        public decimal? SentimentValue { get; set; }
        public string CampaignName { get; set; }
        public string KeyMessage { get; set; }
        public decimal? CompetitiveNewsValue { get; set; }
        public string CompetitiveWhat { get; set; }
        public decimal? KOLValue { get; set; }
        public decimal? OtherValue { get; set; }
        public decimal? TasteValue { get; set; }
        public decimal? PriceValue { get; set; }
        public decimal? NutritionFactValue { get; set; }
        public decimal? VitaminValue { get; set; }
        public decimal? GoodForHealthValue { get; set; }
        public decimal? Bottle_CanDesignValue { get; set; }
        public bool? IsCoding { get; set; }
        public bool? IsAnalysis { get; set; }
        public bool? IsDownload { get; set; }
        public bool? IsSend { get; set; }
        public bool? IsCopy { get; set; }
        public string CampaignKeyMessage { get; set; }
        public string TitleProperty { get; set; }
        public int? CopyVersion { get; set; }
        public int? SourceProperty { get; set; }
        public DateTime? DateCoding { get; set; }
        public bool? IsDailyDownload { get; set; }
        public DateTime? DateDailyDownload { get; set; }
    }
}