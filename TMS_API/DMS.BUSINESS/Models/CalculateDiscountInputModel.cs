﻿using DMS.CORE.Entities.BU;

namespace DMS.BUSINESS.Models
{
    public class CalculateDiscountInputModel
    {
        public TblBuCalculateDiscount Header { get; set; } = new TblBuCalculateDiscount();
        public List<TblBuInputPrice> InputPrice { get; set; } = new List<TblBuInputPrice>();
        public List<TblBuInputMarket> Market { get; set; } = new List<TblBuInputMarket>();
        public List<TblBuInputCustomerDb> CustomerDb { get; set; } = new List<TblBuInputCustomerDb>();
        public List<TblBuInputCustomerPt> CustomerPt { get; set; } = new List<TblBuInputCustomerPt>();
        public List<TblBuInputCustomerPts> CustomerPts { get; set; } = new List<TblBuInputCustomerPts>();
        public List<TblBuInputCustomerFob> CustomerFob { get; set; } = new List<TblBuInputCustomerFob>();
        public List<TblBuInputCustomerTnpp> CustomerTnpp { get; set; } = new List<TblBuInputCustomerTnpp>();
        public List<TblBuInputCustomerBbdo> CustomerBbdo { get; set; } = new List<TblBuInputCustomerBbdo>();
    }

    public class CustomBBDOExportWord
    {
        public string code { get; set; }
        public string deliveryGroupCode { get; set; }
    }
    public class CustomInput
    {
        public string code { get; set; }
        public string name { get; set; }
    }
}
