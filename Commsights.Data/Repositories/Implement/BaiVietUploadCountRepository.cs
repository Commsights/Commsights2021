using Commsights.Data.DataTransferObject;
using Commsights.Data.Helpers;
using Commsights.Data.Models;
using Microsoft.EntityFrameworkCore.Internal;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Commsights.Data.Repositories
{
    public class BaiVietUploadCountRepository : Repository<BaiVietUploadCount>, IBaiVietUploadCountRepository
    {
        private readonly CommsightsContext _context;

        public BaiVietUploadCountRepository(CommsightsContext context) : base(context)
        {
            _context = context;
        }
        public List<BaiVietReport> GetReportByDateBeginAndDateEndAndIndustryIDAndEmployeeIDToList(DateTime dateBegin, DateTime dateEnd, int industryID, int employeeID)
        {
            List<BaiVietReport> list = new List<BaiVietReport>();
            if (industryID > 0)
            {
                try
                {
                    dateBegin = new DateTime(dateBegin.Year, dateBegin.Month, dateBegin.Day, 0, 0, 0);
                    dateEnd = new DateTime(dateEnd.Year, dateEnd.Month, dateEnd.Day, 23, 59, 59);
                    SqlParameter[] parameters =
                    {
                    new SqlParameter("@DateBegin",dateBegin),
                    new SqlParameter("@DateEnd",dateEnd),
                    new SqlParameter("@IndustryID",industryID),
                    new SqlParameter("@EmployeeID",employeeID),
                    };
                    DataTable dt = SQLHelper.Fill(AppGlobal.ConectionString, "sp_BaiVietUploadReportByDateBeginAndDateEndAndIndustryIDAndEmployeeID", parameters);
                    list = SQLHelper.ToList<BaiVietReport>(dt);
                }
                catch (Exception e)
                {
                    string mes = e.Message;
                }
            }
            return list;
        }
        public List<BaiVietReport> GetReportIndustryByDateBeginAndDateEndToList(DateTime dateBegin, DateTime dateEnd)
        {
            List<BaiVietReport> list = new List<BaiVietReport>();
            try
            {
                dateBegin = new DateTime(dateBegin.Year, dateBegin.Month, dateBegin.Day, 0, 0, 0);
                dateEnd = new DateTime(dateEnd.Year, dateEnd.Month, dateEnd.Day, 23, 59, 59);
                SqlParameter[] parameters =
                {
                    new SqlParameter("@DateBegin",dateBegin),
                    new SqlParameter("@DateEnd",dateEnd),
                    };
                DataTable dt = SQLHelper.Fill(AppGlobal.ConectionString, "sp_BaiVietUploadReportIndustryByDateBeginAndDateEnd", parameters);
                list = SQLHelper.ToList<BaiVietReport>(dt);
            }
            catch (Exception e)
            {
                string mes = e.Message;
            }
            return list;
        }
        public List<BaiVietReport> GetReportByDateBeginAndDateEndToList(DateTime dateBegin, DateTime dateEnd)
        {
            List<BaiVietReport> list = new List<BaiVietReport>();
            try
            {
                dateBegin = new DateTime(dateBegin.Year, dateBegin.Month, dateBegin.Day, 0, 0, 0);
                dateEnd = new DateTime(dateEnd.Year, dateEnd.Month, dateEnd.Day, 23, 59, 59);
                SqlParameter[] parameters =
                {
                    new SqlParameter("@DateBegin",dateBegin),
                    new SqlParameter("@DateEnd",dateEnd),
                    };
                DataTable dt = SQLHelper.Fill(AppGlobal.ConectionString, "sp_BaiVietUploadReportByDateBeginAndDateEnd", parameters);
                list = SQLHelper.ToList<BaiVietReport>(dt);
            }
            catch (Exception e)
            {
                string mes = e.Message;
            }
            return list;
        }
        public List<BaiVietReport> GetReportByDateBeginAndDateEndAndIndustryIDToList(DateTime dateBegin, DateTime dateEnd, int industryID)
        {
            List<BaiVietReport> list = new List<BaiVietReport>();
            if (industryID > 0)
            {
                try
                {
                    dateBegin = new DateTime(dateBegin.Year, dateBegin.Month, dateBegin.Day, 0, 0, 0);
                    dateEnd = new DateTime(dateEnd.Year, dateEnd.Month, dateEnd.Day, 23, 59, 59);
                    SqlParameter[] parameters =
                    {
                    new SqlParameter("@DateBegin",dateBegin),
                    new SqlParameter("@DateEnd",dateEnd),
                    new SqlParameter("@IndustryID",industryID),
                    };
                    DataTable dt = SQLHelper.Fill(AppGlobal.ConectionString, "sp_BaiVietUploadReportByDateBeginAndDateEndAndIndustryID", parameters);
                    list = SQLHelper.ToList<BaiVietReport>(dt);
                }
                catch (Exception e)
                {
                    string mes = e.Message;
                }
            }
            return list;
        }

        public List<BaiVietUpload> GetByDateBeginAndDateEndAndRequestUserIDAndIsFilterToList(DateTime dateBegin, DateTime dateEnd, int RequestUserID, bool isFilter)
        {
            List<BaiVietUpload> list = new List<BaiVietUpload>();
            if (RequestUserID > 0)
            {
                try
                {
                    dateBegin = new DateTime(dateBegin.Year, dateBegin.Month, dateBegin.Day, 0, 0, 0);
                    dateEnd = new DateTime(dateEnd.Year, dateEnd.Month, dateEnd.Day, 23, 59, 59);
                    SqlParameter[] parameters =
                    {
                        new SqlParameter("@DateBegin",dateBegin),
                        new SqlParameter("@DateEnd",dateEnd),
                        new SqlParameter("@RequestUserID",RequestUserID),
                        new SqlParameter("@IsFilter",isFilter),
                    };
                    DataTable dt = SQLHelper.Fill(AppGlobal.ConectionString, "sp_BaiVietUploadSelectByDateBeginAndDateEndAndRequestUserIDAndIsFilter", parameters);
                    list = SQLHelper.ToList<BaiVietUpload>(dt);
                }
                catch (Exception e)
                {
                    string mes = e.Message;
                }
            }
            return list;
        }
    }
}
