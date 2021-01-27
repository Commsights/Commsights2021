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
    public class BaiVietUploadRepository : Repository<BaiVietUpload>, IBaiVietUploadRepository
    {
        private readonly CommsightsContext _context;

        public BaiVietUploadRepository(CommsightsContext context) : base(context)
        {
            _context = context;
        }
        public List<BaiVietUpload> GetByDateBeginAndDateEndAndRequestUserIDAndIsFilterToList(DateTime dateBegin, DateTime dateEnd,int RequestUserID, bool isFilter)
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
                    DataTable dt = SQLHelper.Fill(AppGlobal.ConectionString, "sp_CodeDataSelectByDateUpdatedBeginAndDateUpdatedEndAndEmployeeIDAndIsFilter", parameters);
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
