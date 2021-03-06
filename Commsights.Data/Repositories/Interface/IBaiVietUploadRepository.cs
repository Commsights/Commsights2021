﻿using Commsights.Data.DataTransferObject;
using Commsights.Data.Models;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Commsights.Data.Repositories
{
    public interface IBaiVietUploadRepository : IRepository<BaiVietUpload>
    {
        public List<BaiVietUpload> GetByDateBeginAndDateEndAndRequestUserIDAndIsFilterToList(DateTime dateBegin, DateTime dateEnd, int RequestUserID, bool isFilter);
    }
}
