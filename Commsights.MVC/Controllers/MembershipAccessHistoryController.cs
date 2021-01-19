using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Commsights.Data.DataTransferObject;
using Commsights.Data.Enum;
using Commsights.Data.Helpers;
using Commsights.Data.Models;
using Commsights.Data.Repositories;
using Commsights.MVC.Models;
using Kendo.Mvc.Extensions;
using Kendo.Mvc.UI;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Mvc.ViewEngines;
using Microsoft.AspNetCore.Mvc.ViewFeatures;

namespace Commsights.MVC.Controllers
{
    public class MembershipAccessHistoryController : BaseController
    {
        private readonly IMembershipAccessHistoryRepository _membershipAccessHistoryRepository;

        public MembershipAccessHistoryController(IMembershipAccessHistoryRepository membershipAccessHistoryRepository) : base(membershipAccessHistoryRepository)
        {
            _membershipAccessHistoryRepository = membershipAccessHistoryRepository;
        }
        private void Initialization()
        {

        }
        public IActionResult Index()
        {
            CodeDataViewModel model = new CodeDataViewModel();
            model.DatePublishBegin = DateTime.Now;
            model.DatePublishEnd = DateTime.Now;            
            return View(model);
        }

        public ActionResult GetByDateBeginAndDateEndAndMembershipIDToList([DataSourceRequest] DataSourceRequest request, DateTime dateBegin, DateTime dateEnd, int membershipID)
        {
            var data = _membershipAccessHistoryRepository.GetByDateBeginAndDateEndAndMembershipIDToList(dateBegin, dateEnd, membershipID);
            return Json(data.ToDataSourceResult(request));
        }
    }
}
