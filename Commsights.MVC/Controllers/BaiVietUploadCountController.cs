using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Commsights.Data.Enum;
using Commsights.Data.Helpers;
using Commsights.Data.Models;
using Commsights.Data.Repositories;
using Kendo.Mvc.Extensions;
using Kendo.Mvc.UI;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Commsights.MVC.Controllers
{
    public class BaiVietUploadCountController : BaseController
    {
        private readonly IWebHostEnvironment _hostingEnvironment;
        private readonly IBaiVietUploadCountRepository _baiVietUploadCountRepository;
        public BaiVietUploadCountController(IWebHostEnvironment hostingEnvironment, IBaiVietUploadCountRepository baiVietUploadCountRepository, IMembershipAccessHistoryRepository membershipAccessHistoryRepository) : base(membershipAccessHistoryRepository)
        {
            _hostingEnvironment = hostingEnvironment;
            _baiVietUploadCountRepository = baiVietUploadCountRepository;
        }
        public ActionResult GetReportIndustryByDateBeginAndDateEndToList([DataSourceRequest] DataSourceRequest request, DateTime dateBegin, DateTime dateEnd)
        {
            var data = _baiVietUploadCountRepository.GetReportIndustryByDateBeginAndDateEndToList(dateBegin, dateEnd);
            return Json(data.ToDataSourceResult(request));
        }

        public ActionResult GetReportByDateBeginAndDateEndToList([DataSourceRequest] DataSourceRequest request, DateTime dateBegin, DateTime dateEnd)
        {
            var data = _baiVietUploadCountRepository.GetReportByDateBeginAndDateEndToList(dateBegin, dateEnd);
            return Json(data.ToDataSourceResult(request));
        }
        public ActionResult GetReportByDateBeginAndDateEndAndIndustryIDToList([DataSourceRequest] DataSourceRequest request, DateTime dateBegin, DateTime dateEnd, int industryID)
        {
            var data = _baiVietUploadCountRepository.GetReportByDateBeginAndDateEndAndIndustryIDToList(dateBegin, dateEnd, industryID);
            return Json(data.ToDataSourceResult(request));
        }
        public ActionResult GetReportByDateBeginAndDateEndAndIndustryIDAndEmployeeIDToList([DataSourceRequest] DataSourceRequest request, string dateBeginString, string dateEndString, int industryID, int employeeID)
        {
            DateTime dateBegin = DateTime.Parse(dateBeginString);
            DateTime dateEnd = DateTime.Parse(dateEndString);
            var data = _baiVietUploadCountRepository.GetReportByDateBeginAndDateEndAndIndustryIDAndEmployeeIDToList(dateBegin, dateEnd, industryID, employeeID);
            return Json(data.ToDataSourceResult(request));
        }
        public string ExportURLExcel(CancellationToken cancellationToken, DateTime dateBegin, DateTime dateEnd)
        {
            List<BaiVietUpload> list = new List<BaiVietUpload>();
            string excelName = @"URL_" + AppGlobal.DateTimeCode + ".xlsx";
            string sheetName = AppGlobal.DateTimeCode;
            list = _baiVietUploadCountRepository.GetByDateBeginAndDateEndAndRequestUserIDAndIsFilterToList(dateBegin, dateEnd, RequestUserID, false);
            var streamExport = new MemoryStream();
            Color color = Color.FromArgb(int.Parse("#c00000".Replace("#", ""), System.Globalization.NumberStyles.AllowHexSpecifier));
            Color colorTitle = Color.FromArgb(int.Parse("#ed7d31".Replace("#", ""), System.Globalization.NumberStyles.AllowHexSpecifier));

            using (var package = new ExcelPackage(streamExport))
            {
                var workSheet = package.Workbook.Worksheets.Add(sheetName);

                int column = 1;
                int rowExcel = 1;
                workSheet.Cells[rowExcel, column].Value = "URL";
                workSheet.Cells[rowExcel, column].Style.Font.Bold = true;
                workSheet.Cells[rowExcel, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[rowExcel, column].Style.Font.Color.SetColor(System.Drawing.Color.White);
                workSheet.Cells[rowExcel, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells[rowExcel, column].Style.Fill.BackgroundColor.SetColor(color);
                workSheet.Cells[rowExcel, column].Style.Font.Name = "Times New Roman";
                workSheet.Cells[rowExcel, column].Style.Font.Size = 11;
                workSheet.Cells[rowExcel, column].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowExcel, column].Style.Border.Top.Color.SetColor(Color.Black);
                workSheet.Cells[rowExcel, column].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowExcel, column].Style.Border.Left.Color.SetColor(Color.Black);
                workSheet.Cells[rowExcel, column].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowExcel, column].Style.Border.Right.Color.SetColor(Color.Black);
                workSheet.Cells[rowExcel, column].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowExcel, column].Style.Border.Bottom.Color.SetColor(Color.Black);
                column = column + 1;
                workSheet.Cells[rowExcel, column].Value = "Headline  (Vie)";
                workSheet.Cells[rowExcel, column].Style.Font.Bold = true;
                workSheet.Cells[rowExcel, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[rowExcel, column].Style.Font.Color.SetColor(System.Drawing.Color.White);
                workSheet.Cells[rowExcel, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells[rowExcel, column].Style.Fill.BackgroundColor.SetColor(color);
                workSheet.Cells[rowExcel, column].Style.Font.Name = "Times New Roman";
                workSheet.Cells[rowExcel, column].Style.Font.Size = 11;
                workSheet.Cells[rowExcel, column].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowExcel, column].Style.Border.Top.Color.SetColor(Color.Black);
                workSheet.Cells[rowExcel, column].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowExcel, column].Style.Border.Left.Color.SetColor(Color.Black);
                workSheet.Cells[rowExcel, column].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowExcel, column].Style.Border.Right.Color.SetColor(Color.Black);
                workSheet.Cells[rowExcel, column].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowExcel, column].Style.Border.Bottom.Color.SetColor(Color.Black);
                column = column + 1;
                workSheet.Cells[rowExcel, column].Value = "Content";
                workSheet.Cells[rowExcel, column].Style.Font.Bold = true;
                workSheet.Cells[rowExcel, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[rowExcel, column].Style.Font.Color.SetColor(System.Drawing.Color.White);
                workSheet.Cells[rowExcel, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells[rowExcel, column].Style.Fill.BackgroundColor.SetColor(color);
                workSheet.Cells[rowExcel, column].Style.Font.Name = "Times New Roman";
                workSheet.Cells[rowExcel, column].Style.Font.Size = 11;
                workSheet.Cells[rowExcel, column].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowExcel, column].Style.Border.Top.Color.SetColor(Color.Black);
                workSheet.Cells[rowExcel, column].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowExcel, column].Style.Border.Left.Color.SetColor(Color.Black);
                workSheet.Cells[rowExcel, column].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowExcel, column].Style.Border.Right.Color.SetColor(Color.Black);
                workSheet.Cells[rowExcel, column].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowExcel, column].Style.Border.Bottom.Color.SetColor(Color.Black);
                column = column + 1;
                workSheet.Cells[rowExcel, column].Value = "Date";
                workSheet.Cells[rowExcel, column].Style.Font.Bold = true;
                workSheet.Cells[rowExcel, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[rowExcel, column].Style.Font.Color.SetColor(System.Drawing.Color.White);
                workSheet.Cells[rowExcel, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells[rowExcel, column].Style.Fill.BackgroundColor.SetColor(color);
                workSheet.Cells[rowExcel, column].Style.Font.Name = "Times New Roman";
                workSheet.Cells[rowExcel, column].Style.Font.Size = 11;
                workSheet.Cells[rowExcel, column].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowExcel, column].Style.Border.Top.Color.SetColor(Color.Black);
                workSheet.Cells[rowExcel, column].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowExcel, column].Style.Border.Left.Color.SetColor(Color.Black);
                workSheet.Cells[rowExcel, column].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowExcel, column].Style.Border.Right.Color.SetColor(Color.Black);
                workSheet.Cells[rowExcel, column].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowExcel, column].Style.Border.Bottom.Color.SetColor(Color.Black);
                int index = 0;
                rowExcel = rowExcel + 1;
                for (int row = rowExcel; row <= list.Count + rowExcel - 1; row++)
                {

                    workSheet.Cells[row, 1].Value = list[index].URLCode;
                    if (!string.IsNullOrEmpty(list[index].URLCode))
                    {
                        try
                        {
                            workSheet.Cells[row, 1].Hyperlink = new Uri(list[index].URLCode);
                        }
                        catch
                        {
                        }
                        workSheet.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                    }
                    workSheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    workSheet.Cells[row, 1].Style.Font.Name = "Times New Roman";
                    workSheet.Cells[row, 1].Style.Font.Size = 11;
                    workSheet.Cells[row, 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[row, 1].Style.Border.Top.Color.SetColor(Color.Black);
                    workSheet.Cells[row, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[row, 1].Style.Border.Left.Color.SetColor(Color.Black);
                    workSheet.Cells[row, 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[row, 1].Style.Border.Right.Color.SetColor(Color.Black);
                    workSheet.Cells[row, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    workSheet.Cells[row, 1].Style.Border.Bottom.Color.SetColor(Color.Black);
                }
                for (int i = 1; i <= 4; i++)
                {
                    workSheet.Column(i).AutoFit();
                }
                package.Save();
            }
            streamExport.Position = 0;
            var physicalPathCreate = Path.Combine(_hostingEnvironment.WebRootPath, AppGlobal.FTPDownloadReprotMonth, excelName);
            using (var stream = new FileStream(physicalPathCreate, FileMode.Create))
            {
                streamExport.CopyTo(stream);
            }
            string result = AppGlobal.DomainSub + AppGlobal.URLDownloadReprotMonth + excelName;
            return result;
        }
    }
}
