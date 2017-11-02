using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Microsoft.AspNetCore.Hosting;

namespace GenerateExcel.Controllers
{
    [Route("api/[controller]")]
    public class DownloadsController : Controller
    {
        private readonly IHostingEnvironment _hostingEnvironment;

        public DownloadsController(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }

        // GET api/values
        [HttpGet]
        public IActionResult Get()
        {
            var contentRootPath = _hostingEnvironment.ContentRootPath;
            var template = Path.Combine(contentRootPath, @"Templates\Template.xlsx");
            var ms = new MemoryStream();

            using (var fs = new FileStream(template, FileMode.Open, FileAccess.Read))
            {
                fs.Position = 0;
                IWorkbook workbook = new XSSFWorkbook(fs);

                ISheet sheet1 = workbook.GetSheet("Sheet1");
                sheet1.GetRow(1).CreateCell(0).SetCellValue($"This is a dynamic cell, created at {DateTime.Now.ToShortTimeString()}.");
                sheet1.GetRow(1).GetCell(1).SetCellValue("I'm still green!");
                sheet1.AutoSizeColumn(0);

                workbook.Write(ms);
            }

            ms.Position = 0;
            return File(ms, System.Net.Mime.MediaTypeNames.Application.Octet, "file.xlsx");
        }
    }
}
