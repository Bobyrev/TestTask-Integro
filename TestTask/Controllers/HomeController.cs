using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using ExcelDataReader;
using System.IO;
using System.Data;
using Microsoft.AspNetCore.Hosting;

namespace TestTask.Controllers
{
    public class HomeController : Controller
    {
        IHostingEnvironment _env;
        public HomeController(IHostingEnvironment env) 
        {
            _env = env;
        }

        public IActionResult Index()
        {
            return View();
        }


        [HttpPost]
        public IActionResult Upload(IFormFile uploadFile)
        {
            var excelData = ExcelWork.ReadExcel(uploadFile);
            if (excelData == null) 
            {
                return RedirectToAction("Index");
            }
            ExcelWork.SaveFile(uploadFile, _env.WebRootPath);
            return View(ExcelWork.Filter(excelData));
        }



        [HttpPost]
        public IActionResult Result(List<int> selected) 
        {
            string fileName = ExcelWork.ExcelReport(selected, _env.WebRootPath);
            if (fileName == "") 
            {
                return RedirectToAction("Index");
            }
            return PhysicalFile(fileName, "application/excel", "Report.xls");
        }
    }


}
