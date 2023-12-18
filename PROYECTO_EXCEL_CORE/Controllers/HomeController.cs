using Microsoft.AspNetCore.Mvc;
using PROYECTO_EXCEL_CORE.Models;
using System.Diagnostics;

using System.Data;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Business;

namespace PROYECTO_EXCEL_CORE.Controllers
{
    public class HomeController : Controller
    {
        private readonly Report _reportService;

        public HomeController(Report reportService)
        {
            _reportService = reportService;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Exportar_Excel()
        {
            try
            {
                byte[] excelData = _reportService.ExportarExcel();

                var nombreExcel = $"Reporte_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";

                return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombreExcel);
            }
            catch (Exception ex)
            {
                // Manejo de excepciones...
                return RedirectToAction("Error", "Home");
            }
        }

        public IActionResult ExportarWord()
        {
            // OJO NO SE TIENE EL ESTILO DE AUTOAJUSTADO DE LA TABLA DEL ESPORTADO EN WORD!!!
            try
            {               
                byte[] wordData = _reportService.ExportarWord();

                var nombreWord = $"Reporte_{DateTime.Now.ToString("yyyyMMddHHmmss")}.docx";

                return File(wordData, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", nombreWord);
            }
            catch (Exception ex)
            {
                // Manejo de excepciones...
                return RedirectToAction("Error", "Home");
            }
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}