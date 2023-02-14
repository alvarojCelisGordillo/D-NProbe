using System;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using DNPrueba.Web.Models;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace DNPrueba.Web.Controllers
{
    public class ReportController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<ContentResult> MakeReport()
        {
            DataResponse Response = new DataResponse();
            HttpClient client = new HttpClient();

            HttpResponseMessage apiresponse = await client.GetAsync(URL.GetReport);

            if (apiresponse.StatusCode == HttpStatusCode.OK)
            {
                var fileName = "Reporte_Anual" + Guid.NewGuid();
                var jsonString = await apiresponse.Content.ReadAsStringAsync();
                Response = JsonConvert.DeserializeObject<DataResponse>(jsonString);

                return Content(Response.Base64);
            }

            return null;
        }
    }

}
