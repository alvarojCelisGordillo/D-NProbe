using DNPrueba.API.Core;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DNPrueba.API.Persistance;

namespace DNPrueba.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReportController : ControllerBase
    {
        public readonly IUnitOfWork _iUnitOfWork;


        public ReportController()
        {
            _iUnitOfWork = new UnitOfWork();
        }

        [HttpGet]
        public async Task<IActionResult> GetReport()
        {
            var report = await _iUnitOfWork.Reports.GetComparisonReport();

            if (report == null)
            {
                return NotFound();
            }

            return Ok(report);
        }
    }
}
