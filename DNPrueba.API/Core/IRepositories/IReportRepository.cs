using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DNPrueba.API.Models;

namespace DNPrueba.API.Core.IRepositories
{
    public interface IReportRepository
    {
        Task<DataResponse> GetComparisonReport();
    }
}
