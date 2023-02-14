using DNPrueba.API.Core.IRepositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DNPrueba.API.Core
{
    public interface IUnitOfWork
    {
        IReportRepository Reports { get; }
    }
}
