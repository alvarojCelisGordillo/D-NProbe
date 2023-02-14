using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DN.DAL;
using DNPrueba.API.Core;
using DNPrueba.API.Core.IRepositories;
using DNPrueba.API.Persistance.Repositories;

namespace DNPrueba.API.Persistance
{
    public class UnitOfWork : IUnitOfWork
    {
        public readonly DAL _DAL;
        public IReportRepository Reports { get; private set; }

        public UnitOfWork()
        {
            _DAL = new DAL();
            Reports = new ReportRepository(_DAL);
        }
    }
}
