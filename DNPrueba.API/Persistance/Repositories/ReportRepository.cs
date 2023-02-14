using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DN.DAL;
using DNPrueba.API.Core.IRepositories;
using DNPrueba.API.Models;
using Microsoft.AspNetCore.Mvc.DataAnnotations;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace DNPrueba.API.Persistance.Repositories
{
    public class ReportRepository : IReportRepository
    {
        private readonly DAL _DAL;

        public ReportRepository(DAL DAL)
        {
            _DAL = DAL;
        }


        public async Task<DataResponse> GetComparisonReport()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            DataTable MonthBudgetDT = _DAL.GetTable(SPN.GetMonthBudget);
            DataTable OrderedExpenses = _DAL.GetTable(SPN.GetOrderedExpensesList);
            DataTable PABalance = _DAL.GetTable(SPN.GetPABalanceByMonth);
            DataTable COBalance = _DAL.GetTable(SPN.GetCOBalanceByMonth);


            var fileReport = new FileInfo(@"C:\Users\Nelly\Desktop\D&NReports\report.xlsx");

            var file = new byte[]{};

            if (fileReport.Exists)
            {
                fileReport.Delete();
            }

            using (var package = new ExcelPackage(fileReport))
            {
                var ws = package.Workbook.Worksheets.Add("Validacion Cierre");

                //var start = ws.Dimension.Start;
                //var end = ws.Dimension.End;

                for (int i = 0; i < OrderedExpenses.Rows.Count; i++)
                {
                    var rubro = OrderedExpenses.Rows[i]["RubroDeGastos"].ToString();
                    if (MonthBudgetDT.AsEnumerable().Any(x => x.Field<string>("RubroHomologado") == rubro))
                    {
                        ws.Cells[i + 3, 1].Value = OrderedExpenses.Rows[i]["RubroDeGastos"].ToString();
                        ws.Cells[i + 3, 2].Value = MonthBudgetDT.AsEnumerable()
                            .Where(x => x.Field<string>("RubroHomologado") == rubro).Select(x => x.Field<decimal>("Ene"));
                        ws.Cells[i + 3, 5].Value = MonthBudgetDT.AsEnumerable()
                            .Where(x => x.Field<string>("RubroHomologado") == rubro).Select(x => x.Field<decimal>("Feb"));
                        ws.Cells[i + 3, 10].Value = MonthBudgetDT.AsEnumerable()
                            .Where(x => x.Field<string>("RubroHomologado") == rubro).Select(x => x.Field<decimal>("Mar"));
                        ws.Cells[i + 3, 15].Value = MonthBudgetDT.AsEnumerable()
                            .Where(x => x.Field<string>("RubroHomologado") == rubro).Select(x => x.Field<decimal>("Abr"));
                        ws.Cells[i + 3, 20].Value = MonthBudgetDT.AsEnumerable()
                            .Where(x => x.Field<string>("RubroHomologado") == rubro).Select(x => x.Field<decimal>("May"));
                        ws.Cells[i + 3, 25].Value = MonthBudgetDT.AsEnumerable()
                            .Where(x => x.Field<string>("RubroHomologado") == rubro).Select(x => x.Field<decimal>("Jun"));
                        ws.Cells[i + 3, 30].Value = MonthBudgetDT.AsEnumerable()
                            .Where(x => x.Field<string>("RubroHomologado") == rubro).Select(x => x.Field<decimal>("Jul"));
                        ws.Cells[i + 3, 35].Value = MonthBudgetDT.AsEnumerable()
                            .Where(x => x.Field<string>("RubroHomologado") == rubro).Select(x => x.Field<decimal>("Ago"));
                        ws.Cells[i + 3, 40].Value = MonthBudgetDT.AsEnumerable()
                            .Where(x => x.Field<string>("RubroHomologado") == rubro).Select(x => x.Field<decimal>("Sep"));
                        ws.Cells[i + 3, 45].Value = MonthBudgetDT.AsEnumerable()
                            .Where(x => x.Field<string>("RubroHomologado") == rubro).Select(x => x.Field<decimal>("Oct"));
                        ws.Cells[i + 3, 50].Value = MonthBudgetDT.AsEnumerable()
                            .Where(x => x.Field<string>("RubroHomologado") == rubro).Select(x => x.Field<decimal>("Nov"));
                        ws.Cells[i + 3, 55].Value = MonthBudgetDT.AsEnumerable()
                            .Where(x => x.Field<string>("RubroHomologado") == rubro).Select(x => x.Field<decimal>("Dic"));

                        ws.Cells[i + 3, 3].Value = PABalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 1)
                            .Sum(x => x.Field<decimal>("Balance")) + COBalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 1)
                            .Sum(x => x.Field<decimal>("Balance"));
                        ws.Cells[i + 3, 6].Value = PABalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 2)
                            .Sum(x => x.Field<decimal>("Balance")) + COBalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 2)
                            .Sum(x => x.Field<decimal>("Balance"));
                        ws.Cells[i + 3, 11].Value = PABalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 3)
                            .Sum(x => x.Field<decimal>("Balance")) + COBalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 3)
                            .Sum(x => x.Field<decimal>("Balance"));
                        ws.Cells[i + 3, 16].Value = PABalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 4)
                            .Sum(x => x.Field<decimal>("Balance")) + COBalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 4)
                            .Sum(x => x.Field<decimal>("Balance"));
                        ws.Cells[i + 3, 21].Value = PABalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 5)
                            .Sum(x => x.Field<decimal>("Balance")) + COBalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 5)
                            .Sum(x => x.Field<decimal>("Balance"));
                        ws.Cells[i + 3, 26].Value = PABalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 6)
                            .Sum(x => x.Field<decimal>("Balance")) + COBalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 6)
                            .Sum(x => x.Field<decimal>("Balance"));
                        ws.Cells[i + 3, 31].Value = PABalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 7)
                            .Sum(x => x.Field<decimal>("Balance")) + COBalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 7)
                            .Sum(x => x.Field<decimal>("Balance"));
                        ws.Cells[i + 3, 36].Value = PABalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 8)
                            .Sum(x => x.Field<decimal>("Balance")) + COBalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 8)
                            .Sum(x => x.Field<decimal>("Balance"));
                        ws.Cells[i + 3, 41].Value = PABalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 9)
                            .Sum(x => x.Field<decimal>("Balance")) + COBalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 9)
                            .Sum(x => x.Field<decimal>("Balance"));
                        ws.Cells[i + 3, 46].Value = PABalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 10)
                            .Sum(x => x.Field<decimal>("Balance")) + COBalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 10)
                            .Sum(x => x.Field<decimal>("Balance"));
                        ws.Cells[i + 3, 51].Value = PABalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 11)
                            .Sum(x => x.Field<decimal>("Balance")) + COBalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 11)
                            .Sum(x => x.Field<decimal>("Balance"));
                        ws.Cells[i + 3, 56].Value = PABalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 12)
                            .Sum(x => x.Field<decimal>("Balance")) + COBalance.AsEnumerable()
                            .Where(x => x.Field<string>("RubroDeGastos") == rubro && x.Field<int>("Mes") == 12)
                            .Sum(x => x.Field<decimal>("Balance"));

                        ws.Cells[i + 3, 4].Formula = $"=IF(${ws.Cells[i + 3, 2]} < 0, SUM(${ws.Cells[i + 3, 2]}:${ws.Cells[i + 3, 3]}),${ws.Cells[i + 3, 2]}-ABS(${ws.Cells[i + 3, 3]}))";
                        ws.Cells[i + 3, 7].Formula = $"=IF(${ws.Cells[i + 3, 5]} < 0, SUM(${ws.Cells[i + 3, 5]}:${ws.Cells[i + 3, 6]}),${ws.Cells[i + 3, 5]}-ABS(${ws.Cells[i + 3, 6]}))";
                        ws.Cells[i + 3, 12].Formula = $"=IF(${ws.Cells[i + 3, 10]} < 0, SUM(${ws.Cells[i + 3, 10]}:${ws.Cells[i + 3, 11]}),${ws.Cells[i + 3, 10]}-ABS(${ws.Cells[i + 3, 11]}))";
                        ws.Cells[i + 3, 17].Formula = $"=IF(${ws.Cells[i + 3, 15]} < 0, SUM(${ws.Cells[i + 3, 15]}:${ws.Cells[i + 3, 16]}),${ws.Cells[i + 3, 15]}-ABS(${ws.Cells[i + 3, 16]}))";
                        ws.Cells[i + 3, 22].Formula = $"=IF(${ws.Cells[i + 3, 20]} < 0, SUM(${ws.Cells[i + 3, 20]}:${ws.Cells[i + 3, 21]}),${ws.Cells[i + 3, 20]}-ABS(${ws.Cells[i + 3, 21]}))";
                        ws.Cells[i + 3, 27].Formula = $"=IF(${ws.Cells[i + 3, 25]} < 0, SUM(${ws.Cells[i + 3, 25]}:${ws.Cells[i + 3, 26]}),${ws.Cells[i + 3, 25]}-ABS(${ws.Cells[i + 3, 26]}))";
                        ws.Cells[i + 3, 32].Formula = $"=IF(${ws.Cells[i + 3, 30]} < 0, SUM(${ws.Cells[i + 3, 30]}:${ws.Cells[i + 3, 31]}),${ws.Cells[i + 3, 30]}-ABS(${ws.Cells[i + 3, 31]}))";
                        ws.Cells[i + 3, 37].Formula = $"=IF(${ws.Cells[i + 3, 35]} < 0, SUM(${ws.Cells[i + 3, 35]}:${ws.Cells[i + 3, 36]}),${ws.Cells[i + 3, 35]}-ABS(${ws.Cells[i + 3, 36]}))";
                        ws.Cells[i + 3, 42].Formula = $"=IF(${ws.Cells[i + 3, 40]} < 0, SUM(${ws.Cells[i + 3, 40]}:${ws.Cells[i + 3, 41]}),${ws.Cells[i + 3, 40]}-ABS(${ws.Cells[i + 3, 41]}))";
                        ws.Cells[i + 3, 47].Formula = $"=IF(${ws.Cells[i + 3, 45]} < 0, SUM(${ws.Cells[i + 3, 45]}:${ws.Cells[i + 3, 46]}),${ws.Cells[i + 3, 45]}-ABS(${ws.Cells[i + 3, 46]}))";
                        ws.Cells[i + 3, 52].Formula = $"=IF(${ws.Cells[i + 3, 50]} < 0, SUM(${ws.Cells[i + 3, 50]}:${ws.Cells[i + 3, 51]}),${ws.Cells[i + 3, 50]}-ABS(${ws.Cells[i + 3, 51]}))";
                        ws.Cells[i + 3, 57].Formula = $"=IF(${ws.Cells[i + 3, 55]} < 0, SUM(${ws.Cells[i + 3, 55]}:${ws.Cells[i + 3, 56]}),${ws.Cells[i + 3, 55]}-ABS(${ws.Cells[i + 3, 56]}))";

                        ws.Cells[i + 3, 8].Formula = $"=SUM(${ws.Cells[i + 3, 6]},${ws.Cells[i + 3, 3]})-SUM(${ws.Cells[i + 3, 2]},${ws.Cells[i + 3, 5]})";
                        ws.Cells[i + 3, 9].Formula = $"=${ws.Cells[i + 3, 6]}-${ws.Cells[i + 3, 3]}";
                        ws.Cells[i + 3, 13].Formula = $"=SUM(${ws.Cells[i + 3, 11]},${ws.Cells[i + 3, 6]})-SUM(${ws.Cells[i + 3, 5]},${ws.Cells[i + 3, 10]})";
                        ws.Cells[i + 3, 14].Formula = $"=${ws.Cells[i + 3, 11]}-${ws.Cells[i + 3, 6]}";
                        ws.Cells[i + 3, 18].Formula = $"=SUM(${ws.Cells[i + 3, 16]},${ws.Cells[i + 3, 11]})-SUM(${ws.Cells[i + 3, 10]},${ws.Cells[i + 3, 15]})";
                        ws.Cells[i + 3, 19].Formula = $"=${ws.Cells[i + 3, 16]}-${ws.Cells[i + 3, 11]}";
                        ws.Cells[i + 3, 23].Formula = $"=SUM(${ws.Cells[i + 3, 21]},${ws.Cells[i + 3, 16]})-SUM(${ws.Cells[i + 3, 15]},${ws.Cells[i + 3, 20]})";
                        ws.Cells[i + 3, 24].Formula = $"=${ws.Cells[i + 3, 21]}-${ws.Cells[i + 3, 16]}";
                        ws.Cells[i + 3, 28].Formula = $"=SUM(${ws.Cells[i + 3, 26]},${ws.Cells[i + 3, 21]})-SUM(${ws.Cells[i + 3, 20]},${ws.Cells[i + 3, 25]})";
                        ws.Cells[i + 3, 29].Formula = $"=${ws.Cells[i + 3, 26]}-${ws.Cells[i + 3, 21]}";
                        ws.Cells[i + 3, 33].Formula = $"=SUM(${ws.Cells[i + 3, 31]},${ws.Cells[i + 3, 26]})-SUM(${ws.Cells[i + 3, 25]},${ws.Cells[i + 3, 30]})";
                        ws.Cells[i + 3, 34].Formula = $"=${ws.Cells[i + 3, 31]}-${ws.Cells[i + 3, 26]}";
                        ws.Cells[i + 3, 38].Formula = $"=SUM(${ws.Cells[i + 3, 36]},${ws.Cells[i + 3, 31]})-SUM(${ws.Cells[i + 3, 30]},${ws.Cells[i + 3, 35]})";
                        ws.Cells[i + 3, 39].Formula = $"=${ws.Cells[i + 3, 36]}-${ws.Cells[i + 3, 31]}";
                        ws.Cells[i + 3, 43].Formula = $"=SUM(${ws.Cells[i + 3, 41]},${ws.Cells[i + 3, 36]})-SUM(${ws.Cells[i + 3, 35]},${ws.Cells[i + 3, 40]})";
                        ws.Cells[i + 3, 44].Formula = $"=${ws.Cells[i + 3, 41]}-${ws.Cells[i + 3, 36]}";
                        ws.Cells[i + 3, 48].Formula = $"=SUM(${ws.Cells[i + 3, 46]},${ws.Cells[i + 3, 41]})-SUM(${ws.Cells[i + 3, 40]},${ws.Cells[i + 3, 45]})";
                        ws.Cells[i + 3, 49].Formula = $"=${ws.Cells[i + 3, 46]}-${ws.Cells[i + 3, 41]}";
                        ws.Cells[i + 3, 53].Formula = $"=SUM(${ws.Cells[i + 3, 51]},${ws.Cells[i + 3, 46]})-SUM(${ws.Cells[i + 3, 45]},${ws.Cells[i + 3, 50]})";
                        ws.Cells[i + 3, 54].Formula = $"=${ws.Cells[i + 3, 51]}-${ws.Cells[i + 3, 46]}";
                        ws.Cells[i + 3, 58].Formula = $"=SUM(${ws.Cells[i + 3, 56]},${ws.Cells[i + 3, 51]})-SUM(${ws.Cells[i + 3, 50]},${ws.Cells[i + 3, 55]})";
                        ws.Cells[i + 3, 59].Formula = $"=${ws.Cells[i + 3, 56]}-${ws.Cells[i + 3, 51]}";
                    }
                }

                ws.Cells[2, 2].Formula = $"=SUM(${ws.Cells[3, 2]}:${ws.Cells[ws.Dimension.End.Row, 2]})";
                ws.Cells[2, 3].Formula = $"=SUM(${ws.Cells[3, 3]}:${ws.Cells[ws.Dimension.End.Row, 3]})";
                ws.Cells[2, 4].Formula = $"=SUM(${ws.Cells[3, 4]}:${ws.Cells[ws.Dimension.End.Row, 4]})";
                ws.Cells[2, 5].Formula = $"=SUM(${ws.Cells[3, 5]}:${ws.Cells[ws.Dimension.End.Row, 5]})";
                ws.Cells[2, 6].Formula = $"=SUM(${ws.Cells[3, 6]}:${ws.Cells[ws.Dimension.End.Row, 6]})";
                ws.Cells[2, 7].Formula = $"=SUM(${ws.Cells[3, 7]}:${ws.Cells[ws.Dimension.End.Row, 7]})";
                ws.Cells[2, 8].Formula = $"=SUM(${ws.Cells[3, 8]}:${ws.Cells[ws.Dimension.End.Row, 8]})";
                ws.Cells[2, 9].Formula = $"=SUM(${ws.Cells[3, 9]}:${ws.Cells[ws.Dimension.End.Row, 9]})";
                ws.Cells[2, 10].Formula = $"=SUM(${ws.Cells[3, 10]}:${ws.Cells[ws.Dimension.End.Row, 10]})";
                ws.Cells[2, 11].Formula = $"=SUM(${ws.Cells[3, 11]}:${ws.Cells[ws.Dimension.End.Row, 11]})";
                ws.Cells[2, 12].Formula = $"=SUM(${ws.Cells[3, 12]}:${ws.Cells[ws.Dimension.End.Row, 12]})";
                ws.Cells[2, 13].Formula = $"=SUM(${ws.Cells[3, 13]}:${ws.Cells[ws.Dimension.End.Row, 13]})";
                ws.Cells[2, 14].Formula = $"=SUM(${ws.Cells[3, 14]}:${ws.Cells[ws.Dimension.End.Row, 14]})";
                ws.Cells[2, 15].Formula = $"=SUM(${ws.Cells[3, 15]}:${ws.Cells[ws.Dimension.End.Row, 15]})";
                ws.Cells[2, 16].Formula = $"=SUM(${ws.Cells[3, 16]}:${ws.Cells[ws.Dimension.End.Row, 16]})";
                ws.Cells[2, 17].Formula = $"=SUM(${ws.Cells[3, 17]}:${ws.Cells[ws.Dimension.End.Row, 17]})";
                ws.Cells[2, 18].Formula = $"=SUM(${ws.Cells[3, 18]}:${ws.Cells[ws.Dimension.End.Row, 18]})";
                ws.Cells[2, 19].Formula = $"=SUM(${ws.Cells[3, 19]}:${ws.Cells[ws.Dimension.End.Row, 19]})";
                ws.Cells[2, 20].Formula = $"=SUM(${ws.Cells[3, 20]}:${ws.Cells[ws.Dimension.End.Row, 20]})";
                ws.Cells[2, 21].Formula = $"=SUM(${ws.Cells[3, 21]}:${ws.Cells[ws.Dimension.End.Row, 21]})";
                ws.Cells[2, 22].Formula = $"=SUM(${ws.Cells[3, 22]}:${ws.Cells[ws.Dimension.End.Row, 22]})";
                ws.Cells[2, 23].Formula = $"=SUM(${ws.Cells[3, 23]}:${ws.Cells[ws.Dimension.End.Row, 23]})";
                ws.Cells[2, 24].Formula = $"=SUM(${ws.Cells[3, 24]}:${ws.Cells[ws.Dimension.End.Row, 24]})";
                ws.Cells[2, 25].Formula = $"=SUM(${ws.Cells[3, 25]}:${ws.Cells[ws.Dimension.End.Row, 25]})";
                ws.Cells[2, 26].Formula = $"=SUM(${ws.Cells[3, 26]}:${ws.Cells[ws.Dimension.End.Row, 26]})";
                ws.Cells[2, 27].Formula = $"=SUM(${ws.Cells[3, 27]}:${ws.Cells[ws.Dimension.End.Row, 27]})";
                ws.Cells[2, 28].Formula = $"=SUM(${ws.Cells[3, 28]}:${ws.Cells[ws.Dimension.End.Row, 28]})";
                ws.Cells[2, 29].Formula = $"=SUM(${ws.Cells[3, 29]}:${ws.Cells[ws.Dimension.End.Row, 29]})";
                ws.Cells[2, 30].Formula = $"=SUM(${ws.Cells[3, 30]}:${ws.Cells[ws.Dimension.End.Row, 30]})";
                ws.Cells[2, 31].Formula = $"=SUM(${ws.Cells[3, 31]}:${ws.Cells[ws.Dimension.End.Row, 31]})";
                ws.Cells[2, 32].Formula = $"=SUM(${ws.Cells[3, 32]}:${ws.Cells[ws.Dimension.End.Row, 32]})";
                ws.Cells[2, 33].Formula = $"=SUM(${ws.Cells[3, 33]}:${ws.Cells[ws.Dimension.End.Row, 33]})";
                ws.Cells[2, 34].Formula = $"=SUM(${ws.Cells[3, 34]}:${ws.Cells[ws.Dimension.End.Row, 34]})";
                ws.Cells[2, 35].Formula = $"=SUM(${ws.Cells[3, 35]}:${ws.Cells[ws.Dimension.End.Row, 35]})";
                ws.Cells[2, 36].Formula = $"=SUM(${ws.Cells[3, 36]}:${ws.Cells[ws.Dimension.End.Row, 36]})";
                ws.Cells[2, 37].Formula = $"=SUM(${ws.Cells[3, 37]}:${ws.Cells[ws.Dimension.End.Row, 37]})";
                ws.Cells[2, 38].Formula = $"=SUM(${ws.Cells[3, 38]}:${ws.Cells[ws.Dimension.End.Row, 38]})";
                ws.Cells[2, 39].Formula = $"=SUM(${ws.Cells[3, 39]}:${ws.Cells[ws.Dimension.End.Row, 39]})";
                ws.Cells[2, 40].Formula = $"=SUM(${ws.Cells[3, 40]}:${ws.Cells[ws.Dimension.End.Row, 40]})";
                ws.Cells[2, 41].Formula = $"=SUM(${ws.Cells[3, 41]}:${ws.Cells[ws.Dimension.End.Row, 41]})";
                ws.Cells[2, 42].Formula = $"=SUM(${ws.Cells[3, 42]}:${ws.Cells[ws.Dimension.End.Row, 42]})";
                ws.Cells[2, 43].Formula = $"=SUM(${ws.Cells[3, 43]}:${ws.Cells[ws.Dimension.End.Row, 43]})";
                ws.Cells[2, 44].Formula = $"=SUM(${ws.Cells[3, 44]}:${ws.Cells[ws.Dimension.End.Row, 44]})";
                ws.Cells[2, 45].Formula = $"=SUM(${ws.Cells[3, 45]}:${ws.Cells[ws.Dimension.End.Row, 45]})";
                ws.Cells[2, 46].Formula = $"=SUM(${ws.Cells[3, 46]}:${ws.Cells[ws.Dimension.End.Row, 46]})";
                ws.Cells[2, 47].Formula = $"=SUM(${ws.Cells[3, 47]}:${ws.Cells[ws.Dimension.End.Row, 47]})";
                ws.Cells[2, 48].Formula = $"=SUM(${ws.Cells[3, 48]}:${ws.Cells[ws.Dimension.End.Row, 48]})";
                ws.Cells[2, 49].Formula = $"=SUM(${ws.Cells[3, 49]}:${ws.Cells[ws.Dimension.End.Row, 49]})";
                ws.Cells[2, 50].Formula = $"=SUM(${ws.Cells[3, 50]}:${ws.Cells[ws.Dimension.End.Row, 50]})";
                ws.Cells[2, 51].Formula = $"=SUM(${ws.Cells[3, 51]}:${ws.Cells[ws.Dimension.End.Row, 51]})";
                ws.Cells[2, 52].Formula = $"=SUM(${ws.Cells[3, 52]}:${ws.Cells[ws.Dimension.End.Row, 52]})";
                ws.Cells[2, 53].Formula = $"=SUM(${ws.Cells[3, 53]}:${ws.Cells[ws.Dimension.End.Row, 53]})";
                ws.Cells[2, 54].Formula = $"=SUM(${ws.Cells[3, 54]}:${ws.Cells[ws.Dimension.End.Row, 54]})";
                ws.Cells[2, 55].Formula = $"=SUM(${ws.Cells[3, 55]}:${ws.Cells[ws.Dimension.End.Row, 55]})";
                ws.Cells[2, 56].Formula = $"=SUM(${ws.Cells[3, 56]}:${ws.Cells[ws.Dimension.End.Row, 56]})";
                ws.Cells[2, 57].Formula = $"=SUM(${ws.Cells[3, 57]}:${ws.Cells[ws.Dimension.End.Row, 57]})";
                ws.Cells[2, 58].Formula = $"=SUM(${ws.Cells[3, 58]}:${ws.Cells[ws.Dimension.End.Row, 58]})";
                ws.Cells[2, 59].Formula = $"=SUM(${ws.Cells[3, 59]}:${ws.Cells[ws.Dimension.End.Row, 59]})";

                ws.Columns[1].Width = 34;
                for (int i = 2; i <= 59; i++)
                {
                    ws.Columns[i].Width = 12;
                }

                for (int i = 1; i <= 59; i++)
                {
                    ws.Cells[2, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[2, i].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                }

                for (int i = 3; i <= ws.Dimension.End.Row; i++)
                {
                    ws.Cells[i, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 2].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 3].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 5].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 6].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 10].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 11].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 15].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 16].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 16].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 20].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 20].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 21].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 21].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 25].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 25].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 26].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 26].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 30].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 30].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 31].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 31].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 35].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 35].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 36].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 36].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 40].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 40].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 41].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 41].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 45].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 45].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 46].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 46].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 50].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 50].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 51].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 51].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 55].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 55].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                    ws.Cells[i, 56].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[i, 56].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 225, 242));
                }

                ws.Rows[1].Height = 34;
                ws.Rows[1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Rows[1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Rows[1].Style.WrapText = true;
                ws.Rows[1].Style.Font.Size = 9;
                ws.Rows[1].Style.Font.Italic = true;
                ws.Rows[1].Style.Font.Bold = true;
                
                ws.Columns[2].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[3].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[4].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[5].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[6].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[7].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[8].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[9].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[10].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[11].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[12].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[13].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[14].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[15].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[16].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[17].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[18].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[19].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[20].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[21].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[22].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[23].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[24].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[25].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[26].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[27].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[28].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[29].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[30].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[31].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[32].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[33].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[34].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[35].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[36].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[37].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[38].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[39].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[40].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[41].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[42].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[43].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[44].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[45].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[46].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[47].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[48].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[49].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[50].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[51].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[52].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[53].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[54].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[55].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[56].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[57].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[58].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Columns[59].Style.Numberformat.Format = "#,##0;(#,##0)";
                ws.Cells["A1"].Value = "Rubro";
                ws.Cells["A1"].Style.Font.Size = 10;
                ws.Cells["A2"].Value = "GASTO TOTAL";
                ws.Cells["B1"].Value = "Presupuesto Enero";
                ws.Cells["C1"].Value = "Real Enero";
                ws.Cells["D1"].Value = "Variacion Ppto";
                ws.Cells["E1"].Value = "Presupuesto Febrero";
                ws.Cells["F1"].Value = "Real Febrero";
                ws.Cells["G1"].Value = "Variacion Ppto";
                ws.Cells["H1"].Value = "Variacion Real Acumulado Vs Ppto Acumulado";
                ws.Cells["I1"].Value = "Variacion Real Mes Anterior";
                ws.Cells["J1"].Value = "Presupuesto Marzo";
                ws.Cells["K1"].Value = "Real Marzo";
                ws.Cells["L1"].Value = "Variacion Ppto";
                ws.Cells["M1"].Value = "Variacion Real Acumulado Vs Ppto Acumulado";
                ws.Cells["N1"].Value = "Variacion Real Mes Anterior";
                ws.Cells["O1"].Value = "Presupuesto Abril";
                ws.Cells["P1"].Value = "Real Abril";
                ws.Cells["Q1"].Value = "Variacion Ppto";
                ws.Cells["R1"].Value = "Variacion Real Acumulado Vs Ppto Acumulado";
                ws.Cells["S1"].Value = "Variacion Real Mes Anterior";
                ws.Cells["T1"].Value = "Presupuesto Mayo";
                ws.Cells["U1"].Value = "Real Mayo";
                ws.Cells["V1"].Value = "Variacion Ppto";
                ws.Cells["W1"].Value = "Variacion Real Acumulado Vs Ppto Acumulado";
                ws.Cells["X1"].Value = "Variacion Real Mes Anterior";
                ws.Cells["Y1"].Value = "Presupuesto Junio";
                ws.Cells["Z1"].Value = "Real Junio";
                ws.Cells["AA1"].Value = "Variacion Ppto";
                ws.Cells["AB1"].Value = "Variacion Real Acumulado Vs Ppto Acumulado";
                ws.Cells["AC1"].Value = "Variacion Real Mes Anterior";
                ws.Cells["AD1"].Value = "Presupuesto Julio";
                ws.Cells["AE1"].Value = "Real Julio";
                ws.Cells["AF1"].Value = "Variacion Ppto";
                ws.Cells["AG1"].Value = "Variacion Real Acumulado Vs Ppto Acumulado";
                ws.Cells["AH1"].Value = "Variacion Real Mes Anterior";
                ws.Cells["AI1"].Value = "Presupuesto Agosto";
                ws.Cells["AJ1"].Value = "Real Agosto";
                ws.Cells["AK1"].Value = "Variacion Ppto";
                ws.Cells["AL1"].Value = "Variacion Real Acumulado Vs Ppto Acumulado";
                ws.Cells["AM1"].Value = "Variacion Real Mes Anterior";
                ws.Cells["AN1"].Value = "Presupuesto Septiembre";
                ws.Cells["AO1"].Value = "Real Septiembre";
                ws.Cells["AP1"].Value = "Variacion Ppto";
                ws.Cells["AQ1"].Value = "Variacion Real Acumulado Vs Ppto Acumulado";
                ws.Cells["AR1"].Value = "Variacion Real Mes Anterior";
                ws.Cells["AS1"].Value = "Presupuesto Octubre";
                ws.Cells["AT1"].Value = "Real Octubre";
                ws.Cells["AU1"].Value = "Variacion Ppto";
                ws.Cells["AV1"].Value = "Variacion Real Acumulado Vs Ppto Acumulado";
                ws.Cells["AW1"].Value = "Variacion Real Mes Anterior";
                ws.Cells["AX1"].Value = "Presupuesto Noviembre";
                ws.Cells["AY1"].Value = "Real Noviembre";
                ws.Cells["AZ1"].Value = "Variacion Ppto";
                ws.Cells["BA1"].Value = "Variacion Real Acumulado Vs Ppto Acumulado";
                ws.Cells["BB1"].Value = "Variacion Real Mes Anterior";
                ws.Cells["BC1"].Value = "Presupuesto Diciembre";
                ws.Cells["BD1"].Value = "Real Diciembre";
                ws.Cells["BE1"].Value = "Variacion Ppto";
                ws.Cells["BF1"].Value = "Variacion Real Acumulado Vs Ppto Acumulado";
                ws.Cells["BG1"].Value = "Variacion Real Mes Anterior";

                ws.Workbook.Calculate();


                await package.SaveAsync();
                file = package.GetAsByteArray();

            }

            if (file.Length > 1)
            {
                return new DataResponse()
                {
                    Code = 1,
                    Message = "El reporte se ha creado satisfactoriamente.",
                    Base64 = Convert.ToBase64String(file)
                };
            }

            return new DataResponse()
            {
                Code = 0,
                Message = "Se ha producido un error",
                Base64 = ""
            };
        }
    }
}
