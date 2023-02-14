using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DNPrueba.API
{
    public class SPN
    {
        public static string GetMonthBudget = "dbo.GetMonthlyBudgets";
        public static string GetOrderedExpensesList = "dbo.GetOrderedExpensesList";
        public static string GetPABalanceByMonth = "dbo.GetPABalanceByMonth";
        public static string GetCOBalanceByMonth = "dbo.GetCOBalanceByMonth";
    }
}
