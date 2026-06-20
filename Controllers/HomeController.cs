using Microsoft.AspNetCore.Mvc;
using ExpenseTracker.Models;

namespace ExpenseTracker.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            var expenses = ExpenseStore.Expenses;

            ViewBag.TotalAmount = expenses.Sum(e => e.Amount);
            ViewBag.MonthTotal = expenses
                .Where(e => e.Date.Month == DateTime.Today.Month && e.Date.Year == DateTime.Today.Year)
                .Sum(e => e.Amount);
            ViewBag.TransactionCount = expenses.Count;
            ViewBag.TopCategory = expenses
                .GroupBy(e => e.Category)
                .OrderByDescending(g => g.Sum(e => e.Amount))
                .FirstOrDefault()?.Key ?? "N/A";

            ViewBag.CategoryTotals = expenses
                .GroupBy(e => e.Category)
                .Select(g => new CategoryTotal { Category = g.Key, Total = g.Sum(e => e.Amount) })
                .OrderByDescending(x => x.Total)
                .ToList();

            ViewBag.RecentExpenses = expenses
                .OrderByDescending(e => e.Date)
                .Take(5)
                .ToList();

            return View();
        }
    }
}
