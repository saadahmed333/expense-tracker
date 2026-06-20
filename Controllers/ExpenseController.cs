using Microsoft.AspNetCore.Mvc;
using ExpenseTracker.Models;

namespace ExpenseTracker.Controllers
{
    public class ExpenseController : Controller
    {
        // GET: Expense/Index
        public IActionResult Index(string? category, string? search, string? dateFrom, string? dateTo)
        {
            var expenses = ExpenseStore.Expenses.AsQueryable();

            if (!string.IsNullOrEmpty(category) && category != "All")
                expenses = expenses.Where(e => e.Category == category);

            if (!string.IsNullOrEmpty(search))
                expenses = expenses.Where(e => e.Description.Contains(search, StringComparison.OrdinalIgnoreCase));

            if (DateTime.TryParse(dateFrom, out var from))
                expenses = expenses.Where(e => e.Date >= from);

            if (DateTime.TryParse(dateTo, out var to))
                expenses = expenses.Where(e => e.Date <= to);

            ViewBag.Categories = ExpenseStore.Categories;
            ViewBag.SelectedCategory = category;
            ViewBag.Search = search;
            ViewBag.DateFrom = dateFrom;
            ViewBag.DateTo = dateTo;
            ViewBag.Total = expenses.Sum(e => e.Amount);

            return View(expenses.OrderByDescending(e => e.Date).ToList());
        }

        // GET: Expense/Create
        public IActionResult Create()
        {
            ViewBag.Categories = ExpenseStore.Categories;
            var model = new Expense { Date = DateTime.Today };
            return View(model);
        }

        // POST: Expense/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create(Expense expense)
        {
            if (ModelState.IsValid)
            {
                expense.Id = ExpenseStore.GetNextId();
                ExpenseStore.Expenses.Add(expense);
                TempData["Success"] = "Expense added successfully!";
                return RedirectToAction(nameof(Index));
            }
            ViewBag.Categories = ExpenseStore.Categories;
            return View(expense);
        }

        // GET: Expense/Edit/5
        public IActionResult Edit(int id)
        {
            var expense = ExpenseStore.Expenses.FirstOrDefault(e => e.Id == id);
            if (expense == null) return NotFound();

            ViewBag.Categories = ExpenseStore.Categories;
            return View(expense);
        }

        // POST: Expense/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Edit(int id, Expense updated)
        {
            if (ModelState.IsValid)
            {
                var expense = ExpenseStore.Expenses.FirstOrDefault(e => e.Id == id);
                if (expense == null) return NotFound();

                expense.Description = updated.Description;
                expense.Amount = updated.Amount;
                expense.Category = updated.Category;
                expense.Date = updated.Date;
                expense.Notes = updated.Notes;

                TempData["Success"] = "Expense updated successfully!";
                return RedirectToAction(nameof(Index));
            }
            ViewBag.Categories = ExpenseStore.Categories;
            return View(updated);
        }

        // POST: Expense/Delete/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Delete(int id)
        {
            var expense = ExpenseStore.Expenses.FirstOrDefault(e => e.Id == id);
            if (expense != null)
            {
                ExpenseStore.Expenses.Remove(expense);
                TempData["Success"] = "Expense deleted successfully!";
            }
            return RedirectToAction(nameof(Index));
        }
    }
}
