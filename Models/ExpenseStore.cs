namespace ExpenseTracker.Models
{
    // In-memory data store (no database required for this lab)
    public static class ExpenseStore
    {
        private static int _nextId = 1;

        public static List<Expense> Expenses { get; } = new List<Expense>
        {
            new Expense { Id = _nextId++, Description = "Grocery shopping", Amount = 85.50m, Category = "Food", Date = DateTime.Today.AddDays(-1), Notes = "Weekly groceries" },
            new Expense { Id = _nextId++, Description = "Bus pass", Amount = 30.00m, Category = "Transport", Date = DateTime.Today.AddDays(-2), Notes = "Monthly bus card" },
            new Expense { Id = _nextId++, Description = "Netflix subscription", Amount = 15.99m, Category = "Entertainment", Date = DateTime.Today.AddDays(-3) },
            new Expense { Id = _nextId++, Description = "Electricity bill", Amount = 120.00m, Category = "Bills", Date = DateTime.Today.AddDays(-5), Notes = "Monthly electricity" },
            new Expense { Id = _nextId++, Description = "Restaurant dinner", Amount = 45.00m, Category = "Food", Date = DateTime.Today.AddDays(-6) },
        };

        public static readonly List<string> Categories = new List<string>
        {
            "Food", "Transport", "Shopping", "Entertainment", "Bills", "Health", "Education", "Other"
        };

        public static int GetNextId() => _nextId++;
    }
}
