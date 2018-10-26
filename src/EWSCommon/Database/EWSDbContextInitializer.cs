using System;
using System.Data.Entity;

namespace EWS.Common.Database
{
    /// <summary>
    /// Seeds the database with initial data
    /// </summary>
    internal class EWSDbContextInitializer : CreateDatabaseIfNotExists<EWSDbContext>
    {
        /// <summary>
        /// Seeds the specified context.
        /// </summary>
        /// <param name="context">The context.</param>
        protected override void Seed(EWSDbContext context)
        {
            base.Seed(context);


            var runDate = DateTime.UtcNow;
            var firstDayOfMonth = new DateTime(runDate.Year, runDate.Month, 1);
            var formattedDate = firstDayOfMonth.ToString("MM/dd/yyyy");
            var totalDays = DateTime.DaysInMonth(runDate.Year, runDate.Month);



            context.SaveChanges();
        }
    }
}
