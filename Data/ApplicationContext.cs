using Microsoft.EntityFrameworkCore;
using Microsoft.Identity.Client;
using StudentApplication.Models.Excel;

namespace StudentApplication.Data
{
    public class ApplicationContext : DbContext
    {
        internal readonly object dtDuplicates;

        public ApplicationContext(DbContextOptions<ApplicationContext> options) : base(options) { }

        public DbSet<Customer> Customers { get; set; }


        public DbSet<Student> Students { get; set; }

    }
 
}
