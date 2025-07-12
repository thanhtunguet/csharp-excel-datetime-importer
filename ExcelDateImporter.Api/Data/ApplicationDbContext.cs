using Microsoft.EntityFrameworkCore;
using ExcelDateImporter.Api.Models;

namespace ExcelDateImporter.Api.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options)
        {
        }

        public DbSet<ExcelDateEntry> ExcelDateEntries { get; set; }
    }
}