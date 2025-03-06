using Curd_with_import_export_excel.Models;
using Microsoft.EntityFrameworkCore;

namespace Curd_with_import_export_excel.Data
{
    public class ProductDbContext : DbContext
    {
        public ProductDbContext(DbContextOptions options) : base(options)
        {
        }

        public DbSet<Product> Products { get; set; }
    }
}
