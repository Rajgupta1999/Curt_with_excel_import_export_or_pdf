using Curd_with_import_export_excel.Models;

namespace Curd_with_import_export_excel.Services.Interface
{
    public interface IProductService
    {
        Task<List<Product>> GetProductsAsync();
        Task<Product> GetProductByIdAsync(int id);
        Task AddAsync (Product product);
        Task UpdateAsync (Product product);
        Task DeleteAsync (int id);
        Task<byte[]> ExportToExcelAsync();
        Task<byte[]> ExportToPdfAsync();
        Task ImportFromExcelAsync(IFormFile file);
    }
}
