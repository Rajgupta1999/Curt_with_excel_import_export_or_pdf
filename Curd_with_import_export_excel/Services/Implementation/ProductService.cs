using Curd_with_import_export_excel.Data;
using Curd_with_import_export_excel.Models;
using Curd_with_import_export_excel.Services.Interface;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.Diagnostics;
using System.Reflection.Metadata;
using Document = iTextSharp.text.Document;

namespace Curd_with_import_export_excel.Services.Implementation
{
    public class ProductService : IProductService
    {
        private readonly ProductDbContext _context;

        public ProductService(ProductDbContext context)
        {
            _context = context;
        }

        public async Task<List<Product>> GetProductsAsync() => await _context.Products.ToListAsync();


        public async Task<Product> GetProductByIdAsync(int id) => await _context.Products.FindAsync(id);

        public async Task AddAsync(Product product)
        {
            _context.Products.Add(product);
            await _context.SaveChangesAsync();
        }

        public async Task UpdateAsync(Product product)
        {
           _context.Products.Update(product);
            await _context.SaveChangesAsync();
        }

        public async Task DeleteAsync(int id)
        {
           var product = await _context.Products.FindAsync(id);
            if(product != null)
            {
                _context.Products.Remove(product);
                await _context.SaveChangesAsync();
            }
        }

        public async Task<byte[]> ExportToExcelAsync()
        {
           var Products =  await  _context.Products.ToListAsync();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using  var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Products");
            worksheet.Cells[1, 1].Value = "Id";
            worksheet.Cells[1,2].Value= "Name";
            worksheet.Cells[1,3].Value= "Description";
            worksheet.Cells[1, 4].Value = "Price";

            int row = 2;
            foreach(var item in Products)
            {
                worksheet.Cells[row,1].Value = item.Id;
                worksheet.Cells[row,2].Value = item.Name;
                worksheet.Cells[row,3].Value = item.Description;
                worksheet.Cells[row, 4].Value = item.Price;
                row++;
            }
            return package.GetAsByteArray();
        }

        public async Task<byte[]> ExportToPdfAsync()
        {
            var products = await _context.Products.ToListAsync();

            using var stream = new MemoryStream();
            var document = new Document(PageSize.A4);
            PdfWriter.GetInstance(document, stream);
            document.Open();

            //Title
            var titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16);
            var title = new Paragraph("Product List\n\n", titleFont);
            title.Alignment = Element.ALIGN_CENTER;
            document.Add(title);

            //Table
            //var HeadingFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16);
            var table = new PdfPTable(3) {WidthPercentage =100};
            table.AddCell("Name");
            table.AddCell("Description");
            table.AddCell("Price");

            foreach(var item in products)
            {
                table.AddCell(item.Name);
                table.AddCell(item.Description);
                table.AddCell(item.Price.ToString("F2"));
            }
            document.Add(table);
            document.Close();
            return stream.ToArray();
        }
        

        public async Task ImportFromExcelAsync(IFormFile file)
        {
            using var stream = new MemoryStream();
            await file.CopyToAsync(stream);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(stream);
            var worksheet = package.Workbook.Worksheets[0];
            var rowcount = worksheet.Dimension.Rows;

            for(int row=2; row<rowcount; row++)
            {
                var product = new Product
                {
                    Name = worksheet.Cells[row,2].Value?.ToString(),
                    Description= worksheet.Cells[row,3].Value?.ToString(),
                    Price=decimal.Parse(worksheet.Cells[row,4].Value?.ToString() ?? "0")
                };

                _context.Products.Add(product);
            }
            await _context.SaveChangesAsync();  
        }
    }
}
