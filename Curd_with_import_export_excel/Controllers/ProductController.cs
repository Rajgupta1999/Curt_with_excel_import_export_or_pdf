using Curd_with_import_export_excel.Models;
using Curd_with_import_export_excel.Services.Implementation;
using Curd_with_import_export_excel.Services.Interface;
using Microsoft.AspNetCore.Mvc;

namespace Curd_with_import_export_excel.Controllers
{
   
    public class ProductController : Controller
    {
        private readonly IProductService _service;

        public ProductController(IProductService service)
        {
            _service = service;
        }
        public async Task<IActionResult> Index()
        {
            var products = await _service.GetProductsAsync();
            return View(products);
        }

        public IActionResult Create() => View();

        [HttpPost]
        public async Task<IActionResult> Create(Product product)
        {
            if (ModelState.IsValid)
            {
                await _service.AddAsync(product);
                return RedirectToAction("Index");
            }
            return View(product);
        }

        public async Task<IActionResult> Edit(int id)
        {
            var product = await _service.GetProductByIdAsync(id);
            if (product == null) return NotFound();
            return View(product);
        }

        [HttpPost]
        public async Task<IActionResult> Edit(int id, Product product)
        {
            if (id != product.Id) return BadRequest();
            if (ModelState.IsValid)
            {
                await _service.UpdateAsync(product);
                return RedirectToAction("Index");
            }
            return View(product);
        }

        public async Task<IActionResult> Delete(int id)
        {
            var product = await _service.GetProductByIdAsync(id);
            if (product == null) return NotFound();
            return View(product);
        }

        [HttpPost, ActionName("Delete")]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            await _service.DeleteAsync(id);
            return RedirectToAction("Index");
        }


        public IActionResult Import() => View();

        [HttpPost]
        public async Task<IActionResult> Import(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("Invalid file.");

            await _service.ImportFromExcelAsync(file);
            return RedirectToAction("Index");
        }

        public async Task<IActionResult> Export()
        {
            var fileContent = await _service.ExportToExcelAsync();
            return File(fileContent, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Products.xlsx");
        }

        public async Task<IActionResult> ExportPdf()
        {
            var filecontent =await _service.ExportToPdfAsync();
            return File(filecontent, "application/pdf", "product.pdf");
        }
    }
}
