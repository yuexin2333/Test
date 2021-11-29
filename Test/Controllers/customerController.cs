using System;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Linq;
using Test.Models;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace Test.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class CustomerController : ControllerBase
    {
        private readonly NorthwindContext _northwindContext;

        public CustomerController(NorthwindContext northwindContext)
        {
            _northwindContext = northwindContext;
        }

        // GET: 顯示所有資料
        [HttpGet("all")]
        public ActionResult<IEnumerable<Customer>> Get()
        {
            return _northwindContext.Customers.ToList();
        }

        // GET 尋找某筆資料
        [HttpGet("find")]
        public ActionResult<Customer> Get([FromBody] Customer value)
        {
            var result = _northwindContext.Customers.Find(value.CustomerId);
            if (result == null)
            {
                return NotFound("找不到這筆資料");
            }

            return result;
        }

        // POST 增加一筆資料
        [HttpPost("add")]
        public ActionResult<Customer> Post([FromBody] Customer value)
        {
            _northwindContext.Customers.Add(value);
            _northwindContext.SaveChanges();
            
            return CreatedAtAction(nameof(Get), value.CustomerId);
            
        }

        // PUT 修改某筆資料
        [HttpPut("edit")]
        public IActionResult Put([FromBody] Customer value)
        {
            return IsEdit(value);
        }

        private IActionResult IsEdit(Customer value)
        {
            return value.CustomerId != null ? SaveEdit(value) : NotFound("修改失敗(false)");
        }

        private IActionResult SaveEdit(Customer value)
        {
            _northwindContext.Entry(value).State = EntityState.Modified;
            _northwindContext.SaveChanges();
            return Content("修改成功(true)");
        }

        // DELETE 刪除某筆資料
        [HttpDelete("delete")]
        public IActionResult Delete([FromBody] Customer value)
        {
            return IsDelete(value);
        }

        private IActionResult IsDelete(Customer value)
        {
            var result = _northwindContext.Customers.Find(value.CustomerId);
            if (result == null) return NotFound("刪除失敗(false)");
            _northwindContext.Customers.Remove(value);
            _northwindContext.SaveChanges();
            return Content("刪除成功(true)");
        }
    }
}