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

        // GET: api/<customerController>
        [HttpGet("all")]
        public ActionResult<IEnumerable<Customer>> Get()
        {
            return _northwindContext.Customers.ToList();
        }

        // GET api/<customerController>/5
        [HttpGet("find")]
        public ActionResult<Customer> Get([FromBody] Customer value)
        {
            var result = _northwindContext.Customers.Find(value.CustomerId);
            if (result == null)
            {
                return NotFound("NotFound");
            }
            return result;
        }

        // POST api/<customerController>
        [HttpPost("add")]
        public ActionResult<Customer> Post([FromBody] Customer value)
        {
            _northwindContext.Customers.Add(value);
            _northwindContext.SaveChanges();
            return CreatedAtAction(nameof(Get), new { customerId = value.CustomerId }, value.CustomerId);
        }

        // PUT api/<customerController>/5
        [HttpPut("edit")]
        public IActionResult Put([FromBody] Customer value)
        {
            if (value.CustomerId != null)
            {
                _northwindContext.Entry(value).State = EntityState.Modified;
                _northwindContext.SaveChanges();
                return Content("成功(true)");
            }
            else
            {
                return NotFound("失敗(false)");

            }
        }

        // DELETE api/<customerController>/5
        [HttpDelete("delete")]
        public IActionResult Delete([FromBody] Customer value)
        {
            _northwindContext.Customers.Remove(value);
            _northwindContext.SaveChanges();
            return Content("成功(true)");
        }
    }
}
