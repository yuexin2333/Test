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
    public class customerController : ControllerBase
    {
        private readonly NorthwindContext _northwindContext;
        public customerController(NorthwindContext northwindContext)
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
        [HttpGet("find/{customerId}")]
        public ActionResult<Customer> Get(string customerId)
        {
            var result = _northwindContext.Customers.Find(customerId);
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
        public IActionResult Put(string customerId, [FromBody] Customer value)
        {
            _northwindContext.Entry(value).State = Microsoft.EntityFrameworkCore.EntityState.Modified;
            try
            {
                _northwindContext.SaveChanges();
                return Content("成功(true)");
            }
            catch (DbUpdateException)
            {
                if (_northwindContext.Customers.Any(e => e.CustomerId == customerId))
                {
                    return NotFound("失敗(false)");
                }
                else
                {
                    return StatusCode(500, "error");
                }
            }
            return NoContent();
        }

        // DELETE api/<customerController>/5
        [HttpDelete("delete/{customerId}")]
        public IActionResult Delete(string customerId, [FromBody] Customer value)
        {
            if (customerId != value.CustomerId)
            {
                return BadRequest("失敗(false)");
            }
            _northwindContext.Customers.Remove(value);
            _northwindContext.SaveChanges();
            return Content("成功(true)");
        }
    }
}
