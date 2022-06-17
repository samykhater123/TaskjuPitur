using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using TaskjuPitur.DTOs;
using TaskjuPitur.Model;

namespace TaskjuPitur.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EmplyeeController : ControllerBase
    {

        private readonly AppDBContext db;

        public EmplyeeController( AppDBContext _db)
        {  
            db = _db;
        }

       
        [HttpGet]
        public async Task<IActionResult> getdata()
        {
            var all = db.employees.ToList();

            if (all == null)
            {
                return NotFound();
            }
            return Ok(all);
        }

        [HttpPost]
        public async Task<IActionResult> add([FromBody] EmplyeeDto model)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest();
            }

            var emp = new employee()
            {
                fName=model.fName,
                lName=model.lName,
                HiringDate=model.HiringDate

            };
            db.employees.Add(emp);
            db.SaveChanges();
            return Ok();
            
        }

        [HttpPut]
        public async Task<IActionResult> update(int id, employee model)
        {
            var emp = db.employees.Where(x=>x.id==id).FirstOrDefault();

            if (emp==null)
            {
                return NotFound();
            }

            emp.fName = model.fName;
            emp.lName = model.lName;
            emp.HiringDate = model.HiringDate;

            db.SaveChanges();
            return Ok();
        }

        [HttpDelete]
        public async Task<IActionResult> delete(int id)
        {
            var emp = db.employees.Find(id);
            if (emp==null)
            {
                return NotFound();
            }

            db.employees.Remove(emp);
            db.SaveChanges();
            return Ok();
        }

        [HttpGet("getdatatoexcel")]
        public IActionResult getdatatoexcel()
        {
           
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("employees");
                var currentrow = 1;
                worksheet.Cell(currentrow, 1).Value = "Id";
                worksheet.Cell(currentrow, 2).Value = "FirstName";
                worksheet.Cell(currentrow, 3).Value = "LastName";
                worksheet.Cell(currentrow, 4).Value = "Date";

                var emplyees = db.employees.ToList();

                foreach (var emp in emplyees)
                {
                    currentrow++;
                    worksheet.Cell(currentrow, 1).Value = emp.id;
                    worksheet.Cell(currentrow, 2).Value = emp.fName;
                    worksheet.Cell(currentrow, 3).Value = emp.lName;
                    DateTime d = emp.HiringDate;
                    worksheet.Cell(currentrow, 4).Value = d.ToString("dd/MM/yyyy");

                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var contant = stream.ToArray();
                    return File(contant, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Emplyee.xlsx");
                }
                return Ok();
            }
        }


    }
}
