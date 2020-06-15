using ExcelToConvertTable.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToConvertTable.Controllers
{
    public class HomeController : Controller
    {
        PersonDbEntities db = new PersonDbEntities();
        public ActionResult Index()
        {
            PersonModel model = new PersonModel();
            return View(model);
        }


        [HttpPost]
        public PartialViewResult UploadExcel(HttpPostedFileBase excelFile)
        {
            PersonModel model = new PersonModel();

            if (excelFile == null
            || excelFile.ContentLength == 0)
            {
                ViewBag.Error = "Lütfen dosya seçimi yapınız.";

                return PartialView("~/Views/Home/_PartialTableList.cshtml", model);
            }
            else
            {
                //Dosyanın uzantısı xls ya da xlsx ise;
                if (excelFile.FileName.EndsWith("xls")
                || excelFile.FileName.EndsWith("xlsx"))
                {

                    //Seçilen dosyanın nereye yükleneceği seçiliyor.
                    string path = Server.MapPath("~/Documents/" + excelFile.FileName);

                    //Dosya kontrol edilir, varsa silinir.
                    //Dosya yoksa kaydet
                    if (!System.IO.File.Exists(path))
                    {
                        //System.IO.File.Delete(path);
                        //Excel path altına kaydedilir.
                        excelFile.SaveAs(path);
                    }

                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(path);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;


                    if (range.Columns.Count == 5)
                    {
                        model.PersonList = new List<PersonTable>();

                        for (int i = 2; i <= range.Rows.Count; i++)
                        {
                            model.Person = new PersonTable();

                            model.Person.FullName = range.Cells[i, 1].Text;
                            model.Person.Email = ((Excel.Range)range.Cells[i, 2]).Text;
                            model.Person.Address = ((Excel.Range)range.Cells[i, 3]).Text;
                            model.Person.PhoneNumber = ((Excel.Range)range.Cells[i, 4]).Text;
                            model.Person.TC = ((Excel.Range)range.Cells[i, 5]).Text;

                            model.PersonList.Add(model.Person);
                        }

                        application.Quit();
                        return PartialView("~/Views/Home/_PartialTableList.cshtml", model);
                    }
                    else
                    {
                        ViewBag.Error = "Exceldeki kolon sayısı istenilen sayıda değil";
                        return PartialView("~/Views/Home/_PartialTableList.cshtml", model);
                    }
                }
                else
                {
                    ViewBag.Error = "Dosya tipiniz yanlış, lütfen '.xls' yada '.xlsx' uzantılı dosya yükleyiniz.";
                    return PartialView("~/Views/Home/_PartialTableList.cshtml", model);
                }
            }
        }

        [HttpPost]
        public JsonResult SaveDatabase(IEnumerable<PersonTable> PersonList)
        {
            var persons = new List<Person>();

            foreach (var item in PersonList)
            {
                var person = new Person();
                person.Tc = item.TC;
                person.FullName = item.FullName;
                person.Address = item.Address;
                person.PhoneNumber = item.PhoneNumber;
                person.Email = item.Email;

                persons.Add(person);
            }

            db.People.AddRange(persons);
            var isSave = db.SaveChanges();
            if (isSave > 0)
            {
                return Json(true, JsonRequestBehavior.AllowGet);
            }
            else
            {
                return Json(false, JsonRequestBehavior.AllowGet);
            }
        }

    }
}