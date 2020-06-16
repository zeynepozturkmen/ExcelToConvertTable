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
        public JsonResult SaveDatabase(IEnumerable<PersonTable> PersonTable)
        {
            var persons = new List<Person>();
            var isSave = 0;

            foreach (var item in PersonTable)
            {
                var person = new Person();

                var AnyPerson = db.People.Any(x => x.Tc == item.TC);
                if (AnyPerson == false)
                {
                    person.Tc = item.TC;
                    person.FullName = item.FullName;
                    person.Address = item.Address;
                    person.PhoneNumber = item.PhoneNumber;
                    person.Email = item.Email;

                    persons.Add(person);
                }
            }

            if (persons != null)
            {
                db.People.AddRange(persons);
                isSave = db.SaveChanges();

                return Json(true, JsonRequestBehavior.AllowGet);
            }
            //veri tabanına daha önceden kayıt edilmiş demek
            else
            {
                return Json(false, JsonRequestBehavior.AllowGet);
            }


        }

    }
}