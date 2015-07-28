using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using ExportToExcel.Models;
using System.Web.UI.WebControls;
using System.IO;
using System.Web.UI;

namespace ExportToExcel.Controllers
{
    public class ExportController : Controller
    {
        private ExportConn db = new ExportConn();

        // GET: /Export/


        public ActionResult ExportData()
        {
            GridView gv = new GridView();
            gv.DataSource = db.exports.ToList();
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=Marklist.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter sw = new StringWriter();
            HtmlTextWriter htw = new HtmlTextWriter(sw);
            gv.RenderControl(htw);
            Response.Output.Write(sw.ToString());
            Response.Flush();
            Response.End();

            return RedirectToAction("StudentDetails");
        }

        public ActionResult Index()
        {
            return View(db.exports.ToList());
        }

        // GET: /Export/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            export export = db.exports.Find(id);
            if (export == null)
            {
                return HttpNotFound();
            }
            return View(export);
        }

        // GET: /Export/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: /Export/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include="Id,fname,lname,age")] export export)
        {
            if (ModelState.IsValid)
            {
                db.exports.Add(export);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(export);
        }

        // GET: /Export/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            export export = db.exports.Find(id);
            if (export == null)
            {
                return HttpNotFound();
            }
            return View(export);
        }

        // POST: /Export/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include="Id,fname,lname,age")] export export)
        {
            if (ModelState.IsValid)
            {
                db.Entry(export).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(export);
        }

        // GET: /Export/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            export export = db.exports.Find(id);
            if (export == null)
            {
                return HttpNotFound();
            }
            return View(export);
        }

        // POST: /Export/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            export export = db.exports.Find(id);
            db.exports.Remove(export);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
