using FlexCel.Core;
using FlexCel.XlsAdapter;
using MultipartDataMediaFormatter.Infrastructure;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web;
using System.Web.Mvc;
using webApp.Models;
using webApp.Utilities;

namespace webApp.Controllers
{
    public class ImportController : Controller
    {
        // GET: Import
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public bool Upload(HttpPostedFileBase file)
        {
            MemoryStream target = new MemoryStream();
            file.InputStream.CopyTo(target);
            byte[] data = target.ToArray();
            HttpFile httpFile = new HttpFile(file.FileName, file.ContentType, data);

            DataSet dataSet = ReadExcel.ToDataSet(httpFile.Buffer);

            DataTable resultDt = dataSet.Tables[0];

            List<Student> students = new List<Student>();
            students = Convertor.ConvertToList<Student>(resultDt);

            return true;
        }

        public bool Upload2(HttpPostedFileBase file)
        {
            bool formatted = false;
            List<string> sheetNames = new List<string>();
            DataSet dataSet = new DataSet();

            MemoryStream target = new MemoryStream();
            file.InputStream.CopyTo(target);
            byte[] data = target.ToArray();
            HttpFile httpFile = new HttpFile(file.FileName, file.ContentType, data);

            Stream stream = new MemoryStream(httpFile.Buffer);

            XlsFile xls = new XlsFile(false);
            xls.Open(stream);

            for (int sheet = 1; sheet <= xls.SheetCount; sheet++)
            {
                xls.ActiveSheet = sheet;

                sheetNames.Add(xls.SheetName);

                DataTable Data = dataSet.Tables.Add("Sheet" + sheet.ToString());
                Data.BeginLoadData();
                try
                {
                    int ColCount = xls.ColCount;
                    //Add one column on the dataset for each used column on Excel.
                    //for (int c = 1; c <= ColCount; c++)
                    //{
                    //    Data.Columns.Add(TCellAddress.EncodeColumn(c), typeof(String));  //Here we will add all strings, since we do not know what we are waiting for.
                    //}

                    for (int c = 1; c <= ColCount; c++)
                    {
                        int r = 1;
                        int XF = 0; //This is the cell format, we will not use it here.
                        object val = xls.GetCellValueIndexed(r, c, ref XF);
                        string propName = string.Empty;

                        TFormula Fmla = val as TFormula;
                        propName = Convert.ToString(Fmla != null ? Fmla.Result : val);

                        Data.Columns.Add(propName, typeof(String));
                    }

                    string[] dr = new string[ColCount];

                    int RowCount = xls.RowCount;
                    for (int r = 2; r <= RowCount; r++)
                    {
                        Array.Clear(dr, 0, dr.Length);
                        //This loop will only loop on used cells. It is more efficient than looping on all the columns.
                        for (int cIndex = xls.ColCountInRow(r); cIndex > 0; cIndex--)  //reverse the loop to avoid calling ColCountInRow more than once.
                        {
                            int Col = xls.ColFromIndex(r, cIndex);

                            if (formatted)
                            {
                                TRichString rs = xls.GetStringFromCell(r, Col);
                                dr[Col - 1] = rs.Value;
                            }
                            else
                            {
                                int XF = 0; //This is the cell format, we will not use it here.
                                object val = xls.GetCellValueIndexed(r, cIndex, ref XF);

                                TFormula Fmla = val as TFormula;
                                if (Fmla != null)
                                {
                                    //When we have formulas, we want to write the formula result. 
                                    //If we wanted the formula text, we would not need this part.
                                    dr[Col - 1] = Convert.ToString(Fmla.Result);
                                }
                                else
                                {
                                    dr[Col - 1] = Convert.ToString(val);
                                }
                            }
                        }
                        Data.Rows.Add(dr);
                    }
                }
                finally
                {
                    Data.EndLoadData();
                }
            }

            DataTable resultDt = dataSet.Tables[0];

            List<Student> Studentlist = new List<Student>();
            Studentlist = Convertor.ConvertToList<Student>(resultDt);

            //AutoMapper.Mapper.Initialize(cfg => cfg.CreateMissingTypeMaps = true);
            //var students = AutoMapper.Mapper.Map<IDataReader, IEnumerable<Student>>(resultDt.CreateDataReader());

            return true;
        }
    }
}