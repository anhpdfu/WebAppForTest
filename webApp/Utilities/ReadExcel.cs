using System;
using System.Data;
using System.IO;
using FlexCel.Core;
using FlexCel.XlsAdapter;

namespace webApp.Utilities
{
    public static class ReadExcel
    {
        public static DataSet ToDataSet(byte[] data, bool firstRowIsPropertyName = true)
        {
            bool formatted = false;
            DataSet dataSet = new DataSet();

            Stream stream = new MemoryStream(data);

            XlsFile xls = new XlsFile(false);
            xls.Open(stream);

            for (int sheet = 1; sheet <= xls.SheetCount; sheet++)
            {
                xls.ActiveSheet = sheet;

                DataTable Data = dataSet.Tables.Add("Sheet" + sheet.ToString());
                Data.BeginLoadData();

                try
                {
                    int ColCount = xls.ColCount;
                    int beginIndexRow = 1;

                    if (firstRowIsPropertyName)
                    {
                        beginIndexRow = 2;
                        for (int c = 1; c <= ColCount; c++)
                        {
                            int r = 1;
                            int XF = 0; // This is the cell format, we will not use it here.
                            object val = xls.GetCellValueIndexed(r, c, ref XF);
                            string propName = string.Empty;

                            TFormula Fmla = val as TFormula;
                            propName = Convert.ToString(Fmla != null ? Fmla.Result : val);

                            Data.Columns.Add(propName, typeof(String));
                        }
                    }
                    else
                    {
                        beginIndexRow = 1;
                        // Add one column on the dataset for each used column on Excel.
                        for (int c = 1; c <= ColCount; c++)
                        {
                            // Here we will add all strings, since we do not know what we are waiting for.
                            Data.Columns.Add(TCellAddress.EncodeColumn(c), typeof(String));
                        }
                    }

                    string[] dr = new string[ColCount];

                    int rowCount = xls.RowCount;
                    for (int r = beginIndexRow; r <= rowCount; r++)
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

            return dataSet;
        }
    }
}