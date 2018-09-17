using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;

namespace webApp.Utilities
{
    public static class Convertor
    {
        public static List<T> ConvertToList<T>(DataTable dt, bool propertyNameIsColumnName = true)
        {
            var properties = typeof(T).GetProperties();

            var dtEnu = dt.AsEnumerable();

            var list = new List<T>();

            foreach(var row in dtEnu)
            {
                var objT = Activator.CreateInstance<T>();
                var colInd = 0;
                foreach (var prop in properties)
                {
                    try
                    {
                        PropertyInfo propertyInfo = objT.GetType().GetProperty(prop.Name);

                        propertyInfo.SetValue(objT,
                            propertyNameIsColumnName
                                ? Convert.ChangeType(row[prop.Name], propertyInfo.PropertyType)
                                : Convert.ChangeType(row[colInd++], propertyInfo.PropertyType), null);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Convert data exception", ex);
                    }
                }
                list.Add(objT);
            }

            return list;
        }

        public static List<T> ConvertToList3<T>(DataTable dt, bool propertyNameIsColumnName = true)
        {
            var columnNames = dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName.ToLower()).ToList();
            var properties = typeof(T).GetProperties();

            var dtEnu = dt.AsEnumerable();

            var list = new List<T>();

            foreach (var row in dtEnu)
            {
                var objT = Activator.CreateInstance<T>();
                var colInd = 0;
                foreach (var prop in properties)
                {
                    try
                    {
                        PropertyInfo propertyInfo = objT.GetType().GetProperty(prop.Name);
                        propertyInfo.SetValue(objT,
                            propertyNameIsColumnName
                                ? Convert.ChangeType(row[prop.Name], propertyInfo.PropertyType)
                                : Convert.ChangeType(row[colInd++], propertyInfo.PropertyType), null);
                    }
                    catch
                    {
                        continue;
                    }

                    //try
                    //{
                    //    prop.SetValue(objT, row[colInd++]);
                    //}
                    //catch (Exception ex) { }

                    //if (columnNames.Contains(prop.Name.ToLower()))
                    //{
                    //    try
                    //    {
                    //        prop.SetValue(objT, row[prop.Name]);
                    //    }
                    //    catch (Exception ex) { }
                }
                list.Add(objT);
            }

            return list;

            //return dt.AsEnumerable().Select(row => {
            //    var objT = Activator.CreateInstance<T>();
            //    foreach (var pro in properties)
            //    {
            //        if (columnNames.Contains(pro.Name.ToLower()))
            //        {
            //            try
            //            {
            //                pro.SetValue(objT, row[pro.Name]);
            //            }
            //            catch (Exception ex) { }
            //        }
            //    }
            //    return objT;
            //}).ToList();
        }
    }
}