using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ALogic.DBConnector
{
    public static class EntityToTableConverter
    {
        public static DataTable ToDataTable<T>(this IList<T> data)
        {
            PropertyDescriptorCollection props =
                TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            for (int i = 0; i < props.Count; i++)
            {
                PropertyDescriptor prop = props[i];
                table.Columns.Add(prop.Name, prop.PropertyType);
            }
            object[] values = new object[props.Count];
            foreach (T item in data)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = props[i].GetValue(item);
                }
                table.Rows.Add(values);
            }
            return table;
        }

        public static DataTable ConvertToDataTable<T>(this IList<T> data)
        {
            PropertyDescriptorCollection properties =
               TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;

        }

    }



    public static class Helper
    {
        private static readonly IDictionary<Type, ICollection<PropertyInfo>> _Properties =
            new Dictionary<Type, ICollection<PropertyInfo>>();

        /// <summary>
        /// Converts a DataTable to a list with generic objects
        /// </summary>
        /// <typeparam name="T">Generic object</typeparam>
        /// <param name="table">DataTable</param>
        /// <returns>List with generic objects</returns>
        public static IEnumerable<T> DataTableToList<T>(this DataTable table) where T : class, new()
        {
            try
            {
                var objType = typeof(T);
                ICollection<PropertyInfo> properties;

                lock (_Properties)
                {
                    if (!_Properties.TryGetValue(objType, out properties))
                    {
                        properties = objType.GetProperties().Where(property => property.CanWrite).ToList();
                        _Properties.Add(objType, properties);
                    }
                }

                var list = new List<T>(table.Rows.Count);

                foreach (var row in table.AsEnumerable().Skip(0))
                {
                    var obj = new T();

                    foreach (var prop in properties)
                    {
                        try
                        {
                            var propType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                            if (prop.Name == "DbState")
                            {
                                prop.SetValue(obj, (int)row.RowState, null);
                                continue;
                            }
                            if (row.RowState == DataRowState.Deleted)
                            {
                                var safeValue = row[prop.Name, DataRowVersion.Original] == null ? null : Convert.ChangeType(row[prop.Name, DataRowVersion.Original], propType);
                                prop.SetValue(obj, safeValue, null);
                            }
                            else
                            {
                                var safeValue = row[prop.Name] == null ? null : Convert.ChangeType(row[prop.Name], propType);
                                prop.SetValue(obj, safeValue, null);
                            }

                           
                        }
                        catch
                        {
                            // ignored
                        }
                    }

                    

                    list.Add(obj);
                }

                return list;
            }
            catch
            {
                return Enumerable.Empty<T>();
            }
        }

        public static void Fill(this object obj, object data)
        {
            try
            {
                var objType = obj.GetType();
                //ICollection<PropertyInfo> properties;

                /*lock (_Properties)
                {
                    if (!_Properties.TryGetValue(objType, out properties))
                    {
                        properties = objType.GetProperties().Where(property => property.CanWrite).ToList();
                        _Properties.Add(objType, properties);
                    }
                }*/

                foreach (var prop in /*properties*/   objType.GetProperties().Where(property => property.CanWrite).ToList())
                {
                    try
                    {
                        var safeValue = prop.GetValue(data, null);
                        prop.SetValue(obj, safeValue, null);

                    }
                    catch
                    {
                        // ignored
                    }
                }
            }
            catch
            {

            }
        }

       
    }



}
