using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Entity;
using System.Data;
using System.Linq;
using System.Web;
using TestClientsProject.Models;

namespace TestClientsProject.CSV
{
    public class CSVExport
    {
        /// <summary>
        /// Creates a response as a CSV with a header row and results of a data table 
        /// </summary>
        /// <param name="dt">DataTable which holds the data</param>
        /// <param name="fileName">File name for the outputted file</param>
        public static void WriteDataTableToCSV(DbSet db, string fileName)
        {
            WriteOutCSVResponseHeaders(fileName);
            WriteOutDataTable(ToDataTable<Client>(db));
            HttpContext.Current.Response.End();
        }

        private static DataTable ToDataTable<T>(DbSet data)
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

        /// <summary>
        /// Writes out the response headers needed for outputting a CSV file.
        /// </summary>
        /// <param name="fileName">File name for the outputted file</param>
        private static void WriteOutCSVResponseHeaders(string fileName)
        {
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.ClearHeaders();
            HttpContext.Current.Response.ClearContent();
            HttpContext.Current.Response.AddHeader("content-disposition", string.Format("attachment; filename={0}-{1}.csv", fileName, DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss")));
            HttpContext.Current.Response.AddHeader("Pragma", "public");
            HttpContext.Current.Response.ContentType = "text/csv";
            HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.UTF8;
        }


        /// <summary>
        /// Writes out the header row and data rows from a data table.
        /// </summary>
        /// <param name="dt">DataTable which holds the data</param>
        private static void WriteOutDataTable(DataTable dt)
        {
            WriteOutHeaderRow(dt, dt.Columns.Count);
            WriteOutDataRows(dt, dt.Columns.Count, dt.Rows.Count);
        }

        /// <summary>
        /// Writes the header row from a datatable as Http Response
        /// </summary>
        /// <param name="dt">DataTable which holds the data</param>
        /// <param name="colCount">Number of columns</param>
        private static void WriteOutHeaderRow(DataTable dt, int colCount)
        {
            string CSVHeaderRow = string.Empty;
            for (int col = 0; col <= colCount - 1; col++)
            {
                CSVHeaderRow = string.Format("{0}{1},", CSVHeaderRow, dt.Columns[col].ColumnName);
            }
            WriteRow(CSVHeaderRow);
        }

        /// <summary>
        /// Writes the data rows of a datatable as Http Responses
        /// </summary>
        /// <param name="dt">DataTable which holds the data</param>
        /// <param name="colCount">Number of columns</param>
        /// <param name="rowCount">Number of columns</param>
        private static void WriteOutDataRows(DataTable dt, int colCount, int rowCount)
        {
            string CSVDataRow = string.Empty;
            for (int row = 0; row <= rowCount - 1; row++)
            {
                var dataRow = dt.Rows[row];
                CSVDataRow = string.Empty;
                for (int col = 0; col <= colCount - 1; col++)
                {
                    CSVDataRow = string.Format("{0}{1},", CSVDataRow, dataRow[col]);
                }
                WriteRow(CSVDataRow);
            }
        }

        /// <summary>
        /// Write out a row as an Http Response.
        /// </summary>
        /// <param name="row">The data row to write out</param>
        private static void WriteRow(string row)
        {
            HttpContext.Current.Response.Write(row.TrimEnd(','));
            HttpContext.Current.Response.Write(Environment.NewLine);
        }
    }
}