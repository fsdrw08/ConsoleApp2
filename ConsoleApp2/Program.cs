using ClosedXML.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"C:\Users\drw_0\Documents\kv.txt";
            
            if(Directory.Exists(FilePath))
            { 
                string[] Fileline = File.ReadAllLines(FilePath);
                Hashtable hashtable = new Hashtable();

                foreach (string line in Fileline)
                {
                    if (line.Contains("="))
                    {
                        hashtable.Add(line.Split('=')[0], line.Split('=')[1]);
                    }
                }

                Console.WriteLine(hashtable["k2"]);

                foreach (DictionaryEntry line in hashtable)
                {
                    Console.Write(line.Key + "\t:");
                    Console.WriteLine(hashtable[line.Key]);
                }
            }


            var dt = GetDataFromExcel(@"C:\Users\drw_0\OneDrive\Documents\Test\Book1.xlsx", "sheet1");

            foreach (DataRow row in dt.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    //if (!String.IsNullOrEmpty(item.ToString()))
                    Console.WriteLine(item);
                }
            }
            Console.ReadLine();
            

            var distinctIds = dt.AsEnumerable()
                              .Select(s => new
                              {
                                  id = s.Field<string>("c1"),
                              })
                              .Distinct().ToList();
            string items = string.Join(Environment.NewLine, distinctIds);
            Console.WriteLine(items);
            Console.ReadLine();
            /*
            select new
            {
                Value = groupby.Key,
                ColumnValues = groupby
            };
            */

            /*
            var groupby = dt.AsEnumerable().GroupBy ( d=> new
            {
                c1 = d.Field<string>("c1"),
                c2 = d.Field<string>("c2"),
                c3 = d.Field<string>("c3"),
                c4 = d.Field<string>("c4"),
            })
            .Select(x => new {
                c1 = x.Key.c1,
                c2 = x.Key.c2,
                c3 = x.Key.c3,
                c4 = x.Key.c4,
            });
            

            foreach (var key in grouped)

            {

                Console.WriteLine(key.Value.c1);

                Console.WriteLine("---------------------------");

                foreach (var columnValue in key.ColumnValues)

                {

                    Console.WriteLine(columnValue["c2"].ToString());
                    Console.WriteLine(columnValue["c3"].ToString());
                    Console.WriteLine(columnValue["c3"].ToString());

                }

                Console.WriteLine();

            }
            Console.ReadLine();
        */
        }


        public static DataTable GetDataFromExcel(string path, dynamic worksheet)
        {
            //Save the uploaded Excel file.


            DataTable dt = new DataTable();
            //Open the Excel file using ClosedXML.
            using (XLWorkbook workBook = new XLWorkbook(path))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = workBook.Worksheet(worksheet);

                //Create a new DataTable.

                //Loop through the Worksheet rows.
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            if (!string.IsNullOrEmpty(cell.Value.ToString()))
                            {
                                dt.Columns.Add(cell.Value.ToString());
                            }
                            else
                            {
                                break;
                            }
                        }
                        firstRow = false;
                    }
                    else
                    {
                        int i = 0;
                        DataRow toInsert = dt.NewRow();
                        foreach (IXLCell cell in row.Cells(1, dt.Columns.Count))
                        {
                           try
                           {
                             toInsert[i] = cell.Value.ToString();
                           }
                            catch (Exception ex)
                            {

                            }
                            i++;
                        }
                        dt.Rows.Add(toInsert);
                    }
                }
                return dt;
            }
        }

    }
}
