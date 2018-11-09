using System;
using System.Text;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;
using QuickBookConnector;



namespace QuickBookConnector
{
    class Program
    {
        private static OdbcConnection _cn;
        static void Main() 
        {


            String line;
            using (var file = System.IO.File.OpenText("C:\\development\\Code\\AllTables.csv"))
            {
                // read each line, ensuring not null (EOF)
                while ((line = file.ReadLine()) != null)
                {
                    // return trimmed line
                    // yield return line.Trim();
                    //_cn = new OdbcConnection(string.Format("DSN={0}", "QuickBooks Data QRemote"));
                    //_cn.ConnectionTimeout = 60;
                    //_cn.Open();
                    
                    Console.WriteLine("!!!Downloading Quickbooks data!!!!");
                    Console.WriteLine( "SizeOf IntPtr is: {0}", IntPtr.Size );
                    //Console.WriteLine("Please enter FileName");
                    //string fileName = Console.ReadLine();
                    string fileName = line;

                    //Console.WriteLine("Please enter Query");
                    //string query = Console.ReadLine();
                    string query = "select * from "+fileName;
                    Console.WriteLine("Processing for "+fileName);

                    string filePath = "C:\\development\\Code\\QuickBooks\\"+fileName+".csv";
                    if(File.Exists(filePath)){
                        Console.WriteLine("Ignoring file : "+fileName);
                        continue;
                    }

                    // Keep the console window open in debug mode.
                    //Console.WriteLine("Press any key to exit.");
                    //Console.ReadKey();
                    
                    //string query = string.Format("select *  from VendorContacts ");
                    //ProcessQuery(query);
                    DataSet dataSet = new DataSet();
                    dataSet = GetDataSetFromAdapter(dataSet,string.Format("DSN={0}", "QuickBooks Data QRemote"),query);
                    //CreateExcel(dataSet, "c:\\Demo.xls");

                    string text = ConvertToCSV(dataSet);
                    System.IO.File.WriteAllText(@"C:\\development\\Code\\QuickBooks\\"+fileName+".csv", text);

                    
                    //_cn.Close();
                    //_cn.Dispose();
                    //_cn = null;
                }
            }


        

        }
        
        
		public static DataSet GetDataSetFromAdapter(
			DataSet dataSet, string connectionString, string queryString)
		{
			using (OdbcConnection connection = 
					   new OdbcConnection(connectionString))
			{
				OdbcDataAdapter adapter = 
					new OdbcDataAdapter(queryString, connection);

				// Open the connection and fill the DataSet.
				try
				{
					connection.Open();
					adapter.Fill(dataSet);
				}
				catch (Exception ex)
				{
					Console.WriteLine(ex.Message);
				}
				// The connection is automatically closed when the
				// code exits the using block.
			}
			return dataSet;
		}        
        
        private static string ConvertToCSV(DataSet objDataSet)
        {
            StringBuilder content = new StringBuilder();

            if (objDataSet.Tables.Count >= 1)
            {
                DataTable table = objDataSet.Tables[0];

                if (table.Rows.Count > 0)
                {
                    DataRow dr1 = (DataRow) table.Rows[0];
                    int intColumnCount = dr1.Table.Columns.Count;
                    int index=1;

                    //add column names
                    foreach (DataColumn item in dr1.Table.Columns)
                    {
                        content.Append(String.Format("\"{0}\"", item.ColumnName));
                        if (index < intColumnCount)
                            content.Append(",");
                        else
                            content.Append("\r\n");
                        index++;
                    }

                    //add column data
                    foreach (DataRow currentRow in table.Rows)
                    {
                        string strRow = string.Empty;
                        for (int y = 0; y <= intColumnCount - 1; y++)
                        {
                            strRow += "\"" + ReplaceNewlines(currentRow[y].ToString(),"") + "\"";

                            if (y < intColumnCount - 1 && y >= 0)
                                strRow += ",";
                        }
                        content.Append(strRow + "\r\n");
                    }
                }
            }

            return content.ToString();
        }

        static string ReplaceNewlines(string blockOfText, string replaceWith)
        {
            return blockOfText.Replace("\r\n", replaceWith).Replace("\n", replaceWith).Replace("\r", replaceWith);
        }

		public static void CreateExcel(DataSet ds, string excelPath)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            try
            {
                //Previous code was referring to the wrong class, throwing an exception
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                {
                    for (int j = 0; j <= ds.Tables[0].Columns.Count - 1; j++)
                    {
                        xlWorkSheet.Cells[i + 1, j + 1] = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    }
                }

                xlWorkBook.SaveAs(excelPath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlApp);
                releaseObject(xlWorkBook);
                releaseObject(xlWorkSheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }        
		
		
		private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        } 
		
		private static void ProcessQuery(string query)
        {
			
            var cmd = new OdbcCommand(query, _cn);
            //OdbcDataAdapter adapter = new OdbcDataAdapter(cmd);
            //DataSet dataSet = new DataSet();
            //adapter.Fill(dataSet);
            
			
			
			OdbcDataReader reader = cmd.ExecuteReader();
			DataTable schemaTable = reader.GetSchemaTable();

			foreach (DataRow row in schemaTable.Rows)
			{
				foreach (DataColumn column in schemaTable.Columns)
				{
					Console.WriteLine(String.Format("{0} = {1}",
					column.ColumnName, row[column]));
				}
			}
			
			
			
			
			if (reader.HasRows)
			{
				while (reader.Read())
				{
					Console.WriteLine("{0}\t{1}", reader.GetString(0),
						reader.GetString(1));
				}
			}
			else
			{
				Console.WriteLine("No rows found.");
			}
			
			
			
			
			
			//DataTable dt = new DataTable();
			//dt.Load(reader);
			
			//adapter.Fill(dataSet);
			//ExcelUtility.CreateExcel(ds, "D:\\Demo.xls");
			//CreateExcel(dataSet, "c:\\Demo.xls");
			
			//adapter.close();
			reader.Close();
			
			
	

        }

        private static void DisplayInvoiceInGrid(string invoiceRefNumber)
        {

            string query = string.Format("select RefNumber,CustomerRefFullName,InvoiceLineItemRefFullName, InvoiceLineDesc, InvoiceLineRate, InvoiceLineAmount  from InvoiceLine ");
            ProcessQuery(query);
        }
    }
}
