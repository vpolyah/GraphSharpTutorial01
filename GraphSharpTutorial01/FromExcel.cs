using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using GraphSharp.Algorithms.Layout.Simple.Tree;
using System.Data;
using System.Data.OleDb;

namespace WpfApplication1
{
    public class FromExcel
    {
        List<ObjectClass> TempGraph = new List<ObjectClass>();
        List<string> Node_list = new List<string>();
        public void InputParam(int variant)
        {           

            if (Global.tb.Rows.Count == 0)
            {
                Node_list.Clear();
                Global.Graph.Clear();
                Application excel = new Application();
                string directory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string path = System.IO.Path.Combine(directory, "111.xlsx");
                Workbook wb = excel.Workbooks.Open(path);
                Sheets excelSheets = wb.Worksheets;
                Worksheet excelWorksheet = (Worksheet)excelSheets.Item[1];

                string ConnectionString = String.Format(
                "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=\"Excel 8.0;HDR=No\";Data Source={0}", path);
                DataSet ds = new DataSet();
                OleDbConnection cn = new OleDbConnection(ConnectionString);
                cn.Open();

                
                System.Data.DataTable schemaTable =
                    cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
                            new object[] { null, null, null, "TABLE" });

                
                string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];
                
                string select = String.Format("SELECT * FROM [{0}]", sheet1);
                OleDbDataAdapter ad = new OleDbDataAdapter(select, cn);
                ad.Fill(ds);
                Global.tb = ds.Tables[0];


                string node_elem = Global.NodeName;
                Node_list.Add(node_elem);

                for (int i = 1; i < Global.tb.Rows.Count; i++)
                {
                    TempGraph.Add(new ObjectClass {Parent = Global.tb.Rows[i][1].ToString(), Name = Global.tb.Rows[i][2].ToString(), Value = Global.tb.Rows[i][0].ToString() });
                }
                    //for (int i = 2; i < excelWorksheet.UsedRange.Rows.Count; i++)
                    //{
                    //    Microsoft.Office.Interop.Excel.Range name = excelWorksheet.Cells[i, 3];
                    //    Microsoft.Office.Interop.Excel.Range parent = excelWorksheet.Cells[i, 2];
                    //    Microsoft.Office.Interop.Excel.Range value = excelWorksheet.Cells[i, 1];
                    //    TempGraph.Add(new ObjectClass { Name = name.Value.ToString(), Parent = parent.Value.ToString(), Value = value.Value.ToString() });
                    //}

                wb.Close();
                excel.Quit();
            }
            else
            {
                Node_list.Clear();
                Global.Graph.Clear();

                string node_elem = Global.NodeName;
                Node_list.Add(node_elem);

                for (int i = 1; i < Global.tb.Rows.Count; i++)
                {
                    TempGraph.Add(new ObjectClass { Name = Global.tb.Rows[i][2].ToString(), Parent = Global.tb.Rows[i][1].ToString(), Value = Global.tb.Rows[i][0].ToString() });
                }
            }
            for (int i=0;i<TempGraph.Count; i++)
            {            
                for (int j = 0; j < Node_list.Count; j++)
                {
                    if (variant==1)
                    {
                        if (Node_list[j] == TempGraph[i].Parent)
                        {
                            Node_list.Add(TempGraph[i].Name);
                            Global.Graph.Add(new ObjectClass { Name = TempGraph[i].Name, Parent = TempGraph[i].Parent, Value = TempGraph[i].Value});
                            TempGraph.Remove(TempGraph[i]);
                            i = 0;
                            break;
                        }
                    }
                    if (variant == 2)
                    {
                        if (Node_list[j] == TempGraph[i].Name)
                        {
                            Node_list.Add(TempGraph[i].Parent);
                            Global.Graph.Add(new ObjectClass { Name = TempGraph[i].Name, Parent = TempGraph[i].Parent, Value=TempGraph[i].Value });
                            TempGraph.Remove(TempGraph[i]);
                            i = 0;
                            break;
                        }
                    }
                }
            }

        }
    }

    class ObjectClass
    {
        public string Name { get; set; }
        public string Parent { get; set; }

        public string Value { get; set; }
    }
    class Global
    {
        static public System.Data.DataTable tb = new System.Data.DataTable();
        static public string NodeName { get; set; }
        static public List<ObjectClass> Graph = new List<ObjectClass>();
    }


}
