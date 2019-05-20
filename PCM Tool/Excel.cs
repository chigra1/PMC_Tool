using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace PCM_Tool
{
    class Excel
    {
        //string path = "";
        int header_row = 3;
        int first_data_row = 4;
        _Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
        Workbooks wbs;
        Workbook wb;
        Worksheet ws;
       

        public Excel(string path, int Sheet)
        {
            if (String.IsNullOrWhiteSpace(path))
            {
                excel = new _Excel.Application();
                wbs = excel.Workbooks;
                wb = null;
                ws = null;
                wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                // ExcelWorkBook.Worksheets.Add(); //Adding New Sheet in Excel Workbook
                ws = wb.Worksheets[1]; // Compulsory Line in which sheet you want to write data
                wb.Worksheets[1].Name = "Sheet";//Renaming the Sheet1 to Sheet

                if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + @"template.xlsx"))
                {
                    File.Delete(AppDomain.CurrentDomain.BaseDirectory + @"template.xlsx");
                    wb.SaveAs(AppDomain.CurrentDomain.BaseDirectory + @"template.xlsx");
                }
                else
                {
                    wb.SaveAs(AppDomain.CurrentDomain.BaseDirectory + @"template.xlsx");
                }
            }
            else
            {
                //this.path = path;
                wbs = excel.Workbooks;
                wb = wbs.Open(path);
                ws = wb.Worksheets[Sheet];
            }
        }

        public string ReadCell(int i, int j)
        {
            //i++; Interop.Excel library starts counting cells and columns from zero, so this should be used for incrementing values
            //j++;

            if (ws.Cells[i, j].Value2 != null)
                return ws.Cells[i, j].Value2.ToString();
            else
                return "";
        }

        public void WriteToCell(int i, int j, string s)
        {
            ws.Cells[i, j].Value2 = s;
        }

        public void Save()
        {
            wb.Save();
        }

        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }

        public int GetNumberOfUsedRows()
        {
            int number_of_rows = ws.UsedRange.Rows.Count;
            return number_of_rows;
        }

        public BOM getInfoFromDatabaseforBomComponents(BOM bom, string column)
        {

            column = column + ":" + column;

            for (int i = 0; i < bom.item.Length; i++)
            {
                string item = bom.item[i].name;

                _Excel.Range colRange = ws.Columns[column];//get the range object where you want to search from

                _Excel.Range resultRange = colRange.Find(

                What: item,

                LookIn: _Excel.XlFindLookIn.xlValues,

                LookAt: _Excel.XlLookAt.xlWhole,

                SearchOrder: _Excel.XlSearchOrder.xlByRows,

                SearchDirection: _Excel.XlSearchDirection.xlNext

                );

                if (resultRange == null)
                {
                    //MessageBox.Show("Did not found " + item + " in column " + column.Substring(1));
                    Program.mainForm.addToErrorList = "Did not found " + item + " in Database file - column " + column.Substring(1);
                }

                else
                {
                    int row = resultRange.Row;

                    //ovo su podaci za taj item

                    if (ws.Cells[row, 28].Value2 == "HDT")
                        {
                            bom.item[i].HDTflag = true;
                            bom.numberOfHDTFlags++;
                        }

                    try//ako nije upisan mpq u stocku da bi preskocio null
                        {
                            bom.item[i].mpq = ws.Cells[row, 13].Value2;
                            if (bom.item[i].mpq == 0)
                                Program.mainForm.addToErrorList = "POQ is 0 for item: " + bom.item[i].name;
                        }
                    catch
                        {
                            Program.mainForm.addToErrorList = "POQ is not available for item: " + bom.item[i].name;
                        }

                    try//ako nije upisan moq u stocku da bi preskocio null
                        {
                            bom.item[i].moq = ws.Cells[row, 14].Value2;
                            if (bom.item[i].moq == 0)
                                Program.mainForm.addToErrorList = "MOQ is 0 for item: " + bom.item[i].name;
                        }
                    catch
                        {
                            Program.mainForm.addToErrorList = "MOQ is not available for item: " + bom.item[i].name;
                        }

                    try//ako nije upisan unit_price u stocku da bi preskocio null
                        {
                            bom.item[i].unit_price = ws.Cells[row, 17].Value2;
                            if (bom.item[i].unit_price == 0)
                                Program.mainForm.addToErrorList = "Unit price is 0 for item: " + bom.item[i].name;
                        }
                    catch
                        {
                            Program.mainForm.addToErrorList = "Unit price is not available for item: " + bom.item[i].name;
                        }

                    try//ako nije upisan Delivery Time u stocku da bi preskocio null
                        {
                            bom.item[i].delivery_time = (int)ws.Cells[row, 12].Value2;
                        }
                    catch
                        {
                            Program.mainForm.addToErrorList = "Delivery Time is not available for item: " + bom.item[i].name;
                        }

                    try//ako nije upisan description u stocku da bi preskocio null
                        {
                            bom.item[i].description = ws.Cells[row, 3].Value2.ToString();
                        }
                    catch
                        {
                            Program.mainForm.addToErrorList = "Description is not available for item: " + bom.item[i].name;
                        }
                    try//ako nije upisan chineseDescription da bi preskocio null
                        {
                            bom.item[i].chineseDescription = ws.Cells[row, 2].Value2.ToString();
                        }
                    catch
                        {
                            Program.mainForm.addToErrorList = "Chinese Description is not available for item: " + bom.item[i].name;
                        }
                    try//ako nije upisan dobavljac u stocku da bi preskocio null
                        {
                            bom.item[i].supplier = ws.Cells[row, 11].Value2.ToString();
                        }
                    catch
                        {
                            Program.mainForm.addToErrorList = "Supplier is not available for item: " + bom.item[i].name;
                        }
                }
            }

            return bom;
        }

        public BOM getQuantityFromStockforBomComponents(BOM bom, string column)
        {

            column = column + ":" + column;

            for (int i = 0; i < bom.item.Length; i++)
            {
                string item = bom.item[i].name;

                _Excel.Range colRange = ws.Columns[column];//get the range object where you want to search from

                _Excel.Range resultRange = colRange.Find(

                What: item,

                LookIn: _Excel.XlFindLookIn.xlValues,

                LookAt: _Excel.XlLookAt.xlWhole,

                SearchOrder: _Excel.XlSearchOrder.xlByRows,

                SearchDirection: _Excel.XlSearchDirection.xlNext

                );

                if (resultRange == null)
                {
                    //MessageBox.Show("Did not found " + item + " in column " + column.Substring(1));
                    Program.mainForm.addToErrorList = "Did not found " + item + " in Stock file - column " + column.Substring(1);
                }

                else
                {
                    int row = resultRange.Row;

                    //oko je koliko ima na stoku
                    try//ako nije upisan stockqty u stocku da bi preskocio null
                    {
                        double intStock = ws.Cells[row, 16].Value2;// P kolona
                        bom.item[i].stockqty = intStock;
                    }
                    catch
                    {
                        Program.mainForm.addToErrorList = "Stock qty is not available for item: " + bom.item[i].name;
                    }


                }
            }

            return bom;
        }

        public BOM getQuantityFromPOforBomComponents(BOM bom, string column)
        {
            //int number_of_rows = ws.UsedRange.Rows.Count;
            column = column + ":" + column;

            for (int i = 0; i < bom.item.Length; i++)
            {
                string item = bom.item[i].name;

                _Excel.Range colRange = ws.Columns[column];//get the range object where you want to search from

                _Excel.Range resultRange = colRange.Find(

                What: item,

                LookIn: _Excel.XlFindLookIn.xlValues,

                LookAt: _Excel.XlLookAt.xlWhole,

                SearchOrder: _Excel.XlSearchOrder.xlByRows,

                SearchDirection: _Excel.XlSearchDirection.xlNext

                );

                if (resultRange == null)
                {
                    //Program.mainForm.addToErrorList = "Did not found " + item + " in PO - column " + column.Substring(1);
                }
                else
                {
                    _Excel.Range firstRange = resultRange;

                    while (resultRange != null)
                    {
                        bom.item[i].po += ws.Cells[resultRange.Row, 15].Value2;

                        _Excel.Range temp = resultRange;
                        resultRange = colRange.FindNext(temp);

                        if (resultRange.Address == firstRange.Address)
                            resultRange = null;
                    }
                }

            }

            return bom;
        }

        public BOM getQuantityFromScrapforBomComponents(BOM bom, string column)
        {
            column = column + ":" + column;

            for (int i = 0; i < bom.item.Length; i++)
            {
                string item = bom.item[i].name;

                _Excel.Range colRange = ws.Columns[column];//get the range object where you want to search from

                _Excel.Range resultRange = colRange.Find(

                What: item,

                LookIn: _Excel.XlFindLookIn.xlValues,

                LookAt: _Excel.XlLookAt.xlWhole,

                SearchOrder: _Excel.XlSearchOrder.xlByRows,

                SearchDirection: _Excel.XlSearchDirection.xlNext

                );

                if (resultRange == null)
                {
                    //Program.mainForm.addToErrorList = "Did not found " + item + " in PO - column " + column.Substring(1);
                }
                else
                {
                    _Excel.Range firstRange = resultRange;

                    while (resultRange != null)
                    {
                        bom.item[i].scrap += ws.Cells[resultRange.Row, 16].Value2; //p kolona

                        _Excel.Range temp = resultRange;
                        resultRange = colRange.FindNext(temp);

                        if (resultRange.Address == firstRange.Address)
                            resultRange = null;
                    }
                }

            }

            return bom;
        }

        public BOM getQuantityFromInvRepforBomComponents(BOM bom, string column)
        {

            //int number_of_rows = ws.UsedRange.Rows.Count;
            column = column + ":" + column;

            int[] columns = {12,15,18,33,36}; //brojevi kolona koje treba da saberemo

            for (int i = 0; i < bom.item.Length; i++)
            {
                string item = bom.item[i].name;

                _Excel.Range colRange = ws.Columns[column];//get the range object where you want to search from

                _Excel.Range resultRange = colRange.Find(

                What: item,

                LookIn: _Excel.XlFindLookIn.xlValues,

                LookAt: _Excel.XlLookAt.xlWhole,

                SearchOrder: _Excel.XlSearchOrder.xlByRows,

                SearchDirection: _Excel.XlSearchDirection.xlNext

                );

                if (resultRange == null)
                {
                    //Program.mainForm.addToErrorList = "Did not found " + item + " in PO - column " + column.Substring(1);
                }
                else
                {
                    _Excel.Range firstRange = resultRange;

                    while (resultRange != null)
                    {
                        for (int j = 0; j < columns.Length; j++)//columns[j]////l,o,r,u,ag,aj kolona = definisani brojevi kolona u nizu u vrhu funkcije
                        {
                            if(ws.Cells[resultRange.Row, columns[j]].Value2 != null)//sabiramo ako nisu null
                            bom.item[i].stockqty += ws.Cells[resultRange.Row, columns[j]].Value2;
                        }
                        
                        _Excel.Range temp = resultRange;
                        resultRange = colRange.FindNext(temp);

                        if (resultRange.Address == firstRange.Address)
                            resultRange = null;
                    }
                }

            }

            return bom;
        }

        public BOM getQuantityFromWIPforBomComponents(BOM bom, string column)
        {
            column = column + ":" + column;

            for (int i = 0; i < bom.item.Length; i++)
            {
                string item = bom.item[i].name;

                _Excel.Range colRange = ws.Columns[column];//get the range object where you want to search from

                _Excel.Range resultRange = colRange.Find(

                What: item,

                LookIn: _Excel.XlFindLookIn.xlValues,

                LookAt: _Excel.XlLookAt.xlWhole,

                SearchOrder: _Excel.XlSearchOrder.xlByRows,

                SearchDirection: _Excel.XlSearchDirection.xlNext

                );

                if (resultRange == null)
                {
                    ///Program.mainForm.addToErrorList = "Did not found " + item + " in WIP - column " + column.Substring(1);
                }
                else
                {
                    _Excel.Range firstRange = resultRange;

                    while (resultRange != null)
                    {                                                                                       
                        bom.item[i].wip += ws.Cells[resultRange.Row, 20].Value2;

                        _Excel.Range temp = resultRange;
                        resultRange = colRange.FindNext(temp);

                        if (resultRange.Address == firstRange.Address)
                            resultRange = null;
                    }
                }
            }
            return bom;
        }

        public void getQuantityFromWIPforAssembly(List<Assembly> assembly, string column)
        {
            column = column + ":" + column;

            for (int i = 0; i < assembly.Count; i++)
            {
                string item = assembly[i].name;

                _Excel.Range colRange = ws.Columns[column];//get the range object where you want to search from

                _Excel.Range resultRange = colRange.Find(

                What: item,

                LookIn: _Excel.XlFindLookIn.xlValues,

                LookAt: _Excel.XlLookAt.xlWhole,

                SearchOrder: _Excel.XlSearchOrder.xlByRows,

                SearchDirection: _Excel.XlSearchDirection.xlNext

                );

                if (resultRange == null)
                {
                    ///Program.mainForm.addToErrorList = "Did not found " + item + " in WIP - column " + column.Substring(1);
                }
                else
                {
                    _Excel.Range firstRange = resultRange;

                    while (resultRange != null)
                    {
                        assembly[i].wip_quantity += ws.Cells[resultRange.Row, 20].Value2;

                        _Excel.Range temp = resultRange;
                        resultRange = colRange.FindNext(temp);

                        if (resultRange.Address == firstRange.Address)
                            resultRange = null;
                    }
                }
            }
        }

        public void getQuantityFromInventoryforAssembly(List<Assembly> assembly, string column)
        {
            column = column + ":" + column;

            for (int i = 0; i < assembly.Count; i++)
            {
                string item = assembly[i].name;

                _Excel.Range colRange = ws.Columns[column];//get the range object where you want to search from

                _Excel.Range resultRange = colRange.Find(

                What: item,

                LookIn: _Excel.XlFindLookIn.xlValues,

                LookAt: _Excel.XlLookAt.xlWhole,

                SearchOrder: _Excel.XlSearchOrder.xlByRows,

                SearchDirection: _Excel.XlSearchDirection.xlNext

                );

                if (resultRange == null)
                {
                    ///Program.mainForm.addToErrorList = "Did not found " + item + " in WIP - column " + column.Substring(1);
                }
                else
                {
                    _Excel.Range firstRange = resultRange;

                    while (resultRange != null)
                    {
                        try
                        {
                            assembly[i].inventory_quantity += ws.Cells[resultRange.Row, 12].Value2;
                        }

                        catch 
                        {
                            
                        }

                        _Excel.Range temp = resultRange;
                            resultRange = colRange.FindNext(temp);

                            if (resultRange.Address == firstRange.Address)
                                resultRange = null;

                    }
                }
            }
        }
        public void getQuantityFromOP_POforAssembly(List<Assembly> assembly, string column)
        {
            column = column + ":" + column;

            for (int i = 0; i < assembly.Count; i++)
            {
                string item = assembly[i].name;

                _Excel.Range colRange = ws.Columns[column];//get the range object where you want to search from

                _Excel.Range resultRange = colRange.Find(

                What: item,

                LookIn: _Excel.XlFindLookIn.xlValues,

                LookAt: _Excel.XlLookAt.xlWhole,

                SearchOrder: _Excel.XlSearchOrder.xlByRows,

                SearchDirection: _Excel.XlSearchDirection.xlNext

                );

                if (resultRange == null)
                {
                    ///Program.mainForm.addToErrorList = "Did not found " + item + " in WIP - column " + column.Substring(1);
                }
                else
                {
                    _Excel.Range firstRange = resultRange;

                    while (resultRange != null)
                    {
                        try
                        {
                            assembly[i].opPO_quantity += ws.Cells[resultRange.Row, 11].Value2;
                        }

                        catch
                        {

                        }

                        _Excel.Range temp = resultRange;
                        resultRange = colRange.FindNext(temp);

                        if (resultRange.Address == firstRange.Address)
                            resultRange = null;

                    }
                }
            }
        }

        public BOM getQuantityFromSSforBomComponents(BOM bom, string column)
        {
            column = column + ":" + column;

            for (int i = 0; i < bom.item.Length; i++)
            {
                string item = bom.item[i].name;

                _Excel.Range colRange = ws.Columns[column];//get the range object where you want to search from

                _Excel.Range resultRange = colRange.Find(

                What: item,

                LookIn: _Excel.XlFindLookIn.xlValues,

                LookAt: _Excel.XlLookAt.xlWhole,

                SearchOrder: _Excel.XlSearchOrder.xlByRows,

                SearchDirection: _Excel.XlSearchDirection.xlNext

                );

                if (resultRange == null)
                {
                    //Program.mainForm.addToErrorList = "Did not found " + item + " in SafetyStock - column " + column.Substring(1);
                }
                else
                {
                    _Excel.Range firstRange = resultRange;

                    while (resultRange != null)
                    {
                        if (ws.Cells[resultRange.Row, 9].Value2 == null)
                        {

                        }
                        else
                        {
                            bom.item[i].safetystock += ws.Cells[resultRange.Row, 9].Value2;
                        }
                        _Excel.Range temp = resultRange;
                        resultRange = colRange.FindNext(temp);

                        if (resultRange.Address == firstRange.Address)
                            resultRange = null;
                    }
                }
            }

            return bom;
        }

        public BOM getQuantityFromOP_POforBomComponents(BOM bom, string column)
        {
            column = column + ":" + column;

            for (int i = 0; i < bom.item.Length; i++)
            {
                string item = bom.item[i].name;

                _Excel.Range colRange = ws.Columns[column];//get the range object where you want to search from

                _Excel.Range resultRange = colRange.Find(

                What: item,

                LookIn: _Excel.XlFindLookIn.xlValues,

                LookAt: _Excel.XlLookAt.xlWhole,

                SearchOrder: _Excel.XlSearchOrder.xlByRows,

                SearchDirection: _Excel.XlSearchDirection.xlNext

                );

                if (resultRange == null)
                {
                    ///Program.mainForm.addToErrorList = "Did not found " + item + " in WIP - column " + column.Substring(1);
                }
                else
                {
                    _Excel.Range firstRange = resultRange;

                    while (resultRange != null)
                    {
                        bom.item[i].op_po += ws.Cells[resultRange.Row, 11].Value2;

                        _Excel.Range temp = resultRange;
                        resultRange = colRange.FindNext(temp);

                        if (resultRange.Address == firstRange.Address)
                            resultRange = null;
                    }
                }
            }
            return bom;
        }

        public BOM CopyBOM(int MSmulti,string MSproduct, List<string> pcbaFilenames)
        {
            List<Assembly> assemblies = Program.mainForm.assemblies;
            
            int assembly_num = 0;
            //bool assembly = false;
            //MessageBox.Show( ws.UsedRange.Rows.Count.ToString());
            int number_of_rows = ws.UsedRange.Rows.Count;
            //MessageBox.Show(number_of_rows.ToString());
            //ukupan broj redova - 3
            int number_of_items = number_of_rows - 3;

            BOM bom = new BOM(number_of_items);
            for (int i = 4, z = 0; i <= number_of_rows; i++, z++)
            {
                bom.item[z].name = ws.Cells[i, 2].Value2.ToString();
                bom.item[z].product = MSproduct;
                try//ako nije upisana kolicina u bomum da bi preskocio null
                {
                    bom.item[z].bomqty = Math.Round(ws.Cells[i, 6].Value2 * MSmulti, 1);
                }
                catch
                {
                    //MessageBox.Show("Component " + bom.item[z].name + " value in BOM is null!");
                    Program.mainForm.addToErrorList = bom.item[z].name + " value in BOM " + MSproduct + " is null!";
                }

                //for (int j = 0; j < pcbaFilenames.Count; j++)
                //{
                //    if (pcbaFilenames[j] == bom.item[z].name)
                //    {
                //        assemblies.Add(new Assembly(bom.item[z].name));
                //        assembly_num = assemblies.Count - 1;
                //        assemblies[assembly_num].ms_quantity = MSmulti;
                //    }
                //}


                if ( i <= number_of_rows && i > 4)
                {
                    
                    int level = Convert.ToInt32(ws.Cells[i, 1].Value2.Replace(".", string.Empty));
                    int next_row_level = 0;
                    if (i < number_of_rows)
                        next_row_level = Convert.ToInt32(ws.Cells[i + 1, 1].Value2.Replace(".", string.Empty));


                    if (next_row_level > level)
                    {
                        int redni_br = SearchIfAssemblyExists(bom.item[z].name);
                        if (redni_br == -1)
                        {
                            assemblies.Add(new Assembly(bom.item[z].name));
                            assembly_num = assemblies.Count - 1;
                            //assembly = true;
                            assemblies[assembly_num].ms_quantity = MSmulti;
                            assemblies[assembly_num].row = i;
                            assemblies[assembly_num].level = level;
                            int sledeci_item = 1;
                            
                            while (assemblies[assembly_num].level < Convert.ToInt32(ws.Cells[i + sledeci_item, 1].Value2.Replace(".", string.Empty)))
                            {
                                if (Convert.ToInt32(ws.Cells[i + sledeci_item, 1].Value2.Replace(".", string.Empty)) == assemblies[assembly_num].level + 1)
                                {
                                    assemblies[assembly_num].AddItem(ws.Cells[i + sledeci_item, 2].Value2, ws.Cells[i + sledeci_item, 6].Value2);
                                }
                                    sledeci_item++;
                                if (i + sledeci_item > number_of_rows)
                                    break;
                            }
                        }
                        else
                        {
                            assemblies[redni_br].ms_quantity += MSmulti;
                        }
                    }

                    //else if (next_row_level < level && assembly == true)
                    //{
                    //    assemblies[assembly_num].AddItem(bom.item[z].name, ws.Cells[i, 6].Value2);
                    //    assembly = false;
                    //}
                    //else
                    //{
                    //    if (assembly == true)
                    //    {
                    //        assemblies[assembly_num].AddItem(bom.item[z].name, ws.Cells[i, 6].Value2);
                    //    }

                    //}
                }
            }
            return bom;
        }

        public void CopyPCBA(Assembly assembly)
        {
            int number_of_rows = ws.UsedRange.Rows.Count;
            //MessageBox.Show(number_of_rows.ToString());
            //ukupan broj redova - 3
            int number_of_items = number_of_rows - 4;
            //Assembly assembly = new Assembly(assemblyName);

            for (int i = 5, z = 0; i <= number_of_rows; i++, z++)
            {
                Assembly.Item item = new Assembly.Item();
                item.material_name = ws.Cells[i, 2].Value2;
                item.material_qty = ws.Cells[i, 6].Value2;

                assembly.lstItems.Add(item);
            }
        }



                private int SearchIfAssemblyExists(string assemb_name) //ako postoji item vraca se njegov redni broj u listi, ako ne postoji vraca se -1
        {
            List<Assembly> assemblies = Program.mainForm.assemblies;
            for (int z = 0; z < assemblies.Count; z++)
            {
                if (assemblies[z].name == assemb_name)
                    return z;
            }
            return -1;

        }
        public Masterschedule CopyMasterschedule()
        {

            int number_of_rows = ws.UsedRange.Rows.Count;

            int number_of_items = number_of_rows - 8;
            

            Masterschedule mstemp = new Masterschedule(number_of_items);//init strukture, pre je bilo numberOfMP

            //ponovo vrtimo da upisemo samo "Material purchase"
            for (int i = 0, z=0, j = 9; i < number_of_items; i++,j++)
            {
                try 
                {
                    //if (ws.Cells[j, 14].Value2.ToString() == "Material purchase")
                    //{
                    mstemp.item[z].product = ws.Cells[j, 7].Value.ToString();
                    //mstemp.item[z].status = ws.Cells[j, 14].Value2.ToString();
                    mstemp.item[z].quantity = Convert.ToInt32(ws.Cells[j, 10].Value);
                    z++;
                    //}
                }
                catch
                {
                    MessageBox.Show("For item in row " + j.ToString() + " there is a null value for product name or quantity", "Information",MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            return mstemp;
        }



        public void SaveMaterialOnDemand(string path, DataGridView dgv)
        {
            int numberOfColumns = dgv.Columns.Count;
            wb = excel.Workbooks.Add(1);
            ws = wb.Worksheets.get_Item(1);

            ws.Cells[1, 1].Value2 = "Date and time: " + DateTime.Now.ToString("dd/MM/yyyy h:mm:ss tt");
            for (int z = 0; z < numberOfColumns; z++)
            {
                ws.Cells[header_row, z+1].Value2 = dgv.Columns[z].HeaderText;
            }


            for (int i = 0, j = first_data_row; i < dgv.RowCount; i++)
            {
                if (dgv.Rows[i].Cells["diff"].Style.BackColor == Color.Crimson)
                {
                    try
                    {
                        for (int z=0;z<numberOfColumns;z++)
                        {
                            ws.Cells[j, z+1].Value2 = dgv.Rows[i].Cells[z].Value;
                        }
                        j++;

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }

            ws.Columns.AutoFit();
            _Excel.Range header = ws.Range[ws.Cells[header_row, 1], ws.Cells[header_row, numberOfColumns]];
            header.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Beige);
            _Excel.Range modList = ws.Range[ws.Cells[header_row, 1], ws.Cells[ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count]];
            modList.Borders.LineStyle = _Excel.XlLineStyle.xlContinuous;
            modList.Borders.Weight = _Excel.XlBorderWeight.xlThin;

            try
            {
                wb.SaveAs(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void SaveWholeList(string path, DataGridView dgv)
        {
            //var array = new object[dgv.RowCount, dgv.ColumnCount];
            //foreach (DataGridViewRow i in dgv.Rows)
            //{
            //    if (i.IsNewRow) continue;
            //    foreach (DataGridViewCell j in i.Cells)
            //    {
            //        array[j.RowIndex, j.ColumnIndex] = j.Value;
            //    }
            //}
            int numberOfColumns = dgv.Columns.Count;
            wb = excel.Workbooks.Add(1);
            ws = wb.Worksheets.get_Item(1);
            
            try
            {
                ws.Cells[1, 1].Value2 = "Date and time: " + DateTime.Now.ToString("dd/MM/yyyy h:mm:ss tt");
                for (int z = 0; z < numberOfColumns; z++)
                {
                    ws.Cells[header_row, z + 1].Value2 = dgv.Columns[z].HeaderText;
                }

                for (int i = 0, j = first_data_row; i < dgv.RowCount; i++, j++)
                {
                    for (int z = 0; z < numberOfColumns; z++)
                    {
                        ws.Cells[j, z + 1].Value2 = dgv.Rows[i].Cells[z].Value;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            ws.Columns.AutoFit();
            _Excel.Range header = ws.Range[ws.Cells[header_row, 1], ws.Cells[3, numberOfColumns]];
            header.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Beige);
            _Excel.Range wholeList = ws.Range[ws.Cells[header_row, 1], ws.Cells[ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count]];
            wholeList.Borders.LineStyle = _Excel.XlLineStyle.xlContinuous;
            wholeList.Borders.Weight = _Excel.XlBorderWeight.xlThin;

            try
            {
                wb.SaveAs(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void SaveErrorsList(string path, DataGridView dgv)
        {
            wb = excel.Workbooks.Add(1);
            ws = wb.Worksheets.get_Item(1);

            ws.Cells[1, 1].Value2 = "Date and time: " + DateTime.Now.ToString("dd/MM/yyyy h:mm:ss tt");
            ws.Cells[3, 1].Value2 = dgv.Columns[0].HeaderText;
            ws.Cells[3, 2].Value2 = dgv.Columns[1].HeaderText;


            for (int i = 4; i < dgv.RowCount + 4; i++)
            {
                try
                {
                    ws.Cells[i, 1].Value2 = dgv.Rows[i - 4].Cells[0].Value;
                    ws.Cells[i, 2].Value2 = dgv.Rows[i - 4].Cells[1].Value;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            ws.Columns.AutoFit();
            _Excel.Range header = ws.Range[ws.Cells[3, 1], ws.Cells[3, 2]];
            header.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Beige);
            _Excel.Range wholeList = ws.Range[ws.Cells[3, 1], ws.Cells[ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count]];
            wholeList.Borders.LineStyle = _Excel.XlLineStyle.xlContinuous;
            wholeList.Borders.Weight = _Excel.XlBorderWeight.xlThin;

            try
            {
                wb.SaveAs(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Dispose()
        {
            // Cleanup
            excel.DisplayAlerts = false;
            wb.Close(SaveChanges: false);
            wbs.Close();
            excel.Quit();

            // Manual disposal because of COM
            while (Marshal.ReleaseComObject(excel) != 0) { }
            while (Marshal.ReleaseComObject(wbs) != 0) { }
            while (Marshal.ReleaseComObject(wb) != 0) { }
            while (Marshal.ReleaseComObject(ws) != 0) { }
            excel = null;
            wbs = null;
            wb = null;
            wbs = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
