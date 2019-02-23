using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using PMC_Tool;
using System.Diagnostics;

namespace PCM_Tool
{
    public partial class Form1 : Form
    {
        //Excel excel;
        _Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
        string excelFile;
        string path;
        //string dbPath;
        public BOM bom;
        //public int numberOfHDTFlags = 0;
        public Masterschedule ms;
        List<object> MSproductName = new List<object>();
        List<object> MSproductQty = new List<object>();
        List<object> MSproductStatus = new List<object>();
        //List<string> lstBom;
        List<object> BOMmaterialNumber;
        List<object> BOMmaterialQty;
        List<string> BOMproductName;
        List<string> Errors = new List<string>();
        List<string> pcbaFilenames = new List<string>();
        private Thread t = null;
        public List<Assembly> assemblies = new List<Assembly>();

        public Form1()
        {
            InitializeComponent();
        }

        public string addToErrorList
        {
            set{ Errors.Add(value);
                AddErrorsToDataGridView();
            }
        }

        private void btnCalculate_Click(object sender, EventArgs e)
        {
            label5.Text = "START Date and time: " + DateTime.Now.ToString("dd/MM/yyyy h:mm:ss tt");
            btnCalculate.Enabled = false;
            Calculating c = new Calculating();
            c.Show(this);
            Settings s = new Settings();
            s.LoadMPQTypes();

            CopyBOMs();
            //CopyPCBAs();

            //assembly
            addWIPQuantityToAssembly();
            addInventoryQuantityToAssembly();

            //database
            addDatabaseToBOMList();
            
            bom = removeHDTbrandFromBOMS();//ovo ce da obrise sve komponente koje imaju HDT u U koloni u database fajlu

            addAssemblyQtyToBOMS();

            //addStockToBOMList(); //ovo je ponovo iskljuceno jer citamo stock iz database fajla
            //addScrapToStockList(); //ovo je iskljuceno 15.08.18 Berti trazio
            addInvRepToStockList(); //dodato 15.08.18 Berti trazio\ umesto stock qty-a
            addPOToStockList();
            addWIPToStockList();
            addSSToStockList();
            //addOP_POToStockList();// iskljuceno 01.2019

            //write to gui
            WriteBomDataToDataGridView();
         
            btnExportAll.Enabled = true;
            btnExportDemand.Enabled = true;
            label2.Visible = false;
            c.Close();
            System.Media.SoundPlayer player = new System.Media.SoundPlayer(AppDomain.CurrentDomain.BaseDirectory + @"Sound\audio.wav");
            player.Play();
            label6.Text = "END Date and time: " + DateTime.Now.ToString("dd/MM/yyyy h:mm:ss tt");
        }

        private BOM removeHDTbrandFromBOMS()
        {
            BOM bom_temp = new BOM(bom.item.Length- bom.numberOfHDTFlags);
            for(int i = 0,j=0; i < bom.item.Length; i++)
            {
                if(bom.item[i].HDTflag == false)
                {
                    bom_temp.item[j] = bom.item[i];
                    j++;
                }
            }
            return bom_temp;
        }
        private void addAssemblyQtyToBOMS()
        {
            for (int i = 0; i < bom.item.Length; i++)
            {
                for (int j = 0; j < assemblies.Count; j++)
                {
                    for (int z = 0; z < assemblies[j].lstItems.Count; z++)
                    {
                        if (assemblies[j].lstItems[z].material_name == bom.item[i].name)
                        {
                            if(assemblies[j].ms_quantity >= assemblies[j].inventory_quantity + assemblies[j].wip_quantity)
                                bom.item[i].assembly_qty += assemblies[j].lstItems[z].material_qty * (assemblies[j].inventory_quantity + assemblies[j].wip_quantity);
                            else
                                bom.item[i].assembly_qty += assemblies[j].lstItems[z].material_qty * assemblies[j].ms_quantity;
                        }
                    }
                }
            }
        }
        private void CopyBOMs()
        {
            path = AppDomain.CurrentDomain.BaseDirectory + @"BOMs\";
            BOMmaterialNumber = new List<object>();
            BOMmaterialQty = new List<object>();
            BOMproductName = new List<string>();
            //BOMMultiplyMS = new List<object>();

            for (int i = 0; i < ms.item.Length; i++)// ucitavanje i poredjenje sledecih bomova iz ms strukture
            {
                Excel excel = new Excel(path + ms.item[i].product + ".xls", 1);
                bom = excel.CopyBOM(ms.item[i].quantity, ms.item[i].product, pcbaFilenames);
                for (int j = 0; j < bom.item.Length; j++)
                {
                    int redniBroj = SearchIfItemExists(j);
                    if (redniBroj == -1)
                    {
                        BOMmaterialNumber.Add(bom.item[j].name);
                        BOMmaterialQty.Add(bom.item[j].bomqty);
                        BOMproductName.Add(bom.item[j].product);
                    }
                    else
                    {
                        double total = (double)BOMmaterialQty[redniBroj] + bom.item[j].bomqty;
                        BOMmaterialQty[redniBroj] = total;
                        if (BOMproductName[redniBroj].Contains(bom.item[j].product))
                        {
                            // Do Something // or not :)
                        }
                        else
                        {
                            BOMproductName[redniBroj] = BOMproductName[redniBroj] + ";" + bom.item[j].product;
                        }
                    }
                }
                excel.Dispose();
            }

            bom = new BOM(BOMmaterialNumber.Count);
            for (int i = 0; i < BOMmaterialNumber.Count; i++)// vracanje liste u bom strukturu
            {
                bom.item[i].name = BOMmaterialNumber[i].ToString();
                bom.item[i].bomqty = Convert.ToDouble(BOMmaterialQty[i]);
                bom.item[i].product = BOMproductName[i].ToString();
            }
        }

        private void CopyPCBAs()
        {
            path = AppDomain.CurrentDomain.BaseDirectory + @"PCBA\";

            for (int i = 0; i < assemblies.Count; i++)
            {
                Excel excel = new Excel(path + assemblies[i].name + ".xls", 1);
                excel.CopyPCBA(assemblies[i]);

                excel.Dispose();
            }
        }


        private int SearchIfItemExists(int redniBroj) //ako postoji item vraca se njegov redni broj u listi, ako ne postoji vraca se -1
        {
            for (int z = 0; z < BOMmaterialNumber.Count; z++)
            {
                if (bom.item[redniBroj].name == BOMmaterialNumber[z].ToString())
                return z;
            }
            return -1;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            foreach (Process excelProcess in Process.GetProcesses())
            {
                if (excelProcess.ProcessName.Equals("EXCEL"))
                {
                    excelProcess.Kill();
                    break;
                }
            }

            btnExportAll.Enabled = false;
            btnExportDemand.Enabled = false;
            progressBar1.Visible = false;
            label2.Visible = false;
            LoadMasterSchedule();
            WriteErrorsToDataGridView();
            //GetPCBAFilenames();
        }

        private int SearchIfBOMExistsInList(int redniBroj) //ako 
        {
            for (int z = 0; z < MSproductName.Count; z++)
            {
                if (ms.item[redniBroj].product == MSproductName[z].ToString())
                    return z;
            }
            return -1;

        }

        private void LoadMasterSchedule()
        {
            path = AppDomain.CurrentDomain.BaseDirectory + @"MasterSchedule\Masterschedule.xlsx";
            Excel excel = new Excel(path, 1);

            ms = excel.CopyMasterschedule();//kopiranje u strukturu MS samo onih sto imaju status "Material purchase"


            path = AppDomain.CurrentDomain.BaseDirectory + @"BOMs\";

            //proveravanje da li postoji BOM u stukturi i sabiranje istih
            for (int i = 0; i < ms.item.Length; i++)
            {
                int redniBroj = SearchIfBOMExistsInList(i);
                if (redniBroj == -1)
                {
                    MSproductName.Add(ms.item[i].product);
                    MSproductQty.Add(ms.item[i].quantity);
                    //MSproductStatus.Add(ms.item[i].status);
                }
                else
                {
                    MSproductQty[redniBroj] = ms.item[i].quantity + (int)MSproductQty[redniBroj];
                }
            }
            ms = new Masterschedule(MSproductName.Count);// vracanje liste u ms strukturu
            for (int z = 0; z < MSproductName.Count; z++)
            {
                try
                {
                    ms.item[z].product = MSproductName[z].ToString();
                    ms.item[z].quantity = Convert.ToInt32(MSproductQty[z]);
                    //ms.item[z].status = MSproductStatus[z].ToString();
                }
                catch
                { }
            }
            MSproductName = new List<object>();
            MSproductQty = new List<object>();
            MSproductStatus = new List<object>();
            /////////////////////////////////////////


            //proveravanje da li postoji BOM u folderu i ubacivanje u novu listu
            for (int i = 0; i < ms.item.Length; i++)
            {
                string fullPathBOM = path + ms.item[i].product + ".xls";
                if (File.Exists(fullPathBOM) == false)
                {
                    Errors.Add("There is no " + ms.item[i].product + ".xls in BOMs folder!");
                    ms.item[i].found = false;
                }
                else
                {
                    MSproductName.Add(ms.item[i].product);
                    MSproductQty.Add(ms.item[i].quantity);
                    MSproductStatus.Add(ms.item[i].status);
                    ms.item[i].found = true;
                }
            }

            //ispis u datagridview pronadjenih i nepronadjenih bomova
            WriteToMasterScheduleDataGridView();


            /////////////////////////////////////
            ms = new Masterschedule(MSproductName.Count);// vracanje liste u ms strukturu samo pronadjenih bomova
            for (int z = 0; z < MSproductName.Count; z++)
            {
                ms.item[z].product = MSproductName[z].ToString();
                ms.item[z].quantity = Convert.ToInt32(MSproductQty[z]);
                //ms.item[z].status = MSproductStatus[z].ToString();
                ms.item[z].found = true;
            }

            //ispis koliko ima pronadjenih bomova
            lbFoundBoms.Text = lbFoundBoms.Text + ms.item.Length.ToString();

            excel.Dispose();
        }

        public void WriteBomDataToDataGridView()
        {
            dGVmaterials.Rows.Clear();
            for (int i = 0; i < bom.item.Length; i++)
            {
                //if (bom.item[i].HDTflag == false)
                //{
                //izracunaj razliku
                bom.calcDifference(i);
                if (bom.item[i].difference < 0)
                {
                    //izracunaj final demand
                    bom.calcOrderQty(i);
                    //izracunaj ukupnu cenu
                    bom.calcTotalPrice(i);
                }

                string name = bom.item[i].name;
                string product = bom.item[i].product;

                dGVmaterials.Rows.Add();
                dGVmaterials.Rows[i].Cells["itemNumber"].Value = i + 1;
                dGVmaterials.Rows[i].Cells["itemName"].Value = name;
                dGVmaterials.Rows[i].Cells["Description"].Value = bom.item[i].description;
                dGVmaterials.Rows[i].Cells["Supplier"].Value = bom.item[i].supplier;
                dGVmaterials.Rows[i].Cells["Delivery_time"].Value = bom.item[i].delivery_time;
                dGVmaterials.Rows[i].Cells["Product"].Value = product;

                dGVmaterials.Rows[i].Cells["bomQty"].Value = bom.item[i].bomqty;
                dGVmaterials.Rows[i].Cells["safetyStock"].Value = Math.Round(bom.item[i].safetystock, 0);
                dGVmaterials.Rows[i].Cells["stockQty"].Value = bom.item[i].stockqty;
                dGVmaterials.Rows[i].Cells["WIP"].Value = Math.Round(bom.item[i].wip, 0);
                dGVmaterials.Rows[i].Cells["PO"].Value = Math.Round(bom.item[i].po, 0);
                dGVmaterials.Rows[i].Cells["OP_PO"].Value = Math.Round(bom.item[i].op_po, 0);
                dGVmaterials.Rows[i].Cells["Assembly_Quantity"].Value = Math.Round(bom.item[i].assembly_qty, 0);
                dGVmaterials.Rows[i].Cells["diff"].Value = Math.Round(bom.item[i].difference, 0);
                dGVmaterials.Rows[i].Cells["MOQ"].Value = Math.Round(bom.item[i].moq, 0);
                dGVmaterials.Rows[i].Cells["POQ"].Value = Math.Round(bom.item[i].mpq, 0);
                dGVmaterials.Rows[i].Cells["Final_demand"].Value = Math.Round(bom.item[i].calculated_order_qty, 0);
                dGVmaterials.Rows[i].Cells["Unit_Price"].Value = Math.Round(bom.item[i].unit_price, 3);
                dGVmaterials.Rows[i].Cells["Total_price"].Value = Math.Round(bom.item[i].total_price, 0);
                //dataGridView1.Rows[i].Cells["Scrap"].Value = Math.Round(bom.item[i].scrap, 0);
                //dataGridView1.Rows[i].Cells["HDT"].Value = bom.item[i].HDTflag;

                if (bom.item[i].difference < 0)
                {
                    paintDatagridview(i, "diff", Color.Crimson);
                }
                else
                {
                    paintDatagridview(i, "diff", Color.LawnGreen);
                }
            }
        }

        private void addDatabaseToBOMList()
        {
            path = AppDomain.CurrentDomain.BaseDirectory + @"SUDA\Database.xlsx";
            Excel excel = new Excel(path, 1);
            //int usedRows = excel.GetNumberOfUsedRows();
            bom = excel.getInfoFromDatabaseforBomComponents(bom, "A");
            excel.Dispose();
        }

        private void addPOToStockList()
        {
            path = AppDomain.CurrentDomain.BaseDirectory + @"SUDA\PO_Record.xlsx";
            Excel excel = new Excel(path, 1);
            //int usedRows = excel.GetNumberOfUsedRows();
            bom = excel.getQuantityFromPOforBomComponents(bom, "F");
            excel.Dispose();
        }

        private void addScrapToStockList()
        {
            path = AppDomain.CurrentDomain.BaseDirectory + @"SUDA\Scrap.xlsx";
            Excel excel = new Excel(path, 1);
            //int usedRows = excel.GetNumberOfUsedRows();
            bom = excel.getQuantityFromScrapforBomComponents(bom, "A");
            excel.Dispose();
        }

        private void addInvRepToStockList()
        {
            path = AppDomain.CurrentDomain.BaseDirectory + @"SUDA\Inventory_Report.xlsx";
            Excel excel = new Excel(path, 1);
            //int usedRows = excel.GetNumberOfUsedRows();
            bom = excel.getQuantityFromInvRepforBomComponents(bom, "A");
            excel.Dispose();
        }

        private void addWIPToStockList()
        {
            path = AppDomain.CurrentDomain.BaseDirectory + @"SUDA\WIP.xlsx";
            Excel excel = new Excel(path, 1);
            //int usedRows = excel.GetNumberOfUsedRows();
            bom = excel.getQuantityFromWIPforBomComponents(bom, "P");
            excel.Dispose();
        }

        private void addSSToStockList()
        {
            path = AppDomain.CurrentDomain.BaseDirectory + @"SUDA\SafetyStock.xlsx";
            Excel excel = new Excel(path, 1);
            //int usedRows = excel.GetNumberOfUsedRows();
            bom = excel.getQuantityFromSSforBomComponents(bom, "C");
            excel.Dispose();
        }

        private void addWIPQuantityToAssembly()
        {
            path = AppDomain.CurrentDomain.BaseDirectory + @"SUDA\WIP.xlsx";
            Excel excel = new Excel(path, 1);
            excel.getQuantityFromWIPforAssembly(assemblies, "P");
            excel.Dispose();
        }

        private void addInventoryQuantityToAssembly()
        {
            path = AppDomain.CurrentDomain.BaseDirectory + @"SUDA\Inventory_Report.xlsx";
            Excel excel = new Excel(path, 1);
            excel.getQuantityFromInventoryforAssembly(assemblies, "A");
            excel.Dispose();
        }

        private void addOP_POToStockList()
        {
            path = AppDomain.CurrentDomain.BaseDirectory + @"SUDA\Out_Processing_PO.xlsx";
            Excel excel = new Excel(path, 1);
            //int usedRows = excel.GetNumberOfUsedRows();
            bom = excel.getQuantityFromOP_POforBomComponents(bom, "D");
            excel.Dispose();
        }

        private void paintDatagridview(int rowIndex, string columnName, Color color)
        {
            dGVmaterials.Rows[rowIndex].Cells[columnName].Style.BackColor = color;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {          
            this.Close();
        }

        private void btnExportAll_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog sd = new SaveFileDialog();
                sd.Title = "Save Excel File";
                sd.Filter = "Excel Files|*.xlsx";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    excelFile = sd.FileName;
                }
                else return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            // start the animation (I used a progress bar, start your circle here)
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;

            // start the job and the timer, which polls the thread
            btnExportAll.Enabled = false;
            t = new Thread(SaveAll);
            t.Start();
            timer1.Interval = 100;
            timer1.Start();
        }

        public void WriteErrorsToDataGridView()
        {
            dGVerrors.Rows.Clear();
            for (int i = 0; i < Errors.Count; i++)
            {
                dGVerrors.Rows.Add();
                dGVerrors.Rows[i].Cells[0].Value = i + 1;
                dGVerrors.Rows[i].Cells[1].Value = Errors[i].ToString();
            }
            dGVerrors.FirstDisplayedScrollingRowIndex = dGVerrors.RowCount - 1;
        }
        public void AddErrorsToDataGridView()
        {
            dGVerrors.Rows.Add();
            dGVerrors.Rows[Errors.Count-1].Cells[0].Value = Errors.Count;
            dGVerrors.Rows[Errors.Count-1].Cells[1].Value = Errors[Errors.Count-1].ToString();

            dGVerrors.FirstDisplayedScrollingRowIndex = dGVerrors.RowCount - 1;
        }
        public void WriteToMasterScheduleDataGridView()
        {
            ////////////////////////////////////////////////////////
            for (int i = 0; i < ms.item.Length; i++)//prikazujem u datagridview
            {
                dGVboms.Rows.Add();
                dGVboms.Rows[i].Cells[0].Value = i + 1;
                dGVboms.Rows[i].Cells[1].Value = ms.item[i].product;
                dGVboms.Rows[i].Cells[2].Value = ms.item[i].quantity;
                if (ms.item[i].found == false)
                {
                    dGVboms.Rows[i].Cells[1].Style.BackColor = Color.Crimson;
                }
                else
                {
                    dGVboms.Rows[i].Cells[1].Style.BackColor = Color.LawnGreen;
                }
            }
        }

        private void btnExportDemand_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog sd = new SaveFileDialog();
                sd.Title = "Save Excel File";
                sd.Filter = "Excel Files|*.xlsx";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    excelFile = sd.FileName;
                }
                else return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            // start the animation (I used a progress bar, start your circle here)
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;

            // start the job and the timer, which polls the thread
            btnExportDemand.Enabled = false;
            t = new Thread(SaveMOD);
            t.Start();
            timer1.Interval = 100;
            timer1.Start();
        }

        private void SaveAll()
        {
            path = "";
            Excel excel = new Excel(path, 1); 
            excel.SaveWholeList(excelFile, dGVmaterials);
            excel.Dispose();

            File.Delete(AppDomain.CurrentDomain.BaseDirectory + @"template.xlsx");

            MessageBox.Show("File " + excelFile + " saved!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void SaveErrors()
        {
            path = "";
            Excel excel = new Excel(path, 1);
            excel.SaveErrorsList(excelFile, dGVerrors);
            excel.Dispose();

            File.Delete(AppDomain.CurrentDomain.BaseDirectory + @"template.xlsx");

            MessageBox.Show("File " + excelFile + " saved!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void SaveMOD()
        {
            path = "";
            Excel excel = new Excel(path, 1);
            excel.SaveMaterialOnDemand(excelFile, dGVmaterials);
            excel.Dispose();

            File.Delete(AppDomain.CurrentDomain.BaseDirectory + @"template.xlsx");

            MessageBox.Show("File " + excelFile + " saved!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            label2.Visible = true;
            if (t == null)
            {
                timer1.Stop();
                return;
            }

            // still works: exiting
            if (t.IsAlive)
                return;

            // finished
            btnExportAll.Enabled = true;
            btnExportDemand.Enabled = true;
            timer1.Stop();
            progressBar1.Visible = false;
            label2.Visible = false;
            t = null;
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void informationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About a = new About();
            a.ShowDialog();
        }

        private void btnExportErrors_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog sd = new SaveFileDialog();
                sd.Title = "Save Excel File";
                sd.Filter = "Excel Files|*.xlsx";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    excelFile = sd.FileName;
                }
                else return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            // start the animation (I used a progress bar, start your circle here)
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;

            // start the job and the timer, which polls the thread
            btnExportAll.Enabled = false;
            t = new Thread(SaveErrors);
            t.Start();
            timer1.Interval = 100;
            timer1.Start();
        }

        private void GetPCBAFilenames()
        {           
            string[] files = Directory.GetFiles(@"PCBA");
            string name = "";
            foreach (string file in files)
            {
                name = file.Substring(5);

                pcbaFilenames.Add(name.Replace(".xls", string.Empty));
            }
        }

        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Settings s = new Settings();
            s.ShowDialog();
        }
    }
}
