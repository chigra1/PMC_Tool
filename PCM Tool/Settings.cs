using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PMC_Tool
{
    public partial class Settings : Form
    {
        public static List<string> lines;
        public static string path = AppDomain.CurrentDomain.BaseDirectory + @"Settings\MPQ_Types.txt";
        private bool dataSourceChanged = false;

        public Settings()
        {
            InitializeComponent();
        }

        public void LoadMPQTypes()
        {
            var read = System.IO.File.ReadAllLines(path).Where(s => s.Trim() != string.Empty).ToArray();
            lines = new List<string>(read);
        }

        private void Settings_Load(object sender, EventArgs e)
        {
            LoadMPQTypes();
            listBoxMPQTypes.DataSource = lines;
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty (textBoxAdd.Text))
            {
                lines.Add(textBoxAdd.Text);
                dataSourceChanged = true;
            }
            listBoxMPQTypes.DataSource = null;
            listBoxMPQTypes.DataSource = lines;

            textBoxAdd.Clear();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            System.IO.File.WriteAllLines(path, lines);
            this.Close();
        }

        private void buttonRemove_Click(object sender, EventArgs e)
        {
            if (listBoxMPQTypes.SelectedItem != null)
            {
                for (int i = 0; i < lines.Count; i++)
                {
                    if (listBoxMPQTypes.SelectedItem.ToString() == lines[i])
                    {
                        lines.RemoveAt(i);
                        listBoxMPQTypes.DataSource = null;
                        listBoxMPQTypes.DataSource = lines;
                        dataSourceChanged = true;
                    }
                }
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            
            if (dataSourceChanged == true)
            {
                DialogResult result = MessageBox.Show("Do you want to save changes?", "Save changes?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    btnSave_Click(null, null);
                }
                else if (result == DialogResult.No)
                {
                    this.Close();
                }
                else
                {

                }
            }
            else
            {
                this.Close();
            }
   
        }
    }
}
