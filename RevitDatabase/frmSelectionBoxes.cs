using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace RevitDatabase
{
    public partial class frmSelectionBoxes : Form
    {
        public SqlConnection con;

        public frmSelectionBoxes()
        {
            InitializeComponent();
            con = MainForm.con;
        }

        /// <summary>
        /// Loads all the table names from the currently selected database into a ComboBox. Table names are filtered with a prefix of 'tbl'.
        /// </summary>
        /// <param name="_comboBox"></param>
        private void GetAllTablesFromDatabase(System.Windows.Forms.ComboBox _comboBox)
        {
            con.Open();
            using (SqlCommand com = new SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES", con))
            {
                using (SqlDataReader reader = com.ExecuteReader())
                {
                    _comboBox.Items.Clear();
                    while (reader.Read())
                    {
                        string tblName = "";
                        tblName = (string)reader["TABLE_NAME"];

                        if (tblName.StartsWith("tbl"))
                        {
                            _comboBox.Items.Add(tblName);
                        }
                    }
                }
            }
            con.Close();
        }

        private void frmSelectionBoxes_Load(object sender, EventArgs e)
        {
            btnOK.Enabled = false;
        }

        private void cbDatabases_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbDatabases.SelectedIndex != -1)
            {
                con.ConnectionString = "";
                MainForm.SelectedDatabase = "";
                MainForm.SelectedDatabase = cbDatabases.SelectedItem.ToString();
                con.ConnectionString = @"Server=BPRMEPCM-7\SQLEXPRESS;Database=" + MainForm.SelectedDatabase + @";Trusted_Connection=True;";

                GetAllTablesFromDatabase(cbTables);
            }

            if (cbDatabases.SelectedIndex == -1 
                && cbTables.SelectedIndex == -1)

                btnOK.Enabled = false;
            else
                btnOK.Enabled = true;

        }

        private void cbTables_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (cbTables.SelectedIndex != -1)
                MainForm.SelectedTable = cbTables.SelectedItem.ToString();

            if (cbDatabases.SelectedIndex == -1
                && cbTables.SelectedIndex == -1)

                btnOK.Enabled = false;
            else
                btnOK.Enabled = true;

        }
    }
}
