using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.InteropServices;

namespace RevitDatabase
{
    public partial class MainForm : System.Windows.Forms.Form
    {
        #region Class Level Variables

        //Revit
        private UIApplication uIApp;
        private Document doc;

        //System.Data
        private DataSet ds = new DataSet();
        private DataTable dt = new DataTable();
        private SqlDataAdapter da;
        public static SqlConnection con = new SqlConnection();

        //System
        private List<string> markValues = new List<string>();
        public static string SelectedDatabase = "";
        public static string SelectedTable = "";
        private string CopiedValue;
        protected StatusBar mainStatusBar = new StatusBar();
        protected StatusBarPanel databasePanel = new StatusBarPanel();
        protected StatusBarPanel tablePanel = new StatusBarPanel();
        private bool tableIsBeingFiltered = false;

        #endregion

        public MainForm()
        {
            InitializeComponent();
        }

        public MainForm(UIApplication incomingUIApp)
        {            
            InitializeComponent();
            uIApp = incomingUIApp;
            doc = uIApp.ActiveUIDocument.Document;

            con.ConnectionString = "";
            
            CreateMainMenu();
            CreateStatusBar();

            tableIsBeingFiltered = false;
            SetMenuItemChecked("filter", false);

            SetFilterToolsEnabled(false);

            GetAllDatabases(cbProjects);
        }

        #region Voids
        
        private void SetFilterToolsEnabled(bool _enabled)
        {
            if (_enabled)
            {
                txtFilter.Text = "";
                txtFilter.Enabled = true;
                btnFilter.Enabled = true;
                btnClear.Enabled = true;
            }
            else
            {
                txtFilter.Text = "";
                txtFilter.Enabled = false;
                btnFilter.Enabled = false;
                btnClear.Enabled = false;
            }
        }

        private void GetAllDatabases(System.Windows.Forms.ComboBox _comboBox)
        {
            con.ConnectionString = @"Server=BPRMEPCM-7\SQLEXPRESS;Trusted_Connection=True;";
            con.Open();
            using (SqlCommand com = new SqlCommand("SELECT name from sys.databases", con))
            {
                using (SqlDataReader reader = com.ExecuteReader())
                {
                    _comboBox.Items.Clear();
                    while (reader.Read())
                    {
                        string dbName = "";
                        dbName = (string)reader[0];

                        if (dbName.StartsWith("db"))
                            _comboBox.Items.Add(dbName);
                    }
                }
            }
            con.Close();
        }

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
                            _comboBox.Items.Add(tblName);
                    }
                }
            }
            con.Close();
        }

        private void SetMenuItemEnabled(string _name, bool _enabled)
        {
            foreach (MenuItem mainMenuItem in Menu.MenuItems)
            {
                foreach (MenuItem subMenuItem in mainMenuItem.MenuItems)
                {
                    if (subMenuItem.Name == _name)
                        subMenuItem.Enabled = _enabled;
                }
            }
        }

        private void SetMenuItemChecked(string _name, bool _checked)
        {
            foreach (MenuItem mainMenuItem in Menu.MenuItems)
            {
                foreach (MenuItem subMenuItem in mainMenuItem.MenuItems)
                {
                    if (subMenuItem.Name == _name)
                        subMenuItem.Checked = _checked;
                }
            }
        }

        private void UpdateTable(DataGridView _dgv, string _tableName)
        {
            TaskDialog td = new TaskDialog("Update Table");
            td.MainIcon = TaskDialogIcon.TaskDialogIconNone;
            td.MainInstruction = "Are you sure you want to update the database table?";
            td.MainContent = "This will not update elements in the Revit document. This will only update the database table.";
            td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;

            if (td.Show() == TaskDialogResult.Yes)
            {
                foreach (DataGridViewRow row in _dgv.Rows)
                {
                    int currentDictIndex = 0;

                    SqlCommand command = new SqlCommand();
                    command.Connection = con;
                    command.CommandType = CommandType.Text;
                    command.Connection = con;
                    command.CommandType = CommandType.Text;

                    StringBuilder commandText = new StringBuilder();
                    commandText.Append("UPDATE [" + _tableName + "] SET ");

                    con.Open();
                    Dictionary<string, string> dict = new Dictionary<string, string>();
                    dict = CellValueDictionary(row);

                    foreach (string param in dict.Keys)
                    {
                        string mark = "";
                        mark = row.Cells["Mark"].Value.ToString();
                        string parameterValue = "";
                        parameterValue = dict[param];
                        command.Parameters.AddWithValue("@" + param, parameterValue + "");

                        if (currentDictIndex < dict.Count - 1)
                            commandText.Append(param + "=@" + param + ",");
                        else if (currentDictIndex == dict.Count - 1)
                            commandText.Append(param + "=@" + param + " WHERE Mark=\'" + mark + "\'");

                        currentDictIndex++;
                    }

                    command.CommandText = commandText.ToString();
                    command.ExecuteNonQuery();
                    con.Close();
                }

                if (!tableIsBeingFiltered)
                    LoadTable(_tableName);
                else
                    FilterTable(_tableName);
            }
        }

        private void LoadElements(string _tableName)
        {
            FilteredElementCollector vsCol = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule));
            foreach (ViewSchedule vs in vsCol)
            {
                if (vs.ViewName == _tableName)
                {
                    List<Element> elems = new List<Element>();
                    elems = new FilteredElementCollector(doc, vs.Id).OfClass(typeof(FamilyInstance)).ToElements().ToList();

                    foreach (Element elem in elems)
                    {
                        string mark = "";

                        mark = elem.get_Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString();

                        if (!markValues.Contains(mark))
                            InsertTableEntry(elem, _tableName);
                        else
                            UpdateTableEntry(elem, _tableName);
                    }

                    LoadTable(_tableName);

                }
            }

        }
        
        private void UpdateTableEntry(Element _element, string _tableName)
        {
            string id = "";
            string mark = "";

            id = _element.Id.ToString();
            mark = _element.get_Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString();

            Dictionary<string, Guid> paramDict = new Dictionary<string, Guid>();
            paramDict = ElementParameterDictionary(_element);

            con.Open();

            SqlCommand command = new SqlCommand();
            command.Connection = con;
            command.CommandType = CommandType.Text;
            command.Connection = con;
            command.CommandType = CommandType.Text;

            StringBuilder commandText = new StringBuilder();

            commandText.Append("UPDATE [" + _tableName + "] SET ElementId=@ElementId,");
            command.Parameters.AddWithValue("@ElementId", id);
            command.Parameters.AddWithValue("@Mark", mark);

            int currentIndex = 0;

            foreach (string param in paramDict.Keys)
            {
                Guid guid = paramDict[param];
                string parameterValue = _element.get_Parameter(guid).AsString();

                if (parameterValue == null)
                    parameterValue = "";

                command.Parameters.AddWithValue("@" + param, parameterValue + "");

                if (currentIndex < paramDict.Count - 1)
                    commandText.Append(param + "=@" + param + ",");
                else if (currentIndex == paramDict.Count - 1)
                    commandText.Append(param + "=@" + param + " WHERE Mark=@Mark");

                currentIndex += 1;
            }

            command.CommandText = commandText.ToString();
            command.ExecuteNonQuery();
            con.Close();

        }

        private void InsertTableEntry(Element _element, string _tableName)
        {
            string id = "";
            string mark = "";

            id = _element.Id.ToString();
            mark = _element.get_Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString();

            Dictionary<string, Guid> paramDict = new Dictionary<string, Guid>();
            paramDict = ElementParameterDictionary(_element);

            con.Open();

            SqlCommand command = new SqlCommand();
            command.Connection = con;
            command.CommandType = CommandType.Text;
            StringBuilder commandText = new StringBuilder();

            commandText.Append("INSERT INTO [" + _tableName + "] (ElementId,Mark,");
            command.Parameters.AddWithValue("@ElementId", id);
            command.Parameters.AddWithValue("@Mark", mark);

            int currentIndex = 0;

            foreach (string param in paramDict.Keys)
            {
                Guid guid = paramDict[param];

                string parameterValue = _element.get_Parameter(guid).AsString();

                if (parameterValue == null)
                    parameterValue = "";

                command.Parameters.AddWithValue(("@" + param), parameterValue);

                if (currentIndex < paramDict.Count - 1)
                    commandText.Append(param + ",");
                else if (currentIndex == paramDict.Count - 1)
                    commandText.Append((param + ") VALUES (@ElementId,@Mark,"));

                currentIndex += 1;
            }

            currentIndex = 0;

            foreach (string param in paramDict.Keys)
            {
                if (currentIndex < paramDict.Count - 1)
                    commandText.Append("@" + param + ",");
                else if (currentIndex == paramDict.Count - 1)
                    commandText.Append("@" + param + ")");

                currentIndex += 1;
            }

            command.CommandText = commandText.ToString();
            command.ExecuteNonQuery();
            con.Close();

        }

        private void UpdateParameterValue(Element _element, Guid _guid, string _paramValue)
        {
            Parameter param = _element.get_Parameter(_guid);
            param.Set(_paramValue);
        }

        private void UpdateElements(DataGridView _dgv)
        {
            TaskDialog td = new TaskDialog("Update Elements");
            td.MainIcon = TaskDialogIcon.TaskDialogIconNone;
            td.MainInstruction = "Are you sure you want to update the elements?";
            td.MainContent = "This will only update the elements. This will not update the database table.";
            td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;

            if (td.Show() == TaskDialogResult.Yes)
            {
                try
                {
                    using (Transaction t = new Transaction(doc, "Update Elements"))
                    {
                        t.Start();

                        foreach (DataGridViewRow row in _dgv.Rows)
                        {
                            int intElemId = 0;
                            string strElemId = row.Cells["ElementId"].Value.ToString();
                            string mark = row.Cells["Mark"].Value.ToString();

                            intElemId = Convert.ToInt32(strElemId);
                            ElementId id = new ElementId(intElemId);

                            Element elem;
                            elem = SelectElement(uIApp, id);

                            Dictionary<string, Guid> dict = new Dictionary<string, Guid>();
                            dict = this.ElementParameterDictionary(elem);

                            foreach (string param in dict.Keys)
                            {
                                Guid guid = dict[param];

                                string paramValue = row.Cells[param].Value.ToString();

                                if ((paramValue == null))
                                    paramValue = "";

                                UpdateParameterValue(elem, guid, paramValue);
                            }
                        }

                        t.Commit();
                        Close();

                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", ex.Message + "\n\n" + ex.Source);
                }
            }

        }

        private void UpdateElementsAndTable(string _tableName)
        {
            TaskDialog td = new TaskDialog("Update Elements & Table");
            td.MainIcon = TaskDialogIcon.TaskDialogIconNone;
            td.MainInstruction = "Are you sure you want to update the elements and the database table?";
            td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;

            if (td.Show() == TaskDialogResult.Yes)
            {
                try
                {
                    using (Transaction t = new Transaction(doc, "Update Elements & Table"))
                    {
                        t.Start();

                        foreach (DataGridViewRow row in dgvData.Rows)
                        {
                            int intElemId = 0;
                            string strElemId = row.Cells["ElementId"].Value.ToString();
                            string mark = row.Cells["Mark"].Value.ToString();

                            intElemId = Convert.ToInt32(strElemId);
                            ElementId id = new ElementId(intElemId);

                            Element elem;
                            elem = SelectElement(uIApp, id);

                            Dictionary<string, Guid> dict = new Dictionary<string, Guid>();
                            dict = this.ElementParameterDictionary(elem);

                            foreach (string param in dict.Keys)
                            {
                                Guid guid = dict[param];

                                string paramValue = row.Cells[param].Value.ToString();

                                if ((paramValue == null))
                                    paramValue = "";

                                UpdateParameterValue(elem, guid, paramValue);
                            }

                            UpdateTableEntry(elem, _tableName);
                        }

                        t.Commit();
                        LoadTable(_tableName);
                        Close();

                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", ex.Message + "\n\n" + ex.Source);
                }
            }
        }

        private void FillDataGridView()
        {

            DrawingControl.SetDoubleBuffered(dgvData);
            DrawingControl.SuspendDrawing(dgvData);

            dgvData.DataSource = null;

            dgvData.Columns.Clear();

            dt.Columns.Clear();
            dt.Clear();

            da.Fill(dt);

            dgvData.DataSource = dt.DefaultView;

            dgvData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
            dgvData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvData.ReadOnly = false;

            //ID
            dgvData.Columns["ID"].Visible = false;
            dgvData.Columns["ID"].ReadOnly = true;

            //ElementId
            dgvData.Columns["ElementId"].Visible = false;
            dgvData.Columns["ElementId"].ReadOnly = true;

            //Mark
            dgvData.Columns["Mark"].Visible = true;
            dgvData.Columns["Mark"].ReadOnly = true;

            dgvData.ClearSelection();

            DrawingControl.ResumeDrawing(dgvData);

        }

        private void CreateTable(string _tableName)
        {
            tableIsBeingFiltered = false;
            FilteredElementCollector vsCol;
            vsCol = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule));

            ViewScheduleExportOptions vsOpt;
            vsOpt = new ViewScheduleExportOptions();

            vsOpt.Title = false;
            vsOpt.HeadersFootersBlanks = false;
            vsOpt.FieldDelimiter = ",";

            string dir = "";
            dir = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            string fileName = "";
            fileName = _tableName + ".txt";

            string file = "";
            file = dir + @"\" + fileName;

            //iterate through all the ViewSchedules until you get the selected one
            foreach (ViewSchedule vs in vsCol)
            {
                if (vs.ViewName == _tableName)
                    vs.Export(dir, fileName, vsOpt); //export the ViewSchedule to MyDocuments
            }

            //read from the file that was just exported above
            //read only the first line
            //the first line contains the fields to create in the database table
            System.IO.StreamReader sr = new System.IO.StreamReader(file);

            string entry = "";
            entry = sr.ReadLine();

            char[] cs = new char[] { ',' };
            string[] ar = entry.Split(cs, StringSplitOptions.None);

            string curParam = "";
            int curArIndex = 0;

            //create a string builder to dynamically create an sql query string
            StringBuilder commandText = new StringBuilder();
            commandText.Append("CREATE TABLE " + _tableName + " (");
            commandText.Append("ID int PRIMARY KEY IDENTITY,");
            commandText.Append("ElementId varchar(MAX),");

            //interate through all the parameter names in the ar array
            //append them to the string builder
            do
            {
                curParam = ar[curArIndex];
                curParam = curParam.Replace(@"""", "");

                if (curArIndex < ar.Length - 1)
                    commandText.Append(curParam + " varchar(MAX),");
                else if (curArIndex == ar.Length - 1)
                    commandText.Append(curParam + " varchar(MAX));");

                curArIndex++;

            } while (curArIndex <= ar.Length - 1);

            //connect to the database and create the table
            con.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = con;
            command.CommandType = CommandType.Text;
            command.CommandText = commandText.ToString();
            command.ExecuteNonQuery();
            con.Close();

            LoadTable(_tableName);
            LoadElements(_tableName);
        }

        private void LoadTable(string _tableName)
        {
            tableIsBeingFiltered = false;
            con.Open();
            string sql = "SELECT * FROM [" + _tableName + "]";

            FillDataSet(sql, _tableName);
            FillDataGridView();

            con.Close();
            markValues.Clear();

            foreach (DataGridViewRow row in dgvData.Rows)
                markValues.Add(row.Cells["Mark"].Value.ToString());

            SetFilterToolsEnabled(true);
            dgvData.ClearSelection();
        }

        private void GetAllViewSchedules(System.Windows.Forms.ComboBox _comboBox)
        {
            _comboBox.Items.Clear();
            FilteredElementCollector vsCol = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule));

            foreach (ViewSchedule vs in vsCol)
            {
                if (!vs.IsTemplate && !vs.IsTitleblockRevisionSchedule)
                {
                    string vsName = "";
                    vsName = vs.ViewName;

                    if (vsName.StartsWith("tbl"))
                        _comboBox.Items.Add(vsName);

                }
            }
        }

        private void DeleteElementsFromTable(string _tableName)
        {
            TaskDialog td = new TaskDialog("Delete");
            td.MainIcon = TaskDialogIcon.TaskDialogIconNone;
            td.MainInstruction = "Are you sure you want to delete the selected entries?";
            td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;

            if (td.Show() == TaskDialogResult.Yes)
            {
                con.Open();

                foreach (DataGridViewCell cell in dgvData.SelectedCells)
                {
                    int index = 0;
                    index = cell.RowIndex;
                    int id = Convert.ToInt32(dgvData["ID", index].Value);
                    SQLCommand("DELETE FROM [" + _tableName + "] WHERE ID = " + id + "");
                }

                con.Close();

                if (!tableIsBeingFiltered)
                    LoadTable(_tableName);
                else
                    FilterTable(_tableName);
            }

        }

        private void SQLCommand(string _sql)
        {
            SqlCommand command = new SqlCommand(_sql);
            command.Connection = con;
            command.ExecuteNonQuery();
        }

        private void FilterTable(string _tableName)
        {
            tableIsBeingFiltered = true;
            List<string> tableFields = new List<string>();

            con.Open();
            DataColumnCollection col;
            col = dt.Columns;

            foreach (DataColumn column in col)
            {
                if (column.ColumnName != "ID")
                {
                    string param = column.ColumnName;
                    tableFields.Add(param);
                }
            }

            string searchString = txtFilter.Text;

            DrawingControl.SetDoubleBuffered(dgvData);
            DrawingControl.SuspendDrawing(dgvData);

            dgvData.DataSource = null;
            dgvData.Columns.Clear();

            DataView dsView = new DataView();
            dsView = ds.Tables[0].DefaultView;

            BindingSource bs = new BindingSource();
            bs.DataSource = dsView;

            string filterString = FilterLikeString(tableFields, searchString);
            bs.Filter = filterString;

            dgvData.DataSource = bs;

            con.Close();

            dgvData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
            dgvData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvData.ReadOnly = false;

            //ID
            dgvData.Columns["ID"].Visible = false;
            dgvData.Columns["ID"].ReadOnly = true;

            //ElementId
            dgvData.Columns["ElementId"].Visible = false;
            dgvData.Columns["ElementId"].ReadOnly = true;

            //Mark
            dgvData.Columns["Mark"].Visible = true;
            dgvData.Columns["Mark"].ReadOnly = true;

            dgvData.ClearSelection();

            DrawingControl.ResumeDrawing(dgvData);
        }
        
        private void ClearValues()
        {
            TaskDialog td = new TaskDialog("Clear Values");
            td.MainIcon = TaskDialogIcon.TaskDialogIconNone;
            td.MainInstruction = "Are you sure you want to clear the values of the selected cells?";
            td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;

            if (td.Show() == TaskDialogResult.Yes)
            {
                foreach (DataGridViewCell cell in dgvData.SelectedCells)
                    cell.Value = "";
            }
        }

        private void PasteValue()
        {
            TaskDialog td = new TaskDialog("Paste Value");
            td.MainIcon = TaskDialogIcon.TaskDialogIconNone;
            td.MainInstruction = "Are you sure you want to paste the value below to the selected cells?";
            td.MainContent = CopiedValue;
            td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;

            if (td.Show() == TaskDialogResult.Yes)
            {
                foreach (DataGridViewCell cell in dgvData.SelectedCells)
                    cell.Value = CopiedValue;
            }
        }

        private void CreateStatusBar()
        {
            databasePanel.BorderStyle = StatusBarPanelBorderStyle.Sunken;
            databasePanel.Text = "Database:";
            databasePanel.AutoSize = StatusBarPanelAutoSize.Spring;
            mainStatusBar.Panels.Add(databasePanel);

            tablePanel.BorderStyle = StatusBarPanelBorderStyle.Sunken;
            tablePanel.Text = "Table:";
            tablePanel.AutoSize = StatusBarPanelAutoSize.Spring;
            mainStatusBar.Panels.Add(tablePanel);

            mainStatusBar.ShowPanels = true;

            this.Controls.Add(mainStatusBar);
        }

        private void CreateMainMenu()
        {
            MainMenu mainMenu = new MainMenu();

            #region FileMenu

            MenuItem fileMenu = new MenuItem();
            MenuItem menuCreateTable = new MenuItem();
            MenuItem menuCloseTable = new MenuItem();
            MenuItem menuExit = new MenuItem();

            menuCreateTable.Name = "create_table";
            menuCloseTable.Name = "close_table";

            fileMenu.Text = "File";

            menuCreateTable.Text = "Create Table\u2026";
            menuCloseTable.Text = "Close Table";

            menuExit.Text = "Exit";

            fileMenu.MenuItems.Add(menuCreateTable);
            fileMenu.MenuItems.Add(menuCloseTable);
            fileMenu.MenuItems.Add("-");
            fileMenu.MenuItems.Add(menuExit);

            menuCreateTable.Click += new System.EventHandler(this.menuCreateTable_Click);
            menuCloseTable.Click += new System.EventHandler(this.menuCloseTable_Click);
            menuExit.Click += new System.EventHandler(this.menuExit_Click);

            #endregion

            #region EditMenu

            MenuItem editMenu = new MenuItem();
            MenuItem menuLoadElements = new MenuItem();
            MenuItem menuUpdateElements = new MenuItem();
            MenuItem menuUpdateTable = new MenuItem();
            MenuItem menuUpdateElementsAndTable = new MenuItem();
            MenuItem menuDeleteElementsFromTable = new MenuItem();

            menuLoadElements.Name = "load_elements";
            menuUpdateElements.Name = "update_elements";
            menuUpdateTable.Name = "update_table";
            menuUpdateElementsAndTable.Name = "update_elements_and_table";
            menuDeleteElementsFromTable.Name = "delete_elements_from_table";

            editMenu.Text = "Edit";

            menuLoadElements.Text = "Load Elements";
            menuUpdateElements.Text = "Update Elements";
            menuUpdateTable.Text = "Update Table";
            menuUpdateElementsAndTable.Text = "Update Elements && Table";
            menuDeleteElementsFromTable.Text = "Delete";

            editMenu.MenuItems.Add(menuLoadElements);
            editMenu.MenuItems.Add("-");
            editMenu.MenuItems.Add(menuUpdateElements);
            editMenu.MenuItems.Add(menuUpdateTable);
            editMenu.MenuItems.Add(menuUpdateElementsAndTable);
            editMenu.MenuItems.Add("-");
            editMenu.MenuItems.Add(menuDeleteElementsFromTable);

            menuLoadElements.Click += new System.EventHandler(this.menuLoadElements_Click);
            menuUpdateElements.Click += new System.EventHandler(this.menuUpdateElements_Click);
            menuUpdateTable.Click += new System.EventHandler(this.menuUpdateTable_Click);
            menuUpdateElementsAndTable.Click += new System.EventHandler(this.menuUpdateElementsAndTable_Click);
            menuDeleteElementsFromTable.Click += new System.EventHandler(this.menuDeleteElementsFromTable_Click);

            #endregion

            mainMenu.MenuItems.Add(fileMenu);
            mainMenu.MenuItems.Add(editMenu);

            this.Menu = mainMenu;

        }

        #endregion

        #region Functions

        private Dictionary<string, string> CellValueDictionary(DataGridViewRow _row)
        {
            DataColumnCollection col;
            col = dt.Columns;
            Dictionary<string, string> dict = new Dictionary<string, string>();

            foreach (DataColumn column in col)
            {
                if (column.ColumnName != "ID" && column.ColumnName != "ElementId" && column.ColumnName != "Mark")
                {
                    string columnName = "";
                    columnName = column.ColumnName;
                    string cellValue = "";
                    cellValue = _row.Cells[columnName].Value.ToString();
                    dict.Add(columnName, cellValue);
                }

            }

            return dict;
        }

        private Element SelectElement(UIApplication _uiApp, ElementId _id)
        {
            Element elem = _uiApp.ActiveUIDocument.Document.GetElement(_id);
            return elem;
        }

        private Dictionary<string, Guid> ElementParameterDictionary(Element _element)
        {
            DataColumnCollection col;
            col = dt.Columns;

            Dictionary<string, Guid> dict = new Dictionary<string, Guid>();

            foreach (DataColumn column in col)
            {
                if (column.ColumnName != "ID"
                    && column.ColumnName != "ElementId"
                    && column.ColumnName != "Mark")
                {
                    string param = column.ColumnName;
                    Guid guid = _element.LookupParameter(column.ColumnName).GUID;
                    dict.Add(param, guid);
                }
            }

            return dict;
        }

        private DataSet FillDataSet(string _sql, string _tableName)
        {
            da = new SqlDataAdapter(_sql, con);
            ds.Tables.Clear();
            ds.Clear();
            da.Fill(ds, "[" + _tableName + "]");

            return ds;
        }

        private bool TableExists(string _tableName)
        {
            bool tableExists;

            try
            {
                con.Open();
                var cmd = new SqlCommand("select case when exists((select * from information_schema.tables where table_name = '" + _tableName + "')) then 1 else 0 end");
                cmd.Connection = con;

                tableExists = (int)cmd.ExecuteScalar() == 1;
            }
            catch
            {
                tableExists = false;
            }
            finally
            {
                con.Close();
            }

            return tableExists;
        }

        private string FilterLikeString(List<string> _tableFields, string _searchString)
        {
            string filterString = "";
            StringBuilder likeFilter = new StringBuilder();
            int counter = 0;

            foreach (string field in _tableFields)
            {
                int fieldCount = _tableFields.Count - 1;

                if (counter < fieldCount)
                    likeFilter.Append(field + " LIKE'%" + _searchString + "%' or ");
                else
                    likeFilter.Append(field + " LIKE'%" + _searchString + "%'");

                counter += 1;
            }

            filterString = likeFilter.ToString();

            return filterString;
        }

        private ContextMenu TableContextMenu()
        {
            ContextMenu mnu = new ContextMenu();
            MenuItem cxmnuCopyValue = new MenuItem("Copy Value");
            MenuItem cxmnuPasteValue = new MenuItem("Paste Value");
            MenuItem cxmnuClearValues = new MenuItem("Clear Values");
            MenuItem cxmnuDeleteTableEntry = new MenuItem("Delete");

            cxmnuDeleteTableEntry.Click += new EventHandler(cxmnuDeleteTableEntry_Click);
            cxmnuClearValues.Click += new EventHandler(cxmnuClearValues_Click);
            cxmnuCopyValue.Click += new EventHandler(cxmnuCopyValue_Click);
            cxmnuPasteValue.Click += new EventHandler(cxmnuPasteValue_Click);
            
            mnu.MenuItems.Add(cxmnuCopyValue);
            mnu.MenuItems.Add(cxmnuPasteValue);
            mnu.MenuItems.Add("-");
            mnu.MenuItems.Add(cxmnuClearValues);
            mnu.MenuItems.Add("-");
            mnu.MenuItems.Add(cxmnuDeleteTableEntry);

            return mnu;
        }
        
        #endregion

        #region MainMenu Events

        private void menuLoadElements_Click(object sender, EventArgs e)
        { 
            LoadElements(SelectedTable);
        }

        private void menuUpdateElements_Click(object sender, EventArgs e)
        {
            UpdateElements(dgvData);
        }

        private void menuUpdateTable_Click(object sender, EventArgs e)
        {
            UpdateTable(dgvData, SelectedTable);
        }

        private void menuUpdateElementsAndTable_Click(object sender, EventArgs e)
        {
            UpdateElementsAndTable(SelectedTable);
        }

        private void menuDeleteElementsFromTable_Click(object sender, EventArgs e)
        { 
            DeleteElementsFromTable(SelectedTable);
        }

        private void menuSelectDatabase_Click(object sender, EventArgs e)
        {
            frmSelectionBox new_frmSelectionBox = new frmSelectionBox();
            new_frmSelectionBox.Text = "Select Database";
            new_frmSelectionBox.lblInstructions.Text = "Select the database you want to connect to\nfrom the drop-down list below";

            System.Windows.Forms.ComboBox cbDatabases;
            cbDatabases = new_frmSelectionBox.cbItems;

            GetAllDatabases(cbDatabases);

            if (new_frmSelectionBox.ShowDialog() == DialogResult.OK)
            {
                con.ConnectionString = "";
                SelectedDatabase = "";
                SelectedDatabase = cbDatabases.SelectedItem.ToString();
                con.ConnectionString = @"Server=BPRMEPCM-7\SQLEXPRESS;Database=" + SelectedDatabase + @";Trusted_Connection=True;";
                databasePanel.Text = "Database: " + SelectedDatabase;
                tablePanel.Text = "Table:";
                tableIsBeingFiltered = false;
            }
        }

        private void menuCreateTable_Click(object sender, EventArgs e)
        {
            frmSelectionBox new_frmSelectionBox = new frmSelectionBox();
            new_frmSelectionBox.Text = "Create Table";
            new_frmSelectionBox.lblInstructions.Text = "Select the schedule you want to create\na table for from the down list below";

            System.Windows.Forms.ComboBox cbSchedules;
            cbSchedules = new_frmSelectionBox.cbItems;

            GetAllViewSchedules(cbSchedules);

            if (new_frmSelectionBox.ShowDialog() == DialogResult.OK)
            {
                SelectedTable = "";
                SelectedTable = cbSchedules.SelectedItem.ToString();
                if (TableExists(SelectedTable))
                {
                    TaskDialog td = new TaskDialog("Table Exists");
                    td.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
                    td.MainInstruction = "The selected schedule already exists in the database";
                    td.CommonButtons = TaskDialogCommonButtons.Ok;

                    td.Show();
                }
                else
                {
                    SelectedTable = cbSchedules.SelectedItem.ToString();
                    CreateTable(SelectedTable);
                    databasePanel.Text = "Database: " + SelectedDatabase;
                    tablePanel.Text = "Table: " + SelectedTable;
                    GetAllTablesFromDatabase(cbTables);
                    cbTables.SelectedIndex = cbTables.FindStringExact(SelectedTable);

                }
            }
        }

        private void menuLoadTable_Click(object sender, EventArgs e)
        {
            frmSelectionBox new_frmSelectionBox = new frmSelectionBox();
            new_frmSelectionBox.Text = "Load Table";
            new_frmSelectionBox.lblInstructions.Text = "Select the table you want to load\nfrom the drop-down list below";

            System.Windows.Forms.ComboBox cbTables;
            cbTables = new_frmSelectionBox.cbItems;

            if (SelectedDatabase != "")
                GetAllTablesFromDatabase(cbTables);
            
            if (new_frmSelectionBox.ShowDialog() == DialogResult.OK)
            {
                SelectedTable = cbTables.SelectedItem.ToString();
                LoadTable(SelectedTable);
                databasePanel.Text = "Database: " + SelectedDatabase;
                tablePanel.Text = "Table: " + SelectedTable;
            }
        }

        private void menuCloseTable_Click(object sender, EventArgs e)
        {
            dgvData.DataSource = null;
            dgvData.Columns.Clear();
            dt.Columns.Clear();
            dt.Clear();
            SelectedTable = "";
            tableIsBeingFiltered = false;
            databasePanel.Text = "Database: " + SelectedDatabase;
            tablePanel.Text = "Table:";
            SetFilterToolsEnabled(false);
            cbTables.SelectedIndex = -1;
        }

        private void menuExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #endregion

        #region Button Events

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtFilter.Text = "";
            FilterTable(SelectedTable);
            tableIsBeingFiltered = false;
        }

        private void btnFilter_Click(object sender, EventArgs e)
        {
            FilterTable(SelectedTable);
        }

        #region ContextMenu

        private void cxmnuDeleteTableEntry_Click(object sender, EventArgs e)
        {
            DeleteElementsFromTable(SelectedTable);
        }

        private void cxmnuClearValues_Click(object sender, EventArgs e)
        {
            ClearValues();
        }

        private void cxmnuCopyValue_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in dgvData.SelectedCells)
            {
                if (dgvData.SelectedCells.Count == 1)
                    CopiedValue = cell.Value.ToString();
                else
                    return;
            }
        }

        private void cxmnuPasteValue_Click(object sender, EventArgs e)
        {
            PasteValue();
        }
        
        #endregion

        #endregion

        private void dgvData_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenu contextMenu = new ContextMenu();
                contextMenu = TableContextMenu();
                contextMenu.Show(dgvData, new System.Drawing.Point(e.X, e.Y));
            }

        }

        private void cbProjects_SelectedIndexChanged(object sender, EventArgs e)
        {
            SelectedDatabase = cbProjects.SelectedItem.ToString();
            if (SelectedDatabase != "" && cbProjects.SelectedIndex != -1)
            {
                con.ConnectionString = "";
                SelectedDatabase = "";
                SelectedDatabase = cbProjects.SelectedItem.ToString();
                con.ConnectionString = @"Server=BPRMEPCM-7\SQLEXPRESS;Database=" + SelectedDatabase + @";Trusted_Connection=True;";
                databasePanel.Text = "Database: " + SelectedDatabase;
                tablePanel.Text = "Table:";
                tableIsBeingFiltered = false;
                GetAllTablesFromDatabase(cbTables);
            }
        }

        private void cbTables_SelectedIndexChanged(object sender, EventArgs e)
        {
            SelectedTable = cbTables.SelectedItem.ToString();
            if (SelectedDatabase != "" && SelectedTable != "")
            {
                LoadTable(SelectedTable);
                databasePanel.Text = "Database: " + SelectedDatabase;
                tablePanel.Text = "Table: " + SelectedTable;
            }
        }
    }

    #region DrawingControl

    public static class DrawingControl
    {
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr _hWnd, Int32 _wMsg, bool _wParam, Int32 _lParam);

        private const int WM_SETREDRAW = 11;

        public static void SetDoubleBuffered(System.Windows.Forms.Control _ctrl)
        {
            if (!SystemInformation.TerminalServerSession)
            {
                typeof(System.Windows.Forms.Control).InvokeMember("DoubleBuffered", (System.Reflection.BindingFlags.SetProperty
                                | (System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)), null, _ctrl, new object[] {
                            true});
            }
        }

        public static void SetDoubleBuffered_ListControls(List<System.Windows.Forms.Control> _ctrlList)
        {
            if (!SystemInformation.TerminalServerSession)
            {
                foreach (System.Windows.Forms.Control ctrl in _ctrlList)
                {
                    typeof(System.Windows.Forms.Control).InvokeMember("DoubleBuffered", (System.Reflection.BindingFlags.SetProperty
                                    | (System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)), null, ctrl, new object[] {
                                true});
                }
            }
        }

        public static void SuspendDrawing(System.Windows.Forms.Control _ctrl)
        {
            SendMessage(_ctrl.Handle, WM_SETREDRAW, false, 0);
        }

        public static void SuspendDrawing_ListControls(List<System.Windows.Forms.Control> _ctrlList)
        {
            foreach (System.Windows.Forms.Control ctrl in _ctrlList)
            {
                SendMessage(ctrl.Handle, WM_SETREDRAW, false, 0);
            }
        }

        public static void ResumeDrawing(System.Windows.Forms.Control _ctrl)
        {
            SendMessage(_ctrl.Handle, WM_SETREDRAW, true, 0);
            _ctrl.Refresh();
        }

        public static void ResumeDrawing_ListControls(List<System.Windows.Forms.Control> _ctrlList)
        {
            foreach (System.Windows.Forms.Control ctrl in _ctrlList)
            {
                SendMessage(ctrl.Handle, WM_SETREDRAW, true, 0);
                ctrl.Refresh();
            }
        }
    }

    #endregion

}
