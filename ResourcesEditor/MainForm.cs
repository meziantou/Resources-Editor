using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.IsolatedStorage;
using System.Linq;
using System.Net.NetworkInformation;
using System.Resources;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ResourcesEditor
{
    public partial class MainForm : Form
    {
        private bool _isDirty = false;
        private string _file;
        private DataTable _table;
        private DataColumn _keyColumn;
        private DataColumn _commentColumn;
        private readonly string _recentFilesPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Meziantou", "ResourcesEditor", "Recents.txt");
        private bool _confirmExit = false;

        public MainForm()
        {
            InitializeComponent();
        }

        private bool ExitApplication()
        {
            if (!_confirmExit && _isDirty)
            {
                DialogResult dialogResult = MessageBox.Show("Do you want to save data before exiting?", "Save?", MessageBoxButtons.YesNoCancel);
                if (dialogResult == DialogResult.Cancel)
                    return false;

                if (dialogResult == DialogResult.Yes)
                {
                    Save();
                }
                _confirmExit = true;
            }

            Application.Exit();
            return true;
        }

        private void Save()
        {
            if (_table == null)
                return;

            dataGridView1.EndEdit();

            for (int i = 1; i < _table.Columns.Count - 1; i++)
            {
                string fullPath = _table.Columns[i].ExtendedProperties["File"] as string;
                if (fullPath == null)
                    continue;

                UnlockFile(fullPath);

                IEnumerable<ResXDataNode> FileRefs = ReadResXFile(fullPath).Where(_ => _.FileRef != null);

                using (ResXResourceWriter writer = new ResXResourceWriter(fullPath))
                {
                    foreach (var resXDataNode in FileRefs)
                    {
                        writer.AddResource(resXDataNode);
                    }

                    foreach (DataRow row in _table.Rows)
                    {
                        var resXFileRef = row[i] as ResXFileRef;
                        if (resXFileRef != null)
                        {
                            var node = new ResXDataNode((string)row[_keyColumn], resXFileRef);
                            node.Comment = row[_commentColumn] as string;
                            writer.AddResource(node);
                        }
                        else
                        {
                            var txt = row[i] as string;
                            if (txt != null)
                            {
                                var node = new ResXDataNode((string)row[_keyColumn], row[i]);
                                node.Comment = row[_commentColumn] as string;
                                writer.AddResource(node);
                            }
                        }
                    }

                    writer.Generate();
                }
            }

            _isDirty = false;
            MessageBox.Show("Data saved");
        }

        private void UnlockFile(string path)
        {
            try
            {
                if (!File.Exists(path))
                    return;

                FileInfo fi = new FileInfo(path);
                fi.IsReadOnly = false;
            }
            catch
            {
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About about = new About();
            about.ShowDialog();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExitApplication();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!ExitApplication())
                e.Cancel = true;
        }

        private void openFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = _resxOpenFileDialog.ShowDialog();
            if (dialogResult == DialogResult.OK || dialogResult == DialogResult.Yes)
            {
                string fullPath = Path.GetFullPath(_resxOpenFileDialog.FileName);

                OpenFile(fullPath);
            }
        }

        private void OpenFile(string fullPath)
        {
            string prefix;
            try
            {
                List<string> resXFiles = new List<string>();
                string directory = Path.GetDirectoryName(fullPath);
                string fileName = Path.GetFileNameWithoutExtension(fullPath);

                var culture = GetCultureFromName(fileName, out prefix);
                try
                {
                    string filter = prefix + "*.resx";
                    string[] files = Directory.GetFiles(directory, filter);
                    foreach (var file in files)
                    {
                        string p;
                        GetCultureFromName(file, out p);
                        if (string.Equals(prefix, p, StringComparison.OrdinalIgnoreCase))
                            resXFiles.Add(file);
                    }
                }
                catch
                {
                }

                _table = new DataTable();


                _keyColumn = _table.Columns.Add("Key", typeof(string));
                _keyColumn.Unique = true;

                _commentColumn = _table.Columns.Add("Comment", typeof(string));

                foreach (var resXFile in resXFiles)
                {
                    List<ResXDataNode> nodes = ReadResXFile(resXFile);

                    CultureInfo c = GetCultureFromName(resXFile, out prefix);
                    DataColumn cultureColumn = new DataColumn(c == null ? " " : c.Name);
                    cultureColumn.ExtendedProperties.Add("File", resXFile);
                    _table.Columns.Add(cultureColumn);

                    foreach (var node in nodes)
                    {
                        if (node.FileRef != null)
                        {
                            continue;
                        }
                        // Search row
                        bool found = false;
                        for (int i = 0; i < _table.Rows.Count; i++)
                        {
                            if (((string)_table.Rows[i][_keyColumn]) == node.Name)
                            {
                                _table.Rows[i][cultureColumn] = node.GetValue((ITypeResolutionService)null);
                                if (_table.Rows[i][_commentColumn] == null)
                                    _table.Rows[i][_commentColumn] = node.Comment;
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            DataRow dataRow = _table.NewRow();
                            dataRow[_keyColumn] = node.Name;
                            dataRow[_commentColumn] = node.Comment;
                            dataRow[cultureColumn] = node.GetValue((ITypeResolutionService)null);
                            _table.Rows.Add(dataRow);
                        }
                    }
                }

                _commentColumn.SetOrdinal(_table.Columns.Count - 1);

                bindingSource1.DataSource = _table;
                bindingSource1.ResetBindings(true);

                dataGridView1.DataSource = bindingSource1;
                dataGridView1.ResetBindings();

                _file = fullPath;
                _isDirty = false;
                AddRecentFile(fullPath);
                this.Text = "Resources Editor - " + Path.GetFileName(prefix);
            }
            catch (Exception)
            {
                MessageBox.Show("An error occurs.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddRecentFile(string path)
        {
            List<string> files = LoadRecentFiles();
            files.Insert(0, Path.GetFullPath(path));

            SetRecentFiles(files);

            Directory.CreateDirectory(Path.GetDirectoryName(_recentFilesPath));
            File.WriteAllLines(_recentFilesPath, files.Distinct().Take(10));
        }

        private void SetRecentFiles(IEnumerable<string> files)
        {
            recentToolStripMenuItem.DropDownItems.Clear();
            foreach (var file in files)
            {
                var item = new ToolStripMenuItem(file);
                item.Click += recentItem_Click;
                recentToolStripMenuItem.DropDownItems.Add(item);
            }
        }

        private void recentItem_Click(object sender, EventArgs e)
        {
            OpenFile(((ToolStripMenuItem)sender).Text);
        }

        private List<string> LoadRecentFiles()
        {
            List<string> result = new List<string>();

            if (File.Exists(_recentFilesPath))
            {
                result.AddRange(File.ReadAllLines(_recentFilesPath));
            }

            return result;
        }

        private static CultureInfo GetCultureFromName(string path, out string prefix)
        {
            if (path == null) throw new ArgumentNullException("path");

            path = Path.GetFileNameWithoutExtension(path);

            prefix = path;

            int index = path.LastIndexOf('.');
            if (index >= 0)
            {
                string cultureName = path.Substring(index + 1);
                try
                {
                    prefix = path.Substring(0, index);
                    return CultureInfo.GetCultureInfo(cultureName);
                }
                catch
                {
                }
            }

            return null;
        }

        private List<ResXDataNode> ReadResXFile(string path)
        {
            List<ResXDataNode> nodes = new List<ResXDataNode>();
            using (ResXResourceReader reader = new ResXResourceReader(path))
            {
                reader.UseResXDataNodes = true;
                foreach (DictionaryEntry d in reader)
                {
                    ResXDataNode dataNode = (ResXDataNode)d.Value;
                    nodes.Add(dataNode);
                }
            }

            return nodes;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            List<string> files = LoadRecentFiles();
            if (files.Count > 0)
                SetRecentFiles(files);
        }

        private void dataGridView1_CurrentCellChanged(object sender, EventArgs e)
        {
            _isDirty = true;
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex == dataGridView1.RowCount - 1)
                return; // Skip add row

            if (e.ColumnIndex < dataGridView1.ColumnCount - 1) // Skip Comment Column
            {
                if (e.Value == DBNull.Value || string.IsNullOrWhiteSpace((string)e.Value))
                {
                    e.CellStyle.BackColor = Color.MediumVioletRed;
                }
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (_table == null)
                return;

            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                bindingSource1.RemoveFilter();
                return;
            }

            string filter = null;
            foreach (DataColumn column in _table.Columns)
            {
                if (filter != null)
                    filter += " OR ";

                filter += string.Format("[{0}] like '%{1}%'", column.ColumnName, textBox1.Text);
            }

            bindingSource1.Filter = filter;
        }

        private void MainForm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        private void MainForm_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length > 0)
            {
                this.OpenFile(files[0]);
            }

        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if ((keyData & Keys.Control) == Keys.Control && (keyData & Keys.F) == Keys.F)
            {
                textBox1.Focus();
                textBox1.SelectAll();
                return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}
