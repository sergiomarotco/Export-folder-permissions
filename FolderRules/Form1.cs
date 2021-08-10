using System;
using System.Diagnostics;
using System.IO;
using System.Security.AccessControl;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Threading;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;

namespace FolderRules
{
    internal partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            systemusers = richTextBox1.Text.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            systemfolders = richTextBox1.Text.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            systemusers = richTextBox2.Text.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < systemusers.Length; i++)
                systemusers[i] = systemusers[i].Trim('\r');
            for (int i = 0; i < systemfolders.Length; i++)
                systemfolders[i] = systemfolders[i].Trim('\r');
        }
        string[] systemfolders;
        /// <summary>
        /// Start main process (analisyng folder access rules)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button1_Click(object sender, EventArgs e)
        {
            allow_work = true;
            button3.Enabled = true;
            button3.Visible = true;
            button1.Enabled = false;
            listView1.Items.Clear();
            Real_list.Clear();
            root_layer = textBox1.Text.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries).Length;
            Task.Factory.StartNew(() => Main_Task(textBox1.Text));
        }
        int root_layer = 0;
        List<ListItem_my> Real_list = new List<ListItem_my>(0);
        private void Main_Task(string RootDirectory)
        {
            Invoke((ThreadStart)delegate{listView1.Visible = false;});
            WalkDirectoryTree(RootDirectory);
            List<ListViewItem> L = new List<ListViewItem>(0);
            for (int i = 0; i < Real_list.Count; i++)
            {
                try
                {
                    Invoke((ThreadStart)delegate
                    {
                        ListViewItem dfsadfsd = new ListViewItem(Real_list[i].one_item);
                        //dfsadfsd.SubItems[0].Text = dfsadfsd.SubItems[0].Text.Replace(@"\", ".");
                        //dfsadfsd.SubItems[1].Text = "1";
                        //dfsadfsd.SubItems[2].Text = "2";
                        //dfsadfsd.SubItems[3].Text = "3";
                        //dfsadfsd.SubItems[4].Text = "4";
                        //dfsadfsd.SubItems[5].Text = "5";
                        listView1.Items.Add(dfsadfsd);
                    });
                }
                catch (Exception ee)
                {
                    MessageBox.Show(ee.Message);
                }
            }

            MessageBox.Show("Done");

            Invoke((ThreadStart)delegate
            {
                button1.Enabled = true;
                button3.Enabled = false;
                button3.Visible = false;
            });
            Invoke((ThreadStart)delegate
            {
                listView1.Visible = true;
            });
        }
        /// <summary>
        /// Exporting permissions of directory 
        /// </summary>
        /// <param name="directory">Папка права которой требуется вычислить</param>
        private void WalkDirectoryTree(string directory)
        {
            if (allow_work)
            {   
                DirectoryInfo root = new DirectoryInfo(directory);
                if (root.Exists)
                {
                    AuthorizationRuleCollection ACLs_temp = root.GetAccessControl().GetAccessRules(true, true, typeof(System.Security.Principal.NTAccount));
                    AuthorizationRuleCollection ACLs = new AuthorizationRuleCollection();
                    foreach (FileSystemAccessRule ACL_temp in ACLs_temp)
                    {
                        if (checkBox2.Checked && !IsSystemUser(ACL_temp.IdentityReference.ToString()))
                        {
                            ACLs.AddRule(ACL_temp);//add only not system                            
                        }
                        else //not skip
                        {
                            ACLs.AddRule(ACL_temp);//add all
                        }
                    }
                    ACLs_temp = null;
                    if (ACLs.Count > 0)
                    {
                        foreach (FileSystemAccessRule ACL in ACLs)
                        {
                            string[] rights_Temp = ACL.FileSystemRights.ToString().Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (string right_Temp in rights_Temp)
                            {
                                string ACL_string = ACL.IdentityReference.ToString();
                                string ACL_IF = ACL.InheritanceFlags.ToString();
                                string ACL_IsI = ACL.IsInherited.ToString();
                                string right = right_Temp.Trim();
                                switch (right_Temp)//Печатаем большими буквами чтобы визуально отличать, но при сортировке и фильтре чтобы не отличалось
                                {
                                    case "268435456":
                                        right = "FULLCONTROLL";
                                        break;
                                    case "-536805376":
                                        right = "EXECUTE";
                                        break;
                                    case "-1073741824":
                                        right = "WRITE";
                                        break;
                                    case "2147483648":
                                        right = "READ";
                                        break;
                                    case "-1610612736"://https://coderoad.ru/26427967/Get-acl-%D0%BF%D0%BE%D0%B2%D1%82%D0%BE%D1%80%D1%8F%D1%8E%D1%89%D0%B8%D0%B5%D1%81%D1%8F-%D0%B3%D1%80%D1%83%D0%BF%D0%BF%D1%8B-Powershell
                                        right = "READANDEXECUTE";

                                        if (ACL.AccessControlType.ToString().Equals("Allow"))
                                        {
                                            //Invoke((ThreadStart)delegate{richTextBox3.Text += directory + ";" + ACL_string + ";" + "SYNCHRONIZE" + ";" + string.Empty + ";" + ACL_IF + ";" + ACL_IsI + Environment.NewLine;});
                                            Invoke((ThreadStart)delegate { listView1.Items.Add(new ListViewItem(new string[] { directory, ACL_string, "SYNCHRONIZE", string.Empty, ACL_IF, ACL_IsI })); });
                                            //Invoke((ThreadStart)delegate { Real_list.Add(new ListItem_my(new string[] { directory, ACL_string, "SYNCHRONIZE", String.Empty, ACL_IF, ACL_IsI })); });
                                            //Invoke((ThreadStart)delegate { listView1.Refresh(); });
                                        }
                                        if (ACL.AccessControlType.ToString().Equals("Deny"))
                                        {
                                            //Invoke((ThreadStart)delegate{richTextBox3.Text += directory + ";" + ACL_string + ";" + string.Empty + ";" + "SYNCHRONIZE" + ";" + ACL_IF + ";" + ACL_IsI + Environment.NewLine;});
                                            Invoke((ThreadStart)delegate{listView1.Items.Add(new ListViewItem(new string[] { directory, ACL_string, string.Empty, "SYNCHRONIZE", ACL_IF, ACL_IsI }));});
                                            //Invoke((ThreadStart)delegate { Real_list.Add(new ListItem_my(new string[] { directory, ACL_string, String.Empty, "SYNCHRONIZE", ACL_IF, ACL_IsI })); });
                                            //Invoke((ThreadStart)delegate { listView1.Refresh(); });                                            
                                        }
                                        break;
                                    default:
                                        break;
                                }
                                try
                                {
                                    if (ACL.AccessControlType.ToString().Equals("Allow"))
                                    {
                                        //Invoke((ThreadStart)delegate{richTextBox3.Text += directory + ";" + ACL_string + ";" + right + ";" + string.Empty + ";" + ACL_IF + ";" + ACL_IsI + Environment.NewLine;});
                                        Invoke((ThreadStart)delegate { listView1.Items.Add(new ListViewItem(new string[] { directory, ACL_string, right, string.Empty, ACL_IF, ACL_IsI })); });
                                        //Invoke((ThreadStart)delegate { Real_list.Add(new ListItem_my(new string[] { directory, ACL_string, right, String.Empty, ACL_IF, ACL_IsI })); });
                                        //Invoke((ThreadStart)delegate { listView1.Refresh(); });
                                    }
                                    if (ACL.AccessControlType.ToString().Equals("Deny"))
                                    {
                                        //Invoke((ThreadStart)delegate{richTextBox3.Text += directory + ";" + ACL_string + ";" + string.Empty + ";" + right + ";" + ACL_IF + ";" + ACL_IsI + Environment.NewLine;});
                                        Invoke((ThreadStart)delegate {listView1.Items.Add(new ListViewItem(new string[] { directory, ACL_string, string.Empty, right, ACL_IF, ACL_IsI })); });
                                        //Invoke((ThreadStart)delegate { Real_list.Add(new ListItem_my(new string[] { directory, ACL_string, String.Empty, right, ACL_IF, ACL_IsI })); });
                                        //Invoke((ThreadStart)delegate { listView1.Refresh(); });
                                    }
                                }
                                catch (UnauthorizedAccessException ee) { MessageBox.Show(ee.StackTrace); }
                                catch (Exception ee) { MessageBox.Show(ee.StackTrace); }
                            }
                        }
                    }
                    try
                    {
                        List<DirectoryInfo> subDirs = new List<DirectoryInfo>(root.GetDirectories());// рекурсией проверяем все подпапки
                        if (checkBox2.Checked == true)//remove system directories
                        {
                            foreach (string systemfolder in systemfolders)
                                foreach (DirectoryInfo subDir in subDirs)
                                    if (subDir.Name.ToString().Equals(systemfolder, StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        subDirs.Remove(subDir);
                                        break;
                                    }
                            foreach (string systemfolder in systemfolders)
                                foreach (DirectoryInfo subDir in subDirs)
                                    if (subDir.Attributes.ToString().Contains("System"))
                                    {
                                        subDirs.Remove(subDir);
                                        break;
                                    }
                        }

                        List<string> Tree = new List<string>(root.FullName.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries));
                        if (Tree.Count < root_layer + Convert.ToInt32(textBox2.Text))
                            foreach (DirectoryInfo dirInfo in subDirs)
                                if (allow_work)
                                    WalkDirectoryTree(dirInfo.FullName);
                    }
                    catch { Invoke((ThreadStart)delegate { listView1.Items.Add(new ListViewItem(new string[] { directory, "Acceess denied", string.Empty, string.Empty, string.Empty, string.Empty })); }); }
                }
                else
                {
                    MessageBox.Show("The folder does not exist. Canceling an operation.");
                }
            }
        }
        /// <summary>
        /// system users
        /// </summary>
        private string[] systemusers;
        /// <summary>
        /// Check is users is system users
        /// </summary>
        /// <param name="whattochek">Имя пользователя для сверки его со списком имен юзеров на предмет требуется ли исключить проверяемого юзера</param>
        /// <returns></returns>
        public bool IsSystemUser(string whattochek)
        {
            bool IsSystemRule = false;
            try
            {
                for (int i = 0; i < systemusers.Length; i++)
                    if (whattochek.Contains(systemusers[i]))
                        return true;
            }
            catch (Exception ww) { MessageBox.Show(ww.Message); }
            return IsSystemRule;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.ShowDialog();
            if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
                textBox1.Text = fbd.SelectedPath;
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.DefaultDirectory = textBox1.Text;
            Properties.Settings.Default.Upgrade();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            allow_work = false;
            button3.Enabled = false; button1.Enabled = true; button3.Visible = false;
        }

        private bool allow_work = false;

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void ListView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                DirectoryInfo DI = new DirectoryInfo(listView1.SelectedItems[0].SubItems[0].Text);
                if (DI.Exists)
                {
                    try
                    {
                        if (!DI.Attributes.ToString().Contains("ReparsePoint"))//проверка возможности доступа
                            Process.Start("explorer", listView1.SelectedItems[0].SubItems[0].Text);
                    }
                    catch { }
                }
            }
        }

        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://msdn.microsoft.com/ru-ru/library/system.security.accesscontrol.filesystemrights.aspx");
        }

        private void LinkLabel1_MouseEnter(object sender, EventArgs e)
        {
            linkLabel1.Text = "https://msdn.microsoft.com/ru-ru/library/system.security.accesscontrol.filesystemrights.aspx";
        }

        private void LinkLabel1_MouseLeave(object sender, EventArgs e)
        {
            linkLabel1.Text = "Access rights description";
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            if (listView1.Items.Count > 0)
            {
                btnExcel.Enabled = false;
                saveFileDialog1.Filter = "excel files (*.xls)|*.xls | csv files (*.csv)|*.csv";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    if (saveFileDialog1.FilterIndex == 1)
                    {
                        Excel.Application app = new Excel.Application();
                        Excel.Workbook wb = app.Workbooks.Add(1);
                        Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];
                        int i = 1;
                        int i2 = 1;
                        foreach (ListViewItem lvi in listView1.Items)
                        {
                            i = 1;
                            foreach (ListViewItem.ListViewSubItem lvs in lvi.SubItems)
                            {
                                ws.Cells[i2, i] = lvs.Text;
                                i++;
                            }
                            i2++;
                        }
                        wb.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
                        wb.Close(false, Type.Missing, Type.Missing);
                        app.Quit();
                    }
                    else if (saveFileDialog1.FilterIndex == 2)
                    {
                        StringBuilder sb = new StringBuilder();
                        foreach (ColumnHeader ch in listView1.Columns)
                        {
                            sb.Append(ch.Text + ",");
                        }
                        sb.AppendLine();
                        foreach (ListViewItem lvi in listView1.Items)
                        {
                            foreach (ListViewItem.ListViewSubItem lvs in lvi.SubItems)
                            {
                                if (lvs.Text.Trim() == string.Empty)
                                    sb.Append(" ,");
                                else
                                    sb.Append(lvs.Text + ",");
                            }
                            sb.AppendLine();
                        }
                        StreamWriter sw = new StreamWriter(saveFileDialog1.FileName);
                        sw.Write(sb.ToString());
                        sw.Close();
                    }
                }
                btnExcel.Enabled = true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
    class ListItem_my
    {
        public string[] one_item;
        public ListItem_my(string[] one_item)
        {
            this.one_item = one_item;
        }
    }
}
