using System;
using System.Diagnostics;
using System.IO;
using System.Security.AccessControl;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Threading;
using System.Collections.Generic;

namespace FolderRules
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            lvwColumnSorter = new ListViewColumnSorter();
            this.listView1.ListViewItemSorter = lvwColumnSorter;
            systemusers = richTextBox1.Text.Split('\n');

            textBox1.Text = Properties.Settings.Default.DefaultDirectory;
            textBox2.Text = Properties.Settings.Default.scan_dept.ToString();
            richTextBox1.Text = Properties.Settings.Default.System_folders;
            richTextBox2.Text = Properties.Settings.Default.System_users;

            systemfolders = richTextBox1.Text.Split('\n');
            systemusers = richTextBox2.Text.Split('\n');
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
            Task.Factory.StartNew(() => Main_Task(textBox1.Text));
        }
        private void Main_Task(string RootDirectory)
        {
            WalkDirectoryTree(RootDirectory);
            Invoke((ThreadStart)delegate
            {
                button1.Enabled = true;
                button3.Enabled = false;
                button3.Visible = false;
            });
        }
        /// <summary>
        /// Exporting permissions of directory 
        /// </summary>
        /// <param name="RootDirectory">Папка права которой требуется вычислить</param>
        private void WalkDirectoryTree(string RootDirectory)
        {
            if (allow_work)
            {
                // Проверяем текущую папку, проверка подпапок в конце функции
                DirectoryInfo root = new DirectoryInfo(RootDirectory);
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
                            string[] rights = ACL.FileSystemRights.ToString().Split(',');
                            if (ACL.AccessControlType.ToString().Equals("Allow"))
                            {
                                foreach (string right in rights)
                                {
                                    try
                                    {
                                        string ACL_string = ACL.IdentityReference.ToString();
                                        string right_TTT = right.Trim();
                                        switch (right)
                                        {
                                            case "268435456":
                                                right_TTT = "FULLCONTROLL";
                                                break;
                                            case "-536805376":
                                                right_TTT = "EXECUTE";
                                                break;
                                            case "-1073741824":
                                                right_TTT = "WRITE";
                                                break;
                                            case "2147483648":
                                                right_TTT = "READ";
                                                break;
                                            case "-1610612736"://https://coderoad.ru/26427967/Get-acl-%D0%BF%D0%BE%D0%B2%D1%82%D0%BE%D1%80%D1%8F%D1%8E%D1%89%D0%B8%D0%B5%D1%81%D1%8F-%D0%B3%D1%80%D1%83%D0%BF%D0%BF%D1%8B-Powershell
                                                right_TTT = "READANDEXECUTE";
                                                Invoke((ThreadStart)delegate { listView1.Items.Add(new ListViewItem(new string[] { RootDirectory, ACL_string, "SYNCHRONIZE", String.Empty })); });
                                                break;
                                            default:
                                                break;
                                        }
                                        Invoke((ThreadStart)delegate { listView1.Items.Add(new ListViewItem(new string[] { RootDirectory, ACL_string, right_TTT, String.Empty })); });
                                    }
                                    catch { }
                                }
                            }
                            else if (ACL.AccessControlType.ToString().Equals("Deny"))
                            {
                                foreach (string right in rights)
                                {
                                    Invoke((ThreadStart)delegate { listView1.Items.Add(new ListViewItem(new string[] { RootDirectory, ACL.IdentityReference.ToString(), String.Empty, right })); });
                                }
                            }
                        }
                    }

                    // Теперь рекурсией проверяем все подпапки
                    List<DirectoryInfo> subDirs = new List<DirectoryInfo>(root.GetDirectories());
                    if (checkBox2.Checked == true)//remove system directories
                        foreach (string systemfolder in systemfolders)
                        {
                            foreach (DirectoryInfo subDir in subDirs)
                            {
                                if (subDir.Name.Equals(systemfolder))
                                {
                                    subDirs.Remove(subDir);
                                    break;
                                }
                            }
                        }
                    string[] Tree = root.FullName.Split('\\');
                    
                    if (Tree.Length < Convert.ToInt32(Properties.Settings.Default.scan_dept) + 1)
                    {
                        foreach (DirectoryInfo dirInfo in subDirs)
                        {
                            if (allow_work)
                                WalkDirectoryTree(dirInfo.FullName);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("The folder does not exist. Canceling an operation.");
                }
                Invoke((ThreadStart)delegate { listView1.Refresh(); });
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
                {
                    if (whattochek.Contains(systemusers[i]))
                    {
                        return true;
                    }
                }
            }
            catch (Exception ww)
            { MessageBox.Show(ww.Message); }
            return IsSystemRule;
        }

        private void LinkLabel1_Click(object sender, EventArgs e)
        {
            Process.Start("https://msdn.microsoft.com/ru-ru/library/system.security.accesscontrol.filesystemrights.aspx");
        }

        private void LinkLabel1_MouseEnter(object sender, EventArgs e)
        {
            linkLabel1.Text = "https://msdn.microsoft.com/ru-ru/library/system.security.accesscontrol.filesystemrights.aspx";
        }

        private void LinkLabel1_MouseLeave(object sender, EventArgs e)
        {
            linkLabel1.Text = "Описание прав доступа";
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();
            if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
                textBox1.Text = fbd.SelectedPath;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

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

        private void linkLabel1_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string s = listView1.Items[listView1.SelectedIndices[1]].Text;
            //System.Diagnostics.Process.Start("explorer", listView1.Items[listView1.SelectedIndices[1]].);
        }
        private ListViewColumnSorter lvwColumnSorter;
        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            //https://docs.microsoft.com/ru-ru/troubleshoot/dotnet/csharp/sort-listview-by-column
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            this.listView1.Sort();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.System_folders = richTextBox1.Text;
            Properties.Settings.Default.System_users = richTextBox2.Text;
            Properties.Settings.Default.scan_dept = Convert.ToInt32(textBox2.Text);
            Properties.Settings.Default.Ignore_system_users = checkBox2.Checked;
            Properties.Settings.Default.Ignore_system_directories = checkBox1.Checked;
            Properties.Settings.Default.Save();
            Properties.Settings.Default.Reload();
        }


        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.Save();
            Properties.Settings.Default.Reload();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
