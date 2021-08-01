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
            systemusers = richTextBox1.Text.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);

            textBox1.Text = Properties.Settings.Default.DefaultDirectory;
            textBox2.Text = Properties.Settings.Default.scan_dept.ToString();
            richTextBox1.Text = Properties.Settings.Default.System_folders;
            richTextBox2.Text = Properties.Settings.Default.System_users;

            systemfolders = richTextBox1.Text.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            systemusers = richTextBox2.Text.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
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
                            string[] rights_Temp = ACL.FileSystemRights.ToString().Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (string right_Temp in rights_Temp)
                            {
                                try
                                {
                                    string ACL_string = ACL.IdentityReference.ToString();
                                    string right = right_Temp.Trim();//ос
                                    switch (right_Temp)//Печатаем большими буквами чтобы визуально отличать,, но при сортировке и фильтре чтобы не отличалось
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
                                            try
                                            {
                                                if (ACL.AccessControlType.ToString().Equals("Allow"))
                                                {
                                                    Invoke((ThreadStart)delegate { listView1.Items.Add(new ListViewItem(new string[] { RootDirectory, ACL_string, "SYNCHRONIZE", String.Empty, ACL.InheritanceFlags.ToString(), ACL.IsInherited.ToString() })); });
                                                }
                                                if (ACL.AccessControlType.ToString().Equals("Deny"))
                                                {
                                                    Invoke((ThreadStart)delegate { listView1.Items.Add(new ListViewItem(new string[] { RootDirectory, ACL_string, String.Empty, "SYNCHRONIZE", ACL.InheritanceFlags.ToString(), ACL.IsInherited.ToString() })); });
                                                }
                                            }
                                            catch (Exception EE) { MessageBox.Show("1\n" + EE.Message); }
                                            break;
                                        default:
                                            break;
                                    }
                                    try
                                    {
                                        if (ACL.AccessControlType.ToString().Equals("Allow"))
                                        {
                                            Invoke((ThreadStart)delegate { listView1.Items.Add(new ListViewItem(new string[] { RootDirectory, ACL_string, right, String.Empty, ACL.InheritanceFlags.ToString(), ACL.IsInherited.ToString() })); });
                                        }
                                        if (ACL.AccessControlType.ToString().Equals("Deny"))
                                        {
                                            Invoke((ThreadStart)delegate { listView1.Items.Add(new ListViewItem(new string[] { RootDirectory, ACL.IdentityReference.ToString(), String.Empty, right, ACL.InheritanceFlags.ToString(), ACL.IsInherited.ToString() })); });
                                        }
                                    }
                                    catch (Exception EE) { MessageBox.Show("2\n" + EE.Message); }
                                }
                                catch (Exception EE) { MessageBox.Show("3\n" + EE.Message); }
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
                    List<string> Tree = new List<string>(root.FullName.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries));              
                    if (Tree.Count < Convert.ToInt32(Properties.Settings.Default.scan_dept))
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
                try
                {
                    Invoke((ThreadStart)delegate { listView1.Refresh(); });
                }catch(Exception EE) { MessageBox.Show("5\n" + EE.Message); }
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
            if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
                textBox1.Text = fbd.SelectedPath;
            }
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

        private ListViewColumnSorter lvwColumnSorter;
        /// <summary>
        /// Отсортировать по Listview при нажатии на заголовок столба
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListView1_ColumnClick(object sender, ColumnClickEventArgs e)
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
            listView1.Sort();
        }

        private void Button4_Click(object sender, EventArgs e)
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

        private void TextBox2_TextChanged(object sender, EventArgs e)
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
                        {
                            Process.Start("explorer", listView1.SelectedItems[0].SubItems[0].Text);
                        }
                    }
                    catch { }
                }
            }
        }
    }
}
