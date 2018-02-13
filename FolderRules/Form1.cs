using System;
using System.Diagnostics;
using System.IO;
using System.Security.AccessControl;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Threading;

namespace FolderRules
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            allow_work = true;
            button3.Visible = true;
            richTextBox1.Clear();
            richTextBox1.Text += "Имя папки,\t,Имя учетной записи,\tРазрешено,\tЗапрещено" + Environment.NewLine;
            string[] drives = Environment.GetLogicalDrives();
            Task.Factory.StartNew(() => WalkDirectoryTree(textBox1.Text));
        }
        //List<string> filesanddirecoties;
        //DirectoryInfo[] files;
        private void WalkDirectoryTree(string RootDirectory)
        {
            DirectoryInfo root = new DirectoryInfo(RootDirectory);
            if (root.Exists)
            {
                try
                {
                    //filesanddirecoties = new List<string>();
                    DirectoryInfo[] subDirs = root.GetDirectories();
                    AuthorizationRuleCollection ACLs = root.GetAccessControl().GetAccessRules(true, true, typeof(System.Security.Principal.NTAccount));
                    string[] Tree = root.FullName.Split('\\');

                    //строим карту папок и файлов
                    /* for (int i = 0; i < Tree.Length; i++)
                     {
                         filesanddirecoties.Add(Tree[i]);
                         string[] files = Directory.GetFiles(RootDirectory);
                         for (int j = 0; j < files.Length; j++)
                         {
                             if(files.Length!=0)
                                 filesanddirecoties.Add(files[j]);
                         }
                     }*/
                    //получаем права на папку
                    foreach (FileSystemAccessRule ACL in ACLs)
                    {
                        if (!IsSystemRules(ACL.IdentityReference.ToString()))
                        {
                            if (allow_work)
                            {
                                if (ACL.AccessControlType.ToString().Equals("Allow"))
                                {
                                    Invoke((ThreadStart)delegate { richTextBox1.Text += RootDirectory + ",\t" + ACL.IdentityReference + ",\t" + ACL.FileSystemRights + Environment.NewLine; });
                                }
                                else if (ACL.AccessControlType.ToString().Equals("Deny"))
                                {
                                    Invoke((ThreadStart)delegate { richTextBox1.Text += RootDirectory + ",\t" + ACL.IdentityReference + ",\t,\t" + ACL.FileSystemRights + Environment.NewLine; });
                                }
                            }
                        }
                    }
                  
                    if (Tree.Length < Convert.ToInt32(textBox2.Text) + 2)
                    {
                        foreach (DirectoryInfo dirInfo in subDirs)
                        {
                            if (allow_work)
                                WalkDirectoryTree(dirInfo.FullName);
                        }
                    }
                }
                catch (UnauthorizedAccessException ee)
                {
                    Invoke((ThreadStart)delegate { richTextBox1.Text += RootDirectory + ",\tAccess Denied" + Environment.NewLine /*+ ee.StackTrace*/; });
                }
                catch (Exception ww) { MessageBox.Show(ww.StackTrace); }
            }
            else
            {
                Invoke((ThreadStart)delegate { richTextBox1.Text = "Папка не существует. Отмена операции."; });
            }
        }

        private readonly string[] systemrules = new string[] { "СОЗДАТЕЛЬ-ВЛАДЕЛЕЦ", "BUILTIN", "NT AUTHORITY" };
        public bool IsSystemRules(string whattochek)
        {
            bool IsSystemRule = false;
            for (int i = 0; i < systemrules.Length; i++)
            {
                if (whattochek.Contains(systemrules[i]))
                {
                    return IsSystemRule = true;
                }
            }

            return IsSystemRule;
        }

        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

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
            textBox1.Text = Properties.Settings.Default.DefaultDirectory;
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.DefaultDirectory = textBox1.Text;
            Properties.Settings.Default.Upgrade();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            allow_work = false;
            button3.Visible = false;
        }

        private bool allow_work = false;
    }
}
