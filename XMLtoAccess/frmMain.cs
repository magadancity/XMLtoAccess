using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ADOX;
//using ADODB;

namespace XMLtoAccess
{
    public partial class frmMain : Form
    {
        string dbName, pathToDb;
        List<String> tableListH = new List<string>();
        List<String> tableListL = new List<string>();
        Dictionary<string, List<string>> addFieldsH = new Dictionary<string, List<string>>();
        Dictionary<string, List<string>> addFieldsL = new Dictionary<string, List<string>>();

        public frmMain()
        {
            InitializeComponent();
            initData();
        }

        private void initData()
        {
            tableListH =new List<string>(new string[]{ "ZGLV","SCHET","ZAP","PACIENT","SLUCH","HMP","NAPR_FROM","CONS",
                "ONK_SL","B_DIAG","B_PROT","ONK_USL","LEK_PR","USL","SL_KOEF"});
            tableListL = new List<string>(new string[] { "ZGLV","PERS"});

            addFieldsH.Add("ZGLV", new List<string>(new string[] { "H"}));
            addFieldsH.Add("SCHET", new List<string>(new string[] { }));
            addFieldsH.Add("ZAP", new List<string>(new string[] { "CODE","PLAT","NSCHET" }));
            addFieldsH.Add("PACIENT", new List<string>(new string[] { "N_ZAP","PLAT" }));
            addFieldsH.Add("SLUCH", new List<string>(new string[] { "N_ZAP", "PLAT" }));
            addFieldsH.Add("HMP", new List<string>(new string[] { "IDCASE" }));
            addFieldsH.Add("NAPR_FROM", new List<string>(new string[] { "IDCASE" }));
            addFieldsH.Add("CONS", new List<string>(new string[] { "IDCASE"}));
            addFieldsH.Add("ONK_SL", new List<string>(new string[] { "IDCASE" }));
            addFieldsH.Add("B_DIAG", new List<string>(new string[] { "IDCASE" }));
            addFieldsH.Add("B_PROT", new List<string>(new string[] { }));
            addFieldsH.Add("ONK_USL", new List<string>(new string[] { "IDCASE" }));
            addFieldsH.Add("LEK_PR", new List<string>(new string[] { "IDCASE","USL_TIP" }));
            addFieldsH.Add("USL", new List<string>(new string[] { "N_ZAP", "IDCASE","PLAT"}));
            addFieldsH.Add("SL_KOEF", new List<string>(new string[] {"IDCASE" }));

            addFieldsL.Add("PERS", new List<string>(new string[] { "PLAT" }));
        }
        private bool createDB()
        {
            bool isres = false;

            ADOX.Catalog cat = new ADOX.Catalog();
            //ADOX.Table table = new ADOX.Table();

            ////Create the table and it's fields. 
            //table.Name = "Table1";
            //table.Columns.Append("Field1");
            //table.Columns.Append("Field2");

            try
            {
                dbName = getDBName();
                string pathToDb = txtPathToDB.Text + $@"{dbName}";
                string connString = $@"Provider=Microsoft.Jet.OLEDB.4.0; Data Source={pathToDb}; Jet OLEDB:Engine Type=5";
                //cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;" + @"Data Source=D:\data\text.mdb" + "; Jet OLEDB:Engine Type=5");
                cat.Create(connString);
                
                //cat.Tables.Append(table);

                //OleDbConnection con = new OleDbConnection();
                OleDbConnection con = cat.ActiveConnection as OleDbConnection;

                //Now Close the database
                //ADODB.Connection con = cat.ActiveConnection as ADODB.Connection;
                if (con != null)
                    con.Close();

                isres = true;
            }
            catch (Exception ex)
            {
                isres = false;
            }
            cat = null;
            return isres;
        }

        private void btnSelectXml_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Файлы архива zip|*.zip";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtPathToArc.Text = openFileDialog.FileName;
                    txtPathToDB.Text = Path.GetDirectoryName(openFileDialog.FileName) + @"\db\";
                }
                checkPath(txtPathToDB.Text);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void btnSelectPathToDb_Click(object sender, EventArgs e)
        {
            try
            {
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                fbd.ShowNewFolderButton = true;
                fbd.SelectedPath = txtPathToDB.Text;
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    txtPathToDB.Text = fbd.SelectedPath+@"\db\";
                }
                checkPath(fbd.SelectedPath);
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message);
            }

        }

        private string getDBName()
        {
            string res = "";
            string dt = DateTime.Now.ToString("yyyyMMdd");
            string tm = DateTime.Now.ToString("HHmm");
            res = $"RokbSamson_{dt}tm{tm}_fXML_in97.mdb";
            return res;
        }
        private void checkPath(string strPath)
        {
            string strCaption = "Сохранение файла в хранилище";
            if (!Directory.Exists(strPath))
            {
                DialogResult dr = MessageBox.Show("Рабочая папка по указанному пути отсутствует. Создать?",
                    strCaption, MessageBoxButtons.OKCancel);
                if (dr == DialogResult.OK)
                {
                    DirectoryInfo di = Directory.CreateDirectory(strPath);
                }
                else
                {
                    MessageBox.Show("Продолжение невозможно");
                }
            }
            else
            {
                if (MessageBox.Show("Путь существует. Заменить Файлы?") == DialogResult.OK)
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(strPath);

                    foreach (FileInfo file in dirInfo.GetFiles())
                    {
                        file.Delete();
                    }
                }
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            txtLog.Clear();
            txtLog.AppendLine("Разархивирование");
            if (!@unZip()) { MessageBox.Show("Ошибка разархивирования файлов"); return; }
            txtLog.AppendLine("Разархивирование закончено");
            txtLog.AppendLine("Создание базы данных");
            if (!createDB()) 
            {
                MessageBox.Show("Ошибка создания БД");
                txtLog.AppendLine("Ошибка создания БД");
            }
            else
            {
                txtLog.AppendLine($@"База данных создана, расположение {txtPathToDB.Text}{dbName}");
            }
            txtLog.AppendLine("Чтение XML файлов");
            readXMLFiles();
        }

        private bool unZip()
        {
            bool isres = false;
            try
            {
                if (!File.Exists(txtPathToArc.Text)) { return isres; }
                String dir = Path.GetDirectoryName(txtPathToArc.Text);
                String unzippath = dir + @"\unarc\";
                checkPath(unzippath);
                ZipFile.ExtractToDirectory(txtPathToArc.Text, unzippath);
                isres = true;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return isres;
            //string startPath = @".\start";
            //string zipPath = @".\result.zip";
            //string extractPath = @".\extract";

            //ZipFile.CreateFromDirectory(startPath, zipPath);

            //ZipFile.ExtractToDirectory(zipPath, extractPath);
        }

        private bool readXMLFiles()
        {
            bool isres = false;
            try
            {
                DirectoryInfo dir = new DirectoryInfo($@"{Path.GetDirectoryName(txtPathToArc.Text)}\unarc");
                FileInfo[] files = dir.GetFiles();
                if (files.Length == 0)
                {
                    string strMsg = "Файлы для обработки отсутствуют";
                    MessageBox.Show(strMsg);
                    txtLog.AppendLine(strMsg);
                    return isres;
                }

                foreach(FileInfo fi in files)
                {
                    txtLog.AppendLine($"Обрабатывается файл {fi.Name}");
                    if (!ImportFilesToDatabase(fi))
                    {
                        string strMsg = $"Ошибка импорта файла {fi.Name}";
                        MessageBox.Show(strMsg);
                        txtLog.AppendLine(strMsg);
                        return isres;
                    }
                }
                isres = false;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return isres;
        }
        public bool ImportFilesToDatabase(FileInfo fi)
        {
            bool result = false;
            string strMes = "";
            try
            {
                List<string> ignoretables = new List<string>() { "DS2", "DS3" };

                DataSet ds = new DataSet();
                
                ds.ReadXml(fi.FullName);

                if (!ds.Tables.Contains("ZGLV"))
                {
                    strMes = "Отсутствует обязательная таблица ZGLV";
                    MessageBox.Show(strMes);
                    txtLog.AppendLine(strMes);
                    return result;
                }

                string fileName = ds.Tables["ZGLV"].Rows[0]["ZGLV"].ToString();
                if (fileName.StartsWith("H"))
                {
                    result = hFile(ds);
                }
                else if (fileName.StartsWith("L"))
                {
                    result = lFile(ds);
                }
                else
                {
                    strMes = $"Неизвестный тип файла {fileName.Substring(0,1)}";
                    MessageBox.Show(strMes);
                    txtLog.AppendLine(strMes);
                    return result;
                }
                //if (!AddColumnsToTables()) { return result; }

                //if (!CheckTables()) { return result; }

                //if (!InsertData()) { return result; }

                result = true;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            return result;
        }

        private bool hFile(DataSet ds)
        {
            bool result = false;
            try
            {

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return result;
        }

        private bool lFile(DataSet ds)
        {
            bool result = false;
            try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return result;
        }

    }
}
