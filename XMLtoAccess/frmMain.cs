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
//using System.Threading.Tasks;
using System.Diagnostics;
using ADOX;
using ADODB;
using DAO;

namespace XMLtoAccess
{
    public partial class frmMain : Form
    {
        string dbName, pathToDb, connString;
        List<String> tableListH = new List<string>();
        List<String> tableListL = new List<string>();
        Dictionary<string, List<string>> addFieldsH = new Dictionary<string, List<string>>();
        Dictionary<string, List<string>> addFieldsL = new Dictionary<string, List<string>>();
        Dictionary<string, OleDbCommand> commands = new Dictionary<string, OleDbCommand>();
        string _PLAT;
        public frmMain()
        {
            InitializeComponent();
            initData();
        }

        private void initData()
        {
            tableListH =new List<string>(new string[]{ "ZGLV","SCHET","ZAP","PACIENT","SLUCH","HMP","NAPR_FROM","CONS",
                "ONK_SL","B_DIAG","B_PROT","ONK_USL","LEK_PR","USL","SL_KOEF"});
            tableListL = new List<string>(new string[] { "PERS"});

            addFieldsH.Add("ZGLV", new List<string>(new string[] { "H","FILENAME1"}));
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
            addFieldsH.Add("LEK_PR", new List<string>(new string[] { "IDCASE","USL_TIP","DATE_INJ" }));
            addFieldsH.Add("USL", new List<string>(new string[] { "N_ZAP", "IDCASE","PLAT"}));
            addFieldsH.Add("SL_KOEF", new List<string>(new string[] {"IDCASE" }));

            addFieldsL.Add("PERS", new List<string>(new string[] { "PLAT" }));
        }
        private bool createDB()
        {
            bool isres = false;

            ADOX.Catalog cat = new ADOX.Catalog();
            try
            {
                dbName = getDBName();
                pathToDb = txtPathToDB.Text + $@"{dbName}";
                connString = $@"Provider=Microsoft.Jet.OLEDB.4.0; Data Source={pathToDb}; Jet OLEDB:Engine Type=5";
                cat.Create(connString);

                OleDbConnection con = cat.ActiveConnection as OleDbConnection;

                if (con != null)
                    con.Close();

                isres = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
            string strMsg = "";
            try
            {
                txtLog.AppendLine(DateTime.Now.ToString());
                DirectoryInfo dir = new DirectoryInfo($@"{Path.GetDirectoryName(txtPathToArc.Text)}\unarc");
                FileInfo[] files = dir.GetFiles();
                if (files.Length == 0)
                {
                    strMsg = "Файлы для обработки отсутствуют";
                    MessageBox.Show(strMsg);
                    txtLog.AppendLine(strMsg);
                    return isres;
                }

                foreach(FileInfo fi in files)
                {
                    txtLog.AppendLine($"Обрабатывается файл {fi.Name}");
                    if (!ImportFilesToDatabase(fi))
                    {
                        strMsg = $"Ошибка импорта файла {fi.Name}";
                        MessageBox.Show(strMsg);
                        txtLog.AppendLine(strMsg);
                        return isres;
                    }
                }

                strMsg = "Обработка файла закончена";
                txtLog.AppendLine(strMsg);
                txtLog.AppendLine(DateTime.Now.ToString());
                MessageBox.Show(strMsg);

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
                int indxUnder = fi.Name.IndexOf('_');
                int indxType = 0;
                if (fi.Name.Contains("S")) { indxType = fi.Name.IndexOf('S'); }
                else if (fi.Name.Contains("T")) { indxType = fi.Name.IndexOf('S'); }
                else { txtLog.AppendLine($"Не известный тип файла {fi.Name}"); return result; }
                _PLAT=
                DataSet ds = new DataSet();
                
                ds.ReadXml(fi.FullName);

                if (!ds.Tables.Contains("ZGLV"))
                {
                    strMes = "Отсутствует обязательная таблица ZGLV";
                    MessageBox.Show(strMes);
                    txtLog.AppendLine(strMes);
                    return result;
                }

                string fileName = ds.Tables["ZGLV"].Rows[0]["FILENAME"].ToString();
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

                result = true;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            return result;
        }

        private bool hFile(DataSet ds)
        {
            bool result = false;
            string strMes = "";
            OleDbConnection conn;
            ADODB.Connection adodbCon = new ADODB.Connection();
            adodbCon.ConnectionString = connString;
            ADOX.Catalog cat = new ADOX.Catalog();
            ADOX.Table tab;
            int SLUCH_ID = 0;
            DAO.DBEngine dbEngine = new DAO.DBEngine();
            try
            {
                adodbCon.Open();
                cat.ActiveConnection = adodbCon;
                conn = new OleDbConnection(connString);

                conn.Open();
                if (conn.State != ConnectionState.Open)
                {
                    strMes = "Отсутствует соединение к БД";
                    MessageBox.Show(strMes);
                    txtLog.AppendLine(strMes);
                    return result;
                }

                txtLog.AppendLine("Создание структуры БД");
                Application.DoEvents();
                
                foreach (string tabName in tableListH)
                {
                    if (!ds.Tables.Contains(tabName))
                    {
                        strMes = $"Отсутствует таблица {tabName}";
                        MessageBox.Show(strMes);
                        txtLog.AppendLine(strMes);
                        continue;
                    }

                    string[] columnNames = ds.Tables[tabName]
                        .Columns.Cast<DataColumn>()
                        .Select(x => x.ColumnName)
                        .ToArray();

                    tab = new ADOX.Table();
                    tab.Name = $"{tabName}rokb";
                    //id
                    ADOX.Column column = new ADOX.Column();
                    column.Name = "id";
                    column.Type = ADOX.DataTypeEnum.adInteger;
                    column.ParentCatalog = cat;
                    column.Properties["AutoIncrement"].Value = true;
                    tab.Columns.Append(column);

                    foreach(string str in addFieldsH[tabName])
                    {
                        tab.Columns.Append(defCol(str));
                    }
                    foreach(string str in columnNames)
                    {
                        tab.Columns.Append(defCol(str));
                    }
                    cat.Tables.Append(tab);
                }

                conn.Close();

                txtLog.AppendLine("Внесение данных");
                Application.DoEvents();
                //внесение данных
                string _CODE = "", _NSCHET = "", _N_ZAP = "", _IDCASE = "", _USL_TIP = "2";
                txtLog.AppendLine("Таблица ZGLV");
                Application.DoEvents();
                DAO.Database db = dbEngine.OpenDatabase(pathToDb);
                DAO.Recordset rs = db.OpenRecordset("ZGLVrokb");
                foreach (DataRow dr in ds.Tables["ZGLV"].Rows)
                {
                    rs.AddNew();
                    rs.Fields["H"].Value = "H";
                    foreach (DataColumn dc in ds.Tables["ZGLV"].Columns)
                    {
                        rs.Fields[dc.ColumnName].Value=dr[dc.ColumnName].ToString();
                    }
                    rs.Update();
                }
                rs.Close();
                //
                //SCHET
                //
                txtLog.AppendLine("Таблица SCHET");
                Application.DoEvents();
                rs = db.OpenRecordset("SCHETrokb");
                DataRow drSchet = ds.Tables["SCHET"].Rows[0];
                _CODE = drSchet["CODE"].ToString();
                _NSCHET = drSchet["NSCHET"].ToString();
                rs.AddNew();
                foreach (DataColumn dc in ds.Tables["SCHET"].Columns)
                {
                    rs.Fields[dc.ColumnName].Value=drSchet[dc.ColumnName].ToString();
                }
                rs.Update();
                rs.Close();
                //
                //ZAP
                //
                txtLog.AppendLine("Таблица ZAP");
                Application.DoEvents();
                pb.Minimum = 0;
                pb.Maximum = ds.Tables["ZAP"].Rows.Count;
                pb.Value = 0;
                int counter = 0;
                rs = db.OpenRecordset("ZAProkb");
                foreach (DataRow dr in ds.Tables["ZAP"].Rows)
                {
                    pb.Value = counter++;
                    Application.DoEvents();
                    rs.AddNew();
                    _N_ZAP = dr["N_ZAP"].ToString();
                    int ZAP_Id = int.Parse(dr["ZAP_Id"].ToString());
                    rs.Fields["CODE"].Value = _CODE;
                    rs.Fields["PLAT"].Value = _PLAT;
                    rs.Fields["NSCHET"].Value = _NSCHET;
                    foreach (DataColumn dc in ds.Tables["ZAP"].Columns)
                    {
                        rs.Fields[dc.ColumnName].Value = dr[dc.ColumnName].ToString();
                    }
                    rs.Update();
                    //
                    //PACIENT
                    //
                    DAO.Recordset rsPacient = db.OpenRecordset("PACIENTrokb");
                    List<DataRow> pacList = ds.Tables["PACIENT"].AsEnumerable().Where(m => m.Field<Int32>("ZAP_Id") == ZAP_Id).ToList<DataRow>();
                    foreach (DataRow drPac in pacList)
                    {
                        rsPacient.AddNew();
                        rsPacient.Fields["N_ZAP"].Value = _N_ZAP;
                        rsPacient.Fields["PLAT"].Value = _PLAT;
                        foreach (DataColumn dc in ds.Tables["PACIENT"].Columns)
                        {
                            rsPacient.Fields[dc.ColumnName].Value = drPac[dc.Ordinal].ToString();
                        }
                        rsPacient.Update();
                    }
                    rsPacient.Close();
                    //
                    //SLUCH
                    //
                    DAO.Recordset rsSLUCH = db.OpenRecordset("SLUCHrokb");
                    List<DataRow> sluchList = ds.Tables["SLUCH"].AsEnumerable().Where(m => m.Field<Int32>("ZAP_Id") == ZAP_Id).ToList<DataRow>();
                    foreach (DataRow drSluch in sluchList)
                    {
                        _IDCASE = drSluch[ds.Tables["SLUCH"].Columns["IDCASE"].Ordinal].ToString();
                        SLUCH_ID = int.Parse(drSluch["SLUCH_Id"].ToString());
                        if (SLUCH_ID == 1098)
                        {
                            int zz = 0;
                        }
                        int SLUCH_Id = int.Parse(drSluch[ds.Tables["SLUCH"].Columns["SLUCH_Id"].Ordinal].ToString());
                        rsSLUCH.AddNew();
                        rsSLUCH.Fields["N_ZAP"].Value = _N_ZAP;
                        rsSLUCH.Fields["PLAT"].Value = _PLAT;
                        foreach (DataColumn dc in ds.Tables["SLUCH"].Columns)
                        {
                            rsSLUCH.Fields[dc.ColumnName].Value = drSluch[dc.Ordinal].ToString();
                        }
                        rsSLUCH.Update();
                        //
                        //HMP
                        //
                        DAO.Recordset rsHMP = db.OpenRecordset("HMProkb");
                        List<DataRow> hmpList = ds.Tables["HMP"].AsEnumerable().Where(m => m.Field<Int32>("SLUCH_Id") == SLUCH_Id).ToList<DataRow>();
                        if (hmpList != null && hmpList.Count > 0)
                        {
                            foreach (DataRow drHMP in hmpList)
                            {
                                rsHMP.AddNew();
                                rsHMP.Fields["IDCASE"].Value = _IDCASE;
                                foreach (DataColumn dc in ds.Tables["HMP"].Columns)
                                {
                                    rsHMP.Fields[dc.ColumnName].Value = drHMP[dc.Ordinal].ToString();
                                }
                                rsHMP.Update();
                            }
                        }
                        rsHMP.Close();
                        //
                        //NAPR_FROM
                        //
                        DAO.Recordset rsNaprFrom = db.OpenRecordset("NAPR_FROMrokb");
                        List<DataRow> naprFromList = ds.Tables["NAPR_FROM"].AsEnumerable().Where(m => m.Field<Int32>("SLUCH_Id") == SLUCH_Id).ToList<DataRow>();
                        if (naprFromList != null && naprFromList.Count > 0)
                        {
                            foreach (DataRow drNaprFrom in naprFromList)
                            {
                                rsNaprFrom.AddNew();
                                rsNaprFrom.Fields["IDCASE"].Value = _IDCASE;
                                foreach (DataColumn dc in ds.Tables["NAPR_FROM"].Columns)
                                {
                                    rsNaprFrom.Fields[dc.ColumnName].Value = drNaprFrom[dc.Ordinal].ToString();
                                }
                                rsNaprFrom.Update();
                            }
                        }
                        rsNaprFrom.Close();
                        //
                        //CONS
                        //
                        DAO.Recordset rsCons = db.OpenRecordset("CONSrokb");
                        List<DataRow> consList = ds.Tables["CONS"].AsEnumerable().Where(m => m.Field<Int32>("SLUCH_Id") == SLUCH_Id).ToList<DataRow>();
                        if (consList != null && consList.Count > 0)
                        {
                            foreach (DataRow drCons in consList)
                            {
                                rsCons.AddNew();
                                rsCons.Fields["IDCASE"].Value = _IDCASE;
                                foreach (DataColumn dc in ds.Tables["CONS"].Columns)
                                {
                                    rsCons.Fields[dc.ColumnName].Value = drCons[dc.Ordinal].ToString();
                                }
                                rsCons.Update();
                            }
                        }
                        rsCons.Close();
                        //
                        //ONK_SL
                        //
                        DAO.Recordset rsOnkSl = db.OpenRecordset("ONK_SLrokb");
                        List<DataRow> onkSlList = ds.Tables["ONK_SL"].AsEnumerable().Where(m => m.Field<Int32>("SLUCH_Id") == SLUCH_Id).ToList<DataRow>();
                        if (onkSlList != null && onkSlList.Count > 0)
                        {
                            foreach (DataRow drOnkSl in onkSlList)
                            {
                                int OnkSl_Id = int.Parse(drOnkSl[ds.Tables["ONK_SL"].Columns["ONK_SL_Id"].Ordinal].ToString());
                                rsOnkSl.AddNew();
                                rsOnkSl.Fields["IDCASE"].Value = _IDCASE;
                                foreach (DataColumn dc in ds.Tables["ONK_SL"].Columns)
                                {
                                    rsOnkSl.Fields[dc.ColumnName].Value = drOnkSl[dc.Ordinal].ToString();
                                }
                                rsOnkSl.Update();
                                //
                                //B_DIAG
                                //
                                DAO.Recordset rsBDiag = db.OpenRecordset("B_DIAGrokb");
                                List<DataRow> bDiagList = ds.Tables["B_DIAG"].AsEnumerable().Where(m => m.Field<Int32>("ONK_SL_Id") == OnkSl_Id).ToList<DataRow>();
                                if (bDiagList != null && bDiagList.Count > 0)
                                {
                                    foreach (DataRow drBDiag in bDiagList)
                                    {
                                        rsBDiag.AddNew();
                                        rsBDiag.Fields["IDCASE"].Value = _IDCASE;
                                        foreach (DataColumn dc in ds.Tables["B_DIAG"].Columns)
                                        {
                                            rsBDiag.Fields[dc.ColumnName].Value = drBDiag[dc.Ordinal].ToString();
                                        }
                                        rsBDiag.Update();
                                    }
                                }
                                rsBDiag.Close();
                                //
                                //ONK_USL
                                //
                                DAO.Recordset rsOnkUsl = db.OpenRecordset("ONK_USLrokb");
                                List<DataRow> onkUslList = ds.Tables["ONK_USL"].AsEnumerable().Where(m => m.Field<Int32>("ONK_SL_Id") == OnkSl_Id).ToList<DataRow>();
                                if (onkUslList != null && onkUslList.Count > 0)
                                {
                                    foreach (DataRow drOnkUsl in onkUslList)
                                    {
                                        rsOnkUsl.AddNew();
                                        rsOnkUsl.Fields["IDCASE"].Value = _IDCASE;
                                        int ONK_USL_Id = int.Parse(drOnkUsl["ONK_USL_Id"].ToString());
                                        foreach (DataColumn dc in ds.Tables["ONK_USL"].Columns)
                                        {
                                            rsOnkUsl.Fields[dc.ColumnName].Value = drOnkUsl[dc.Ordinal].ToString();
                                        }
                                        rsOnkUsl.Update();
                                        //
                                        //LEK_PR
                                        //
                                        DAO.Recordset rsLekPr = db.OpenRecordset("LEK_PRrokb");
                                        List<DataRow> lekPrList = ds.Tables["LEK_PR"].AsEnumerable().Where(m => m.Field<Int32>("ONK_USL_Id") == ONK_USL_Id).ToList<DataRow>();
                                        if (lekPrList != null && lekPrList.Count > 0)
                                        {
                                            foreach (DataRow drLekPr in lekPrList)
                                            {
                                                int LEK_PR_Id = int.Parse(drLekPr["LEK_PR_Id"].ToString());
                                                List<DataRow> dateInjList = ds.Tables["DATE_INJ"].AsEnumerable().Where(m => m.Field<Int32>("LEK_PR_Id") == LEK_PR_Id).ToList<DataRow>();
                                                if (dateInjList != null && dateInjList.Count > 0)
                                                {
                                                    foreach (DataRow drDateInj in dateInjList)
                                                    {
                                                        rsLekPr.AddNew();
                                                        foreach (DataColumn dc in ds.Tables["LEK_PR"].Columns)
                                                        {
                                                            rsLekPr.Fields[dc.ColumnName].Value = drLekPr[dc.Ordinal].ToString();
                                                        }
                                                        rsLekPr.Fields["IDCASE"].Value = _IDCASE;
                                                        rsLekPr.Fields["USL_TIP"].Value = "2";
                                                        rsLekPr.Fields["DATE_INJ"].Value = drDateInj["DATE_INJ_Text"].ToString();
                                                        rsLekPr.Update();
                                                    }
                                                }
                                                else
                                                {
                                                    rsLekPr.AddNew();
                                                    rsLekPr.Fields["IDCASE"].Value = _IDCASE;
                                                    rsLekPr.Fields["USL_TIP"].Value = "2";
                                                    rsLekPr.Update();
                                                }
                                            }
                                        }
                                        rsLekPr.Close();
                                    }
                                }
                                rsOnkUsl.Close();
                            }
                        }
                        rsOnkSl.Close();
                        //
                        //USL
                        //
                        DAO.Recordset rsUSL = db.OpenRecordset("USLrokb");
                        List<DataRow> uslList = ds.Tables["USL"].AsEnumerable().Where(m => m.Field<Int32>("SLUCH_Id") == SLUCH_Id).ToList<DataRow>();
                        if (uslList != null && uslList.Count > 0)
                        {
                            foreach (DataRow drUSL in uslList)
                            {
                                int USL_Id = int.Parse(drUSL["USL_Id"].ToString());
                                rsUSL.AddNew();
                                rsUSL.Fields["IDCASE"].Value = _IDCASE;
                                rsUSL.Fields["N_ZAP"].Value = _N_ZAP;
                                rsUSL.Fields["PLAT"].Value = _PLAT;
                                foreach (DataColumn dc in ds.Tables["USL"].Columns)
                                {
                                    rsUSL.Fields[dc.ColumnName].Value = drUSL[dc.Ordinal].ToString();
                                }
                                rsUSL.Update();
                                //
                                //SL_KOEF
                                //
                                DAO.Recordset rsSlKoef = db.OpenRecordset("SL_KOEFrokb");
                                List<DataRow> slKoefList = ds.Tables["SL_KOEF"].AsEnumerable().Where(m => m.Field<Int32>("USL_Id") == USL_Id).ToList<DataRow>();
                                if (slKoefList != null && slKoefList.Count > 0)
                                {
                                    foreach (DataRow drSlKoef in slKoefList)
                                    {
                                        rsSlKoef.AddNew();
                                        rsSlKoef.Fields["IDCASE"].Value = _IDCASE;
                                        foreach (DataColumn dc in ds.Tables["SL_KOEF"].Columns)
                                        {
                                            rsSlKoef.Fields[dc.ColumnName].Value = drSlKoef[dc.Ordinal].ToString();
                                        }
                                        rsSlKoef.Update();
                                    }
                                }
                                rsSlKoef.Close();
                            }
                        }
                        rsUSL.Close();
                    }
                    rsSLUCH.Close();

                }
                rs.Close();

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(dbEngine);
                dbEngine = null;
            }

            cat = null;
            return result;
        }

        private bool lFile(DataSet ds)
        {
            bool result = false;
            string strMes = "";
            OleDbConnection conn;
            ADODB.Connection adodbCon = new ADODB.Connection();
            adodbCon.ConnectionString = connString;
            ADOX.Catalog cat = new ADOX.Catalog();
            ADOX.Table tab;
            DAO.DBEngine dbEngine = new DAO.DBEngine();
            
            try
            {

                adodbCon.Open();
                cat.ActiveConnection = adodbCon;
                conn = new OleDbConnection(connString);

                conn.Open();
                if (conn.State != ConnectionState.Open)
                {
                    strMes = "Отсутствует соединение к БД";
                    MessageBox.Show(strMes);
                    txtLog.AppendLine(strMes);
                    return result;
                }
                txtLog.AppendLine("Создание структуры БД");
                Application.DoEvents();
                foreach (string tabName in tableListL)
                {
                    if (!ds.Tables.Contains(tabName))
                    {
                        strMes = $"Отсутствует таблица {tabName}";
                        MessageBox.Show(strMes);
                        txtLog.AppendLine(strMes);
                        continue;
                    }

                    string[] columnNames = ds.Tables[tabName]
                        .Columns.Cast<DataColumn>()
                        .Select(x => x.ColumnName)
                        .ToArray();

                    tab = new ADOX.Table();
                    tab.Name = $"{tabName}rokb";
                    //id
                    ADOX.Column column = new ADOX.Column();
                    column.Name = "id";
                    column.Type = ADOX.DataTypeEnum.adInteger;
                    column.ParentCatalog = cat;
                    column.Properties["AutoIncrement"].Value = true;
                    tab.Columns.Append(column);

                    foreach (string str in addFieldsL[tabName])
                    {
                        tab.Columns.Append(defCol(str));
                    }
                    foreach (string str in columnNames)
                    {
                        tab.Columns.Append(defCol(str));
                    }
                    cat.Tables.Append(tab);
                }

                conn.Close();


                //внесение данных
                string _PLAT = "61";
                txtLog.AppendLine("Внесение данных");
                Application.DoEvents();
                DAO.Database db = dbEngine.OpenDatabase(pathToDb);
                DAO.Recordset rs = db.OpenRecordset("ZGLVrokb");
                foreach (DataRow dr in ds.Tables["ZGLV"].Rows)
                {
                    rs.AddNew();
                    rs.Fields["H"].Value = "L";
                    foreach (DataColumn dc in ds.Tables["ZGLV"].Columns)
                    {
                        rs.Fields[dc.ColumnName].Value = dr[dc.ColumnName].ToString();
                    }
                    rs.Update();
                }
                rs.Close();
                //
                //PERS
                //
                txtLog.AppendLine("Таблица PERS");
                Application.DoEvents();
                System.Threading.Thread.Sleep(10000);
                pb.Minimum = 0;
                pb.Maximum = ds.Tables["PERS"].Rows.Count;
                pb.Value = 0;
                Application.DoEvents();
                rs = db.OpenRecordset("PERSrokb");
                int counter = 0;
                foreach (DataRow drSchet in ds.Tables["PERS"].Rows)
                {
                    pb.Value = counter++;
                    Application.DoEvents();
                    rs.AddNew();
                    rs.Fields["PLAT"].Value = _PLAT;
                    foreach (DataColumn dc in ds.Tables["PERS"].Columns)
                    {
                        rs.Fields[dc.ColumnName].Value = drSchet[dc.ColumnName].ToString();
                    }
                    rs.Update();
                }
                rs.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(dbEngine);
                dbEngine = null;
            }
            cat = null;
            return result;
        }

        private ADOX.Column defCol(string colName)
        {
            ADOX.Column col = new ADOX.Column();
            col.Name = colName;
            col.Type = ADOX.DataTypeEnum.adVarWChar;
            col.Attributes = ADOX.ColumnAttributesEnum.adColNullable;
            col.DefinedSize = 255;
            return col;
        }
    }
}
