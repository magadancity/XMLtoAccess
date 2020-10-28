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
using System.Threading.Tasks;
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
            string strMes = "";
            OleDbConnection conn;
            ADODB.Connection adodbCon = new ADODB.Connection();
            adodbCon.ConnectionString = connString;
            ADOX.Catalog cat = new ADOX.Catalog();
            ADOX.Table tab;
            string fieldName, valueName, insCommand;
            OleDbCommand cmd;
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
                
                foreach (string tabName in tableListH)
                {
                    fieldName = "";
                    valueName = "";
                    insCommand = "";
                    if (!ds.Tables.Contains(tabName))
                    {
                        strMes = $"Отсутствует таблица {tabName}";
                        MessageBox.Show(strMes);
                        txtLog.AppendLine(strMes);
                        continue;
                    }

                    //DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,new object[] { null, null, $"{tabName}rokb", "TABLE" });
                    //if (schemaTable.Rows.Count>0)
                    //{
                    //    if (tabName == "ZGLV")
                    //    {

                    //    }
                    //    continue;
                    //}

                    string[] columnNames = ds.Tables[tabName]
                        .Columns.Cast<DataColumn>()
                        .Select(x => x.ColumnName)
                        .ToArray();

                    tab = new ADOX.Table();
                    tab.Name = $"{tabName}rokb";
                    //id
                    fieldName += "";
                    ADOX.Column column = new ADOX.Column();
                    column.Name = "id";
                    column.Type = ADOX.DataTypeEnum.adInteger;
                    column.ParentCatalog = cat;
                    column.Properties["AutoIncrement"].Value = true;
                    tab.Columns.Append(column);

                    foreach(string str in addFieldsH[tabName])
                    {
                        fieldName += $"[{str}],";
                        valueName += $"@{str},";
                        
                        tab.Columns.Append(defCol(str));
                    }
                    foreach(string str in columnNames)
                    {
                        fieldName += $"[{str}],";
                        valueName += $"@{str},";
                        tab.Columns.Append(defCol(str));
                    }
                    cat.Tables.Append(tab);
                    fieldName = fieldName.TrimEnd(',');
                    valueName = valueName.TrimEnd(',');
                    insCommand = $"insert into [{tabName}rokb]({fieldName}) values({valueName})";
                    //
                    cmd = new OleDbCommand(insCommand, conn);
                    commands[tabName] = cmd;
                    //cmd.Parameters.Clear();
                }

                conn.Close();

                //insDataH(ds, conn);

                //conn.Open();
                //using (OleDbTransaction trans = conn.BeginTransaction())
                //{
                //    string _CODE = "", _PLAT = "61", _NSCHET = "", _N_ZAP = "", _IDCASE = "", _USL_TIP = "2", _DATE_INJ = "";
                //    OleDbCommand cOleDbCommand = new OleDbCommand();

                //    cOleDbCommand = commands["ZGLV"];
                //    cOleDbCommand.Transaction = trans;
                //    //System.Threading.Thread.Sleep(60000);

                //    //DataTable schemaTable2 = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, $"ZGLVrokb", "TABLE" });
                //    //if (schemaTable2.Rows.Count > 0)
                //    //{
                //    //    int z = 0;
                //    //}

                //    foreach (DataRow dr in ds.Tables["ZGLV"].Rows)
                //    {
                //        cOleDbCommand.Parameters.AddWithValue("H", "H");
                //        foreach (DataColumn dc in ds.Tables["ZGLV"].Columns)
                //        {
                //            cOleDbCommand.Parameters.AddWithValue(dc.ColumnName, dr[dc.ColumnName].ToString());
                //        }
                //        cOleDbCommand.ExecuteNonQuery();
                //    }

                //    //
                //    //SCHET
                //    //
                //    cOleDbCommand = commands["SCHET"];
                //    cOleDbCommand.Transaction = trans;
                //    DataRow drSchet = ds.Tables["SCHET"].Rows[0];
                //    _CODE = drSchet["CODE"].ToString();
                //    _NSCHET = drSchet["NSCHET"].ToString();
                //    foreach (DataColumn dc in ds.Tables["SCHET"].Columns)
                //    {
                //        cOleDbCommand.Parameters.AddWithValue(dc.ColumnName, drSchet[dc.ColumnName].ToString());
                //    }
                //    cOleDbCommand.ExecuteNonQuery();

                //    //
                //    //ZAP
                //    //
                //    cOleDbCommand = commands["ZAP"];
                //    cOleDbCommand.Transaction = trans;
                //    foreach (DataRow dr in ds.Tables["ZAP"].Rows)
                //    {
                //        _N_ZAP = dr["N_ZAP"].ToString();
                //        string ZAP_Id = dr["ZAP_Id"].ToString();
                //        cOleDbCommand.Parameters.AddWithValue("CODE", _CODE);
                //        cOleDbCommand.Parameters.AddWithValue("PLAT", _PLAT);
                //        cOleDbCommand.Parameters.AddWithValue("NSCHET", _NSCHET);
                //        foreach (DataColumn dc in ds.Tables["ZAP"].Columns)
                //        {
                //            cOleDbCommand.Parameters.AddWithValue(dc.ColumnName, dr[dc.ColumnName].ToString());
                //        }
                //        cOleDbCommand.ExecuteNonQuery();
                //    }
                //    trans.Commit();
                //}

                //conn.Close();

                DAO.DBEngine dbEngine = new DAO.DBEngine();
                Boolean CheckFl = false;

                try
                {
                    DataTable dtOutData = ds.Tables["ZAP"];
                    DAO.Database db = dbEngine.OpenDatabase(pathToDb);
                    DAO.Recordset AccesssRecordset = db.OpenRecordset("ZAProkb");
                    DAO.Field[] AccesssFields = new DAO.Field[dtOutData.Columns.Count];

                    //Loop on each row of dtOutData
                    for (Int32 rowCounter = 0; rowCounter < dtOutData.Rows.Count; rowCounter++)
                    {
                        AccesssRecordset.AddNew();
                        //Loop on column
                        for (Int32 colCounter = 0; colCounter < dtOutData.Columns.Count; colCounter++)
                        {
                            // for the first time... setup the field name.
                            if (!CheckFl)
                                AccesssFields[colCounter] = AccesssRecordset.Fields[dtOutData.Columns[colCounter].ColumnName];
                            AccesssFields[colCounter].Value = dtOutData.Rows[rowCounter][colCounter];
                        }

                        AccesssRecordset.Update();
                        CheckFl = true;
                    }

                    AccesssRecordset.Close();
                    db.Close();
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(dbEngine);
                    dbEngine = null;
                }

                //conn.Open();

                //string _CODE = "", _PLAT = "61", _NSCHET = "", _N_ZAP = "", _IDCASE = "", _USL_TIP = "2", _DATE_INJ = "";
                //OleDbCommand cOleDbCommand = new OleDbCommand();

                //cOleDbCommand = commands["ZGLV"];
                ////cOleDbCommand.Transaction = trans;

                //foreach (DataRow dr in ds.Tables["ZGLV"].Rows)
                //{
                //    cOleDbCommand.Parameters.AddWithValue("H", "H");
                //    foreach (DataColumn dc in ds.Tables["ZGLV"].Columns)
                //    {
                //        cOleDbCommand.Parameters.AddWithValue(dc.ColumnName, dr[dc.ColumnName].ToString());
                //    }
                //    cOleDbCommand.ExecuteNonQuery();
                //}

                ////
                ////SCHET
                ////
                //cOleDbCommand = commands["SCHET"];
                ////cOleDbCommand.Transaction = trans;
                //DataRow drSchet = ds.Tables["SCHET"].Rows[0];
                //_CODE = drSchet["CODE"].ToString();
                //_NSCHET = drSchet["NSCHET"].ToString();
                //foreach (DataColumn dc in ds.Tables["SCHET"].Columns)
                //{
                //    cOleDbCommand.Parameters.AddWithValue(dc.ColumnName, drSchet[dc.ColumnName].ToString());
                //}
                //cOleDbCommand.ExecuteNonQuery();

                ////
                ////ZAP
                ////
                //cOleDbCommand = commands["ZAP"];
                ////cOleDbCommand.Transaction = trans;
                //cOleDbCommand.CommandText = "insert into ZAProkb(CODE) values (@CODE)";
                //txtLog.AppendLine($"{DateTime.Now} начало");
                //for (int index = 0; index < 6000; index++)
                //{
                //    if (index % 500 == 0)
                //    {
                //        conn.Close();
                //        txtLog.AppendLine($"{DateTime.Now} - {index}");
                //        conn.Open();
                //    }
                //    cOleDbCommand.Parameters.AddWithValue("CODE", index.ToString());
                //    cOleDbCommand.ExecuteNonQuery();
                //}
                ////foreach (DataRow dr in ds.Tables["ZAP"].Rows)
                ////{
                ////    _N_ZAP = dr["N_ZAP"].ToString();
                ////    string ZAP_Id = dr["ZAP_Id"].ToString();
                ////    cOleDbCommand.Parameters.AddWithValue("CODE", _CODE);
                ////    cOleDbCommand.Parameters.AddWithValue("PLAT", _PLAT);
                ////    cOleDbCommand.Parameters.AddWithValue("NSCHET", _NSCHET);
                ////    foreach (DataColumn dc in ds.Tables["ZAP"].Columns)
                ////    {
                ////        cOleDbCommand.Parameters.AddWithValue(dc.ColumnName, dr[dc.ColumnName].ToString());
                ////    }
                ////    await cOleDbCommand.ExecuteNonQueryAsync();
                ////}

                //txtLog.AppendLine($"{DateTime.Now} окончание");
                
                //conn.Close();

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
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
            string fieldName, valueName, insCommand;
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

                foreach (string tabName in tableListL)
                {
                    fieldName = "";
                    valueName = "";
                    insCommand = "";
                    if (!ds.Tables.Contains(tabName))
                    {
                        strMes = $"Отсутствует таблица {tabName}";
                        MessageBox.Show(strMes);
                        txtLog.AppendLine(strMes);
                        continue;
                    }

                    DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, tabName, "TABLE" });
                    if (schemaTable.Rows.Count > 0)
                    {
                        if (tabName == "ZGLAV")
                        {

                        }
                        continue;
                    }

                    string[] columnNames = ds.Tables[tabName]
                        .Columns.Cast<DataColumn>()
                        .Select(x => x.ColumnName)
                        .ToArray();

                    tab = new ADOX.Table();
                    tab.Name = $"{tabName}rokb";
                    //id
                    fieldName += "id,";
                    ADOX.Column column = new ADOX.Column();
                    column.Name = "id";
                    column.Type = ADOX.DataTypeEnum.adInteger;
                    column.ParentCatalog = cat;
                    column.Properties["AutoIncrement"].Value = true;
                    tab.Columns.Append(column);

                    if (!addFieldsL.ContainsKey(tabName)) { continue; }

                    foreach (string str in addFieldsL[tabName])
                    {
                        fieldName += $"{str},";
                        valueName += $"@{str},";
                        tab.Columns.Append(str, ADOX.DataTypeEnum.adVarWChar, 255);
                    }
                    foreach (string str in columnNames)
                    {
                        fieldName += $"{str},";
                        valueName += $"@{str},";
                        tab.Columns.Append(str, ADOX.DataTypeEnum.adVarWChar, 255);
                    }
                    cat.Tables.Append(tab);
                    fieldName = fieldName.TrimEnd(',');
                    valueName = valueName.TrimEnd(',');
                    insCommand = $"insert into {tabName}({fieldName}) values({valueName})";
                }
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            cat = null;
            return result;
        }


        private async void insDataH(DataSet ds, OleDbConnection conn)
        {
            try
            {
                Stopwatch sw = new Stopwatch();
                sw.Start();
                //OleDbConnection conn = new OleDbConnection(connString);
                await Task.Run(async () =>
                {
                    conn.Open();
                    using (OleDbTransaction trans = conn.BeginTransaction())
                    {
                        string _CODE = "", _PLAT = "61", _NSCHET = "", _N_ZAP = "", _IDCASE = "", _USL_TIP = "2", _DATE_INJ = "";
                        OleDbCommand cOleDbCommand = new OleDbCommand();

                        cOleDbCommand = commands["ZGLV"];
                        cOleDbCommand.Transaction = trans;

                        foreach (DataRow dr in ds.Tables["ZGLV"].Rows)
                        {
                            cOleDbCommand.Parameters.AddWithValue("H", "H");
                            foreach (DataColumn dc in ds.Tables["ZGLV"].Columns)
                            {
                                cOleDbCommand.Parameters.AddWithValue(dc.ColumnName, dr[dc.ColumnName].ToString());
                            }
                            await cOleDbCommand.ExecuteNonQueryAsync();
                        }

                        //
                        //SCHET
                        //
                        cOleDbCommand = commands["SCHET"];
                        cOleDbCommand.Transaction = trans;
                        DataRow drSchet = ds.Tables["SCHET"].Rows[0];
                        _CODE = drSchet["CODE"].ToString();
                        _NSCHET = drSchet["NSCHET"].ToString();
                        foreach (DataColumn dc in ds.Tables["SCHET"].Columns)
                        {
                            cOleDbCommand.Parameters.AddWithValue(dc.ColumnName, drSchet[dc.ColumnName].ToString());
                        }
                        await cOleDbCommand.ExecuteNonQueryAsync();

                        //
                        //ZAP
                        //
                        cOleDbCommand = commands["ZAP"];
                        cOleDbCommand.Transaction = trans;
                        cOleDbCommand.CommandText = "insert into ZAProkb(CODE) values (@CODE)";
                        txtLog.AppendLine($"{DateTime.Now} начало");
                        for (int index = 0; index < 6000; index++)
                        {
                            if (index % 500 == 0)
                            {
                                txtLog.AppendLine($"{DateTime.Now} - {index}");
                            }
                            cOleDbCommand.Parameters.AddWithValue("CODE", index.ToString());
                            await cOleDbCommand.ExecuteNonQueryAsync();
                        }
                        //foreach (DataRow dr in ds.Tables["ZAP"].Rows)
                        //{
                        //    _N_ZAP = dr["N_ZAP"].ToString();
                        //    string ZAP_Id = dr["ZAP_Id"].ToString();
                        //    cOleDbCommand.Parameters.AddWithValue("CODE", _CODE);
                        //    cOleDbCommand.Parameters.AddWithValue("PLAT", _PLAT);
                        //    cOleDbCommand.Parameters.AddWithValue("NSCHET", _NSCHET);
                        //    foreach (DataColumn dc in ds.Tables["ZAP"].Columns)
                        //    {
                        //        cOleDbCommand.Parameters.AddWithValue(dc.ColumnName, dr[dc.ColumnName].ToString());
                        //    }
                        //    await cOleDbCommand.ExecuteNonQueryAsync();
                        //}
                        trans.Commit();
                        txtLog.AppendLine($"{DateTime.Now} окончание");
                    }

                    conn.Close();
                });
                sw.Stop();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
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
