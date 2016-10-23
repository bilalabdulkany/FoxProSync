using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FoxProUpdate
{
    public partial class Form1 : Form
    {
        string[] confini;
        string strLogConnectionString;
        String connectionString = "";
        SqlConnection con;
        SqlCommand com;
        SqlDataReader reader;

        OleDbConnection strConLog;
        OleDbCommand oComm;
        OleDbDataReader oReader;
        DirectoryInfo di;
        FileInfo[] allFiles;
        bool hasRows = false;
        string[] columns;
        bool chkFields;
        string commandText;
        int rowcount = 0;
        int currentrow = 0;
        string currentDB;
        int currentDBn;
        int dbCount;
        int maxRow;
        int updateCount = 0;
        int insertCount = 0;
        public Form1()
        {
            InitializeComponent();
            outputToFile();
            //System.Globalization.CultureInfo.CurrentCulture.ClearCachedData();
            Console.WriteLine("============================================");
            Console.WriteLine("************* FoxPro Sync ******************");
            Console.WriteLine("        Date: " + DateTime.Now.ToString("dd-MM-yyyy HH:MM:ss tt"));
            Console.WriteLine("============================================");
            Console.WriteLine("Reading configuration file..");
            confini = File.ReadAllLines("config.ini");
            strLogConnectionString = @"Provider=vfpoledb;Data Source=" + confini[0] + @";Collating Sequence=machine;Mode=ReadWrite;";
            strConLog = new OleDbConnection(strLogConnectionString);

            oComm = new OleDbCommand();
            oComm.Connection = strConLog;
            Console.WriteLine("Established OleDbConnection..");
            connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["Zahira_SISConnectionString"].ToString();
            con = new SqlConnection(connectionString);
            /* con = new SqlConnection("user id=" + confini[2] + ";" +                                 "password=" + confini[3] + ";server=" + confini[4] + ";" +
                             //           "Trusted_Connection=yes;" +
                                        "database=" + confini[5] + "; ");
             */
            com = new SqlCommand();
            com.Connection = con;
            Console.WriteLine("Established SQLDbConnection..");
            
            di = new DirectoryInfo(confini[1]);
            Console.WriteLine("Zip file directory: " + confini[6]);
            Console.WriteLine("============================================");


        }

        private void button1_Click(object sender, EventArgs e)
        {
            currentDBn = 0;
            backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            allFiles = di.GetFiles("*.dbf");
            dbCount = allFiles.Length;
            foreach (FileInfo fi in allFiles)
            {
                currentDB = fi.Name.ToLower().Replace(".dbf", "");

                Console.WriteLine("Parsing table:" + currentDB);


                if (fi.Name.ToLower() == "student.dbf")
                {
                    strLogConnectionString = @"Provider=vfpoledb;Data Source=" + fi.FullName + @";Collating Sequence=machine;Mode=ReadWrite;";
                    strConLog.ConnectionString = strLogConnectionString;
                    try
                    {
                        strConLog.Open();
                    }
                    catch
                    {
                        currentDBn++;
                        continue;
                    }

                    oComm.CommandText = "select count(*) from " + fi.Name.ToLower().Replace(".dbf", "");
                    oReader = oComm.ExecuteReader();
                    while (oReader.Read())
                    {
                        rowcount = Decimal.ToInt32((decimal)oReader[0]);
                    }
                    Console.WriteLine("row count in " + currentDB + ": " + rowcount);
                    strConLog.Close();
                    strConLog.Open();
                    oComm.CommandText = "select * from " + fi.Name.ToLower().Replace(".dbf", "");
                    oReader = oComm.ExecuteReader();
                    chkFields = true;
                    currentrow = 0;
                    while (oReader.Read())
                    {
                        if (chkFields)
                        {
                            DataTable tableSchema = oReader.GetSchemaTable();
                            columns = new string[tableSchema.Rows.Count];
                            foreach (DataRow row in tableSchema.Rows)
                            {
                                columns[tableSchema.Rows.IndexOf(row)] = row["ColumnName"].ToString();
                            }
                            chkFields = false;
                        }
                        con.Open();
                        com.CommandText = "select * from " + fi.Name.ToLower().Replace(".dbf", "") + " where key_fld='" + oReader["key_fld"] + "'";
                        reader = com.ExecuteReader();
                        hasRows = false;
                        while (reader.Read())
                        {
                            hasRows = true;
                        }
                        con.Close();
                        con.Open();
                        if (hasRows)
                        {

                            commandText = "update " + fi.Name.ToLower().Replace(".dbf", "") + " set ";
                            foreach (string column in columns)
                            {
                                if (oReader[column].GetType() == typeof(DateTime))
                                {
                                    commandText = commandText + "[" + column + "]='" + String.Format("{0:MM/dd/yy hh:mm:ss tt}", (DateTime)oReader[column]) + "',";
                                }
                                else
                                {
                                    commandText = commandText + "[" + column + "]='" + oReader[column].ToString().Replace("'", "''") + "',";
                                }
                            }
                            commandText = commandText.Substring(0, commandText.Length - 1) + " where key_fld='" + oReader["key_fld"] + "'";
                            Console.WriteLine(commandText);
                            com.CommandText = commandText;
                            com.ExecuteNonQuery();
                            updateCount++;
                        }
                        else
                        {
                            commandText = "insert into " + fi.Name.ToLower().Replace(".dbf", "") + " (";
                            foreach (string column in columns)
                            {
                                commandText = commandText + "[" + column + "],";
                            }
                            commandText = commandText.Substring(0, commandText.Length - 1) + ") values (";
                            foreach (string column in columns)
                            {
                                if (oReader[column].GetType() == typeof(DateTime))
                                {
                                    commandText = commandText + "'" + String.Format("{0:MM/dd/yy hh:mm:ss tt}", (DateTime)oReader[column]) + "',";
                                }
                                else
                                {
                                    commandText = commandText + "'" + oReader[column].ToString().Replace("'", "''") + "',";
                                }
                            }
                            commandText = commandText.Substring(0, commandText.Length - 1) + ")";
                            com.CommandText = commandText;
                            Console.WriteLine(commandText);
                            com.ExecuteNonQuery();
                            insertCount++;
                        }
                        con.Close();
                        currentrow++;
                        if (rowcount == 0)
                            backgroundWorker1.ReportProgress(100);
                        else
                            backgroundWorker1.ReportProgress((currentrow * 100) / rowcount);
                    }
                    strConLog.Close();
                    currentDBn++;
                }
                else
                {
                    con.Open();
                    com.CommandText = "select max(key_fld) from " + currentDB;
                    try
                    {
                        reader = com.ExecuteReader();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                        currentDBn++;
                        con.Close();
                        continue;
                    }
                    while (reader.Read())
                    {
                        if (reader[0].ToString().Trim() == "")
                        {
                            currentDBn++;
                            con.Close();
                            break;
                        }
                        else
                        {
                            maxRow = (int)reader[0];
                            Console.WriteLine("max rows:" + maxRow);
                        }
                    }
                    con.Close();
                    strLogConnectionString = @"Provider=vfpoledb;Data Source=" + fi.FullName + @";Collating Sequence=machine;Mode=ReadWrite;";
                    strConLog.ConnectionString = strLogConnectionString;
                    try
                    {
                        strConLog.Open();
                    }
                    catch
                    {
                        currentDBn++;
                        continue;
                    }

                    oComm.CommandText = "select count(*) from " + fi.Name.ToLower().Replace(".dbf", "") + " where key_fld>" + maxRow;
                    oReader = oComm.ExecuteReader();
                    while (oReader.Read())
                    {
                        rowcount = Decimal.ToInt32((decimal)oReader[0]);
                    }
                    strConLog.Close();
                    strConLog.Open();
                    oComm.CommandText = "select * from " + fi.Name.ToLower().Replace(".dbf", "") + " where key_fld>" + maxRow;
                    if (fi.Name.ToLower().Equals("stumnthhst.dbf"))
                    {
                        Console.WriteLine("stumnthst table parsing..");
                        //oComm.CommandText = "SELECT [key_fld],[feesmnth],[key_stu],(STR(mfees)) as mfees,[paid],(STR(topay)) as topay,[paidon],[key_fr],[key_mfp],[key_change],[key_trn],[trntype]  FRom stumnthhst";
                    }
                    oReader = oComm.ExecuteReader();
                    chkFields = true;
                    currentrow = 0;
                    while (oReader.Read())
                    {
                        if (chkFields)
                        {
                            DataTable tableSchema = oReader.GetSchemaTable();
                            columns = new string[tableSchema.Rows.Count];
                            foreach (DataRow row in tableSchema.Rows)
                            {
                                columns[tableSchema.Rows.IndexOf(row)] = row["ColumnName"].ToString();
                            }
                            chkFields = false;
                        }

                        con.Open();
                        String colReader = "";
                        commandText = "insert into " + fi.Name.ToLower().Replace(".dbf", "") + " (";
                        foreach (string column in columns)
                        {
                            commandText = commandText + "[" + column + "],";
                        }
                        commandText = commandText.Substring(0, commandText.Length - 1) + ") values (";
                        foreach (string column in columns)
                        {


                            if (fi.Name.ToLower() == "stumnthhst.dbf")
                            {
                                //  Console.WriteLine(column);
                                if (column == "mfees" || column == "topay" || column == "paid")
                                {
                                    //boolean found=true;
                                    String colData = "";


                                    String[] colArray = { "mfees", "topay", "paid" };

                                    int updateCount = 0;
                                    //TODO
                                    /**
                                     * check for combinations of mfees, topay, paid if two or more columns are null
                                     * */
                                    //   Console.WriteLine("++++++++");

                                    try
                                    {
                                        oReader[column].GetType();
                                        //Console.WriteLine(column + "---" + oReader[column].GetType());
                                    }
                                    catch (Exception et)
                                    {

                                        foreach (String col in colArray)
                                        {
                                            try
                                            {
                                                //Check if the column does not throw exceptions.


                                                colReader = oReader[col] + "";
                                                //Console.WriteLine("-----------: " + colReader);
                                            }
                                            catch (Exception e3)
                                            {
                                                Console.WriteLine(e3);
                                                //If exceptions are thrown here, then update the sql statement.
                                                try
                                                {
                                                    colData = colData + col + " = 0,";
                                                    updateCount++;//increment the update count to check how many rows need to get updated.

                                                    //executeExcemptedColumn(oReader["key_fld"].ToString(),column,strConLog);
                                                    // Console.WriteLine(colData);
                                                    //continue;
                                                }
                                                catch (Exception e4)
                                                {
                                                    Console.WriteLine("exception----" + e4);
                                                    continue;
                                                }
                                            }

                                        }
                                    }

                                    if (updateCount >= 1)
                                    {
                                        colData = colData.Remove(colData.Length - 1);
                                        Console.WriteLine("updating statement:" + colData);
                                        executeExcemptedColumn(oReader["key_fld"].ToString(), colData, strConLog);
                                        updateCount = 0;
                                    }
                                }

                            }
                            // Console.WriteLine("***");
                            if (oReader[column].GetType() == typeof(DateTime))
                            {

                                commandText = commandText + "'" + String.Format("{0:MM/dd/yy hh:mm:ss tt}", (DateTime)oReader[column]) + "',";
                            }
                            else
                            {

                                commandText = commandText + "'" + oReader[column].ToString().Replace("'", "''") + "',";
                            }

                        }
                        commandText = commandText.Substring(0, commandText.Length - 1) + ")";
                        Console.WriteLine(commandText);
                        com.CommandText = commandText;
                        com.ExecuteNonQuery();
                        insertCount++;
                        currentrow++;
                        if (rowcount == 0)
                            backgroundWorker1.ReportProgress(100);
                        else
                        {
                            backgroundWorker1.ReportProgress((currentrow * 100) / rowcount);

                        }

                        con.Close();
                    }
                    currentrow++;

                    if (rowcount == 0)
                        backgroundWorker1.ReportProgress(100);
                    else
                    {
                        backgroundWorker1.ReportProgress((currentrow * 100) / rowcount);

                    }
                }
                strConLog.Close();
                currentDBn++;
            }
            MessageBox.Show(updateCount + " Rows Updated, " + insertCount + " Rows Inserted");
        }
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            textBox1.Text = currentDB;
            textBox2.Text = e.ProgressPercentage.ToString() + "%";
            progressBar1.Value = (currentDBn * 100 / dbCount);
            progressBar2.Value = e.ProgressPercentage;
            lblTables.Text = currentDBn.ToString().Trim() + "/" + dbCount.ToString().Trim();
            lblRecords.Text = currentrow.ToString().Trim() + "/" + rowcount.ToString().Trim();
        }

        private void executeExcemptedColumn(String key, String col, OleDbConnection con)
        {
            try
            {
                //con.Close();
                //con.Open();   
                OleDbCommand oCom = new OleDbCommand();
                oCom.Connection = strConLog;

                String sql = "";
                sql = "update stumnthhst set " + col + " where key_fld=" + key;
                Console.WriteLine(sql);
                oCom.CommandText = sql;
                oCom.ExecuteNonQuery();

            }
            catch (Exception e)
            {
                Console.WriteLine("executeExcempt:" + e);
            }
        }

        private void outputToFile()
        {
            DateTime today = DateTime.Today;
            String date = today.ToString("dd-MM-yyyy");
            FileStream filestream = new FileStream("./trace." + date + ".log", FileMode.Append);
            var streamwriter = new StreamWriter(filestream);
            streamwriter.AutoFlush = true;
            Console.SetOut(streamwriter);
            Console.SetError(streamwriter);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            di = new DirectoryInfo(confini[6]);
            allFiles= di.GetFiles("*.zip");
            //String filesToconvert=allFiles[0].ToString();
            //Console.WriteLine(filesToconvert.Split('T')[0].ToString().Substring(6));
            foreach (FileInfo fi in allFiles)
            {

                Console.WriteLine(fi.FullName.ToString());
            }
            string path = confini[6].ToString()+allFiles[allFiles.Length-1];
            Console.WriteLine("--------------------------------------------");           
            Console.WriteLine("Chosen Backup file: " + path);
            Console.WriteLine("--------------------------------------------");
            lblPath.Text=path;
            string directory = "Lctn";
            {
                using (var file = File.OpenRead(path))
                using (var zip = new ZipArchive(file, ZipArchiveMode.Read))
                {
                    var result = from currEntry in zip.Entries
                                 where Path.GetDirectoryName(currEntry.FullName) == directory
                                 where !String.IsNullOrEmpty(currEntry.Name)
                                 select currEntry;

                    string dbffile = "";
                    foreach (ZipArchiveEntry entry in result)
                    {
                        using (var stream = entry.Open())
                        {
                            if (entry.FullName.EndsWith(".dbf", StringComparison.OrdinalIgnoreCase))
                            {
                                dbffile = entry.FullName.ToString();
                                dbffile = dbffile.Split('/')[1];
                                Console.WriteLine(dbffile);
                            }

                        }
                    }
                }
            }
        }
    }
}
