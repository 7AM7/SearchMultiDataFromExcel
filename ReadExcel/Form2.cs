using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReadExcel
{


    public partial class Form2 : Form
    {
        StreamWriter FileData = new StreamWriter("Text_File.Text");
        DataSet result;
        List<string> Files;
        List<string> FilesName;
        OpenFileDialog opfd;
        string[] lines;
        string selectedPath = "";
        public Form2()
        {
            InitializeComponent();
            Files = new List<string>();
            FilesName = new List<string>();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog opfd = new FolderBrowserDialog();
            if (opfd.ShowDialog() == DialogResult.OK)
            {
                if (CheckExcal() == false)
                {
                    textBox1.Text = opfd.SelectedPath;
                    Files = Directory.GetFiles(opfd.SelectedPath, "*.xlsx")
                                         .Select(Path.GetFullPath)
                                         .ToList();
                    FilesName = Directory.GetFiles(opfd.SelectedPath, "*.xlsx")
                     .Select(Path.GetFileName)
                     .ToList();
                    button4.Enabled = true;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

            var t = new Thread((ThreadStart)(() =>
            {
                opfd = new OpenFileDialog() { Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*" };
                if (opfd.ShowDialog() == DialogResult.OK)
                {
                    lines = System.IO.File.ReadAllLines(opfd.FileName);
                }


                selectedPath = opfd.FileName;
            }));

            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
            foreach (var item in lines)
            {
                listBox1.Invoke(p => p.Items.Add(item));
            }

        }
        struct DataParmeter
        {
            public int Process;
            public int Delay;
        };
        bool CheckExcal()
        {
            Process[] processlist = Process.GetProcesses();
            foreach (Process theprocess in processlist)
            {
                //Console.WriteLine("Process: {0} ID: {1}", theprocess.ProcessName, theprocess.Id);
                if (theprocess.ProcessName == "EXCEL")
                {
                    //MessageBox.Show("Close Excel First Plzz");
                    return true;
                }
            }
            return false;
        }
        private void button4_Click(object sender, EventArgs e)
        {

            if (listBox1.Items.Count < 1)
            {
                MessageBox.Show("Select Id List Plzz");
                return;
            }
            if (CheckExcal() == true)
            {
                MessageBox.Show("Close Excel First Plzz");
                return;
            }



            if (!backgroundWorker2.IsBusy)
            {
                inputparmeter.Delay = 10;

                inputparmeter.Process = 100;

                backgroundWorker2.RunWorkerAsync(inputparmeter);
            }


           // MessageBox.Show("Finished");
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        void GetData()
        {

          
        }

        private DataParmeter inputparmeter;
        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            int process = ((DataParmeter)e.Argument).Process;
            int delay = ((DataParmeter)e.Argument).Delay;
            int index = 1;
            int t = 1;
            try
            {

                    if (!backgroundWorker2.CancellationPending)
                    {
                    foreach (var item in Files)
                    {
                        Console.WriteLine(item);
                        FileStream fs = File.Open(item, FileMode.Open, FileAccess.Read);
                        IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                        result = reader.AsDataSet();

                        reader.Close();

                        if (result.Tables[0].Rows.Count > listBox1.Items.Count)
                        {

                                for (int k = 0; k < result.Tables[0].Rows.Count; k++)
                            {
                                for (int j = 0; j < listBox1.Items.Count; j++)
                                {
                                    if ((Convert.ToInt64(result.Tables[0].Rows[k][0]) == Convert.ToInt64(listBox1.Items[j])))
                                    {

                                            backgroundWorker2.ReportProgress(index++ * 100 / process, string.Format("Process1 {0}/{1}", k, result.Tables[0].Rows.Count));

                                        Thread.Sleep(delay);
                                        richTextBox1.Invoke(p => p.AppendText(Convert.ToInt64(result.Tables[0].Rows[k][0]) + " , " + Convert.ToInt64(result.Tables[0].Rows[k][1]).ToString() + "\n"));
                                    }
                                }

                            }
                        }
                        else
                        {

                            for (int k = 0; k < listBox1.Items.Count; k++)
                            {
                                for (int j = 0; j < result.Tables[0].Rows.Count; j++)
                                {
                                    if ((Convert.ToInt64(result.Tables[0].Rows[j][0]) == Convert.ToInt64(listBox1.Items[k])))
                                    {

                                        backgroundWorker2.ReportProgress(index++ * 100 / process, string.Format("Process1 {0}/{1}", k, listBox1.Items.Count));
                                        
                                        Thread.Sleep(delay);
                                            richTextBox1.Invoke(p => p.AppendText(Convert.ToInt64(result.Tables[0].Rows[j][0]) + " , " + Convert.ToInt64(result.Tables[0].Rows[j][1]).ToString() + "\n"));
                                    }
                                }
                            }
                        }

                        result.Clear();
                        //MessageBox.Show("done");
                    }


                }
               
            }
            catch (Exception ex)
            {
                backgroundWorker2.CancelAsync();
                MessageBox.Show(ex.Message,"Infon",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }

        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
     
           
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Task Completed!!", "Info",MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (backgroundWorker2.IsBusy)
            {
                backgroundWorker2.CancelAsync();
                MessageBox.Show("Task Stoped!!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
