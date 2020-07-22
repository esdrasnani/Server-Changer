using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml;
using MaterialSkin;
using Ionic.Zip;
using System.Threading;
using System.Diagnostics;

namespace prj_serverchanger_materialskin
{
    public partial class Form1 : MaterialSkin.Controls.MaterialForm
    {
        MaterialSkinManager skinManager;
        int _TotalArquivos = 0;
        int _count = 0;
        public Form1()
        {
            InitializeComponent();

            skinManager = MaterialSkinManager.Instance;
            skinManager.AddFormToManage(this);
            skinManager.Theme = MaterialSkinManager.Themes.DARK;
            skinManager.ColorScheme = new ColorScheme(Primary.Blue600, Primary.Blue900, Primary.Blue900, Accent.LightBlue200, TextShade.WHITE);

            materialListView1.AllowDrop = true;

            backgroundWorker1.WorkerReportsProgress = true;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            materialListView1.Columns[0].Width = materialListView1.Width - 4;
        }

        private void materialListView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void materialListView1_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            for (int i = 0; i < s.Length; i++)
            {
                FileInfo fi = new FileInfo(s[i]);
                if (fi.Extension == ".twb" || fi.Extension == ".twbx")
                    materialListView1.Items.Add(s[i]);
                else
                {
                    MessageBox.Show("Arquivo " + s[i] + " Inválido!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (s.Length == 1)
                        return;
                }
            }

            List<String> files_path = new List<String>();
            for (int i = 0; i < materialListView1.Items.Count; i++)
                files_path.Add(materialListView1.Items[i].Text);

            _TotalArquivos = materialListView1.Items.Count;

            materialProgressBar1.Maximum = _TotalArquivos;
            lbl_pct.Text = "";

            BackgroundWorker bgw;
            for (int i = 0; i < materialListView1.Items.Count; i++)
            {
                List<Object> arguments = new List<object>();
                arguments.Add(files_path);
                arguments.Add(i);

                bgw = new BackgroundWorker();
                bgw.DoWork += backgroundWorker1_DoWork;
                bgw.RunWorkerCompleted += backgroundWorker1_RunWorkerCompleted;
                bgw.ProgressChanged += backgroundWorker1_ProgressChanged;                
                bgw.RunWorkerAsync(arguments);
            }
        }

        private void sairToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            notifyIcon1.Visible = true;

            if (this.WindowState == FormWindowState.Minimized)
                Hide();
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            this.WindowState = FormWindowState.Normal;
            notifyIcon1.Visible = true;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.WindowState = FormWindowState.Minimized;
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

            List<Object> arguments = e.Argument as List<object>;
            List<String> files_path = (List<String>)arguments[0];
            int i = (int)arguments[1];

            FileInfo fi = new FileInfo(files_path[i]);

            if (fi.Extension == ".twb")
            {
                File.Move(files_path[i], Path.ChangeExtension(files_path[i], ".xml"));


                files_path[i] = Path.ChangeExtension(files_path[i], ".xml");


                XDocument doc;
                try
                {
                    doc = XDocument.Load(files_path[i]);



                }
                catch (XmlException xmle)
                {
                    MessageBox.Show("Erro ao Carregar o Arquivo!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show(xmle.ToString());
                    return;
                }

                var elements = doc.Descendants("connection");


                foreach (XElement element in elements)
                {
                    if (element.Attribute("server") != null)
                    {
                        if ((string)element.Attribute("server") != null || element.Attribute("server").ToString() != "")
                        {
                            element.Attribute("server").Value = "tableau.scanntech.com";
                        }
                    }
                }

                doc.Save(files_path[i]);


                File.Move(files_path[i], Path.ChangeExtension(files_path[i], ".twb"));

            }
            else if (fi.Extension == ".twbx")
            {

                File.Move(files_path[i], Path.ChangeExtension(files_path[i], ".zip"));

                files_path[i] = Path.ChangeExtension(files_path[i], ".zip");

                string file_twb = files_path[i].Replace("zip", "twb");
                int index = files_path[i].LastIndexOf(@"\");
                index++;
                file_twb = file_twb.Substring(index, file_twb.Length - index);

                string file_twb_aux;
                string path = files_path[i].Substring(0, index);
                using (ZipFile zip = ZipFile.Read(files_path[i]))
                {
                    ZipEntry file = new ZipEntry();

                    foreach (ZipEntry entry in zip.Entries)
                    {
                        if (entry.FileName.Contains("twb"))
                        {
                            file = entry;
                            file.Extract(path);
                            break;
                        }
                    }

                    zip.RemoveEntry(file);
                    zip.Save();

                    file_twb_aux = path + file.FileName;

                    File.Move(file_twb_aux, Path.ChangeExtension(file_twb_aux, ".xml"));
                    file_twb_aux = Path.ChangeExtension(file_twb_aux, ".xml");
                }

                XDocument doc;
                try
                {
                    doc = XDocument.Load(file_twb_aux);
                }
                catch (XmlException xmle)
                {
                    MessageBox.Show("Erro ao Carregar o Arquivo!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show(xmle.ToString());
                    return;
                }

                var elements = doc.Descendants("connection");

                foreach (XElement element in elements)
                {
                    if (element.Attribute("server") != null)
                    {
                        if ((string)element.Attribute("server") != null || element.Attribute("server").ToString() != "")
                        {
                            element.Attribute("server").Value = "tableau.scanntech.com";
                        }
                    }
                }

                doc.Save(file_twb_aux);

                File.Move(file_twb_aux, Path.ChangeExtension(file_twb_aux, ".twb"));
                file_twb_aux = Path.ChangeExtension(file_twb_aux, ".twb");


                using (ZipFile zip = ZipFile.Read(files_path[i]))
                {
                    zip.AddFile(file_twb_aux, "");
                    zip.Save();
                }

                File.Delete(file_twb_aux);
                File.Move(files_path[i], Path.ChangeExtension(files_path[i], ".twbx"));
            }

            backgroundWorker1.ReportProgress(i);
            Thread.Sleep(500);
            e.Result = true;
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int total = e.ProgressPercentage + 1;
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
                MessageBox.Show("Operação Cancelada", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
            {
                _count++;
                lbl_pct.Text = _count + "/" + _TotalArquivos;
                materialProgressBar1.Value = _count;

                if (_count == _TotalArquivos)
                {
                    lbl_pct.Text = "Concluído!";

                    if (Environment.Is64BitOperatingSystem)
                    {
                        for (int i = 0; i < materialListView1.Items.Count; i++)
                        {
                            string fullpath = materialListView1.Items[i].Text;
                            ProcessStartInfo psi = new ProcessStartInfo();
                            psi.FileName = Path.GetFileName(fullpath);
                            psi.WorkingDirectory = Path.GetDirectoryName(fullpath);
                            Process.Start(psi);
                        }
                    }
                    else
                        MessageBox.Show("Não é Possível Iniciar o Tableau!\nSistema 32-bits Não Suportado!", "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    
                    materialListView1.Items.Clear();
                    
                    _count = 0;                    
                }
            }
        }
    }
}
