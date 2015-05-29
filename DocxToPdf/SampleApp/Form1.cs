using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace SampleApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            this.AllowDrop = true;
            this.DragEnter += Form1_DragEnter;
            this.DragDrop += Form1_DragDrop;
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            bool processSomething = false;
            var watch = Stopwatch.StartNew();
            foreach (string file in files)
            {
                if (Path.GetExtension(file) != ".docx")
                    MessageBox.Show(this, String.Format("{0} is not a docx file", file));
                else
                {
                    try
                    {
                        DocxToPdf.DocxToPdf convertor = new DocxToPdf.DocxToPdf(file, file.Replace(".docx", ".pdf"));
                        processSomething = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, ex.Message);
                    }
                }
            }
            watch.Stop();
            if (processSomething)
            {
                MessageBox.Show(this, "Conversion Finish!\n" + watch.ElapsedMilliseconds + "ms");
            }
        }
    }
}
