using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelModuleSuitable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            //DD(@"C:\Users\deve-yamada\Documents\Excel_Module_BackUp\EmpAddin\20110311170639\Module1.bas");
        }
        private void FF(string folderPath)
        {
            string[] paths = Directory.GetFiles(folderPath);
            for (int i = 0; i < paths.Length; i++)
            {
                string p = paths[i];
                if (Path.GetExtension(p).ToLower() != ".frx")
                {
                    this.DD(p);
                }
                
            }
        }
        private void DD(string path)
        {
            using (StreamWriter sw = new StreamWriter(Path.Combine(Path.GetDirectoryName(path), "ks_" + Path.GetFileName(path)), false, Encoding.GetEncoding("shift-jis")))
            using (StreamReader sr = new StreamReader(path, Encoding.GetEncoding("shift-jis")))
            {
                string s;
                while (!sr.EndOfStream)
                {
                    s = sr.ReadLine().Trim();
                    if (s != string.Empty && 
                        s != "\n" && 
                        s != "\r\n" && 
                        s != "\t" && 
                        !s.StartsWith("'"))
                    {

                        s = ConvertTypeString(s);
                        s = ConverConstString(s);
                        s = RemoveEmpty(s);
                        sw.WriteLine(s);


                        //Console.WriteLine(s);
                    }
                }
            }
        }

        private string ConvertTypeString(string s)
        {
            return Regex.Replace(s, @"(\w+)(\(\)|\(\w+\s+To\s+\w+\))?\s+As\s+((?:Integer|Long|String|Single|Double|Currency)(?!\(\)))", m =>
            {
                return m.Groups[1].Value + GetTypeName(m.Groups[3].Value) + m.Groups[2].Value;
            });
        }
        private string ConverConstString(string s)
        {
            return Regex.Replace(s, @"(adInteger|adDate|adDBTimeStamp|adVarWChar|vbCritical|vbQuestion|vbOKCancel|vbYesNo|vbYesNoCancel|vbDefaultButton2|vbNarrow|vbLowerCase|vbKatakana)\b", m =>
            {
                return GetConstNumber(m.Groups[1].Value);
            });
        }
        private string RemoveEmpty(string s)
        {
            if (s.StartsWith("MultiUse"))
                return s;

            s = Regex.Replace(s, @"\s+=\s+", "=");
            return Regex.Replace(s, @"(?:\s+)?,\s+([^_])", m =>
            {
                return "," + m.Groups[1].Value;
            });
        }

        private static string GetTypeName(string s)
        {
            switch (s)
            {
                case "Long":
                    return "&";
                case "Integer":
                    return "%";
                case "Double":
                    return "#";
                case "String":
                    return "$";
                case "Single":
                    return "!";
                case "Currency":
                    return "@";
            }

            return s;
        }
        private static string GetConstNumber(string s)
        {
            switch (s)
            {
                case "adInteger": return "3";
                case "adDate": return "7";
                case "adDBTimeStamp": return "135";
                case "adVarWChar": return "202";
                case "vbCritical": return "16";
                case "vbNarrow": return "8";
                case "vbLowerCase": return "2";
                case "vbKatakana": return "16";
                case "vbQuestion": return "32";
                case "vbOKCancel": return "1";
                case "vbDefaultButton2": return "256";
                case "vbYesNo": return "4";
                case "vbYesNoCancel": return "3";
            }
            return s;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = this.folderBrowserDialog1.SelectedPath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(this.textBox1.Text))
            {
                this.FF(this.textBox1.Text);

                MessageBox.Show("•ÏŠ·I—¹‚µ‚Ü‚µ‚½");
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        private void Form1_DragOver(object sender, DragEventArgs e)
        {

        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (files.Length > 0 && !string.IsNullOrEmpty(files[0]) && Directory.Exists(files[0]))
            {
                this.textBox1.Text = files[0];
            }
        }
    }
}