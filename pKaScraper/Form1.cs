﻿using OfficeOpenXml;
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

namespace pKaScraper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonRun_Click(object sender, EventArgs e)
        {
            if (textBoxPath.Text == String.Empty)
                return;
            //var result = saveFileDialog1.ShowDialog();
            //if (result != DialogResult.OK)
            //    return;
            var methoddirs = Directory.GetDirectories(textBoxPath.Text);
            var logfile = textBoxPath.Text + "\\log.txt";
            File.WriteAllText(logfile, "Starting processing...\n");
            foreach (var method in methoddirs)
            {
                File.AppendAllText(logfile, method + "\n");
                //var savefile = saveFileDialog1.FileName;
                var savefile = textBoxPath.Text + "\\" + method.Substring(method.LastIndexOf("\\")) + ".csv";
                var sampledirs = Directory.GetDirectories(method);

                //Find the hydrogen run in the folder
                var hdirindex = -1;
                for (int i = 0; i < sampledirs.Length; ++i)
                {
                    if (sampledirs[i].Substring(sampledirs[i].LastIndexOf('\\') + 1).Equals("H"))
                        hdirindex = i;
                }
                if (hdirindex == -1)
                {
                    File.AppendAllText(logfile, "Error: No hydrogen simulation found!\n");
                }

                //Get the hydrogen run values
                if (!File.Exists(sampledirs[hdirindex] + "\\gp\\H.out"))
                {
                    File.AppendAllText(logfile, "Error: Hydrogen gas-phase output not found!\n");
                }
                if (!File.Exists(sampledirs[hdirindex] + "\\smd\\H.out"))
                {
                    File.AppendAllText(logfile, "Error: Hydrogen SMD output not found!\n");
                }
                var GgH = 1.0m;
                var GaqH = 1.0m;
                var filelines = File.ReadAllLines(sampledirs[hdirindex] + "\\gp\\H.out");
                foreach (var line in filelines)
                {
                    if (line.Contains("Sum of electronic and thermal Free Energies"))
                    {
                        GgH = decimal.Parse(line.Split((char[])null, StringSplitOptions.RemoveEmptyEntries)[7]);
                    }
                    if (line.Contains("imaginary frequenc"))
                    {
                        File.AppendAllText(logfile, "Error: Hydrogen gas-phase has imaginary frequency somehow\n");
                    }
                }

                filelines = File.ReadAllLines(sampledirs[hdirindex] + "\\smd\\H.out");
                foreach (var line in filelines)
                {
                    if (line.Contains("Sum of electronic and thermal Free Energies"))
                    {
                        GaqH = decimal.Parse(line.Split((char[])null, StringSplitOptions.RemoveEmptyEntries)[7]);
                    }
                    if (line.Contains("imaginary frequenc"))
                    {
                        File.AppendAllText(logfile, "Error: Hydrogen SMD has imaginary frequency somehow\n");
                        return;
                    }
                }
                Console.WriteLine("GgH: " + GgH + "\tGaqH: " + GaqH);
                Console.WriteLine("Getting Data...");
                var data = new List<List<List<decimal>>>();
                var samples = new List<string>();
                for (int i = 0; i < sampledirs.Length; ++i)
                {
                    if (i != hdirindex)
                    {
                        data.Add(new List<List<decimal>>());
                        samples.Add(sampledirs[i].Substring(sampledirs[i].LastIndexOf("\\") + 1));
                        var gpouts = Directory.GetFiles(sampledirs[i] + "\\gp\\", "*.out");
                        var smdouts = Directory.GetFiles(sampledirs[i] + "\\smd\\", "*.out");
                        var outcount = Math.Max(gpouts.Length, smdouts.Length);

                        for (int j = 0; j < outcount; ++j)
                        {
                            data[data.Count - 1].Add(new List<decimal>());
                            var currout = "";
                            if (j == 0)
                            {
                                currout = "A.out";
                            }
                            else if (j == 1)
                            {
                                currout = "HA.out";
                            }
                            else
                            {
                                currout = "H" + j + "A.out";
                            }

                            data[data.Count - 1][j].Add(1000000.0m);
                            data[data.Count - 1][j].Add(1000000.0m);
                            var outcontents = File.ReadAllLines(sampledirs[i] + "\\gp\\" + currout);
                            foreach (var line in outcontents)
                            {
                                if (line.Contains("Sum of electronic and thermal Free Energies"))
                                {
                                    data[data.Count - 1][j][0] = decimal.Parse(line.Split((char[])null, StringSplitOptions.RemoveEmptyEntries)[7]);
                                }
                                if (line.Contains("imaginary frequenc"))
                                {
                                    File.AppendAllText(logfile, "Error: Imaginary frequency: " + gpouts[j] + "\n");
                                }

                            }
                            outcontents = File.ReadAllLines(sampledirs[i] + "\\smd\\" + currout);
                            foreach (var line in outcontents)
                            {
                                if (line.Contains("Sum of electronic and thermal Free Energies"))
                                {
                                    data[data.Count - 1][j][1] = decimal.Parse(line.Split((char[])null, StringSplitOptions.RemoveEmptyEntries)[7]);
                                }
                                if (line.Contains("imaginary frequenc"))
                                {
                                    File.AppendAllText(logfile, "Error: Imaginary frequency: " + gpouts[j] + "\n");
                                }
                            }
                        }
                    }
                }

                Console.WriteLine("Processing Data");
                string output = "Compound:,State:,Gg:,Gaq:,\n";
                output += "H,H," + GgH + "," + GaqH + ",\n";
                for (int i = 0; i < data.Count; ++i)
                {
                    for (int j = 0; j < data[i].Count; ++j)
                    {
                        if (j == 0)
                            output += samples[i] + ",A," + data[i][j][0] + "," + data[i][j][1] + ",\n";
                        else if (j == 1)
                            output += ",HA," + data[i][j][0] + "," + data[i][j][1] + ",\n";
                        else
                            output += ",H" + j + "A," + data[i][j][0] + "," + data[i][j][1] + ",\n";
                    }
                }
                File.WriteAllText(savefile, output);
            }
            Console.WriteLine("Done!");
        }
    }
}