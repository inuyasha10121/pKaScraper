using OfficeOpenXml;
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

        Dictionary<string, decimal> experimentalValues = new Dictionary<string, decimal>
        {
            {"Acetic", 4.75m},
            {"Chloracetic", 2.85m},
            {"Cyanoacetic", 2.45m},
            {"Formic", 3.75m},
            {"H3PO4_1", 2.15m},
            {"H3PO4_2", 7.20m},
            {"H3PO4_3", 12.35m},
            {"Morpholine", 8.36m},
            {"Oxalic_1", 1.23m},
            {"Oxalic_2", 4.19m},
            {"Pivalic", 5.03m},
            {"m-Chlorobenzoic", 3.83m},
            {"Methylhydroperoxide", 11.5m},
            {"m-Nitrobenzoic", 2.45m},
            {"N-Methylmorpholine", 7.38m},
            {"o-Chlorobenzoic", 2.94m},
            {"o-Nitrobenzoic", 2.17m},
            {"p-Chlorobenzoic", 3.99m},
            {"Peracetic", 8.2m},
            {"p-Nitrobenzoic", 3.44m},
            {"Quinuclidine", 11.0m},
            {"TBA", 17.0m},
            {"Trifluoroethanol", 12.5m}
        };

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
                var GgH = 1000000.0m;
                var GaqH = 1000000.0m;
                var filelines = File.ReadAllLines(sampledirs[hdirindex] + "\\gp\\H.out");
                foreach (var line in filelines)
                {
                    if (line.Contains("Sum of electronic and thermal Free Energies"))
                    {
                        GgH = decimal.Parse(line.Split((char[])null, StringSplitOptions.RemoveEmptyEntries)[7]);
                    }
                    if (line.Contains("imaginary frequencies (negative"))
                    {
                        File.AppendAllText(logfile, "Error: Hydrogen gas-phase has imaginary frequency somehow\n");
                        break;
                    }
                }
                if(GgH == 1000000.0m)
                {
                    File.AppendAllText(logfile, "Error: Hydrogen gas-phase energy not found!\n");
                }

                filelines = File.ReadAllLines(sampledirs[hdirindex] + "\\smd\\H.out");
                foreach (var line in filelines)
                {
                    if (line.Contains("Sum of electronic and thermal Free Energies"))
                    {
                        GaqH = decimal.Parse(line.Split((char[])null, StringSplitOptions.RemoveEmptyEntries)[7]);
                    }
                    if (line.Contains("imaginary frequencies (negative"))
                    {
                        File.AppendAllText(logfile, "Error: Hydrogen SMD has imaginary frequency somehow\n");
                        break;
                    }
                }
                if (GaqH == 1000000.0m)
                {
                    File.AppendAllText(logfile, "Error: Hydrogen SMD energy not found!\n");
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
                        if(gpouts.Length != smdouts.Length)
                        {
                            Console.WriteLine("ERROR:");
                            return;
                        }

                        for (int j = 0; j < gpouts.Length; ++j)
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
                                if (line.Contains("imaginary frequencies (negative"))
                                {
                                    File.AppendAllText(logfile, "Error: Gas-phase imaginary frequency: " + samples[samples.Count-1] + "_" + currout + "\n");
                                    break;
                                }

                            }
                            if (data[data.Count - 1][j][0] == 1000000.0m)
                            {
                                File.AppendAllText(logfile, "Error: Gas-phase energy not found: " + samples[samples.Count - 1] + "_" + currout + "\n");
                            }

                            outcontents = File.ReadAllLines(sampledirs[i] + "\\smd\\" + currout);
                            foreach (var line in outcontents)
                            {
                                if (line.Contains("Sum of electronic and thermal Free Energies"))
                                {
                                    data[data.Count - 1][j][1] = decimal.Parse(line.Split((char[])null, StringSplitOptions.RemoveEmptyEntries)[7]);
                                }
                                if (line.Contains("imaginary frequencies (negative"))
                                {
                                    File.AppendAllText(logfile, "Error: SMD imaginary frequency: " + samples[samples.Count - 1] + "_" + currout + "\n");
                                    break;
                                }
                            }
                            if (data[data.Count - 1][j][1] == 1000000.0m)
                            {
                                File.AppendAllText(logfile, "Error: SMD energy not found: " + samples[samples.Count - 1] + "_" + currout + "\n");
                            }

                        }
                    }
                }

                Console.WriteLine("Calculating pKa values when able...");
                var thermovals = new List<List<List<decimal>>>();
                for(int i = 0; i < data.Count; ++i)
                {
                    thermovals.Add(new List<List<decimal>>());
                    for (int j = 1; j < data[i].Count; ++j)
                    {
                        //dGsHA, dGsA, dGsH, dGg, dGaq, pKag, pKaaq, %Eg, %Eaq
                        thermovals[i].Add(new List<decimal>(new decimal[] { 1000000.0m, 1000000.0m, 1000000.0m, 1000000.0m, 1000000.0m, 1000000.0m, 1000000.0m, 1000000.0m, 1000000.0m }));

                        //Get relevant data values
                        var GgHA = data[i][j][0];
                        var GgA = data[i][j - 1][0];
                        var GaqHA = data[i][j][1];
                        var GaqA = data[i][j - 1][1];

                        //Calculate dGsHA
                        if (GgHA != 1000000.0m && GaqHA != 1000000.0m)
                        {
                            thermovals[i][j - 1][0] = GaqHA - GgHA;
                        }
                        //Calculate dGsA
                        if (GgA != 1000000.0m && GaqA != 1000000.0m)
                        {
                            thermovals[i][j - 1][1] = GaqA - GgA;
                        }
                        //Calculate dGsH
                        if (GgH != 1000000.0m && GaqH != 1000000.0m)
                        {
                            thermovals[i][j - 1][2] = GaqH - GgH;
                        }
                        //Calculate dGg
                        if (GgHA != 1000000.0m && GgA != 1000000.0m && GgH != 1000000.0m)
                        {
                            thermovals[i][j - 1][3] = (GgA + GgH) - GgHA;
                        }
                        //Calculate dGaq
                        if (GaqHA != 1000000.0m && GaqA != 1000000.0m && GaqH != 1000000.0m)
                        {
                            thermovals[i][j - 1][4] = (GaqA + GaqH) - GaqHA;
                        }
                        //Calculate pKag
                        if(thermovals[i][j - 1][0] != 1000000.0m && thermovals[i][j - 1][1] != 1000000.0m && thermovals[i][j - 1][2] != 1000000.0m && thermovals[i][j - 1][3] != 1000000.0m)
                        {
                            thermovals[i][j - 1][5] = (thermovals[i][j - 1][3] + (1.9872036E-3m * 298.15m * (decimal)Math.Log(24.46))) / (2.303m * 1.9872036E-3m * 298.15m);
                        }
                        //Calculate pKaaq
                        if (thermovals[i][j-1][4] != 1000000.0m)
                        {
                            thermovals[i][j - 1][6] = thermovals[i][j - 1][4] / (2.303m * 1.9872036E-3m * 298.15m);
                        }
                        var expkey = samples[i];
                        if(data[i].Count != 2)
                        {
                            expkey += "_" + j;
                        }
                        var expval = experimentalValues[expkey];
                        //Calculate %Eg
                        if(thermovals[i][j-1][5] != 1000000.0m)
                        {
                            
                        }
                    }
                }

                Console.WriteLine("Writing data to file...");
                string output = "Compound:,State:,Gg:,Gaq:,\n";
                output += "H,H," + GgH + "," + GaqH + ",\n";
                for (int i = 0; i < data.Count; ++i)
                {
                    for (int j = 0; j < data[i].Count; ++j)
                    {
                        if (j == 0)
                            output += samples[i] + ",A,";
                        else if (j == 1)
                            output += ",HA,";
                        else
                            output += ",H" + j + "A,";

                        if(data[i][j][0] != 1000000.0m)
                        {
                            output += data[i][j][0];
                        }
                        output += ",";
                        if (data[i][j][1] != 1000000.0m)
                        {
                            output += data[i][j][1];
                        }
                        output += ",\n";
                    }
                }
                File.WriteAllText(savefile, output);
            }
            Console.WriteLine("Done!");
        }
    }
}
