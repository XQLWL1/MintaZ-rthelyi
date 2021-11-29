using Olympics.Entities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Olympics
{
    public partial class Form1 : Form
    {

        List<OlympicResult> results = new List<OlympicResult>();

        public Form1()
        {
            InitializeComponent();

            Betolt("Summer_olympic_Medals.csv");
            ComboFeltolt();

            Osztalyozas();

            dataGridView1.DataSource = results;
        }

        private void Osztalyozas()
        {
            foreach (OlympicResult item in results)
            {
                item.Position = Helyezes(item);
            }
        }

        int Helyezes(OlympicResult olympicResult)
        {
            int counter = 0;
            var szurt = from x in results where x.Year == olympicResult.Year && x.Country != olympicResult.Country select x;

            foreach (OlympicResult item in szurt)
            {
                if (item.Medals[0] > olympicResult.Medals[0])
                {
                    counter++;
                }

                else
                {
                    if ((item.Medals[0] == olympicResult.Medals[0]) && (item.Medals[1] <= olympicResult.Medals[1]))
                    {
                        counter++;
                    }

                    else
                    {
                        if ((item.Medals[0] == olympicResult.Medals[0])
                            && (item.Medals[1] == olympicResult.Medals[1])
                            && (item.Medals[2] > olympicResult.Medals[2]))
                        {
                            counter++;
                        }

                    }
                }


            }

            return counter + 1;

        }

        private void ComboFeltolt()
        {
            var years = (from x in results orderby x.Year select x.Year).Distinct();
            comboBoxEv.DataSource = years.ToList();
        }

        private void Betolt(string fajlnev)
        {

            using (StreamReader streamReader = new StreamReader(fajlnev))
            {
                streamReader.ReadLine();

                while (!streamReader.EndOfStream)
                {
                    string sor = streamReader.ReadLine();

                    string[] mezok = sor.Split(',');

                    OlympicResult or = new OlympicResult();
                    or.Year = int.Parse(mezok[0]);
                    or.Country = mezok[3];

                    //Medálaknak 3 elemű tömböt kell létrehozni
                    int[] medaltomb = new int[3];

                    //5,6,7 lesz az arany, ezüt, bronz
                    medaltomb[0] = int.Parse(mezok[5]);
                    medaltomb[1] = int.Parse(mezok[6]);
                    medaltomb[2] = int.Parse(mezok[7]);

                    or.Medals = medaltomb;

                    results.Add(or);
                }

            }

        }
    }
}
