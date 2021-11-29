using Olympics.Entities;
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
        }

        private void ComboFeltolt()
        {
            var years = (from x in results orderby x.Year select x.Year).Distinct();
            comboBoxEv.DataSource = years.ToList();
        }

        void Betolt(string fajlnev)
        {

           using (StreamReader streamReader = new StreamReader(fajlnev))
            {
                streamReader.ReadLine();

                while (!streamReader.EndOfStream)
                {
                    string sor = streamReader.ReadLine();

                    string[] mezok = sor.Split(',');

                    OlympicResult olympicResult = new OlympicResult();
                    olympicResult.Year = int.Parse(mezok[0]);
                    olympicResult.Country = mezok[3];

                    //Medálaknak 3 elemű tömböt kell létrehozni
                    int[] medaltomb = new int[3];

                    //5,6,7 lesz az arany, ezüt, bronz
                    medaltomb[0] = int.Parse(mezok[5]);
                    medaltomb[1] = int.Parse(mezok[6]);
                    medaltomb[2] = int.Parse(mezok[7]);

                    results.Add(olympicResult);
                }

            }

        }

    }
}
