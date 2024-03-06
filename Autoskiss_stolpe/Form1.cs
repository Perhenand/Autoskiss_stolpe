using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using Tekla.Structures.Model;
using Tekla.Structures.Drawing;
using T3D = Tekla.Structures.Geometry3d;
using TSMUI = Tekla.Structures.Model.UI;
using MathNet.Numerics.LinearAlgebra;


namespace Autoskiss_stolpe
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //myModel = new Model();
        }

        Model myModel = new Model();
        List<Guid> tilltek_GUID = new List<Guid>();
        List<string> tilltek_idnr = new List<string>();
        List<string> tilltek_startnod = new List<string>();
        List<string> tilltek_slutnod = new List<string>();
        List<string> tilltek_startPl = new List<string>();
        List<string> tilltek_slutPl = new List<string>();
        List<string> tilltek_typ = new List<string>();
        List<string> tilltek_dim = new List<string>();
        List<string> tilltek_kvalitet = new List<string>();
        List<double> tilltek_rotera = new List<double>();
        List<int> tilltek_pos = new List<int>();
        List<double> tilltek_x1 = new List<double>();
        List<double> tilltek_y1 = new List<double>();
        List<double> tilltek_z1 = new List<double>();
        List<double> tilltek_x2 = new List<double>();
        List<double> tilltek_y2 = new List<double>();
        List<double> tilltek_z2 = new List<double>();
        List<int> tilltek_blt = new List<int>();
        List<int> tilltek_bltN = new List<int>();
        List<double> tilltek_blte1 = new List<double>();
        List<double> tilltek_bltp1 = new List<double>();
        List<double> tilltek_bltd0 = new List<double>();

        List<string> tilltekPl_idnr = new List<string>();
        List<Guid> tilltekPl_GUID = new List<Guid>();
        List<int> tilltekPl_typ = new List<int>();
        List<double> tilltekPl_t = new List<double>();
        List<string> tilltekPl_ram = new List<string>();
        List<double> tilltekPl_ramc = new List<double>();
        List<double> tilltekPl_ramb = new List<double>();
        List<double> tilltekPl_ramt = new List<double>();
        List<string[]> tilltekPl_diag = new List<string[]>();
        List<int> tilltekPl_blt = new List<int>();
        List<int> tilltekPl_bltN = new List<int>();
        List<double> tilltekPl_blte1 = new List<double>();
        List<double> tilltekPl_bltp1 = new List<double>();
        List<double> tilltekPl_bltd0 = new List<double>();
        List<double> tilltekPl_cntr = new List<double>();

        private void Form1_Load(object sender, EventArgs e)
        {
            RegistreraStartvarden();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Knapp 1");
        }

        // Registrerar de nominella indata-värden som visas när skriptet startas.
        private void RegistreraStartvarden()
        {
            textBox1.Text = "5";
            dataGridView1.Rows.Add(1, 2, null, null, 2, null, null, null, null, null, 0.1, null, null, null, null, 5, 1);
            dataGridView1.Rows.Add(2, 2, null, null, 2, null, 0.4, null, null, null, 0.1, null, null, null, null, 5, 1);
            dataGridView1.Rows.Add(3, 1, null, null, 3, null, null, null, 0.1, 0.1, 0.1, null, null, null, null, 5, 1);
            dataGridView1.Rows.Add(4, 1, null, null, 2.5, null, 0.4, 0.1, 0.1, 0.1, 0.1, null, null, null, null, 5, 1);
            dataGridView1.Rows.Add(5, 1, null, null, 2, 1.6, null, 0.1, 0.1, 0.1, 0.1, null, null, null, null, 5, 1);
            dataGridView1.Rows.Add(6, 1, null, null, 1.5, null, 0.4, 0.1, null, null, 0.1, null, null, null, null, 5, 1);
            dataGridView1.Rows.Add(7, 1, null, null, 1.5, null, null, 0.1, null, null, 0.1, null, null, null, null, 5, 1);
            dataGridView1.Rows.Add(8, 1, null, null, 1.5, null, null, 0.1, null, null, 0.1, null, null, null, null, 5, 1);
            dataGridView1.Rows.Add(9, 1, null, null, 1.5, null, 0.4, 0.1, null, null, 0.1, null, null, null, null, 5, 1);
            dataGridView1.Rows.Add(10, 1, null, null, 1.5, null, null, 0.05, null, null, 0.05, null, null, null, null, 5, 1);
            dataGridView1.Rows.Add(11, 1, null, null, 1.5, null, null, null, null, null, null, null, null, null, null, 5, 1);
            dataGridView1.Rows.Add(12, 1, null, null, 1.5, 0.8, null, null, null, null, null, null, null, null, null, 5, 1);
            dataGridView2.Rows.Add(18, 10, null, 10);
            dataGridView2.Rows.Add(18, 10, null, 10);
            dataGridView2.Rows.Add(16, 10, 10, null);
            dataGridView2.Rows.Add(16, 10, 10, null);
            dataGridView2.Rows.Add(15, 8, 8, 8);
            dataGridView2.Rows.Add(15, 8, null, 8);
            dataGridView2.Rows.Add(14, 8, null, 8);
            dataGridView2.Rows.Add(14, 8, null, 8);
            dataGridView2.Rows.Add(14, 8, null, 8);
            dataGridView2.Rows.Add(12, 8, null, 8);
            dataGridView2.Rows.Add(12, 8, null, 8);
            dataGridView2.Rows.Add(12, 8, null, 8);
            dataGridView3.Rows.Add(3, 2, 2, 2);
            dataGridView3.Rows.Add(3, 2, 2, 2);
            dataGridView3.Rows.Add(3, 2, 2, null);
            dataGridView3.Rows.Add(3, 2, 2, null);
            dataGridView3.Rows.Add(3, 2, 2, 2);
            dataGridView3.Rows.Add(3, 2, 2, 2);
            dataGridView3.Rows.Add(3, 2, 2, 2);
            dataGridView3.Rows.Add(3, 2, 2, 2);
            dataGridView3.Rows.Add(3, 2, 2, 2);
            dataGridView3.Rows.Add(3, 2, 2, 2);
            dataGridView3.Rows.Add(3, 2, 2, 2);
            dataGridView3.Rows.Add(3, 1, 1, 2);
            dataGridView4.Rows.Add(3, 2, 1, 1);
            dataGridView4.Rows.Add(3, 2, 1, 1);
            dataGridView4.Rows.Add(3, 2, 1, null);
            dataGridView4.Rows.Add(3, 2, 1, null);
            dataGridView4.Rows.Add(3, 2, 1, 1);
            dataGridView4.Rows.Add(3, 2, 1, 1);
            dataGridView4.Rows.Add(3, 2, 1, 1);
            dataGridView4.Rows.Add(3, 2, 1, 1);
            dataGridView4.Rows.Add(3, 2, 1, 1);
            dataGridView4.Rows.Add(3, 2, 1, 1);
            dataGridView4.Rows.Add(3, 1, 1, 1);
            dataGridView4.Rows.Add(3, 1, 1, 1);
            dataGridView5.Rows.Add(0, 0);
            dataGridView5.Rows.Add(0, 0);
            dataGridView5.Rows.Add(0, 0);
            dataGridView5.Rows.Add(0, 0);
            dataGridView5.Rows.Add(0, 0);
            dataGridView5.Rows.Add(0, 0);
            dataGridView5.Rows.Add(0, 0);
            dataGridView5.Rows.Add(0, 0);
            dataGridView5.Rows.Add(0, 0);
            dataGridView5.Rows.Add(0, 0);
            dataGridView5.Rows.Add(0, 0);
            dataGridView5.Rows.Add(0, 0);
            dataGridView11.Rows.Add(null, null, 12);
            dataGridView11.Rows.Add(null, null, 12);
            dataGridView11.Rows.Add(10, 12, null);
            dataGridView11.Rows.Add(10, 12, null);
            dataGridView11.Rows.Add(10, 12, null);
            dataGridView11.Rows.Add(10, null, null);
            dataGridView11.Rows.Add(10, null, null);
            dataGridView11.Rows.Add(10, null, null);
            dataGridView11.Rows.Add(null, null, null);
            dataGridView11.Rows.Add(null, null, null);
            dataGridView11.Rows.Add(null, null, null);
            dataGridView11.Rows.Add(null, null, null);
            dataGridView10.Rows.Add(null, null, 2);
            dataGridView10.Rows.Add(null, null, 2);
            dataGridView10.Rows.Add(2, 2, null);
            dataGridView10.Rows.Add(2, 2, null);
            dataGridView10.Rows.Add(2, 2, null);
            dataGridView10.Rows.Add(2, null, null);
            dataGridView10.Rows.Add(2, null, null);
            dataGridView10.Rows.Add(2, null, null);
            dataGridView10.Rows.Add(null, null, null);
            dataGridView10.Rows.Add(null, null, null);
            dataGridView10.Rows.Add(null, null, null);
            dataGridView10.Rows.Add(null, null, null);
            dataGridView9.Rows.Add(null, null, 3);
            dataGridView9.Rows.Add(null, null, 3);
            dataGridView9.Rows.Add(3, 3, null);
            dataGridView9.Rows.Add(3, 3, null);
            dataGridView9.Rows.Add(3, 3, null);
            dataGridView9.Rows.Add(3, null, null);
            dataGridView9.Rows.Add(3, null, null);
            dataGridView9.Rows.Add(3, null, null);
            dataGridView9.Rows.Add(null, null, null);
            dataGridView9.Rows.Add(null, null, null);
            dataGridView9.Rows.Add(null, null, null);
            dataGridView9.Rows.Add(null, null, null);
            textBox2.Text = "2";
            textBox3.Text = "1,5";
            textBox4.Text = "1,5";
            textBox5.Text = "3";
            textBox6.Text = "10";
            textBox7.Text = "10";
            textBox8.Text = "1";
            textBox9.Text = "10";
            dataGridView7.Rows.Add(1, "M12", "12", "8.8", 12, 13.5, 18.7);
            dataGridView7.Rows.Add(2, "M16", "16", "8.8", 16, 17.5, 24, 7);
            dataGridView7.Rows.Add(3, "M20", "20", "8.8", 20, 21.5, 31.1);
            dataGridView7.Rows.Add(4, "M24", "24", "8.8", 24, 26, 37.3);
            dataGridView7.Rows.Add(5, "M30", "30", "8.8", 30, 32, 47.9);
            dataGridView8.Rows.Add(1, "L45x5", "L45x5", "S355J2", 45, 5, 12.8);
            dataGridView8.Rows.Add(2, "L50x5", "L50x5", "S355J2", 50, 5, 14);
            dataGridView8.Rows.Add(3, "L55x6", "L55x6", "S355J2", 55, 6, 15.6);
            dataGridView8.Rows.Add(4, "L60x6", "L60x6", "S355J2", 60, 6, 16.9);
            dataGridView8.Rows.Add(5, "L65x7", "L65x7", "S355J2", 65, 7, 18.5);
            dataGridView8.Rows.Add(6, "L70x7", "L70x7", "S355J2", 70, 7, 19.7);
            dataGridView8.Rows.Add(7, "L75x7", "L75x7", "S355J2", 75, 7, 20.9);
            dataGridView8.Rows.Add(8, "L80x8", "L80x8", "S355J2", 80, 8, 22.6);
            dataGridView8.Rows.Add(9, "L90x9", "L90x9", "S355J2", 90, 9, 25.4);
            dataGridView8.Rows.Add(10, "L100x10", "L100x10", "S355J2", 100, 10, 28.2);
            dataGridView8.Rows.Add(11, "L110x10", "L110x10", "S355J2", 110, 10, 30.7);
            dataGridView8.Rows.Add(12, "L120x11", "L120x11", "S355J2", 120, 11, 33.6);
            dataGridView8.Rows.Add(13, "L130x12", "L130x12", "S355J2", 130, 12, 36.4);
            dataGridView8.Rows.Add(14, "L140x13", "L140x13", "S355J2", 140, 13, 39.2);
            dataGridView8.Rows.Add(15, "L150x14", "L150x14", "S355J2", 150, 14, 42.1);
            dataGridView8.Rows.Add(16, "L160x15", "L160x15", "S355J2", 160, 15, 44.9);
            dataGridView8.Rows.Add(17, "L180x16", "L180x16", "S355J2", 180, 16, 50.2);
            dataGridView8.Rows.Add(18, "L180x18", "L180x18", "S355J2", 180, 18, 51);
            dataGridView8.Rows.Add(19, "L200x16", "L200x16", "S355J2", 200, 16, 55.2);
            dataGridView8.Rows.Add(20, "L200x18", "L200x18", "S355J2", 200, 18, 56);
        }

        // Skapa ny rad, dataGridView1-5
        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 1 && dataGridView1.Rows.Count > 1)
            {
                int index = dataGridView1.CurrentCell.RowIndex;
                int antrader = dataGridView1.Rows.Count;
                DataGridViewRow dr = new DataGridViewRow();
                dr.CreateCells(dataGridView1);
                dr.Cells[1].Value = dataGridView1.Rows[index].Cells[1].Value;
                dr.Cells[2].Value = dataGridView1.Rows[index].Cells[2].Value;
                dr.Cells[3].Value = dataGridView1.Rows[index].Cells[3].Value;
                dr.Cells[4].Value = dataGridView1.Rows[index].Cells[4].Value;
                dr.Cells[5].Value = dataGridView1.Rows[index].Cells[5].Value;
                dr.Cells[6].Value = dataGridView1.Rows[index].Cells[6].Value;
                dr.Cells[7].Value = dataGridView1.Rows[index].Cells[7].Value;
                dr.Cells[8].Value = dataGridView1.Rows[index].Cells[8].Value;
                dr.Cells[9].Value = dataGridView1.Rows[index].Cells[9].Value;
                dr.Cells[10].Value = dataGridView1.Rows[index].Cells[10].Value;
                dr.Cells[11].Value = dataGridView1.Rows[index].Cells[11].Value;
                dr.Cells[12].Value = dataGridView1.Rows[index].Cells[12].Value;
                dr.Cells[13].Value = dataGridView1.Rows[index].Cells[13].Value;
                dr.Cells[14].Value = dataGridView1.Rows[index].Cells[14].Value;
                dr.Cells[15].Value = dataGridView1.Rows[index].Cells[15].Value;
                dr.Cells[16].Value = dataGridView1.Rows[index].Cells[16].Value;
                dataGridView1.Rows.Insert(index + 1, dr);
                for (int i = 0; i < antrader + 1; i++)
                {
                    dataGridView1.Rows[i].Cells[0].Value = i + 1;
                }

                dr = new DataGridViewRow();
                dr.CreateCells(dataGridView2);
                dr.Cells[0].Value = dataGridView2.Rows[index].Cells[0].Value;
                dr.Cells[1].Value = dataGridView2.Rows[index].Cells[1].Value;
                dr.Cells[2].Value = dataGridView2.Rows[index].Cells[2].Value;
                dr.Cells[3].Value = dataGridView2.Rows[index].Cells[3].Value;
                dataGridView2.Rows.Insert(index + 1, dr);

                dr = new DataGridViewRow();
                dr.CreateCells(dataGridView3);
                dr.Cells[0].Value = dataGridView3.Rows[index].Cells[0].Value;
                dr.Cells[1].Value = dataGridView3.Rows[index].Cells[1].Value;
                dr.Cells[2].Value = dataGridView3.Rows[index].Cells[2].Value;
                dr.Cells[3].Value = dataGridView3.Rows[index].Cells[3].Value;
                dataGridView3.Rows.Insert(index + 1, dr);

                dr = new DataGridViewRow();
                dr.CreateCells(dataGridView4);
                dr.Cells[0].Value = dataGridView4.Rows[index].Cells[0].Value;
                dr.Cells[1].Value = dataGridView4.Rows[index].Cells[1].Value;
                dr.Cells[2].Value = dataGridView4.Rows[index].Cells[2].Value;
                dr.Cells[3].Value = dataGridView4.Rows[index].Cells[3].Value;
                dataGridView4.Rows.Insert(index + 1, dr);

                dr = new DataGridViewRow();
                dr.CreateCells(dataGridView5);
                dr.Cells[0].Value = dataGridView5.Rows[index].Cells[0].Value;
                dr.Cells[1].Value = dataGridView5.Rows[index].Cells[1].Value;
                dataGridView5.Rows.Insert(index + 1, dr);

                dr = new DataGridViewRow();
                dr.CreateCells(dataGridView9);
                dr.Cells[0].Value = dataGridView9.Rows[index].Cells[0].Value;
                dr.Cells[1].Value = dataGridView9.Rows[index].Cells[1].Value;
                dr.Cells[2].Value = dataGridView9.Rows[index].Cells[2].Value;
                dataGridView9.Rows.Insert(index + 1, dr);

                dr = new DataGridViewRow();
                dr.CreateCells(dataGridView10);
                dr.Cells[0].Value = dataGridView10.Rows[index].Cells[0].Value;
                dr.Cells[1].Value = dataGridView10.Rows[index].Cells[1].Value;
                dr.Cells[2].Value = dataGridView10.Rows[index].Cells[2].Value;
                dataGridView10.Rows.Insert(index + 1, dr);

                dr = new DataGridViewRow();
                dr.CreateCells(dataGridView11);
                dr.Cells[0].Value = dataGridView11.Rows[index].Cells[0].Value;
                dr.Cells[1].Value = dataGridView11.Rows[index].Cells[1].Value;
                dr.Cells[2].Value = dataGridView11.Rows[index].Cells[2].Value;
                dataGridView11.Rows.Insert(index + 1, dr);
            }
        }

        // Tar bort rad, dataGridView1-5
        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 1 && dataGridView1.Rows.Count > 1)
            {
                int index = dataGridView1.CurrentCell.RowIndex;
                int antrader = dataGridView1.Rows.Count;
                dataGridView1.Rows.RemoveAt(index);
                for (int i = 0; i < antrader - 1; i++)
                {
                    dataGridView1.Rows[i].Cells[0].Value = i + 1;
                }
                dataGridView2.Rows.RemoveAt(index);
                dataGridView3.Rows.RemoveAt(index);
                dataGridView4.Rows.RemoveAt(index);
                dataGridView5.Rows.RemoveAt(index);
            }
        }

        // Numrera om rader när rad skapas eller tas bort
        private void dataGridView7_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            int antrader = dataGridView7.Rows.Count;
            for (int i = 0; i < antrader - 1; i++)
            {
                dataGridView7.Rows[i].Cells[0].Value = i + 1;
            }
        }

        // Numrera om rader när rad skapas eller tas bort
        private void dataGridView7_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            int antrader = dataGridView7.Rows.Count;
            for (int i = 0; i < antrader - 1; i++)
            {
                dataGridView7.Rows[i].Cells[0].Value = i + 1;
            }
        }

        // Numrera om rader när rad skapas eller tas bort
        private void dataGridView8_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            int antrader = dataGridView8.Rows.Count;
            for (int i = 0; i < antrader - 1; i++)
            {
                dataGridView8.Rows[i].Cells[0].Value = i + 1;
            }
        }

        // Numrera om rader när rad skapas eller tas bort
        private void dataGridView8_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            int antrader = dataGridView8.Rows.Count;
            for (int i = 0; i < antrader - 1; i++)
            {
                dataGridView8.Rows[i].Cells[0].Value = i + 1;
            }
        }

        // Sparar indatavärden till fil
        private void sparaSomToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sparafil1 = new SaveFileDialog();
            sparafil1.Filter = "Normal text file|*.txt";
            sparafil1.Title = "Spara som";
            sparafil1.InitialDirectory = Directory.GetCurrentDirectory();
            // sparafil1.InitialDirectory = Path.GetDirectoryName(Application.ExecutablePath);
            sparafil1.ShowDialog();

            if (sparafil1.FileName != "")
            {
                var delimiter = "\t";
                int antrader;
                int antkol;
                String textrad;
                using (StreamWriter sw = new StreamWriter(sparafil1.FileName))
                {
                    sw.WriteLine(" Nr Värde  Parameter");
                    sw.WriteLine("---------------------------------------------------------------------");
                    sw.WriteLine("1" + delimiter + textBox1.Text + delimiter + delimiter + delimiter + "Basbredd [m]");
                    sw.WriteLine("2" + delimiter + textBox2.Text + delimiter + delimiter + delimiter + "Minsta kantavstånd e1, ben, (e1/d0) [mm]");
                    sw.WriteLine("3" + delimiter + textBox3.Text + delimiter + delimiter + delimiter + "Minsta kantavstånd e1, (e1/d0) [mm]");
                    sw.WriteLine("4" + delimiter + textBox4.Text + delimiter + delimiter + delimiter + "Minsta kantavstånd e2, (e2/d0) [mm]");
                    sw.WriteLine("5" + delimiter + textBox5.Text + delimiter + delimiter + delimiter + "Minsta kantavstånd p1, (p1/d0) [mm]");
                    sw.WriteLine("6" + delimiter + checkBox1.Checked + delimiter + delimiter + delimiter + "Kontinuerliga horisontaler vid anslutning mot korsande diagonaler");
                    sw.WriteLine("7" + delimiter + textBox6.Text + delimiter + delimiter + delimiter + "Minsta kantavstånd stänger, skarv av ben [mm]");
                    sw.WriteLine("8" + delimiter + textBox7.Text + delimiter + delimiter + delimiter + "Minsta kantavstånd stänger, övriga skarvar [mm]");
                    sw.WriteLine("9" + delimiter + textBox9.Text + delimiter + delimiter + delimiter + "Minsta kantavstånd stänger, övriga stänger [mm]");
                    sw.WriteLine("");
                    sw.WriteLine(" Bultdata");
                    sw.WriteLine("---------------------------------------------------------------------");
                    antrader = dataGridView7.Rows.Count;
                    antkol = dataGridView7.Columns.Count;
                    for (int i = 0; i < antrader; i++)
                    {
                        textrad = "";
                        for (int j = 0; j < antkol; j++)
                        {
                            textrad = textrad + dataGridView7.Rows[i].Cells[j].Value + delimiter;
                        }
                        sw.WriteLine(textrad);
                    }
                    sw.WriteLine("");
                    sw.WriteLine(" Tvärsnittsdata");
                    sw.WriteLine("---------------------------------------------------------------------");
                    antrader = dataGridView8.Rows.Count;
                    antkol = dataGridView8.Columns.Count;
                    for (int i = 0; i < antrader; i++)
                    {
                        textrad = "";
                        for (int j = 0; j < antkol; j++)
                        {
                            textrad = textrad + dataGridView8.Rows[i].Cells[j].Value + delimiter;
                        }
                        sw.WriteLine(textrad);
                    }
                    sw.WriteLine("");
                    sw.WriteLine(" Utformning och geometri");
                    sw.WriteLine("---------------------------------------------------------------------");
                    antrader = dataGridView1.Rows.Count;
                    antkol = dataGridView1.Columns.Count;
                    for (int i = 0; i < antrader; i++)
                    {
                        textrad = "";
                        for (int j = 0; j < antkol; j++)
                        {
                            textrad = textrad + dataGridView1.Rows[i].Cells[j].Value + delimiter;
                        }
                        sw.WriteLine(textrad);
                    }
                    sw.WriteLine("");
                    sw.WriteLine(" Stänger och bultar: profiler");
                    sw.WriteLine("---------------------------------------------------------------------");
                    antrader = dataGridView2.Rows.Count;
                    antkol = dataGridView2.Columns.Count;
                    for (int i = 0; i < antrader; i++)
                    {
                        textrad = "";
                        for (int j = 0; j < antkol; j++)
                        {
                            textrad = textrad + dataGridView2.Rows[i].Cells[j].Value + delimiter;
                        }
                        sw.WriteLine(textrad);
                    }
                    sw.WriteLine("");
                    sw.WriteLine(" Stänger och bultar: typ av bult");
                    sw.WriteLine("---------------------------------------------------------------------");
                    antrader = dataGridView3.Rows.Count;
                    antkol = dataGridView3.Columns.Count;
                    for (int i = 0; i < antrader; i++)
                    {
                        textrad = "";
                        for (int j = 0; j < antkol; j++)
                        {
                            textrad = textrad + dataGridView3.Rows[i].Cells[j].Value + delimiter;
                        }
                        sw.WriteLine(textrad);
                    }
                    sw.WriteLine("");
                    sw.WriteLine(" Stänger och bultar: antal bult");
                    sw.WriteLine("---------------------------------------------------------------------");
                    antrader = dataGridView4.Rows.Count;
                    antkol = dataGridView4.Columns.Count;
                    for (int i = 0; i < antrader; i++)
                    {
                        textrad = "";
                        for (int j = 0; j < antkol; j++)
                        {
                            textrad = textrad + dataGridView4.Rows[i].Cells[j].Value + delimiter;
                        }
                        sw.WriteLine(textrad);
                    }
                }
            }
        }

        // Importera indatavärden från fil
        // EJ FÄRDIG !!!
        private void öppnaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog oppnafil1 = new OpenFileDialog();
            oppnafil1.Filter = "Normal text file|*.txt";
            oppnafil1.Title = "Spara som";
            if (oppnafil1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                MessageBox.Show("Error!! Varning!!");
            }
        }

        // Ritar upp stolpe i Tekla
        private void button2_Click(object sender, EventArgs e)
        {
            if (Kvalkontroll())
            {
                IndatatTillTekla();
                RitaTekla();
            }
            else
            {
                IndatatTillTekla();
                RitaTekla();
                //MessageBox.Show("Felaktig indata");
            }
        }

        // Utför kvalitetskontroll av indata
        private bool Kvalkontroll()
        {
            bool indik1 = true;
            int rad1 = dataGridView1.Rows.Count;
            int rad2 = dataGridView2.Rows.Count;
            int rad3 = dataGridView3.Rows.Count;
            int rad4 = dataGridView4.Rows.Count;
            int rad7 = dataGridView7.Rows.Count;
            int rad8 = dataGridView8.Rows.Count;
            decimal tabellvarde1, tabellvarde2, tabellvarde3, tabellvarde4, tabellvarde5, tabellvarde6, tabellvarde7, tabellvarde8, tabellvarde9;
            decimal tabellvarde10, tabellvarde11, tabellvarde12, tabellvarde13, tabellvarde14, tabellvarde15, tabellvarde16, tabellvarde17;
            tabellvarde1 = tabellvarde2 = tabellvarde3 = tabellvarde4 = tabellvarde5 = tabellvarde6 = tabellvarde7 = tabellvarde8 = tabellvarde9 = 0;
            tabellvarde10 = tabellvarde11 = tabellvarde12 = tabellvarde13 = tabellvarde14 = tabellvarde15 = tabellvarde16 = tabellvarde17 = 0;

            if (rad1 != rad2 || rad1 != rad3) { indik1 = false; }
            else if (rad7 == 0 || rad8 == 0) { indik1 = false; }
            else
            {
                for (int i = 0; i < rad1; i++)
                {
                    bool tabellkont1 = (dataGridView1.Rows[i].Cells[1].Value != null) ? decimal.TryParse(dataGridView1.Rows[i].Cells[1].Value.ToString(), out tabellvarde1) : false;
                    bool tabellkont2 = (dataGridView1.Rows[i].Cells[2].Value != null) ? decimal.TryParse(dataGridView1.Rows[i].Cells[2].Value.ToString(), out tabellvarde2) : false;
                    bool tabellkont3 = (dataGridView1.Rows[i].Cells[3].Value != null) ? decimal.TryParse(dataGridView1.Rows[i].Cells[3].Value.ToString(), out tabellvarde3) : false;
                    bool tabellkont4 = (dataGridView1.Rows[i].Cells[4].Value != null) ? decimal.TryParse(dataGridView1.Rows[i].Cells[4].Value.ToString(), out tabellvarde4) : false;
                    bool tabellkont5 = (dataGridView1.Rows[i].Cells[5].Value != null) ? decimal.TryParse(dataGridView1.Rows[i].Cells[5].Value.ToString(), out tabellvarde5) : false;
                    bool tabellkont6 = (dataGridView2.Rows[i].Cells[0].Value != null) ? decimal.TryParse(dataGridView2.Rows[i].Cells[0].Value.ToString(), out tabellvarde6) : false;
                    bool tabellkont7 = (dataGridView2.Rows[i].Cells[1].Value != null) ? decimal.TryParse(dataGridView2.Rows[i].Cells[1].Value.ToString(), out tabellvarde7) : false;
                    bool tabellkont8 = (dataGridView2.Rows[i].Cells[2].Value != null) ? decimal.TryParse(dataGridView2.Rows[i].Cells[2].Value.ToString(), out tabellvarde8) : false;
                    bool tabellkont9 = (dataGridView2.Rows[i].Cells[3].Value != null) ? decimal.TryParse(dataGridView2.Rows[i].Cells[3].Value.ToString(), out tabellvarde9) : false;
                    bool tabellkont10 = (dataGridView3.Rows[i].Cells[0].Value != null) ? decimal.TryParse(dataGridView3.Rows[i].Cells[0].Value.ToString(), out tabellvarde10) : false;
                    bool tabellkont11 = (dataGridView3.Rows[i].Cells[1].Value != null) ? decimal.TryParse(dataGridView3.Rows[i].Cells[1].Value.ToString(), out tabellvarde11) : false;
                    bool tabellkont12 = (dataGridView3.Rows[i].Cells[2].Value != null) ? decimal.TryParse(dataGridView3.Rows[i].Cells[2].Value.ToString(), out tabellvarde12) : false;
                    bool tabellkont13 = (dataGridView3.Rows[i].Cells[3].Value != null) ? decimal.TryParse(dataGridView3.Rows[i].Cells[3].Value.ToString(), out tabellvarde13) : false;
                    bool tabellkont14 = (dataGridView4.Rows[i].Cells[0].Value != null) ? decimal.TryParse(dataGridView4.Rows[i].Cells[0].Value.ToString(), out tabellvarde14) : false;
                    bool tabellkont15 = (dataGridView4.Rows[i].Cells[1].Value != null) ? decimal.TryParse(dataGridView4.Rows[i].Cells[1].Value.ToString(), out tabellvarde15) : false;
                    bool tabellkont16 = (dataGridView4.Rows[i].Cells[2].Value != null) ? decimal.TryParse(dataGridView4.Rows[i].Cells[2].Value.ToString(), out tabellvarde16) : false;
                    bool tabellkont17 = (dataGridView4.Rows[i].Cells[3].Value != null) ? decimal.TryParse(dataGridView4.Rows[i].Cells[3].Value.ToString(), out tabellvarde17) : false;

                    // Kontrollerar Typ, höjd + skruv och elementval m.a.p. L1 och D1
                    if (!tabellkont1 || !tabellkont4 || !tabellkont6 || !tabellkont7 || !tabellkont10 || !tabellkont11 || !tabellkont14 || !tabellkont15) { indik1 = false; }
                    else
                    {
                        if (tabellvarde4 <= 0) { indik1 = false; }
                        if ((tabellvarde6 < 1 || tabellvarde6 > rad8) || (tabellvarde7 < 1 || tabellvarde7 > rad8)) { indik1 = false; }
                        if ((tabellvarde10 < 1 || tabellvarde10 > rad7) || (tabellvarde11 < 1 || tabellvarde11 > rad7)) { indik1 = false; }
                        if ((tabellvarde14 < 1) || (tabellvarde15 < 1)) { indik1 = false; }

                        // Kontrollerar bredd på sektion. Sista rad i tabell måste ha bredd.
                        if (!tabellkont5)
                        {
                            if (i == rad1 - 1) { indik1 = false; }
                            else if (dataGridView1.Rows[i].Cells[5].Value != null) { indik1 = false; }
                        }
                        else
                        {
                            if (tabellvarde5 <= 0) { indik1 = false; }
                        }

                        // Kontrollerar val typ av sektion
                        if (tabellvarde1 < 1 || tabellvarde1 > 3) { indik1 = false; }

                        if (tabellvarde1 == 1)
                        {
                            if (!tabellkont8 || !tabellkont12 || !tabellkont16)
                            {
                                //if (dataGridView2.Rows[i].Cells[2].Value != null || dataGridView3.Rows[i].Cells[2].Value != null || dataGridView4.Rows[i].Cells[2].Value != null) { indik1 = false;}
                            }
                            else
                            {
                                if ((tabellvarde8 < 1) || (tabellvarde8 > rad8)) { indik1 = false; }
                                if ((tabellvarde12 < 1) || (tabellvarde12 > rad7)) { indik1 = false; }
                                if (tabellvarde16 < 1) { indik1 = false; }
                            }
                            if (!tabellkont9 || !tabellkont13 || !tabellkont17)
                            {
                                if (dataGridView2.Rows[i].Cells[3].Value != null || dataGridView3.Rows[i].Cells[3].Value != null || dataGridView4.Rows[i].Cells[3].Value != null) { indik1 = false; }
                            }
                            else
                            {
                                if ((tabellvarde9 < 1) || (tabellvarde9 > rad8)) { indik1 = false; }
                                if ((tabellvarde13 < 1) || (tabellvarde13 > rad7)) { indik1 = false; }
                                if (tabellvarde17 < 1) { indik1 = false; }
                            }
                        }
                        else if (tabellvarde1 == 2)
                        {
                            if (!tabellkont9 || !tabellkont13 || !tabellkont17) { indik1 = false; }
                            else
                            {
                                if ((tabellvarde9 < 1) || (tabellvarde9 > rad8)) { indik1 = false; }
                                if ((tabellvarde13 < 1) || (tabellvarde13 > rad7)) { indik1 = false; }
                                if (tabellvarde17 < 1) { indik1 = false; }
                            }
                        }
                    }
                }
            }

            bool kont1 = decimal.TryParse(textBox1.Text, out decimal input1);
            bool kont2 = decimal.TryParse(textBox2.Text, out decimal input2);
            bool kont3 = decimal.TryParse(textBox3.Text, out decimal input3);
            bool kont4 = decimal.TryParse(textBox4.Text, out decimal input4);
            bool kont5 = decimal.TryParse(textBox5.Text, out decimal input5);
            bool kont6 = decimal.TryParse(textBox6.Text, out decimal input6);
            bool kont7 = decimal.TryParse(textBox7.Text, out decimal input7);
            bool kont8 = int.TryParse(textBox8.Text, out int input8);
            bool kont9 = decimal.TryParse(textBox9.Text, out decimal input9);

            if (!kont1 || !kont2 || !kont3 || !kont4 || !kont5 || !kont6 || !kont7 || !kont8 || !kont9) { indik1 = false; }
            else
            {
                if (input8 < 1) { textBox8.Text = "1"; }
            }
            return indik1;
        }

        // Beräknar start- och slutkoordinater för stolpens ramstänger
        // Indata
        //      x1:         x-koordinat, start på tyndpunktsline, typ = tal
        //      z1:         z-koordinat, start på tyndpunktsline, typ = tal
        //      x2:         x-koordinat, slut på tyndpunktsline, typ = tal
        //      z2:         z-koordinat, slut på tyndpunktsline, typ = tal
        //      z3:         z-koordinat, start på aktuell ramstång, typ = tal
        //      z4:         z-koordinat, slut på aktuell ramstång, typ = tal
        //      delta1:     Avstånd från tp till insida fläns, ramstång start tyndpunktsline, typ = tal
        //      delta2:     Avstånd från tp till insida fläns, ramstång slut tyndpunktsline, typ = tal
        //      delta3:     Avstånd från tp till insida fläns, aktuell ramstång, typ = tal
        // Utdata:
        //      noder:      Vektor = [x1,z1,x2,z2], x1,z1 är startnod och x2,z2 slutnod för stångens tyndpunktslinje
        // Noter:
        // L3 användes förut felaktit för att beräkna xt,zt
        private double[] Ramstangslinje(double x1, double z1, double x2, double z2, double z3, double z4, double delta1, double delta2, double delta3)
        {
            double[] noder = new double[4];
            double L1 = Math.Sqrt((x1 - x2) * (x1 - x2) + (z1 - z2) * (z1 - z2));
            double beta = Math.Asin((x1 - x2) / L1);
            double x11 = x1;
            double z11 = z1;
            double x21 = x2;
            double z21 = z2;

            //if (Math.Abs(delta1 - delta2) >= 0.001)
            //{
            if (delta1 == delta2)
            {
                x11 = x1 + (z2 - z1) * (delta2 - delta3) / L1;
                z11 = z1 + (x1 - x2) * (delta2 - delta3) / L1;
                x21 = x2 + (z2 - z1) * (delta2 - delta3) / L1;
                z21 = z2 + (x1 - x2) * (delta2 - delta3) / L1;
            }
            else if (delta1 > delta2)
            {
                double L2 = delta1 - delta2;
                double L3 = Math.Sqrt(L1 * L1 - L2 * L2);
                double alph = Math.Asin(L2 / L1);
                double xt = x2 + L3 * Math.Sin(alph + beta);
                double zt = z2 + L3 * -Math.Cos(alph + beta);
                x11 = xt + (xt - x1) * (delta2 - delta3) / L2;
                z11 = zt + (zt - z1) * (delta2 - delta3) / L2;
                x21 = x2 + (xt - x1) * (delta2 - delta3) / L2;
                z21 = z2 + (zt - z1) * (delta2 - delta3) / L2;
            }
            else
            {
                double L2 = delta2 - delta1;
                double alph = Math.PI * 0.5 - Math.Asin(L2 / L1);
                double xt = x2 + L2 * Math.Sin(alph + beta);
                double zt = z2 + L2 * -Math.Cos(alph + beta);
                x11 = x1 + (xt - x2) * (delta1 - delta3) / L2;
                z11 = z1 + (zt - z2) * (delta1 - delta3) / L2;
                x21 = xt + (xt - x2) * (delta1 - delta3) / L2;
                z21 = zt + (zt - z2) * (delta1 - delta3) / L2;
            }
            //}
            noder[0] = x11 + ((z3 - z11) / (z21 - z11)) * (x21 - x11);
            noder[1] = z3;
            noder[2] = x11 + ((z4 - z11) / (z21 - z11)) * (x21 - x11);
            noder[3] = z4;
            return noder;
        }

        // Skapar listor med element
        // Arbetet sker i tre steg:
        // 1 Knutkoordinater    I detta steg beräknas koordinaterna för alla knutpunkter där kraftlijer skär varandra.
        //                      Start- och slutkoordnater för ramstängernas tyndpunktslinjer väljs så att 1) insidan 
        //                      av ramstängerna med samma lutning bildar en rak linje och 2) avståndet mellan 
        //                      slutkoordinaterna motsvarar angiven bredd på sektionen (undantag: understa sektionen
        //                      där basbredden sätts lika mellan avståndet mellan startkoordinaterna). 
        //                      I detta steg sparas även sektions och materialdata för alla stänger.
        //
        // 2 Ändföskjutningar   Tar fram förskjutnings-vektorer (FV) för start- och slutpunkterna hos alla diagonaler 
        //                      och ramstänger. FV används för att 1) undvika krockar mellan element och 2) positionera 
        //                      stänger rätt mot insida/utsida av ramstång eller plåt.
        //              
        // 3 Elementlista       Skapa lista med alla element som skall skapas i Tekla genom att kombinera resultat från
        //                      steg 1 och 2.
        //
        // Utdata:      Listor med elementdata (starnod slutnod, material, dimension, rotation m.m)
        private void IndatatTillTekla()
        {
            tilltek_GUID.Clear();
            tilltek_idnr.Clear();
            tilltek_pos.Clear();
            tilltek_startnod.Clear();
            tilltek_slutnod.Clear();
            tilltek_startPl.Clear();
            tilltek_slutPl.Clear();
            tilltek_typ.Clear();
            tilltek_dim.Clear();
            tilltek_kvalitet.Clear();
            tilltek_rotera.Clear();
            tilltek_x1.Clear();
            tilltek_y1.Clear();
            tilltek_z1.Clear();
            tilltek_x2.Clear();
            tilltek_y2.Clear();
            tilltek_z2.Clear();
            tilltek_blt.Clear();
            tilltek_bltN.Clear();
            tilltek_blte1.Clear();
            tilltek_bltp1.Clear();
            tilltek_bltd0.Clear();

            tilltekPl_idnr.Clear();
            tilltekPl_GUID.Clear();
            tilltekPl_typ.Clear();
            tilltekPl_t.Clear();
            tilltekPl_ram.Clear();
            tilltekPl_ramc.Clear();
            tilltekPl_ramb.Clear();
            tilltekPl_ramt.Clear();
            tilltekPl_diag.Clear();
            tilltekPl_blt.Clear();
            tilltekPl_bltN.Clear();
            tilltekPl_blte1.Clear();
            tilltekPl_bltp1.Clear();
            tilltekPl_bltd0.Clear();
            tilltekPl_cntr.Clear();

            List<int> sekdata_D1blt = new List<int>();
            List<int> sekdata_H1blt = new List<int>();
            List<int> sekdata_H2blt = new List<int>();
            List<int> sekdata_D1bltN = new List<int>();
            List<int> sekdata_H1bltN = new List<int>();
            List<int> sekdata_H2bltN = new List<int>();
            List<double> sekdata_D1blte1 = new List<double>();
            List<double> sekdata_H1blte1 = new List<double>();
            List<double> sekdata_H2blte1 = new List<double>();
            List<double> sekdata_D1bltp1 = new List<double>();
            List<double> sekdata_H1bltp1 = new List<double>();
            List<double> sekdata_H2bltp1 = new List<double>();
            List<double> sekdata_D1bltd0 = new List<double>();
            List<double> sekdata_H1bltd0 = new List<double>();
            List<double> sekdata_H2bltd0 = new List<double>();

            List<double> sekdata_1x = new List<double>();       // x-koord, start på ramstångens mittlinje 
            List<double> sekdata_1z = new List<double>();       // z-koord, start på ramstångens mittlinje 
            List<double> sekdata_2x = new List<double>();       // x-koord, slut på ramstångens mittlinje 
            List<double> sekdata_2z = new List<double>();       // z-koord, slut på ramstångens mittlinje
            List<double> sekdata_3x = new List<double>();       // x-koord, knutpunkt, start på ramstången
            List<double> sekdata_3y = new List<double>();       // y-koord, knutpunkt, start på ramstången
            List<double> sekdata_3z = new List<double>();       // z-koord, knutpunkt, start på ramstången 
            List<double> sekdata_4x = new List<double>();       // x-koord, knutpunkt, slut på ramstången
            List<double> sekdata_4y = new List<double>();       // y-koord, knutpunkt, slut på ramstången
            List<double> sekdata_4z = new List<double>();       // z-koord, knutpunkt, slut på ramstången
            List<double> sekdata_5y = new List<double>();       // x-koord, knutpunkt, mitt övre horisontal
            List<double> sekdata_5z = new List<double>();       // z-koord, knutpunkt, mitt övre horisontal
            List<double> sekdata_6y = new List<double>();       // x-koord, knutpunkt, mitt undre horisontal
            List<double> sekdata_6z = new List<double>();       // z-koord, knutpunkt, mitt undre horisontal
            List<double> sekdata_7x = new List<double>();       // x-koord, knutpunkt, ände undre horisontal
            List<double> sekdata_7z = new List<double>();       // z-koord, knutpunkt, ände undre horisontal
            List<double> sekdata_8y = new List<double>();       // x-koord, knutpunkt
            List<double> sekdata_8z = new List<double>();       // z-koord, knutpunkt

            List<double> sekd_d1y = new List<double>();         // x-koord, förskjutninsvektor för att placera diagonal / horisontal på utsidan av ramstångens fläns
            List<double> sekd_d1z = new List<double>();         // z-koord, förskjutninsvektor för att placera diagonal / horisontal på utsidan av ramstångens fläns
            List<double> sekd_d2y = new List<double>();         // x-koord, förskjutninsvektor för att placera diagonal / horisontal på insidan av ramstångens fläns
            List<double> sekd_d2z = new List<double>();         // z-koord, förskjutninsvektor för att placera diagonal / horisontal på insidan av ramstångens fläns

            List<double> sekd_H2_dL1 = new List<double>();      // Förlängning [m] av övre horisontal vid nod 4
            List<double> sekd_H2_dL2 = new List<double>();      // Förlängning [m] av övre horisontal vid nod 5
            List<double> sekd_H1_dL1 = new List<double>();      // Förlängning [m] av undre horisontal vid nod 7
            List<double> sekd_H1_dL2 = new List<double>();      // Förlängning [m] av undre horisontal vid nod 6
            List<double> sekd_D1_dL1 = new List<double>();      // Förlängning [m] av diagonal vid undre nod (nod 3 eller 8 beroende på på param. sekdata_D1typ)
            List<double> sekd_D1_dL2 = new List<double>();      // Förlängning [m] av diagonal vid mittnod 
            List<double> sekd_D1_dL3 = new List<double>();      // Förlängning [m] av diagonal vid mittnod 
            List<double> sekd_D1_dL4 = new List<double>();      // Förlängning [m] av diagonal vid övre nod (nod 3 eller 5 beroende på param. sekdata_D1typ)

            List<string> sekdata_L1kdim = new List<string>();
            List<string> sekdata_L1kvalitet = new List<string>();
            List<int> sekdata_D1typ = new List<int>();
            List<string> sekdata_D1kdim = new List<string>();
            List<string> sekdata_D1kvalitet = new List<string>();
            List<int> sekdata_D1dorient = new List<int>();
            List<double> sekdata_D1rot = new List<double>();
            List<string> sekdata_H1kdim = new List<string>();
            List<string> sekdata_H1kvalitet = new List<string>();
            List<double> sekdata_Hrot = new List<double>();
            List<int> sekdata_H1dorient = new List<int>();
            List<string> sekdata_H2kdim = new List<string>();
            List<string> sekdata_H2kvalitet = new List<string>();
            List<int> sekdata_H2dorient = new List<int>();
            List<decimal> sekdata_L1t = new List<decimal>();
            List<decimal> sekdata_L1b = new List<decimal>();
            List<decimal> sekdata_L1cz = new List<decimal>();

            int antrader = dataGridView2.Rows.Count;    // Antal sektioner
            decimal hsek1 = 0;                          // Z-koord, underkant sektion
            decimal hsek2 = 0;                          // Höjd (i z-riktning), sektion
            decimal hsek3 = 0;                          // Z-koord, start på ramstångslinje
            decimal hsek4 = 0;                          // Z-koord, slut på ramstångslinje
            decimal bsek3 = 0;                          // Bredd, start på ramstångslinje
            decimal bsek4 = 0;                          // Bredd, slut på ramstångslinje
            decimal cz1 = 0;                            // Avstånd från kant till tp, aktuell stång  
            decimal cz3 = 0;                            // Avstånd från kant till tp, stång underst i ramstångslinje  
            decimal cz4 = 0;                            // Avstånd från kant till tp, stång överst i ramstångslinje  
            decimal t1 = 0;                             // Tjocklek fläns, aktuell stång 
            decimal t3 = 0;                             // Tjocklek fläns, stång underst i ramstångslinje  
            decimal t4 = 0;                             // Tjocklek fläns, stång överst i ramstångslinje  
            decimal b1 = 0;                             // Bredd fläns, aktuell stång
            double k2 = 0;                              // Param som anger var mellan ramstångens tp (k2=0) och centrumline (k2=1) som övre knutpunkten placeras
            double k3 = 0;                              // Param som anger var mellan ramstångens tp (k3=0) och centrumline (k3=1) som nedre knutpunkten placeras
            int k1 = 0;                                 // Param som anger om två diagonaler får dela skruv (0 = nej, 1 = ja)
            double[] koord = new double[4];


            // 1 Knutkoordinater och elementdata    
            for (int i = 0; i < antrader; i++)
            {
                // 1.2 Bultar
                int typD1 = 0;
                int typH1 = 0;
                int typH2 = 0;
                double eimin = double.Parse(textBox3.Text);
                double pimin = double.Parse(textBox5.Text);

                typD1 = int.Parse(dataGridView3.Rows[i].Cells[1].Value.ToString()) - 1;
                sekdata_D1bltN.Add(int.Parse(dataGridView4.Rows[i].Cells[1].Value.ToString()));

                if ((dataGridView3.Rows[i].Cells[2].Value != null))
                {
                    typH1 = int.Parse(dataGridView3.Rows[i].Cells[2].Value.ToString()) - 1;
                    sekdata_H1bltN.Add(int.Parse(dataGridView4.Rows[i].Cells[2].Value.ToString()));
                }
                else
                { sekdata_H1bltN.Add(9); }
                if ((dataGridView3.Rows[i].Cells[3].Value != null))
                {
                    typH2 = int.Parse(dataGridView3.Rows[i].Cells[3].Value.ToString()) - 1;
                    sekdata_H2bltN.Add(int.Parse(dataGridView4.Rows[i].Cells[3].Value.ToString()));
                }
                else
                { sekdata_H2bltN.Add(9); }

                double doD1 = double.Parse(dataGridView7.Rows[typD1].Cells[5].Value.ToString());
                double doH1 = double.Parse(dataGridView7.Rows[typH1].Cells[5].Value.ToString());
                double doH2 = double.Parse(dataGridView7.Rows[typH2].Cells[5].Value.ToString());
                sekdata_D1blt.Add(int.Parse(dataGridView7.Rows[typD1].Cells[2].Value.ToString()));
                sekdata_H1blt.Add(int.Parse(dataGridView7.Rows[typH1].Cells[2].Value.ToString()));
                sekdata_H2blt.Add(int.Parse(dataGridView7.Rows[typH2].Cells[2].Value.ToString()));
                sekdata_D1bltd0.Add(doD1);
                sekdata_H1bltd0.Add(doH1);
                sekdata_H2bltd0.Add(doH2);
                sekdata_D1blte1.Add(eimin * doD1);
                sekdata_H1blte1.Add(eimin * doH1);
                sekdata_H2blte1.Add(eimin * doH2);
                sekdata_D1bltp1.Add(pimin * doD1);
                sekdata_H1bltp1.Add(pimin * doH1);
                sekdata_H2bltp1.Add(pimin * doD1);

                // 1.3 Ramstänger
                int radel = int.Parse(dataGridView2.Rows[i].Cells[0].Value.ToString()) - 1;
                b1 = decimal.Parse(dataGridView8.Rows[radel].Cells[4].Value.ToString()) / 1000;
                t1 = decimal.Parse(dataGridView8.Rows[radel].Cells[5].Value.ToString()) / 1000;
                cz1 = decimal.Parse(dataGridView8.Rows[radel].Cells[6].Value.ToString()) / 1000;
                k2 = double.Parse(dataGridView5.Rows[i].Cells[1].Value.ToString());
                if (i > 0) { k3 = double.Parse(dataGridView5.Rows[i - 1].Cells[1].Value.ToString()); ; }
                else { k3 = 0; }
                sekdata_L1kdim.Add(dataGridView8.Rows[radel].Cells[2].Value.ToString());
                sekdata_L1kvalitet.Add(dataGridView8.Rows[radel].Cells[3].Value.ToString());
                sekdata_L1t.Add(t1);
                sekdata_L1cz.Add(cz1);

                // Tar fram start- och slutkkordinater för ramstänger
                hsek2 = decimal.Parse(dataGridView1.Rows[i].Cells[4].Value.ToString());
                if ((i == 0) || ((dataGridView1.Rows[i - 1].Cells[5].Value != null)))
                {
                    if (i == 0)
                    {
                        bsek3 = decimal.Parse(textBox1.Text);
                        t3 = t1;
                        cz3 = cz1;
                    }
                    else
                    {
                        bsek3 = decimal.Parse(dataGridView1.Rows[i - 1].Cells[5].Value.ToString());
                        t3 = sekdata_L1t[i - 1];
                        cz3 = sekdata_L1cz[i - 1];
                    }
                    hsek3 = hsek1;
                    hsek4 = hsek1;
                    for (int j = i; j < antrader; j++)
                    {
                        hsek4 = hsek4 + decimal.Parse(dataGridView1.Rows[j].Cells[4].Value.ToString());
                        if (dataGridView1.Rows[j].Cells[5].Value != null)
                        {
                            bsek4 = decimal.Parse(dataGridView1.Rows[j].Cells[5].Value.ToString());
                            radel = int.Parse(dataGridView2.Rows[j].Cells[0].Value.ToString()) - 1;
                            t4 = decimal.Parse(dataGridView8.Rows[radel].Cells[5].Value.ToString()) / 1000;
                            cz4 = decimal.Parse(dataGridView8.Rows[radel].Cells[6].Value.ToString()) / 1000;
                            break;
                        }
                    }
                }
                koord[0] = 0.5 * Convert.ToDouble(bsek3 + (bsek4 - bsek3) * (hsek1 - hsek3) / (hsek4 - hsek3));
                koord[1] = Convert.ToDouble(hsek1);
                koord[2] = 0.5 * Convert.ToDouble(bsek3 + (bsek4 - bsek3) * (hsek1 + hsek2 - hsek3) / (hsek4 - hsek3));
                koord[3] = Convert.ToDouble(hsek1 + hsek2);
                double[] ramvek = { koord[2] - koord[0], koord[2] - koord[0], koord[3] - koord[1] };
                ramvek = Normvek(ramvek, 3);

                // Beräknar koordinater för knutpunkter 1-2 (stångens mittlinje)
                double dx0 = Convert.ToDouble(cz1 - t1) * -1 * ramvek[2];
                double dz0 = Convert.ToDouble(cz1 - t1) * 2 * ramvek[0];
                double dx1 = Convert.ToDouble((b1 / 2) - cz1) * -1 * ramvek[2];
                double dz1 = Convert.ToDouble((b1 / 2) - cz1) * 2 * ramvek[0];
                double dx3 = 0;
                double dz3 = 0;
                double dx4 = 0;
                double dz4 = 0;
                if (i != 0)
                {
                    if (dataGridView1.Rows[i - 1].Cells[5].Value != null && i + 1 < antrader)
                    {
                        dx3 = 0.005 * ramvek[0];
                        dz3 = 0.005 * ramvek[2];
                    }
                    else if (dataGridView1.Rows[i - 1].Cells[6].Value != null)
                    {
                        dx3 = (double.Parse(dataGridView1.Rows[i - 1].Cells[6].Value.ToString()) + 0.01) / 1 * ramvek[0];
                        dz3 = (double.Parse(dataGridView1.Rows[i - 1].Cells[6].Value.ToString()) + 0.01) / 1 * ramvek[2];
                    }
                    //else if (dataGridView2.Rows[i].Cells[0].Value.ToString() != dataGridView2.Rows[i - 1].Cells[0].Value.ToString())
                    //{
                    //    dx3 = 0.21 * ramvek[0];
                    //    dz3 = 0.21 * ramvek[2];
                    //}
                }
                if (i < antrader - 1)
                {
                    if (dataGridView1.Rows[i].Cells[5].Value != null && i + 1 < antrader)
                    {
                        dx4 = -0.005 * ramvek[0];
                        dz4 = -0.005 * ramvek[2];
                    }
                    else if (dataGridView1.Rows[i].Cells[6].Value != null)
                    {
                        dx4 = double.Parse(dataGridView1.Rows[i].Cells[6].Value.ToString()) / 1 * ramvek[0];
                        dz4 = double.Parse(dataGridView1.Rows[i].Cells[6].Value.ToString()) / 1 * ramvek[2];
                    }
                    //else if (dataGridView2.Rows[i].Cells[0].Value.ToString() != dataGridView2.Rows[i+1].Cells[0].Value.ToString())
                    //{
                    //    dx4 = 0.200 * ramvek[0];
                    //    dz4 = 0.200 * ramvek[2];
                    //}
                }
                sekdata_1x.Add(koord[0] + dx0 + dx1 + dx3);
                sekdata_1z.Add(koord[1] + dz0 + dz1 + dz3);
                sekdata_2x.Add(koord[2] + dx0 + dx1 + dx4);
                sekdata_2z.Add(koord[3] + dz0 + dz1 + dz4);

                // Beräknar koordinater för knutpunkter 3-7
                sekdata_3x.Add(koord[0] + dx0 + k3 * dx1);
                sekdata_3y.Add(koord[0] + dx0);
                sekdata_3z.Add(koord[1] + dz0 + k3 * dz1);
                sekdata_4x.Add(koord[2] + dx0 + k2 * dx1);
                sekdata_4y.Add(koord[2] + dx0);
                sekdata_4z.Add(koord[3] + dz0 + k2 * dz1);
                sekdata_5y.Add(koord[2] + dx0);
                sekdata_5z.Add(koord[3] + dz0);
                sekdata_6y.Add(koord[0] + dx0 + (koord[2] - koord[0]) * (koord[0] + dx0) / (koord[0] + koord[2] + 2 * dx0));
                sekdata_6z.Add(koord[1] + dz0 + (koord[3] - koord[1]) * (koord[0] + dx0) / (koord[0] + koord[2] + 2 * dx0));
                sekdata_7x.Add(koord[0] + dx0 + (koord[2] - koord[0]) * (koord[0] + dx0) / (koord[0] + koord[2] + 2 * dx0));
                sekdata_7z.Add(koord[1] + dz0 + (koord[3] - koord[1]) * (koord[0] + dx0) / (koord[0] + koord[2] + 2 * dx0));
                sekdata_8y.Add(koord[0] + dx0);
                sekdata_8z.Add(koord[1] + dz0);
                hsek1 = hsek1 + hsek2;

                // 1.4 Diagonaler
                int D1typ = int.Parse(dataGridView1.Rows[i].Cells[1].Value.ToString());
                radel = int.Parse(dataGridView2.Rows[i].Cells[1].Value.ToString()) - 1;
                sekdata_D1kdim.Add(dataGridView8.Rows[radel].Cells[2].Value.ToString());
                sekdata_D1kvalitet.Add(dataGridView8.Rows[radel].Cells[3].Value.ToString());
                sekdata_D1typ.Add(D1typ);
                if ((dataGridView2.Rows[i].Cells[2].Value != null))
                {
                    // Ena diagonalen på utsida ramstång, andra på insida ramstång
                    sekdata_D1dorient.Add(1);
                }
                else
                {
                    // Ena diagonalen på utsida ramstång, andra på insida ramstång
                    sekdata_D1dorient.Add(2);
                }

                // 1.5 Horisontaler
                if ((dataGridView2.Rows[i].Cells[2].Value != null))
                {
                    int radelh1 = int.Parse(dataGridView2.Rows[i].Cells[2].Value.ToString()) - 1;
                    sekdata_H1kdim.Add(dataGridView8.Rows[radelh1].Cells[2].Value.ToString());
                    sekdata_H1kvalitet.Add(dataGridView8.Rows[radelh1].Cells[3].Value.ToString());
                    sekdata_H1dorient.Add(2);
                }
                else
                {
                    sekdata_H1kdim.Add("");
                    sekdata_H1kvalitet.Add("");
                    sekdata_H1dorient.Add(0);
                }
                if ((dataGridView2.Rows[i].Cells[3].Value != null))
                {
                    int radelh1 = int.Parse(dataGridView2.Rows[i].Cells[3].Value.ToString()) - 1;
                    sekdata_H2kdim.Add(dataGridView8.Rows[radelh1].Cells[2].Value.ToString());
                    sekdata_H2kvalitet.Add(dataGridView8.Rows[radelh1].Cells[3].Value.ToString());
                    sekdata_H1dorient.Add(2);
                }
                else
                {
                    sekdata_H2kdim.Add("");
                    sekdata_H2kvalitet.Add("");
                    sekdata_H1dorient.Add(2);
                }

                // 1.6 Beräknar vilken roation av diagonaler / horisontaler som krävs för att dessa elements flänsar 
                //     skall löpa parallellt med ramstångens fläns.
                double betaram = Math.Atan((koord[0] - koord[2]) / (koord[3] - koord[1]));
                double betadiag = 0;
                double betadiag2 = 0;
                double normx = 0;
                double normy = 0;
                double normz = 0;
                double dxy = 0;
                if (D1typ == 1)
                {
                    dxy = Math.Sqrt(Math.Pow(sekdata_3x[i] + sekdata_4x[i], 2) + Math.Pow(sekdata_4x[i] - sekdata_3x[i], 2));
                    normx = (sekdata_4z[i] - sekdata_3z[i]) * (sekdata_3x[i] + sekdata_4x[i]) / dxy;
                    normy = (sekdata_4z[i] - sekdata_3z[i]) * (sekdata_4x[i] - sekdata_3x[i]) / dxy;
                    normz = dxy;
                    //betadiag = Math.Atan((sekdata_3x[i] + sekdata_4x[i]) / (sekdata_4z[i] - sekdata_3z[i]));
                    //betadiag2 = Math.Atan((sekdata_4x[i] - sekdata_3x[i]) / Math.Sqrt(Math.Pow(sekdata_3x[i] + sekdata_4x[i], 2) + Math.Pow(sekdata_4z[i] - sekdata_3z[i], 2)));
                }
                else if (D1typ == 2)
                {
                    dxy = Math.Sqrt(Math.Pow(sekdata_3x[i], 2) + Math.Pow(sekdata_4x[i] - sekdata_3x[i], 2));
                    normx = (sekdata_4z[i] - sekdata_3z[i]) * sekdata_3x[i] / dxy;
                    normy = (sekdata_4z[i] - sekdata_3z[i]) * (sekdata_4x[i] - sekdata_3x[i]) / dxy;
                    normz = dxy;
                    //betadiag = Math.Atan((sekdata_3x[i]) / (sekdata_4z[i] - sekdata_3z[i]));
                    //betadiag2 = Math.Atan((sekdata_4x[i] - sekdata_3x[i]) / Math.Sqrt(Math.Pow(sekdata_3x[i], 2) + Math.Pow(sekdata_4z[i] - sekdata_3z[i], 2)));
                }
                else if (D1typ == 3)
                {
                    dxy = Math.Sqrt(Math.Pow(sekdata_4x[i], 2) + Math.Pow(sekdata_4x[i] - sekdata_3x[i], 2));
                    normx = (sekdata_4z[i] - sekdata_3z[i]) * sekdata_4x[i] / dxy;
                    normy = (sekdata_4z[i] - sekdata_3z[i]) * (sekdata_4x[i] - sekdata_3x[i]) / dxy;
                    normz = dxy;
                    //betadiag = Math.Atan((sekdata_4x[i]) / (sekdata_4z[i] - sekdata_3z[i]));
                    //betadiag2 = Math.Atan((sekdata_4x[i] - sekdata_3x[i]) / Math.Sqrt(Math.Pow(sekdata_4x[i], 2) + Math.Pow(sekdata_4z[i] - sekdata_3z[i], 2)));
                }
                betadiag = Math.Atan(1 * normz / normx);
                betadiag2 = Math.Atan(normy / Math.Sqrt(Math.Pow(normx, 2) + Math.Pow(normz, 2)));
                sekdata_D1rot.Add((-1 * Math.Atan(Math.Tan(betaram) * Math.Sin(betadiag)) + betadiag2) * 180 / Math.PI);
                sekdata_Hrot.Add(-1 * betaram * 180 / Math.PI);

                // 1.6 Beräknar förskjutninsvektor som krävs för att positionera diagonaler / horisontaler på utsidan 
                //     (sekd_d1) respekve på insidan (sekd_d2) av ramstångens fläns. 
                sekd_d1y.Add(Convert.ToDouble(cz1) * ramvek[2]);
                sekd_d1z.Add(Convert.ToDouble(cz1) * -1 * ramvek[0]);
                sekd_d2y.Add(Convert.ToDouble(cz1 - t1) * ramvek[2]);
                sekd_d2z.Add(Convert.ToDouble(cz1 - t1) * -1 * ramvek[0]);
            }

            // 2 Ändföskjutningar
            // 2.1 Sektionsdata för knutpunkter (diagonaler och horisontaler)
            for (int i = 0; i < antrader; i++)
            {
                //double[] vek = { sekdata_D1x2[i]+ sekdata_D1x1[i], sekdata_D1y2[i] + sekdata_D1y1[i], sekdata_D1z2[i] + sekdata_D1z1[i] };
                //vek = Normvek(vek);
                double tmp_H1dL1 = 0;
                double tmp_H1dL2 = 0;
                double tmp_H2dL1 = 0;
                double tmp_H2dL2 = 0;
                double tmp_D1dL1 = 0;
                double tmp_D1dL2 = 0;
                double tmp_D1dL3 = 0;
                double tmp_D1dL4 = 0;
                if ((dataGridView1.Rows[i].Cells[11].Value != null)) { double.TryParse(dataGridView1.Rows[i].Cells[11].Value.ToString(), out tmp_H1dL1); }
                if ((dataGridView1.Rows[i].Cells[12].Value != null)) { double.TryParse(dataGridView1.Rows[i].Cells[12].Value.ToString(), out tmp_H1dL2); }
                if ((dataGridView1.Rows[i].Cells[13].Value != null)) { double.TryParse(dataGridView1.Rows[i].Cells[13].Value.ToString(), out tmp_H2dL1); }
                if ((dataGridView1.Rows[i].Cells[14].Value != null)) { double.TryParse(dataGridView1.Rows[i].Cells[14].Value.ToString(), out tmp_H2dL2); }
                if ((dataGridView1.Rows[i].Cells[7].Value != null)) { double.TryParse(dataGridView1.Rows[i].Cells[7].Value.ToString(), out tmp_D1dL1); }
                if ((dataGridView1.Rows[i].Cells[8].Value != null)) { double.TryParse(dataGridView1.Rows[i].Cells[8].Value.ToString(), out tmp_D1dL2); }
                if ((dataGridView1.Rows[i].Cells[9].Value != null)) { double.TryParse(dataGridView1.Rows[i].Cells[9].Value.ToString(), out tmp_D1dL3); }
                if ((dataGridView1.Rows[i].Cells[10].Value != null)) { double.TryParse(dataGridView1.Rows[i].Cells[10].Value.ToString(), out tmp_D1dL4); }
                sekd_H1_dL1.Add(tmp_H1dL1 - sekdata_H1blte1[i] / 1000);      // Förlängning [m] av övre horisontal vid nod 4
                sekd_H1_dL2.Add(tmp_H1dL2 - sekdata_H1blte1[i] / 1000);      // Förlängning [m] av övre horisontal vid nod 5
                sekd_H2_dL1.Add(tmp_H2dL1 - sekdata_H2blte1[i] / 1000);      // Förlängning [m] av mitten-horisontal vid nod 7
                sekd_H2_dL2.Add(tmp_H2dL2 - sekdata_H2blte1[i] / 1000);      // Förlängning [m] av mitten-horisontal vid nod 6
                sekd_D1_dL1.Add(tmp_D1dL1 - sekdata_D1blte1[i] / 1000);      // Förlängning [m] av diagonal vid undre nod (nod 3 eller 8 beroende på på param. sekdata_D1typ)
                sekd_D1_dL2.Add(tmp_D1dL2 - sekdata_D1blte1[i] / 1000);      // Förlängning [m] av diagonal vid mittnod
                sekd_D1_dL3.Add(tmp_D1dL3 - sekdata_D1blte1[i] / 1000);      // Förlängning [m] av diagonal vid mittnod
                sekd_D1_dL4.Add(tmp_D1dL4 - sekdata_D1blte1[i] / 1000);      // Förlängning [m] av diagonal vid övre nod (nod 3 eller 5 beroende på param. sekdata_D1typ)
            }

            // 3 Elementlista
            // 3.1 Ramstänger  (stolpkropp)
            for (int i = 0; i < sekdata_L1kdim.Count; i++)
            {
                int[] iternr = { -45, 45, 135, 225 };
                int k = 0;
                foreach (int ii in iternr)
                {
                    tilltek_typ.Add("L1");
                    tilltek_dim.Add(sekdata_L1kdim[i]);
                    tilltek_kvalitet.Add(sekdata_L1kvalitet[i]);
                    tilltek_pos.Add(0);
                    tilltek_rotera.Add(135);
                    tilltek_x1.Add(sekdata_1x[i] * Math.Sign(Math.Cos((ii) * Math.PI / 180)));
                    tilltek_y1.Add(sekdata_1x[i] * Math.Sign(Math.Sin((ii) * Math.PI / 180)));
                    tilltek_x2.Add(sekdata_2x[i] * Math.Sign(Math.Cos((ii) * Math.PI / 180)));
                    tilltek_y2.Add(sekdata_2x[i] * Math.Sign(Math.Sin((ii) * Math.PI / 180)));
                    tilltek_z1.Add(sekdata_1z[i]);
                    tilltek_z2.Add(sekdata_2z[i]);
                    tilltek_blt.Add(0);
                    tilltek_bltN.Add(0);
                    tilltek_blte1.Add(0);
                    tilltek_bltp1.Add(0);
                    tilltek_bltd0.Add(0);
                    tilltek_idnr.Add(string.Format("N:{0}_L1:{1}", i, k));
                    tilltek_startnod.Add(string.Format("N:{0}_L1:{1}", i - 1, k));
                    tilltek_slutnod.Add(string.Format("N:{0}_L1:{1}", i + 1, k));
                    tilltek_startPl.Add("");
                    tilltek_slutPl.Add("");
                    k = k + 1;
                }
            }

            // 3.2 Horisontaler (stolpkropp)
            for (int i = 0; i < sekdata_H1kdim.Count; i++)
            {
                int[] iternr = { 0, 90, 180, 270 };
                int[] hjap1 = { 0, 1, 2, 3, 0 };
                if (sekdata_H1kdim[i] != string.Empty)
                {
                    int k = 0;
                    foreach (int ii in iternr)
                    {
                        tilltek_typ.Add("H1");
                        tilltek_pos.Add(2);
                        tilltek_dim.Add(sekdata_H1kdim[i]);
                        tilltek_kvalitet.Add(sekdata_H1kvalitet[i]);
                        tilltek_rotera.Add(-sekdata_Hrot[i]);
                        double a_x1 = sekdata_7x[i] * Math.Sign(Math.Cos((ii + 45) * Math.PI / 180)) + Math.Cos((ii) * Math.PI / 180) * sekd_d2y[i];
                        double a_y1 = sekdata_7x[i] * Math.Sign(Math.Sin((ii + 45) * Math.PI / 180)) + Math.Sin((ii) * Math.PI / 180) * sekd_d2y[i];
                        double a_z1 = sekdata_7z[i] + sekd_d2z[i];
                        double a_x2 = sekdata_7x[i] * Math.Sign(Math.Cos((ii - 45) * Math.PI / 180)) + Math.Cos((ii) * Math.PI / 180) * sekd_d2y[i];
                        double a_y2 = sekdata_7x[i] * Math.Sign(Math.Sin((ii - 45) * Math.PI / 180)) + Math.Sin((ii) * Math.PI / 180) * sekd_d2y[i];
                        double a_z2 = sekdata_7z[i] + sekd_d2z[i];
                        double[] vek = new double[3] { a_x2 - a_x1, a_y2 - a_y1, a_z2 - a_z1 };
                        vek = Normvek(vek);
                        tilltek_x1.Add(a_x1 + sekd_H1_dL1[i] * vek[0]);
                        tilltek_y1.Add(a_y1 + sekd_H1_dL1[i] * vek[1]);
                        tilltek_z1.Add(a_z1 + sekd_H1_dL1[i] * vek[2]);
                        tilltek_x2.Add(a_x2 - sekd_H1_dL1[i] * vek[0]);
                        tilltek_y2.Add(a_y2 - sekd_H1_dL1[i] * vek[1]);
                        tilltek_z2.Add(a_z2 - sekd_H1_dL1[i] * vek[2]);
                        tilltek_blt.Add(sekdata_H1blt[i]);
                        tilltek_bltN.Add(sekdata_H1bltN[i]);
                        tilltek_blte1.Add(sekdata_H1blte1[i]);
                        tilltek_bltp1.Add(sekdata_H1bltp1[i]);
                        tilltek_bltd0.Add(sekdata_H1bltd0[i]);
                        tilltek_idnr.Add(string.Format("N:{0}_H1:{1}", i, k));
                        tilltek_startnod.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k + 1]));
                        tilltek_slutnod.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k]));
                        tilltek_startPl.Add("");
                        tilltek_slutPl.Add("");
                        k = k + 1;
                    }
                }
            }
            for (int i = 0; i < sekdata_H2kdim.Count; i++)
            {
                int[] iternr = { 0, 90, 180, 270 };
                int[] hjap1 = { 0, 1, 2, 3, 0 };
                if (sekdata_H2kdim[i] != string.Empty)
                {
                    int k = 0;
                    foreach (int ii in iternr)
                    {
                        tilltek_typ.Add("H2");
                        tilltek_pos.Add(2);
                        tilltek_dim.Add(sekdata_H2kdim[i]);
                        tilltek_kvalitet.Add(sekdata_H2kvalitet[i]);
                        tilltek_rotera.Add(-sekdata_Hrot[i]);
                        double a_x1 = Math.Abs(sekdata_4x[i] * Math.Sin(ii * Math.PI / 180) + sekdata_4y[i] * Math.Cos(ii * Math.PI / 180)) * Math.Sign(Math.Cos((ii + 45) * Math.PI / 180)) + Math.Cos((ii) * Math.PI / 180) * sekd_d2y[i];
                        double a_y1 = Math.Abs(sekdata_4x[i] * Math.Cos(ii * Math.PI / 180) + sekdata_4y[i] * Math.Sin(ii * Math.PI / 180)) * Math.Sign(Math.Sin((ii + 45) * Math.PI / 180)) + Math.Sin((ii) * Math.PI / 180) * sekd_d2y[i];
                        double a_z1 = sekdata_4z[i] + sekd_d2z[i];
                        double a_x2 = Math.Abs(sekdata_4x[i] * Math.Sin(ii * Math.PI / 180) + sekdata_4y[i] * Math.Cos(ii * Math.PI / 180)) * Math.Sign(Math.Cos((ii - 45) * Math.PI / 180)) + Math.Cos((ii) * Math.PI / 180) * sekd_d2y[i];
                        double a_y2 = Math.Abs(sekdata_4x[i] * Math.Cos(ii * Math.PI / 180) + sekdata_4y[i] * Math.Sin(ii * Math.PI / 180)) * Math.Sign(Math.Sin((ii - 45) * Math.PI / 180)) + Math.Sin((ii) * Math.PI / 180) * sekd_d2y[i];
                        double a_z2 = sekdata_4z[i] + sekd_d2z[i];
                        double[] vek = new double[3] { a_x2 - a_x1, a_y2 - a_y1, a_z2 - a_z1 };
                        vek = Normvek(vek);
                        tilltek_x1.Add(a_x1 + sekd_H2_dL1[i] * vek[0]);
                        tilltek_y1.Add(a_y1 + sekd_H2_dL1[i] * vek[1]);
                        tilltek_z1.Add(a_z1 + sekd_H2_dL1[i] * vek[2]);
                        tilltek_x2.Add(a_x2 - sekd_H2_dL1[i] * vek[0]);
                        tilltek_y2.Add(a_y2 - sekd_H2_dL1[i] * vek[1]);
                        tilltek_z2.Add(a_z2 - sekd_H2_dL1[i] * vek[2]);
                        tilltek_blt.Add(sekdata_H2blt[i]);
                        tilltek_bltN.Add(sekdata_H2bltN[i]);
                        tilltek_blte1.Add(sekdata_H2blte1[i]);
                        tilltek_bltp1.Add(sekdata_H2bltp1[i]);
                        tilltek_bltd0.Add(sekdata_H2bltd0[i]);
                        tilltek_idnr.Add(string.Format("N:{0}_H2:{1}", i, k));
                        tilltek_startnod.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k + 1]));
                        tilltek_slutnod.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k]));
                        tilltek_startPl.Add("");
                        tilltek_slutPl.Add("");
                        k = k + 1;
                    }
                }
            }

            // 3.3 Diagonaler (stolpkropp)
            for (int i = 0; i < sekdata_D1kdim.Count; i++)
            {
                int[] iternr = { 0, 90, 180, 270 };
                int[] hjap1 = { 0, 1, 2, 3, 0 };
                int k = 0;
                foreach (int ii in iternr)
                {
                    tilltek_typ.Add("D1");
                    tilltek_typ.Add("D1");
                    tilltek_dim.Add(sekdata_D1kdim[i]);
                    tilltek_dim.Add(sekdata_D1kdim[i]);
                    tilltek_kvalitet.Add(sekdata_D1kvalitet[i]);
                    tilltek_kvalitet.Add(sekdata_D1kvalitet[i]);
                    tilltek_blt.Add(sekdata_D1blt[i]);
                    tilltek_bltN.Add(sekdata_D1bltN[i]);
                    tilltek_blte1.Add(sekdata_D1blte1[i]);
                    tilltek_bltp1.Add(sekdata_D1bltp1[i]);
                    tilltek_bltd0.Add(sekdata_D1bltd0[i]);
                    tilltek_blt.Add(sekdata_D1blt[i]);
                    tilltek_bltN.Add(sekdata_D1bltN[i]);
                    tilltek_blte1.Add(sekdata_D1blte1[i]);
                    tilltek_bltp1.Add(sekdata_D1bltp1[i]);
                    tilltek_bltd0.Add(sekdata_D1bltd0[i]);
                    if (sekdata_D1typ[i] == 1 && sekdata_D1dorient[i] == 1)
                    {
                        tilltek_typ.Add("D1");
                        tilltek_typ.Add("D1");
                        tilltek_dim.Add(sekdata_D1kdim[i]);
                        tilltek_dim.Add(sekdata_D1kdim[i]);
                        tilltek_kvalitet.Add(sekdata_D1kvalitet[i]);
                        tilltek_kvalitet.Add(sekdata_D1kvalitet[i]);
                        tilltek_blt.Add(sekdata_D1blt[i]);
                        tilltek_bltN.Add(sekdata_D1bltN[i]);
                        tilltek_blte1.Add(sekdata_D1blte1[i]);
                        tilltek_bltp1.Add(sekdata_D1bltp1[i]);
                        tilltek_bltd0.Add(sekdata_D1bltd0[i]);
                        tilltek_blt.Add(sekdata_D1blt[i]);
                        tilltek_bltN.Add(sekdata_D1bltN[i]);
                        tilltek_blte1.Add(sekdata_D1blte1[i]);
                        tilltek_bltp1.Add(sekdata_D1bltp1[i]);
                        tilltek_bltd0.Add(sekdata_D1bltd0[i]);
                    }
                    if (sekdata_D1typ[i] == 1 && sekdata_D1dorient[i] == 1)
                    {
                        // Båda diagonalerna på utsidan av ramstång.
                        tilltek_pos.Add(1);
                        tilltek_pos.Add(1);
                        tilltek_pos.Add(1);
                        tilltek_pos.Add(1);

                        double a_x1 = Math.Abs(sekdata_3x[i] * Math.Sin(ii * Math.PI / 180) + sekdata_3y[i] * Math.Cos(ii * Math.PI / 180)) * Math.Sign(Math.Cos((ii - 45) * Math.PI / 180)) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        double a_y1 = Math.Abs(sekdata_3x[i] * Math.Cos(ii * Math.PI / 180) + sekdata_3y[i] * Math.Sin(ii * Math.PI / 180)) * Math.Sign(Math.Sin((ii - 45) * Math.PI / 180)) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        double a_z1 = sekdata_3z[i] + sekd_d1z[i];
                        double a_x2 = sekdata_6y[i] * Math.Cos(ii * Math.PI / 180) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        double a_y2 = sekdata_6y[i] * Math.Sin(ii * Math.PI / 180) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        double a_z2 = sekdata_6z[i] + sekd_d1z[i];
                        double[] vek = new double[3] { a_x2 - a_x1, a_y2 - a_y1, a_z2 - a_z1 };
                        vek = Normvek(vek);
                        tilltek_x1.Add(a_x1 + sekd_D1_dL1[i] * vek[0]);
                        tilltek_y1.Add(a_y1 + sekd_D1_dL1[i] * vek[1]);
                        tilltek_z1.Add(a_z1 + sekd_D1_dL1[i] * vek[2]);
                        tilltek_x2.Add(a_x2 - sekd_D1_dL2[i] * vek[0]);
                        tilltek_y2.Add(a_y2 - sekd_D1_dL2[i] * vek[1]);
                        tilltek_z2.Add(a_z2 - sekd_D1_dL2[i] * vek[2]);
                        a_x1 = sekdata_6y[i] * Math.Cos(ii * Math.PI / 180) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_y1 = sekdata_6y[i] * Math.Sin(ii * Math.PI / 180) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_z1 = sekdata_6z[i] + sekd_d1z[i];
                        a_x2 = Math.Abs(sekdata_4x[i] * Math.Sin(ii * Math.PI / 180) + sekdata_4y[i] * Math.Cos(ii * Math.PI / 180)) * Math.Sign(Math.Cos((ii + 45) * Math.PI / 180)) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_y2 = Math.Abs(sekdata_4x[i] * Math.Cos(ii * Math.PI / 180) + sekdata_4y[i] * Math.Sin(ii * Math.PI / 180)) * Math.Sign(Math.Sin((ii + 45) * Math.PI / 180)) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_z2 = sekdata_4z[i] + sekd_d1z[i];
                        double[] vekb = new double[3] { a_x2 - a_x1, a_y2 - a_y1, a_z2 - a_z1 };
                        vek = Normvek(vekb);
                        tilltek_x1.Add(a_x1 + sekd_D1_dL3[i] * vek[0]);
                        tilltek_y1.Add(a_y1 + sekd_D1_dL3[i] * vek[1]);
                        tilltek_z1.Add(a_z1 + sekd_D1_dL3[i] * vek[2]);
                        tilltek_x2.Add(a_x2 - sekd_D1_dL4[i] * vek[0]);
                        tilltek_y2.Add(a_y2 - sekd_D1_dL4[i] * vek[1]);
                        tilltek_z2.Add(a_z2 - sekd_D1_dL4[i] * vek[2]);
                        a_x1 = sekdata_6y[i] * Math.Cos(ii * Math.PI / 180) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_y1 = sekdata_6y[i] * Math.Sin(ii * Math.PI / 180) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_z1 = sekdata_6z[i] + sekd_d1z[i];
                        a_x2 = Math.Abs(sekdata_3x[i] * Math.Sin(ii * Math.PI / 180) + sekdata_3y[i] * Math.Cos(ii * Math.PI / 180)) * Math.Sign(Math.Cos((ii + 45) * Math.PI / 180)) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_y2 = Math.Abs(sekdata_3x[i] * Math.Cos(ii * Math.PI / 180) + sekdata_3y[i] * Math.Sin(ii * Math.PI / 180)) * Math.Sign(Math.Sin((ii + 45) * Math.PI / 180)) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_z2 = sekdata_3z[i] + sekd_d1z[i];
                        double[] vekc = new double[3] { a_x2 - a_x1, a_y2 - a_y1, a_z2 - a_z1 };
                        vek = Normvek(vekc);
                        tilltek_x1.Add(a_x1 + sekd_D1_dL2[i] * vek[0]);
                        tilltek_y1.Add(a_y1 + sekd_D1_dL2[i] * vek[1]);
                        tilltek_z1.Add(a_z1 + sekd_D1_dL2[i] * vek[2]);
                        tilltek_x2.Add(a_x2 - sekd_D1_dL1[i] * vek[0]);
                        tilltek_y2.Add(a_y2 - sekd_D1_dL1[i] * vek[1]);
                        tilltek_z2.Add(a_z2 - sekd_D1_dL1[i] * vek[2]);
                        a_x1 = Math.Abs(sekdata_4x[i] * Math.Sin(ii * Math.PI / 180) + sekdata_4y[i] * Math.Cos(ii * Math.PI / 180)) * Math.Sign(Math.Cos((ii - 45) * Math.PI / 180)) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_y1 = Math.Abs(sekdata_4x[i] * Math.Cos(ii * Math.PI / 180) + sekdata_4y[i] * Math.Sin(ii * Math.PI / 180)) * Math.Sign(Math.Sin((ii - 45) * Math.PI / 180)) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_z1 = sekdata_4z[i] + sekd_d1z[i];
                        a_x2 = sekdata_6y[i] * Math.Cos(ii * Math.PI / 180) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_y2 = sekdata_6y[i] * Math.Sin(ii * Math.PI / 180) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_z2 = sekdata_6z[i] + sekd_d1z[i];
                        double[] vekd = new double[3] { a_x2 - a_x1, a_y2 - a_y1, a_z2 - a_z1 };
                        vek = Normvek(vekd);
                        tilltek_x1.Add(a_x1 + sekd_D1_dL4[i] * vek[0]);
                        tilltek_y1.Add(a_y1 + sekd_D1_dL4[i] * vek[1]);
                        tilltek_z1.Add(a_z1 + sekd_D1_dL4[i] * vek[2]);
                        tilltek_x2.Add(a_x2 - sekd_D1_dL3[i] * vek[0]);
                        tilltek_y2.Add(a_y2 - sekd_D1_dL3[i] * vek[1]);
                        tilltek_z2.Add(a_z2 - sekd_D1_dL3[i] * vek[2]);
                        tilltek_rotera.Add(sekdata_D1rot[i]);
                        tilltek_rotera.Add(sekdata_D1rot[i]);
                        tilltek_rotera.Add(sekdata_D1rot[i]);
                        tilltek_rotera.Add(sekdata_D1rot[i]);
                        tilltek_idnr.Add(string.Format("N:{0}_D1:{1}.vn", i, k));
                        tilltek_idnr.Add(string.Format("N:{0}_D1:{1}.ho", i, k));
                        tilltek_idnr.Add(string.Format("N:{0}_D1:{1}.hn", i, k));
                        tilltek_idnr.Add(string.Format("N:{0}_D1:{1}.vo", i, k));
                        if (i > 0 && (dataGridView1.Rows[i - 1].Cells[6].Value != null))
                        {
                            tilltek_startnod.Add(string.Format("N:{0}_L1:{1}", Math.Max(i - 1, 0), hjap1[k]));
                            tilltek_slutnod.Add(string.Format("N:{0}_H1:{1}", i, k));
                            tilltek_startnod.Add(string.Format("N:{0}_H1:{1}", i, k));
                            tilltek_slutnod.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k + 1]));
                            tilltek_startnod.Add(string.Format("N:{0}_H1:{1}", i, k));
                            tilltek_slutnod.Add(string.Format("N:{0}_L1:{1}", Math.Max(i - 1, 0), hjap1[k + 1]));
                            tilltek_startnod.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k]));
                            tilltek_slutnod.Add(string.Format("N:{0}_H1:{1}", i, k));
                        }
                        else
                        {
                            tilltek_startnod.Add(string.Format("N:{0}_L1:{1}", Math.Max(i, 0), hjap1[k]));
                            tilltek_slutnod.Add(string.Format("N:{0}_H1:{1}", i, k));
                            tilltek_startnod.Add(string.Format("N:{0}_H1:{1}", i, k));
                            tilltek_slutnod.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k + 1]));
                            tilltek_startnod.Add(string.Format("N:{0}_H1:{1}", i, k));
                            tilltek_slutnod.Add(string.Format("N:{0}_L1:{1}", Math.Max(i, 0), hjap1[k + 1]));
                            tilltek_startnod.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k]));
                            tilltek_slutnod.Add(string.Format("N:{0}_H1:{1}", i, k));
                        }
                        tilltek_startPl.Add("");
                        tilltek_slutPl.Add(string.Format("N:{0}_PL3:{1}", i, k));
                        tilltek_startPl.Add(string.Format("N:{0}_PL3:{1}", i, k));
                        tilltek_slutPl.Add("");
                        tilltek_startPl.Add(string.Format("N:{0}_PL3:{1}", i, k));
                        tilltek_slutPl.Add("");
                        tilltek_startPl.Add("");
                        tilltek_slutPl.Add(string.Format("N:{0}_PL3:{1}", i, k));
                        k = k + 1;
                    }
                    else if (sekdata_D1typ[i] == 1)
                    {
                        // Ena diagonalen på utsida ramstång, andra på insida ramstång.
                        tilltek_pos.Add(1);
                        tilltek_pos.Add(2);
                        double a_x1 = Math.Abs(sekdata_3x[i] * Math.Sin(ii * Math.PI / 180) + sekdata_3y[i] * Math.Cos(ii * Math.PI / 180)) * Math.Sign(Math.Cos((ii - 45) * Math.PI / 180)) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        double a_y1 = Math.Abs(sekdata_3x[i] * Math.Cos(ii * Math.PI / 180) + sekdata_3y[i] * Math.Sin(ii * Math.PI / 180)) * Math.Sign(Math.Sin((ii - 45) * Math.PI / 180)) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        double a_z1 = sekdata_3z[i] + sekd_d1z[i];
                        double a_x2 = Math.Abs(sekdata_4x[i] * Math.Sin(ii * Math.PI / 180) + sekdata_4y[i] * Math.Cos(ii * Math.PI / 180)) * Math.Sign(Math.Cos((ii + 45) * Math.PI / 180)) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        double a_y2 = Math.Abs(sekdata_4x[i] * Math.Cos(ii * Math.PI / 180) + sekdata_4y[i] * Math.Sin(ii * Math.PI / 180)) * Math.Sign(Math.Sin((ii + 45) * Math.PI / 180)) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        double a_z2 = sekdata_4z[i] + sekd_d1z[i];
                        double[] vek = new double[3] { a_x2 - a_x1, a_y2 - a_y1, a_z2 - a_z1 };
                        vek = Normvek(vek);
                        tilltek_x1.Add(a_x1 + sekd_D1_dL1[i] * vek[0]);
                        tilltek_y1.Add(a_y1 + sekd_D1_dL1[i] * vek[1]);
                        tilltek_z1.Add(a_z1 + sekd_D1_dL1[i] * vek[2]);
                        tilltek_x2.Add(a_x2 - sekd_D1_dL4[i] * vek[0]);
                        tilltek_y2.Add(a_y2 - sekd_D1_dL4[i] * vek[1]);
                        tilltek_z2.Add(a_z2 - sekd_D1_dL4[i] * vek[2]);
                        a_x1 = Math.Abs(sekdata_3x[i] * Math.Sin(ii * Math.PI / 180) + sekdata_3y[i] * Math.Cos(ii * Math.PI / 180)) * Math.Sign(Math.Cos((ii + 45) * Math.PI / 180)) + Math.Cos((ii) * Math.PI / 180) * sekd_d2y[i];
                        a_y1 = Math.Abs(sekdata_3x[i] * Math.Cos(ii * Math.PI / 180) + sekdata_3y[i] * Math.Sin(ii * Math.PI / 180)) * Math.Sign(Math.Sin((ii + 45) * Math.PI / 180)) + Math.Sin((ii) * Math.PI / 180) * sekd_d2y[i];
                        a_z1 = sekdata_3z[i] + sekd_d2z[i];
                        a_x2 = Math.Abs(sekdata_4x[i] * Math.Sin(ii * Math.PI / 180) + sekdata_4y[i] * Math.Cos(ii * Math.PI / 180)) * Math.Sign(Math.Cos((ii - 45) * Math.PI / 180)) + Math.Cos((ii) * Math.PI / 180) * sekd_d2y[i];
                        a_y2 = Math.Abs(sekdata_4x[i] * Math.Cos(ii * Math.PI / 180) + sekdata_4y[i] * Math.Sin(ii * Math.PI / 180)) * Math.Sign(Math.Sin((ii - 45) * Math.PI / 180)) + Math.Sin((ii) * Math.PI / 180) * sekd_d2y[i];
                        a_z2 = sekdata_4z[i] + sekd_d2z[i];
                        double[] vekb = new double[3] { a_x2 - a_x1, a_y2 - a_y1, a_z2 - a_z1 };
                        vek = Normvek(vekb);
                        tilltek_x1.Add(a_x1 + sekd_D1_dL1[i] * vek[0]);
                        tilltek_y1.Add(a_y1 + sekd_D1_dL1[i] * vek[1]);
                        tilltek_z1.Add(a_z1 + sekd_D1_dL1[i] * vek[2]);
                        tilltek_x2.Add(a_x2 - sekd_D1_dL4[i] * vek[0]);
                        tilltek_y2.Add(a_y2 - sekd_D1_dL4[i] * vek[1]);
                        tilltek_z2.Add(a_z2 - sekd_D1_dL4[i] * vek[2]);
                        tilltek_rotera.Add(sekdata_D1rot[i]);
                        tilltek_rotera.Add(-sekdata_D1rot[i]);
                        tilltek_idnr.Add(string.Format("N:{0}_D1:{1}.5", i, k));
                        tilltek_idnr.Add(string.Format("N:{0}_D1:{1}.6", i, k));
                        if (i > 0 && (dataGridView1.Rows[i - 1].Cells[6].Value != null))
                        {
                            tilltek_startnod.Add(string.Format("N:{0}_L1:{1}", Math.Max(i - 1, 0), hjap1[k]));
                            tilltek_slutnod.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k + 1]));
                            tilltek_startnod.Add(string.Format("N:{0}_L1:{1}", Math.Max(i - 1, 0), hjap1[k + 1]));
                            tilltek_slutnod.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k]));
                        }
                        else
                        {
                            tilltek_startnod.Add(string.Format("N:{0}_L1:{1}", Math.Max(i, 0), hjap1[k]));
                            tilltek_slutnod.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k + 1]));
                            tilltek_startnod.Add(string.Format("N:{0}_L1:{1}", Math.Max(i, 0), hjap1[k + 1]));
                            tilltek_slutnod.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k]));
                        }
                        tilltek_startPl.Add("");
                        tilltek_slutPl.Add("");
                        tilltek_startPl.Add("");
                        tilltek_slutPl.Add("");
                        k = k + 1;
                    }
                    else if (sekdata_D1typ[i] == 2)
                    {
                        // K-sektion
                        tilltek_pos.Add(1);
                        tilltek_pos.Add(1);
                        double a_x1 = sekdata_5y[i] * Math.Cos(ii * Math.PI / 180) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        double a_y1 = sekdata_5y[i] * Math.Sin(ii * Math.PI / 180) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        double a_z1 = sekdata_5z[i] + sekd_d1z[i];
                        double a_x2 = Math.Abs(sekdata_3x[i] * Math.Sin(ii * Math.PI / 180) + sekdata_3y[i] * Math.Cos(ii * Math.PI / 180)) * Math.Sign(Math.Cos((ii + 45) * Math.PI / 180)) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        double a_y2 = Math.Abs(sekdata_3x[i] * Math.Cos(ii * Math.PI / 180) + sekdata_3y[i] * Math.Sin(ii * Math.PI / 180)) * Math.Sign(Math.Sin((ii + 45) * Math.PI / 180)) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        double a_z2 = sekdata_3z[i] + sekd_d1z[i];
                        double[] vek = new double[3] { a_x2 - a_x1, a_y2 - a_y1, a_z2 - a_z1 };
                        vek = Normvek(vek);
                        tilltek_x1.Add(a_x1 + sekd_D1_dL4[i] * vek[0]);
                        tilltek_y1.Add(a_y1 + sekd_D1_dL4[i] * vek[1]);
                        tilltek_z1.Add(a_z1 + sekd_D1_dL4[i] * vek[2]);
                        tilltek_x2.Add(a_x2 - sekd_D1_dL1[i] * vek[0]);
                        tilltek_y2.Add(a_y2 - sekd_D1_dL1[i] * vek[1]);
                        tilltek_z2.Add(a_z2 - sekd_D1_dL1[i] * vek[2]);
                        a_x1 = Math.Abs(sekdata_3x[i] * Math.Sin(ii * Math.PI / 180) + sekdata_3y[i] * Math.Cos(ii * Math.PI / 180)) * Math.Sign(Math.Cos((ii - 45) * Math.PI / 180)) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_y1 = Math.Abs(sekdata_3x[i] * Math.Cos(ii * Math.PI / 180) + sekdata_3y[i] * Math.Sin(ii * Math.PI / 180)) * Math.Sign(Math.Sin((ii - 45) * Math.PI / 180)) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_z1 = sekdata_3z[i] + sekd_d1z[i];
                        a_x2 = sekdata_5y[i] * Math.Cos(ii * Math.PI / 180) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_y2 = sekdata_5y[i] * Math.Sin(ii * Math.PI / 180) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_z2 = sekdata_5z[i] + sekd_d1z[i];
                        double[] vekb = new double[3] { a_x2 - a_x1, a_y2 - a_y1, a_z2 - a_z1 };
                        vek = Normvek(vekb);
                        tilltek_x1.Add(a_x1 + sekd_D1_dL1[i] * vek[0]);
                        tilltek_y1.Add(a_y1 + sekd_D1_dL1[i] * vek[1]);
                        tilltek_z1.Add(a_z1 + sekd_D1_dL1[i] * vek[2]);
                        tilltek_x2.Add(a_x2 - sekd_D1_dL4[i] * vek[0]);
                        tilltek_y2.Add(a_y2 - sekd_D1_dL4[i] * vek[1]);
                        tilltek_z2.Add(a_z2 - sekd_D1_dL4[i] * vek[2]);
                        tilltek_rotera.Add(sekdata_D1rot[i]);
                        tilltek_rotera.Add(sekdata_D1rot[i]);
                        tilltek_idnr.Add(string.Format("N:{0}_D1:{1}.7", i, k));
                        tilltek_idnr.Add(string.Format("N:{0}_D1:{1}.8", i, k));
                        if (i > 0 && (dataGridView1.Rows[i - 1].Cells[6].Value != null))
                        {
                            tilltek_startnod.Add(string.Format("N:{0}_H2:{1}", i, k));
                            tilltek_slutnod.Add(string.Format("N:{0}_L1:{1}", Math.Max(i - 1, 0), hjap1[k + 1]));
                            tilltek_startnod.Add(string.Format("N:{0}_L1:{1}", Math.Max(i - 1, 0), hjap1[k]));
                            tilltek_slutnod.Add(string.Format("N:{0}_H2:{1}", i, k));
                        }
                        else
                        {
                            tilltek_startnod.Add(string.Format("N:{0}_H2:{1}", i, k));
                            tilltek_slutnod.Add(string.Format("N:{0}_L1:{1}", Math.Max(i, 0), hjap1[k + 1]));
                            tilltek_startnod.Add(string.Format("N:{0}_L1:{1}", Math.Max(i, 0), hjap1[k]));
                            tilltek_slutnod.Add(string.Format("N:{0}_H2:{1}", i, k));
                        }
                        tilltek_startPl.Add(string.Format("N:{0}_PL4:{1}", i, k));
                        tilltek_slutPl.Add("");
                        tilltek_startPl.Add("");
                        tilltek_slutPl.Add(string.Format("N:{0}_PL4:{1}", i, k));
                        k = k + 1;
                    }
                    else if (sekdata_D1typ[i] == 3)
                    {
                        // Omvänd K-sektion
                        tilltek_pos.Add(1);
                        tilltek_pos.Add(1);
                        double a_x1 = Math.Abs(sekdata_4x[i] * Math.Sin(ii * Math.PI / 180) + sekdata_4y[i] * Math.Cos(ii * Math.PI / 180)) * Math.Sign(Math.Cos((ii - 45) * Math.PI / 180)) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        double a_y1 = Math.Abs(sekdata_4x[i] * Math.Cos(ii * Math.PI / 180) + sekdata_4y[i] * Math.Sin(ii * Math.PI / 180)) * Math.Sign(Math.Sin((ii - 45) * Math.PI / 180)) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        double a_z1 = sekdata_4z[i] + sekd_d1z[i];
                        double a_x2 = sekdata_8y[i] * Math.Cos(ii * Math.PI / 180) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        double a_y2 = sekdata_8y[i] * Math.Sin(ii * Math.PI / 180) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        double a_z2 = sekdata_8z[i] + sekd_d1z[i];
                        double[] vek = new double[3] { a_x2 - a_x1, a_y2 - a_y1, a_z2 - a_z1 };
                        vek = Normvek(vek);
                        tilltek_x1.Add(a_x1 + sekd_D1_dL4[i] * vek[0]);
                        tilltek_y1.Add(a_y1 + sekd_D1_dL4[i] * vek[1]);
                        tilltek_z1.Add(a_z1 + sekd_D1_dL4[i] * vek[2]);
                        tilltek_x2.Add(a_x2 - sekd_D1_dL1[i] * vek[0]);
                        tilltek_y2.Add(a_y2 - sekd_D1_dL1[i] * vek[1]);
                        tilltek_z2.Add(a_z2 - sekd_D1_dL1[i] * vek[2]);
                        a_x1 = sekdata_8y[i] * Math.Cos(ii * Math.PI / 180) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_y1 = sekdata_8y[i] * Math.Sin(ii * Math.PI / 180) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_z1 = sekdata_8z[i] + sekd_d1z[i];
                        a_x2 = Math.Abs(sekdata_4x[i] * Math.Sin(ii * Math.PI / 180) + sekdata_4y[i] * Math.Cos(ii * Math.PI / 180)) * Math.Sign(Math.Cos((ii + 45) * Math.PI / 180)) + Math.Cos((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_y2 = Math.Abs(sekdata_4x[i] * Math.Cos(ii * Math.PI / 180) + sekdata_4y[i] * Math.Sin(ii * Math.PI / 180)) * Math.Sign(Math.Sin((ii + 45) * Math.PI / 180)) + Math.Sin((ii) * Math.PI / 180) * sekd_d1y[i];
                        a_z2 = sekdata_4z[i] + sekd_d1z[i];
                        double[] vekb = new double[3] { a_x2 - a_x1, a_y2 - a_y1, a_z2 - a_z1 };
                        vek = Normvek(vekb);
                        tilltek_x1.Add(a_x1 + sekd_D1_dL1[i] * vek[0]);
                        tilltek_y1.Add(a_y1 + sekd_D1_dL1[i] * vek[1]);
                        tilltek_z1.Add(a_z1 + sekd_D1_dL1[i] * vek[2]);
                        tilltek_x2.Add(a_x2 - sekd_D1_dL4[i] * vek[0]);
                        tilltek_y2.Add(a_y2 - sekd_D1_dL4[i] * vek[1]);
                        tilltek_z2.Add(a_z2 - sekd_D1_dL4[i] * vek[2]);
                        tilltek_rotera.Add(sekdata_D1rot[i]);
                        tilltek_rotera.Add(sekdata_D1rot[i]);
                        tilltek_idnr.Add(string.Format("N:{0}_D1:{1}.9", i, k));
                        tilltek_idnr.Add(string.Format("N:{0}_D1:{1}.10", i, k));
                        if (i > 0 && (dataGridView1.Rows[i - 1].Cells[6].Value != null))
                        {
                            tilltek_startnod.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k]));
                            tilltek_slutnod.Add(string.Format("N:{0}_H2:{1}", Math.Max(i - 1, 0), k));
                            tilltek_startnod.Add(string.Format("N:{0}_H2:{1}", Math.Max(i - 1, 0), k));
                            tilltek_slutnod.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k + 1]));
                        }
                        else
                        {
                            tilltek_startnod.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k]));
                            tilltek_slutnod.Add(string.Format("N:{0}_H2:{1}", Math.Max(i, 0), k));
                            tilltek_startnod.Add(string.Format("N:{0}_H2:{1}", Math.Max(i, 0), k));
                            tilltek_slutnod.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k + 1]));
                        }
                        tilltek_startPl.Add("");
                        tilltek_slutPl.Add(string.Format("N:{0}_PL4:{1}", i - 1, k));
                        tilltek_startPl.Add(string.Format("N:{0}_PL4:{1}", i - 1, k));
                        tilltek_slutPl.Add("");
                        k = k + 1;
                    }
                }
            }

            // Plåtar
            for (int i = 0; i < antrader; i++)
            {
                // 1.1 Plåtar
                double tmp_PL2_t = 0;
                double tmp_PL3_t = 0;
                double tmp_PL4_t = 0;
                int tmp_PL2_b = 1;
                int tmp_PL3_b = 1;
                int tmp_PL4_b = 1;
                int tmp_PL2_n = 1;
                int tmp_PL3_n = 1;
                int tmp_PL4_n = 1;
                int tmp_PL2_blt = 1;
                int tmp_PL3_blt = 1;
                int tmp_PL4_blt = 1;
                double tmp_PL2_blte1 = 1;
                double tmp_PL3_blte1 = 1;
                double tmp_PL4_blte1 = 1;
                double tmp_PL2_bltp1 = 1;
                double tmp_PL3_bltp1 = 1;
                double tmp_PL4_bltp1 = 1;
                double tmp_PL2_bltd0 = 1;
                double tmp_PL3_bltd0 = 1;
                double tmp_PL4_bltd0 = 1;

                if ((dataGridView11.Rows[i].Cells[0].Value != null)) { double.TryParse(dataGridView11.Rows[i].Cells[0].Value.ToString(), out tmp_PL2_t); }
                if ((dataGridView11.Rows[i].Cells[1].Value != null)) { double.TryParse(dataGridView11.Rows[i].Cells[1].Value.ToString(), out tmp_PL3_t); }
                if ((dataGridView11.Rows[i].Cells[2].Value != null)) { double.TryParse(dataGridView11.Rows[i].Cells[2].Value.ToString(), out tmp_PL4_t); }
                if ((dataGridView10.Rows[i].Cells[0].Value != null)) { int.TryParse(dataGridView10.Rows[i].Cells[0].Value.ToString(), out tmp_PL2_b); }
                if ((dataGridView10.Rows[i].Cells[1].Value != null)) { int.TryParse(dataGridView10.Rows[i].Cells[1].Value.ToString(), out tmp_PL3_b); }
                if ((dataGridView10.Rows[i].Cells[2].Value != null)) { int.TryParse(dataGridView10.Rows[i].Cells[2].Value.ToString(), out tmp_PL4_b); }
                if ((dataGridView9.Rows[i].Cells[0].Value != null)) { int.TryParse(dataGridView9.Rows[i].Cells[0].Value.ToString(), out tmp_PL2_n); }
                if ((dataGridView9.Rows[i].Cells[1].Value != null)) { int.TryParse(dataGridView9.Rows[i].Cells[1].Value.ToString(), out tmp_PL3_n); }
                if ((dataGridView9.Rows[i].Cells[2].Value != null)) { int.TryParse(dataGridView9.Rows[i].Cells[2].Value.ToString(), out tmp_PL4_n); }
                double eimin = double.Parse(textBox3.Text);
                double pimin = double.Parse(textBox5.Text);
                tmp_PL2_blt = int.Parse(dataGridView7.Rows[tmp_PL2_b - 1].Cells[2].Value.ToString());
                tmp_PL3_blt = int.Parse(dataGridView7.Rows[tmp_PL3_b - 1].Cells[2].Value.ToString());
                tmp_PL4_blt = int.Parse(dataGridView7.Rows[tmp_PL4_b - 1].Cells[2].Value.ToString());
                tmp_PL2_bltd0 = double.Parse(dataGridView7.Rows[tmp_PL2_b - 1].Cells[5].Value.ToString());
                tmp_PL3_bltd0 = double.Parse(dataGridView7.Rows[tmp_PL3_b - 1].Cells[5].Value.ToString());
                tmp_PL4_bltd0 = double.Parse(dataGridView7.Rows[tmp_PL4_b - 1].Cells[5].Value.ToString());
                tmp_PL2_blte1 = eimin * tmp_PL2_bltd0;
                tmp_PL3_blte1 = eimin * tmp_PL3_bltd0;
                tmp_PL4_blte1 = eimin * tmp_PL4_bltd0;
                tmp_PL2_bltp1 = pimin * tmp_PL2_bltd0;
                tmp_PL3_bltp1 = pimin * tmp_PL3_bltd0;
                tmp_PL4_bltp1 = pimin * tmp_PL4_bltd0;

                int[] iternr = { 0, 1, 2, 3 };
                int[] hjap1 = { 0, 1, 2, 3, 0 };
                int k = 0;
                foreach (int ii in iternr)
                {

                    //if (sekdata_D1typ[i] == 1 &&  sekdata_D1dorient[i] == 1)
                    int typ11 = sekdata_D1typ[i];       // Typ av diogonalisering
                    int typ12 = sekdata_D1dorient[i];   // Typ av diogonalisering
                    int typ13 = 0;                      // H1
                    int typ14 = 0;                      // H2
                    int typ21 = 0;
                    int typ22 = 0;
                    int typ23 = 0;
                    int typ24 = 0;
                    if ((dataGridView2.Rows[i].Cells[2].Value != null)) { int.TryParse(dataGridView2.Rows[i].Cells[2].Value.ToString(), out typ13); }
                    if ((dataGridView2.Rows[i].Cells[3].Value != null)) { int.TryParse(dataGridView2.Rows[i].Cells[3].Value.ToString(), out typ14); }

                    if (i < antrader - 1)
                    {
                        typ21 = sekdata_D1typ[i + 1];
                        typ22 = sekdata_D1dorient[i + 1];
                        if ((dataGridView2.Rows[i + 1].Cells[2].Value != null)) { int.TryParse(dataGridView2.Rows[i + 1].Cells[2].Value.ToString(), out typ23); }
                        if ((dataGridView2.Rows[i + 1].Cells[3].Value != null)) { int.TryParse(dataGridView2.Rows[i + 1].Cells[3].Value.ToString(), out typ24); }
                    }


                    if (tmp_PL2_t != 0)
                    {
                        // !!!!!!!
                        // !!!!!!!
                        int radel = int.Parse(dataGridView2.Rows[i].Cells[0].Value.ToString()) - 1;
                        double bpl1 = double.Parse(dataGridView8.Rows[radel].Cells[4].Value.ToString());
                        //tpl1 = decimal.Parse(dataGridView8.Rows[radel].Cells[5].Value.ToString()) / 1000;
                        double czpl1 = double.Parse(dataGridView8.Rows[radel].Cells[6].Value.ToString());
                        double tpl1 = double.Parse(dataGridView8.Rows[radel].Cells[5].Value.ToString());
                        // Rätta !!!!!
                        string el1 = "";
                        string el2 = "";
                        string el3 = "";
                        string el4 = "";
                        string el5 = "";
                        string el6 = "";
                        int len = 0;
                        if (i < antrader - 1)
                        {
                            if (sekdata_D1typ[i + 1] == 1 && sekdata_D1dorient[i + 1] == 1)
                            {
                                el3 = string.Format("N:{0}_D1:{1}.vn", i + 1, ii);
                                el6 = string.Format("N:{0}_D1:{1}.hn", i + 1, ii);
                                len = len + 1;
                            }
                            else if (sekdata_D1typ[i + 1] == 1)
                            {
                                el3 = string.Format("N:{0}_D1:{1}.5", i + 1, ii);
                                el6 = string.Format("N:{0}_D1:{1}.6", i + 1, ii);
                                len = len + 1;
                            }
                            else if (sekdata_D1typ[i + 1] == 2)
                            {
                                el3 = string.Format("N:{0}_D1:{1}.8", i + 1, ii);
                                el6 = string.Format("N:{0}_D1:{1}.7", i + 1, ii);
                                len = len + 1;
                            }
                        }
                        if (sekdata_H2kdim[i] != string.Empty)
                        {
                            el2 = string.Format("N:{0}_H2:{1}", i, ii);
                            el5 = string.Format("N:{0}_H2:{1}", i, ii);
                            len = len + 1;
                        }


                        if (sekdata_D1typ[i] == 1 && sekdata_D1dorient[i] == 1)
                        {
                            el1 = string.Format("N:{0}_D1:{1}.vo", i, ii);
                            el4 = string.Format("N:{0}_D1:{1}.ho", i, ii);
                            len = len + 1;
                        }
                        else if (sekdata_D1typ[i] == 1)
                        {
                            el1 = string.Format("N:{0}_D1:{1}.6", i, ii);
                            el4 = string.Format("N:{0}_D1:{1}.5", i, ii);
                            len = len + 1;
                        }
                        else if (sekdata_D1typ[i] == 3)
                        {
                            el1 = string.Format("N:{0}_D1:{1}.96", i, ii);
                            el4 = string.Format("N:{0}_D1:{1}.10", i, ii);
                            len = len + 1;
                        }
                        if (len > 0)
                        {
                            string[] elPl2_1 = new string[len];
                            string[] elPl2_2 = new string[len];
                            int j = 0;
                            if (el3 != "")
                            {
                                elPl2_1[j] = el3;
                                elPl2_2[j] = el6;
                                j = j + 1;
                            }
                            if (el2 != "")
                            {
                                elPl2_1[j] = el2;
                                elPl2_2[j] = el5;
                                j = j + 1;
                            }
                            if (el1 != "")
                            {
                                elPl2_1[j] = el1;
                                elPl2_2[j] = el4;
                                j = j + 1;
                            }

                            double forskj = 0;
                            if ((dataGridView5.Rows[i].Cells[1].Value != null)) { double.TryParse(dataGridView5.Rows[i].Cells[1].Value.ToString(), out forskj); }
                            tilltekPl_idnr.Add(string.Format("N:{0}_PL2.1:{1}", i, ii));
                            tilltekPl_idnr.Add(string.Format("N:{0}_PL2.1:{1}", i, ii));
                            tilltekPl_diag.Add(elPl2_1);
                            tilltekPl_diag.Add(elPl2_2);
                            tilltekPl_cntr.Add(forskj);
                            tilltekPl_cntr.Add(forskj);
                            //tilltekPl_GUID.Add();
                            tilltekPl_typ.Add(2);
                            tilltekPl_typ.Add(2);
                            tilltekPl_t.Add(tmp_PL2_t);
                            tilltekPl_t.Add(tmp_PL2_t);
                            tilltekPl_ram.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k]));
                            tilltekPl_ram.Add(string.Format("N:{0}_L1:{1}", i, hjap1[k + 1]));
                            tilltekPl_ramc.Add(czpl1);
                            tilltekPl_ramc.Add(czpl1);
                            tilltekPl_ramb.Add(bpl1);
                            tilltekPl_ramb.Add(bpl1);
                            tilltekPl_ramt.Add(tpl1);
                            tilltekPl_ramt.Add(tpl1);
                            tilltekPl_blt.Add(tmp_PL2_blt);
                            tilltekPl_blt.Add(tmp_PL2_blt);
                            tilltekPl_bltN.Add(tmp_PL2_n);
                            tilltekPl_bltN.Add(tmp_PL2_n);
                            tilltekPl_blte1.Add(tmp_PL2_blte1);
                            tilltekPl_blte1.Add(tmp_PL2_blte1);
                            tilltekPl_bltp1.Add(tmp_PL2_bltp1);
                            tilltekPl_bltp1.Add(tmp_PL2_bltp1);
                            tilltekPl_bltd0.Add(tmp_PL2_bltd0);
                            tilltekPl_bltd0.Add(tmp_PL2_bltd0);
                        }
                    }



                    if ((tmp_PL3_t != 0) && (typ11 == 1) && (typ12 == 1))
                    {
                        int radel = int.Parse(dataGridView2.Rows[i].Cells[2].Value.ToString()) - 1;
                        double bpl1 = double.Parse(dataGridView8.Rows[radel].Cells[4].Value.ToString());
                        //tpl1 = decimal.Parse(dataGridView8.Rows[radel].Cells[5].Value.ToString()) / 1000;
                        double czpl1 = double.Parse(dataGridView8.Rows[radel].Cells[6].Value.ToString());
                        double tpl1 = double.Parse(dataGridView8.Rows[radel].Cells[5].Value.ToString());
                        string el1 = string.Format("N:{0}_D1:{1}.vn", i, ii);
                        string el2 = string.Format("N:{0}_D1:{1}.hn", i, ii);
                        string el3 = string.Format("N:{0}_D1:{1}.ho", i, ii);
                        string el4 = string.Format("N:{0}_D1:{1}.vo", i, ii);
                        string[] elPl3_1 = new string[] { el1, el2, el3, el4 };
                        double forskj = 0;
                        if ((dataGridView5.Rows[i].Cells[1].Value != null)) { double.TryParse(dataGridView5.Rows[i].Cells[1].Value.ToString(), out forskj); }
                        tilltekPl_idnr.Add(string.Format("N:{0}_PL3:{1}", i, ii));
                        tilltekPl_cntr.Add(forskj);
                        //tilltekPl_GUID.Add();
                        tilltekPl_typ.Add(3);
                        tilltekPl_t.Add(tmp_PL3_t);
                        tilltekPl_ram.Add(string.Format("N:{0}_H1:{1}", i, ii));
                        tilltekPl_ramc.Add(czpl1);
                        tilltekPl_ramb.Add(bpl1);
                        tilltekPl_ramt.Add(tpl1);
                        tilltekPl_diag.Add(elPl3_1);
                        tilltekPl_blt.Add(tmp_PL3_blt);
                        tilltekPl_bltN.Add(tmp_PL3_n);
                        tilltekPl_blte1.Add(tmp_PL3_blte1);
                        tilltekPl_bltp1.Add(tmp_PL3_bltp1);
                        tilltekPl_bltd0.Add(tmp_PL3_bltd0);
                    }

                    if (tmp_PL4_t != 0)
                    {
                        if ((typ11 == 2) || ((i < antrader - 1) && (sekdata_D1typ[i + 1] == 3)))
                        {
                            //sekdata_D1typ[i]
                            int radel = int.Parse(dataGridView2.Rows[i].Cells[3].Value.ToString()) - 1;
                            double bpl1 = double.Parse(dataGridView8.Rows[radel].Cells[4].Value.ToString());
                            double czpl1 = double.Parse(dataGridView8.Rows[radel].Cells[6].Value.ToString());
                            double tpl1 = double.Parse(dataGridView8.Rows[radel].Cells[5].Value.ToString());
                            string[] elPl3_1;
                            if ((typ11 == 2) && ((i < antrader - 1) && (sekdata_D1typ[i + 1] == 3)))
                            {
                                string el1 = string.Format("N:{0}_D1:{1}.7", i, ii);
                                string el2 = string.Format("N:{0}_D1:{1}.8", i, ii);
                                string el3 = string.Format("N:{0}_D1:{1}.9", i + 1, ii);
                                string el4 = string.Format("N:{0}_D1:{1}.10", i + 1, ii);
                                elPl3_1 = new string[] { el1, el2, el3, el4 };
                            }
                            else if ((i < antrader - 1) && (sekdata_D1typ[i + 1] == 3))
                            {
                                string el3 = string.Format("N:{0}_D1:{1}.9", i + 1, ii);
                                string el4 = string.Format("N:{0}_D1:{1}.10", i + 1, ii);
                                elPl3_1 = new string[] { el3, el4 };
                            }
                            else
                            {
                                string el1 = string.Format("N:{0}_D1:{1}.7", i, ii);
                                string el2 = string.Format("N:{0}_D1:{1}.8", i, ii);
                                elPl3_1 = new string[] { el1, el2};
                            }
                            double forskj = 0;
                            if ((dataGridView5.Rows[i].Cells[1].Value != null)) { double.TryParse(dataGridView5.Rows[i].Cells[1].Value.ToString(), out forskj); }
                            tilltekPl_idnr.Add(string.Format("N:{0}_PL4:{1}", i, ii));
                            tilltekPl_cntr.Add(forskj);
                            //tilltekPl_GUID.Add();
                            tilltekPl_typ.Add(4);
                            tilltekPl_t.Add(tmp_PL4_t);
                            tilltekPl_ram.Add(string.Format("N:{0}_H2:{1}", i, ii));
                            tilltekPl_ramc.Add(czpl1);
                            tilltekPl_ramb.Add(bpl1);
                            tilltekPl_ramt.Add(tpl1);
                            tilltekPl_diag.Add(elPl3_1);
                            tilltekPl_blt.Add(tmp_PL4_blt);
                            tilltekPl_bltN.Add(tmp_PL4_n);
                            tilltekPl_blte1.Add(tmp_PL4_blte1);
                            tilltekPl_bltp1.Add(tmp_PL4_bltp1);
                            tilltekPl_bltd0.Add(tmp_PL4_bltd0);
                        }
                    }
                    k = k + 1;
                }
            }


        }

        private void RitaTekla()
        {

            if (!myModel.GetConnectionStatus()) { MessageBox.Show("Inget anslutning mot Tekla"); }
            else
            {
                //TransformationPlane aa = myModel.GetWorkPlaneHandler().GetCurrentTransformationPlane();
                ModelInfo modelinfo = myModel.GetInfo();
                string name = modelinfo.ModelName;
                MessageBox.Show(string.Format("Hej! Namnet på denna modell är: {0}", name));

                // Skapar stänger
                for (int i = 0; i < this.tilltek_idnr.Count; i++)
                {
                    Beam myBeam = new Beam();
                    myBeam.Material.MaterialString = tilltek_kvalitet[i];
                    myBeam.Profile.ProfileString = tilltek_dim[i];
                    myBeam.Class = "3";
                    myBeam.StartPoint.X = tilltek_x1[i] * 1000;
                    myBeam.StartPoint.Y = tilltek_y1[i] * 1000;
                    myBeam.StartPoint.Z = tilltek_z1[i] * 1000;
                    myBeam.EndPoint.X = tilltek_x2[i] * 1000;
                    myBeam.EndPoint.Y = tilltek_y2[i] * 1000;
                    myBeam.EndPoint.Z = tilltek_z2[i] * 1000;
                    myBeam.Position.Depth = Position.DepthEnum.MIDDLE;
                    myBeam.Position.Rotation = Position.RotationEnum.TOP;
                    if (tilltek_typ[i] == "L1")
                    {
                        myBeam.Position.Plane = Position.PlaneEnum.MIDDLE;
                    }
                    else if (tilltek_typ[i] == "D1")
                    {
                        myBeam.Position.Plane = Position.PlaneEnum.RIGHT;
                    }
                    else if (tilltek_typ[i] == "H1" || tilltek_typ[i] == "H2")
                    {
                        myBeam.Position.Plane = Position.PlaneEnum.RIGHT;
                        myBeam.Position.Rotation = Position.RotationEnum.BACK;
                        myBeam.Position.Depth = Position.DepthEnum.MIDDLE;
                    }
                    myBeam.Position.RotationOffset = tilltek_rotera[i];
                    myBeam.Insert();
                    tilltek_GUID.Add(myBeam.Identifier.GUID); //GUID
                }
                
                // Skapar plåtar
                for (int i = 0; i < this.tilltekPl_idnr.Count; i++)
                {
                    Beam objekt1 = new Beam();
                    Beam objekt2 = new Beam();
                    string elid = tilltekPl_ram[i];
                    double[] x1 = new double[tilltekPl_diag[i].Length];
                    double[] y1 = new double[tilltekPl_diag[i].Length];
                    double[] z1 = new double[tilltekPl_diag[i].Length];
                    double[] x2 = new double[tilltekPl_diag[i].Length];
                    double[] y2 = new double[tilltekPl_diag[i].Length];
                    double[] z2 = new double[tilltekPl_diag[i].Length];

                    for (int k = 0; k < this.tilltek_idnr.Count; k++)
                    {
                        if (tilltek_idnr[k] == elid)
                        {
                            objekt1.Identifier.GUID = tilltek_GUID[k];
                            objekt1.Select();
                        }
                    }
                    //MessageBox.Show(string.Format("d0: {0}", Bram));
                    //ArrayList sNames = new ArrayList();
                    //ArrayList iNames = new ArrayList();
                    //ArrayList dNames = new ArrayList();
                    //dNames.Add("HEIGHT");
                    //dNames.Add("WIDTH");
                    //Hashtable sValues = new Hashtable(sNames.Count + dNames.Count + iNames.Count);
                    //objekt1.GetAllReportProperties(sNames, dNames, iNames, ref sValues);
                    //double Bram = double.Parse(sValues["WIDTH"].ToString());
                    //double Hram = double.Parse(sValues["HEIGHT"].ToString());
                    double Bram = tilltekPl_ramb[i];
                    double Cram = tilltekPl_ramc[i];
                    //double Lram = Math.Sqrt(Math.Pow((objekt1.EndPoint.X - objekt1.StartPoint.X), 2) + Math.Pow((objekt1.EndPoint.Y - objekt1.StartPoint.Y), 2) + Math.Pow((objekt1.EndPoint.Z - objekt1.StartPoint.Z), 2));
                    myModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane(objekt1.GetCoordinateSystem()));

                    // Beräknar mittpunkten för plåtens skruvrad + bestäm arbetsplan (x-z eller y-z)
                    for (int k = 0; k < this.tilltek_idnr.Count; k++)
                    {
                        for (int j = 0; j < this.tilltekPl_diag[i].Length; j++)
                        {
                            if (tilltek_idnr[k] == tilltekPl_diag[i][j])
                            {
                                objekt2.Identifier.GUID = tilltek_GUID[k];
                                objekt2.Select();
                                x1[j] = objekt2.StartPoint.X;
                                y1[j] = objekt2.StartPoint.Y;
                                z1[j] = objekt2.StartPoint.Z;
                                x2[j] = objekt2.EndPoint.X;
                                y2[j] = objekt2.EndPoint.Y;
                                z2[j] = objekt2.EndPoint.Z;
                            }
                        }
                    }
                    int arbplan = 1;
                    if (Math.Abs(y2[0] - y1[0]) > Math.Abs(z2[0]- z1[0]))
                    {
                        arbplan = 2;
                    }
                    double[] p = { x1[0], y1[0], z1[0] };
                    double[] q = { x1[1], y1[1], z1[1] };
                    double[] r = { x2[0] - x1[0], y2[0] - y1[0], z2[0] - z1[0] };
                    double[] s = { x2[1] - x1[1], y2[1] - y1[1], z2[1] - z1[1] };
                    double[] delt = Vektkors(p, q, r, s);
                    double xlok = p[0] + r[0] * delt[0];

                    // Beräkna plåtens fyra knutpunkter
                    double xmax = Bram / 2;
                    double xmin = -Bram / 2;
                    double ymax = Bram / 2;
                    double ymin = -Bram / 2;
                    ArrayList sNames = new ArrayList();
                    ArrayList iNames = new ArrayList();
                    ArrayList dNames = new ArrayList();
                    dNames.Add("HEIGHT");
                    dNames.Add("WIDTH");
                    Hashtable sValues = new Hashtable(sNames.Count + dNames.Count + iNames.Count);
                    for (int k = 0; k < this.tilltek_idnr.Count; k++)
                    {
                        for (int j = 0; j < this.tilltekPl_diag[i].Length; j++)
                        {
                            if (tilltek_idnr[k] == tilltekPl_diag[i][j])
                            {
                                objekt2.Identifier.GUID = tilltek_GUID[k];
                                objekt2.Select();
                                double x1_tmp = objekt2.StartPoint.X;
                                double x2_tmp = objekt2.EndPoint.X;
                                double y1_tmp = objekt2.StartPoint.Y;
                                double y2_tmp = objekt2.EndPoint.Y;
                                double z1_tmp = objekt2.StartPoint.Z;
                                double z2_tmp = objekt2.EndPoint.Z;
                                double yz1_tmp = y1_tmp;
                                double yz2_tmp = y2_tmp;
                                if (arbplan == 1)
                                {
                                    yz1_tmp = z1_tmp;
                                    yz2_tmp = z2_tmp;
                                }
                                objekt2.GetAllReportProperties(sNames, dNames, iNames, ref sValues);
                                double Bdiag = double.Parse(sValues["WIDTH"].ToString());
                                double Lskr = tilltek_blte1[k] * 2 + (tilltek_bltN[k] - 1) * tilltek_bltp1[k];
                                double Ldiag = Math.Sqrt(Math.Pow((x2_tmp - x1_tmp), 2) + Math.Pow((yz2_tmp - yz1_tmp), 2));
                                double avst1 = Math.Sqrt(Math.Pow((x1_tmp - xlok), 2) + Math.Pow((yz1_tmp), 2));
                                double avst2 = Math.Sqrt(Math.Pow((x2_tmp - xlok), 2) + Math.Pow((yz2_tmp), 2));
                                double[] Vek = { (x2_tmp - x1_tmp) / Ldiag, (yz2_tmp - yz1_tmp) / Ldiag };
                                double[] Nvek = { Vek[0], -Vek[1] };
                                double avsX1 = 0;
                                double avsX2 = 0;
                                double avsY1 = 0;
                                double avsY2 = 0;

                                if (avst2 >= avst1)
                                {
                                    avsX1 = x1_tmp + Vek[0] * Lskr + Nvek[0] * Bdiag / 2 - xlok;
                                    avsX2 = x1_tmp + Vek[0] * Lskr + Nvek[0] * -Bdiag / 2 - xlok;
                                    avsY1 = yz1_tmp + Vek[1] * Lskr + Nvek[1] * Bdiag / 2;
                                    avsY2 = yz1_tmp + Vek[1] * Lskr + Nvek[1] * -Bdiag / 2;
                                }
                                else
                                {
                                    avsX1 = x1_tmp + Vek[0] * (Ldiag - Lskr) + Nvek[0] * Bdiag / 2 - xlok;
                                    avsX2 = x1_tmp + Vek[0] * (Ldiag - Lskr) + Nvek[0] * -Bdiag / 2 - xlok;
                                    avsY1 = yz1_tmp + Vek[1] * (Ldiag - Lskr) + Nvek[1] * Bdiag / 2;
                                    avsY2 = yz1_tmp + Vek[1] * (Ldiag - Lskr) + Nvek[1] * -Bdiag / 2;
                                }
                                if (Math.Max(avsX1, avsX2) > xmax) { xmax = Math.Max(avsX1, avsX2); }
                                if (Math.Max(avsY1, avsY2) > ymax) { ymax = Math.Max(avsY1, avsY2); }
                                if (Math.Min(avsX1, avsX2) < xmin) { xmin = Math.Min(avsX1, avsX2); }
                                if (Math.Min(avsY1, avsY2) < ymin) { ymin = Math.Min(avsY1, avsY2); }

                                // Anpassa diagonal till plåt
                                //objekt2.StartPoint.Y = objekt2.StartPoint.Y - tilltekPl_ramt[i];
                                //objekt2.StartPoint.Z = objekt2.StartPoint.Z - tilltekPl_ramt[i];
                                //objekt2.EndPoint.Y = objekt2.EndPoint.Y - tilltekPl_ramt[i];
                                //objekt2.EndPoint.Z = objekt2.EndPoint.Z - tilltekPl_ramt[i];
                                if ((avst2 >= avst1) && (tilltekPl_typ[i] == 2))
                                {
                                    if (arbplan == 1)
                                    {
                                        if (objekt2.StartPoint.Y <= -Bram / 2 + tilltekPl_ramt[i]/2)
                                        {
                                            objekt2.StartPoint.Y = objekt2.StartPoint.Y - tilltekPl_t[i];
                                        }
                                        else if (Math.Abs(objekt2.StartPoint.Z) > Bram / 2)
                                        {
                                            objekt2.StartPoint.Y = objekt2.StartPoint.Y - tilltekPl_ramt[i];
                                        }
                                    }
                                    else
                                    {
                                        if (objekt2.StartPoint.Z <= -Bram / 2 + tilltekPl_ramt[i] / 2)
                                        {
                                            objekt2.StartPoint.Z = objekt2.StartPoint.Z - tilltekPl_t[i];
                                        }
                                        else if (Math.Abs(objekt2.StartPoint.Y) > Bram / 2)
                                        {
                                            objekt2.StartPoint.Z = objekt2.StartPoint.Z - tilltekPl_ramt[i];
                                        }
                                    }
                                }
                                else if (tilltekPl_typ[i] == 2)
                                {
                                    if (arbplan == 1)
                                    {
                                        if (objekt2.EndPoint.Y <= -Bram / 2 + tilltekPl_ramt[i] / 2)
                                        {
                                            objekt2.EndPoint.Y = objekt2.EndPoint.Y - tilltekPl_t[i];
                                        }
                                        else if (Math.Abs(objekt2.EndPoint.Z) > Bram / 2)
                                        {
                                            objekt2.EndPoint.Y = objekt2.EndPoint.Y - tilltekPl_ramt[i];
                                        }
                                    }
                                    else
                                    {
                                        if (objekt2.EndPoint.Z <= -Bram / 2 + tilltekPl_ramt[i] / 2)
                                        {
                                            objekt2.EndPoint.Z = objekt2.EndPoint.Z - tilltekPl_t[i];
                                        }
                                        else if (Math.Abs(objekt2.EndPoint.Y) > Bram / 2)
                                        {
                                            objekt2.EndPoint.Z = objekt2.EndPoint.Z - tilltekPl_ramt[i];
                                        }
                                    }
                                }
                                objekt2.Modify();
                            }
                        }
                    }

                    ContourPlate NyPlat = new ContourPlate();
                    NyPlat.AssemblyNumber.Prefix = "";
                    NyPlat.AssemblyNumber.StartNumber = 0;
                    NyPlat.PartNumber.Prefix = "PL";
                    NyPlat.PartNumber.StartNumber = 9001;
                    NyPlat.Name = "PLÅT";
                    NyPlat.Profile.ProfileString = string.Format("PL{0}", tilltekPl_t[i]);
                    NyPlat.Material.MaterialString = "S355J2";
                    NyPlat.Finish = "";
                    NyPlat.Class = "99";
                    //for (int k = 0; k < tilltekPl_diag.Count; k++)
                    //{
                    //    NyPlat.AddBoltDistX(tilltekPl_diag[k]);
                    //}
                    //NyPlat.AddContourPoint(new T3D.Point(100, 0, 0););
                    //MessageBox.Show(string.Format("d0: {0}", xmax));
                    //MessageBox.Show(string.Format("d0: {0}", xmin));
                    //MessageBox.Show(string.Format("d0: {0}", ymax));
                    //MessageBox.Show(string.Format("d0: {0}", ymin));

                    if (arbplan == 1)
                    {
                        NyPlat.Position.Depth = Position.DepthEnum.FRONT;
                        NyPlat.AddContourPoint(new ContourPoint(new T3D.Point(xlok + xmax, -Bram / 2, ymin), new Chamfer(0, 0, Chamfer.ChamferTypeEnum.CHAMFER_LINE)));
                        NyPlat.AddContourPoint(new ContourPoint(new T3D.Point(xlok + xmax, -Bram / 2, ymax), new Chamfer(0, 0, Chamfer.ChamferTypeEnum.CHAMFER_LINE)));
                        NyPlat.AddContourPoint(new ContourPoint(new T3D.Point(xlok + xmin, -Bram / 2, ymax), new Chamfer(0, 0, Chamfer.ChamferTypeEnum.CHAMFER_LINE)));
                        NyPlat.AddContourPoint(new ContourPoint(new T3D.Point(xlok + xmin, -Bram / 2, ymin), new Chamfer(0, 0, Chamfer.ChamferTypeEnum.CHAMFER_LINE)));
                    }
                    else
                    {
                        NyPlat.Position.Depth = Position.DepthEnum.BEHIND;
                        NyPlat.AddContourPoint(new ContourPoint(new T3D.Point(xlok + xmax, ymin, -Bram / 2), new Chamfer(0, 0, Chamfer.ChamferTypeEnum.CHAMFER_LINE)));
                        NyPlat.AddContourPoint(new ContourPoint(new T3D.Point(xlok + xmax, ymax, -Bram / 2), new Chamfer(0, 0, Chamfer.ChamferTypeEnum.CHAMFER_LINE)));
                        NyPlat.AddContourPoint(new ContourPoint(new T3D.Point(xlok + xmin, ymax, -Bram / 2), new Chamfer(0, 0, Chamfer.ChamferTypeEnum.CHAMFER_LINE)));
                        NyPlat.AddContourPoint(new ContourPoint(new T3D.Point(xlok + xmin, ymin, -Bram / 2), new Chamfer(0, 0, Chamfer.ChamferTypeEnum.CHAMFER_LINE)));
                    }
                    NyPlat.Insert();
                    tilltekPl_GUID.Add(NyPlat.Identifier.GUID);

                    // Skapar bultar
                    if ((tilltekPl_typ[i] == 3) || (tilltekPl_typ[i] == 4) || (tilltekPl_typ[i] == 2))
                    {

                        BoltArray NyBultgrupp = new BoltArray();
                        NyBultgrupp.BoltSize = tilltekPl_blt[i];
                        NyBultgrupp.Tolerance = tilltekPl_bltd0[i] - tilltekPl_blt[i];
                        if (tilltekPl_bltN[i] == 1) { NyBultgrupp.AddBoltDistX(0); }
                        else
                        {
                            for (int k = 0; k < tilltekPl_bltN[i] - 1; k++)
                            {
                                NyBultgrupp.AddBoltDistX(tilltekPl_bltp1[i]);
                            }
                        }

                        NyBultgrupp.BoltStandard = "EN 14399-4-FZV";
                        NyBultgrupp.AddBoltDistY(0);
                        NyBultgrupp.StartPointOffset.Dx = tilltekPl_blte1[i];
                        NyBultgrupp.BoltType = BoltGroup.BoltTypeEnum.BOLT_TYPE_SITE;
                        NyBultgrupp.CutLength = 200;
                        NyBultgrupp.Position.Rotation = Position.RotationEnum.BELOW;

                        NyBultgrupp.PartToBoltTo = objekt1;
                        NyBultgrupp.PartToBeBolted = NyPlat;

                        if (arbplan == 1)
                        {
                            NyBultgrupp.Position.Rotation = Position.RotationEnum.TOP;
                        }
                        else
                        {
                            NyBultgrupp.Position.Rotation = Position.RotationEnum.BACK;
                        }
                        if (tilltekPl_typ[i] == 2)
                        {
                            NyBultgrupp.FirstPosition = new T3D.Point(xlok - tilltekPl_blte1[i] - 0.5 * (tilltekPl_bltN[i] - 1) * tilltekPl_bltp1[i], -tilltekPl_ramc[i], -tilltekPl_ramc[i]);
                            NyBultgrupp.SecondPosition = new T3D.Point(xlok + 100, -tilltekPl_ramc[i], -tilltekPl_ramc[i]);
                        }
                        else
                        {
                            NyBultgrupp.FirstPosition = new T3D.Point(xlok - tilltekPl_blte1[i] - 0.5 * (tilltekPl_bltN[i] - 1) * tilltekPl_bltp1[i], 0, 0);
                            NyBultgrupp.SecondPosition = new T3D.Point(xlok + 100, 0, 0);
                        }
                        NyBultgrupp.Insert();
                        myModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());
                    }  
                }


                // Skapar bultar
                for (int i = 0; i < this.tilltek_idnr.Count; i++)
                {
                    if (tilltek_typ[i] == "D1" || tilltek_typ[i] == "H1" || tilltek_typ[i] == "H2")
                    {
                        Beam objekt1 = new Beam();
                        Beam objekt2 = new Beam();
                        Beam objekt3 = new Beam();
                        ContourPlate objekt4 = new ContourPlate();
                        ContourPlate objekt5 = new ContourPlate();
                        string startID = tilltek_startnod[i];
                        string slutiID = tilltek_slutnod[i];
                        string startPlID = tilltek_startPl[i];
                        string slutPlID = tilltek_slutPl[i];
                        bool Plstart = false;
                        bool Plslut = false;
                        double Lst = 1000 * Math.Sqrt(Math.Pow((tilltek_x2[i] - tilltek_x1[i]), 2) + Math.Pow((tilltek_y2[i] - tilltek_y1[i]), 2) + Math.Pow((tilltek_z2[i] - tilltek_z1[i]), 2));
                        objekt1.Identifier.GUID = tilltek_GUID[i];
                        objekt1.Select();
                        for (int k = 0; k < this.tilltek_idnr.Count; k++)
                        {
                            if (tilltek_idnr[k] == startID)
                            {
                                objekt2.Identifier.GUID = tilltek_GUID[k];
                                objekt2.Select();
                            }
                            if (tilltek_idnr[k] == slutiID)
                            {
                                objekt3.Identifier.GUID = tilltek_GUID[k];
                                objekt3.Select();
                            }
                        }
                        for (int k = 0; k < this.tilltekPl_idnr.Count; k++)
                        {
                            if (tilltekPl_idnr[k] == startPlID)
                            {
                                objekt4.Identifier.GUID = tilltekPl_GUID[k];
                                objekt4.Select();
                                Plstart = true;
                            }
                            if (tilltekPl_idnr[k] == slutPlID)
                            {
                                objekt5.Identifier.GUID = tilltekPl_GUID[k];
                                objekt5.Select();
                                Plslut = true;
                            }
                        }

                        myModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane(objekt1.GetCoordinateSystem()));
                        BoltArray NyBultgrupp = new BoltArray();
                        NyBultgrupp.BoltSize = tilltek_blt[i];
                        NyBultgrupp.Tolerance = tilltek_bltd0[i] - tilltek_blt[i];
                        if (tilltek_bltN[i] == 1) { NyBultgrupp.AddBoltDistX(0); }
                        else
                        {
                            for (int k = 0; k < tilltek_bltN[i] - 1; k++)
                            {
                                NyBultgrupp.AddBoltDistX(tilltek_bltp1[i]);
                            }
                        }

                        NyBultgrupp.BoltStandard = "EN 14399-4-FZV";
                        NyBultgrupp.AddBoltDistY(0);
                        NyBultgrupp.StartPointOffset.Dx = tilltek_blte1[i];
                        NyBultgrupp.BoltType = BoltGroup.BoltTypeEnum.BOLT_TYPE_SITE;
                        NyBultgrupp.Position.Rotation = Position.RotationEnum.FRONT;
                        NyBultgrupp.CutLength = 200;
                        if (tilltek_typ[i] == "H1" || tilltek_typ[i] == "H2")
                        {
                            NyBultgrupp.Position.Rotation = Position.RotationEnum.BELOW;
                        }
                        else if (tilltek_pos[i] == 2)
                        {
                            NyBultgrupp.Position.Rotation = Position.RotationEnum.BACK;
                        }
                        else
                        {
                            NyBultgrupp.Position.Rotation = Position.RotationEnum.FRONT;
                        }
                        NyBultgrupp.PartToBeBolted = objekt1;
                        if (Plslut)
                        {
                            //NyBultgrupp.AddOtherPartToBolt(objekt5);
                            NyBultgrupp.PartToBoltTo = objekt5;
                        }
                        else
                        {
                            NyBultgrupp.PartToBoltTo = objekt3;
                        }
                        NyBultgrupp.FirstPosition = new T3D.Point(Lst, 0, 0);
                        NyBultgrupp.SecondPosition = new T3D.Point(0, 0, 0);
                        NyBultgrupp.Insert();
                        if (Plslut)
                        {
                            NyBultgrupp.RemoveOtherPartToBolt(objekt5);
                        }
                        if (tilltek_typ[i] == "H1" || tilltek_typ[i] == "H2")
                        {
                            NyBultgrupp.Position.Rotation = Position.RotationEnum.TOP;
                        }
                        NyBultgrupp.PartToBeBolted = objekt1;
                        if (Plstart)
                        {
                            //NyBultgrupp.AddOtherPartToBolt(objekt4);
                            NyBultgrupp.PartToBoltTo = objekt4;
                        }
                        else
                        {
                            NyBultgrupp.PartToBoltTo = objekt2;
                        }
                        NyBultgrupp.FirstPosition = new T3D.Point(0, 0, 0);
                        NyBultgrupp.SecondPosition = new T3D.Point(Lst, 0, 0);
                        NyBultgrupp.Insert();
                    }
                }

                //myModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(aa);
                myModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());
                myModel.CommitChanges();
            }
        }

        // Normaliserar vektor 
        // Indata
        //      v:      Punkt 1, 2 eller 3 dimensoner, typ = array
        //      typ:    Anger om vektor ar 3 dim (typ = 3) eller 2 dim (typ != 3)
        private double[] Normvek(double[] v, int typ = 3)
        {
            if (typ == 3)
            {
                double Lv = Math.Sqrt(Math.Pow(v[0], 2) + Math.Pow(v[1], 2) + Math.Pow(v[2], 2));
                v[0] = v[0] / Lv;
                v[1] = v[1] / Lv;
                v[2] = v[2] / Lv;
            }
            else
            {
                double Lv = Math.Sqrt(Math.Pow(v[0], 2) + Math.Pow(v[1], 2));
                v[0] = v[0] / Lv;
                v[1] = v[1] / Lv;
            }
            return v;
        }

        // Skalärprodukten v1*v2
        // Indata
        //      v1:     Vektor 1, 3 dimensoner, typ = array
        //      v2:     Vektor 2, 3 dimensoner, typ = array
        private double Skalarprod(double[] v1, double[] v2)
        { return v1[0] * v2[0] + v1[1] * v2[1] + v1[2] * v2[2]; }

        // Tar fram normalvektorn (normaliserad) till vektor V1 i planet V1-V2 
        // Fungerar enbart med vektorer med längd 3
        // Indata
        //      v1:     Vektor 1, 3 dimensoner, typ = array
        //      v2:     Vektor 2, 3 dimensoner, typ = array
        private double[] Normvektor(double[] v1, double[] v2)
        {
            double[] v3 = new double[3];
            double[] n = new double[3];
            v3[0] = v1[1] * v2[2] - v1[2] * v2[1];
            v3[1] = v1[2] * v2[0] - v1[0] * v2[2];
            v3[2] = v1[0] * v2[1] - v1[1] * v2[0];
            n[0] = v1[1] * v3[2] - v1[2] * v3[1];
            n[1] = v1[2] * v3[0] - v1[0] * v3[2];
            n[2] = v1[0] * v3[1] - v1[1] * v3[0];
            double Lv = Math.Sqrt(Math.Pow(n[0], 2)+ Math.Pow(n[1], 2)+ Math.Pow(n[2], 2));
            n[0] = n[0] / Lv;
            n[1] = n[1] / Lv;
            n[2] = n[2] / Lv;
            return n;
        }

        // Kontrollerar om linje 1 (p,r) oc line 2 (q,s) korsar varandra i planet rxs
        // Indata
        //      p:      Startkoordnat, linje 1, 3 dimensoner, typ = array
        //      q:      Startkoordnat, linje 2, 3 dimensoner, typ = array
        //      r:      Vektor, linje 1, 3 dimensoner, längd motsvarar längd hos linje 1, typ = array
        //      s:      Vektor, linje 2, 3 dimensoner, längd motsvarar längd hos linje 2, typ = array    
        private bool Vektkors2(double[] p, double[] q, double[] r, double[] s)
        {
            bool ut = false;
            double[] N = new double[3];
            N[0] = s[1] * r[2] - s[2] * r[1];
            N[1] = s[2] * r[0] - s[0] * r[2];
            N[2] = s[0] * r[1] - s[1] * r[0];
            double Ln = Math.Sqrt(Math.Pow(N[0], 2) + Math.Pow(N[1], 2) + Math.Pow(N[2], 2));
            double Lr = Math.Sqrt(Math.Pow(r[0], 2) + Math.Pow(r[1], 2) + Math.Pow(r[2], 2));
            double Ls = Math.Sqrt(Math.Pow(s[0], 2) + Math.Pow(s[1], 2) + Math.Pow(s[2], 2));
            N[0] = N[0] / Ln;
            N[1] = N[1] / Ln;
            N[2] = N[2] / Ln;

            var A = Matrix<double>.Build.DenseOfArray(new double[,] {
                { r[0], N[0], -s[0] },
                { r[1], N[1], -s[1] },
                { r[2], N[2], -s[2] }
            });
            var B = Vector<double>.Build.Dense(new double[] { q[0] - p[0], q[1] - p[1], q[2] - p[2] });
            var x = A.Solve(B);
            if ((x[0]>= 0 && x[0] <= Lr) && (x[2] >= 0 && x[2] <= Ls)) { ut = true; }
            return ut;
        }

        // Identifierar nod där två vektorer möts
        // Indata
        //      p:      Startkoordnat, linje 1, 3 dimensoner, typ = array
        //      q:      Startkoordnat, linje 2, 3 dimensoner, typ = array
        //      r:      Vektor, linje 1, 3 dimensoner, längd motsvarar längd hos linje 1, typ = array
        //      s:      Vektor, linje 2, 3 dimensoner, längd motsvarar längd hos linje 2, typ = array   
        private double[] Vektkors(double[] p, double[] q, double[] r, double[] s)
        {
            double[] ut = new double[3];
            double[] N = new double[3];
            N[0] = s[1] * r[2] - s[2] * r[1];
            N[1] = s[2] * r[0] - s[0] * r[2];
            N[2] = s[0] * r[1] - s[1] * r[0];
            double Ln = Math.Sqrt(Math.Pow(N[0], 2) + Math.Pow(N[1], 2) + Math.Pow(N[2], 2));
            double Lr = Math.Sqrt(Math.Pow(r[0], 2) + Math.Pow(r[1], 2) + Math.Pow(r[2], 2));
            double Ls = Math.Sqrt(Math.Pow(s[0], 2) + Math.Pow(s[1], 2) + Math.Pow(s[2], 2));
            N[0] = N[0] / Ln;
            N[1] = N[1] / Ln;
            N[2] = N[2] / Ln;

            var A = Matrix<double>.Build.DenseOfArray(new double[,] {
                { r[0], N[0], -s[0] },
                { r[1], N[1], -s[1] },
                { r[2], N[2], -s[2] }
            });
            var B = Vector<double>.Build.Dense(new double[] { q[0] - p[0], q[1] - p[1], q[2] - p[2] });
            var x = A.Solve(B);
            ut = x.AsArray();
            return ut;
        }

        // Kontrollerar avstånd mellan stänger oc skruvar i en knutunkt med
        // Notera att v1 och v2 måste defineras så att de båda har kp som
        // startpunkt
        // Indata
        //      kp:     Koordinat knutpunkt, 3 dimensoner [m], typ = array
        //      vr:     Vektor, ramstång, 3 dimensoner [m], typ = array
        //      v1:     Vektor, diagonal 1, 3 dimensoner [m], typ = array
        //      v2:     Vektor, diagonal 2, 3 dimensoner [m], typ = array
        //      L:      Avstånd kant diagonal till knutpunkt [m], 2 dimensoner [m], typ = array
        //      B:      Bredd diagonal 1-3 [m], 2 dimensoner [m], typ = array
        //      Dmax:   Max. ytterdiam. bult diagonal 1-3 [m], 2 dimensoner [m], typ = array
        //      e1:     e1 [m] diagonal 1-3, 2 dimensoner [m], typ = array
        //      p1:     p1 [m] diagonal 1-3, 2 dimensoner [m], typ = array
        //      n:      Antal bultar diagonal 1-3, 2 dimensoner [m], typ = array
        //      sida:   Parameter som anger vilken sida om ramstångsflänsen diagonalen befinner sig 
        //              (0 = yttersida, 1 = innersida), 2 dimensoner [m], typ = array
        //      gemen:  Parameter som anger om två diagoanler får dela skruv (0 = nej, 1 = ja)
        //      Lf:     Avstånd från knutpunkt till insida av fläns ramstång (inkluderar radie)
        //      Lk:     Avstånd från knutpunkt till kant ramstång
        // Utdata:      int, 1 om ok med skruvar inom ramstångsfläns, 2 om ok med vissa skruvar utanför ramstångsfläns, 3 om ej ok
        private int Diagavstand(double[] kp, double[] vr, double[] v1, double[] v2, double[] v3, double[] L, double[] B,
           double[] Dmax, double[] e1, double[] p1, double[] n, double[] sida, double gemen, double Lf, double Lk)
        {
            int ut = 1;
            // Normaliserar alla vektorer
            vr = Normvek(vr);
            v1 = Normvek(v1);
            v2 = Normvek(v2);

            // Tra fram normalvektorer till samtliga vektorer. Notera att två normalvektorer beräknas för vr 
            double[] nr1 = Normvektor(vr, v1);
            double[] nr2 = Normvektor(vr, v2);
            double[] n1 = Normvektor(v1, vr);
            double[] n2 = Normvektor(v2, vr);

            // Kontrollerar avstånd mellan skruvar
            double p1max = p1.Max();
            double[] skr1 = new double[3];
            double[] skr2 = new double[3];
            double Lskr;
            for (int i = 0; i < n[0]; i++)
            {
                skr1[0] = 0 + v1[0] * (L[0] + e1[0] + p1[0] * i);
                skr1[1] = 0 + v1[1] * (L[0] + e1[0] + p1[0] * i);
                skr1[2] = 0 + v1[2] * (L[0] + e1[0] + p1[0] * i);
                for (int j = 0; j < n[1]; i++)
                {
                    skr2[0] = 0 + v2[0] * (L[1] + e1[1] + p1[1] * j);
                    skr2[1] = 0 + v2[1] * (L[1] + e1[1] + p1[1] * j);
                    skr2[2] = 0 + v2[2] * (L[1] + e1[1] + p1[1] * j);
                    Lskr = Math.Sqrt(Math.Pow(skr1[0] - skr2[0], 2) + Math.Pow(skr1[1] - skr2[1], 2) + Math.Pow(skr1[2] - skr2[2], 2));
                    if ((Lskr < p1max) && !(gemen == 1 & Lskr <= 0.00001))
                    {
                        ut = 3;
                        break;
                    }
                }
                if (ut == 3) { break; }
            }

            // Kontrollerar avstånd till insida fläns ramstång
            if (ut != 3)
            { if (((L[0]+e1[0])* Skalarprod(v1, nr1) > Lf - Dmax[0]) || ((L[1] + e1[1]) * Skalarprod(v2, nr2) > Lf - Dmax[1])) { ut = 3; } }

            // Kontrollerar avstånd från innersta och yttersta bult till kant ramstång
            if (ut != 3)
            {
                if (((L[0] + e1[0] + (n[0] - 1) * p1[0]) * Skalarprod(v1, nr1) >= e1[0] - Lk) || ((L[1] + e1[1] + (n[1] - 1) * p1[1]) * Skalarprod(v2, nr2) >= e1[1] - Lk))
                { ut = 1; }
                else if (((L[0] + e1[0]) * Skalarprod(v1, nr1) <= -Lk - Dmax[0]) || ((L[1] + e1[1]) * Skalarprod(v2, nr2) <= -Lk - Dmax[1]))
                { ut = 2; }
                else
                { ut = 3; }
            }

            // Kontrollerar avstånd mellan stänger samt mellan stänger och insida ramstång
            if (ut != 3)
            {
                double[] v1v = Normvek(vr);
                double[] v2v = Normvek(vr);
                double[] v1p1 = Normvek(vr);
                double[] v1p2 = Normvek(vr);
                double[] v2p1 = Normvek(vr);
                double[] v2p2 = Normvek(vr);
                v1p1[0] = v1[0] * (L[0] + e1[0]) + n1[0] * B[0] * 0.5;
                v1p1[1] = v1[1] * (L[0] + e1[0]) + n1[1] * B[0] * 0.5;
                v1p1[2] = v1[2] * (L[0] + e1[0]) + n1[2] * B[0] * 0.5;
                v1p2[0] = v1[0] * (L[0] + e1[0]) - n1[0] * B[0] * 0.5;
                v1p2[1] = v1[1] * (L[0] + e1[0]) - n1[1] * B[0] * 0.5;
                v1p2[2] = v1[2] * (L[0] + e1[0]) - n1[2] * B[0] * 0.5;
                v2p1[0] = v2[0] * (L[2] + e1[2]) + n2[0] * B[2] * 0.5;
                v2p1[1] = v2[1] * (L[2] + e1[2]) + n2[1] * B[2] * 0.5;
                v2p1[2] = v2[2] * (L[2] + e1[2]) + n2[2] * B[2] * 0.5;
                v2p2[0] = v2[0] * (L[2] + e1[2]) - n2[0] * B[2] * 0.5;
                v2p2[1] = v2[1] * (L[2] + e1[2]) - n2[1] * B[2] * 0.5;
                v2p2[2] = v2[2] * (L[2] + e1[2]) - n2[2] * B[2] * 0.5;
                v1v[0] = v1[0] * n[0] * 10 * p1[0];
                v1v[1] = v1[1] * n[0] * 10 * p1[0];
                v1v[2] = v1[2] * n[0] * 10 * p1[0];
                v2v[0] = v2[0] * n[1] * 10 * p1[1];
                v2v[1] = v2[1] * n[1] * 10 * p1[1];
                v2v[2] = v2[2] * n[1] * 10 * p1[1];
                bool kors1 = Vektkors2(v1p1, v2p1, v1v, v2v);
                bool kors2 = Vektkors2(v1p1, v2p2, v1v, v2v);
                bool kors3 = Vektkors2(v1p2, v2p1, v1v, v2v);
                bool kors4 = Vektkors2(v1p2, v2p2, v1v, v2v);
                bool kors5 = false;
                if ((Skalarprod(v1p1, nr1) > Lf) || (Skalarprod(v1p2, nr1) > Lf) || (Skalarprod(v2p1, nr2) > Lf) || (Skalarprod(v2p2, nr2) > Lf)) { kors5 = true; }
                if (kors1 || kors2 || kors3 || kors4 || kors5) { ut = 3; }
            }
            return ut; 
        }

    }

    // ANVÄNDS EJ!!!
    public class Knutpnkt
    {
        public double X1 { get; set; }          // x-koord, knutpunkt, start på ramstången
        public double Y1 { get; set; }          // z-koord, knutpunkt, start på ramstången 
        public double X2 { get; set; }          // x-koord, knutpunkt, slut på ramstången
        public double Y2 { get; set; }          // z-koord, knutpunkt, slut på ramstången
        public double Test1 { get; set; }
        public double Test2 { get; set; }

        //objekt1.Profile.ProfileString = "L65x7";
        //objekt1.Modify();
        //MessageBox.Show(string.Format("Hej! Detta elements GUID är: {0}", objekt1.Identifier.GUID.ToString()));
        //MessageBox.Show(string.Format("Bult: {0}", tilltek_blt[i]));
        //MessageBox.Show(string.Format("Antal bult: {0}", tilltek_bltN[i]));
        //MessageBox.Show(string.Format("e1: {0}", tilltek_blte1[i]));
        //MessageBox.Show(string.Format("p1 {0}", tilltek_bltp1[i]));
        //MessageBox.Show(string.Format("d0: {0}", tilltek_bltd0[i]));
        //myModel.SelectModelObject(myBeam.Identifier);
        //Tekla.Structures.Model.Part ValdDelEtt = null;
        //Tekla.Structures.Model.Part ValdDelTva = null;
        //tilltek_idnr[i];
        //tilltek_startnod[i];
        //tilltek_slutnod[i];




    }

}
