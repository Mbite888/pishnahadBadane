using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace pishnahadBadane
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        public Form1.Bime[] bm;
        private int page, pages;

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void Form2_Load(object sender, EventArgs e)
        {
            loadpishnahad(0);
            page = 0;
            pages = bm.Length;
            
            pagesLb.Text = pages.ToString();
        }

        private void loadpishnahad(int i)
        {
            pageLb.Text = (page+1).ToString();

            //////////person 
            pname.Text = bm[i].person.name;
            pid.Text = bm[i].person.id;

            ////// poosheshha
            asidPashi.Checked = bm[i].poosheshHa.asidPashi;
            avamelTabiyi.Checked = bm[i].poosheshHa.avamelTabiyi;
            ayabzahab.Checked = bm[i].poosheshHa.ayabzahab;
            feranshiz.Checked = bm[i].poosheshHa.feranshiz;
            estelak.Checked = bm[i].poosheshHa.estelak;
            serghatGHataat.Checked = bm[i].poosheshHa.serghatGHataat;
            havadesShakhsi.Checked = bm[i].poosheshHa.havadesShakhsi;
            navasanat.Checked = bm[i].poosheshHa.navasanat;
            shishe.Checked = bm[i].poosheshHa.shishe;

            //////// car
            cname.Text = bm[i].car.name;
            carzesh.Text = bm[i].car.arzesh;
            carzeshYadak.Text = bm[i].car.arzeshYadak;
            cmotor.Text = bm[i].car.motor;
            cshasi.Text = bm[i].car.shasi;
            ctype.Text = bm[i].car.type;
            carzeshLavazem.Text = bm[i].car.arzeshLavazem;
            crang.Text = bm[i].car.rang;
            cuse.Text = bm[i].car.use;
            csaleSakht.Text = bm[i].car.saleSakht;

            //////////// pelak

            piran.Text = bm[i].car.pelak.iran;
            pdoRagham.Text = bm[i].car.pelak.doRagham;
            pseRagham.Text = bm[i].car.pelak.seRagham;
            pharf.Text = bm[i].car.pelak.harf;

            //////////////// bime 
            pishnahadNum.Text = bm[i].pishnahadNum;
            bimeNum.Text = bm[i].num;
            bimeId.Text = bm[i].id;
            lastCompany.Text = bm[i].lastCompany;
            lastBimeNum.Text = bm[i].lastBimeNum;
            lastBimeId.Text = bm[i].lastBimeId;
            startDate.Text = bm[i].startDate;
            endDate.Text = bm[i].endDate;
            lastBimeEndDate.Text = bm[i].lastBimeEndDate;
            mablagh.Text = bm[i].mablagh;
            salTakhfif.Text = bm[i].salTakhfif;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (page > 0)
            {
                page--;
                loadpishnahad(page);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            page= 0;
            loadpishnahad(page);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            page = pages - 1;
            loadpishnahad(page);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (page < pages-1)
            {
                page++;
                loadpishnahad(page);
            }
               
        }
    }
}
