using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace pishnahadBadane
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }


        public Form1.Bime[] bm;
        private List<int> selectedbm = new List<int>();
        public string[,] pishnahad;
        DataTable dt = new DataTable();
        DataTable DataTable = new DataTable();
        DataSet dataSet = new DataSet();
        private void Form3_Load(object sender, EventArgs e)
        {
            pishnahad = new string[bm.Length, 36];
            dt.Columns.Add("id");
            dt.Columns.Add("شماره بیمه");
            dt.Columns.Add("بیمه گذار");
            dt.Columns.Add("نام خودرو");
            dt.Columns.Add("select").DataType = typeof(bool);


            //classToArry();
            clastoGv();

            gv.DataSource = dt;
            gv.Columns[0].Visible = false;
            gv.RightToLeft = RightToLeft.Yes;
            gv.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            gv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            gv.MultiSelect = false;
            gv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            gv.AllowUserToAddRows = false;
            gv.AllowUserToDeleteRows = false;
            gv.AllowUserToOrderColumns = false;
            gv.AllowUserToResizeRows = false;
            gv.RowHeadersVisible = false;
            gv.Columns[0].ReadOnly = true;
            gv.Columns[1].ReadOnly = true;
            gv.Columns[2].ReadOnly = true;
            gv.Columns[3].ReadOnly = true;
            gv.Columns[4].ReadOnly = false;
            foreach (DataGridViewColumn column in gv.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            DataGridViewImageColumn imageCol = new DataGridViewImageColumn();
            imageCol.HeaderText = "printed";
            imageCol.ImageLayout = DataGridViewImageCellLayout.Stretch;
            imageCol.Image = Image.FromFile("c:\\users\\mega\\documents\\visual studio 2015\\Projects\\pishnahadBadane\\pishnahadBadane\\cancel.png"); ;
            gv.Columns.Add(imageCol);
            gv.Columns[5].Visible = false;


        }

        private void clastoGv()
        {

            for (int i = 0; i < bm.Length; i++)
            {
                dt.Rows.Add(i, bm[i].num, bm[i].person.name, bm[i].car.name, false);

            }
        }


        private void loadpishnahad(int i)
        {

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
        private void gv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
                loadpishnahad(Int32.Parse(gv.Rows[e.RowIndex].Cells[0].Value.ToString())); ;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            countSelected();
            printPishnahad(selectedbm);
        }


        string[] bmToPdfData(int i)
        {
            string[] pdfDataStrings = new string[36];
            pdfDataStrings[0] = bm[i].num + " / " + bm[i].year;
            pdfDataStrings[1] = bm[i].person.name;
            pdfDataStrings[2] = bm[i].person.codeMeli;
            pdfDataStrings[3] = bm[i].car.type;
            pdfDataStrings[4] = bm[i].car.shasi;
            pdfDataStrings[5] = bm[i].car.motor;
            pdfDataStrings[6] = bm[i].car.rang;
            pdfDataStrings[7] = bm[i].car.saleSakht;
            pdfDataStrings[8] = bm[i].car.pelak.doRagham + " " + bm[i].car.pelak.harf + " " + bm[i].car.pelak.seRagham;
            pdfDataStrings[9] = bm[i].car.pelak.iran + " ایران ";
            pdfDataStrings[10] = bm[i].car.name;

            pdfDataStrings[11] = bm[i].lastBimeEndDate;
            if (bm[i].lastCompany == "آسيا(بيمه گر)")
            {
                pdfDataStrings[12] = bm[i].lastBimeId;
            }
            else
            {
                pdfDataStrings[12] = bm[i].lastBimeNum;
            }
            pdfDataStrings[13] = bm[i].lastCompany;
            pdfDataStrings[14] = bm[i].poosheshHa.feranshiz ? "■" : " ";
            pdfDataStrings[15] = bm[i].poosheshHa.navasanat ? "■" : " ";
            pdfDataStrings[16] = bm[i].poosheshHa.avamelTabiyi ? "■" : " ";
            pdfDataStrings[17] = bm[i].poosheshHa.asidPashi ? "■" : " ";
            pdfDataStrings[18] = bm[i].poosheshHa.shishe ? "■" : " ";
            pdfDataStrings[19] = bm[i].poosheshHa.estelak ? "■" : " ";
            pdfDataStrings[20] = "■";
            pdfDataStrings[21] = bm[i].poosheshHa.serghatGHataat ? "■" : " ";
            pdfDataStrings[22] = bm[i].poosheshHa.ayabzahab ? "■" : " ";
            pdfDataStrings[23] = bm[i].poosheshHa.havadesShakhsi ? "■" : " ";

            pdfDataStrings[24] = string.Format("{0:n0}", Int64.Parse(bm[i].mablagh));
            pdfDataStrings[25] = string.Format("{0:n0}", Int64.Parse(bm[i].car.arzeshYadak));
            pdfDataStrings[26] = "ارزش یدک";
            pdfDataStrings[30] = string.Format("{0:n0}", Int64.Parse(bm[i].car.arzeshLavazem));
            pdfDataStrings[31] = "ارزش لوازم اضافی";

            pdfDataStrings[27] = bm[i].person.name;

            Int64 s = Int64.Parse(bm[i].car.arzesh);
            s += Int64.Parse(bm[i].car.arzeshLavazem);
            s += Int64.Parse(bm[i].car.arzeshYadak);
            pdfDataStrings[28] = string.Format("{0:n0}", s);

            pdfDataStrings[29] = bm[i].startDate;

            pdfDataStrings[32] = " کد: " + bm[i].person.id;
            pdfDataStrings[33] = bm[i].person.address;
            pdfDataStrings[34] = bm[i].person.tel;
            pdfDataStrings[35] = bm[i].person.phone;
            return pdfDataStrings;
        }


        private void printPishnahad(List<int> v)
        {

            PdfDocument doc = new PdfDocument();
            doc.PageSettings.SetMargins(20);
            PdfFont font1 = new PdfTrueTypeFont(new Font("Dast Nevis2", 12, FontStyle.Regular), true);
            PdfFont font2 = new PdfTrueTypeFont(new Font("Lucida Handwriting", 10, FontStyle.Regular), true);
            PdfFont font3 = new PdfTrueTypeFont(new Font("Arial", 13, FontStyle.Regular), true);

            /////////
            PdfStringFormat format = new PdfStringFormat();
            format.LineSpacing = 1;
            format.TextDirection = PdfTextDirection.RightToLeft;
            format.Alignment = PdfTextAlignment.Right;
            format.LineAlignment = PdfVerticalAlignment.Middle;

            PdfStringFormat format2 = new PdfStringFormat();
            format2.LineSpacing = 1;
            format2.TextDirection = PdfTextDirection.RightToLeft;
            format2.Alignment = PdfTextAlignment.Left;
            format2.LineAlignment = PdfVerticalAlignment.Middle;

            PdfStringFormat format3 = new PdfStringFormat();
            format3.LineSpacing = 0;
            format3.TextDirection = PdfTextDirection.RightToLeft;
            format3.Alignment = PdfTextAlignment.Center;
            format3.LineAlignment = PdfVerticalAlignment.Middle;
            format3.WordWrap = PdfWordWrapType.Character;


            PdfPen pen = new PdfPen(Color.Black, 1);


            dataSet.ReadXml("data.xml");
            DataTable = dataSet.Tables[0];
            string[] pdfDataStrings = new string[36];

            for (int n = 0; n < v.Count; n++)
            {
                pdfDataStrings = bmToPdfData(v[n]);

                PdfPage page = doc.Pages.Add();
                PdfGraphics graphics = page.Graphics;
                ////////


                for (int i = 0; i < pdfDataStrings.Length; i++)
                {
                    float x, y, h, w;
                    x = float.Parse(DataTable.Rows[i][1].ToString());
                    y = float.Parse(DataTable.Rows[i][2].ToString());
                    h = float.Parse(DataTable.Rows[i][3].ToString());
                    w = float.Parse(DataTable.Rows[i][4].ToString());
                    //graphics.DrawRectangle(pen, x, y, w, h);
                    if (i == 4 || i == 5)
                    {
                        graphics.DrawString(pdfDataStrings[i], font2, PdfBrushes.Black,
                            new RectangleF(x, y, w, h), format3);
                    }
                    else if (i >= 14 && i <= 23)
                    {
                        graphics.DrawString(pdfDataStrings[i], font3, PdfBrushes.Black,
                            new RectangleF(x, y, w, h), format3);
                    }
                    else
                    {
                        graphics.DrawString(pdfDataStrings[i], font1, PdfBrushes.Black,
                            new RectangleF(x, y, w, h), format3);
                    }

                }
            }

            string pdfFilename = "firstpage.pdf";
            doc.Save(pdfFilename);
            Process.Start(pdfFilename);
        }

        private void makePdf()
        {
            PdfDocument doc = new PdfDocument();
            doc.PageSettings.SetMargins(20);
            PdfPage page = doc.Pages.Add();
            PdfGraphics graphics = page.Graphics;
            PdfPen pen = new PdfPen(Color.DarkGray, 1);
            for (int i = 0; i < 100; i++)
            {
                graphics.DrawLine(pen, 10 * i, 0, 10 * i, 1500);
            }
            for (int j = 0; j < 300; j++)
            {
                graphics.DrawLine(pen, 0, 10 * j, 1500, 10 * j);
            }
            string pdfFilename = "firstpage.pdf";
            doc.Save(pdfFilename);
            Process.Start(pdfFilename);
        }

        private void countSelected()
        {
            int t;
            for (int i = 0; i < gv.Rows.Count; i++)
            {
                if (bool.Parse(gv.Rows[i].Cells[4].Value.ToString()))
                {
                    t = int.Parse(gv.Rows[i].Cells[0].Value.ToString());
                    selectedbm.Add(t);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            makePdf();
        }

        private void gv_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex == 4 && e.RowIndex == -1)
            {
                for (int i = 0; i < gv.Rows.Count; i++)
                {
                    gv.Rows[i].Cells[4].Value = !(bool)gv.Rows[i].Cells[4].Value;

                }

            }
        }

    }
}
