using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Management.Instrumentation;
using System.Windows.Forms;
using Syncfusion.XPS;

namespace pishnahadBadane
{
    public partial class Form1 : Form
    {
        string excelBadaneHa = "C:\\Users\\Mega\\Desktop\\badane ha.xlsx";
        string excelPersons = "C:\\Users\\Mega\\Desktop\\badane Persons.xlsx";
        public Bime[] bm;
        public Person[] persons;


        public class Bime
        {
            public Car car;
            public Person person;
            public PoosheshHa poosheshHa;

            public string pishnahadNum,
                num,
                id,
                lastCompany,
                lastBimeNum,
                lastBimeId,
                startDate,
                endDate,
                lastBimeEndDate,
                mablagh,
                salTakhfif,
                year;

        }

        public class PoosheshHa
        {
            public bool navasanat,
                avamelTabiyi,
                asidPashi,
                ayabzahab,
                serghatGHataat,
                shishe,
                feranshiz,
                estelak,
                havadesShakhsi;
        }

        public class Person
        {
            public string name, id , phone , address , codeMeli , tel;
        }
        public class Car
        {
            public string shasi, motor, arzesh, arzeshYadak, arzeshLavazem, saleSakht, rang, type, name, use;
            public Pelak pelak;
        }

        public class Pelak
        {
            public string iran, seRagham, doRagham, harf;
        }


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "excel Files (*.xlsx)|*.xlsx";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                string sFileName = choofdlog.FileName;
                textBox1.Text = sFileName.Replace("//", "/");
                excelBadaneHa = sFileName;
            }
        }

        DataTable readBadaneHaexcel()
        {
            String name = "CarBdBNVer";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            excelBadaneHa +
                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            sda.Fill(data);
            return data;
        }
        DataTable readPersonsexcel()
        {
            String name = "NewDataSet";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            excelPersons +
                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            sda.Fill(data);
            return data;
        }



        Bime[] Table1toClass(DataTable data)
        {
            Bime[] bml = new Bime[data.Rows.Count];
            for (int i = 0; i < data.Rows.Count; i++)
            {
                bml[i] = new Bime();


                ////////////////////////person
                bml[i].person = new Person();
                bml[i].person.id = giveCodeFromName(data.Rows[i]["بيمه گذار"].ToString());
                bml[i].person.name = giveNameFromName(data.Rows[i]["بيمه گذار"].ToString());

                ////////////////////////poosheshHa
                bml[i].poosheshHa = new PoosheshHa();
                bml[i].poosheshHa.asidPashi = chkPooshesh(data.Rows[i]["حق بيمه پاشيدن رنگ"].ToString());
                bml[i].poosheshHa.avamelTabiyi = chkPooshesh(data.Rows[i]["حق بيمه سيل و زلزله"].ToString());
                bml[i].poosheshHa.ayabzahab = chkPooshesh(data.Rows[i]["حق بيمه اياب و ذهاب / هزینه توقف"].ToString());
                bml[i].poosheshHa.estelak = chkPooshesh(data.Rows[i]["حذف استهلاک"].ToString());
                bml[i].poosheshHa.feranshiz = chkPooshesh(data.Rows[i]["حذف فرانشیز"].ToString());
                bml[i].poosheshHa.havadesShakhsi = chkPooshesh(data.Rows[i]["حوادث شخصی جاری"].ToString());
                bml[i].poosheshHa.serghatGHataat = chkPooshesh(data.Rows[i]["حق بيمه سرقت درجا"].ToString());
                bml[i].poosheshHa.navasanat = chkPooshesh(data.Rows[i]["حق بیمه نوسانات قیمت"].ToString());
                bml[i].poosheshHa.shishe = chkPooshesh(data.Rows[i]["حق بیمه شكست شيشه به تنهايي"].ToString());


                ////////////////////////car
                bml[i].car = new Car();
                bml[i].car.name = data.Rows[i]["نوع وسيله نقليه"].ToString();
                bml[i].car.arzesh = data.Rows[i]["ارزش وسيله نقليه"].ToString();
                bml[i].car.arzeshYadak = data.Rows[i]["ارزش يدک"].ToString();
                bml[i].car.motor = data.Rows[i]["شماره موتور"].ToString();
                bml[i].car.shasi = data.Rows[i]["شماره شاسي"].ToString();
                bml[i].car.type = data.Rows[i]["مورد استفاده وسيله نقليه"].ToString();
                bml[i].car.arzeshLavazem = data.Rows[i]["ارزش لوازم اضافی"].ToString();
                bml[i].car.rang = data.Rows[i]["رنگ خودرو"].ToString();
                bml[i].car.use = data.Rows[i]["گروه تعرفه اي بدنه"].ToString();
                bml[i].car.saleSakht = data.Rows[i]["سال ساخت وسيله نقليه"].ToString();
                ////////////////////////pelak
                bml[i].car.pelak = new Pelak();
                bml[i].car.pelak = setPelak(data.Rows[i]["شماره پلاک"].ToString());
                bml[i].car.pelak.iran = data.Rows[i]["سريال پلاک"].ToString();

                //////////////////////////bime
                bml[i].pishnahadNum = data.Rows[i]["پيشنهاد"].ToString();
                bml[i].num = data.Rows[i]["شماره بيمه نامه"].ToString();
                bml[i].id = data.Rows[i]["کد رایانه بیمه نامه"].ToString();
                bml[i].lastCompany = data.Rows[i]["شرکت بیمه سال قبل"].ToString();
                bml[i].lastBimeNum = data.Rows[i]["شماره بیمه نامه سال قبل"].ToString();
                bml[i].lastBimeId = data.Rows[i]["بیمه نامه سال قبل"].ToString();
                bml[i].startDate = data.Rows[i]["تاریخ شروع"].ToString();
                bml[i].endDate = data.Rows[i]["تاریخ پایان"].ToString();
                bml[i].lastBimeEndDate = data.Rows[i]["تاريخ انقضاء بيمه نامه سال قبل"].ToString();
                bml[i].mablagh = data.Rows[i]["حق بیمه با مالیات"].ToString();
                bml[i].salTakhfif = data.Rows[i]["تعداد سال عدم خسارت"].ToString();
                bml[i].year = bml[i].startDate.Split('/')[0];



            }
            return bml;
        }

        private Pelak setPelak(string p)
        {
            Pelak pelak = new Pelak();

            if (p.Length > 4)
            {
                pelak.harf = p[3].ToString();
                pelak.seRagham = p[0].ToString() + p[1].ToString() + p[2].ToString();
                pelak.doRagham = p[4].ToString() + p[5].ToString();
            }
            else
            {
                pelak.harf = "";
                pelak.seRagham = "";
                pelak.doRagham = "";
            }
            return pelak;
        }

        public bool chkPooshesh(string p)
        {
            if (p == "0")
                return false;
            else
            {
                return true;
            }
        }


        public string giveNameFromName(string name)
        {
            name = name.Split(new string[] { " کد" }, StringSplitOptions.None)[0];
            return name;
        }

        public string giveCodeFromName(string name)
        {
            name = name.Split(new string[] { " کد" }, StringSplitOptions.None)[1];
            return name;
        }



        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "excel Files (*.xlsx)|*.xlsx";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                string sFileName = choofdlog.FileName;
                textBox2.Text = sFileName.Replace("//", "/");
                excelPersons = sFileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (excelBadaneHa != null && excelPersons != null )
            {

                DataTable data = readBadaneHaexcel();
                DataTable data2 = readPersonsexcel();
                
                /*var results = (from table1 in data.AsEnumerable()
                    join table2 in data2.AsEnumerable() on (double) table1["کد رایانه بیمه نامه"] equals (double) table2["کد رایانه بیمه نامه"]
                    select new {T1 = table1, T2 = table2 }).ToList();*/

                persons = personsToClass(data2);
                bm = Table1toClass(data);
                bm = joinPersonToBm(bm,persons);

                Form3 f3 = new Form3();
                f3.bm = bm;
                f3.Show();

            }
            else
            {
                MessageBox.Show("insert excel file", "bla bla bla",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Bime[] joinPersonToBm(Bime[] bm, Person[] persons)
        {
            for (int i = 0; i < bm.Length; i++)
            {
                for (int j = 0; j < persons.Length; j++)
                {
                    if (bm[i].person.id == persons[j].id)
                    {
                        bm[i].person.address = persons[j].address;
                        bm[i].person.codeMeli = persons[j].codeMeli;
                        bm[i].person.phone = persons[j].phone;
                        bm[i].person.tel = persons[j].tel;
                    }
                }
            }

            return bm;
        }

        private Person[] personsToClass(DataTable data2)
        {
            persons = new Person[data2.Rows.Count];
            for (int i = 0; i < data2.Rows.Count; i++)
            {
                persons[i] = new Person();
                persons[i].id = giveCodeFromName(data2.Rows[i]["نام بيمه گذار"].ToString());
                persons[i].name = giveNameFromName(data2.Rows[i]["نام بيمه گذار"].ToString());
                persons[i].codeMeli = data2.Rows[i]["کد / شناسه ملي"].ToString();
                persons[i].address = data2.Rows[i]["آدرس"].ToString();
                persons[i].phone = data2.Rows[i]["موبایل"].ToString();
                persons[i].tel = data2.Rows[i]["تلفن"].ToString();
            }

            return persons;
        }
    }
}
