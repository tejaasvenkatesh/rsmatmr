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
using Microsoft.Office;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Diagnostics;
using Novacode;
using ExcelLibrary.SpreadSheet;
using System.Data.OleDb;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        string firstGap = "";
        string secondGap = "    ";
        string gender;
        string lang;
        int checkLnth = 54;
        int template = 1;
      
        List<string> univlist = new List<string>()
        {
            "Acharya N G Ranga Agricultural University (ANGRAU), Hyderabad",
"Acharya Nagarjuna University, Guntur",
"Acharya Nagarjuna University - Center for Distance Education, Guntur",
"Adikavi Nannaya University (ANU), East Godavari",
"Andhra Pradesh University of Law, Visakhapatnam",
"Andhra University (AU), Visakhapatnam",
"Andhra University School of Distance Education (AU SDE), Visakhapatnam",
"Dr B R Ambedkar Open University (BROU), Hyderabad",
"Dr BR Ambedkar University, Srikakulam",
"Dr Y S R Horticultural University, West Godavari",
"Dravidian University, Chitoor",
"Dravidian University Directorate of Distance Education, Chitoor",
"English and Foreign Languages University (EFLU), Hyderabad",
"English and Foreign Languages University - School of Distance Education (EFLU SDE)",
"ICFAI Foundation for Higher Education, Hyderabad",
"International Institute of Information Technology (IIIT), Hyderabad",
"GITAM University, Visakhapatnam",
"GITAM Centre for Distance Learning, Visakhapatnam",
"Jawaharlal Nehru Technological University Anantapur (JNTU Anantapur)",
"Jawaharlal Nehru Technological University Kakinada (JNTU Kakinada), East Godavari",
"Jawaharlal Nehru Technological University Hyderabad (JNTU Hyderabad)",
"JNTU - School of Continuing & Distance Education, Hyderabad",
"Jawaharlal Nehru Architecture and Fine Arts University, Hyderabad",
"Kakatiya University (KU), Warangal",
"Kakatiya University - School of Distance Learning & Continuing Education (KU-SDLC), Warangal",
"Koneru Lakshmaiah University, Vijayawada",
"Krishna University, Krishna",
"Mahatma Gandhi University (MGU), Nalgonda",
"Mahatma Gandhi University Centre for Distance Education (MGUCDE)",
"Maulana Azad National Urdu University",
"Maulana Azad National Urdu University - Directorate of Distance Education",
"NALSAR University of Law, Ranga Reddy",
"National Institute of Technology (NIT) Warangal",
"Nizam's Institute of Medical Sciences, Hyderabad",
"NTR University of Health Sciences, Vijayawada",
"Osmania University (OU), Hyderabad",
"Osmania University - PGRR Center for Distance Education, Hyderabad",
"Palamuru University, Mahabubnagar",
"Potti Sreeramulu Telugu University (PSTU), Hyderabad",
"Rashtriya Sanskrit Vidyapeeth, Tirupathi",
"Rashtriya Sanskrit Vidyapeeth - Directorate of Distance Education, Tirupathi",
"Rayalaseema University, Kurnool",
"Rayalaseema University Directorate of Distance Education, Kurnool",
"Rajiv Gandhi University of Knowledge Technologies (RGUKS), Hyderabad",
"Satavahana University, Karimnagar",
"Sri Krishnadevaraya University (SKU) Anantapur",
"Sri Krishnadevaraya University Centre for Distance Education (SKUCDE), Anantapur",
"Sri Padmavathi Mahila Viswa Vidyalayam (SPMVV), Tirupati",
"Sri Padmavati Mahila Visvavidyalayam Distance Education Centre, Tirupati",
"Sri Venkateswara University (SVU), Tirupathi",
"Sri Venkateswara University Directorate of Distance Education (SVUDDE), Tirupati",
"Sri Venkateswara Institute of Medical Sciences (SVIMS), Tirupathi",
"Sri Venkateswara Vedic University, Tirupathi",
"Sri Venkateswara Veterinary University (SVVU), Tirupati",
"Sri Sathya Sai University, Prasanthinilayam, Anantapur",
"Telangana University, Nizamabad",
"Vignan University, Guntur",
"Vikram Simhapuri University, Nellore",
"University of Hyderabad (UoH), Hyderabad",
"University of Hyderabad Center for Distance Education (UoH CDE), Hyderabad",
"Yogi Vemana University, Cuddapah"

        };

        public Form1()
        {
            InitializeComponent();
            foreach(string s in univlist)
            {
                comboBox1.Items.Add(s);
            }  
           
        }



        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Add("JNTUK");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!Validate(textBox6.Text)) return;
            if (!Validate(textBox29.Text)) return;

            string path = SaveFile();

            //string fName = "Resume.docx";
            string fileName = @Path.Combine(path,"Resume.docx");

            string objective = "OBJECTIVE:";
            string eduQ = "EDUCATIONAL QUALIFICATION:";
            string PD = "PERSONAL DETAILS:";
            string title = "RESUME";
            if (template == 1)
            {
                title = "RESUME";
            }
            if (template == 2)
            {
                title = textBox1.Text + " " + textBox2.Text;
            }

            string TS = "TECHNICAL SKILLS:";
            string NL = "\n";
            string DC = "DECLARATION:";

            string para0 = "S.No";
            string para1 = "Qualification";
            string para2 = "Institute";
            string para3 = "University/Board";
            string para4 = "Year of Pass";
            string para5 = "Percentage/GPA";

            string para6 = "1";
            string para7 = "B.Tech";
            string para8 = textBox10.Text;
            string para9 = comboBox1.Text;
            string para10 = textBox12.Text;
            string para11 = textBox13.Text;

            string para12 = "2";
            string para13 = "Intermediate";
            string para14 = textBox17.Text;
            string para15 = textBox16.Text;
            string para16 = textBox15.Text;
            string para17 = textBox14.Text;

            string para18 = "3";
            string para19 = "SSC";
            string para20 = textBox21.Text;
            string para21 = textBox20.Text;
            string para22 = textBox19.Text;
            string para23 = textBox18.Text;



            if (checkBox1.Checked) lang += "Telugu ";
            if (checkBox2.Checked) lang += "Hindi ";
            if (checkBox3.Checked) lang += "Tamil ";
            if (checkBox4.Checked) lang += "English ";

            string place = "Place: " + textBox22.Text;
            string date = "Date: " + DateTime.Now.ToString("dd/MM/yyyy");

            string line0 = "    I hereby declare that all the details furnished above are true to the best of my knowledge.";

            string line1 = "      " + richTextBox2.Text ;

                              
            
            string line2= "Name\t\t\t\t" + firstGap + ":" + secondGap + textBox1.Text + " " + textBox2.Text ;
            string line3= "Father's Name\t\t\t" + firstGap + ":" + secondGap + textBox4.Text ;
            string line4= "Mother's Name\t\t" + firstGap + ":" + secondGap + textBox5.Text ;
            string line5= "Gender\t\t\t" + firstGap + ":" + secondGap + gender ;
            string line6= "Date of Birth\t\t\t" + firstGap + ":" + secondGap + textBox3.Text ;
            string line7= "Programming Languages\t" + firstGap + ":" + secondGap + textBox6.Text ;
            string line8= "Languages Known\t\t" + firstGap + ":" + secondGap + lang ;
            string line9= "Phone Number\t\t" + firstGap + ":" + secondGap + textBox28.Text ;
            string line10= "Address\t\t\t" + firstGap + ":" + secondGap + textBox25.Text + "," + textBox24.Text + textBox23.Text + "," ;
            string line101 = "\t\t\t\t" + firstGap + " " + secondGap + textBox22.Text + "-" + textBox26.Text;
            string line16= "Packages Known\t\t" + firstGap + ":" + secondGap + textBox29.Text ;
            string line17= "Name\t\t\t" + firstGap + ":" + secondGap + textBox1.Text ;

            var headformat = new Formatting();
            headformat.Size = 20D;
            headformat.Bold = true;
            
            var objformat = new Formatting();
            objformat.Size = 16D;
            objformat.Bold = true;
            objformat.FontFamily = new System.Drawing.FontFamily("Calibri");

            var lineFormat = new Formatting();
            lineFormat.Size = 12D;
            lineFormat.FontFamily = new System.Drawing.FontFamily("Calibri");

            var endformat = new Formatting();
            endformat.Size = 12D;
            endformat.Bold = true;
            endformat.FontFamily = new System.Drawing.FontFamily("Calibri");

            var doc = DocX.Create(fileName);
            doc.InsertParagraph(title, false, headformat).Alignment = Alignment.center;
            doc.InsertParagraph(NL, false, lineFormat).LineSpacingAfter = 10;

            doc.InsertParagraph(objective, false, objformat).LineSpacingAfter = 10;
            doc.InsertParagraph(line1, false, lineFormat).LineSpacingAfter = 10;
            doc.InsertParagraph(NL, false, lineFormat);
            //Educational Qualification
            doc.InsertParagraph(eduQ, false, objformat).LineSpacingAfter = 10;
            Table t = doc.AddTable(4,6);
            
            t.Rows[0].Cells[0].InsertParagraph(para0);
            t.Rows[0].Cells[1].InsertParagraph(para1);
            t.Rows[0].Cells[2].InsertParagraph(para2);
            t.Rows[0].Cells[3].InsertParagraph(para3);
            t.Rows[0].Cells[4].InsertParagraph(para4);
            t.Rows[0].Cells[5].InsertParagraph(para5);
            t.Rows[1].Cells[0].InsertParagraph(para6);
            t.Rows[1].Cells[1].InsertParagraph(para7);
            t.Rows[1].Cells[2].InsertParagraph(para8);
            t.Rows[1].Cells[3].InsertParagraph(para9);
            t.Rows[1].Cells[4].InsertParagraph(para10);
            t.Rows[1].Cells[5].InsertParagraph(para11);
            t.Rows[2].Cells[0].InsertParagraph(para12);
            t.Rows[2].Cells[1].InsertParagraph(para13);
            t.Rows[2].Cells[2].InsertParagraph(para14);
            t.Rows[2].Cells[3].InsertParagraph(para15);
            t.Rows[2].Cells[4].InsertParagraph(para16);
            t.Rows[2].Cells[5].InsertParagraph(para17);
            t.Rows[3].Cells[0].InsertParagraph(para18);
            t.Rows[3].Cells[1].InsertParagraph(para19);
            t.Rows[3].Cells[2].InsertParagraph(para20);
            t.Rows[3].Cells[3].InsertParagraph(para21);
            t.Rows[3].Cells[4].InsertParagraph(para22);
            t.Rows[3].Cells[5].InsertParagraph(para23);
            doc.InsertTable(t);

            //table place here
            doc.InsertParagraph(NL, false, lineFormat);
            //Technical Skills
            doc.InsertParagraph(TS, false, objformat).LineSpacingAfter = 10;
            doc.InsertParagraph(line7, false, lineFormat).LineSpacingAfter = 10;
            doc.InsertParagraph(line16, false, lineFormat).LineSpacingAfter = 10;
            if (textBox31.Text != "")
            {
                //split line
                if (textBox31.Text.Length > checkLnth)
                {

                    var lines = Splitter(textBox31.Text);
                    string line15 = "Projects\t\t\t" + firstGap + ":" + secondGap + lines.Item1;
                    doc.InsertParagraph(line15, false, lineFormat).LineSpacingAfter = 10;
                    //split line
                    if (lines.Item2.Length > checkLnth)
                    {
                        var linee = Splitter(lines.Item2);
                        string line151 = "\t\t\t\t" + firstGap + " " + secondGap + linee.Item1;
                        doc.InsertParagraph(line151, false, lineFormat).LineSpacingAfter = 10;
                        //split line
                        if (linee.Item2.Length > checkLnth)
                        {
                            var lineee = Splitter(linee.Item2);
                            string line1511 = "\t\t\t\t" + firstGap + " " + secondGap + lineee.Item1;
                            doc.InsertParagraph(line1511, false, lineFormat).LineSpacingAfter = 10;
                            //split line
                            if (lineee.Item2.Length > checkLnth)
                            {
                                var lineeee = Splitter(lineee.Item2);
                                string line15111 = "\t\t\t\t" + firstGap + " " + secondGap + lineeee.Item1;
                                doc.InsertParagraph(line15111, false, lineFormat).LineSpacingAfter = 10;


                                string liney111 = "\t\t\t\t" + firstGap + " " + secondGap + lineeee.Item2;
                                doc.InsertParagraph(liney111, false, lineFormat).LineSpacingAfter = 10;

                            }
                            else
                            {
                                string liney11 = "\t\t\t\t" + firstGap + " " + secondGap + lineee.Item2;
                                doc.InsertParagraph(liney11, false, lineFormat).LineSpacingAfter = 10;
                            }
                        }
                        else
                        {
                            string liney1 = "\t\t\t\t" + firstGap + " " + secondGap + linee.Item2;
                            doc.InsertParagraph(liney1, false, lineFormat).LineSpacingAfter = 10;
                        }
                    }
                    else
                    {
                        string liney = "\t\t\t\t" + firstGap + " " + secondGap + lines.Item2;
                        doc.InsertParagraph(liney, false, lineFormat).LineSpacingAfter = 10;
                    }
                }
                else
                {
                    string line15 = "Projects\t\t\t" + firstGap + ":" + secondGap + textBox31.Text;
                    doc.InsertParagraph(line15, false, lineFormat).LineSpacingAfter = 10;
                }
            }
            if (textBox32.Text != "")
            {
                //split line
                if (textBox32.Text.Length > checkLnth)
                {

                    var lines = Splitter(textBox32.Text);
                    string line13 = "Internship/Experience\t" + firstGap + ":" + secondGap + lines.Item1;
                    doc.InsertParagraph(line13, false, lineFormat).LineSpacingAfter = 10;
                    //split line
                    if (lines.Item2.Length > checkLnth)
                    {
                        var linee = Splitter(lines.Item2);
                        string line131 = "\t\t\t\t" + firstGap + " " + secondGap + linee.Item1;
                        doc.InsertParagraph(line131, false, lineFormat).LineSpacingAfter = 10;
                        //split line
                        if (linee.Item2.Length > checkLnth)
                        {
                            var lineee = Splitter(linee.Item2);
                            string line1311 = "\t\t\t\t" + firstGap + " " + secondGap + lineee.Item1;
                            doc.InsertParagraph(line1311, false, lineFormat).LineSpacingAfter = 10;
                            //split line
                            if (lineee.Item2.Length > checkLnth)
                            {
                                var lineeee = Splitter(lineee.Item2);
                                string line13111 = "\t\t\t\t" + firstGap + " " + secondGap + lineeee.Item1;
                                doc.InsertParagraph(line13111, false, lineFormat).LineSpacingAfter = 10;


                                string liney111 = "\t\t\t\t" + firstGap + " " + secondGap + lineeee.Item2;
                                doc.InsertParagraph(liney111, false, lineFormat).LineSpacingAfter = 10;

                            }
                            else
                            {
                                string liney11 = "\t\t\t\t" + firstGap + " " + secondGap + lineee.Item2;
                                doc.InsertParagraph(liney11, false, lineFormat).LineSpacingAfter = 10;
                            }
                        }
                        else
                        {
                            string liney1 = "\t\t\t\t" + firstGap + " " + secondGap + linee.Item2;
                            doc.InsertParagraph(liney1, false, lineFormat).LineSpacingAfter = 10;
                        }
                    }
                    else
                    {
                        string liney = "\t\t\t\t" + firstGap + " " + secondGap + lines.Item2;
                        doc.InsertParagraph(liney, false, lineFormat).LineSpacingAfter = 10;
                    }
                }
                else
                {
                    string line13 = "Internship/Experience\t" + firstGap + ":" + secondGap + textBox32.Text;
                    doc.InsertParagraph(line13, false, lineFormat).LineSpacingAfter = 10;
                }
            }
            if (textBox9.Text != "")
            {               
                //split line
                if(textBox9.Text.Length > checkLnth)
                {
                   
                    var lines = Splitter(textBox9.Text);
                    string line14 = "Achievements\t\t\t" + firstGap + ":" + secondGap + lines.Item1;
                    doc.InsertParagraph(line14, false, lineFormat).LineSpacingAfter = 10;
                    //split line
                    if (lines.Item2.Length > checkLnth)
                    {
                        var linee = Splitter(lines.Item2);
                        string line141 = "\t\t\t\t" + firstGap + " " + secondGap + linee.Item1;
                        doc.InsertParagraph(line141, false, lineFormat).LineSpacingAfter = 10;
                        //split line
                        if (linee.Item2.Length > checkLnth)
                        {
                            var lineee = Splitter(linee.Item2);
                            string line1411 = "\t\t\t\t" + firstGap + " " + secondGap + lineee.Item1;
                            doc.InsertParagraph(line1411, false, lineFormat).LineSpacingAfter = 10;
                            //split line
                            if (lineee.Item2.Length > checkLnth)
                            {
                                var lineeee = Splitter(lineee.Item2);
                                string line14111 = "\t\t\t\t" + firstGap + " " + secondGap + lineeee.Item1;
                                doc.InsertParagraph(line14111, false, lineFormat).LineSpacingAfter = 10;


                                string liney111 = "\t\t\t\t" + firstGap + " " + secondGap + lineeee.Item2;
                                doc.InsertParagraph(liney111, false, lineFormat).LineSpacingAfter = 10;

                            }
                            else
                            {
                                string liney11 = "\t\t\t\t" + firstGap + " " + secondGap + lineee.Item2;
                                doc.InsertParagraph(liney11, false, lineFormat).LineSpacingAfter = 10;
                            }
                        }
                        else
                        {
                            string liney1 = "\t\t\t\t" + firstGap + " " + secondGap + linee.Item2;
                            doc.InsertParagraph(liney1, false, lineFormat).LineSpacingAfter = 10;
                        }
                    }
                    else
                    {
                        string liney = "\t\t\t\t" + firstGap + " " + secondGap + lines.Item2;
                        doc.InsertParagraph(liney, false, lineFormat).LineSpacingAfter = 10;
                    }
                }
                else
                {
                    string line14 = "Achievements\t\t\t" + firstGap + ":" + secondGap + textBox9.Text;
                    doc.InsertParagraph(line14, false, lineFormat).LineSpacingAfter = 10;
                }
              
            }
            if (richTextBox1.Text != "")
            {
                //split line
                if (richTextBox1.Text.Length > checkLnth)
                {

                    var lines = Splitter(richTextBox1.Text);
                    string line11 = "Other Skills\t\t\t" + firstGap + ":" + secondGap + lines.Item1;
                    doc.InsertParagraph(line11, false, lineFormat).LineSpacingAfter = 10;
                    //split line
                    if (lines.Item2.Length > checkLnth)
                    {
                        var linee = Splitter(lines.Item2);
                        string line111 = "\t\t\t\t" + firstGap + " " + secondGap + linee.Item1;
                        doc.InsertParagraph(line111, false, lineFormat).LineSpacingAfter = 10;
                        //split line
                        if (linee.Item2.Length > checkLnth)
                        {
                            var lineee = Splitter(linee.Item2);
                            string line1111 = "\t\t\t\t" + firstGap + " " + secondGap + lineee.Item1;
                            doc.InsertParagraph(line1111, false, lineFormat).LineSpacingAfter = 10;
                            //split line
                            if (lineee.Item2.Length > checkLnth)
                            {
                                var lineeee = Splitter(lineee.Item2);
                                string line11111 = "\t\t\t\t" + firstGap + " " + secondGap + lineeee.Item1;
                                doc.InsertParagraph(line11111, false, lineFormat).LineSpacingAfter = 10;


                                string liney111 = "\t\t\t\t" + firstGap + " " + secondGap + lineeee.Item2;
                                doc.InsertParagraph(liney111, false, lineFormat).LineSpacingAfter = 10;

                            }
                            else
                            {
                                string liney11 = "\t\t\t\t" + firstGap + " " + secondGap + lineee.Item2;
                                doc.InsertParagraph(liney11, false, lineFormat).LineSpacingAfter = 10;
                            }
                        }
                        else
                        {
                            string liney1 = "\t\t\t\t" + firstGap + " " + secondGap + linee.Item2;
                            doc.InsertParagraph(liney1, false, lineFormat).LineSpacingAfter = 10;
                        }
                    }
                    else
                    {
                        string liney = "\t\t\t\t" + firstGap + " " + secondGap + lines.Item2;
                        doc.InsertParagraph(liney, false, lineFormat).LineSpacingAfter = 10;
                    }
                }
                else
                {
                    string line11 = "Other Skills\t\t\t" + firstGap + ":" + secondGap + richTextBox1.Text;
                    doc.InsertParagraph(line11, false, lineFormat).LineSpacingAfter = 10;
                }

                
            }
            doc.InsertParagraph(NL, false, lineFormat);
            //Personal Details
            doc.InsertParagraph(PD, false, objformat).LineSpacingAfter = 10;
            doc.InsertParagraph(line2, false, lineFormat).LineSpacingAfter = 10;
            doc.InsertParagraph(line3, false, lineFormat).LineSpacingAfter = 10;
            doc.InsertParagraph(line4, false, lineFormat).LineSpacingAfter = 10;
            doc.InsertParagraph(line6, false, lineFormat).LineSpacingAfter = 10;
            doc.InsertParagraph(line5, false, lineFormat).LineSpacingAfter = 10;
            doc.InsertParagraph(line8, false, lineFormat).LineSpacingAfter = 10;
            doc.InsertParagraph(line9, false, lineFormat).LineSpacingAfter = 10;
            doc.InsertParagraph(line10, false, lineFormat).LineSpacingAfter = 10;
            doc.InsertParagraph(line101, false, lineFormat).LineSpacingAfter = 10;

            if (textBox7.Text != "")
            {
                string line12 = "Hobbies\t\t\t" + firstGap + ":" + secondGap + textBox7.Text;
                doc.InsertParagraph(line12, false, lineFormat).LineSpacingAfter = 10;
            }
            doc.InsertParagraph(NL, false, lineFormat);
            //Declaration
            doc.InsertParagraph(DC, false, objformat).LineSpacingAfter = 10;
            doc.InsertParagraph(line0, false, lineFormat).LineSpacingAfter = 10;
            doc.InsertParagraph(NL, false, lineFormat);
            doc.InsertParagraph(place, false, endformat).LineSpacingAfter = 10;
            doc.InsertParagraph(date, false, endformat).LineSpacingAfter = 10;

            //            doc.InsertParagraph(line17, false, lineFormat).LineSpacingAfter = 10;

            doc.Save();


        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)//Gender
        {
            
            if(radioButton1.Checked == true)
            {
                gender = "Male";
            }
            if (radioButton2.Checked == true)
            {
                gender = "Female";
            }

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void textBox28_TextChanged(object sender, EventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }
                
        private void next4_Click(object sender, EventArgs e)
        {
            if (!Validate(textBox17.Text)) return;
            if (!Validate(textBox16.Text)) return;
            if (!ValidateNum(textBox15.Text)) return;
            if (!ValidateGPA(textBox14.Text)) return;
            tabControl1.SelectedTab = tabPage5;
            //Btech(true);
            // Inter(false);
        }

        private void next3_Click(object sender, EventArgs e)
        {
            if (!Validate(textBox21.Text)) return;
            if (!Validate(textBox20.Text)) return;
            if (!ValidateNum(textBox19.Text)) return;
            if (!ValidateGPA(textBox18.Text)) return;
            tabControl1.SelectedTab = tabPage4;

            //  Inter(true);
            //  SSC(false);
        }

        private void next2_Click(object sender, EventArgs e)
        {
            if(!ValidateText(textBox2.Text)) return; //ln
            if(!ValidateText(textBox1.Text)) return; //fn
            if(!Validate(textBox3.Text)) return; //dob
            if(!ValidateText(textBox4.Text)) return; //father
            if(!ValidateText(textBox5.Text)) return; //mother
            if(!ValidateNum(textBox28.Text)) return; //phone
            if(!ValidateNum(textBox26.Text)) return;
            if(!ValidateText(textBox22.Text)) return;
            if(!Validate(textBox23.Text)) return;
            if(!Validate(textBox24.Text)) return;
            if(!Validate(textBox25.Text)) return;
            tabControl1.SelectedTab = tabPage2;
            //SSC(true);
            // Personal(false);
        }             

        private void button7_Click(object sender, EventArgs e)
        {
          
            if (!Validate(textBox10.Text)) return;
            if (!Validate(comboBox1.Text)) return;
            if (!ValidateNum(textBox12.Text)) return;
            if (!ValidateGPA(textBox13.Text)) return;
            if (!Validate(comboBox1.Text)) return;
            tabControl1.SelectedTab = tabPage6;


            // Skills(true);
            //  Btech(false);
        }

        private void Personal(bool result)
        {
            next2.Visible = result;
            label35.Visible = result;
            label34.Visible = result;
            textBox2.Visible = result;
            textBox1.Visible = result;
            textBox3.Visible = result;
            groupBox4.Visible = result;
            radioButton1.Visible = result;
            label2.Visible = result;
            textBox4.Visible = result;
            label3.Visible = result;
            label18.Visible = result;
            label4.Visible = result;
            textBox5.Visible = result;
            radioButton2.Visible = result;
            label5.Visible = result;
            textBox28.Visible = result;
            label6.Visible = result;
        }

        private void SSC(bool result)
        {
            label37.Visible = result;
            groupBox2.Visible = result;
            next3.Visible = result;
        }

        private void Inter(bool result)
        {
            next4.Visible = result;
            label38.Visible = result;
            groupBox1.Visible = result;
        }

        private void Btech(bool result)
        {
            button7.Visible = result;
            label36.Visible = result;
            groupBox3.Visible = result;
            comboBox1.Visible = result;
        }

        private void Skills(bool result)
        {
            label39.Visible = result;
            label10.Visible = result;
            checkBox2.Visible = result;
            label14.Visible = result;
            label13.Visible = result;
            richTextBox1.Visible = result;
            checkBox4.Visible = result;
            label9.Visible = result;
            textBox32.Visible = result;
            textBox7.Visible = result;
            label8.Visible = result;
            label12.Visible = result;
            textBox6.Visible = result;
            checkBox1.Visible = result;
            textBox31.Visible = result;
            textBox29.Visible = result;
            label11.Visible = result;
            button1.Visible = result;
            checkBox3.Visible = result;
            textBox9.Visible = result;
            label7.Visible = result;
        }

        private void Template(bool result)
        {
            button13.Visible = result;
            button14.Visible = result;
            pictureBox1.Visible = result;
            pictureBox2.Visible = result;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Personal(true);
            SSC(false);
            Inter(false);
            Btech(false);
            Skills(false);
            Template(false);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Personal(false);
            SSC(true);
            Inter(false);
            Btech(false);
            Skills(false);
            Template(false);

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Personal(false);
            SSC(false);
            Inter(true);
            Btech(false);
            Skills(false);
            Template(false);

        }

        private void button5_Click(object sender, EventArgs e)
        {
            Personal(false);
            SSC(false);
            Inter(false);
            Btech(true);
            Skills(false);
            Template(false);

        }

        private void button6_Click(object sender, EventArgs e)
        {
            Personal(false);
            SSC(false);
            Inter(false);
            Btech(false);
            Template(false);
            Skills(true);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            SSC(false);
            Inter(false);
            Personal(false);
            Skills(false);
            Btech(false);

            Template(true);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Resume Automation is a Freeware Windows Software Tool that is designed for students which lets user to create the Resume.docx automatically by providing the necessary information.\nFirst Release - July 2017\n\nCopyright 2017. ");
        }//About

        private void button10_Click(object sender, EventArgs e)
        {
            MessageBox.Show("For any queries or information regarding this software\nPlease Contact/Follow at:\n\nFacebook - fb.com/tejavictorystark\nGmail        - tejaasvenkatesh@gmail.com\nLinkedIn   - linkedin.com/in/tejavictorystark\nTwitter      - twitter.com/tejaasvenkatesh");
        }//Contact

        private void button11_Click(object sender, EventArgs e)//Credits
        {
            MessageBox.Show("Developer Information:\n\n K. Sai Krishna Teja -- B.Tech 4th Year\n148W1A05E5\nV.R.Siddhartha Engineering College\n                              All Rights Reserved.");
        }

        private Tuple<string,string> Splitter (string s)
        {
            int prevIndex = 0;
            int count = 0;
            foreach(char c in s)
            {
                if(c == ' ')
                {
                    count ++;
                    int index = IndexOfNth(s, count);
                    if (index > checkLnth)
                    {
                        string one = s.Substring(0,prevIndex);
                        string two = s.Substring(prevIndex + 1, s.Length-one.Length-2);
                        return new Tuple<string, string>(one, two);
                    }
                    prevIndex = index;
                }
            }
            return new Tuple<string, string>(" "," ");
        }

        public static int IndexOfNth(string str, int nth)
        {
            string value = " ";
            int offset = str.IndexOf(value);
            for (int i = 1; i < nth; i++)
            {
                offset = str.IndexOf(value, offset + 1);
            }
            return offset;
        }

        private void button12_Click(object sender, EventArgs e)//RESET
        {
            textBox2.Text = "";
            textBox1.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox28.Text = "";
            textBox14.Text = "";
            textBox15.Text = "";
            textBox16.Text = "";
            textBox17.Text = "";
            textBox13.Text = "";
            textBox12.Text = "";
            comboBox1.Text = "";
            textBox10.Text = "";
            textBox27.Text = "";
            textBox26.Text = "";
            textBox22.Text = "";
            textBox23.Text = "";
            textBox24.Text = "";
            textBox25.Text = "";
            textBox18.Text = "";
            textBox19.Text = "";
            textBox20.Text = "";
            textBox21.Text = "";
            richTextBox1.Text = "";
            textBox32.Text = "";
            textBox7.Text = "";
            textBox6.Text = "";
            textBox31.Text = "";
            textBox29.Text = "";
            textBox9.Text = "";
            radioButton1.Checked = true;
            richTextBox2.Text = "To work in a challenging environment, where I can utilize my skills and continuously enhance my skills. I would work to be a good team player.";
            label43.Text = "Default : Template 1";
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
        }

        private bool Validate (string s)
        {
            if (String.IsNullOrWhiteSpace(s))
            {
                MessageBox.Show("Please Enter a Value");
                return false;
            }
            return true;
        }

        private bool ValidateText (string s)
        {
            if (String.IsNullOrWhiteSpace(s))
            {
                MessageBox.Show("Please Enter a Value");
                return false;
            }
            
            //if (!s.All(c => Char.IsLetter(c)))
            //{
            //    MessageBox.Show("Please Enter Valid Characters Only");
            //    return false;
            //}
            return true;
        }

        private bool ValidateNum (string s)
        {
            if (String.IsNullOrWhiteSpace(s))
            {
                MessageBox.Show("Please Enter a Value");
                return false;
            }
            if (!s.All(c => Char.IsNumber(c)))
            {
                MessageBox.Show("Please Enter Valid Numbers Only");
                return false;
            }
            return true;
        }

        private bool ValidateTextNum (string s)
        {
            if (String.IsNullOrWhiteSpace(s))
            {
                MessageBox.Show("Please Enter a Value");
                return false;
            }
            if (!s.All(c => Char.IsLetterOrDigit(c)))
            {
                MessageBox.Show("Please Enter Valid Characters Only");
                return false;
            }
            return true;
        }

        private bool ValidateGPA(string s)
        {
            if (String.IsNullOrWhiteSpace(s))
            {
                MessageBox.Show("Please Enter a Value");
                return false;
            }
            foreach (char c in s)
            {
                if (Char.IsLetter(c))
                {
                    MessageBox.Show("Please Enter Valid Marks/GPA Only");
                    return false;
                }
            }
            
            return true;
        }
        
        private string SaveFile()
        {
            // Show the FolderBrowserDialog.
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                string folderName = folderBrowserDialog1.SelectedPath;
                MessageBox.Show("Saved To:\n" + folderName);
                return folderName;
            }
            return " ";
        }

        private void button13_Click(object sender, EventArgs e)//Template 1
        {
            template = 1;
            label43.Text = "Template 1";
            //Personal(true);
            //Template(false);
        }

        private void button14_Click(object sender, EventArgs e)//Template 2
        {
            template = 2;
            label43.Text = "Template 2";
            // Personal(true);
            //Template(false);
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            
                //exit application when form is closed
                Application.Exit();
            
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (!Validate(richTextBox2.Text)) return;
            tabControl1.SelectedTab = tabPage3;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)//link to fb page
        {
            System.Diagnostics.Process.Start("https://www.facebook.com/BaymaxSoft/");
        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)//calendar show
        {
            if (monthCalendar1.Visible == false)
            {
                monthCalendar1.Show();
            }
            else
            {
                monthCalendar1.Hide();
            }
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            textBox3.Text = monthCalendar1.SelectionStart.ToShortDateString();
            this.monthCalendar1.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar1_DateSelected);

        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            monthCalendar1.Hide();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show("New Improvements in V2.0.1:\n1. UI Improved\n2. GUI Elements Validated\n3. Calendar added for Date of Birth\n4. Splash Screen added\n5. New Box added for Objective\n\n Follow and Support us @fb.com/BaymaxSoft");
        }
    }
}
