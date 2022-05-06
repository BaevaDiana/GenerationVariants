using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Words.NET;

namespace GenerationVariants
{
    public partial class Form1 : Form
    {
        public int n;
        public string varPoint = "0";
        public Form1()
        {
            InitializeComponent();
        }
        private int Factorial(int x)
        {
            return (x > 1) ? x * Factorial(x - 1) : 1;
        }

        // факториал в диапазоне [border; x]
        private int FactorialPlus(int x, int border)
        {
            return (x > border) ? x * FactorialPlus(x - 1, border) : 1;
        }

        private double Combinations(int k, int n)
        {
            return FactorialPlus(n, n - k) / Factorial(k);
        }

        private void generateButton_Click(object sender, EventArgs e)
        {
            string fileName = @"C:\Users\Дианочка\source\repos\GenerationVariants\ex.docx";
            string fileName2 = @"C:\Users\Дианочка\source\repos\GenerationVariants\ans.docx";
            var doc = DocX.Create(fileName);
            var doc2 = DocX.Create(fileName2);
            int k, z;
            Random random = new Random();
            for (k = 1; k <= n; k++)
            {
                z = 0;
                string s = "" + "Вариант №" + k + Environment.NewLine;
                doc.InsertParagraph(s);
                doc2.InsertParagraph(s);
                if (checkBox1.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {
                        int h, i = 8;
                        h = random.Next(2, 10);
                        string s2 = "" + z + ")Сколько имеется чисел c" + " " + h + " " + "знаками, все цифры у которых различны?" + Environment.NewLine;
                        doc.InsertParagraph(s2);
                        int ans = 9 * 9;
                        while (h != 2)
                        {
                            ans *= i;
                            i--;
                            h--;
                        }
                        string s3 = "" + "Задание " + z + " -" + ans + Environment.NewLine;
                        doc2.InsertParagraph(s3);
                    }
                    else
                    {
                        double h1, p2;
                        h1 = random.Next(2, 17);//этажи
                        p2 = random.Next(2, 11);//люди
                        string s4 = "" + z + ")" + p2 + " человек вошли в лифт на 1-м этаже" + " дома с" + " " + h1 + " этажами. Сколькими способами пассажиры могут выйти из лифта на нужных этажах?" + Environment.NewLine;
                        doc.InsertParagraph(s4);
                        int ans1 = (int)Math.Pow(h1 - 1, p2);
                        string s5 = "" + "Задание " + z + "- " + ans1 + Environment.NewLine;
                        doc2.InsertParagraph(s5);
                    }
                }

                if (checkBox2.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {
                        int m1, w1, h1, m2, h2, w2, g;
                        m1 = random.Next(2, 16);//всего в группе мужчин
                        w1 = random.Next(2, 16);//женшин в группе
                        g = random.Next(2, m1 + w1);//группа
                        w2 = random.Next(2, g);//в группе2 женщин
                        m2 = random.Next(2, g);//в группе2 мужчин
                        string s5 = "" + z + ")Группа туристов из " + m1 + " юношей и " + w1 + " девушек выбирает по жребию " + g + " человек для приготовления ужина. Сколько существует способов, при которых в эту группу попадут " + w2 + " девушек или " + m2 + " юношей?" + Environment.NewLine;
                        doc.InsertParagraph(s5);
                        double ans5;
                        ans5 = Combinations(w2, w1) + Combinations(m2, m1);
                        string s6 = "" + "Задание " + z + " - " + ans5 + Environment.NewLine;
                        doc2.InsertParagraph(s6);
                    }
                    else
                    {
                        int m1, w1, h1, m2, h2, w2, g;
                        m1 = random.Next(2, 16);//всего деталей
                        w1 = random.Next(2, m1);//бракованные
                        g = random.Next(2, m1);//комплект
                        w2 = random.Next(2, g);//бракованные выпадают
                        string s7 = "" + z + ")В ящике " + m1 + " деталей, среди которых  " + w1 + " бракованных. Наудачу выбирается комплект из  " + g + " деталей. Сколько всего комплектов, в каждом из которых " + w2 + " детали бракованные?" + Environment.NewLine;
                        doc.InsertParagraph(s7);
                        double ans9;
                        ans9 = Combinations(w2, w1) * Combinations(g - w2, m1 - w1);
                        string s6 = "" + "Задание " + z + " - " + ans9 + Environment.NewLine;
                        doc2.InsertParagraph(s6);
                    }
                }

                if (checkBox3.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {
                        int s1, s2;
                        s1 = random.Next(2, 11);
                        s2 = random.Next(2, s1);
                        string s9 = "" + z + ")В комнате имеется " + s1 + " стульев. Сколькими способами можно разместить на них " + s2 + " гостей?" + Environment.NewLine;
                        doc.InsertParagraph(s9);
                        int ans7 = 1, c = 0;
                        while (c != s2)
                        {
                            ans7 *= s1;
                            s1--;
                            c++;
                        }
                        string s10 = "" + "Задание " + z + " - " + ans7 + Environment.NewLine;
                        doc2.InsertParagraph(s10);
                    }
                    else
                    {
                        int r; string l;
                        r = random.Next(1, 4);
                        if (r == 1) l = "КНИГА";
                        else
                            if (r == 2) l = "ВЕСНА";
                        else
                            l = "ЛЕКЦИЯ";
                        string s11 = "" + z + ")Сколько различных «слов» можно получить, переставляя буквы в слове " + l + "?" + Environment.NewLine;
                        doc.InsertParagraph(s11);
                        int ans8 = Factorial(5) / 1;
                        string s12 = "" + "Задание " + z + " - " + ans8 + Environment.NewLine;
                        doc2.InsertParagraph(s12);
                    }
                }

                if (checkBox4.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {
                        int r1; string l;
                        r1 = random.Next(1, 4);
                        if (r1 == 1) l = "ПРОФЕССОР";
                        else
                            if (r1 == 2) l = "АВИАЛИНИЯ";
                        else
                            l = "БАДМИНТОН";
                        string s13 = "" + z + ")Сколько различных «слов» можно получить, переставляя буквы в слове " + l + "?" + Environment.NewLine;
                        doc.InsertParagraph(s13);
                        int ans10;
                        if (r1 == 1) ans10 = 45360;
                        else if (r1 == 2) ans10 = 30240;
                        else ans10 = 181440;
                        string s14 = "" + "Задание " + z + "- " + ans10 + Environment.NewLine;
                        doc2.InsertParagraph(s14);
                    }
                    else
                    {
                        int r2; string l;
                        r2 = random.Next(1, 4);
                        if (r2 == 1) l = "ХОЛОДИЛЬНИК";
                        else
                            if (r2 == 2) l = "БАКАЛАВРИАТ";
                        else
                            l = "КАЛАМБУРНЫЙ";
                        string s15 = "" + z + ")Сколько различных «слов» можно получить, переставляя буквы в слове " + l + "?" + Environment.NewLine;
                        doc.InsertParagraph(s15);
                        int ans11;
                        if (r2 == 1) ans11 = 9979200;
                        else if (r2 == 2) ans11 = 1663200;
                        else ans11 = 19958400;
                        string s16 = "" + "Задание " + z + " - " + ans11 + Environment.NewLine;
                        doc2.InsertParagraph(s16);
                    }
                }

                if (checkBox5.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {
                        double t1, t2, t3;
                        t1 = random.Next(3, 6);
                        t2 = random.Next(3, 10);
                        t3 = random.Next(1, (int)t2);
                        string s13 = "" + z + ")В контрольной работе будет " + t1 + " задач – по одной из каждой пройденной темы. Задачи будут взяты из общего списка по " + t2 + " задач в каждой теме, а всего было пройдено " + t1 + " тем. При подготовке к контрольной Вова решил только по " + t3 + " задач в каждой теме. Найдите общее число всех возможных вариантов контрольной работы." + Environment.NewLine;
                        doc.InsertParagraph(s13);
                        int ans12;
                        ans12 = (int)Math.Pow(t2, t1);
                        string s14 = "" + "Задание " + z + "- " + ans12 + Environment.NewLine;
                        doc2.InsertParagraph(s14);
                    }
                    else
                    {
                        int g1; string l;
                        g1 = random.Next(4, 8);
                        string s15 = "" + z + ")В футбольном турнире участвуют несколько команд. Оказалось, что все они для трусов и футболок использовали  " + g1 + " цветов, причем были представлены все возможные варианты. Сколько команд участвовали в турнире?" + Environment.NewLine;
                        doc.InsertParagraph(s15);
                        int ans13;
                        ans13 = g1 * (g1 - 1);
                        string s16 = "" + "Задание " + z + " - " + ans13 + Environment.NewLine;
                        doc2.InsertParagraph(s16);
                    }
                }

                if (checkBox6.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {
                        double h1;
                        h1 = random.Next(2, 13);
                        string s17 = "" + z + ")Бросаются две игральные кости. Определить вероятность того, что: а) сумма числа очков не превосходит " + h1 + "; б) произведение числа очков не превосходит " + h1 + "; в) произведение числа очков делится на " + h1 + "." + Environment.NewLine;
                        doc.InsertParagraph(s17);
                        double ans13, ans14, ans15;
                        double count1 = 0, count2 = 0, count3 = 0;
                        for (int i = 0; i < 6; i++)
                        {
                            for (int j = 0; j < 6; j++)
                            {
                                if (i + j <= h1) count1++;
                                if (i * j <= h1) count2++;
                                if (i * j % h1 == 0) count3++;
                            }
                        }
                        ans13 = count1 / 36; ans14 = count2 / 36; ans15 = count3 / 36;
                        string s14 = "" + "Задание " + z + "- " + "a)" + ans13 + ";б)" + ans14 + ";в)" + ans15 + Environment.NewLine;
                        doc2.InsertParagraph(s14);
                    }
                    else
                    {
                        double h1;
                        h1 = random.Next(2, 13);
                        string s18 = "" + z + ")Бросаются две игральные кости. Определить вероятность того, что: а) сумма числа очков не превосходит " + h1 + "; б) произведение числа очков не превосходит " + h1 + "; в) произведение числа очков делится на " + h1 + "." + Environment.NewLine;
                        doc.InsertParagraph(s18);
                        double ans13, ans14, ans15;
                        double count1 = 0, count2 = 0, count3 = 0;
                        for (int i = 0; i < 6; i++)
                        {
                            for (int j = 0; j < 6; j++)
                            {
                                if (i + j <= h1) count1++;
                                if (i * j <= h1) count2++;
                                if (i * j % h1 == 0) count3++;
                            }
                        }
                        ans13 = count1 / 36; ans14 = count2 / 36; ans15 = count3 / 36;
                        string s19 = "" + "Задание " + z + "- " + "a)" + ans13 + ";б)" + ans14 + ";в)" + ans15 + Environment.NewLine;
                        doc2.InsertParagraph(s19);
                    }
                }

                if (checkBox7.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {
                        double p1, p2;
                        p1 = Math.Round(random.NextDouble() * 1, 1);
                        p2 = Math.Round(random.NextDouble() * 1, 1);
                        string s20 = "" + z + ")Два студента ищут нужную им книгу в букинистических магазинах. Вероятность того, что книга будет найдена первым студентом, равна " + p1 + ", а вторым " + p2 + ". Какова вероятность того, что: а) только один из студентов найдет книгу; б) оба студента найдут книгу; в) хотя бы один студент найдет книгу. " + Environment.NewLine;
                        doc.InsertParagraph(s20);
                        double ans16, ans17, ans18;
                        ans16 = p1 * (1 - p2) + p2 * (1 - p1);
                        ans17 = p1 * p2;
                        ans18 = (1 - ((1 - p1) * (1 - p2)));
                        string s21 = "" + "Задание " + z + "- " + "a)" + ans16 + ";б)" + ans17 + ";в)" + ans18 + Environment.NewLine;
                        doc2.InsertParagraph(s21);
                    }
                    else
                    {
                        double p1;
                        int c1, c2; ;
                        p1 = Math.Round(random.NextDouble() * 1, 1);
                        c1 = random.Next(2, 6);
                        c2 = random.Next(2, c1);
                        string s20 = "" + z + ")Вероятность выигрыша по лотерейному билету " + p1 + ". Приобретено " + c1 + " билета. Какова вероятность того, что выигрыша: а) только по одному из купленных билетов; б) только по " + c2 + " из купленных билетов; в) хотя бы по одному билету?" + Environment.NewLine;
                        doc.InsertParagraph(s20);
                        double ans19, ans20, ans21;
                        ans19 = Combinations(1, c1) * (double)Math.Pow(p1, 1) * (double)Math.Pow((1 - p1), c1 - 1);
                        ans20 = Combinations(c2, c1) * (double)Math.Pow(p1, c2) * (double)Math.Pow((1 - p1), c1 - c2);
                        ans21 = 1 - (Combinations(0, c1) * (double)Math.Pow(p1, 0) * (double)Math.Pow((1 - p1), c1));
                        string s21 = "" + "Задание " + z + "- " + "a)" + ans19 + ";б)" + ans20 + ";в)" + ans21 + Environment.NewLine;
                        doc2.InsertParagraph(s21);
                    }
                }

                if (checkBox8.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {

                        string s22 = "" + "Задание доделать!" + Environment.NewLine;
                        doc.InsertParagraph(s22);

                    }
                    else
                    {
                        int b1, b2, b3, b4;
                        b1 = random.Next(3, 10);//всего инженеров
                        b2 = random.Next(2, b1);//всего женщин
                        b3 = random.Next(2, b1 - 1);//в смене человек
                        b4 = random.Next(2, b3);//мужчин в смене
                        //while (b4 > b1 - b2) b4 = random.Next(2, b3);
                        string s23 = "" + z + ")На тепловой электростанции " + b1 + " сменных инженеров, из них " + b2 + " женщин. В смену занято " + b3 + " человека. Найти вероятность того, что в случайно выбранную смену окажется " + b4 + " мужчин." + Environment.NewLine;
                        doc.InsertParagraph(s23);
                        double ans23;
                        ans23 = (double)(Combinations(b4, b1 - b2) * Combinations(b3 - b4, b2)) / (Combinations(b3, b1));
                        string s24 = "" + "Задание " + z + " - " + ans23 + Environment.NewLine;
                        doc2.InsertParagraph(s24);
                    }
                }
















            }
            doc.Save();
                doc2.Save();
                Process.Start("WINWORD.EXE", fileName);
                Process.Start("WINWORD.EXE", fileName2);
            
        }


        private void inputNumberOfVariants_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(Count.Text))
            {
                varPoint = "0";
                MessageBox.Show("Ошибка.Введите количество вариантов!");
            }
            else
            {
                varPoint = "1";
                n = Convert.ToInt32(Count.Text);
            }
        }
    }
}

    
        
