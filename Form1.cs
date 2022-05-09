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
                        string s4 = "" + z + ")" + p2 + " человек(-a) вошли в лифт на 1-м этаже" + " дома с" + " " + h1 + " этажами. Сколькими способами пассажиры могут выйти из лифта на нужных этажах?" + Environment.NewLine;
                        doc.InsertParagraph(s4);
                        double ans1 = (int)Math.Pow(h1 - 1, p2);
                        string s5 = "" + "Задание " + z + "- " + ans1 + Environment.NewLine;
                        doc2.InsertParagraph(s5);
                    }
                }

                if (checkBox2.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {
                        int m1, w1, m2, w2, g;
                        m1 = random.Next(2, 16);//всего в группе мужчин
                        w1 = random.Next(2, 16);//женшин в группе
                        g = random.Next(2, m1 + w1);//группа для ужина
                        w2 = random.Next(2, w1);//в группе2 женщин
                        m2 = random.Next(2, m1);//в группе2 мужчин
                        string s5 = "" + z + ")Группа туристов из " + m1 + " юношей и " + w1 + " девушек выбирает по жребию " + g + " человек(-a) для приготовления ужина. Сколько существует способов, при которых в эту группу попадут " + w2 + " девушек или " + m2 + " юношей?" + Environment.NewLine;
                        doc.InsertParagraph(s5);
                        double ans5;
                        ans5 = Combinations(w2, w1) + Combinations(m2, m1);
                        string s6 = "" + "Задание " + z + " - " + ans5 + Environment.NewLine;
                        doc2.InsertParagraph(s6);
                    }
                    else
                    {
                        int m1, w1, w2, g;
                        m1 = random.Next(2, 16);//всего деталей
                        w1 = random.Next(2, m1);//бракованные
                        g = random.Next(2, w1);//комплект
                        w2 = random.Next(2, g);//бракованные выпадают
                        string s7 = "" + z + ")В ящике " + m1 + " детали(-ей), среди которых  " + w1 + " бракованных. Наудачу выбирается комплект из  " + g + " деталей. Сколько всего комплектов, в каждом из которых " + w2 + " детали(-ей) бракованные(-ых)?" + Environment.NewLine;
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
                        string s13 = "" + z + ")В контрольной работе будет " + t1 + " задач(-и) – по одной из каждой пройденной темы. Задачи будут взяты из общего списка по " + t2 + " задач(-и) в каждой теме, а всего было пройдено " + t1 + " тем(-ы). При подготовке к контрольной Вова решил только по " + t3 + " задач(-е) в каждой теме. Найдите общее число всех возможных вариантов контрольной работы." + Environment.NewLine;
                        doc.InsertParagraph(s13);
                        double ans12;
                        ans12 = (int)Math.Pow(t2, t1);
                        string s14 = "" + "Задание " + z + "- " + ans12 + Environment.NewLine;
                        doc2.InsertParagraph(s14);
                    }
                    else
                    {
                        int g1;
                        g1 = random.Next(4, 8);
                        string s15 = "" + z + ")В футбольном турнире участвуют несколько команд. Оказалось, что все они для трусов и футболок использовали " + g1 + " цвета(-ов), причем были представлены все возможные варианты. Сколько команд участвовали в турнире?" + Environment.NewLine;
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
                        string s20 = "" + z + ")Два студента ищут нужную им книгу в букинистических магазинах. Вероятность того, что книга будет найдена первым студентом, равна " + p1 + ", а вторым " + p2 + ". Какова вероятность того, что: а) только один из студентов найдет книгу; б) оба студента найдут книгу; в) хотя бы один студент найдет книгу? " + Environment.NewLine;
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
                        int f1, f2, f3;
                        f1 = random.Next(10, 80);//вероятность
                        f2 = random.Next(2, 20);//всего
                        f3 = random.Next(2, 5);//сколько
                        double ff = f1 * 0.01;double ffff3 = Convert.ToDouble(f3);double raz = Convert.ToDouble(f2 - f3);double raz1 = Convert.ToDouble(1 - ff);
                        string s22 = "" +z+")" +f1 + " % деталей перед поступлением на сборку проходят термическую обработку. Найти вероятность того, что из " + f2 + " поступающих на сборку деталей; " + f3+" были термически обработаны." + Environment.NewLine;
                        doc.InsertParagraph(s22);
                        double ans22;
                        ans22 = Combinations(f3, f2) * (double)Math.Pow(ff, ffff3) * (double)Math.Pow(raz1, raz);
                        string s23 = "" + "Задание " + z + " - " + ans22 + Environment.NewLine;
                        doc2.InsertParagraph(s23);
                    }
                    else
                    {
                        int b1, b2, b3, b4;
                        b1 = random.Next(5, 15);//всего инженеров
                        b2 = random.Next(2, b1 - 3);//всего женщин
                        b3 = random.Next(2, b1 - 1);//в смене человек
                        b4 = random.Next(2, b1 - b2);//мужчин в смене
                        //while (b4 > b1 - b2) b4 = random.Next(2, b3);
                        string s24 = "" + z + ")На тепловой электростанции " + b1 + " сменных инженеров, из них " + b2 + " женщин. В смену занято " + b3 + " человека. Найти вероятность того, что в случайно выбранную смену окажется " + b4 + " мужчин." + Environment.NewLine;
                        doc.InsertParagraph(s24);
                        double ans23;
                        ans23 = (double)(Combinations(b4, b1 - b2) * Combinations(b3 - b4, b2)) / (Combinations(b3, b1));
                        string s25 = "" + "Задание " + z + " - " + ans23 + Environment.NewLine;
                        doc2.InsertParagraph(s25);
                    }
                }

                if (checkBox9.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {
                        int v1, v2, v3, v4, v5, v6;
                        v1 = random.Next(1, 100);//продукции с первой фабрики
                        v2 = random.Next(1, (100 - v1));//продукции со второй фабрики
                        v3 = random.Next(1, (100 - v1 - v2));//продукции с третьей фабрики
                        v4 = random.Next(1, 30);//процент нестандартных с первой фабрики
                        v5 = random.Next(1, v4);//процент нестандартных со второй фабрики
                        v6 = random.Next(1, v5);//процент нестандартных с третьей фабрики
                        string s26 = "" + z + ")На склад поступает продукция трёх фабрик. Причём продукция первой фабрики составляет " + v1 + " %, второй - " + v2 + " % и третьей - " + v3 + " %. Известно также, что средний процент нестандартных изделий для первой фабрики равен " + v4 + " %; для второй - " + v5 + " % и для третьей - " + v6 + " %. Найти вероятность того, что наудачу взятое изделие оказалось нестандартным." + Environment.NewLine;
                        doc.InsertParagraph(s26);
                        double ans24;
                        ans24 = (double)((v1 * 0.01) * (v4 * 0.01) + (v2 * 0.01) * (v5 * 0.01) + (v3 * 0.01) * (v6 * 0.01));
                        string s27 = "" + "Задание " + z + " - " + ans24 + Environment.NewLine;
                        doc2.InsertParagraph(s27);
                    }
                    else
                    {
                        //место для задачи партнера-Кристины!
                        //сделана!!
                        double v1, v2, v3, v4, v5, v6;
                        v1 = Math.Round(random.NextDouble() * 1, 1);//брак первого автомата
                        v2 = Math.Round(random.NextDouble() * 1, 1);//брак 2го автомата
                        v3 = Math.Round(random.NextDouble() * 1, 1);//брак 3го автомата
                        v4 = random.Next(500, 4000);//поступает с 1 завода
                        v5 = random.Next(500, 4000);//поступает со 2 завода
                        v6 = random.Next(500, 4000);//поступает с 3 завода
                        string s28 = "" + z + ")На сборку попадают детали с трёх автоматов. Известно, что первый автомат даёт " + v1 + "% брака, второй -" + v2 + "% и третий - " + v3 + "%. Найти вероятность попадания на сборку бракованной детали, если с первого автомата поступило" + v4 + ", со второго " + v5 + ", с третьего " + v6 + "." + Environment.NewLine;
                        doc.InsertParagraph(s28);
                        double ans25;
                        ans25 = (double)(((v4 / (v4 + v5 + v6) * (v1)) + ((v5 / (v4 + v5 + v6)) * v2) + ((v6 / (v4 + v5 + v6)) * v3)));
                        string s27 = "" + "Задание " + z + " - " + ans25 + Environment.NewLine;
                        doc2.InsertParagraph(s27);
                    }
                }

                if (checkBox10.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {
                        int t1, t2;
                        t1 = random.Next(1, 100);//детали первого автомата
                        t2 = random.Next(1, 100);//детали второго автомата
                        string s30 = "" + z + ")Два автомата производят одинаковые детали, которые сбрасываются на общий конвейер. Производительность первого автомата вдвое больше производительности второго. Первый автомат производит в среднем " + t1 + " % деталей отличного качества, а второй - " + t2 + " %. Наудачу взятая с конвейера деталь оказалась отличного качества. Найти вероятность того, что эта деталь произведена первым автоматом." + Environment.NewLine;
                        doc.InsertParagraph(s30);
                        double ans26;
                        ans26 = (double)(0.66666 * (t1 * 0.01)) / (0.66666 * (t1 * 0.01) + 0.33333 * (t2 * 0.01));
                        string s31 = "" + "Задание " + z + " - " + ans26 + Environment.NewLine;
                        doc2.InsertParagraph(s31);
                    }
                    else
                    {
                        //задача сделана
                        double pr1, pr2, pr3;
                        //место для задачи партнера-Кристины!
                        //string s32,33 ans 27
                        pr1 = random.Next(1, 95);//семян было обработано
                        pr2 = Math.Round(random.NextDouble() * 1, 1);//поражение обработанных
                        pr3 = Math.Round(random.NextDouble() * 1, 1);////поражение необработанных
                        string s32 = "" + z + ")Перед посевом " + pr1 + " % всех семян было обработано ядохимикатами. Вероятность поражения вредителями для растений из обработанных семян равна " + pr2 + ", для растений из необработанных семян - " + pr3 + ". Взятое наудачу растение оказалось пораженным. Какова вероятность того, что оно выращено из партии обработанных семян?" + Environment.NewLine;
                        doc.InsertParagraph(s32);
                        double ans27;
                        ans27 = (double)(((pr1 / 100) * pr3) / ((pr1 / 100) * pr3 + (1 - (pr1 / 100)) * pr2));
                        string s33 = "" + "Задание " + z + " - " + ans27 + Environment.NewLine;
                        doc2.InsertParagraph(s33);
                    }
                }

                if (checkBox11.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {
                        int y2, y3; double y1;
                        y1 = Math.Round(random.NextDouble() * 1, 1);
                        y2 = random.Next(2, 10);
                        y3 = random.Next(2, y2);
                        string s34 = "" + z + ")При каждом выстреле из орудия вероятность попадания в цель равна " + y1 + ". Найти вероятность того, что при " + y2 + " выстрелах будет " + y3 + " выстрела(-ов) мимо." + Environment.NewLine;
                        doc.InsertParagraph(s34);
                        double ans28;
                        ans28 = Combinations(y2 - y3, y2) * (double)Math.Pow(y1, y2 - y3) * (double)Math.Pow(1 - y1, y3);
                        string s35 = "" + "Задание " + z + " - " + ans28 + Environment.NewLine;
                        doc2.InsertParagraph(s35);
                    }
                    else
                    {
                        //место для задачи партнера-Кристины!
                        //Диана привет!!Задачка готова!! ясно С# не умеет считать дроби фу
                        double kos1, kos2, kos;
                        kos = random.Next(2, 7);
                        kos1 = random.Next(1, 7);
                        kos2 = random.Next(1, (int)kos);
                        string s36 = "" + z + ")Найти вероятность того, что при " + kos + " подбрасываниях игральной кости " + kos1 + " очков появится " + kos2 + " раз(-а)." + Environment.NewLine;
                        doc.InsertParagraph(s36);
                        double ans29;
                        ans29 = Combinations((int)kos2, (int)kos) * (double)Math.Pow(0.17, kos2) * (double)Math.Pow(0.83, kos - kos2);
                        string s37 = "" + "Задание " + z + " - " + ans29 + Environment.NewLine;
                        doc2.InsertParagraph(s37);
                    }
                }
                if (checkBox12.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {
                        int u2, u3; double u1;
                        u1 = Math.Round(random.NextDouble() * 0.9, 1);
                        if (u1 == 0) u1 += 0.1;
                        u2 = random.Next(2, 100);//всего испытаний
                        u3 = random.Next(2, u2);//сколько раз
                        string s38 = "" + z + ")Вероятность наступления события в каждом из одинаковых и независимых испытаний равна " + u1 + ".Найти вероятность того, что в " + u2 + " испытаниях событие наступит " + u3 + " раз." + Environment.NewLine;
                        doc.InsertParagraph(s38);
                        double ans30;
                        double vx = (double)((u3 - (u2 * u1)) / (double)(Math.Sqrt(u2 * u1 * (1 - u1))));
                        ans30 = 1 / ((double)Math.Sqrt(u2 * u1 * (1 - u1)));
                        string s39 = "" + "Задание " + z + " - ф(" + vx + ")" + ans30 + Environment.NewLine;
                        doc2.InsertParagraph(s39);
                    }
                    else
                    {
                        //Диана у меня перемнных ховут вера....
                        double vera1, vera2, vera3;
                        vera1 = Math.Round(random.NextDouble() * 0.9, 1);
                        if (vera1 == 0) vera1 += 0.1;
                        vera2 = random.Next(1, 100);//всего выстрелов
                        vera3 = random.Next(1, (int)vera2);//сколько раз
                        string s40 = "" + z + ")Вероятность поражения мишени при одном выстреле равна " + vera1 + ". Найти вероятность того, что при " + vera2 + " выстреле(-ах) мишень будет поражена " + vera3 + " раз." + Environment.NewLine;
                        doc.InsertParagraph(s40);
                        double ans31;
                        double vx = (double)((vera3 - (vera1 * vera2)) / (double)(Math.Sqrt(vera2 * vera1 * (1 - vera1))));
                        ans31 = 1 / ((double)Math.Sqrt(vera2 * vera1 * (1 - vera1)));
                        string s41 = "" + "Задание " + z + " - ф(" + vx + ")*" + ans31 + Environment.NewLine;
                        doc2.InsertParagraph(s41);
                    }
                }

                if (checkBox13.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {
                        int a1, a2, a3;
                        a1 = random.Next(5, 95);
                        a2 = random.Next(2, 100);
                        a3 = random.Next(2, a2);
                        string s42 = "" + z + ")На склад поступают изделия, из которых " + a1 + " оказываются высшего сорта. Найти вероятность того, что из " + a2 + " взятых наудачу не менее " + a3 + " изделий окажется высшего сорта" + Environment.NewLine;
                        doc.InsertParagraph(s42);
                        double ans32, ans33;
                        double ll = (double)(a1 * 0.01);
                        ans32 = (double)(((a3 - (ll * a2))) / (Math.Sqrt((double)(a3 * ll * (1 - ll)))));
                        ans33 = (double)(((a3 - (ll * a2))) / (Math.Sqrt((double)(a3 * ll * (1 - ll)))));
                        string s43 = "" + "Задание " + z + " - Ф(" + ans32 + ")-Ф(" + ans33 + ")" + Environment.NewLine;
                        doc2.InsertParagraph(s43);
                    }
                    else
                    {
                        //место для задачи партнера-Кристины!
                        //string s44,45 ans 34
                        int a1, a2, a3, a4;
                        a1 = random.Next(5, 95);//p
                        a2 = random.Next(700, 1500);//n
                        a3 = random.Next(200, a2 - 100);//a
                        a4 = random.Next(a3 + 1, a2);//b
                        string s44 = "" + z + ")Всхожесть семян составляет " + a1 + "%. Какова вероятность того. Что из " + a2 + " посеянных семян взойдут от " + a3 + " до " + a4 + "?" + Environment.NewLine;
                        doc.InsertParagraph(s44);
                        double ans34, ans35;
                        double ll = (double)(a1 * 0.01);
                        ans34 = (double)(((a3 - (ll * a2))) / (double)(Math.Sqrt((double)(a2 * ll * (1 - ll)))));
                        ans35 = (double)(((a4 - (ll * a2))) / (double)(Math.Sqrt((double)(a2 * ll * (1 - ll)))));
                        string s45 = "" + "Задание " + z + " - Ф(" + ans34 + ")-Ф(" + ans35 + ")" + Environment.NewLine;
                        doc2.InsertParagraph(s45);
                    }
                }


                if (checkBox14.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {
                        int s1; double s2, s3;
                        s1 = random.Next(20, 401);
                        s2 = Math.Round(random.NextDouble() * 1, 1);
                        s3 = Math.Round(random.NextDouble() * 1, 2);
                        string s46 = "" + z + ")В автопарке имеется " + s1 + " автомобиль(-ей). Вероятность безотказной работы каждого из них равна " + s2 + ". С вероятностью " + s3 + " определить границы, в которых будет находиться доля безотказно работавших машин в определенный момент времени." + Environment.NewLine;
                        doc.InsertParagraph(s46);
                        double qq, ww, rr;
                        qq = s1 * s2; ww = (double)((s3 + 1) / 2); rr = (double)Math.Sqrt((double)(s1 * s2 * (1 - s2)));
                        string s47 = "" + "Задание " + z + "- " + qq + "-Ф(" + ww + ")*" + rr + "<=X<=" + qq + "+Ф(" + ww + ")*" + rr + Environment.NewLine;
                        doc2.InsertParagraph(s47);
                    }
                    else
                    {
                        //место для задачи партнера-Кристины!
                        //string s48,49 ans 35
                        int j1; double j2, j3;//сколько
                        j1 = random.Next(20, 401);
                        j2 = Math.Round(random.NextDouble() * 1, 1);//вероятность общая
                        j3 = Math.Round(random.NextDouble() * 1, 2);//граница
                        string s48 = "" + z + ")Вероятность выплавки стабильного сплава в дуговой вакуумной установке равна " + j2 + " в каждой отдельной плавке. Произведена(-о) " + j1 + " плавка(-ок). Найти вероятность того, что относительная частота выплавки стабильного сплава отклонится от вероятности не более чем на " + j3 + " ." + Environment.NewLine;
                        doc.InsertParagraph(s48);
                        double qq, ww, rr;
                        qq = j1 * j2; ww = (double)((j3 + 1) / 2); rr = (double)Math.Sqrt((double)(j1 * j2 * (1 - j2)));
                        string s49 = "" + "Задание " + z + "- " + qq + "-Ф(" + ww + ")*" + rr + "<=X<=" + qq + "+Ф(" + ww + ")*" + rr + Environment.NewLine;
                        doc2.InsertParagraph(s49);
                    }
                }

                if (checkBox15.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {
                        //я ее пока еще не делала
                        //string 50,51 ans 36
                        int x1, x3; double x2;
                        x1 = random.Next(10, 500);//всего
                        x3 = random.Next(1, 11);//отбираются
                        x2 = Math.Round(random.NextDouble() * 0.01, 4);
                        string s50 = "" + z + ")Судно перевозит " + x1 + " упаковок доброкачественного груза. Вероятность того, что в рейсе любая упаковка повредится, равна " + x2 + " . Найти вероятность того, что в порт назначения будет доставлен груз с " + x3 + " упаковками испорченного груза." + Environment.NewLine;
                        doc.InsertParagraph(s50);
                        double lam = Convert.ToDouble(x1 * x2);
                        double xx3 = Convert.ToDouble(x3);
                        double stepen = Math.Pow(lam, xx3);
                        int mfact = Factorial(x3);
                        string s51 = "" + "Задание " + z + "- " + " ( " + stepen + " *e^(-" + lam + "))/" + mfact + Environment.NewLine;
                        doc2.InsertParagraph(s51);
                    }
                    else
                    {
                        //место для задачи партнера-Кристины!
                        //string 52,53 ans 37
                        int vv1, vv3; double vv2;
                        vv1 = random.Next(10, 500);//всего
                        vv3 = random.Next(1, 11);//отбираются
                        vv2 = Math.Round(random.NextDouble() * 1, 3);//вероятность
                        string s52 = "" + z + ")Вероятность того, что человек в период страхования будет травмирован, равна " + vv2 + ". Компанией застраховано " + vv1 + " человек. Какова вероятность того, что травму получат " + vv3 + " человек?" + Environment.NewLine;
                        doc.InsertParagraph(s52);
                        double lam2 = Convert.ToDouble(vv1 * vv2);
                        double xx32 = Convert.ToDouble(vv3);
                        double stepen = Math.Pow(lam2, xx32);
                        int mfact2 = Factorial(vv3);
                        string s53 = "" + "Задание " + z + "- " + " ( " + stepen + " *e^(-" + lam2 + "))/" + mfact2 + Environment.NewLine;
                        doc2.InsertParagraph(s53);
                    }
                }

                if (checkBox16.Checked == true)
                {
                    z++;
                    if (k % 2 != 0)
                    {
                        double d1; int d2;
                        d1 = Math.Round(random.NextDouble() * 1, 1);
                        d2 = random.Next(15, 75);
                        string s54 = "" + z + ")Вероятность получения студентом отличной оценки на экзамене равна " + d1 + ". Найти наивероятнейшее число отличных оценок и вероятность этого числа, если число студентов, сдающих экзамен равно " + d2+" ." + Environment.NewLine;
                        doc.InsertParagraph(s54);
                        double ans38, ans39;
                        ans38 = d1 * 0.01; ans39 = ans38 - 1;
                        string s55 = "" + "Задание " + z + "- " + ans39 + "<=k0<=" + ans38 + Environment.NewLine;
                        doc2.InsertParagraph(s55);
                    }
                    else
                    {
                        //место для задачи партнера-Кристины!
                        //string s44,45 ans 31
                    }
                }

                //Xceed.Document.NET.Paragraph par = doc.InsertParagraph("Таблица Гаусса:");
                //par.AppendPicture(p);

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

        private void checkBox17_CheckedChanged(object sender, EventArgs e)
        {


            string fileName = @"C:\Users\Дианочка\source\repos\GenerationVariants\ex.docx";
            string fileName2 = @"C:\Users\Дианочка\source\repos\GenerationVariants\ans.docx";
            var doc = DocX.Create(fileName);
            var doc2 = DocX.Create(fileName2);
            int k;//z;
            Random random = new Random();
            for (k = 1; k <= n; k++)
            {
                if (checkBox17.Checked == true)
                {
                    string s = "" + "Вариант №" + k + Environment.NewLine;
                    doc.InsertParagraph(s);
                    doc2.InsertParagraph(s);
                    {
                        if (k % 2 != 0)
                        {
                            int h, i = 8;
                            h = random.Next(2, 10);
                            string s2 = "" + "1)Сколько имеется чисел c" + " " + h + " " + "знаками, все цифры у которых различны?" + Environment.NewLine;
                            doc.InsertParagraph(s2);
                            int ans = 9 * 9;
                            while (h != 2)
                            {
                                ans *= i;
                                i--;
                                h--;
                            }
                            string s3 = "" + "Задание 1 " + " -" + ans + Environment.NewLine;
                            doc2.InsertParagraph(s3);
                        }
                        else
                        {
                            double h1, p2;
                            h1 = random.Next(2, 17);//этажи
                            p2 = random.Next(2, 11);//люди
                            string s4 = "" + "2)" + p2 + " человек(-a) вошли в лифт на 1-м этаже" + " дома с" + " " + h1 + " этажами. Сколькими способами пассажиры могут выйти из лифта на нужных этажах?" + Environment.NewLine;
                            doc.InsertParagraph(s4);
                            double ans1 = (int)Math.Pow(h1 - 1, p2);
                            string s5 = "" + "Задание 2 - " + ans1 + Environment.NewLine;
                            doc2.InsertParagraph(s5);
                        }
                    }

                    {
                        if (k % 2 != 0)
                        {
                            int m1, w1, m2, w2, g;
                            m1 = random.Next(2, 16);//всего в группе мужчин
                            w1 = random.Next(2, 16);//женшин в группе
                            g = random.Next(2, m1 + w1);//группа для ужина
                            w2 = random.Next(2, w1);//в группе2 женщин
                            m2 = random.Next(2, m1);//в группе2 мужчин
                            string s5 = "" + "2)Группа туристов из " + m1 + " юношей и " + w1 + " девушек выбирает по жребию " + g + " человек(-a) для приготовления ужина. Сколько существует способов, при которых в эту группу попадут " + w2 + " девушек или " + m2 + " юношей?" + Environment.NewLine;
                            doc.InsertParagraph(s5);
                            double ans5;
                            ans5 = Combinations(w2, w1) + Combinations(m2, m1);
                            string s6 = "" + "Задание 2 " + " - " + ans5 + Environment.NewLine;
                            doc2.InsertParagraph(s6);
                        }
                        else
                        {
                            int m1, w1, w2, g;
                            m1 = random.Next(2, 16);//всего деталей
                            w1 = random.Next(2, m1);//бракованные
                            g = random.Next(2, w1);//комплект
                            w2 = random.Next(2, g);//бракованные выпадают
                            string s7 = "" + "2)В ящике " + m1 + " детали(-ей), среди которых  " + w1 + " бракованных. Наудачу выбирается комплект из  " + g + " деталей. Сколько всего комплектов, в каждом из которых " + w2 + " детали(-ей) бракованные(-ых)?" + Environment.NewLine;
                            doc.InsertParagraph(s7);
                            double ans9;
                            ans9 = Combinations(w2, w1) * Combinations(g - w2, m1 - w1);
                            string s6 = "" + "Задание 2 " + " - " + ans9 + Environment.NewLine;
                            doc2.InsertParagraph(s6);
                        }
                    }

                    {
                        if (k % 2 != 0)
                        {
                            int s1, s2;
                            s1 = random.Next(2, 11);
                            s2 = random.Next(2, s1);
                            string s9 = "" + "3)В комнате имеется " + s1 + " стульев. Сколькими способами можно разместить на них " + s2 + " гостей?" + Environment.NewLine;
                            doc.InsertParagraph(s9);
                            int ans7 = 1, c = 0;
                            while (c != s2)
                            {
                                ans7 *= s1;
                                s1--;
                                c++;
                            }
                            string s10 = "" + "Задание 3 - " + ans7 + Environment.NewLine;
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
                            string s11 = "" + "3)Сколько различных «слов» можно получить, переставляя буквы в слове " + l + "?" + Environment.NewLine;
                            doc.InsertParagraph(s11);
                            int ans8 = Factorial(5) / 1;
                            string s12 = "" + "Задание 3 -" + ans8 + Environment.NewLine;
                            doc2.InsertParagraph(s12);
                        }
                    }

                    {
                        if (k % 2 != 0)
                        {
                            int r1; string l;
                            r1 = random.Next(1, 4);
                            if (r1 == 1) l = "ПРОФЕССОР";
                            else
                                if (r1 == 2) l = "АВИАЛИНИЯ";
                            else
                                l = "БАДМИНТОН";
                            string s13 = "" + "4)Сколько различных «слов» можно получить, переставляя буквы в слове " + l + "?" + Environment.NewLine;
                            doc.InsertParagraph(s13);
                            int ans10;
                            if (r1 == 1) ans10 = 45360;
                            else if (r1 == 2) ans10 = 30240;
                            else ans10 = 181440;
                            string s14 = "" + "Задание 4 -" + ans10 + Environment.NewLine;
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
                            string s15 = "" + "4)Сколько различных «слов» можно получить, переставляя буквы в слове " + l + "?" + Environment.NewLine;
                            doc.InsertParagraph(s15);
                            int ans11;
                            if (r2 == 1) ans11 = 9979200;
                            else if (r2 == 2) ans11 = 1663200;
                            else ans11 = 19958400;
                            string s16 = "" + "Задание 4 - " + ans11 + Environment.NewLine;
                            doc2.InsertParagraph(s16);
                        }
                    }

                    {
                        if (k % 2 != 0)
                        {
                            double t1, t2, t3;
                            t1 = random.Next(3, 6);
                            t2 = random.Next(3, 10);
                            t3 = random.Next(1, (int)t2);
                            string s13 = "" + "5)В контрольной работе будет " + t1 + " задач(-и) – по одной из каждой пройденной темы. Задачи будут взяты из общего списка по " + t2 + " задач(-и) в каждой теме, а всего было пройдено " + t1 + " тем(-ы). При подготовке к контрольной Вова решил только по " + t3 + " задач(-е) в каждой теме. Найдите общее число всех возможных вариантов контрольной работы." + Environment.NewLine;
                            doc.InsertParagraph(s13);
                            double ans12;
                            ans12 = (int)Math.Pow(t2, t1);
                            string s14 = "" + "Задание 5 - " + ans12 + Environment.NewLine;
                            doc2.InsertParagraph(s14);
                        }
                        else
                        {
                            int g1;
                            g1 = random.Next(4, 8);
                            string s15 = "" + "5)В футбольном турнире участвуют несколько команд. Оказалось, что все они для трусов и футболок использовали " + g1 + " цвета(-ов), причем были представлены все возможные варианты. Сколько команд участвовали в турнире?" + Environment.NewLine;
                            doc.InsertParagraph(s15);
                            int ans13;
                            ans13 = g1 * (g1 - 1);
                            string s16 = "" + "Задание 5 - " + ans13 + Environment.NewLine;
                            doc2.InsertParagraph(s16);
                        }
                    }

                    {
                        if (k % 2 != 0)
                        {
                            double h1;
                            h1 = random.Next(2, 13);
                            string s17 = "" + "6)Бросаются две игральные кости. Определить вероятность того, что: а) сумма числа очков не превосходит " + h1 + "; б) произведение числа очков не превосходит " + h1 + "; в) произведение числа очков делится на " + h1 + "." + Environment.NewLine;
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
                            string s14 = "" + "Задание 6 - " + "a)" + ans13 + ";б)" + ans14 + ";в)" + ans15 + Environment.NewLine;
                            doc2.InsertParagraph(s14);
                        }
                        else
                        {
                            double h1;
                            h1 = random.Next(2, 13);
                            string s18 = "" + "6)Бросаются две игральные кости. Определить вероятность того, что: а) сумма числа очков не превосходит " + h1 + "; б) произведение числа очков не превосходит " + h1 + "; в) произведение числа очков делится на " + h1 + "." + Environment.NewLine;
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
                            string s19 = "" + "Задание 6 - " + "a)" + ans13 + ";б)" + ans14 + ";в)" + ans15 + Environment.NewLine;
                            doc2.InsertParagraph(s19);
                        }
                    }

                    {
                        if (k % 2 != 0)
                        {
                            double p1, p2;
                            p1 = Math.Round(random.NextDouble() * 1, 1);
                            p2 = Math.Round(random.NextDouble() * 1, 1);
                            string s20 = "" + "7)Два студента ищут нужную им книгу в букинистических магазинах. Вероятность того, что книга будет найдена первым студентом, равна " + p1 + ", а вторым " + p2 + ". Какова вероятность того, что: а) только один из студентов найдет книгу; б) оба студента найдут книгу; в) хотя бы один студент найдет книгу? " + Environment.NewLine;
                            doc.InsertParagraph(s20);
                            double ans16, ans17, ans18;
                            ans16 = p1 * (1 - p2) + p2 * (1 - p1);
                            ans17 = p1 * p2;
                            ans18 = (1 - ((1 - p1) * (1 - p2)));
                            string s21 = "" + "Задание 7 - " + "a)" + ans16 + ";б)" + ans17 + ";в)" + ans18 + Environment.NewLine;
                            doc2.InsertParagraph(s21);
                        }
                        else
                        {
                            double p1;
                            int c1, c2; ;
                            p1 = Math.Round(random.NextDouble() * 1, 1);
                            c1 = random.Next(2, 6);
                            c2 = random.Next(2, c1);
                            string s20 = "" + "7)Вероятность выигрыша по лотерейному билету " + p1 + ". Приобретено " + c1 + " билета. Какова вероятность того, что выигрыша: а) только по одному из купленных билетов; б) только по " + c2 + " из купленных билетов; в) хотя бы по одному билету?" + Environment.NewLine;
                            doc.InsertParagraph(s20);
                            double ans19, ans20, ans21;
                            ans19 = Combinations(1, c1) * (double)Math.Pow(p1, 1) * (double)Math.Pow((1 - p1), c1 - 1);
                            ans20 = Combinations(c2, c1) * (double)Math.Pow(p1, c2) * (double)Math.Pow((1 - p1), c1 - c2);
                            ans21 = 1 - (Combinations(0, c1) * (double)Math.Pow(p1, 0) * (double)Math.Pow((1 - p1), c1));
                            string s21 = "" + "Задание 7 - " + "a)" + ans19 + ";б)" + ans20 + ";в)" + ans21 + Environment.NewLine;
                            doc2.InsertParagraph(s21);
                        }
                    }



                    {
                        if (k % 2 != 0)
                        {
                            int f1, f2, f3;
                            f1 = random.Next(10, 100);//вероятность
                            f2 = random.Next(10, 50);//всего
                            f3 = random.Next(10, f2);//сколько
                            double ff = f1 * 0.01;
                            string s22 = "" + "8)" + f1 + " % деталей перед поступлением на сборку проходят термическую обработку. Найти вероятность того, что из " + f2 + " поступающих на сборку деталей; " + f3 + " были термически обработаны." + Environment.NewLine;
                            doc.InsertParagraph(s22);
                            double ans22;
                            ans22 = Combinations(f2 - f3, f2) * (double)Math.Pow(ff, f2 - f3) * (double)Math.Pow(1 - ff, f3);
                            string s23 = "" + "Задание 8 - " + ans22 + Environment.NewLine;
                            doc2.InsertParagraph(s23);
                        }
                        else
                        {
                            int b1, b2, b3, b4;
                            b1 = random.Next(3, 10);//всего инженеров
                            b2 = random.Next(2, b1);//всего женщин
                            b3 = random.Next(2, b1 - 1);//в смене человек
                            b4 = random.Next(2, b1 - b2);//мужчин в смене
                                                         //while (b4 > b1 - b2) b4 = random.Next(2, b3);
                            string s24 = "" +  "8)На тепловой электростанции " + b1 + " сменных инженеров, из них " + b2 + " женщин. В смену занято " + b3 + " человек(-а). Найти вероятность того, что в случайно выбранную смену окажется " + b4 + " мужчин." + Environment.NewLine;
                            doc.InsertParagraph(s24);
                            double ans23;
                            ans23 = (double)(Combinations(b4, b1 - b2) * Combinations(b3 - b4, b2)) / (Combinations(b3, b1));
                            string s25 = "" + "Задание 8 - " + ans23 + Environment.NewLine;
                            doc2.InsertParagraph(s25);
                        }
                    }


                }

            }

            doc.Save();
            doc2.Save();
            Process.Start("WINWORD.EXE", fileName);
            Process.Start("WINWORD.EXE", fileName2);
        }
    }
}

