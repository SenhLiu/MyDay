using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;

namespace MyDay
{
    public partial class MyDay : Form
    {
        public static string APP_NAME = "MyDay";

        private DateTime theDate;
        private DateTime today;
        private Day[] calendar;

        private AboutBox aboutBox = null;
        private SearchForm searchForm = null;
        private ExportForm exportForm = null;

        class Day
        {
            public bool isBold;
            public bool isItalic;
            public bool isStrike;

            public int value;
            public string text;

            public Day(string text)
            {
                this.text = text;
            }

            public Day(int value)
            {
                this.value = value;
            }
        }

        public MyDay()
        {
            InitializeComponent();
            LoadDimensions(APP_NAME);
            today = DateTime.Now;
            theDate = today;
            updateLabels();
            updateCalendar(true);
        }

        private void updateLabels()
        {
            lbToday.Text = theDate.ToString("dddd d MMMM [M], yyyy");
            cbMonth.Text = theDate.ToString("MMMM");
            txtYear.Value = theDate.Year;
        }

        /* Determination of the day of the week

        Jan 1st 1 AD is a Monday in Gregorian calendar.
        So Jan 0th 1 AD is a Sunday [It does not exist technically].

        Every 4 years we have a leap year. But xy00 cannot be a leap unless xy divides 4 with reminder 0.
        y/4 - y/100 + y/400 : this gives the number of leap years from 1AD to the given year. As each year has 365 days (divdes 7 with reminder 1), unless it is a leap year or the date is in Jan or Feb, the day of a given date changes by 1 each year. In other case it increases by 2.
        y -= m<3 : If the month is not Jan or Feb, we do not count the 29th Feb (if it exists) of the given year. 
        So y + y/4 - y/100 + y/400  gives the day of Jan 0th (Dec 31st of prev year) of the year. (This gives the reminder with 7 of  the number of days passed before the given year began.)

        Array t:  Number of days passed before the month 'm+1' begins.

        So t[m-1]+d is the number of days passed in year 'y' upto the given date. 
        (y + y/4 - y/100 + y/400 + t[m-1] + d) % 7 is reminder of the number of days from Jan 0  1AD to the given date which will be the day (0=Sunday,6=Saturday).

        Description credits: Sai Teja Pratap (quora.com/How-does-Tomohiko-Sakamotos-Algorithm-work).
        */
        static int DayOfWeek(int y, int m, int d)
        {
            int[] t = { 0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4 };
            if (m < 3)
                y--;
            return (y + y / 4 - y / 100 + y / 400 + t[m - 1] + d) % 7;

            //  string[] days = {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"};
        }

        // get the first dow for the month
        int getFirstDow(int year, int month)
        {
            var dow = DayOfWeek(year, month, 1);
            // wrap around so monday is the first day of the week instead of sunday
            if (dow > 0)
                return dow - 1;
            else
                return 6;            
        }

        private int getDaysInMonth(int year, int month)
        {            
            //DateTime.IsLeapYear(year))
            return DateTime.DaysInMonth(year, month);            
        }

        private Button getDateButton(int day)
        {
            Button[] buttons = { d0,  d1,  d2,  d3,  d4,  d5,  d6,
                                 d7,  d8,  d9,  d10, d11, d12, d13,
                                 d14, d15, d16, d17, d18, d19, d20,
                                 d21, d22, d23, d24, d25, d26, d27,
                                 d28, d29, d30, d31, d32, d33, d34,
                                 d35, d36, d37, d38, d39, d40, d41};
            return buttons[day];
        }        

        private bool isWeekend(int year, int month, int day)
        {
            var dow = DayOfWeek(year, month, day);
            return (dow == 0 || dow == 6);            
        }

        private bool isToday(int year, int month, int day)
        {
            return (day == today.Day && month == today.Month && year == today.Year);
        }

        private string LZ(int m)
        {
            if (m < 10)
                return "0" + m;
            return Convert.ToString(m);
        }

        private string xmldecode(string s)
        {                        
            s = s.Replace("&gt;", ">");
            s = s.Replace("&lt;", "<");
            s = s.Replace("&amp;", "&");
            return s;
        }

        private string xmlencode(string s)
        {            
            s = s.Replace("&", "&amp;");
            s = s.Replace("<", "&lt;");
            s = s.Replace(">", "&gt;");
            return s;
        }

        // escape text for button label
        private string escape(string s)
        {
            return s.Replace("&", "&&");
        }

        private Day[] loadCalendar(int year, int month)
        {
            var dict = new Day[34];

            bool isBold = false;
            bool isItalic = false;
            bool isStrike = false;

            var filename = "data\\" + year + LZ(month) + ".xml";
            if (File.Exists(filename))
            {
                bool inText = false;
                StringBuilder text = new StringBuilder();
                int day = 0;
                string[] lines = File.ReadAllLines(filename);
                foreach (var line in lines)
                {
                    var s = line.Trim();
                    if (s.StartsWith("<Day DayValue=\""))
                    {
                        isBold = s.Contains("bold=\"True\"");
                        isItalic = s.Contains("italic=\"True\"");
                        isStrike = s.Contains("strike=\"True\"");

                        s = s.Substring(15);
                        int p = s.IndexOf('"');
                        if (p != -1)
                            s = s.Substring(0, p);
                        day = Convert.ToInt32(s);
                    }
                    else if (s.StartsWith("<Text>"))
                    {
                        s = s.Substring(6);
                        if (s.StartsWith("-"))
                            s = s.Substring(1);
                        if (s.EndsWith("</Text>"))
                        {
                            var dd = new Day(xmldecode(s.Substring(0, s.Length - 7)));
                            dd.isBold = isBold;
                            dd.isItalic = isItalic;
                            dd.isStrike = isStrike;
                            dict[day] = dd; 
                        }
                        else
                        {
                            inText = true;
                            text.Append(s);
                        }
                    } else if (s.EndsWith("</Text>"))
                    {
                        inText = false;
                        text.Append("\n");
                        text.Append(s.Substring(0, s.Length - 7));
                        var dd = new Day(xmldecode(text.ToString()));
                        dd.isBold = isBold;
                        dd.isItalic = isItalic;
                        dd.isStrike = isStrike;
                        dict[day] = dd;
                        text.Clear();
                    } else if (inText)
                    {
                        text.Append("\n");
                        text.Append(s);
                    }
                }
            }

            return dict;
        }

        private void reloadCalendar(int year, int month)
        {
            if (calendar == null)
            {
                calendar = loadCalendar(year, month);
            }
            else
            {
                if (calendar[32].value == year && calendar[33].value == month)
                    return;

                calendar = loadCalendar(year, month);
            }

            calendar[32] = new Day(year);
            calendar[33] = new Day(month);
        }

        private bool isEmptyCalender()
        {
            for (int n = 1; n <= 31; n++)
                if (calendar[n] != null && !String.IsNullOrEmpty(calendar[n].text.Trim()))
                    return false;
            return true;
        }

        private void saveCalendar()
        {
            if (calendar != null)
            {                
                int year = calendar[32].value;
                int month = calendar[33].value;

                var filename = "data\\" + year + LZ(month) + ".xml";

                if (!isEmptyCalender())
                {
                    if (!Directory.Exists("data"))
                        Directory.CreateDirectory("data");
                    using (var file = new StreamWriter(filename, false))
                    {
                        file.WriteLine("<Year YearValue=\"" + year + "\">");
                        file.WriteLine("\t<Month MonthValue=\"" + month + "\">");
                        for (int n = 1; n <= 31; n++)
                            if (calendar[n] != null)
                            {
                                var dd = calendar[n];
                                var s = dd.text.Trim();
                                if (!String.IsNullOrEmpty(s))
                                {
                                    string format = "";
                                    if (dd.isBold)
                                        format += " bold=\"True\"";
                                    if (dd.isItalic)
                                        format += " italic=\"True\"";
                                    if (dd.isStrike)
                                        format += " strike=\"True\"";
                                    file.WriteLine("\t\t<Day DayValue=\"" + n + "\"" + format + ">");
                                    file.WriteLine("\t\t\t<Text>-" + xmlencode(s) + "</Text>");
                                    file.WriteLine("\t\t</Day>");
                                }
                            }
                        file.WriteLine("\t</Month>");
                        file.WriteLine("</Year>");
                    }
                }
                else
                {
                    if (File.Exists(filename))
                        File.Delete(filename);
                }
            }
        }

        private void disableDateButton(int n)
        {
            Button but = getDateButton(n);
            but.Text = "";
            but.BackColor = Color.White;
            but.FlatAppearance.BorderSize = 0;
            but.Enabled = false;
        }

        private void updateCalendar(bool updateText)
        {
            // get the first day of week for the month
            int year = theDate.Year;
            int month = theDate.Month;

            // load 
            reloadCalendar(year, month);

            int dow = getFirstDow(year, month);
            int dim = getDaysInMonth(year, month);

            var weekdayColor = Color.FromArgb(192, 192, 255);
            var weekendColor = Color.FromArgb(255, 224, 192);
            var selectedColor = Color.FromArgb(128, 128, 255);
            var todayColor = Color.FromArgb(255, 192, 255);

            toolTip1.RemoveAll();

            for (int n = 0; n < dow; n++)            
                disableDateButton(n);

            var toolTip = new ToolTip();
            
            for (int n = 1; n <= dim; n++)
            {
                Button but = getDateButton(dow + n - 1);
                if (calendar[n] != null && !String.IsNullOrEmpty(calendar[n].text))
                {
                    var dd = calendar[n];
                    but.Text = n + "-" + escape(dd.text);
                    if (but.Font.Bold != dd.isBold || but.Font.Italic != dd.isItalic || but.Font.Strikeout != dd.isStrike)
                    {
                        FontStyle fs = FontStyle.Regular;
                        if (dd.isBold)
                            fs |= FontStyle.Bold;
                        if (dd.isItalic)
                            fs |= FontStyle.Italic;
                        if (dd.isStrike)
                            fs |= FontStyle.Strikeout;
                        var ff = new Font(but.Font.FontFamily, but.Font.Size, fs);
                        but.Font = ff;
                    }

                    toolTip1.SetToolTip(but, dd.text);
                }
                else
                {
                    but.Text = Convert.ToString(n);

                    if (but.Font.Bold || but.Font.Italic || but.Font.Strikeout)
                    {
                        FontStyle fs = FontStyle.Regular;                        
                        var ff = new Font(but.Font.FontFamily, but.Font.Size, fs);
                        but.Font = ff;
                    }
                }
                if (n == theDate.Day)
                    but.BackColor = selectedColor;
                else if (isToday(year, month, n))
                    but.BackColor = todayColor;
                else if (isWeekend(year, month, n))
                    but.BackColor = weekendColor;
                else
                    but.BackColor = weekdayColor;
                but.FlatAppearance.BorderSize = 1;
                but.FlatAppearance.BorderColor = Color.White;
                but.Enabled = true;
            }
            for (int n = dim; n + dow <= 41; n++)
                disableDateButton(dow + n);

            if (updateText)
            {
                if (calendar[theDate.Day] != null)
                {
                    var dd = calendar[theDate.Day];                    
                    textBox.Text = dd.text;
                    
                    // need to update font after text to avoid extra text update
                    if (textBox.Font.Bold != dd.isBold || textBox.Font.Italic != dd.isItalic || textBox.Font.Strikeout != dd.isStrike)
                    {
                        FontStyle fs = FontStyle.Regular;
                        if (dd.isBold)
                            fs |= FontStyle.Bold;
                        if (dd.isItalic)
                            fs |= FontStyle.Italic;
                        if (dd.isStrike)
                            fs |= FontStyle.Strikeout;
                        var ff = new Font(textBox.Font.FontFamily, textBox.Font.Size, fs);
                        textBox.Font = ff;
                    }                    
                }
                else
                {
                    textBox.Text = "";

                    // need to update font after text to avoid extra text update
                    if (textBox.Font.Bold || textBox.Font.Italic || textBox.Font.Strikeout)
                    {
                        FontStyle fs = FontStyle.Regular;
                        var ff = new Font(textBox.Font.FontFamily, textBox.Font.Size, fs);
                        textBox.Font = ff;
                    }                    
                }
            }
        }

        private void selectDate(int day)
        {
            int year = theDate.Year;
            int month = theDate.Month;
            int dow = getFirstDow(year, month);
            theDate = new DateTime(year, month, day - dow + 1);
            updateLabels();
            updateCalendar(true);
            textBox.Focus();
        }

        private void d0_Click(object sender, EventArgs e)
        {
            selectDate(0);
        }

        private void d1_Click(object sender, EventArgs e)
        {
            selectDate(1);
        }

        private void d2_Click(object sender, EventArgs e)
        {
            selectDate(2);
        }

        private void d3_Click(object sender, EventArgs e)
        {
            selectDate(3);
        }

        private void d4_Click(object sender, EventArgs e)
        {
            selectDate(4);
        }

        private void d5_Click(object sender, EventArgs e)
        {
            selectDate(5);
        }

        private void d6_Click(object sender, EventArgs e)
        {
            selectDate(6);
        }

        private void d7_Click(object sender, EventArgs e)
        {
            selectDate(7);
        }

        private void d8_Click(object sender, EventArgs e)
        {
            selectDate(8);
        }

        private void d9_Click(object sender, EventArgs e)
        {
            selectDate(9);
        }

        private void d10_Click(object sender, EventArgs e)
        {
            selectDate(10);
        }

        private void d11_Click(object sender, EventArgs e)
        {
            selectDate(11);
        }

        private void d12_Click(object sender, EventArgs e)
        {
            selectDate(12);               
        }

        private void d13_Click(object sender, EventArgs e)
        {
            selectDate(13);
        }

        private void d14_Click(object sender, EventArgs e)
        {
            selectDate(14);
        }

        private void d15_Click(object sender, EventArgs e)
        {
            selectDate(15);
        }

        private void d16_Click(object sender, EventArgs e)
        {
            selectDate(16);
        }

        private void d17_Click(object sender, EventArgs e)
        {
            selectDate(17);
        }

        private void d18_Click(object sender, EventArgs e)
        {
            selectDate(18);
        }

        private void d19_Click(object sender, EventArgs e)
        {
            selectDate(19);
        }

        private void d20_Click(object sender, EventArgs e)
        {
            selectDate(20);
        }

        private void d21_Click(object sender, EventArgs e)
        {
            selectDate(21);
        }

        private void d22_Click(object sender, EventArgs e)
        {
            selectDate(22);
        }

        private void d23_Click(object sender, EventArgs e)
        {
            selectDate(23);
        }

        private void d24_Click(object sender, EventArgs e)
        {
            selectDate(24);
        }

        private void d25_Click(object sender, EventArgs e)
        {
            selectDate(25);
        }

        private void d26_Click(object sender, EventArgs e)
        {
            selectDate(26);
        }

        private void d27_Click(object sender, EventArgs e)
        {
            selectDate(27);
        }

        private void d28_Click(object sender, EventArgs e)
        {
            selectDate(28);
        }

        private void d29_Click(object sender, EventArgs e)
        {
            selectDate(29);
        }

        private void d30_Click(object sender, EventArgs e)
        {
            selectDate(30);
        }

        private void d31_Click(object sender, EventArgs e)
        {
            selectDate(31);
        }

        private void d32_Click(object sender, EventArgs e)
        {
            selectDate(32);
        }

        private void d33_Click(object sender, EventArgs e)
        {
            selectDate(33);
        }

        private void d34_Click(object sender, EventArgs e)
        {
            selectDate(34);
        }

        private void d35_Click(object sender, EventArgs e)
        {
            selectDate(35);
        }

        private void d36_Click(object sender, EventArgs e)
        {
            selectDate(36);
        }

        private void d37_Click(object sender, EventArgs e)
        {
            selectDate(37);
        }

        private void d38_Click(object sender, EventArgs e)
        {
            selectDate(38);
        }

        private void d39_Click(object sender, EventArgs e)
        {
            selectDate(39);
        }

        private void d40_Click(object sender, EventArgs e)
        {
            selectDate(40);
        }

        private void d41_Click(object sender, EventArgs e)
        {
            selectDate(41);
        }

        private void btnToday_Click(object sender, EventArgs e)
        {
            theDate = today;
            updateLabels();
            updateCalendar(true);
            textBox.Focus();
        }

        private void updateTheDate(int year, int month, int day)
        {            
            saveCalendar();

            if (day > getDaysInMonth(year, month))
                day = getDaysInMonth(year, month);

            theDate = new DateTime(year, month, day);

            updateLabels();
            updateCalendar(true);
        }

        private void btnPreviousMonth_Click(object sender, EventArgs e)
        {
            int year = theDate.Year;
            int month = theDate.Month;
            int day = theDate.Day;
            if (month == 1)
            {
                month = 12;
                year--;
            }
            else
                month--;

            updateTheDate(year, month, day);
        }

        private void btnNextMonth_Click(object sender, EventArgs e)
        {
            int year = theDate.Year;
            int month = theDate.Month;
            int day = theDate.Day;
            if (month == 12)
            {
                month = 1;
                year++;
            }
            else
                month++;

            updateTheDate(year, month, day);            
        }

        private void btnPreviousYear_Click(object sender, EventArgs e)
        {
            txtYear.Value--;
        }

        private void btnNextYear_Click(object sender, EventArgs e)
        {
            txtYear.Value++;
        }

        private void txtYear_ValueChanged(object sender, EventArgs e)
        {
            int year = Convert.ToInt32(txtYear.Value);
            int month = theDate.Month;
            int day = theDate.Day;

            updateTheDate(year, month, day);
        }

        private void cbMonth_SelectedIndexChanged(object sender, EventArgs e)
        {
            int year = theDate.Year;
            int month = cbMonth.SelectedIndex + 1;
            int day = theDate.Day;

            updateTheDate(year, month, day);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {            
            Application.Exit();
        }        

        private void textBox_TextChanged(object sender, EventArgs e)
        {
            if (calendar[theDate.Day] == null)
                calendar[theDate.Day] = new Day(textBox.Text);
            else
                calendar[theDate.Day].text = textBox.Text;            

            updateCalendar(false);
        }

        private void MyDay_FormClosing(object sender, FormClosingEventArgs e)
        {
            saveCalendar();
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveCalendar();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (aboutBox == null)
                aboutBox = new AboutBox();
            aboutBox.ShowDialog();
        }

        private bool containsWord(string haystack, string needle)
        {
            string[] words = haystack.Split(new char[] { '\r', '\n', ' ' });
            foreach (var word in words)
                if (word == needle)
                    return true;
            return false;
        }

        private void findText(string text, bool matchCase, bool wholeWord, bool regularX, bool findNext)
        {           
            if (!Directory.Exists("data"))
            {
                MessageBox.Show("Missing data directory", "Find");
                return;
            }

            var needle = matchCase ? text : text.ToLower();

            Regex regex = null;
            if (regularX)
                regex = new Regex(needle);

            string[] files = Directory.GetFiles("data", "*.xml");
            Array.Sort(files);
            foreach (var file in files)
            {
                int year = Convert.ToInt32(file.Substring(5, 4));
                if (findNext && (year < theDate.Year))
                    continue;

                int month = Convert.ToInt32(file.Substring(9, 2));
                var dict = loadCalendar(year, month);
                for (int n = 1; n <= 31; n++)
                    if (dict[n] != null)
                    {
                        if (findNext)
                        {
                            var thisDate = new DateTime(year, month, n);
                            if (thisDate.CompareTo(theDate) <= 0)
                                continue;
                        }

                        var haystack = matchCase ? dict[n].text : dict[n].text.ToLower();
                        if (regularX)
                        {
                            Match match = regex.Match(haystack);
                            if (match.Success)
                            {
                                updateTheDate(year, month, n);
                                return;
                            }
                        }
                        else if (wholeWord ? containsWord(haystack, needle) : haystack.Contains(needle))
                        {
                            updateTheDate(year, month, n);
                            return;
                        }
                        
                    }
            }
            MessageBox.Show("Cannot find \"" + text + "\"", "Find");
        }

        int compareTuple(int a1, int a2, int a3, int b1, int b2, int b3)
        {
            if (a1 != b1)
                return a1 - b1;
            if (a2 != b2)
                return a2 - b2;
            if (a3 != b3)
                return a3 - b3;
            return 0;
        }

        // fromdate, toDate = yyyy-mm-dd
        private void exportText(string fromDate, string toDate, string fileName)
        {
            if (!Directory.Exists("data"))
            {
                MessageBox.Show("Missing data directory", "Export");
                return;
            }

            var ass = fromDate.Split('-');
            int y1 = Convert.ToInt32(ass[0]);
            int m1 = Convert.ToInt32(ass[1]);
            int d1 = Convert.ToInt32(ass[2]);
            ass = toDate.Split('-');
            int y2 = Convert.ToInt32(ass[0]);
            int m2 = Convert.ToInt32(ass[1]);
            int d2 = Convert.ToInt32(ass[2]);

            string[] files = Directory.GetFiles("data", "*.xml");
            Array.Sort(files);

            using (var writer = new StreamWriter(fileName))
            {                
                
                foreach (var file in files)
                {
                    int year = Convert.ToInt32(file.Substring(5, 4));
                    if (year < y1)
                        continue;
                    if (year > y2)
                        break;

                    int month = Convert.ToInt32(file.Substring(9, 2));

                    var dict = loadCalendar(year, month);
                    for (int n = 1; n <= 31; n++)
                        if (dict[n] != null)
                        {
                            // date comparison
                            if (compareTuple(y1, m1, d1, year, month, n) <= 0 && compareTuple(y2, m2, d2, year, month, n) >= 0)
                            {
                                var text = dict[n].text;
                                var ass2 = text.Split('\n');
                                foreach (var s in ass2)
                                    writer.WriteLine(year + "-" + LZ(month) + "-" + LZ(n) + ", " + QuoteCSV(s));
                            }
                        }
                }
            }
            MessageBox.Show("Exported to \"" + fileName + "\"", "Export");
        }

        public static string QuoteCSV(string str)
        {
            bool mustQuote = (str.Contains(",") || str.Contains("\"") || str.Contains("\r") || str.Contains("\n"));
            if (mustQuote)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("\"");
                foreach (char nextChar in str)
                {
                    sb.Append(nextChar);
                    if (nextChar == '"')
                        sb.Append("\"");
                }
                sb.Append("\"");
                return sb.ToString();
            }

            return str;
        }


        private void findToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (searchForm == null)
                searchForm = new SearchForm();
            if (searchForm.ShowDialog() == DialogResult.OK)
            {
                var text = searchForm.GetText();
                if (!String.IsNullOrWhiteSpace(text))
                {
                    //MessageBox.Show("Find [" + text + "]");
                    bool matchCase = searchForm.MatchCase();
                    bool wholeWord = searchForm.WholeWord();
                    bool regularX  = searchForm.RegularExpression();
                    findText(text, matchCase, wholeWord, regularX, false);
                }
            }
        }

        private void findNextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (searchForm != null)
            {
                var text = searchForm.GetText();
                if (!String.IsNullOrWhiteSpace(text))
                {
                    bool matchCase = searchForm.MatchCase();
                    bool wholeWord = searchForm.WholeWord();
                    bool regularX  = searchForm.RegularExpression();
                    findText(text, matchCase, wholeWord, regularX, true);
                }
            }
        }

        private void btnBold_Click(object sender, EventArgs e)
        {
            if (calendar[theDate.Day] != null)
            {
                calendar[theDate.Day].isBold = !calendar[theDate.Day].isBold;
                updateCalendar(true);
            }
        }

        private void btnItalic_Click(object sender, EventArgs e)
        {
            if (calendar[theDate.Day] != null)
            {
                calendar[theDate.Day].isItalic = !calendar[theDate.Day].isItalic;
                updateCalendar(true);
            }
        }

        private void btnStike_Click(object sender, EventArgs e)
        {
            if (calendar[theDate.Day] != null)
            {
                calendar[theDate.Day].isStrike = !calendar[theDate.Day].isStrike;
                updateCalendar(true);
            }
        }

        private void boldToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnBold_Click(sender, e);
        }

        private void italicToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnItalic_Click(sender, e);
        }

        private void strikeoutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnStike_Click(sender, e);
        }

        private void SaveDimensions(string key)
        {
            var subKey = "Software\\" + key + "\\Settings";
            RegistryTools.WriteHKCUString(subKey, "FormLeft", Convert.ToString(Left));
            RegistryTools.WriteHKCUString(subKey, "FormTop", Convert.ToString(Top));
            RegistryTools.WriteHKCUString(subKey, "FormWidth", Convert.ToString(Width));
            RegistryTools.WriteHKCUString(subKey, "FormHeight", Convert.ToString(Height));
        }

        private void LoadDimensions(string key)
        {
            var subKey = "Software\\" + key + "\\Settings";
            string val = RegistryTools.ReadHKCUString(subKey, "FormLeft");
            if (!String.IsNullOrEmpty(val))
            {
                StartPosition = FormStartPosition.Manual;
                Left = Convert.ToInt32(val);

                val = RegistryTools.ReadHKCUString(subKey, "FormTop");
                Top = Convert.ToInt32(val);
                val = RegistryTools.ReadHKCUString(subKey, "FormWidth");
                Width = Convert.ToInt32(val);
                val = RegistryTools.ReadHKCUString(subKey, "FormHeight");
                Height = Convert.ToInt32(val);
            }
        }

        private void MyDay_FormClosed(object sender, FormClosedEventArgs e)
        {
            // save
            SaveDimensions(APP_NAME);
        }

        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (exportForm == null)
                exportForm = new ExportForm();
            var now = DateTime.Now;
            exportForm.SetFromDate(now.Year + "-01-01");
            exportForm.SetToDate(now.Year + "-12-31");
            if (exportForm.ShowDialog() == DialogResult.OK)
            {                
                saveFileDialog1.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var fileName = saveFileDialog1.FileName;
                    exportText(exportForm.GetFromDate(), exportForm.GetToDate(), fileName);                    
                }
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(theDate.ToString("yyyy-MM-dd"));
        }
    }
}

