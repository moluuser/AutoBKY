using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;

namespace GTPAutoReponse
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll")]
        private extern static IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("user32.dll", EntryPoint = "SendMessageA")]
        private static extern int SendMessage(IntPtr hwnd, int wMsg, int wParam, StringBuilder lParam);
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        ////////////////////////////////
        ///鼠标事件代码
        [DllImport("User32")]
        public extern static void mouse_event(int dwFlags, int dx, int dy, int dwData, IntPtr dwExtraInfo);

        [DllImport("User32")]
        public extern static void SetCursorPos(int x, int y);

        [DllImport("User32")]
        public extern static bool GetCursorPos(out POINT p);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        public enum MouseEventFlags
        {
            Move = 0x0001,
            LeftDown = 0x0002,
            LeftUp = 0x0004,
            RightDown = 0x0008,
            RightUp = 0x0010,
            MiddleDown = 0x0020,
            MiddleUp = 0x0040,
            Wheel = 0x0800,
            Absolute = 0x8000
        }

        private void AutoClick(int x, int y)
        {
            POINT p = new POINT();
            GetCursorPos(out p);
            try
            {
                SetCursorPos(x, y);
                mouse_event((int)(MouseEventFlags.LeftDown | MouseEventFlags.Absolute), 0, 0, 0, IntPtr.Zero);
                mouse_event((int)(MouseEventFlags.LeftUp | MouseEventFlags.Absolute), 0, 0, 0, IntPtr.Zero);
            }
            finally
            {
                SetCursorPos(p.X, p.Y);
            }
        }

        ////////////////////////////////

        const int WM_GETTEXT = 0x0D;
        const int buffer_size = 1024 * 10;
        IntPtr myp = IntPtr.Zero;

        //遗留连接SqlServer代码
        /*
        string connStr = "Data source=LAPTOP-EJLC5ICS;Initial Catalog=GTPExam;User ID=bkygtpuser;Password=gtp2009";
        SqlConnection conn = null;
        SqlDataReader dr = null;
        SqlCommand cmd = null;
        */

        //连接MySQL
        static string connetStr = "server=150.158.212.132;port=3306;user=root;password=xDCA^OIdE{t($e2; database=autobky;charset=utf8";
        // server=127.0.0.1/localhost 代表本机，端口号port默认是3306可以不写
        MySqlConnection conn = new MySqlConnection(connetStr);

        //存储面板题目
        StringBuilder QuestionBuffer = new StringBuilder(buffer_size);
        string QuestionContent;
        string SingleQuestionAns;
        //存储面板选项
        StringBuilder AnsBuffer1 = new StringBuilder(buffer_size);
        StringBuilder AnsBuffer2 = new StringBuilder(buffer_size);
        StringBuilder AnsBuffer3 = new StringBuilder(buffer_size);
        StringBuilder AnsBuffer4 = new StringBuilder(buffer_size);
        StringBuilder AnsBuffer5 = new StringBuilder(buffer_size);
        StringBuilder AnsBuffer6 = new StringBuilder(buffer_size);

        //多选题正确答案
        string MultiAns1 = null;
        string MultiAns2 = null;
        string MultiAns3 = null;
        string MultiAns4 = null;
        string MultiAns5 = null;
        string MultiAns6 = null;

        //套卷科目
        StringBuilder ExamTypeBuffer = new StringBuilder(buffer_size);


        public Form1()
        {
            InitializeComponent();
        }

        private int getExamType()
        {
            string ExamType = null;

            IntPtr my0 = new IntPtr(0);
            IntPtr my1 = new IntPtr(0);
            IntPtr my2 = new IntPtr(0);
            IntPtr my3 = new IntPtr(0);

            IntPtr my4 = new IntPtr(0);
            IntPtr my5 = new IntPtr(0);

            my0 = FindWindowEx(myp, IntPtr.Zero, null, null);
            my1 = FindWindowEx(myp, my0, null, null);
            my2 = FindWindowEx(myp, my1, null, null);
            my3 = FindWindowEx(myp, my2, null, null);

            my4 = FindWindowEx(my3, IntPtr.Zero, null, null);
            my5 = FindWindowEx(my3, my4, null, null);
            if (my5 != IntPtr.Zero)
            {
                SendMessage(my5, WM_GETTEXT, buffer_size, ExamTypeBuffer);
                ExamType = ExamTypeBuffer.ToString();
                //测试使用，输出题目信息
                //MessageBox.Show(ExamType);

                if (ExamType.IndexOf("近现代史纲要") != -1)
                {
                    return 1;
                }
                else if (ExamType.IndexOf("思想道德") != -1)
                {
                    return 2;
                }
                else if (ExamType.IndexOf("中国化") != -1)
                {
                    return 3;
                }
                else if (ExamType.IndexOf("马克思") != -1)
                {
                    return 4;
                }
                else
                {
                    MessageBox.Show("未识别套卷！");
                    return -1;
                }
            }
            else
            {
                MessageBox.Show("getQuestionType ERROR");
                return -1;
            }
        }

        //获取面板题目信息
        private bool getQuestion()
        {
            IntPtr my0 = new IntPtr(0);
            IntPtr my1 = new IntPtr(0);
            IntPtr my2 = new IntPtr(0);
            IntPtr my3 = new IntPtr(0);
            IntPtr my4 = new IntPtr(0);

            my0 = FindWindowEx(myp, IntPtr.Zero, null, null);
            my1 = FindWindowEx(my0, IntPtr.Zero, null, null);
            my2 = FindWindowEx(my1, IntPtr.Zero, null, null);
            my3 = FindWindowEx(my2, IntPtr.Zero, null, null);
            my4 = FindWindowEx(my2, my3, "RichTextWndClass", null);
            if (my4 != IntPtr.Zero)
            {
                SendMessage(my4, WM_GETTEXT, buffer_size, QuestionBuffer);
                //测试使用，输出题目信息
                //MessageBox.Show(QuestionBuffer.ToString());
                string temp = QuestionBuffer.ToString();
                temp = temp.Replace("-", "_");
                temp = temp.Substring(0, temp.Length - 1);
                temp = temp.Substring(0, temp.Length - 1);
                QuestionContent = temp.Replace("\"", "_") + "%";
                //MessageBox.Show(QuestionContent);
                return true;
            }
            else
            {
                MessageBox.Show("getQuestion ERROR");
                return false;
            }
        }

        //获取面板选择题选项
        private int getQuestionAns()
        {
            int AnsNum = 4;

            IntPtr my0 = new IntPtr(0);
            IntPtr my1 = new IntPtr(0);
            IntPtr my2 = new IntPtr(0);
            IntPtr my3 = new IntPtr(0);
            IntPtr my4 = new IntPtr(0);

            IntPtr my5 = new IntPtr(0);
            IntPtr my6 = new IntPtr(0);
            IntPtr my7 = new IntPtr(0);
            IntPtr my8 = new IntPtr(0);
            IntPtr my9 = new IntPtr(0);
            //ABCD
            IntPtr my10 = new IntPtr(0);
            IntPtr my11 = new IntPtr(0);
            IntPtr my12 = new IntPtr(0);
            IntPtr my13 = new IntPtr(0);
            IntPtr my14 = new IntPtr(0);
            IntPtr my15 = new IntPtr(0);
            //对应文本句柄
            IntPtr my10a = new IntPtr(0);
            IntPtr my11a = new IntPtr(0);
            IntPtr my12a = new IntPtr(0);
            IntPtr my13a = new IntPtr(0);
            IntPtr my14a = new IntPtr(0);
            IntPtr my15a = new IntPtr(0);

            my0 = FindWindowEx(myp, IntPtr.Zero, null, null);
            my1 = FindWindowEx(my0, IntPtr.Zero, null, null);
            my2 = FindWindowEx(my0, my1, null, null);
            my3 = FindWindowEx(my0, my2, null, null);
            my4 = FindWindowEx(my3, IntPtr.Zero, null, null);

            my5 = FindWindowEx(my3, my4, null, null);
            my6 = FindWindowEx(my3, my5, null, null);

            my7 = FindWindowEx(my6, IntPtr.Zero, null, null);
            //my8 = my7;
            my9 = my7;
            my10 = FindWindowEx(my9, IntPtr.Zero, null, "A");
            my11 = FindWindowEx(my9, IntPtr.Zero, null, "B");
            my12 = FindWindowEx(my9, IntPtr.Zero, null, "C");
            my13 = FindWindowEx(my9, IntPtr.Zero, null, "D");
            my14 = FindWindowEx(my9, IntPtr.Zero, null, "E");
            my15 = FindWindowEx(my9, IntPtr.Zero, null, "F");

            my10a = FindWindowEx(my9, my10, null, null);
            my11a = FindWindowEx(my9, my11, null, null);
            my12a = FindWindowEx(my9, my12, null, null);
            my13a = FindWindowEx(my9, my13, null, null);
            my14a = FindWindowEx(my9, my14, null, null);
            my15a = FindWindowEx(my9, my15, null, null);

            SendMessage(my10a, WM_GETTEXT, buffer_size, AnsBuffer1);
            SendMessage(my11a, WM_GETTEXT, buffer_size, AnsBuffer2);
            SendMessage(my12a, WM_GETTEXT, buffer_size, AnsBuffer3);
            SendMessage(my13a, WM_GETTEXT, buffer_size, AnsBuffer4);
            if (my14a != my10)
            {
                SendMessage(my14a, WM_GETTEXT, buffer_size, AnsBuffer5);
                AnsNum = 5;
            }

            if (my15a != my10)
            {
                SendMessage(my15a, WM_GETTEXT, buffer_size, AnsBuffer6);
                AnsNum = 6;
            }

            //MessageBox.Show(AnsNum.ToString());
            /*
            MessageBox.Show(AnsBuffer1.ToString());
            MessageBox.Show(AnsBuffer2.ToString());
            MessageBox.Show(AnsBuffer3.ToString());
            MessageBox.Show(AnsBuffer4.ToString());
            MessageBox.Show(AnsBuffer5.ToString());
            MessageBox.Show(AnsBuffer6.ToString());
            */
            //返回选项数量
            return AnsNum;
        }

        //获取题目类型
        /*
        private string getQuestionTpye()
        {
            string SQL_getQueType = "select qType from QuestionContent where contentText like '" + QuestionContent + "'";
            cmd = new SqlCommand(SQL_getQueType, conn);
            dr = cmd.ExecuteReader();
            dr.Read();
            string qType = dr[0].ToString();
            dr.Close();
            return qType;
        }
        */

        //获取单选题正确答案
        private void getTrueAns_single()
        {
            //获取答案ID
            string SQL_getAnsID = "select ans1 from alldata where content like '" + QuestionContent + "'";
            //MessageBox.Show(SQL_getAnsID);
            MySqlCommand cmd = new MySqlCommand(SQL_getAnsID, conn);
            MySqlDataReader myData = cmd.ExecuteReader();
            string Ans = "";
            while (myData.Read())
            {
                Ans = myData.GetString(0);
                //MessageBox.Show(myData.GetString(0));
            }
            myData.Close();
            SingleQuestionAns = Ans;
        }

        //获取判断题正确答案
        private bool getTrueAns_judge()
        {
            string SQL_getAnsID = "select ans1 from alldata where content like '" + QuestionContent + "'";
            //MessageBox.Show(SQL_getAnsID);
            MySqlCommand cmd = new MySqlCommand(SQL_getAnsID, conn);
            MySqlDataReader myData = cmd.ExecuteReader();

            string Ans = "";
            while (myData.Read())
            {
                Ans = myData.GetString(0);
                //MessageBox.Show(myData.GetString(0));
            }
            myData.Close();
            if (Ans == "正确")
                return true;
            else if (Ans == "错误")
                return false;
            else
                MessageBox.Show("不是判断题！");
            return false;
        }

        //获取多选题正确答案
        private int getTrueAns_multi()
        {
            int TrueAnsNum = 2;
            //获取答案ID
            string SQL_getAnsID = "select ans1, ans2, ans3, ans4, ans5, ans6 from alldata where content like '" + QuestionContent + "'";
            //MessageBox.Show(SQL_getAnsID);
            //MessageBox.Show(QuestionContent);
            //剪切板
            //Clipboard.SetText(SQL_getAnsID);

            MySqlCommand cmd = new MySqlCommand(SQL_getAnsID, conn);
            MySqlDataReader myData = cmd.ExecuteReader();

            while (myData.Read())
            {
                MultiAns1 = myData.GetString(0);
                MultiAns2 = myData.GetString(1);
                //MessageBox.Show(MultiAns1);
                if (!myData.IsDBNull(2))
                {
                    TrueAnsNum++;
                    MultiAns3 = myData.GetString(2);
                }
                if (!myData.IsDBNull(3))
                {
                    TrueAnsNum++;
                    MultiAns4 = myData.GetString(3);
                }
                if (!myData.IsDBNull(4))
                {
                    TrueAnsNum++;
                    MultiAns5 = myData.GetString(4);
                }
                if (!myData.IsDBNull(5))
                {
                    TrueAnsNum++;
                    MultiAns6 = myData.GetString(5);
                }
            }
            //MessageBox.Show("非法题型！");
            myData.Close();
            return TrueAnsNum;

        }


        //点击“下一题”
        private void clickNext()
        {
            this.AutoClick(int.Parse(textBox1.Text), int.Parse(textBox2.Text));
        }

        //点击选项1-6;-1;0
        private void clickOption(int num)
        {
            switch (num)
            {
                //正确
                case -1:
                    this.AutoClick(int.Parse(textBox9.Text), int.Parse(textBox12.Text));
                    break;
                //错误
                case 0:
                    this.AutoClick(int.Parse(textBox10.Text), int.Parse(textBox12.Text));
                    break;

                case 1:
                    this.AutoClick(int.Parse(textBox3.Text), int.Parse(textBox11.Text));
                    break;
                case 2:
                    this.AutoClick(int.Parse(textBox3.Text), int.Parse(textBox4.Text));
                    break;
                case 3:
                    this.AutoClick(int.Parse(textBox3.Text), int.Parse(textBox8.Text));
                    break;
                case 4:
                    this.AutoClick(int.Parse(textBox3.Text), int.Parse(textBox7.Text));
                    break;
                case 5:
                    this.AutoClick(int.Parse(textBox3.Text), int.Parse(textBox6.Text));
                    break;
                case 6:
                    this.AutoClick(int.Parse(textBox3.Text), int.Parse(textBox5.Text));
                    break;
            }
        }

        //自动单选题
        private void AutoSingle()
        {
            getQuestion();
            getQuestionAns();
            getTrueAns_single();
            //MessageBox.Show(SingleQuestionAns);
            if (SingleQuestionAns == AnsBuffer1.ToString())
            {
                clickOption(1);
            }
            else if (SingleQuestionAns == AnsBuffer2.ToString())
            {
                clickOption(2);
            }
            else if (SingleQuestionAns == AnsBuffer3.ToString())
            {
                clickOption(3);
            }
            else if (SingleQuestionAns == AnsBuffer4.ToString())
            {
                clickOption(4);
            }
            else
                clickOption(4);
            //System.Threading.Thread.Sleep(500);
        }

        //自动判断题
        private void AutoJudge()
        {
            getQuestion();
            //System.Threading.Thread.Sleep(500);
            if (getTrueAns_judge())
                clickOption(-1);
            else
                clickOption(0);
        }

        //自动多选题
        private void AutoMulti()
        {
            getQuestion();
            int DisplayAnsNum = getQuestionAns();
            int TrueAnsNum = getTrueAns_multi();
            List<string> MultiDisplay = new List<string> { AnsBuffer1.ToString(), AnsBuffer2.ToString(), AnsBuffer3.ToString(), AnsBuffer4.ToString(), AnsBuffer5.ToString(), AnsBuffer6.ToString() };
            List<string> MultiAns = new List<string> { MultiAns1, MultiAns2, MultiAns3, MultiAns4, MultiAns5, MultiAns6 };

            for (int i = 0; i < DisplayAnsNum; ++i)
            {
                for (int j = 0; j < TrueAnsNum; ++j)
                {
                    //MessageBox.Show(MultiDisplay[i]);
                    //MessageBox.Show(MultiAns[j]);
                    if (MultiDisplay[i] == MultiAns[j])
                    {
                        clickOption(i + 1);
                        //MessageBox.Show(i.ToString());
                    }
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.timer1.Enabled = true;
            this.timer1.Interval = 10;//timer控件的执行频率

            //测试数据库      
            try
            {
                //打开数据库连接
                conn.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show("数据库连接失败！" + ex.Message);
                Application.Exit();
            }

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            textBox14.Text = Cursor.Position.X.ToString();
            textBox13.Text = Cursor.Position.Y.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            myp = FindWindow(null, "通用考试平台");
            for (int i = 0; i < 40; ++i)
            {
                AutoSingle();
                clickNext();
                System.Threading.Thread.Sleep(1000);
            }
            for (int i = 0; i < 20; ++i)
            {
                AutoMulti();
                clickNext();
                System.Threading.Thread.Sleep(1000);
            }

            for (int i = 0; i < 20; ++i)
            {
                AutoJudge();
                clickNext();
                System.Threading.Thread.Sleep(1000);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            myp = FindWindow(null, "通用考试平台");
            for (int i = 0; i < 50; ++i)
            {
                AutoSingle();
                clickNext();
                System.Threading.Thread.Sleep(1000);
            }
            for (int i = 0; i < 20; ++i)
            {
                AutoMulti();
                clickNext();
                System.Threading.Thread.Sleep(1000);
            }

            for (int i = 0; i < 30; ++i)
            {
                AutoJudge();
                clickNext();
                System.Threading.Thread.Sleep(1000);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            myp = FindWindow(null, "通用考试平台");
            for (int i = 0; i < 50; ++i)
            {
                AutoSingle();
                clickNext();
                System.Threading.Thread.Sleep(1000);
            }
            for (int i = 0; i < 15; ++i)
            {
                AutoMulti();
                clickNext();
                System.Threading.Thread.Sleep(1000);
            }

            for (int i = 0; i < 20; ++i)
            {
                AutoJudge();
                clickNext();
                System.Threading.Thread.Sleep(1000);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            myp = FindWindow(null, "通用考试平台");
            for (int i = 0; i < 40; ++i)
            {
                AutoSingle();
                clickNext();
                System.Threading.Thread.Sleep(1000);
            }
            for (int i = 0; i < 20; ++i)
            {
                AutoMulti();
                clickNext();
                System.Threading.Thread.Sleep(1000);
            }

            for (int i = 0; i < 20; ++i)
            {
                AutoJudge();
                clickNext();
                System.Threading.Thread.Sleep(1000);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            myp = FindWindow(null, "通用考试平台");
            int num = int.Parse(textBox15.Text);
            for (int i = 0; i < num; ++i)
            {
                AutoSingle();
                clickNext();
                System.Threading.Thread.Sleep(1000);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            myp = FindWindow(null, "通用考试平台");
            int num = int.Parse(textBox15.Text);
            for (int i = 0; i < num; ++i)
            {
                AutoMulti();
                clickNext();
                System.Threading.Thread.Sleep(1000);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            myp = FindWindow(null, "通用考试平台");
            int num = int.Parse(textBox15.Text);
            for (int i = 0; i < num; ++i)
            {
                AutoJudge();
                clickNext();
                System.Threading.Thread.Sleep(1000);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            conn.Close();
        }
    }
}

