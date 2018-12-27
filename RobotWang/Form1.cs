using AutoToolHelper;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;

namespace RobotWang
{
    public partial class Form1 : Form
    {
        //qq聊天窗口句柄集合
        private volatile List<IntPtr> hWndCollect = new List<IntPtr>();
        private volatile List<string> qqTitleList = new List<string>();
        private volatile Dictionary<IntPtr, string> dicList = new Dictionary<IntPtr, string>();
        private const int NULL = 0;
        private const int MAXBYTE = 255;
        //释放按键的常量
        private const int KEYEVENTF_KEYUP = 2;
        private const int WM_PASTE = 0x302;
        private const int WM_CUT = 0x300;
        private const int WM_COPY = 0x301;
        //SendMessage 参数
        private const int WM_KEYDOWN = 0X100;
        private const int WM_KEYUP = 0X101;
        private const int WM_SYSCHAR = 0X106;
        private const int WM_SYSKEYUP = 0X105;
        private const int WM_SYSKEYDOWN = 0X104;
        private const int WM_CHAR = 0X102;

        private string qqTitleName = string.Empty;

        /// <summary>
        /// 构造函数，窗体初始化
        /// </summary>
        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 窗体加载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            Control.CheckForIllegalCrossThreadCalls = false;//屏蔽跨线程操作ui控件的异常

            Thread getMessageThread = new Thread(GetQQMessage);
            getMessageThread.SetApartmentState(ApartmentState.STA);
            getMessageThread.IsBackground = true;
            getMessageThread.Start();
        }

        #region  手动发送
        /// <summary>
        /// 发送文字
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            string textStr = this.textBox1.Text;
            if (string.IsNullOrEmpty(textStr))
            {
                MessageBox.Show("将要发送的内容不能为空！", "RobotWang");
                return;
            }
            if (textStr.Length > 3420)
            {
                MessageBox.Show("将要发送的文字内容超长！", "RobotWang");
                return;
            }
            SendTextStr();
        }
        /// <summary>
        /// 发送文件资源
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            string sendFilePath = this.textBox2.Text;
            if (string.IsNullOrEmpty(sendFilePath))
            {
                MessageBox.Show("将要发送的文件路径不能为空！", "RobotWang");
                return;
            }
            if (!File.Exists(sendFilePath))
            {
                MessageBox.Show("将要发送的文件不存在！", "RobotWang");
                return;
            }
            SendImage();
        }
        /// <summary>
        /// 设置comboBox1改变时修改qqTitleName
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            qqTitleName = this.comboBox1.Text;
        }
        /// <summary>
        /// 选择文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();     //显示选择文件对话框
            if (openFileDialog.ShowDialog() == DialogResult.OK)
                this.textBox2.Text = openFileDialog.FileName;          //显示文件路径
        }

        #endregion

        #region  主动发送

        #endregion

        #region  被动发送

        #endregion

        #region  功能函数

        /// <summary>
        /// 发送文字信息
        /// </summary>
        public void SendTextStr()
        {
            string sendStr = this.textBox1.Text;
            if (string.IsNullOrEmpty(sendStr)) return;

            IntPtr intPtr = User32API.FindWindow(null, qqTitleName);
            Clipboard.Clear();
            Clipboard.SetText(sendStr);

            Send(intPtr);
        }
        /// <summary>
        /// 发送图片
        /// </summary>
        public void SendImage()
        {
            string imgPath = this.textBox2.Text;
            IntPtr intPtr = User32API.FindWindow(null, qqTitleName);
            Clipboard.Clear();
            Bitmap bmp = new Bitmap(imgPath);
            Clipboard.SetImage(bmp);

            Send(intPtr);
        }
        /// <summary>
        /// 发送图片链接
        /// </summary>
        public void SendImage(string url)
        {           
            IntPtr intPtr = User32API.FindWindow(null, qqTitleName);
            Clipboard.Clear();
            Bitmap bmp = DowloadImage(url);
            if (bmp == null) return;
            Clipboard.SetImage(bmp);

            Send(intPtr);
        }
        /// <summary>
        /// 发送文件
        /// </summary>
        public void SendFile()
        {
            IntPtr intPtr = User32API.FindWindow(null, qqTitleName);
            Clipboard.Clear();
            System.Collections.Specialized.StringCollection strcoll = new System.Collections.Specialized.StringCollection();
            string filePath = "";
            if (string.IsNullOrEmpty(filePath)) return;
            strcoll.Add(filePath);
            Clipboard.SetFileDropList(strcoll);

            Send(intPtr);
        }
        /// <summary>
        /// 模拟发送操作
        /// </summary>
        /// <param name="intPtr"></param>
        public void Send(IntPtr intPtr)
        {
            User32API.SetForegroundWindow(intPtr);//把找到的的对话框在最前面显示如果使用了这个方法
            User32API.ShowWindow(intPtr, ShowWindowCmd.SW_SHOWNORMAL);
            User32API.SendMessageA(intPtr, WM_PASTE, 0, 0);
            User32API.SendMessageA(intPtr, WM_KEYDOWN, 0X0D, 0);//发
            User32API.SendMessageA(intPtr, WM_KEYUP, 0X0D, 0); //送
            User32API.SendMessageA(intPtr, WM_CHAR, 0X0D, 0); //回车
            Clipboard.Clear();
        }
        /// <summary>
        /// 获取qq聊天信息
        /// </summary>
        public void GetQQMessage()
        {
            try
            {
                int tempLen = 0;
                string messageStr = string.Empty;
                List<string> contentList = new List<string>();
                List<string> tempList = new List<string>();
                while (true)
                {
                    GetSessionWindow();
                    this.comboBox1.Items.AddRange(qqTitleList.ToArray());
                    foreach (var hWnd in hWndCollect)
                    {
                        messageStr = FindUserMessage(hWnd);
                        contentList = messageStr.Split('\r').ToList();
                        foreach (var item in contentList)
                        {
                            if (!string.IsNullOrWhiteSpace(item))
                            {
                                tempList.Add(item.Replace(" ", ""));
                            }
                        }
                        tempLen = tempList.Count();
                        if (tempLen > 1)
                        {
                            if (dicList.ContainsKey(hWnd))
                                this.comboBox1.Text = dicList[hWnd];
                            ResponseOperation(tempList[tempLen - 2], tempList[tempLen - 1]);
                        }
                    }
                    Thread.Sleep(200);
                }
            }
            catch (Exception ex)
            {
                FileRW.WriteToFile(ex.Message);

                Thread getMessageThread = new Thread(GetQQMessage);
                getMessageThread.SetApartmentState(ApartmentState.STA);
                getMessageThread.IsBackground = true;
                getMessageThread.Start();
            }
        }
        /// <summary>
        /// 获取qq聊天窗口的句柄集合
        /// </summary>
        /// <returns></returns>
        public void GetSessionWindow()
        {
            hWndCollect.Clear();
            qqTitleList.Clear();
            dicList.Clear();
            this.comboBox1.Items.Clear();
            User32API.EnumWindows(ScanSessionWindow, NULL);//委托
        }
        /// <summary>
        /// 获取qq聊天窗口的句柄集合
        /// </summary>
        /// <param name="hwnd"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        private bool ScanSessionWindow(IntPtr hwnd, IntPtr lParam)
        {
            StringBuilder buf = new StringBuilder(MAXBYTE);
            if (User32API.GetClassName(hwnd, buf, MAXBYTE) > 0 && buf.ToString() == "TXGuiFoundation")
                if (User32API.GetWindowTextLength(hwnd) > 0 && User32API.GetWindowText(hwnd, buf, MAXBYTE) > 0)
                {
                    string str = buf.ToString();
                    if (str != "TXMenuWindow" && str != "QQ" && str != "增加时长")
                    {
                        FileRW.WriteToFile("\t" + (hWndCollect.Count + 1) + ": " + str);
                        hWndCollect.Add(hwnd);
                        qqTitleList.Add(str);
                        dicList.Add(hwnd, str);
                    }
                }
            return true;
        }
        /// <summary>
        /// 根据句柄获取窗口里的信息
        /// </summary>
        /// <param name="hwnd"></param>
        public string FindUserMessage(IntPtr hwnd)
        {
            string qqMessage = string.Empty;
            if (User32API.IsWindow(hwnd))
            {
                AutomationElement element = AutomationElement.FromHandle(hwnd);
                element = element.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "消息"));
                if (element != null && element.Current.IsEnabled)
                {
                    ValuePattern vpTextEdit = element.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                    if (vpTextEdit != null)
                    {
                        qqMessage = vpTextEdit.Current.Value;
                    }
                }
            }
            return qqMessage;
        }
        /// <summary>
        /// 响应操作
        /// </summary>
        /// <param name="newMessage"></param>
        public void ResponseOperation(string name, string newMessage)
        {
            if (!this.listBox1.Items.Contains(newMessage))
            {
                this.listBox1.Items.Add(name);
                this.listBox1.Items.Add(newMessage);
                this.listBox1.SelectedIndex = this.listBox1.Items.Count;
            }
            switch (newMessage)
            {
                case "$图片":
                    {
                        SendPicture();
                        break;
                    }
                case "$鸡汤":
                    {
                        ChickenSoup();
                        break;
                    }
                case "$天气":
                    {
                        break;
                    }
                case "$谜语":
                    {
                        break;
                    }
                case "$猜拳":
                    {
                        break;
                    }
                case "$$":
                    {
                        break;
                    }
                default:
                    break;
            }
        }
        /// <summary>
        /// 发送图片
        /// </summary>
        public void SendPicture()
        {
            string api = "https://54188.xyz/api/imagespiderapi/getrandompicture";
            string imgUrl = string.Empty;

            string result = GetApi(api);
            if (!string.IsNullOrEmpty(result))
            {
                imgUrl = result.Replace("\"","");
                if (!string.IsNullOrEmpty(imgUrl))
                {
                    SendImage(imgUrl);
                }
            }
        }
        /// <summary>
        /// 发送名人名言
        /// </summary>
        public void ChickenSoup()
        {
            string api = "https://api.lwl12.com/hitokoto/v1";//一言api
            string result = GetApi(api);
            if (string.IsNullOrEmpty(result) || result.Length > 3420)
                return;
            this.textBox1.Text = result;
            SendTextStr();
        }
        /// <summary>
        /// GET方式请求
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        private string GetApi(string url)
        {
            string strHTML = "";
            WebClient myWebClient = new WebClient();
            Stream myStream = myWebClient.OpenRead(url);
            StreamReader sr = new StreamReader(myStream, System.Text.Encoding.GetEncoding("utf-8"));
            strHTML = sr.ReadToEnd();
            myStream.Close();
            return strHTML;
        }
        /// <summary>   
        /// 下载验证码图片并保存到本地   
        /// </summary>   
        /// <param name="Url">验证码URL</param>   
        /// <param name="cookCon">Cookies值</param>   
        /// <param name="savePath">保存位置/文件名</param>   
        public Bitmap DowloadImage(string Url)
        {  
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(Url);
            webRequest.AllowWriteStreamBuffering = true;
            webRequest.Credentials = CredentialCache.DefaultCredentials;
            webRequest.MaximumResponseHeadersLength = -1;
            webRequest.Accept = "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/x-shockwave-flash, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, */*";
            webRequest.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; Maxthon; .NET CLR 1.1.4322)";
            webRequest.ContentType = "application/x-www-form-urlencoded";
            webRequest.Method = "GET";
            webRequest.Headers.Add("Accept-Language", "zh-cn");
            webRequest.Headers.Add("Accept-Encoding", "gzip,deflate");
            webRequest.KeepAlive = true;
            try
            {
                //获取服务器返回的资源   
                using (HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse())
                {
                    using (Stream sream = webResponse.GetResponseStream())
                    {
                        List<byte> list = new List<byte>();
                        while (true)
                        {
                            int data = sream.ReadByte();
                            if (data == -1)
                                break;
                            list.Add((byte)data);
                        }
                        byte[] yzmByte = list.ToArray();
                        Bitmap bitmap = BytesToBitmap(yzmByte);
                        return bitmap;
                    }
                }
            }
            catch (WebException ex)
            {
                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
            return null;
        }
        /// <summary>
        /// byte[] 转图片 
        /// </summary>
        /// <param name="Bytes"></param>
        /// <returns></returns>
        public Bitmap BytesToBitmap(byte[] Bytes)
        {
            MemoryStream stream = null;
            try
            {
                stream = new MemoryStream(Bytes);
                return new Bitmap((Image)new Bitmap(stream));
            }
            catch (ArgumentNullException ex)
            {
                throw ex;
            }
            catch (ArgumentException ex)
            {
                throw ex;
            }
            finally
            {
                stream.Close();
            }
        }
        #endregion
    }
}
