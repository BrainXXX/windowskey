using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Futures
{
    public partial class Form1 : Form
    {
        static string api = "https://fapi.binance.com";

        static string coin1 = "BTCUSDT";
        double lastPrice = 1234.56;

        public Form1()
        {
            this.Location = new Point(0 - 11, Screen.PrimaryScreen.WorkingArea.Height - 72);
            InitializeComponent();
            this.Text = "Futures";
            this.button1.BackgroundImage = Properties.Resources.Update;
            this.button1.BackgroundImageLayout = ImageLayout.Stretch;
            button1.Text = "";
            textBox1.Text = "";
            label1.Text = "1234";

            StartTimer();
        }

        private void StartTimer()
        {
            TimerCallback func = new TimerCallback((x) =>
            {
                lastPrice = double.Parse(GetLastPrice().Replace(".", ","));

                if (lastPrice > double.Parse(label1.Text))
                {
                    label1.ForeColor = Color.Green;
                }
                else
                {
                    label1.ForeColor = Color.Red;
                }

                this.label1.BeginInvoke((MethodInvoker)(() => this.label1.Text = lastPrice.ToString("0.00")));
                //if (label1.Text != textBox1.Text.Split(' ').Last())
                this.textBox1.BeginInvoke((MethodInvoker)(() => this.textBox1.AppendText(lastPrice.ToString("0.00") + "\t" + DateTime.Now.ToString("HH:mm:ss") + "\r\n")));
                if (textBox1.Lines.Length > 100) this.textBox1.BeginInvoke((MethodInvoker)(() => this.textBox1.Text = ""));
            });

            System.Threading.Timer tm = new System.Threading.Timer(func, null, 0, 1000);
        }

        private static string GetLastPrice()
        {
            string byc = getWebResponse(api + "/fapi/v1/ticker/24hr?symbol=" + coin1);

            int pos1 = byc.IndexOf("\"lastPrice\":", 0);
            int pos2 = byc.IndexOf("\"", pos1 + 12);
            int pos3 = byc.IndexOf("\"", pos2 + 1);

            string str1 = byc.Substring(pos2 + 1, pos3 - pos2 - 1);

            return str1;
        }

        private static string getWebResponse(string url)
        {

            // create request..
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(url);

            // use GET method
            webRequest.Method = "GET";

            // POST!
            HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

            // read response into StreamReader
            Stream responseStream = webResponse.GetResponseStream();
            StreamReader _responseStream = new StreamReader(responseStream);

            // get raw result
            return _responseStream.ReadToEnd();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            StartTimer();
        }
    }
}