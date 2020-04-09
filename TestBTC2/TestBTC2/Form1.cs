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
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestBTC2
{
    public partial class Form1 : Form
    {
        string Coin1 = "BTC-NEO";
        string Coin2 = "USDT-NEO";
        string Coin3 = "USDT-BTC";
        string Coin4 = "BTC-QRL";

        public Form1()
        {
            InitializeComponent();

            button1.BackColor = Color.Transparent;
            button1.FlatStyle = FlatStyle.Flat;
            button1.FlatAppearance.BorderSize = 0;
            button1.TabStop = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            outputToDisplayPrice(Coin1);
            outputToDisplayPrice(Coin2);
            outputToDisplayPrice(Coin3);
            outputToDisplayPrice(Coin4);
            outputToDisplayVolume(Coin1);
            outputToDisplayVolume(Coin4);
            textBox1.Text += "35 NEO = " + 35 * getCoinInfo(Coin2) + " $ " + 35 * getCoinInfo(Coin1) + "\r\n";
            textBox1.Text += "2164 QRL = " + 2164 * getCoinInfo(Coin4) * getCoinInfo(Coin3) + " $ " + 2164 * getCoinInfo(Coin4) + "\r\n";
        }

        private void outputToDisplayPrice(string coin)
        {
            textBox1.Text += String.Format(coin + "\t{0} price", getCoinInfo(coin)) + "\r\n";
        }

        private void outputToDisplayVolume(string coin)
        {
            textBox1.Text += String.Format(coin + "\t{0:F2} volume", getCoinVolume(coin)) + "\r\n";
        }

        private static decimal getCoinInfo(string coin)
        {
            string byc = getWebResponse("https://bittrex.com/api/v1.1/public/getticker?market=" + coin);

            int pos1 = byc.IndexOf("Last", 0);
            int pos2 = byc.IndexOf(":", pos1);
            int pos3 = byc.IndexOf("}", pos2);

            return Convert.ToDecimal(byc.Substring(pos2 + 1, pos3 - pos2 - 1), CultureInfo.InvariantCulture);
        }

        private static decimal getCoinVolume(string coin)
        {
            string byc = getWebResponse("https://bittrex.com/api/v1.1/public/getmarketsummary?market=" + coin);

            int pos1 = byc.IndexOf("BaseVolume", 0);
            int pos2 = byc.IndexOf(":", pos1);
            int pos3 = byc.IndexOf(",", pos2);

            return Convert.ToDecimal(byc.Substring(pos2 + 1, pos3 - pos2 - 1), CultureInfo.InvariantCulture);
        }

        public static string getWebResponse(string url)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

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
    }
}