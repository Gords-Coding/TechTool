using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TechTool
{
    public partial class Form2 : Form
    {
        public Form2(ListBox listBox1)
        {
            InitializeComponent();
            Console.WriteLine("Displaying Form2");
            //Form1 f = new Form1(this);
            listBox1.Sorted = true;
            listBox1.Size = new Size(ClientRectangle.Width, ClientRectangle.Height);
            listBox1.BorderStyle = BorderStyle.Fixed3D;
            this.Controls.Add(listBox1);
            /*for (int i = 0; i < softwareAL.Length; i++)
            {
                cartListBox.Items.Add(movieArray[i].ToString());
            }*/
        }
    }
}
