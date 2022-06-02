using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Grafik_Deneme
{
    static class static_Deneme
    {
        public static int x { get; set; }
        static static_Deneme()
        {
            Form1 form1 = new Form1();

            //form1.event
            //Form1.txt_Box.Text = "Deneme";
        }
    }


    /*class Class1
    {
        public Class1()
        {
            Form1.instance.comboBox1.Items.Add("Base");
        }
        public Class1(int x)
        {
            Form1.instance.comboBox1.Items.Add("Base_" + x);
        }

    }
    
    class Class2:Class1
    {
        public Class2()
        {
            Form1.instance.comboBox1.Items.Add("Main");
        }
        public Class2(int y) : base(2)
        {
            Form1.instance.comboBox1.Items.Add("Main_" + y);
        }
    }*/
}
