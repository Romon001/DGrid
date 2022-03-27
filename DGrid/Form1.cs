using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DGrid
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            dataGrid = new DGridLib.DGrid();
        }

        DGridLib.DGrid dataGrid { get; set; }
        object[,] arrRows = new object[,] { {"1", "1", "1", "1", "1", "1" },
                                                { "2", "2", "2", "2", "2", "2" },
                                                { "2", "3", "2", "2", "2", "2" },
                                                { "2", "4", "2", "2", "2", "2" },
                                                { "2", "6", "2", "2", "2", "2" },
                                                { "2", "7", "2", "2", "2", "2" },
                                                { "2", "8", "2", "2", "2", "2" },
                                                { "2", "9", "2", "2", "2", "2" },
                                                { "2", "10", "2", "2", "2", "2" },
                                                { "2", "11", "2", "2", "2", "2" },
                                                { "2", "12", "2", "2", "2", "2" },
            };
        object[,] arrCols = new object[,] { {"1", "1", "1", "1", "1", "1", "1", "1" },
                                                { "2", "2", "2", "2", "2", "2", "2", "2" },
                                                { "2", "3", "2", "2", "2", "2", "2", "2" },
                                                { "2", "4", "2", "2", "2", "2", "2", "2" },
                                                { "2", "5", "2", "2", "2", "2", "2", "2" },
                                                { "2", "6", "2", "2", "2", "2", "2", "2" },
                                                { "2", "7", "2", "2", "2", "2", "2", "2" },
                                                { "2", "8", "2", "2", "2", "2", "2", "2" },
                                                { "2", "555", "2", "2", "2", "2", "2", "2" },
                                                { "2", "aaa", "2", "2", "2", "2", "2", "2" }
            };
        object[,] arrVals = new object[,] { {"1", "1", "1", "1", "1", "1val", "1", "1", "1", "1" },
                                                { "2", "2", "2", "2", "2", "2val", "2", "2", "2", "2" } };
        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGrid == null)
            {
                dataGrid = new DGridLib.DGrid();

            }
            dataGrid.SendInfo
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            object c = "1";
            var a= dataGrid.InvokeSendResult();
            DGridLib.DGrid grid = new DGridLib.DGrid();
            var b = dataGrid.SendInfo();
            a = dataGrid.InvokeSendResult();
        }


    }
}
