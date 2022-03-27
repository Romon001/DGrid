using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DGridLib
{

    [ComVisible(true)]
    [Guid("E349D0EE-F5F6-4ea8-9279-6ED2F2891E47")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ISendResults
    {
        object SendInfo();
    }
    
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComSourceInterfaces(typeof(ISendResults))]
    [Guid("36CF32E5-6157-4b67-BB34-31EAA756AEDB")]
    [ComVisible(true)]
    public class DGrid
    {
        public delegate object WriteMessageDelegate();
        public event WriteMessageDelegate SendResult;

        GridForm form { get; set; }
        AllTableInfo allInfo = new AllTableInfo();
        AllTableInfo allInfo2 = new AllTableInfo();

        AllTableInfo allInfoAfterChange = new AllTableInfo();
        object allInfoAfterChangeObject;
        
        [DispId(1)]
        [ComVisible(true)]
        public void Open(object rowInfo, object columnsInfo, object valuesInfo)
        {
             object[,] arrRows = (object[,])rowInfo;
             object[,] arrColumns = (object[,])columnsInfo;
             object[,] arrValues = (object[,])valuesInfo;

             form = new GridForm(this);


            allInfo.ColumnsInfo = new List<DataColumns>();
            allInfo.RowsInfo = new List<DataRows>();
            allInfo.ValuesInfo = new List<DataValues>();
            // Заполнение инфо по строкам
            for (int i = 0; i < arrRows.GetLength(0); i++)
            {
                DataRows row = new DataRows();
                row.smpr_data_Info_Tables_Rows_TableId = arrRows[i, 0].ToString();
                row.smpr_data_Info_Tables_Rows_Code = arrRows[i, 1].ToString();
                row.smpr_data_Info_Tables_Rows_Descriptor = arrRows[i, 2].ToString();
                row.smpr_data_Info_Tables_Rows_Comment = arrRows[i, 3].ToString();
                row.smpr_data_Info_Tables_Rows_Index = arrRows[i, 4].ToString();
                row.smpr_data_Info_Tables_Rows_ModFlag = arrRows[i, 5].ToString();
                allInfo.RowsInfo.Add(row);

            }


            // Заполнение инфо по столбцам
            for (int i = 0; i < arrColumns.GetLength(0); i++)
            {
                DataColumns col = new DataColumns();
                col.smpr_data_Info_Tables_Columns_TableId = arrColumns[i, 0].ToString();
                col.smpr_data_Info_Tables_Columns_Code = arrColumns[i, 1].ToString();
                col.smpr_data_Info_Tables_Columns_Descriptor = arrColumns[i, 2].ToString();
                col.smpr_data_Info_Tables_Columns_DataType = arrColumns[i, 3].ToString();
                col.smpr_data_Info_Tables_Columns_Precision = arrColumns[i, 4].ToString();
                col.smpr_data_Info_Tables_Columns_Comment = arrColumns[i, 5].ToString();
                col.smpr_data_Info_Tables_Columns_Index = arrColumns[i, 6].ToString();
                col.smpr_data_Info_Tables_Columns_ModFlag = arrColumns[i, 7].ToString();
                allInfo.ColumnsInfo.Add(col);
            }

            // Заполнение инфо по значениям
            for (int i = 0; i < arrValues.GetLength(0); i++)
            {
                DataValues val = new DataValues();
                val.smpr_data_Info_Tables_Values_TableId = arrValues[i, 0].ToString();
                val.smpr_data_Info_Tables_Values_CaseId = arrValues[i, 1].ToString();
                val.smpr_data_Info_Tables_Values_PeriodId = arrValues[i, 2].ToString();
                val.smpr_data_Info_Tables_Values_RowCode = arrValues[i, 3].ToString();
                val.smpr_data_Info_Tables_Values_ColCode = arrValues[i, 4].ToString();
                val.smpr_data_Info_Tables_Values_Value = arrValues[i, 5].ToString();
                val.smpr_data_Info_Tables_Values_Formula = arrValues[i, 6].ToString();
                val.smpr_data_Info_Tables_Values_Comment = arrValues[i, 7].ToString();
                val.smpr_data_Info_Tables_Values_Format = arrValues[i, 8].ToString();
                val.smpr_data_Info_Tables_Values_ModFlag = arrValues[i, 9].ToString();
                allInfo.ValuesInfo.Add(val);
            }

            form.dataGridView1.Columns.Add("RowName", "...");
            //Заполнение столбцов таблицы
            for (int i = 0; i < allInfo.ColumnsInfo.Count; i++)
            {
                form.dataGridView1.Columns.Add(allInfo.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Code,
                    allInfo.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Code);
            }
            //Заполнение строк таблицы
            for (int i = 0; i < allInfo.RowsInfo.Count; i++)
            {
                DataGridViewRow row = new DataGridViewRow();
                form.dataGridView1.Rows.Add(row);
            }
            for (int i = 0; i < allInfo.RowsInfo.Count; i++)
            {
                form.dataGridView1[0, i].Value = allInfo.RowsInfo[i].smpr_data_Info_Tables_Rows_Code;
            }

            //Заполнение значений таблицы
            for (int i = 0; i < allInfo.ValuesInfo.Count; i++)
            {
                
                string value = allInfo.ValuesInfo[i].smpr_data_Info_Tables_Values_Value;
                string rowCode = allInfo.ValuesInfo[i].smpr_data_Info_Tables_Values_RowCode;
                var selectedRow = allInfo.RowsInfo.Where(p => p.smpr_data_Info_Tables_Rows_Code == rowCode).FirstOrDefault();
                int rowIndex = allInfo.RowsInfo.IndexOf(selectedRow);
                string colCode = allInfo.ValuesInfo[i].smpr_data_Info_Tables_Values_ColCode;
                var selectedCol = allInfo.ColumnsInfo.Where(p => p.smpr_data_Info_Tables_Columns_Code == colCode).FirstOrDefault();
                int colIndex = allInfo.ColumnsInfo.IndexOf(selectedCol)+1;
                form.dataGridView1[colIndex,rowIndex].Value = value;

            }
            form.setInfo(allInfo);
            form.Show();


        }
        
        [DispId(2)]
        [ComVisible(true)]
        public virtual void Save() 
        {

            AllTableInfo newInfo = form.getInfo();
            allInfoAfterChange = newInfo;
            
            //Rows
            object[,] arrRows = new object[allInfoAfterChange.RowsInfo.Count, 6];
            
            for(int i=0;i< allInfoAfterChange.RowsInfo.Count;i++)
            {
                arrRows[i, 0] = allInfoAfterChange.RowsInfo[i].smpr_data_Info_Tables_Rows_TableId;
                arrRows[i, 1] = allInfoAfterChange.RowsInfo[i].smpr_data_Info_Tables_Rows_Code;
                arrRows[i, 2] = allInfoAfterChange.RowsInfo[i].smpr_data_Info_Tables_Rows_Descriptor;
                arrRows[i, 3] = allInfoAfterChange.RowsInfo[i].smpr_data_Info_Tables_Rows_Comment;
                arrRows[i, 4] = allInfoAfterChange.RowsInfo[i].smpr_data_Info_Tables_Rows_Index;
                arrRows[i, 5] = allInfoAfterChange.RowsInfo[i].smpr_data_Info_Tables_Rows_ModFlag;
                
            }
            
            //Columns
            object[,] arrCols = new object[allInfoAfterChange.ColumnsInfo.Count,8];
            for(int i=0;i< allInfoAfterChange.ColumnsInfo.Count;i++)
            {
                arrCols[i, 0] = allInfoAfterChange.ColumnsInfo[i].smpr_data_Info_Tables_Columns_TableId;
                arrCols[i, 1] = allInfoAfterChange.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Code;
                arrCols[i, 2] = allInfoAfterChange.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Descriptor;
                arrCols[i, 3] = allInfoAfterChange.ColumnsInfo[i].smpr_data_Info_Tables_Columns_DataType;
                arrCols[i, 4] = allInfoAfterChange.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Precision;
                arrCols[i, 5] = allInfoAfterChange.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Comment;
                arrCols[i, 6] = allInfoAfterChange.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Index;
                arrCols[i, 7] = allInfoAfterChange.ColumnsInfo[i].smpr_data_Info_Tables_Columns_ModFlag;
            }
            

            //Values
            object[,] arrValues = new object[allInfoAfterChange.ValuesInfo.Count,10];
            for (int i = 0; i < allInfoAfterChange.ValuesInfo.Count; i++)
            {
                arrValues[i,0] = allInfoAfterChange.ValuesInfo[i].smpr_data_Info_Tables_Values_TableId;
                arrValues[i, 1] = allInfoAfterChange.ValuesInfo[i].smpr_data_Info_Tables_Values_CaseId;
                arrValues[i, 2] = allInfoAfterChange.ValuesInfo[i].smpr_data_Info_Tables_Values_PeriodId;
                arrValues[i, 3] = allInfoAfterChange.ValuesInfo[i].smpr_data_Info_Tables_Values_RowCode;
                arrValues[i, 4] = allInfoAfterChange.ValuesInfo[i].smpr_data_Info_Tables_Values_ColCode;
                arrValues[i, 5] = allInfoAfterChange.ValuesInfo[i].smpr_data_Info_Tables_Values_Value;
                arrValues[i, 6] = allInfoAfterChange.ValuesInfo[i].smpr_data_Info_Tables_Values_Formula;
                arrValues[i, 7] = allInfoAfterChange.ValuesInfo[i].smpr_data_Info_Tables_Values_Comment;
                arrValues[i, 8] = allInfoAfterChange.ValuesInfo[i].smpr_data_Info_Tables_Values_Format;
                arrValues[i, 9] = allInfoAfterChange.ValuesInfo[i].smpr_data_Info_Tables_Values_ModFlag;
            
            }
           

            object[] returnInfoArray = { arrRows, arrCols, arrValues };
            object returnInfo = (object)returnInfoArray;
            allInfoAfterChangeObject = returnInfo;
        }
        
        [DispId(3)]
        [ComVisible(true)]
        public virtual object InvokeSendResult()
        {
            return SendResult?.Invoke();
        }
        public object SendInfo()
        {
            if (this.allInfoAfterChangeObject != null)
            {
                return this.allInfoAfterChangeObject;
            }
            else
            {
                return "";
            }
        }

    }
    public class GridForm : Form
    {
        public delegate object WriteMessageDelegate();
        
        public GridForm()
        {
            InitializeComponent();
        }
        public GridForm(DGrid owner)
        {
            InitializeComponent();
            _owner = owner;
        }
        public void setInfo(AllTableInfo allinfo)
        {
            info = allinfo;
        }
        public AllTableInfo getInfo()
        {
            return info;
        }
        private void InitializeComponent()
        {
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.toolStrip1.SuspendLayout();
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form2_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButton1});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(829, 25);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(84, 22);
            this.toolStripButton1.Text = "Save Changes";
            this.toolStripButton1.Click += new System.EventHandler(this.toolStripButton1_Click);

            // 
            // dataGridView1
            // 
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(0, 28);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(800, 418);
            this.dataGridView1.TabIndex = 1;
            this.dataGridView1.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellValueChanged);
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(829, 450);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.toolStrip1);
            this.Name = "GridForm";
            this.Text = "GridForm";
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            _owner.SendResult += _owner.SendInfo;
            _owner.Save();
        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            _owner.Save();
            _owner.SendResult += _owner.SendInfo;
            
            //var q=_owner.Save();

        }
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (info != null & e.ColumnIndex != 0)
            {
                var colCode = info.ColumnsInfo[e.ColumnIndex - 1].smpr_data_Info_Tables_Columns_Code;
                var rowCode = info.RowsInfo[e.RowIndex].smpr_data_Info_Tables_Rows_Code;
                var changedValue = info.ValuesInfo.Where(prop => prop.smpr_data_Info_Tables_Values_ColCode == colCode &
                                                        prop.smpr_data_Info_Tables_Values_RowCode == rowCode).FirstOrDefault();
                int changedValueIndex = info.ValuesInfo.IndexOf(changedValue);
                if (changedValueIndex != -1)
                {
                    info.ValuesInfo[changedValueIndex].smpr_data_Info_Tables_Values_Value = dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString();
                }
                else
                {
                    DataValues newValue = new DataValues();
                    newValue.smpr_data_Info_Tables_Values_TableId = info.ValuesInfo[0].smpr_data_Info_Tables_Values_TableId;
                    newValue.smpr_data_Info_Tables_Values_CaseId = info.ValuesInfo[0].smpr_data_Info_Tables_Values_CaseId;
                    newValue.smpr_data_Info_Tables_Values_PeriodId = info.ValuesInfo[0].smpr_data_Info_Tables_Values_PeriodId;
                    newValue.smpr_data_Info_Tables_Values_RowCode = rowCode;
                    newValue.smpr_data_Info_Tables_Values_ColCode = colCode;
                    newValue.smpr_data_Info_Tables_Values_Value = dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString();
                    newValue.smpr_data_Info_Tables_Values_Formula = info.ValuesInfo[0].smpr_data_Info_Tables_Values_Formula;
                    newValue.smpr_data_Info_Tables_Values_Comment = info.ValuesInfo[0].smpr_data_Info_Tables_Values_Comment;
                    newValue.smpr_data_Info_Tables_Values_Format = info.ValuesInfo[0].smpr_data_Info_Tables_Values_Format;
                    newValue.smpr_data_Info_Tables_Values_ModFlag = info.ValuesInfo[0].smpr_data_Info_Tables_Values_ModFlag;
                    info.ValuesInfo.Add(newValue);
                }

            }
        }
        public DGrid _owner { get; set; }
        public AllTableInfo info { get; set; }
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        public System.Windows.Forms.DataGridView dataGridView1;

    }
    public class AllTableInfo
    {
        public List<DataRows> RowsInfo { get; set; }
        public List<DataColumns> ColumnsInfo { get; set; }
        public List<DataValues> ValuesInfo { get; set; }

    }

    public class DataRows
    {

        public string smpr_data_Info_Tables_Rows_TableId { get; set; }
        public string smpr_data_Info_Tables_Rows_Code { get; set; }
        public string smpr_data_Info_Tables_Rows_Descriptor { get; set; }
        public string smpr_data_Info_Tables_Rows_Comment { get; set; }
        public string smpr_data_Info_Tables_Rows_Index { get; set; }
        public string smpr_data_Info_Tables_Rows_ModFlag { get; set; }
    }
    public class DataColumns
    {

        public string smpr_data_Info_Tables_Columns_TableId { get; set; }
        public string smpr_data_Info_Tables_Columns_Code { get; set; }
        public string smpr_data_Info_Tables_Columns_Descriptor { get; set; }
        public string smpr_data_Info_Tables_Columns_DataType { get; set; }
        public string smpr_data_Info_Tables_Columns_Precision { get; set; }
        public string smpr_data_Info_Tables_Columns_Comment { get; set; }
        public string smpr_data_Info_Tables_Columns_Index { get; set; }
        public string smpr_data_Info_Tables_Columns_ModFlag { get; set; }

    }
    public class DataValues
    {

        public string smpr_data_Info_Tables_Values_TableId { get; set; }
        public string smpr_data_Info_Tables_Values_CaseId { get; set; }
        public string smpr_data_Info_Tables_Values_PeriodId { get; set; }
        public string smpr_data_Info_Tables_Values_RowCode { get; set; }
        public string smpr_data_Info_Tables_Values_ColCode { get; set; }
        public string smpr_data_Info_Tables_Values_Value { get; set; }
        public string smpr_data_Info_Tables_Values_Formula { get; set; }
        public string smpr_data_Info_Tables_Values_Comment { get; set; }
        public string smpr_data_Info_Tables_Values_Format { get; set; }
        public string smpr_data_Info_Tables_Values_ModFlag { get; set; }

    }
    public enum smpr_data_Info_Tables_Columns
    {
        smpr_data_Info_Tables_Columns_TableId = 0,         //!< Table Id field index. See <see cref="smpr_data_Table"/> for predefined values.
        smpr_data_Info_Tables_Columns_Code = 1,
        smpr_data_Info_Tables_Columns_Descriptor = 2,
        smpr_data_Info_Tables_Columns_DataType = 3,        //!< I4, values - VT_... (VT_R8 and VT_BSTR) currently
        smpr_data_Info_Tables_Columns_Precision = 4,
        smpr_data_Info_Tables_Columns_Comment = 5,
        smpr_data_Info_Tables_Columns_Index = 6,
        smpr_data_Info_Tables_Columns_ModFlag = 7
    }
    public enum smpr_data_Info_Tables_Values
    {
        smpr_data_Info_Tables_Values_TableId = 0,         //!< Table Id field index. See <see cref="smpr_data_Table"/> for predefined values
        smpr_data_Info_Tables_Values_CaseId = 1,
        smpr_data_Info_Tables_Values_PeriodId = 2,
        smpr_data_Info_Tables_Values_RowCode = 3,
        smpr_data_Info_Tables_Values_ColCode = 4,
        smpr_data_Info_Tables_Values_Value = 5,           //Column DataType (Empty, Text or Double)
        smpr_data_Info_Tables_Values_Formula = 6,
        smpr_data_Info_Tables_Values_Comment = 7,
        smpr_data_Info_Tables_Values_Format = 8,
        smpr_data_Info_Tables_Values_ModFlag = 9
    }
    public enum smpr_data_Info_Tables_Rows
    {
        smpr_data_Info_Tables_Rows_TableId = 0,         //!< Table Id field index. See <see cref="smpr_data_Table"/> for predefined values.
        smpr_data_Info_Tables_Rows_Code = 1,
        smpr_data_Info_Tables_Rows_Descriptor = 2,
        smpr_data_Info_Tables_Rows_Comment = 3,
        smpr_data_Info_Tables_Rows_Index = 4,
        smpr_data_Info_Tables_Rows_ModFlag = 5
    };

}
