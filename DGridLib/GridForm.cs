using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DGridLib
{
    public partial class GridForm : Form
    {
        public Localization localization { get; set; } 
        public GridForm()
        {
            InitializeComponent();

        }
        public GridForm(DGrid owner)
        {
            InitializeComponent();
            _owner = owner;
            SetLocalization();

        }
        public void setInfo(AllTableInfo allinfo)
        {
            info = allinfo;
            //savedInfo = allinfo;

        }
        public void setSavedInfo(AllTableInfo allinfo)
        {
            savedInfo = allinfo;
            //savedInfo = allinfo;

        }
        public AllTableInfo getInfo()
        {
            return info;
        }
        public void SetLocalization() 
        {
            
            localization = new RussianLocalization();
            //if (_owner.locale == 1049)
            //{
            //    localization = new RussianLocalization();
            //}
            //if (_owner.locale == 1033)
            //{
            //    localization = new EnglishLocalization();
            //}

            saveButton.Text = localization.saveButton;
        }
        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            bool hasChangedItems = (!info.isEqual(savedInfo));

            if (hasChangedItems)
            {
                var window = MessageBox.Show(
                                "В таблице остались несохраненные данные. Сохранить?",
                                "Сохранение данных",
                                MessageBoxButtons.YesNoCancel);
                if (window == DialogResult.Yes)
                {
                    e.Cancel = false;

                    object?[] array = _owner.CollectTableData(info);
                    object? arrRows = array[0];
                    object? arrCols = array[1];
                    object? arrValues = array[2];


                    bool isNoException = _owner.InvokeSave(arrRows, arrCols, arrValues);
                    _owner.InvokeClose();

                }
                if (window == DialogResult.No)
                {
                    e.Cancel = false;
                    _owner.InvokeClose();

                }
                if (window == DialogResult.Cancel)
                {
                    e.Cancel = true;

                }
            }
            else
            {
                e.Cancel = false;
                _owner.InvokeClose();
            }



        }
        public void AddEmptyColumnsAndRows(DataGridView dataGrid, int cols, int rows)
        {
            for (int i = 0; i < cols; i++)
            {
                DataGridViewColumn column = new DataGridViewColumn();
                column.CellTemplate = new DataGridViewTextBoxCell();
                dataGrid.Columns.Add(column);
            }
            for (int i = 0; i < rows; i++)
            {
                DataGridViewRow row = new DataGridViewRow();
                dataGrid.Rows.Add(row);
            }
        }
        private void saveButton_Click(object sender, EventArgs e)
        {
            object?[] array = _owner.CollectTableData(info);
            object? arrRows = array[0];
            object? arrCols = array[1];
            object? arrValues = array[2];
            
            bool isNoException = _owner.InvokeSave(arrRows, arrCols, arrValues);
            if (!isNoException)
            {
                var window = MessageBox.Show(
                            "Не удалось сохранить данные в таблице.",
                            "Ошибка ",
                            MessageBoxButtons.OK);
            }
            savedInfo = (AllTableInfo)info.Clone();
        }
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if(info == null)
            {
                return;
            }
            //Действия с колонками
            if (info!= null && e.RowIndex == 0 && e.ColumnIndex>1)
            {
                string colCode = dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString();
                //Изменение или удаление( удаление столбца происходит при сохранении в случае отсутствия назнания у столбца)
                if (e.ColumnIndex!= dataGridView1.Columns.Count - 1)
                {
                    info.ColumnsInfo[e.ColumnIndex - 2].smpr_data_Info_Tables_Columns_Code = colCode;
                    var columnValues = info.ValuesInfo.Where(prop => prop.smpr_data_Info_Tables_Values_ColCode == colCode).ToList();
                    foreach (var value in columnValues)
                    {
                        value.smpr_data_Info_Tables_Values_ColCode = colCode;
                    }
                }
                //Добавление
                if (e.ColumnIndex == dataGridView1.Columns.Count-1)
                {
                    AddColumn(colCode);

                }
            }
            //Действия со строками 
            if (info!= null && e.ColumnIndex == 0 && e.RowIndex> 0)
            {
                string rowCode = dataGridView1[e.ColumnIndex, e.RowIndex].Value != null ?
                        dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString() : null;
                //Изменение и удаление
                if (e.RowIndex != dataGridView1.Rows.Count - 1)
                {
                    info.RowsInfo[e.RowIndex - 1].smpr_data_Info_Tables_Rows_Code= rowCode;
                    var rowValues = info.ValuesInfo.Where(prop => prop.smpr_data_Info_Tables_Values_RowCode == rowCode).ToList();
                    foreach (var value in rowValues)
                    {
                        value.smpr_data_Info_Tables_Values_RowCode = rowCode;
                    }
                }
                //Добавление
                if (e.RowIndex == dataGridView1.Rows.Count - 1)
                {
                    AddRow(rowCode);

                }
            }
            //описание строк
            if (e.ColumnIndex ==1 && e.RowIndex>0 && e.RowIndex < dataGridView1.RowCount-1)
            {
                string descriptor = dataGridView1[e.ColumnIndex, e.RowIndex].Value!=null?
                        dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString() : null;
                string rowCode = dataGridView1[0, e.RowIndex].Value.ToString();
                info.RowsInfo[e.RowIndex - 1].smpr_data_Info_Tables_Rows_Descriptor = descriptor;
            }
            // Действия с ячейками
            if (info != null & e.RowIndex > 0 && e.RowIndex < dataGridView1.Rows.Count-1 && e.ColumnIndex > 1 && e.ColumnIndex < dataGridView1.Columns.Count - 1)
            {
                var colCode = info.ColumnsInfo[e.ColumnIndex - 2].smpr_data_Info_Tables_Columns_Code;
                var rowCode = info.RowsInfo[e.RowIndex-1].smpr_data_Info_Tables_Rows_Code;
                var changedValue = info.ValuesInfo.Where(prop => prop.smpr_data_Info_Tables_Values_ColCode == colCode &
                                                        prop.smpr_data_Info_Tables_Values_RowCode == rowCode).FirstOrDefault();
                int changedValueIndex = info.ValuesInfo.IndexOf(changedValue);
                
                //Удаление значения
                if (dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString() == "")
                {
                    info.ValuesInfo.RemoveAt(changedValueIndex);
                    return;
                }
                // Изменение
                if (changedValueIndex != -1)
                {
                    try
                    {
                        info.ValuesInfo[changedValueIndex].smpr_data_Info_Tables_Values_Value =
                                        Convert.ToDouble(dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString().Replace(".", ","));

                    }
                    catch
                    {
                        dataGridView1[e.ColumnIndex, e.RowIndex].Value = info.ValuesInfo[changedValueIndex].smpr_data_Info_Tables_Values_Value;
                    }
                }
                //Добавление
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
                int precision = info.ColumnsInfo.Where(p => p.smpr_data_Info_Tables_Columns_Code == colCode).FirstOrDefault().smpr_data_Info_Tables_Columns_Precision;
                string format = new String('0', precision);
                dataGridView1[e.ColumnIndex, e.RowIndex].Style.Format = $"##.{format}";
                dataGridView1[e.ColumnIndex, e.RowIndex].Style.ApplyStyle(dataGridView1[e.ColumnIndex, e.RowIndex].Style);
            }
            //Название таблицы
            if (e.ColumnIndex == 0 && e.RowIndex== 0)
            {
                _owner.tableName = dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString();
            }
            bool unchangedCells = (e.RowIndex == 1 && e.ColumnIndex < 2) ||
                                  (e.RowIndex == 0 && e.ColumnIndex == 1) ||
                                  (e.RowIndex == dataGridView1.RowCount - 1 && e.ColumnIndex > 0) ||
                                  (e.ColumnIndex == dataGridView1.ColumnCount - 1 && e.RowIndex > 0);
            if (unchangedCells)
            {
                //dataGridView1.CancelEdit();
                dataGridView1[e.ColumnIndex, e.RowIndex].Value = "";
            }

        }
        public void AddColumn(string colCode)
        {
            DataColumns newColumn = new DataColumns();
            newColumn.smpr_data_Info_Tables_Columns_Code = colCode;
            newColumn.smpr_data_Info_Tables_Columns_Comment = "";
            newColumn.smpr_data_Info_Tables_Columns_DataType = 3;
            newColumn.smpr_data_Info_Tables_Columns_Descriptor = "";
            newColumn.smpr_data_Info_Tables_Columns_Index = info.ColumnsInfo.Count;
            newColumn.smpr_data_Info_Tables_Columns_ModFlag = 0;
            newColumn.smpr_data_Info_Tables_Columns_Precision = 3;
            newColumn.smpr_data_Info_Tables_Columns_TableId = info.ColumnsInfo[0].smpr_data_Info_Tables_Columns_TableId;

            info.ColumnsInfo.Add(newColumn);

            string format = new String('0', newColumn.smpr_data_Info_Tables_Columns_Precision);

            dataGridView1.Columns[dataGridView1.Columns.Count - 1].ValueType = typeof(System.Double);
            //TODO формат числа
            dataGridView1.Columns[dataGridView1.Columns.Count - 1].DefaultCellStyle.Format = $"##.{format}";
            //Новый пустой столбец
            dataGridView1.Columns.Add($"column{dataGridView1.Columns.Count}", "");

        }
        public void AddRow(string rowCode)
        {
            DataRows newRow = new DataRows();
            newRow.smpr_data_Info_Tables_Rows_Code = rowCode;
            newRow.smpr_data_Info_Tables_Rows_Comment = "";
            newRow.smpr_data_Info_Tables_Rows_Descriptor = "";
            newRow.smpr_data_Info_Tables_Rows_Index = info.RowsInfo.Count;
            newRow.smpr_data_Info_Tables_Rows_ModFlag = 0;
            newRow.smpr_data_Info_Tables_Rows_TableId = info.RowsInfo[0].smpr_data_Info_Tables_Rows_TableId;

            info.RowsInfo.Add(newRow);

            dataGridView1.Rows.Add();
            //for(int i = 0; i < dataGridView1.ColumnCount; i++)
            //{
            //    dataGridView1[i, dataGridView1.RowCount - 1].Value = "";
            //}
        }
        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //try
            //{

            //    int precision = info.ColumnsInfo[e.ColumnIndex - 2].smpr_data_Info_Tables_Columns_Precision;
            //    string format = new String('0', precision);
            //    e.CellStyle.Format = $"##.{format}";
            //    e.FormattingApplied = true;
            //}
            //catch
            //{

            //}
        }
        public DGrid _owner { get; set; }
        public AllTableInfo info { get; set; }
        public AllTableInfo savedInfo = new AllTableInfo();
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton saveButton;
        public System.Windows.Forms.DataGridView dataGridView1;

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            AddingForm newForm = new AddingForm();
            newForm.Show();
        }
    }
}
