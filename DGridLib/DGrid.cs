using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
    public interface ISendTableInfo
    {
        //TODO: изменить на Save в конце
        bool Save(object arrRows, object arrColumns, object arrValues);
        public bool Close();
        public object SetAdditionalInfo();
    }

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComSourceInterfaces(typeof(ISendTableInfo))]
    [Guid("36CF32E5-6157-4b67-BB34-31EAA756AEDB")]
    [ComVisible(true)]
    public class DGrid: IDisposable
    {
        public delegate bool SaveDelegate(object arrRows, object arrColumns, object arrValues);
        [DispId(4)]
        //TODO: изменить на Save в конце
        public event SaveDelegate Save;
        //SendTableInfo
        public delegate bool CloseDelegate();
        //Todo: спросить что принимает event Close
        public event CloseDelegate Close;

        public delegate object SetAdditionalInfoDelegate();
        public event SetAdditionalInfoDelegate SetAdditionalInfo;

        GridForm form { get; set; }
        AllTableInfo allInfo { get; set; }
        AllTableInfo allInfoAfterChange = new AllTableInfo();
        object allInfoAfterChangeObject;
        private bool disposed = false;

        [DispId(5)]
        public string? server { get; set; }
        [DispId(6)]
        public string? user { get; set; }
        [DispId(7)]
        public string? modelName { get; set; }
        [DispId(8)]
        public string? caseName { get; set; }
        [DispId(9)]
        public string? period { get; set; }
        [DispId(10)]
        public string? tableName { get; set; }
        [DispId(12)]
        public int locale { get; set; }

        public DGrid()
        {
            allInfo = new AllTableInfo();
        }
        ~DGrid()
        {
            Dispose(disposing: false);
        }
        public void Dispose()
        {
            Dispose(disposing: true);

            GC.SuppressFinalize(this);

        }
        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                disposed = true;
            }
        }
        

        [DispId(1)]
        [ComVisible(true)]
        public bool Open(object rowInfo, object columnsInfo, object valuesInfo)
        {
            bool isNoException = true;
            allInfo = new AllTableInfo();
            try
            {
                //if(SetAdditionalInfo != null)
                //{
                //    object[] addInfo = (object[])SetAdditionalInfo();
                //    server = addInfo[0];
                //    user = addInfo[0];
                //    modelName = addInfo[0];
                //    caseName = addInfo[0];
                //    period = addInfo[0];
                //    tableName = addInfo[0];
                //}


                object[,] arrRows = (object[,])rowInfo;
                object[,] arrColumns = (object[,])columnsInfo;
                object[,] arrValues = (object[,])valuesInfo;

                form = new GridForm(this);
                form.Text = "Таблица";
            
                allInfo.ColumnsInfo = new BindingList<DataColumns>();
                allInfo.RowsInfo = new BindingList<DataRows>();
                allInfo.ValuesInfo = new BindingList<DataValues>();
                // Заполнение инфо по строкам
                if (rowInfo != null)
                {
                    for (int i = 0; i < arrRows.GetLength(0); i++)
                    {
                        DataRows row = new DataRows();
                        row.smpr_data_Info_Tables_Rows_TableId = Convert.ToInt32(arrRows[i, 0]);
                        row.smpr_data_Info_Tables_Rows_Code = ToStringOrNull(arrRows[i, 1]).ToString();
                        row.smpr_data_Info_Tables_Rows_Descriptor = ToStringOrNull(arrRows[i, 2]).ToString();
                        row.smpr_data_Info_Tables_Rows_Comment = ToStringOrNull(arrRows[i, 3]).ToString();
                        row.smpr_data_Info_Tables_Rows_Index = Convert.ToInt32(arrRows[i, 4]);
                        row.smpr_data_Info_Tables_Rows_ModFlag = Convert.ToInt32(arrRows[i, 5]);
                        allInfo.RowsInfo.Add(row);

                    }
                }


                // Заполнение инфо по столбцам
                if (columnsInfo != null) 
                {
                    for (int i = 0; i < arrColumns.GetLength(0); i++)
                    {
                        DataColumns col = new DataColumns();
                        col.smpr_data_Info_Tables_Columns_TableId = Convert.ToInt32(arrColumns[i, 0]);
                        col.smpr_data_Info_Tables_Columns_Code = arrColumns[i, 1].ToString();
                        col.smpr_data_Info_Tables_Columns_Descriptor = ToStringOrNull(arrColumns[i, 2]).ToString();
                        col.smpr_data_Info_Tables_Columns_DataType = Convert.ToInt32(arrColumns[i, 3]);
                        col.smpr_data_Info_Tables_Columns_Precision = Convert.ToInt32(arrColumns[i, 4]);
                        col.smpr_data_Info_Tables_Columns_Comment = ToStringOrNull(arrColumns[i, 5]).ToString();
                        col.smpr_data_Info_Tables_Columns_Index = Convert.ToInt32(arrColumns[i, 6]);
                        col.smpr_data_Info_Tables_Columns_ModFlag = Convert.ToInt32(arrColumns[i, 7]);
                        allInfo.ColumnsInfo.Add(col);
                    }
                }


                // Заполнение инфо по значениям
                if (valuesInfo != null)
                {
                    for (int i = 0; i < arrValues.GetLength(0); i++)
                    {
                        DataValues val = new DataValues();
                        val.smpr_data_Info_Tables_Values_TableId = Convert.ToInt32(arrValues[i, 0]);
                        val.smpr_data_Info_Tables_Values_CaseId = Convert.ToInt32(arrValues[i, 1]);
                        val.smpr_data_Info_Tables_Values_PeriodId = Convert.ToInt32(arrValues[i, 2]);
                        val.smpr_data_Info_Tables_Values_RowCode = ToStringOrNull(arrValues[i, 3]).ToString();
                        val.smpr_data_Info_Tables_Values_ColCode = ToStringOrNull(arrValues[i, 4]).ToString();
                        val.smpr_data_Info_Tables_Values_Value = arrValues[i, 5];
                        val.smpr_data_Info_Tables_Values_Formula = ToStringOrNull(arrValues[i, 6]).ToString();
                        val.smpr_data_Info_Tables_Values_Comment = ToStringOrNull(arrValues[i, 7]).ToString();
                        val.smpr_data_Info_Tables_Values_Format = ToStringOrNull(arrValues[i, 8]).ToString();
                        val.smpr_data_Info_Tables_Values_ModFlag = Convert.ToInt32(arrValues[i, 9]);
                        allInfo.ValuesInfo.Add(val);
                    }
                }
            

                form.dataGridView1.Columns.Add("RowCode", " ");
                form.dataGridView1.Columns["RowCode"].Frozen = true;
                form.dataGridView1.Columns.Add("RowName", " ");
                form.dataGridView1.Columns["RowName"].Frozen = true;
                form.dataGridView1.Rows.Add();

                //Заполнение столбцов таблицы
                for (int i = 0; i < allInfo.ColumnsInfo.Count; i++)
                {
                    int precision = allInfo.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Precision;
                    string format = new String('0', precision);
                    form.dataGridView1.Columns.Add(allInfo.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Code,
                        allInfo.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Code);
                    form.dataGridView1.Columns[i + 2].ValueType = typeof(System.Double);
                    form.dataGridView1[i+2, 0].Value = allInfo.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Code;
                    //TODO формат числа
                    form.dataGridView1.Columns[allInfo.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Code].DefaultCellStyle.Format = $"##.{format}";
                }
                //Заполнение строк таблицы
                for (int i = 0; i < allInfo.RowsInfo.Count; i++)
                {
                    DataGridViewRow row = new DataGridViewRow();
                    form.dataGridView1.Rows.Add(row);
                }
                for (int i = 0; i < allInfo.RowsInfo.Count; i++)
                {
                    form.dataGridView1[0, i+1].Value = allInfo.RowsInfo[i].smpr_data_Info_Tables_Rows_Code;
                    form.dataGridView1[1, i+1].Value = allInfo.RowsInfo[i].smpr_data_Info_Tables_Rows_Descriptor;
                }
                form.dataGridView1.Rows[0].DefaultCellStyle.BackColor = Color.LightGray;
                //Пустая строка и столбец для добавления новых данных
                form.dataGridView1.Columns.Add($"column{form.dataGridView1.Columns.Count}", "");
                form.dataGridView1.Rows.Add();
                //Название таблицы
                form.dataGridView1[0,0].Value = (tableName!=null) ? tableName : "";
                form.Text = (tableName != null) ? $"Таблица - {tableName}" : "Таблица";
                //Заполнение значений таблицы
                for (int i = 0; i < allInfo.ValuesInfo.Count; i++)
                {
       
                    object value = allInfo.ValuesInfo[i].smpr_data_Info_Tables_Values_Value;
                    string rowCode = allInfo.ValuesInfo[i].smpr_data_Info_Tables_Values_RowCode;
                    var selectedRow = allInfo.RowsInfo.Where(p => p.smpr_data_Info_Tables_Rows_Code == rowCode).FirstOrDefault();
                    int rowIndex = allInfo.RowsInfo.IndexOf(selectedRow)+1;
                    string colCode = allInfo.ValuesInfo[i].smpr_data_Info_Tables_Values_ColCode;
                    var selectedCol = allInfo.ColumnsInfo.Where(p => p.smpr_data_Info_Tables_Columns_Code == colCode).FirstOrDefault();
                    int colIndex = allInfo.ColumnsInfo.IndexOf(selectedCol)+2;
                    form.dataGridView1[colIndex,rowIndex].Value = value;

                }
                //form.AddEmptyColumnsAndRows(form.dataGridView1, 10, 3);
                form.info = (AllTableInfo)allInfo.Clone();
                form.savedInfo = (AllTableInfo)allInfo.Clone();
                form.Show();
            }
            catch
            {
                isNoException = false;
            }
            return isNoException;

        }
        [DispId(2)]
        [ComVisible(true)]
        public bool OnClose(bool checkForChanges)
        {
            bool result = false;
            try
            {
                bool hasChangedItems = (!form.info.isEqual(form.savedInfo));

                if (hasChangedItems)
                {
                    var window = MessageBox.Show(
                                    "В таблице остались несохраненные данные. Сохранить?",
                                    "Сохранение данных",
                                    MessageBoxButtons.YesNoCancel);
                    if (window == DialogResult.Yes)
                    {
                        object?[] array = CollectTableData(form.info);
                        object? arrRows = array[0];
                        object? arrCols = array[1];
                        object? arrValues = array[2];
                        bool isNoException = InvokeSave(arrRows, arrCols, arrValues);
                        result=InvokeClose();
                        form.Close();

                    }
                    if (window == DialogResult.No)
                    {
                        result=InvokeClose();
                        form.Close();

                    }
                    if (window == DialogResult.Cancel)
                    {
                        result = false;
                    }

                }
                else
                {
                    result = InvokeClose();
                    form.Close();
                }
            }
            catch
            {
                result = false;
            }
            return result;

            
        }
        public object[] CollectTableData(AllTableInfo info) 
        {


            allInfoAfterChange = info;
            int rowsCount = 0;
            foreach(var row in allInfoAfterChange.RowsInfo)
            {
                if (row.smpr_data_Info_Tables_Rows_Code != "")
                {
                    rowsCount += 1;
                }
            }

            int columnsCount = 0;
            foreach (var col in allInfoAfterChange.ColumnsInfo)
            {
                if (col.smpr_data_Info_Tables_Columns_Code != "")
                {
                    columnsCount += 1;
                }
            }

            int valuesCount = 0;
            foreach (var val in allInfoAfterChange.ValuesInfo)
            {
                if (val.smpr_data_Info_Tables_Values_ColCode != "" && val.smpr_data_Info_Tables_Values_RowCode != "")
                {
                    valuesCount += 1;
                }
            }
            //Rows
            object[,] arrRows = new object[rowsCount, 6];
            int j = 0;
            foreach(var row in allInfoAfterChange.RowsInfo)
            {
                if (row.smpr_data_Info_Tables_Rows_Code == "")
                {
                    continue;
                }
                arrRows[j, 0] = row.smpr_data_Info_Tables_Rows_TableId;
                arrRows[j, 1] = row.smpr_data_Info_Tables_Rows_Code;
                arrRows[j, 2] = row.smpr_data_Info_Tables_Rows_Descriptor;
                arrRows[j, 3] = row.smpr_data_Info_Tables_Rows_Comment;
                arrRows[j, 4] = row.smpr_data_Info_Tables_Rows_Index;
                arrRows[j, 5] = row.smpr_data_Info_Tables_Rows_ModFlag;
                j++;
            }

            //Columns
            j = 0;
            object[,] arrCols = new object[columnsCount,8];
            foreach (var col in allInfoAfterChange.ColumnsInfo)
            {
                if (col.smpr_data_Info_Tables_Columns_Code == "")
                {
                    continue;
                }
                arrCols[j, 0] = col.smpr_data_Info_Tables_Columns_TableId;
                arrCols[j, 1] = col.smpr_data_Info_Tables_Columns_Code;
                arrCols[j, 2] = col.smpr_data_Info_Tables_Columns_Descriptor;
                arrCols[j, 3] = col.smpr_data_Info_Tables_Columns_DataType;
                arrCols[j, 4] = col.smpr_data_Info_Tables_Columns_Precision;
                arrCols[j, 5] = col.smpr_data_Info_Tables_Columns_Comment;
                arrCols[j, 6] = col.smpr_data_Info_Tables_Columns_Index;
                arrCols[j, 7] = col.smpr_data_Info_Tables_Columns_ModFlag;
                j++;
            }


            //Values
            j = 0;
            object[,] arrValues = new object[valuesCount,10];
            foreach (var val in allInfoAfterChange.ValuesInfo)
            {
                if (val.smpr_data_Info_Tables_Values_ColCode == "" || val.smpr_data_Info_Tables_Values_RowCode == "")
                {
                    valuesCount += 1;
                }
                arrValues[j,0]  = val.smpr_data_Info_Tables_Values_TableId;
                arrValues[j, 1] = val.smpr_data_Info_Tables_Values_CaseId;
                arrValues[j, 2] = val.smpr_data_Info_Tables_Values_PeriodId;
                arrValues[j, 3] = val.smpr_data_Info_Tables_Values_RowCode;
                arrValues[j, 4] = val.smpr_data_Info_Tables_Values_ColCode;
                arrValues[j, 5] = val.smpr_data_Info_Tables_Values_Value;
                arrValues[j, 6] = val.smpr_data_Info_Tables_Values_Formula;
                arrValues[j, 7] = val.smpr_data_Info_Tables_Values_Comment;
                arrValues[j, 8] = val.smpr_data_Info_Tables_Values_Format;
                arrValues[j, 9] = val.smpr_data_Info_Tables_Values_ModFlag;
                j++;
            }
           

            object?[] returnInfoArray = { arrRows, arrCols, arrValues };
            object returnInfo = (object)returnInfoArray;
            allInfoAfterChangeObject = returnInfo;
            return returnInfoArray;
        }
        
        [DispId(3)]
        [ComVisible(true)]
        public bool InvokeSave(object arrRows, object arrColumns, object arrValues)
        {
            bool isNoException;

            

            if (Save != null)
            {
                isNoException = Save(arrRows, arrColumns, arrValues);
            }
            else
            {
                isNoException = false;
            }
            return isNoException;
        }

        [DispId(13)]
        public void Activate()
        {

        }

        public bool InvokeClose()
        {
            bool isNoException;
            if (Close != null)
            {
                
                isNoException = Close();
            }
            else
            {
                isNoException = false;
            }
            return isNoException;
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
        public object ToStringOrNull(object str)
        {
            if (str != null)
            {
                return str;
            }
            else
            {
                return "";
            }
        }



    }

    


    public class AllTableInfo : ICloneable
    {
        public bool isEqual( AllTableInfo right)
        {
            bool isRowsEqual = true; 
            bool isColumnsEqual = true; 
            bool isValuesEqual = true;
            if (this.RowsInfo.Count == right.RowsInfo.Count)
            {
                for (int i = 0; i < this.RowsInfo.Count; i++)
                {
                    bool x1 = this.RowsInfo[i].smpr_data_Info_Tables_Rows_Code
                                  == right.RowsInfo[i].smpr_data_Info_Tables_Rows_Code ? true : false;
                    bool x2 = this.RowsInfo[i].smpr_data_Info_Tables_Rows_Comment
                                  == right.RowsInfo[i].smpr_data_Info_Tables_Rows_Comment ? true : false;
                    bool x3 = this.RowsInfo[i].smpr_data_Info_Tables_Rows_Descriptor
                                  == right.RowsInfo[i].smpr_data_Info_Tables_Rows_Descriptor ? true : false;
                    bool x4 = this.RowsInfo[i].smpr_data_Info_Tables_Rows_Index
                                  == right.RowsInfo[i].smpr_data_Info_Tables_Rows_Index ? true : false;
                    bool x5 = this.RowsInfo[i].smpr_data_Info_Tables_Rows_ModFlag
                                  == right.RowsInfo[i].smpr_data_Info_Tables_Rows_ModFlag ? true : false;
                    bool x6 = this.RowsInfo[i].smpr_data_Info_Tables_Rows_TableId
                                  == right.RowsInfo[i].smpr_data_Info_Tables_Rows_TableId ? true : false;
                    isRowsEqual = x1 && x2 && x3 && x4 && x5 && x6;
                    if (isRowsEqual == false)
                    {
                        break;
                    }
                }
            }
            else
            {
                isRowsEqual = false;
            }
            if (this.ValuesInfo.Count == right.ValuesInfo.Count)
            {
                for (int i = 0; i < this.ValuesInfo.Count; i++)
                {
                    bool x1 = this.ValuesInfo[i].smpr_data_Info_Tables_Values_CaseId
                                  == right.ValuesInfo[i].smpr_data_Info_Tables_Values_CaseId ? true : false;
                    bool x2 = this.ValuesInfo[i].smpr_data_Info_Tables_Values_ColCode
                                  == right.ValuesInfo[i].smpr_data_Info_Tables_Values_ColCode ? true : false;
                    bool x3 = this.ValuesInfo[i].smpr_data_Info_Tables_Values_Comment
                                  == right.ValuesInfo[i].smpr_data_Info_Tables_Values_Comment ? true : false;
                    bool x4 = this.ValuesInfo[i].smpr_data_Info_Tables_Values_Format
                                  == right.ValuesInfo[i].smpr_data_Info_Tables_Values_Format ? true : false;
                    bool x5 = this.ValuesInfo[i].smpr_data_Info_Tables_Values_Formula
                                  == right.ValuesInfo[i].smpr_data_Info_Tables_Values_Formula ? true : false;
                    bool x6 = this.ValuesInfo[i].smpr_data_Info_Tables_Values_ModFlag
                                  == right.ValuesInfo[i].smpr_data_Info_Tables_Values_ModFlag ? true : false;
                    bool x7 = this.ValuesInfo[i].smpr_data_Info_Tables_Values_PeriodId
                                  == right.ValuesInfo[i].smpr_data_Info_Tables_Values_PeriodId ? true : false;
                    bool x8 = (this.ValuesInfo[i].smpr_data_Info_Tables_Values_RowCode
                                  == right.ValuesInfo[i].smpr_data_Info_Tables_Values_RowCode);
                    bool x9 = (this.ValuesInfo[i].smpr_data_Info_Tables_Values_TableId
                                  == right.ValuesInfo[i].smpr_data_Info_Tables_Values_TableId);
                    bool x10 = (Convert.ToDouble(this.ValuesInfo[i].smpr_data_Info_Tables_Values_Value)
                                  == Convert.ToDouble(right.ValuesInfo[i].smpr_data_Info_Tables_Values_Value));
                    isValuesEqual = x1 && x2 && x3 && x4 && x5 && x6 && x7 && x8 && x9 && x10;
                    if (isValuesEqual == false)
                    {
                        break;
                    }
                }
            }
            else
            {
                isValuesEqual = false;
            }
            if (this.ColumnsInfo.Count == right.ColumnsInfo.Count)
            {
                for (int i = 0; i < this.ColumnsInfo.Count; i++)
                {
                    bool x1 = this.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Code
                                  == right.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Code ? true : false;
                    bool x2 = this.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Comment
                                  == right.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Comment ? true : false;
                    bool x3 = this.ColumnsInfo[i].smpr_data_Info_Tables_Columns_DataType
                                  == right.ColumnsInfo[i].smpr_data_Info_Tables_Columns_DataType ? true : false;
                    bool x4 = this.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Descriptor
                                  == right.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Descriptor ? true : false;
                    bool x5 = this.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Index
                                  == right.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Index ? true : false;
                    bool x6 = this.ColumnsInfo[i].smpr_data_Info_Tables_Columns_ModFlag
                                  == right.ColumnsInfo[i].smpr_data_Info_Tables_Columns_ModFlag ? true : false;
                    bool x7 = this.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Precision
                                  == right.ColumnsInfo[i].smpr_data_Info_Tables_Columns_Precision ? true : false;
                    bool x8 = this.ColumnsInfo[i].smpr_data_Info_Tables_Columns_TableId
                                  == right.ColumnsInfo[i].smpr_data_Info_Tables_Columns_TableId ? true : false;
                    isColumnsEqual = x1 && x2 && x3 && x4 && x5 && x6 && x7 && x8;
                    if (isColumnsEqual == false)
                    {
                        break;
                    }
                }
            }
            else
            {
                isColumnsEqual = false;
            }
            return (isRowsEqual && isColumnsEqual && isValuesEqual);
        }
        public AllTableInfo()
        {

        }
        public AllTableInfo(BindingList<DataRows> rowsInfo, BindingList<DataColumns> columnsInfo, BindingList<DataValues> valuesInfo)
        {
            RowsInfo = rowsInfo;
            ColumnsInfo = columnsInfo;
            ValuesInfo = valuesInfo;
        }

        public object Clone()
        {
            BindingList<DataRows> newRowsInfo = new BindingList<DataRows>();
            foreach (DataRows row in this.RowsInfo)
            {
                DataRows newRow = (DataRows)row.Clone();
                newRowsInfo.Add(newRow);
            }
            BindingList<DataColumns> newColumnsInfo = new BindingList<DataColumns>();
            foreach (DataColumns col in this.ColumnsInfo)
            {
                DataColumns newCol = (DataColumns)col.Clone();
                newColumnsInfo.Add(newCol);
            }
            BindingList<DataValues> newValsInfo = new BindingList<DataValues>();
            foreach (DataValues val in this.ValuesInfo)
            {
                DataValues newVal = (DataValues)val.Clone();
                newValsInfo.Add(newVal);
            }
            return new AllTableInfo(newRowsInfo, newColumnsInfo, newValsInfo);
        }

        public BindingList<DataRows> RowsInfo { get; set; }
        public BindingList<DataColumns> ColumnsInfo { get; set; }
        public BindingList<DataValues> ValuesInfo { get; set; }

    }

    public class DataRows: ICloneable
    {
        public object Clone() => MemberwiseClone();

        public int smpr_data_Info_Tables_Rows_TableId { get; set; }
        public string smpr_data_Info_Tables_Rows_Code { get; set; }
        public string smpr_data_Info_Tables_Rows_Descriptor { get; set; }
        public string smpr_data_Info_Tables_Rows_Comment { get; set; }
        public int smpr_data_Info_Tables_Rows_Index { get; set; }
        public int smpr_data_Info_Tables_Rows_ModFlag { get; set; }
    }
    public class DataColumns : ICloneable
    {
        public object Clone() => MemberwiseClone();

        public int smpr_data_Info_Tables_Columns_TableId { get; set; }
        public string smpr_data_Info_Tables_Columns_Code { get; set; }
        public string smpr_data_Info_Tables_Columns_Descriptor { get; set; }
        public int smpr_data_Info_Tables_Columns_DataType { get; set; }
        public int smpr_data_Info_Tables_Columns_Precision { get; set; }
        public string smpr_data_Info_Tables_Columns_Comment { get; set; }
        public int smpr_data_Info_Tables_Columns_Index { get; set; }
        public int smpr_data_Info_Tables_Columns_ModFlag { get; set; }

    }
    public class DataValues : ICloneable
    {
        public object Clone() => MemberwiseClone();

        public int smpr_data_Info_Tables_Values_TableId { get; set; }
        public int smpr_data_Info_Tables_Values_CaseId { get; set; }
        public int smpr_data_Info_Tables_Values_PeriodId { get; set; }
        public string smpr_data_Info_Tables_Values_RowCode { get; set; }
        public string smpr_data_Info_Tables_Values_ColCode { get; set; }
        public object smpr_data_Info_Tables_Values_Value { get; set; }
        public string smpr_data_Info_Tables_Values_Formula { get; set; }
        public string smpr_data_Info_Tables_Values_Comment { get; set; }
        public string smpr_data_Info_Tables_Values_Format { get; set; }
        public int smpr_data_Info_Tables_Values_ModFlag { get; set; }

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
    public static class Extensions
    {
        public static IList<T> Clone<T>(this IList<T> listToClone) where T : ICloneable
        {
            return listToClone.Select(item => (T)item.Clone()).ToList();
        }
    }

}
