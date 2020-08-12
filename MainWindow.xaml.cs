using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Excel;
using System.Data;

namespace secondApp
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow :System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public  class GlobalData  //一些共用的数据存储
        { 
            public static string gFileName;
            public static readonly string gMySheetName = "I46功能配置";
            public static readonly int gWantUseColumn = 5;
            public static readonly string gTestPointSheetName = "I46测试点";
            public static readonly int[] gTestPointRange = {2,485 };
        }

        private void chooseFileButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog openFileDialog = new OpenFileDialog();  // 实现.xlsx文件的选取
            openFileDialog.ShowDialog();
            string [] sArray = openFileDialog.FileName.Split('.');
            if (sArray[1] == "xlsx")
            {
                GlobalData.gFileName = fileName.Text = openFileDialog.FileName;
            }
            else {
                System.Windows.Forms.MessageBox.Show("请选择有效xlsx文件");
            }
            
        }

        private void Tool_Click(object sender, RoutedEventArgs e)
        {
            MSWord.Application EXCL1 = new MSWord.Application(); //新建一个应用程序EXC1
            EXCL1.Visible = true; //设置EXC1打开后可见

            MSWord.Workbooks wbs = EXCL1.Workbooks;
            MSWord._Workbook wb = wbs.Add(GlobalData.gFileName);   //打开EXCEL

            MSWord._Worksheet mySheet;

            mySheet = wb.Sheets[wantUseSheetIndex(GlobalData.gMySheetName, wb)]; //找到我要操作的sheet
            mySheet.Activate();
            //tiaoshi.Text = excelElementRead(mySheet, 1, 5).Interior.Color.ToString();
            System.Data.DataTable dt =  readNonredunOneColumn(mySheet,GlobalData.gWantUseColumn);

            MSWord._Worksheet testPointSheet = wb.Sheets[wantUseSheetIndex(GlobalData.gTestPointSheetName, wb)];
            testPointSheet.Activate();
            excelElementWrite(testPointSheet,1, 1, "测试点序号");
            excelElementWrite(testPointSheet, 1, 2, "测试点");
            for (int i =0;i< dt.Rows.Count;i++) {
                excelElementWrite(testPointSheet, i + 2, 1, dt.Rows[i][0].ToString());
                excelElementWrite(testPointSheet, i + 2, 2, dt.Rows[i][1].ToString());
            }

            writeTestPointCount(mySheet,GlobalData.gTestPointRange,GlobalData.gWantUseColumn, dt);

            wb.Save();



        }

        public void excelElementWrite(MSWord._Worksheet worksheet,  int row, int column, object val)
        {
            worksheet.Cells[row, column] = val;
        }
        public MSWord.Range excelElementRead(MSWord._Worksheet worksheet, int row, int column)
        {
            return ((MSWord.Range)worksheet.Cells[row, column]);
        }

        public System.Data.DataTable readNonredunOneColumn(MSWord._Worksheet worksheet, int column, string columnName1 = "TestPoitCount", string columnName2 = "TestPoitData") 
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add(columnName1,Type.GetType("System.Int32"));
            dt.Columns.Add(columnName2,Type.GetType("System.String"));  // 创建容器表

            int nullDataCount = 0;
            int rowIndexCount = 0;
            int dataArrayCount = 0;
            bool columnDataIsTrue = true;  //判断标志
            List<string> gTestPiotDataArray = new List<string> { };
            while (columnDataIsTrue) {
                rowIndexCount++;     
                if ((excelElementRead(worksheet, rowIndexCount, column).Text != "") && (excelElementRead(worksheet, rowIndexCount, column).Interior.Color != 14470546))
                {
                    dataArrayCount++; 
                    gTestPiotDataArray.Add(excelElementRead(worksheet, rowIndexCount, column).Text);
                    nullDataCount = 0;

                }
                else if (excelElementRead(worksheet, rowIndexCount, column).Text == "")
                {
                    nullDataCount++;
                    if (nullDataCount>=10) {
                        columnDataIsTrue = false;  //超过10行为空，则认为数据结束；
                    }
                }               
            }
            string[] str = gTestPiotDataArray.Distinct().ToArray();
            Array.Sort(str);
            for (int i=0;i<str.Length;i++) {    //将序号和数据存放在table中
                DataRow dr = dt.NewRow();
                dr[columnName1] = i + 1;
                dr[columnName2] = str[i];
                dt.Rows.Add(dr);
            }            
            return dt;
        }

        public Int32 wantUseSheetIndex(string indexStr, MSWord._Workbook workbook) 
        {
            int i = 0;
            foreach (MSWord._Worksheet fSheet in workbook.Sheets)
            {            //找到我要操作的sheet
                i++;
                if (fSheet.Name.ToString() == indexStr)
                {
                    break;
                }
            }
            return i;
        }

        public void writeTestPointCount(MSWord._Worksheet worksheet, int[] rowsRange,int column,System.Data.DataTable dt)
        {
            //frist read data from mysheet
            for (int i = rowsRange[0]; i <= rowsRange[1]; i++)
            {
                string elementValue =  excelElementRead(worksheet, i, column).Text;
                if (elementValue == "")
                {
                    continue;
                }
                
                DataRow[] dr = dt.Select("TestPoitData=" +"'"+elementValue+"'");

                excelElementWrite(worksheet, i, column-1, dr[0][0].ToString());
            }
            
        }
    }
}
