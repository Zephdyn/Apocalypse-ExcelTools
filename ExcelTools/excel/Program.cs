using System;
using System.IO;
using System.Windows;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;

namespace excelTools.excel
{
    class Program
    {
        public static List<string[]> arrWrite = [];
        public static List<object[]> arrShee =[];
        /// <summary>
        /// 打开 excel, 完成后会主动关闭
        /// </summary>
        /// <param name="version">传入打开版本(目前只有v1,请传入1</param>
        /// <param name="type">类型1 xls, 2xlsx</param>
        /// <param name="sutype">读取类型枚举</param>
        /// <param name="fileName">文件路径</param>
        /// <param name="sheetId">打开表ID</param>
        public static bool OpenData(int version, int type, string[] sutype, string fileName, int sheetId)
        {
            File.Copy(fileName, fileName + ".temp", true);
            fileName += ".temp";
            arrWrite.Clear(); // 将arrWrite清空
            arrShee.Clear();// 将表清空
            FileStream fs = null;
            IWorkbook wk = null;
            if (type == 1)
            {
                try
                {
                    fs = File.OpenRead(fileName);// 打开文件
                    wk = new HSSFWorkbook(fs);// 连接xls
                }
                catch
                {
                    MessageBox.Show("打开xls失败");
                    fs?.Close();
                    wk?.Close();
                    return false;
                }
            }
            else if (type == 2)
            {
                try
                {
                    wk = new XSSFWorkbook(fileName);// 连接xlsx
                }
                catch
                {
                    MessageBox.Show("打开xlsx失败");
                    wk?.Close();
                    return false;
                }

            }
            else
            {
                MessageBox.Show("文件类型错误");
                return false;
            }
            int sheetCount = wk.NumberOfSheets;//获取表的数量
            if (sheetCount == 0)
            {
                MessageBox.Show("未获取到表名");
                return false;
            }
            if (sheetCount < sheetId + 1)
            {
                MessageBox.Show("超出表数量上限");
                return false;
            }
            string[] sheetName = new string[sheetCount];// 获取所有表名
            arrShee.Add([sheetId, wk.GetSheetName(sheetId)]);
            for (int i = 0; i < sheetCount; i++)
            {
                sheetName[i] = wk.GetSheetName(i);
                if (i == sheetId) continue;
                arrShee.Add([i, sheetName[i]]);
            }

            // 获取sheetId
            var st = wk.GetSheetAt(sheetId);
            // 进行数据读取
            if (version == 1)
            {
                Read_V1(sutype, st);
            }
            // 进行数据写入
            // 关闭打开的文件
            fs?.Close();
            // 关闭workbook
            wk?.Close();
            File.Delete(fileName);
            return true;

        }
        /// <summary>
        /// [v1]读取 excel 数据
        /// </summary>
        /// <param name="sutype"></param>
        /// <param name="st"></param>
        /// <returns></returns>
        public static int Read_V1(string[] sutype, ISheet st)
        {
            // 类数组
            string sTemp;
            var arr = sutype;
            // 读取失败返回信息
            if (arr == null) return 0;
            // 获取总列数
            // 读取第一行
            IRow firstRow = st.GetRow(0);
            // 第一行无数据, 不进行读取
            if (firstRow == null) return 0;
            // 获取数量
            int rowConut = firstRow.LastCellNum;

            List<ExcelType.RCache> arrReadCell = [];
            for (var i = 0; i < rowConut; i++)
            {
                var currentCell = st.GetRow(0).GetCell(i);
                // 防止报错
                if (currentCell == null) continue;
                // 不读取string以外的内容
                if (currentCell.CellType != CellType.String) continue;
                int classId = Array.IndexOf(arr, currentCell.StringCellValue);
                if (classId == -1) continue;
                ExcelType.RCache temp;
                temp._rowId = i;
                temp._classID = classId;
                arrReadCell.Add(temp);
            }
            ExcelType.RCache[] arrCell = [.. arrReadCell];
            // 获取总行数
            rowConut = st.LastRowNum;
            int numConut = arrCell.Length;
            //List<ExcelType.WCache> arrWrite = new List<ExcelType.WCache>();
            int arrLength = arr.Length + 1;
            // 从第二行开始读取
            for (var i = 1; i < rowConut + 1; i++)
            {
                // 防止null报错
                if (st.GetRow(i) == null) continue;
                string[] wTemp = new string[arrLength];
                for (var j = 0; j < arrLength; j++)
                {
                    // 给所有数值赋值空,防止出现null
                    wTemp[j] = "";
                }

                wTemp[0] = i.ToString();
                for (var j = 0; j < numConut; j++)
                {
                    // 读取数据
                    var currentCell = st.GetRow(i).GetCell(arrCell[j]._rowId);
                    if (currentCell == null)
                    {
                        sTemp = "";
                    }
                    else if (currentCell.CellType == CellType.String)
                    {
                        sTemp = currentCell.StringCellValue;
                    }
                    else if (currentCell.CellType == CellType.Numeric)
                    {
                        sTemp = currentCell.NumericCellValue.ToString();
                    }
                    else if (currentCell.CellType == CellType.Formula)
                    {
                        //currentCell.NumericCellValue.ToString();
                        currentCell.SetCellType(CellType.String);
                        sTemp = currentCell.StringCellValue;
                    }
                    else
                    {
                        sTemp = "";
                    }

                    // 写入需要
                    // 1. class_id // 类的ID
                    // 2. class_value // 读取的内容储存
                    wTemp[arrCell[j]._classID + 1] = sTemp;
                    /*ExcelType.WCache temp;
                    temp._classID = arrCell[j]._classID;
                    temp._value = sTemp;
                    arrWrite.Add(temp);*/
                }
                // 防止数组不足,报错
                if (wTemp.Length < 2) continue;
                if (wTemp[1] != "")
                    arrWrite.Add(wTemp);
            }
            // 数据读取完毕,返回读取到的数组

            //WCache = arrWrite.ToArray();
            return arrWrite.Count;
        }
    }

}
