using System;
using System.Text;
using System.Collections.Generic;
using ExcelKit.Core.ExcelWrite;
using ExcelKit.Core.ExcelRead;
using Newtonsoft.Json;
using ExcelKit.Core.Constraint.Enums;
using ExcelKit.Core.Infrastructure.Factorys;
using Sample.Contract.ReadDtos;
using ExcelKit.Core.Helpers;

namespace Sample.Consoles
{
    /// <summary>
    /// Excel读取
    /// </summary>
    public class ExcelReadWrapper
    {
        /// <summary>
        /// 获取Sheet中数据总行数
        /// </summary>
        public static void GetSheetRowCount()
        {
            //1.指定Sheet索引(从1开始)读取
            var count1 = ContextFactory.GetReadContext().ReadSheetRowsCount("用户数据.xlsx", new ReadSheetRowsCountOptions()
            {
                //可以不指定SheetIndex，默认就为1
                SheetIndex = 1,
                //可以不指定，默认为释放，当需要多次读取时，可指定不释放传false
                //比如对于反馈进度的，先读取总行数，再读取内容
                IsDisposeStream = true,
            });
            Console.WriteLine($"指定Sheet索引为1读取后的总行数为：{count1}");

            //2.指定Sheet名称读取
            var count2 = ContextFactory.GetReadContext().ReadSheetRowsCount("用户数据.xlsx", new ReadSheetRowsCountOptions()
            {
                //可以不指定SheetIndex，默认就为1
                SheetName = "Sheet2",
                //可以不指定，默认为释放，当需要多次读取时，可指定不释放传false
                //比如对于反馈进度的，先读取总行数，再读取内容
                IsDisposeStream = true,
            });
            Console.WriteLine($"指定Sheet名称为Sheet2读取后的总行数为：{count2}");
        }

        /// <summary>
        /// 读取Sheet中一行数据(如用来获取表头行)
        /// </summary>
        public static void ReadOneRow()
        {
            //sheetIndex为Sheet索引(从1开始)，rowLine为行号(从1开始)
            var headers = LiteDataHelper.ReadOneRow(filePath: "用户数据.xlsx", sheetIndex: 1, rowLine: 1);
            Console.WriteLine($"表头为：{string.Join("  ", headers)}");
        }

        /// <summary>
        /// 根据Sheet的索引读取行数据，默认SheetIndex为1，读取方式为根据Sheet索引
        /// </summary>
        public static void SheetIndexReadRows()
        {
            var context = ContextFactory.GetReadContext();
            context.ReadRows("用户数据.xlsx", new ReadRowsOptions()
            {
                RowData = rowdata =>
                {
                    Console.WriteLine(JsonConvert.SerializeObject(rowdata));
                }
            });
        }

        /// <summary>
        /// 根据Sheet名称读取行数据，返回一行数据IList<string>
        /// </summary>
        public static void SheetNameReadRows()
        {
            var context = ContextFactory.GetReadContext();
            context.ReadRows("用户数据.xlsx", new ReadRowsOptions()
            {
                ReadWay = ReadWay.SheetName,
                RowData = rowdata =>
                {
                    Console.WriteLine(JsonConvert.SerializeObject(rowdata));
                }
            });
        }

        /// <summary>
        /// 读取后转换为实体类(切记，此处的文件要存在且不可被占用。没有的话，自己可以生成一个)
        /// </summary>
        /// <remarks>更多读取项请查看ReadSheetOptions的定义，如读取结束行，按Sheet索引读取，按Sheet名称读取，读取开始行</remarks>
        public static void ReadSheet()
        {
            var context = ContextFactory.GetReadContext();
            context.ReadSheet("用户数据.xlsx", new ReadSheetOptions<UserImportDto>()
            {
                SucData = (rowdata, rowindex) =>
                {
                    Console.WriteLine(JsonConvert.SerializeObject(rowdata));
                },
                FailData = (odata, failinfo) =>
                {
                    //odata为Excel中的原始数据，FailInfo为失败相关信息
                }
            });
        }

        /// <summary>
        /// 根据指定的字典动态读取
        /// </summary>
        public static void ReadSheetDic()
        {
            ContextFactory.GetReadContext().ReadSheet("用户数据.xlsx", new ReadSheetDicOptions()
            {
                DataEndRow = 10,
                ExcelFields = new (string field, ColumnType type, bool allowNull)[]
                {
                    ("账号",ColumnType.String,false),("昵称",ColumnType.String,false)
                },
                SucData = (rowdata, rowindex) =>
                {
                    Console.WriteLine(JsonConvert.SerializeObject(rowdata));
                },
                FailData = (odata, failinfo) =>
                {
                    //odata为Excel中的原始数据，FailInfo为失败相关信息
                }
            });
        }
    }
}
