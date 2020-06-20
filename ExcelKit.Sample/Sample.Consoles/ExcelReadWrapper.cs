using System;
using System.Text;
using System.Collections.Generic;
using ExcelKit.Core.ExcelWrite;
using ExcelKit.Core.ExcelRead;
using Newtonsoft.Json;
using ExcelKit.Core.Constraint.Enums;
using ExcelKit.Core.Infrastructure.Factorys;
using Sample.Contract.ReadDtos;

namespace Sample.Consoles
{
	/// <summary>
	/// Excel读取
	/// </summary>
	public class ExcelReadWrapper
	{
		/// <summary>
		/// 根据Sheet的索引读取行数据，默认SheetIndex为1，读取方式为根据Sheet索引
		/// </summary>
		public static void SheetIndexReadRows()
		{
			var context = ContextFactory.GetReadContext();
			context.ReadRows("测试导出文件.xlsx", new ReadRowsOptions()
			{
				RowData = rowdata =>
				{
					Console.WriteLine(JsonConvert.SerializeObject(rowdata));
				}
			});
		}

		/// <summary>
		/// 根据Sheet名称读取行数据
		/// </summary>
		public static void SheetNameReadRows()
		{
			var context = ContextFactory.GetReadContext();
			context.ReadRows("测试导出文件.xlsx", new ReadRowsOptions()
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
			context.ReadSheet<UserImportDto>("用户数据-202006201252.xlsx", new ReadSheetOptions<UserImportDto>()
			{
				SucData = (rowdata, rowindex) =>
				{
					Console.WriteLine(JsonConvert.SerializeObject(rowdata));
				},
				FailData = (odata, failinfo) =>
				{

				}
			});
		}

		/// <summary>
		/// 根据指定的字典动态读取
		/// </summary>
		public static void ReadSheetDic()
		{
			var context = ContextFactory.GetReadContext();
			context.ReadSheet("测试导出文件.xlsx", new ReadSheetDicOptions()
			{
				DataEndRow = 10,
				ExcelFields = new (string field, ColumnType type, bool allowNull)[]
				{
					("账号",ColumnType.String,false),("昵称",ColumnType.String,false)
				},
				SucData = (rowdata, rowindex) =>
				{
					Console.WriteLine(JsonConvert.SerializeObject(rowdata));
				}
			});
		}
	}
}
