using System;
using Sample.Contract.ReadDtos;

namespace Sample.Consoles
{
	class Program
	{
		/// <summary>
		/// 1.两个方法所在的类里面还有其他使用示例
		/// 2.如需测试导出，请放开第一个注释，注释第二个测试；测试读取注释第一个，放开第二个
		/// 3.读取的ReadSheet里面，新下载下来的没有Excel文件，读取会报错的，因为读取的文件为：用户数据-202006201252.xlsx
		///   是我用于测试的，可以先执行导出，然后里面指定文件名，进行读取测试
		/// </summary>
		/// <param name="args"></param>
		static void Main(string[] args)
		{
			//1.泛型类导出
			ExcelWriteWrapper.GenericWrite();

			//2.泛型类读取
			//ExcelReadWrapper.ReadSheet();

			//3.获取总行数
			//ExcelReadWrapper.GetSheetRowCount();

			//4.读取一行
			//ExcelReadWrapper.ReadOneRow();
			Console.Read();
		}
	}
}
