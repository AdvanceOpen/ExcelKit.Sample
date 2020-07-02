using System;
using Sample.Contract.ReadDtos;

namespace Sample.Consoles
{
	class Program
	{
		/// <summary>
		/// 如需测试导出，请放开第一个注释，注释第二个测试；测试读取注释第一个，放开第二个
		/// 两个方法所在的类里面还有其他使用示例
		/// </summary>
		/// <param name="args"></param>
		static void Main(string[] args)
		{
			//1.泛型类导出
			//ExcelWriteWrapper.GenericWrite();

			//2.泛型类读取
			ExcelReadWrapper.ReadSheet();
			Console.Read();
		}
	}
}
