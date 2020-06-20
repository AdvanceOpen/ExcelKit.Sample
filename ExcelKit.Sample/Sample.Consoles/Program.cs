using System;
using Sample.Contract.ReadDtos;

namespace Sample.Consoles
{
	class Program
	{
		static void Main(string[] args)
		{
			//ExcelWriteWrapper.GenericWrite();

			ExcelReadWrapper.ReadSheet();
			Console.Read();
		}
	}
}
