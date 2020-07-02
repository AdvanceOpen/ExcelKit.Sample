using System;
using ExcelKit.Core.Attributes;

namespace Sample.Contract.ReadDtos
{
	/// <summary>
	/// 读取
	/// </summary>
	public class UserImportDto
	{
		[ExcelKit(Desc = "账号")]
		public string Account { get; set; }

		[ExcelKit(Desc = "昵称")]
		public string Name { get; set; }

		[ExcelKit(Desc = "金额")]
		public decimal Money { get; set; }

		[ExcelKit(Desc = "创建时间")]
		public DateTime CreateDate { get; set; }
	}
}
