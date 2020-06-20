using System;
using System.Text;
using System.Collections.Generic;
using ExcelKit.Core.Attributes;
using ExcelKit.Core.ExcelWrite;
using ExcelKit.Core.Infrastructure.Factorys;
using System.Threading.Tasks;
using Sample.Contract.WriteDtos;

namespace Sample.Consoles
{
	public class ExcelWriteWrapper
	{
		/// <summary>
		/// 导出用户数据
		/// </summary>
		/// <remarks>
		/// 1.获取GetWriteContext并指定导出文件名
		/// 2.创建Sheet并制定Sheet名（Sheet名作为后期追加数据区分是哪个Sheet的依据）
		/// 3.AppendData向Sheet中追加数据
		/// 4.调用Save保存（默认保存到程序运行目录）或Generate生成Excel信息，web环境调用Generate生成的信息，调用return File(Excel信息)后，可直接用于下载
		/// 5.特别提示，当单个Sheet数据量超过1048200后，后续追加的数据会自动拆分到新的Sheet，使用者不需要自己处理，只管追加数据
		/// </remarks>
		/// <returns>Excel文件路径</returns>
		public static string GenericWrite()
		{
			string filePath;
			using (var context = ContextFactory.GetWriteContext($"用户数据-{DateTime.Now.ToString("yyyyMMddHHmm")}"))
			{
				Parallel.For(0, 4, index =>
				{
					var sheet = context.CrateSheet<UserExportDto>($"Sheet{index}");
					for (int i = 0; i < 1020000; i++)
					{
						sheet.AppendData($"Sheet{index}", new UserExportDto { Account = $"{index}-{i}-2010211", Name = $"{index}-{i}-用户用户" });
					}
				});

				filePath = context.Save();
				Console.WriteLine($"文件路径：{filePath}");
			}

			return filePath;
		}

		public static string DynamicWrite()
		{
			string filePath;
			using (var context = ContextFactory.GetWriteContext($"用户数据-{DateTime.Now.ToString("yyyyMMddHHmm")}"))
			{
				var sheet = context.CrateSheet("Sheet1", new List<ExcelKitAttribute>()
				{
					new ExcelKitAttribute(){ Code = "Account", Desc = "账号",Width=60 },
					new ExcelKitAttribute(){ Code = "Name", Desc = "昵称" }
				}, 5);

				for (int i = 0; i < 10; i++)
				{
					sheet.AppendData("Sheet1", new Dictionary<string, object>()
					{
						{"Account", $"{i}-2010211" }, {"Name", $"{i}-用户用户" }
					});
				}

				filePath = context.Save();
				Console.WriteLine($"文件路径：{filePath}");
			}

			return filePath;
		}
	}
}
