using System;
using System.Text;
using System.Collections.Generic;
using ExcelKit.Core.Attributes;
using ExcelKit.Core.Infrastructure.Factorys;
using System.Threading.Tasks;
using Sample.Contract.WriteDtos;

namespace Sample.Consoles
{
	public class ExcelWriteWrapper
	{
		/// <summary>
		/// 导出用户数据（采用了并发多Sheet导出，一个线程一个Sheet，不需要可以去掉Parallel，采用单纯的for）
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
					var sheetName = $"Sheet{index}";
					var sheet = context.CrateSheet<UserExportDto>(sheetName);

					for (int i = 0; i < 1020; i++)
					{
						sheet.AppendData(sheetName, new UserExportDto { Account = $"{index}-{i}-2010211", Name = $"{index}-{i}-用户用户", IsConfirm = i % 2 == 0, IsMan = i % 2 == 0 });
					}
				});

				filePath = context.Save();
				Console.WriteLine($"文件路径：{filePath}");
			}

			return filePath;
		}

		/// <summary>
		/// 动态导出，不需要建立类，直接指定
		/// </summary>
		/// <returns></returns>
		public static string DynamicWrite()
		{
			string filePath;
			using (var context = ContextFactory.GetWriteContext($"用户数据-{DateTime.Now.ToString("yyyyMMddHHmm")}"))
			{
				//动态指定Code为字段名，自己定义，和AppendData中的数据字段名保持一致即可，Desc为导出的Excel列头名
				//注意CreateSheet方法最后一个字段，指定多少条数据自动拆分一个新Sheet，不指定默认为单Sheet最大数据量1048200
				var sheet = context.CrateSheet("Sheet1", new List<ExcelKitAttribute>()
				{
					new ExcelKitAttribute(){ Code = "Account", Desc = "账号",Width=60 },
					new ExcelKitAttribute(){ Code = "Name", Desc = "昵称" }
				});

				for (int i = 0; i < 104; i++)
				{
					//Dictionary中的Key为上面指定的Code中的字段，Value为数据
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
