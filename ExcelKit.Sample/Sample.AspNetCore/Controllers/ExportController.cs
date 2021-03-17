using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using ExcelKit.Core.Attributes;
using ExcelKit.Core.Helpers;
using ExcelKit.Core.Infrastructure.Factorys;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Sample.AspNetCore.Models;
using Sample.Contract.WriteDtos;

namespace Sample.AspNetCore.Controllers
{
    /// <summary>
    /// Web导出示例（导出字段定义参照UserExportDto）
    /// </summary>
    /// <remarks>
    /// 1.大数据量的导出，集合中的数据会占用内存，请使用标准的导出方式，参照GenericExport的方式
    /// 2.如果数据量不大但是又想简单的使用，可以直接使用内置的LiteDataHelper中的方法进行导出
    /// 3.多Sheet导出参照Sample.Consoles，本示例中仅演示单Sheet
    /// 4.测试的话请使用Ctrl+F5运行
    /// </remarks>
    public class ExportController : Controller
    {
        private const int ExportCount = 104000;
        private readonly ILogger<ExportController> _logger;

        public ExportController(ILogger<ExportController> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// 泛型实体类Web导出
        /// </summary>
        /// <returns></returns>
        public IActionResult GenericExport()
        {
            using (var context = ContextFactory.GetWriteContext("用户数据"))
            {
                for (int i = 1; i <= ExportCount; i++)
                {
                    var sheet = context.CrateSheet<UserExportDto>("Sheet1");
                    sheet.AppendData("Sheet1", new UserExportDto { Account = $"2020-{i}", Name = $"测试用户-{i}" });
                }
                var excelInfo = context.Generate();
                return File(excelInfo.Stream, excelInfo.WebContentType, excelInfo.FileName);
            }
        }

        /// <summary>
        /// 泛型实体类导出并保存到本地
        /// </summary>
        /// <returns></returns>
        public IActionResult GenericSaveToDisk()
        {
            using (var context = ContextFactory.GetWriteContext("用户数据"))
            {
                for (int i = 1; i <= ExportCount; i++)
                {
                    var sheet = context.CrateSheet<UserExportDto>("Sheet1");
                    sheet.AppendData("Sheet1", new UserExportDto { Account = $"2020-{i}", Name = $"测试用户-{i}" });
                }

                //保存路径不指定默认为程序运行目录
                context.Save();
            }
            return Ok();
        }

        /// <summary>
        /// 非大批量数据便捷导出
        /// </summary>
        /// <returns></returns>
        public IActionResult LiteDataExport()
        {
            //此方式对象全部在内存中，故数据量大的时候会占用内存，适合数据量不大使用；大数据量不占用内存的请采用上述的AppendData方式
            var users = Enumerable.Range(1, ExportCount).Select(index => new UserExportDto { Account = $"2021-{index}", Name = $"测试用户-{index}", IsMan = true, IsConfirm = true }).ToList();
            var excelInfo = LiteDataHelper.ExportToWebDown(users, fileName: "用户数据");
            return File(excelInfo.Stream, excelInfo.WebContentType, excelInfo.FileName);
        }

        /// <summary>
        /// 动态数据导出
        /// </summary>
        /// <returns></returns>
        public IActionResult DynamicDataExport()
        {
            using (var context = ContextFactory.GetWriteContext($"用户数据-{DateTime.Now.ToString("yyyyMMddHHmm")}"))
            {
                //动态指定Code为字段名，自己定义，和AppendData中的数据字段名保持一致即可，Desc为导出的Excel列头名
                //注意CreateSheet方法最后一个字段，指定多少条数据自动拆分一个新Sheet，不指定默认为单Sheet最大数据量1048200
                var sheet = context.CrateSheet("Sheet1", new List<ExcelKitAttribute>()
                {
                    new ExcelKitAttribute(){ Code = "Account", Desc = "账号",Width = 60 },
                    new ExcelKitAttribute(){ Code = "Name", Desc = "昵称" }
                });

                for (int i = 1; i <= ExportCount; i++)
                {
                    //Dictionary中的Key为上面指定的Code中的字段，Value为数据
                    sheet.AppendData("Sheet1", new Dictionary<string, object>()
                    {
                        {"Account", $"2020-{i}" }, {"Name",  $"测试用户-{i}" }
                    });
                }

                var excelInfo = context.Generate();
                return File(excelInfo.Stream, excelInfo.WebContentType, excelInfo.FileName);
            }
        }
    }
}
