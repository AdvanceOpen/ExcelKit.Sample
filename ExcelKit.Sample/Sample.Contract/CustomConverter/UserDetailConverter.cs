using ExcelKit.Core.Infrastructure.Converter;
using Sample.Contract.WriteDtos;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sample.Contract.CustomConverter
{
	/// <summary>
	/// 用户详情Converter
	/// </summary>
	public class UserDetailConverter : IExportConverter<UserDetailDto>
	{
		public string Convert(UserDetailDto obj)
		{
			if (obj == null)
				return "";

			//自己实现逻辑，最终返回一个值即可，下述仅仅只是演示
			return $"{obj.Province}-{obj.City} ， 类别：{(obj.Age > 18 ? "成年人" : "未成年人")}， {obj.PhoneNumber}";
		}
	}
}
