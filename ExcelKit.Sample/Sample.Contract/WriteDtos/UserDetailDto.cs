using System;
using System.Collections.Generic;
using System.Text;

namespace Sample.Contract.WriteDtos
{
	/// <summary>
	/// 用户详情(默认值为演示使用，实际数据自己赋值)
	/// </summary>
	public class UserDetailDto
	{
		/// <summary>
		/// 年龄
		/// </summary>
		public int Age { get; set; } = 60 - DateTime.Now.Minute;

		/// <summary>
		/// 省份
		/// </summary>
		public string Province { get; set; } = "四川省";

		/// <summary>
		/// 城市
		/// </summary>
		public string City { get; set; } = "成都市";

		/// <summary>
		/// 电话号码
		/// </summary>
		public string PhoneNumber { get; set; } = "130xxxx3333";
	}
}
