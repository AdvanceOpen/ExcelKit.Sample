﻿using System;
using System.Collections.Generic;
using System.Text;
using ExcelKit.Core.Attributes;
using ExcelKit.Core.Constraint.Enums;
using ExcelKit.Core.Infrastructure.Converter;
using NPOI.XSSF.UserModel;
using Sample.Contract.CustomConverter;
using TextAlign = ExcelKit.Core.Constraint.Enums.TextAlign;

namespace Sample.Contract.WriteDtos
{
	public class UserExportDto
	{
		[ExcelKit(Desc = "账号", Width = 20, IsIgnore = false, Sort = 20, Align = TextAlign.Right, FontColor = DefineColor.LightBlue)]
		public string Account { get; set; }

		[ExcelKit(Desc = "昵称", Width = 50, Sort = 10, FontColor = DefineColor.Rose, ForegroundColor = DefineColor.LemonChiffon)]
		public string Name { get; set; }

		[ExcelKit(Desc = "金额", Width = 20, Sort = 20, Converter = typeof(DecimalPointDigitConverter), ConverterParam = 2)]
		public decimal Money { get; set; } = 20;

		//自定义Converter，示例待完善
		//[ExcelKit(Desc = "详情信息", Width = 60, Sort = 35, Converter = typeof(UserDetailConverter))]
		//public UserDetailDto UserDetail { get; set; }

		[ExcelKit(Desc = "是否确认", Width = 20, Sort = 30, Converter = typeof(BoolConverter), ConverterParam = "√|×")]
		public bool? IsConfirm { get; set; }

		[ExcelKit(Desc = "性别", Width = 20, Sort = 30, Converter = typeof(BoolConverter), ConverterParam = "男|女")]
		public bool? IsMan { get; set; }

		[ExcelKit(Desc = "创建时间", Width = 50, Sort = 40, Converter = typeof(DateTimeFmtConverter), ConverterParam = "yyyy-MM-dd")]
		public DateTime CreateDate { get; set; } = DateTime.Now;
	}
}
