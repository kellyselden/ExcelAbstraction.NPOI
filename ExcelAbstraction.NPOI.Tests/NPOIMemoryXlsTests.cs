using System;
using ExcelAbstraction.Entities;
using ExcelAbstraction.Tests;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelAbstraction.NPOI.Tests
{
	[TestClass]
	public class NpoiMemoryXlsTests : ExcelServiceMemoryTests
	{
		public NpoiMemoryXlsTests() : base(new ExcelService(), ExcelVersion.Xls) { }

		[TestInitialize]
		public override void TestInitialize()
		{
			base.TestInitialize();
		}

		[TestCleanup]
		public override void TestCleanup()
		{
			base.TestCleanup();
		}

		[TestMethod]
		public override void ExcelService_AddValidations()
		{
			base.ExcelService_AddValidations();
		}

		[TestMethod]
		public override void ExcelService_AddValidations_Hack()
		{
			base.ExcelService_AddValidations_Hack();
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public override void ExcelService_AddValidations_EmptyListThrows()
		{
			base.ExcelService_AddValidations_EmptyListThrows();
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public override void ExcelService_AddValidations_NullListThrows()
		{
			base.ExcelService_AddValidations_NullListThrows();
		}

		[TestMethod]
		public override void ExcelService_Worksheet_IsHidden()
		{
			base.ExcelService_Worksheet_IsHidden();
		}

		[TestMethod]
		public override void ExcelService_AddValidations_LotsOfItems()
		{
			base.ExcelService_AddValidations_LotsOfItems();
		}
	}
}