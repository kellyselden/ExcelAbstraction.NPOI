using System.Linq;
using ExcelAbstraction.Tests;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelAbstraction.NPOI.Tests
{
	[TestClass]
	public class NPOITests : ExcelTests
	{
		const string WorksheetName = "Data";

		public NPOITests() : base(new ExcelService(), Strings.FileName) { }

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
		[DeploymentItem(Strings.DeploymentItem)]
		public override void ExcelService_OpenWorkbook_FileNotFound_ReturnsNull()
		{
			base.ExcelService_OpenWorkbook_FileNotFound_ReturnsNull();
		}

		[TestMethod]
		[DeploymentItem(Strings.DeploymentItem)]
		public void Workbook_Worksheets()
		{
			var worksheets = Workbook.Worksheets.ToArray();

			Assert.AreEqual(3, worksheets.Length);
			Assert.AreEqual("Index Plot", worksheets[0].Name);
			Assert.AreEqual("PE (CAPE) Plot", worksheets[1].Name);
			Assert.AreEqual(WorksheetName, worksheets[2].Name);
		}

		[TestMethod]
		[DeploymentItem(Strings.DeploymentItem)]
		public void Worksheet_Rows()
		{
			Assert.AreEqual(2423, Workbook.Worksheets.Single(worksheet => worksheet.Name == WorksheetName).Rows.Count());
		}

		[TestMethod]
		[DeploymentItem(Strings.DeploymentItem)]
		public override void ExcelService_CheckGrid()
		{
			base.ExcelService_CheckGrid();
		}

		[TestMethod]
		[DeploymentItem(Strings.DeploymentItem)]
		public override void ExcelService_UsesCulture()
		{
			base.ExcelService_UsesCulture();
		}

		[TestMethod]
		[DeploymentItem(Strings.DeploymentItem)]
		public override void ExcelService_IgnoresThreadCulture()
		{
			base.ExcelService_IgnoresThreadCulture();
		}

		[TestMethod]
		[DeploymentItem(Strings.DeploymentItem)]
		public override void ExcelService_WriteWorkbook_Xls()
		{
			base.ExcelService_WriteWorkbook_Xls();
		}

		[TestMethod]
		[DeploymentItem(Strings.DeploymentItem)]
		public override void ExcelService_WriteWorkbook_Xlsx()
		{
			base.ExcelService_WriteWorkbook_Xlsx();
		}
	}
}