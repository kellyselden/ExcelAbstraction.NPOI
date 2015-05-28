using System.Linq;
using ExcelAbstraction.Entities;
using ExcelAbstraction.Tests;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelAbstraction.NPOI.Tests
{
	[TestClass]
	public class NpoiDiskXlsxTests : ExcelServiceDiskTests
	{
		const string
			FileName = "worksheet functions.xlsx",
			DeploymentItem = "Excel/" + FileName;

		public NpoiDiskXlsxTests() : base(new ExcelService(), FileName, ExcelVersion.Xlsx) { }

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
		[DeploymentItem(DeploymentItem)]
		public override void ExcelService_OpenWorkbook_FileNotFound_ReturnsNull()
		{
			base.ExcelService_OpenWorkbook_FileNotFound_ReturnsNull();
		}

		[TestMethod]
		[DeploymentItem(DeploymentItem)]
		public void Workbook_Worksheets()
		{
			var worksheets = Workbook.Worksheets.ToArray();

			Assert.AreEqual(2, worksheets.Length);
			Assert.AreEqual("Sheet1", worksheets[0].Name);
			Assert.AreEqual("Sheet2", worksheets[1].Name);
		}

		[TestMethod]
		[DeploymentItem(DeploymentItem)]
		public void Worksheet_Rows()
		{
			Assert.AreEqual(341, Workbook.Worksheets.Single(worksheet => worksheet.Name == "Sheet1").Rows.Count());
		}

		[TestMethod]
		[DeploymentItem(DeploymentItem)]
		public override void ExcelService_CheckGrid()
		{
			base.ExcelService_CheckGrid();
		}

		[TestMethod]
		[DeploymentItem(DeploymentItem)]
		public override void ExcelService_UsesCulture()
		{
			//base.ExcelService_UsesCulture();
		}

		[TestMethod]
		[DeploymentItem(DeploymentItem)]
		public override void ExcelService_IgnoresThreadCulture()
		{
			base.ExcelService_IgnoresThreadCulture();
		}

		[TestMethod]
		[DeploymentItem(DeploymentItem)]
		public override void ExcelService_WriteWorkbook()
		{
			base.ExcelService_WriteWorkbook();
		}
	}
}