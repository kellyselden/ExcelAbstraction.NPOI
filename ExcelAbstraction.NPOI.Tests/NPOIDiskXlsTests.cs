using System.Linq;
using ExcelAbstraction.Entities;
using ExcelAbstraction.Tests;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelAbstraction.NPOI.Tests
{
	[TestClass]
	public class NpoiDiskXlsTests : ExcelServiceDiskTests
	{
		const string
			FileName = "ie_data.xls",
			DeploymentItem = "Excel/" + FileName;

		public NpoiDiskXlsTests() : base(new ExcelService(), FileName, ExcelVersion.Xls) { }

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

			Assert.AreEqual(3, worksheets.Length);
			Assert.AreEqual("Index Plot", worksheets[0].Name);
			Assert.AreEqual("PE (CAPE) Plot", worksheets[1].Name);
			Assert.AreEqual("Data", worksheets[2].Name);
		}

		[TestMethod]
		[DeploymentItem(DeploymentItem)]
		public void Worksheet_Rows()
		{
			Assert.AreEqual(2423, Workbook.Worksheets.Single(worksheet => worksheet.Name == "Data").Rows.Count());
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
			base.ExcelService_UsesCulture();
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