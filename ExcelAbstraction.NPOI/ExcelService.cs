using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using ExcelAbstraction.Entities;
using ExcelAbstraction.Services;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelAbstraction.NPOI
{
	public class ExcelService : IExcelService
	{
		public IFormatProvider Format { get; set; }

		public ExcelService()
		{
			Format = NumberFormatInfo.CurrentInfo;
		}

		public Workbook ReadWorkbook(string path)
		{
			if (!File.Exists(path)) return null;

			using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
				return ReadWorkbook(stream);
		}

		public Workbook ReadWorkbook(Stream stream)
		{
			return CreateWorkbook(WorkbookFactory.Create(stream));
		}

		Workbook CreateWorkbook(IWorkbook workbook)
		{
			var worksheets = new List<Worksheet>();
			for (int i = 0; i < workbook.NumberOfSheets; i++)
				worksheets.Add(CreateWorksheet(workbook.GetSheetAt(i), i));
			return new Workbook(worksheets);
		}

		Worksheet CreateWorksheet(ISheet sheet, int index)
		{
			var rows = new List<IRow>();
			int maxColumns = 0;
			for (int i = 0; i <= sheet.LastRowNum; i++)
			{
				IRow row = sheet.GetRow(i);
				if (row == null) continue;
				maxColumns = Math.Max(maxColumns, row.LastCellNum);
				rows.Add(row);
			}
			return new Worksheet(sheet.SheetName, index, maxColumns, rows.Select(row => CreateRow(row, maxColumns)).ToArray());
		}

		Row CreateRow(IRow row, int columns)
		{
			var cells = new List<Cell>();
			ICell[] iCells = row.Cells.ToArray();
			int skipped = 0;
			for (int i = 0; i < columns; i++)
			{
				Cell cell = null;
				if (i - skipped >= iCells.Length || i != iCells[i - skipped].ColumnIndex)
					skipped++;
				else cell = CreateCell(iCells[i - skipped]);
				cells.Add(cell);
			}
			return new Row(row.RowNum, cells);
		}

		Cell CreateCell(ICell cell)
		{
			string value = null;
			switch (cell.CellType)
			{
				case CellType.String:
					value = cell.StringCellValue;
					break;
				case CellType.Numeric:
					value = cell.NumericCellValue.ToString(Format);
					break;
				case CellType.Boolean:
					value = cell.BooleanCellValue.ToString();
					break;
				case CellType.Formula:
					switch (cell.CachedFormulaResultType)
					{
						case CellType.String:
							value = cell.StringCellValue;
							break;
						case CellType.Numeric:
							//excel trigger is probably out-of-date
							value = (cell.CellFormula == "TODAY()" ? DateTime.Today.ToOADate() : cell.NumericCellValue).ToString(Format);
							break;
					}
					break;
			}
			return new Cell(cell.RowIndex, cell.ColumnIndex, value);
		}

		public void WriteWorkbook(Workbook workbook, ExcelVersion version, string path)
		{
			using (Stream stream = File.Create(path))
				WriteWorkbook(workbook, version, stream);
		}
		public void WriteWorkbook(Workbook workbook, ExcelVersion version, Stream stream)
		{
			CreateWorkbook(workbook, version).Write(stream);
		}

		static IWorkbook CreateWorkbook(Workbook workbook, ExcelVersion version)
		{
			IWorkbook iWorkbook;
			switch (version)
			{
				case ExcelVersion.Xls:
					iWorkbook = new HSSFWorkbook();
					break;
				case ExcelVersion.Xlsx:
					iWorkbook = new XSSFWorkbook();
					break;
				default: throw new InvalidEnumArgumentException("version", (int)version, version.GetType());
			}

			foreach (Worksheet worksheet in workbook.Worksheets)
			{
				ISheet sheet = iWorkbook.CreateSheet(worksheet.Name);
				foreach (Row row in worksheet.Rows)
				{
					IRow iRow = sheet.CreateRow(row.Index);
					foreach (Cell cell in row.Cells)
						if (cell != null)
						{
							ICell iCell = iRow.CreateCell(cell.ColumnIndex);
							if (cell.Value != null)
								iCell.SetCellValue(cell.Value);
						}
				}
			}

			return iWorkbook;
		}
	}
}