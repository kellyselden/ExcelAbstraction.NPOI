using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelAbstraction.Entities;
using ExcelAbstraction.Services;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelAbstraction.NPOI
{
	public class ExcelService : IExcelService
	{
		public Workbook ReadWorkbook(string path)
		{
			if (!File.Exists(path)) return null;

			using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
				return CreateWorkbook(new HSSFWorkbook(stream));
		}

		static Workbook CreateWorkbook(IWorkbook workbook)
		{
			var worksheets = new List<Worksheet>();
			for (int i = 0; i < workbook.NumberOfSheets; i++)
				worksheets.Add(CreateWorksheet(workbook.GetSheetAt(i)));
			return new Workbook(worksheets);
		}

		static Worksheet CreateWorksheet(ISheet sheet)
		{
			var rows = new List<IRow>();
			int maxColumns = 0;
			for (int i = 0; i < sheet.PhysicalNumberOfRows; i++)
			{
				IRow row = sheet.GetRow(i);
				maxColumns = Math.Max(maxColumns, row.LastCellNum);
				rows.Add(row);
			}
			return new Worksheet(sheet.SheetName, maxColumns, rows.Select(row => CreateRow(row, maxColumns)));
		}

		static Row CreateRow(IRow row, int columns)
		{
			var cells = new List<Cell>();
			ICell[] iCells = row.Cells.ToArray();
			int skipped = 0;
			for (int i = 0; i < columns; i++)
			{
				Cell cell;
				if (i - skipped >= iCells.Length || i != iCells[i - skipped].ColumnIndex)
				{
					cell = new Cell(row.RowNum, i, null);
					skipped++;
				}
				else cell = CreateCell(iCells[i - skipped]);
				cells.Add(cell);
			}
			return new Row(cells);
		}

		static Cell CreateCell(ICell cell)
		{
			string value = null;
			switch (cell.CellType)
			{
				case CellType.String:
					value = cell.StringCellValue;
					break;
				case CellType.Numeric:
					value = cell.NumericCellValue.ToString();
					break;
				case CellType.Formula:
					switch (cell.CachedFormulaResultType)
					{
						case CellType.String:
							value = cell.StringCellValue;
							break;
						case CellType.Numeric:
							//excel trigger is probably out-of-date
							value = (cell.CellFormula == "TODAY()" ? DateTime.Today.ToOADate() : cell.NumericCellValue).ToString();
							break;
					}
					break;
			}
			return new Cell(cell.RowIndex, cell.ColumnIndex, value);
		}
	}
}