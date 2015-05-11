using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using ExcelAbstraction.Entities;
using ExcelAbstraction.Helpers;
using ExcelAbstraction.Services;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record;
using NPOI.HSSF.Record.Aggregates;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula;
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
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
			var worksheet = new Worksheet(sheet.SheetName, index, maxColumns, rows.Select(row => CreateRow(row, maxColumns)).ToArray());
			AddToValidations(worksheet.Validations, sheet);
			return worksheet;
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
				AddValidations(sheet, version, worksheet.Validations.ToArray());
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

		public object GetWorkbook(string path)
		{
			return WorkbookFactory.Create(path);
		}

		public object GetWorkbook(Stream stream)
		{
			return WorkbookFactory.Create(stream);
		}

		public void SaveWorkbook(object workbook, string path)
		{
			using (var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write))
				SaveWorkbook(workbook, stream);
		}

		public void SaveWorkbook(object workbook, Stream stream)
		{
			((IWorkbook)workbook).Write(stream);
		}

		static void AddToValidations(ICollection<Validation> validations, ISheet sheet)
		{
			var hssfSheet = sheet as HSSFSheet;
			if (hssfSheet != null)
			{
				AddToValidations(validations, hssfSheet);
			}
			else
			{
				var xssfSheet = sheet as XSSFSheet;
				if (xssfSheet != null)
				{
					AddToValidations(validations, xssfSheet);
				}
			}
		}

		static void AddToValidations(ICollection<Validation> validations, HSSFSheet sheet)
		{
			InternalSheet internalSheet = sheet.Sheet;
			var dataValidityTable = (DataValidityTable)internalSheet.GetType()
				.GetField("_dataValidityTable", BindingFlags.NonPublic | BindingFlags.Instance)
				.GetValue(internalSheet);
			if (dataValidityTable == null) return;

			var validationList = (IList)dataValidityTable.GetType()
				.GetField("_validationList", BindingFlags.NonPublic | BindingFlags.Instance)
				.GetValue(dataValidityTable);
			foreach (DVRecord record in validationList)
			{
				var formula = (Formula)record.GetType()
					.GetField("_formula1", BindingFlags.NonPublic | BindingFlags.Instance)
					.GetValue(record);
				validations.Add(new Validation
				{
					Range = ExcelHelper.ParseRange(record.CellRangeAddress.CellRangeAddresses[0].FormatAsString(), ExcelVersion.Xls),
					List = ((StringPtg)formula.Tokens[0]).Value.Split('\0')
				});
			}
		}

		static void AddToValidations(ICollection<Validation> validations, XSSFSheet sheet)
		{
			CT_DataValidations dataValidations = ((CT_Worksheet)sheet.GetType()
				.GetField("worksheet", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly)
				.GetValue(sheet)).dataValidations;
			if (dataValidations == null) return;

			foreach (CT_DataValidation dataValidation in dataValidations.dataValidation)
			{
				if (dataValidation.formula1 == null) continue;

				var range = ExcelHelper.ParseRange(dataValidation.sqref, ExcelVersion.Xlsx);
				if (range == null) continue;

				validations.Add(new Validation
				{
					Range = range,
					List = dataValidation.formula1.Trim('\"').Split(',')
				});
			}
		}

		public void AddValidations(object workbook, int sheetIndex, ExcelVersion version, params Validation[] validations)
		{
			AddValidations(((IWorkbook)workbook).GetSheetAt(sheetIndex), version, validations);
		}

		static void AddValidations(ISheet sheet, ExcelVersion version, params Validation[] validations)
		{
			IDataValidationHelper helper = sheet.GetDataValidationHelper();
			foreach (Validation validation in validations)
			{
				IDataValidationConstraint constraint = helper.CreateExplicitListConstraint(validation.List.ToArray());
				var range = new CellRangeAddressList(
					validation.Range.RowStart,
					validation.Range.RowEnd ?? ExcelHelper.GetRowMax(version),
					validation.Range.ColumnStart,
					validation.Range.ColumnEnd);
				IDataValidation dataValidation = helper.CreateValidation(constraint, range);
				sheet.AddValidationData(dataValidation);
			}
		}
	}
}