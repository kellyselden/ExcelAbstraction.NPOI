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

		Workbook CreateWorkbook(IWorkbook iWorkbook)
		{
			var worksheets = new List<Worksheet>();
			var workbook = new Workbook(worksheets);
			AddToNames(workbook.Names, iWorkbook);
			string[] names = workbook.Names.Select(name => name.Name).ToArray();
			for (int i = 0; i < iWorkbook.NumberOfSheets; i++)
			{
				ISheet sheet = iWorkbook.GetSheetAt(i);
				Worksheet worksheet = CreateWorksheet(sheet, i);
				worksheet.IsHidden = iWorkbook.IsSheetHidden(i);
				AddToValidations(worksheet.Validations, sheet, names);
				worksheets.Add(worksheet);
			}
			return workbook;
		}

		Worksheet CreateWorksheet(ISheet sheet, int index)
		{
			var rows = new List<IRow>();
			int maxColumns = 0;
			for (int i = 0; i <= sheet.LastRowNum; i++)
			{
				IRow row = sheet.GetRow(i);
				if (row != null)
					maxColumns = Math.Max(maxColumns, row.LastCellNum);
				rows.Add(row);
			}
			return new Worksheet(sheet.SheetName, index, maxColumns, rows.Select(row => CreateRow(row, maxColumns)).ToArray());
		}

		Row CreateRow(IRow row, int columns)
		{
			if (row == null) return null;

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

			AddNames(iWorkbook, version, workbook.Names.ToArray());
			foreach (Worksheet worksheet in workbook.Worksheets)
			{
				ISheet sheet = iWorkbook.CreateSheet(worksheet.Name);
				AddValidations(sheet, version, worksheet.Validations.ToArray());
				AddRows(sheet, worksheet.Rows.ToArray());
				if (worksheet.IsHidden)
					iWorkbook.SetSheetHidden(worksheet.Index, SheetState.Hidden);
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

		static void AddToNames(ICollection<NamedRange> names, IWorkbook workbook)
		{
			string propName;
			ExcelVersion version;
			if (workbook as HSSFWorkbook != null)
			{
				propName = "names";
				version = ExcelVersion.Xls;
			}
			else
			{
				if (workbook as XSSFWorkbook != null)
				{
					propName = "namedRanges";
					version = ExcelVersion.Xlsx;
				}
				else return;
			}

			var namedRanges = ((IList)workbook.GetType()
				.GetField(propName, BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly)
				.GetValue(workbook));

			foreach (IName name in namedRanges)
			{
				if (name.RefersToFormula.Contains("://")) continue;

				Range range = ExcelHelper.ParseRange(name.RefersToFormula, version);
				if (range == null) continue;

				names.Add(new NamedRange
				{
					Name = name.NameName,
					Range = range
				});
			}
		}

		static void AddToValidations(ICollection<DataValidation> validations, ISheet sheet, string[] names)
		{
			var hssfSheet = sheet as HSSFSheet;
			if (hssfSheet != null)
			{
				AddToValidations(validations, hssfSheet, names);
			}
			else
			{
				var xssfSheet = sheet as XSSFSheet;
				if (xssfSheet != null)
				{
					AddToValidations(validations, xssfSheet, names);
				}
			}
		}

		static void AddToValidations(ICollection<DataValidation> validations, HSSFSheet sheet, string[] names)
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

				var validation = new DataValidation
				{
					Range = ExcelHelper.ParseRange(record.CellRangeAddress.CellRangeAddresses[0].FormatAsString(), ExcelVersion.Xls)
				};

				Ptg ptg = formula.Tokens[0];
				var namePtg = ptg as NamePtg;
				if (namePtg != null)
				{
					validation.Type = DataValidationType.Formula;
					validation.Name = names.ElementAt(namePtg.Index);
				}
				else
				{
					var stringPtg = ptg as StringPtg;
					if (stringPtg != null)
					{
						validation.Type = DataValidationType.List;
						validation.List = stringPtg.Value.Split('\0');
					}
					else continue;
				}

				validations.Add(validation);
			}
		}

		static void AddToValidations(ICollection<DataValidation> validations, XSSFSheet sheet, string[] names)
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

				var validation = new DataValidation { Range = range };

				if (names.Contains(dataValidation.formula1))
				{
					validation.Type = DataValidationType.Formula;
					validation.Name = dataValidation.formula1;
				}
				else
				{
					validation.Type = DataValidationType.List;
					validation.List = dataValidation.formula1.Trim('\"').Split(',');
				}

				validations.Add(validation);
			}
		}

		public void AddNames(object workbook, ExcelVersion version, params NamedRange[] names)
		{
			AddNames((IWorkbook)workbook, version, names);
		}

		static void AddNames(IWorkbook workbook, ExcelVersion version, params NamedRange[] names)
		{
			foreach (NamedRange namedRange in names)
			{
				IName name = workbook.CreateName();
				name.NameName = namedRange.Name;
				name.RefersToFormula = ExcelHelper.RangeToString(namedRange.Range, version);
			}
		}

		public void AddRows(object workbook, string sheetName, params Row[] rows)
		{
			AddRows(((IWorkbook)workbook).GetSheet(sheetName), rows);
		}

		static void AddRows(ISheet sheet, params Row[] rows)
		{
			foreach (Row row in rows)
			{
				if (row == null) continue;

				IRow iRow = sheet.CreateRow(row.Index);
				foreach (Cell cell in row.Cells)
				{
					if (cell == null) continue;

					ICell iCell = iRow.CreateCell(cell.ColumnIndex);
					if (cell.Value != null)
						iCell.SetCellValue(cell.Value);
				}
			}
		}

		public void AddValidations(object workbook, string sheetName, ExcelVersion version, params DataValidation[] validations)
		{
			AddValidations(((IWorkbook)workbook).GetSheet(sheetName), version, validations);
		}

		static void AddValidations(ISheet sheet, ExcelVersion version, params DataValidation[] validations)
		{
			IDataValidationHelper helper = sheet.GetDataValidationHelper();
			foreach (DataValidation validation in validations)
			{
				if ((validation.List == null || validation.List.Count == 0) && validation.Name == null)
				{
					throw new InvalidOperationException("Validation is invalid");
				}

				IDataValidationConstraint constraint = validation.Name != null ?
					helper.CreateFormulaListConstraint(validation.Name) :
					helper.CreateExplicitListConstraint(validation.List.ToArray());

				var range = new CellRangeAddressList(
					validation.Range.RowStart ?? 0,
					validation.Range.RowEnd ?? ExcelHelper.GetRowMax(version) - 1,
					validation.Range.ColumnStart ?? 0,
					validation.Range.ColumnEnd ?? ExcelHelper.GetColumnMax(version) - 1);

				IDataValidation dataValidation = helper.CreateValidation(constraint, range);
				sheet.AddValidationData(dataValidation);
			}
		}
	}
}