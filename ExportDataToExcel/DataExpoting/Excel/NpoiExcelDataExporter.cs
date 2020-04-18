using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExportDataToExcel.DataExpoting.Excel
{
    public class NpoiExcelDataExporter : IExcelExporter
    {

        #region Constants

        public const string HeaderCellStyleKey = "Header";
        public const string DateTimeCellStyleKey = "DateTime";

        public const short DefaultDateTimeFormat = 14;

        #endregion

        #region Methods

        public void ExportToStream(IEnumerable<ExportSheetInfo> exportSheets, Stream stream)
        {
            if (exportSheets is null)
            {
                throw new ArgumentNullException(nameof(exportSheets));
            }

            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            XSSFWorkbook workbook = new XSSFWorkbook();

            var styles = CreateStyles(workbook);
            foreach (ExportSheetInfo exportSheetInfo in exportSheets)
            {
                ExportSheet(workbook, styles, exportSheetInfo);
            }

            workbook.Write(stream, true);
        }

        private static void ExportSheet(IWorkbook workbook, IDictionary<string, ICellStyle> styles, ExportSheetInfo sheetInfo)
        {
            ISheet sheet = workbook.CreateSheet(sheetInfo.Header);

            List<PropertyMetadata> propertiesToExport = sheetInfo.Metadata
                .EnumeratePropertiesToExport()
                .ToList();

            FillHeaderRow(sheet, styles, propertiesToExport);

            int rowIndex = 1;
            foreach (object item in sheetInfo.Items)
            {
                FillItem(sheet, styles, propertiesToExport, item, rowIndex++);
            }

            for (int i = 0; i < propertiesToExport.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }
        }

        private static void FillHeaderRow(ISheet sheet, IDictionary<string, ICellStyle> styles, IList<PropertyMetadata> propertiesToExport)
        {
            IRow headerRow = sheet.CreateRow(0);

            for (int i = 0; i < propertiesToExport.Count; i++)
            {
                string propertyName = propertiesToExport[i].DisplayName;

                ICell cell = headerRow.CreateCell(i);
                cell.CellStyle = styles[HeaderCellStyleKey];
                cell.SetCellValue(propertyName);
            }
        }

        private static void FillItem(
            ISheet sheet,
            IDictionary<string, ICellStyle> styles,
            IList<PropertyMetadata> propertiesToExport,
            object item,
            int rowIndex)
        {
            if (item is null)
            {
                throw new ArgumentNullException(nameof(item));
            }

            IRow row = sheet.CreateRow(rowIndex);

            for (int i = 0; i < propertiesToExport.Count; i++)
            {
                PropertyMetadata property = propertiesToExport[i];

                ICell cell = row.CreateCell(i);

                object value = property.GetValue(item);
                if (value != null)
                {
                    Type valueType = value.GetType();
                    TypeCode typeCode = Type.GetTypeCode(valueType);

                    switch (typeCode)
                    {
                        case TypeCode.Boolean:
                            cell.SetCellValue((bool)value);
                            break;
                        case TypeCode.DateTime:
                            cell.CellStyle = styles[DateTimeCellStyleKey];
                            cell.SetCellValue((DateTime)value);
                            break;

                        case TypeCode.Byte:
                            cell.SetCellValue((byte)value);
                            break;
                        case TypeCode.Decimal:
                            cell.SetCellValue((double)(decimal)value);
                            break;
                        case TypeCode.Double:
                            cell.SetCellValue((double)value);
                            break;
                        case TypeCode.Int16:
                            cell.SetCellValue((short)value);
                            break;
                        case TypeCode.Int32:
                            cell.SetCellValue((int)value);
                            break;
                        case TypeCode.Int64:
                            cell.SetCellValue((long)value);
                            break;
                        case TypeCode.SByte:
                            cell.SetCellValue((sbyte)value);
                            break;
                        case TypeCode.Single:
                            cell.SetCellValue((float)value);
                            break;
                        case TypeCode.UInt16:
                            cell.SetCellValue((ushort)value);
                            break;
                        case TypeCode.UInt32:
                            cell.SetCellValue((uint)value);
                            break;
                        case TypeCode.UInt64:
                            cell.SetCellValue((ulong)value);
                            break;

                        default:
                            cell.SetCellValue(value.ToString());
                            break;
                    }
                }
            }
        }

        private IDictionary<string, ICellStyle> CreateStyles(IWorkbook workbook)
        {
            Dictionary<string, ICellStyle> styles = new Dictionary<string, ICellStyle>();

            ICellStyle headerCellStyle = workbook.CreateCellStyle();
            styles.Add(HeaderCellStyleKey, headerCellStyle);

            var headerFont = workbook.CreateFont();
            headerFont.FontHeightInPoints = 11;
            headerFont.IsBold = true;
            headerCellStyle.SetFont(headerFont);

            ICellStyle dateTimeCellStyle = workbook.CreateCellStyle();
            styles.Add(DateTimeCellStyleKey, dateTimeCellStyle);

            dateTimeCellStyle.DataFormat = DefaultDateTimeFormat;

            return styles;
        }

        #endregion

    }
}
