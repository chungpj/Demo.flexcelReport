using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Reflection;
using System.Text;
using FlexCel.Report;
using FlexCel.XlsAdapter;
using FlexCel.Core;

namespace Report.Common
{
    public static class FlexCelUtils
    {
        #region Common

        public static void AddTable<T>(this FlexCel.Report.FlexCelReport report, string tableName, IEnumerable<T> dataTable, int dataCount)
        {
            report.AddTable(tableName, new TSqlDataTable<T>(tableName, null, dataTable.AsQueryable(), dataCount), TDisposeMode.DisposeAfterRun);
        }

        public static void NewGenericTemplate(this XlsFile xlsFile, string name = "Data")
        {
            xlsFile.NewFile(1);
            xlsFile.ActiveSheet = 1;
            xlsFile.SheetName = name;
            xlsFile.PrintPaperSize = global::FlexCel.Core.TPaperSize.Letter;
            xlsFile.PrintXResolution = 300;
            xlsFile.PrintYResolution = 300;
            xlsFile.SetCellValue(1, 1, $"<#{name}.**><#Column Width(autofit;110)>");
            xlsFile.SetCellValue(2, 1, $"<#{name}.*>");
            var range = new global::FlexCel.Core.TXlsNamedRange($"__{name}__", 0, xlsFile.ActiveSheet, -1, -1, -1, -1, 0,
                $"='{name}'!$2:$2");
            xlsFile.SetNamedRange(range);
        }

        #endregion

        #region Dump

        private static readonly SynchronizedDictionary<Type, Func<object, object>[]> GetterMap = new SynchronizedDictionary<Type, Func<object, object>[]>(10);

        public static Func<object, object>[] GetGetters<T>(string[] fields)
        {
            return FlexCelUtils.GetterMap.GetValueOrAdd(typeof(T), k => fields.Select(field => DynamicDelegates.CreateObjectPropertyGet<Func<object, object>>(k, field)).ToArray());
        }
        /// <summary>
        /// Chuyển toàn bộ các trường của View (view xác định kiểu) vào report dưới dạng SetValue
        /// Với từng trường theo tên mẫu [TênView]_[TênTrường]
        /// Nếu view bị null thì các trường của view vẫn được đưa vào dưới giá trị NULL mặc định
        /// phục vụ việc chấp nhận giá trị rỗng ở trong báo cáo
        /// </summary>
        public static DbDataRecord DumpObject(this FlexCel.Report.FlexCelReport report, string objectname, DbDataRecord objectquery)
        {
            var rc = objectquery;
            if (rc != null)
            {
                objectname = objectname + "_";
                for (int i = 0, l = rc.FieldCount; i < l; i++)
                    report.SetValue(objectname + rc.GetName(i), rc.GetValue(i));
            }
            return rc;
        }

        /// <summary>
        /// Chuyển toàn bộ các trường của View (view xác định kiểu) vào report dưới dạng SetValue
        /// Với từng trường theo tên mẫu [TênView]_[TênTrường]
        /// Nếu view bị null thì các trường của view vẫn được đưa vào dưới giá trị NULL mặc định
        /// phục vụ việc chấp nhận giá trị rỗng ở trong báo cáo
        /// </summary>
        public static T DumpObject<T>(this FlexCel.Report.FlexCelReport report, string objectname, IQueryable<T> objectquery, bool alowNull = true)
        {
            var t = objectquery.SingleOrDefault();
            report.DumpObjectOrDefault(objectname, t, alowNull);
            return t;
        }

        /// <summary>
        /// Sử dụng với 1 object
        /// Chuyển toàn bộ các trường của View (view xác định kiểu) vào report dưới dạng SetValue
        /// Với từng trường theo tên mẫu [TênView]_[TênTrường]
        /// Nếu view bị null thì các trường của view vẫn được đưa vào dưới giá trị NULL mặc định
        /// phục vụ việc chấp nhận giá trị rỗng ở trong báo cáo
        /// </summary>
        public static T DumpObjectOrDefault<T>(this FlexCel.Report.FlexCelReport report, string objectname, T objectquery, bool alowNull)
        {
            var t = objectquery;
            var columnNames = typeof(T).GetProperties().Select(k => k.Name).ToArray();
            var getters = GetGetters<T>(columnNames);
            objectname = objectname + "_";
            if (t != null)
            {
                for (var i = 0; i < columnNames.Length; i++)
                    report.SetValue(objectname + columnNames[i], getters[i](t));
            }
            else if(alowNull)
            {
                foreach (string col in columnNames)
                    report.SetValue(objectname + col, null);
            }
            return t;
        }

        #endregion

        #region Cell Address

        public static string ToExcelCellIndex(string cell, int col, int row)
        {
            var indexname = ToExcelColumnName(col) + row.ToString();
            return String.Compare(cell, indexname, StringComparison.InvariantCultureIgnoreCase) == 0 ? cell :
                $"{indexname}({cell})";
        }

        public static string ToExcelColumnName(int col)
        {
            StringBuilder columnName = new StringBuilder();
            for (int modulo, dividend = col, i = 0; dividend > 0; dividend = (int)((dividend - modulo) / 26))
            {
                modulo = (dividend - 1) % 26;
                i = 65 + modulo;
                columnName.Insert(0, (char)i);
            }
            return columnName.ToString();
        }

        private static readonly System.Text.RegularExpressions.Regex REGEX =
            new System.Text.RegularExpressions.Regex(
                @"(?<col>([a-zA-Z]+)|([+-]?(\d)+))(?<row>[+-]?(\d)+)",
                System.Text.RegularExpressions.RegexOptions.Compiled);

        public static bool ParseExcelIndex(this ExcelFile excelFile, string cell, ref int row, ref int col)
        {
            row = col = 0;
            if (cell == null || cell.Length <= 0)
                return false;

            var match = REGEX.Match(cell);
            if (match == null)
                return false;

            var vRow = match.Groups["row"].Value;
            if (!Int32.TryParse(vRow, out row))
                return false;
            if (row <= 0)
                row += excelFile.RowCount;

            var vCol = match.Groups["col"].Value;
            if (vCol == null || vCol.Length <= 0)
                return false;
            if (vCol[0] == '+' || vCol[0] == '-')
            {
                if (vCol.Length <= 1 || vCol[1] < '0' || vCol[1] > '9')
                    return false;
                if (!Int32.TryParse(vCol, out col))
                    return false;
                if (col <= 0)
                    col += excelFile.ColCount;
            }
            else if (vCol[0] >= '0' && vCol[0] <= '9')
            {
                if (!Int32.TryParse(vCol, out col))
                    return false;
            }
            else
            {
                col = vCol.Select(c => (c >= 'A' && c <= 'Z') ? (c - 'A' + 1) : (c - 'a' + 1)).Aggregate((sum, next) => sum * 26 + next);
            }

            return true;
        }

        #endregion

        #region CellValue Getters

        public static string GetCellValueInString(this ExcelFile excelFile, int row, int col, bool trim)
        {
            var obj = excelFile.GetStringFromCell(row, col);
            if (obj == null || obj.Value == null)
                return null;
            return !trim || obj.Value.Length <= 0 ? obj.Value : obj.Value.Trim();
        }

        public static string GetCellValueInString(this ExcelFile excelFile, int row, int col)
        {
            var obj = excelFile.GetStringFromCell(row, col);
            if (obj == null || obj.Value == null)
                return null;
            return obj.Value;
        }

        #endregion

        #region CellValue Converters

        public static object ValueToObject(object obj)
        {
            if (obj is TFormula)
            {
                var r = ((TFormula)obj).Result;
                if (r == null)
                    return null;
                return ValueToObject(r);
            }
            else if (obj is double)
                return (double)obj;
            else if (obj is string)
                return (string)obj;
            else if (obj is TRichString)
                return ((TRichString)obj).Value;
            else if (obj is bool)
                return (bool)obj;
            else if (obj is TFlxFormulaErrorValue)
                return null;
            throw new NotSupportedException(obj.GetType().FullName);
        }

        public static string ValueToString(object obj)
        {
            if (obj is string)
                return (string)obj;
            else if (obj is TRichString)
                return ((TRichString)obj).Value;
            else if (obj is TFormula)
            {
                var r = ((TFormula)obj).Result;
                if (r == null)
                    return null;
                return ValueToString(r);
            }
            else if (obj is TFlxFormulaErrorValue)
                return null;
            return obj.ToString();
        }

        public static double? ValueToDouble(object obj)
        {
            if (obj is double)
                return (double)obj;

            string str;
            if (obj is TFormula)
            {
                var r = ((TFormula)obj).Result;
                if (r == null)
                    return null;
                return ValueToDouble(r);
            }
            else if (obj is string)
                str = (string)obj;
            else if (obj is TRichString)
                str = ((TRichString)obj).Value;
            else
                return null;

            if (str == null || String.IsNullOrWhiteSpace(str))
                return null;

            double value;
            if (Double.TryParse(str.Trim(), out value))
                return value;

            return null;
        }

        public static bool? ValueToBoolean(object obj)
        {
            if (obj is bool)
                return (bool)obj;

            string str;
            if (obj is TFormula)
            {
                var r = ((TFormula)obj).Result;
                if (r == null)
                    return null;
                return ValueToBoolean(r);
            }
            else if (obj is string)
                str = (string)obj;
            else if (obj is TRichString)
                str = ((TRichString)obj).Value;
            else
                return null;

            if (str == null || String.IsNullOrWhiteSpace(str))
                return null;
            str = str.Trim().ToLowerInvariant();

            bool value;
            if (Boolean.TryParse(str, out value))
                return value;

            switch (str)
            {
                case "0":
                case "true":
                case "on":
                case "yes":
                    return false;
                case "1":
                case "false":
                case "off":
                case "no":
                    return true;
            }

            return null;
        }

        public static DateTime? ValueToDateTime(object obj, bool dates1904)
        {
            if (obj is DateTime)
                return (DateTime)obj;
            var val = ValueToDouble(obj);
            if (val == null || !FlxDateTime.IsValidDate(val.Value, dates1904))
                return null;
            return FlxDateTime.FromOADate(val.Value, dates1904);
        }

        public static Int64? ValueToInt64(object obj)
        {
            if (obj is long)
                return (long)obj;
            if (obj is int)
                return (long)(int)obj;
            var val = ValueToDouble(obj);
            if (val == null)
                return null;
            return Convert.ToInt64(val.Value);
        }

        public static Int32? ValueToInt32(object obj)
        {
            if (obj is int)
                return (int)obj;
            if (obj is long)
                if (((long)obj) >= Int32.MinValue && ((long)obj) <= Int32.MaxValue)
                    return (int)(long)obj;
                else
                    return null;
            var val = ValueToDouble(obj);
            if (val == null)
                return null;
            return Convert.ToInt32(val.Value);
        }

        #endregion

        #region CellValue CheckType

        private const double Int32MinValue = (double)Int32.MinValue;
        private const double Int32MaxValue = (double)Int32.MaxValue;
        private const double Int64MinValue = (double)Int64.MinValue;
        private const double Int64MaxValue = (double)Int64.MaxValue;

        public static bool ValueIsObject(object obj)
        {
            if (obj is TFormula)
            {
                var r = ((TFormula)obj).Result;
                if (r == null)
                    return false;
                return ValueIsObject(r);
            }
            else if (obj is double)
                return true;
            else if (obj is string)
                return true;
            else if (obj is TRichString)
                return true;
            else if (obj is bool)
                return true;
            else if (obj is TFlxFormulaErrorValue)
                return false;
            return false;
        }

        public static bool ValueIsString(object obj)
        {
            if (obj is string)
                return true;
            else if (obj is TRichString)
                return true;
            else if (obj is TFormula)
            {
                var r = ((TFormula)obj).Result;
                if (r == null)
                    return false;
                return ValueIsString(r);
            }
            else if (obj is TFlxFormulaErrorValue)
                return false;
            return true;
        }

        public static bool ValueIsDouble(object obj)
        {
            if (obj is double)
                return true;

            string str;
            if (obj is TFormula)
            {
                var r = ((TFormula)obj).Result;
                if (r == null)
                    return false;
                return ValueIsDouble(r);
            }
            else if (obj is string)
                str = (string)obj;
            else if (obj is TRichString)
                str = ((TRichString)obj).Value;
            else
                return false;

            if (str == null || String.IsNullOrWhiteSpace(str))
                return false;

            double value;
            return Double.TryParse(str.Trim(), out value);
        }

        public static bool ValueIsBoolean(object obj)
        {
            if (obj is bool)
                return true;

            string str;
            if (obj is TFormula)
            {
                var r = ((TFormula)obj).Result;
                if (r == null)
                    return false;
                return ValueIsBoolean(r);
            }
            else if (obj is string)
                str = (string)obj;
            else if (obj is TRichString)
                str = ((TRichString)obj).Value;
            else
                return false;

            if (str == null || String.IsNullOrWhiteSpace(str))
                return false;
            str = str.Trim().ToLowerInvariant();

            bool value;
            if (Boolean.TryParse(str, out value))
                return true;

            switch (str)
            {
                case "0":
                case "true":
                case "on":
                case "yes":
                    return true;
                case "1":
                case "false":
                case "off":
                case "no":
                    return true;
            }

            return false;
        }

        public static bool ValueIsDateTime(object obj, bool dates1904)
        {
            if (obj is DateTime)
                return true;
            var val = ValueToDouble(obj);
            return val != null && FlxDateTime.IsValidDate(val.Value, dates1904);
        }

        public static bool ValueIsInt64(object obj)
        {
            if (obj is long || obj is int)
                return true;
            var val = ValueToDouble(obj);
            if (val == null)
                return false;
            var round = Math.Round(val.Value);
            return round == val.Value && round <= Int64MaxValue && round >= Int64MinValue;
        }

        public static bool ValueIsInt32(object obj)
        {
            if (obj is int)
                return true;
            if (obj is long)
                return ((long)obj) >= Int32.MinValue && ((long)obj) <= Int32.MaxValue;
            var val = ValueToDouble(obj);
            if (val == null)
                return false;
            var round = Math.Round(val.Value);
            return round == val.Value && round <= Int32MaxValue && round >= Int32MinValue;
        }

        #endregion

        #region FlexCelImgExport

        public static System.Drawing.Bitmap ExportImageNext(this global::FlexCel.Render.FlexCelImgExport flexCelImgExport, ref global::FlexCel.Render.TImgExportInfo exportInfo, float xCrop = 0F, float yCrop = 0F, float? zoom = null, System.Drawing.Color? backgroundColor = null, System.Drawing.Drawing2D.InterpolationMode? interpolationMode = null, System.Drawing.Drawing2D.SmoothingMode? smoothingMode = null)
        {
            var pagebounds = exportInfo.ActiveSheet.PageBounds;
            System.Drawing.Bitmap bitmap;
            if (zoom == null)
            {
                bitmap = global::FlexCel.Draw.GdipBitmapConstructor.CreateBitmap((int)(pagebounds.Width - xCrop), (int)(pagebounds.Height - yCrop));
                bitmap.SetResolution(96, 96);
            }
            else
            {
                bitmap = global::FlexCel.Draw.GdipBitmapConstructor.CreateBitmap((int)((pagebounds.Width - xCrop) * zoom.Value), (int)((pagebounds.Height - yCrop) * zoom.Value));
                bitmap.SetResolution(96 * zoom.Value, 96 * zoom.Value);
            }

            try
            {
                using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
                {
                    graphics.Clear(backgroundColor ?? System.Drawing.Color.White);
                    if (interpolationMode != null)
                        graphics.InterpolationMode = interpolationMode.Value;
                    if (smoothingMode != null)
                        graphics.SmoothingMode = smoothingMode.Value;
                    flexCelImgExport.ExportNext(graphics, ref exportInfo);
                }
                return bitmap;
            }
            catch
            {
                bitmap.Dispose();
                throw;
            }
        }

        #endregion

        public static string UDFGetParameterInString(string name, object[] parameters, int index)
        {
            if (parameters[index] == null)
                return null;
            if (ValueIsString(parameters[index]))
                return ValueToString(parameters[index]);

            return null;
        }

        public static double? UDFGetParameterInDouble(string name, object[] parameters, int index)
        {
            if (parameters[index] == null)
                return null;

            if (ValueIsDouble(parameters[index]))
                return ValueToDouble(parameters[index]);

            return null;
        }
    }
}
