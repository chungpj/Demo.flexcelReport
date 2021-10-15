using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;
using System.Drawing;
using Aspose.Cells.Drawing;
using Report.Common;

namespace Report.AsposeHelper
{
    public static class CellsUtils
    {
        public static Point LocationIn(this Cell cell, PointF offsetPercent)
        {
            var point = new Point();
            var mergeRange = cell.GetMergedRange();
            var cells = cell.Worksheet.Cells;
            if (mergeRange == null || (mergeRange.ColumnCount <= 1 && mergeRange.RowCount <= 1))
            {
                point.X = cells.GetColumnWidthPixel(cell.Column);
                point.Y = cells.GetRowHeightPixel(cell.Row);
            }
            else
            {
                for (var i = 0; i < mergeRange.ColumnCount; i++)
                    point.X += cells.GetColumnWidthPixel(cell.Column + i);
                for (var i = 0; i < mergeRange.RowCount; i++)
                    point.Y += cells.GetRowHeightPixel(cell.Row + i);
            }

            point.X = (int)(point.X * offsetPercent.X);
            point.Y = (int)(point.Y * offsetPercent.Y);

            return point;
        }

        public static Size DistanceTo(this Cell from, Cell to)
        {
            var size = new Size();
            var cells = from.Worksheet.Cells;
            for (int col = from.Column, inc = to.Column > from.Column ? 1 : -1; col != to.Column; col += inc)
                size.Width += inc * cells.GetColumnWidthPixel(col);
            for (int row = from.Row, inc = to.Row > from.Row ? 1 : -1; row != to.Row; row += inc)
                size.Height += inc * cells.GetRowHeightPixel(row);
            return size;
        }

        public static Aspose.Cells.Drawing.LineShape DrawLine(LineHead from, LineHead to, LineInfo lineInfo)
        {
            var fromLocation = from.Cell.LocationIn(lineInfo.OffsetGetter(from, lineInfo));
            var toLocation = to.Cell.LocationIn(lineInfo.OffsetGetter(from, lineInfo));
            var distanceSize = from.Cell.DistanceTo(to.Cell);

            var line = from.Cell.Worksheet.Shapes.AddLine(
                distanceSize.Height >= 0 ? from.Cell.Row : to.Cell.Row,
                distanceSize.Height >= 0 ? fromLocation.Y : toLocation.Y,
                distanceSize.Width >= 0 ? from.Cell.Column : to.Cell.Column,
                distanceSize.Width >= 0 ? fromLocation.X : toLocation.X,
                Math.Abs(distanceSize.Height) - fromLocation.Y + toLocation.Y,
                Math.Abs(distanceSize.Width) - fromLocation.X + toLocation.X
            );
            if (distanceSize.Height < 0)
                line.IsFlippedVertically = true;
            if (distanceSize.Width < 0)
                line.IsFlippedHorizontally = true;
//            line.BeginArrowheadStyle = Aspose.Cells.Drawing.MsoArrowheadStyle.ArrowOval;
//#if DEBUG
//            line.EndArrowheadStyle = Aspose.Cells.Drawing.MsoArrowheadStyle.ArrowStealth;
//#else
//            line.EndArrowheadStyle = Aspose.Cells.Drawing.MsoArrowheadStyle.ArrowOval;
//#endif
//            line.LineFormat.DashStyle = lineInfo.DashStyleGetter(from, to, lineInfo);
//            line.LineFormat.Weight = lineInfo.WeightGetter(from, to, lineInfo);
//            line.LineFormat.ForeColor = lineInfo.ColorGetter(from, to, lineInfo);

            return lineInfo.LineShapeGetter(from, to, lineInfo, line);
        }

        public static IEnumerable<Cell> Search(this Cells cells, string search, bool regex = false, FindOptions options = null)
        {
            if (options == null)
                options = new FindOptions()
                {
                    CaseSensitive = true,
                    LookAtType = LookAtType.EntireContent,
                    LookInType = LookInType.Values,
                    SearchNext = true,
                    RegexKey = regex,
                };
            for (var cell = cells.Find(search, null, options); cell != null;
                cell = cells.Find(search, cell, options))
                yield return cell;
        }

        public sealed class LineHead
        {
            public static IEnumerable<LineHead> Parse(Cell cell)
            {
                var stringValue = cell.StringValue;
                for (int last = 0, pos = 0, length = stringValue.Length; pos < length; last = pos + 1)
                {
                    pos = stringValue.IndexOf(',', last);
                    if (pos < 0)
                        pos = length;
                    var lineHead = LineHead.Parse(cell, stringValue.Substring(last, pos - last));
                    if (lineHead != null)
                        yield return lineHead;
                }
            }

            private static LineHead Parse(Cell cell, string value)
            {
                int page, order, dot = value.IndexOf('.'), sharp = value.IndexOf('#');
                char? param;

                if (dot < 0)
                    return null;

                if (!Int32.TryParse(value.Substring(1, dot - 1), out page) || page < 0)
                    return null;
                if (sharp < 0)
                {
                    if (!Int32.TryParse(value.Substring(dot + 1), out order) || order < 0)
                        return null;
                    param = null;
                }
                else
                {
                    if (dot > sharp || !Int32.TryParse(value.Substring(dot + 1, sharp - 1 - dot), out order) || order < 0)
                        return null;
                    param = value[sharp + 1];
                }
                return new LineHead()
                {
                    Cell = cell,
                    Type = value[0],
                    Page = page,
                    Order = order,
                    Param = param,
                };
            }

            public Cell Cell;

            public char Type;
            public int Page;
            public int Order;
            public char? Param;
        }

        public sealed class LineInfo
        {
            public static readonly PointF OffsetDefault = new PointF(.5f, .5f);

            public LineInfo(char type)
            {
                this.Type = type;
            }

            public readonly char Type;


            /// <summary>
            /// Giá trị offset (mặc định) của điểm vẽ so với góc trên bên trái của cell từ 0.0 - 1.0 (mặc định là 0.5 - ở giữa)
            /// </summary>
            public PointF? Offset;
            /// <summary>
            /// Hàm điều chỉnh giá trị Offset
            /// </summary>
            public Func<LineHead, LineInfo, PointF> OffsetGetter = _OffsetGetter;
            private static PointF _OffsetGetter(LineHead lineHead, LineInfo lineInfo) { return lineInfo.Offset ?? LineInfo.OffsetDefault; }


            /// <summary>
            /// Giá trị màu (mặc định) của đường vẽ 0 (mặc định là màu đen)
            /// </summary>
            public Color? Color;
            /// <summary>
            /// Hàm điều chỉnh giá trị Color
            /// </summary>
            public Func<LineHead, LineHead, LineInfo, Color> ColorGetter = _ColorGetter;
            private static Color _ColorGetter(LineHead from, LineHead to, LineInfo lineInfo) { return lineInfo.Color ?? System.Drawing.Color.Black; }


            /// <summary>
            /// Giá trị dạng chấm ghạch (mặc định) của đường vẽ 0 (mặc định là liền mạch)
            /// </summary>
            public MsoLineDashStyle? DashStyle;
            /// <summary>
            /// Hàm điều chỉnh giá trị DashStyle
            /// </summary>
            public Func<LineHead, LineHead, LineInfo, MsoLineDashStyle> DashStyleGetter = _DashStyleGetter;
            private static MsoLineDashStyle _DashStyleGetter(LineHead from, LineHead to, LineInfo lineInfo) { return lineInfo.DashStyle ?? MsoLineDashStyle.Solid; }


            /// <summary>
            /// Giá trị độ dày (mặc định) của đường vẽ 0 (mặc định là 1.0 - đường chuẩn)
            /// </summary>
            public double? Weight;
            /// <summary>
            /// Hàm điều chỉnh giá trị Weight
            /// </summary>
            public Func<LineHead, LineHead, LineInfo, double> WeightGetter = _WeightGetter;
            private static double _WeightGetter(LineHead from, LineHead to, LineInfo lineInfo) { return lineInfo.Weight ?? 1.0; }


            /// <summary>
            /// Hàm điều chỉnh đường vẽ ra
            /// </summary>
            public Func<LineHead, LineHead, LineInfo, Aspose.Cells.Drawing.LineShape, Aspose.Cells.Drawing.LineShape> LineShapeGetter = _LineShapeGetter;
            private static Aspose.Cells.Drawing.LineShape _LineShapeGetter(LineHead from, LineHead to, LineInfo lineInfo, Aspose.Cells.Drawing.LineShape lineShape) { return lineShape; }
        }

        public static void DrawConnectedLines(this Worksheet sheet, params LineInfo[] lineInfos)
        {
            var lineInfoMap = lineInfos.ToDictionaryWithCheck(li => li.Type);
            // Tìm các cell có toàn bộ nội dung dưới định dạng x12.213,y243.23, x12.213,y243.23#p...
            // OLD1: (\d)*\.(\d)*
            // OLD2: ([a-zA-Z]\d+\.\d+,?)+
            foreach (var typedItems in sheet.Cells.Search(@"([a-zA-Z]\d+\.\d+(#+\w)?,?)+", true).SelectMany(c => LineHead.Parse(c)).GroupBy(l => l.Type))
            {
                var lineInfo = lineInfoMap[typedItems.Key];
                foreach (var pagedTypedItems in typedItems.GroupBy(ti => ti.Page))
                {
                    var orderedPagedTypedItems = pagedTypedItems.OrderBy(pti => pti.Order).ToArray();
                    if (orderedPagedTypedItems.Length == 1)
                    {
                        CellsUtils.DrawLine(orderedPagedTypedItems[0], orderedPagedTypedItems[0], lineInfo);
                        orderedPagedTypedItems[0].Cell.Value = null;
                    }
                    else if (orderedPagedTypedItems.Length > 1)
                    {
                        for (int i = 0, l = orderedPagedTypedItems.Length - 1; i < l; i++)
                        {
                            CellsUtils.DrawLine(orderedPagedTypedItems[i], orderedPagedTypedItems[i + 1], lineInfo);
                            orderedPagedTypedItems[i].Cell.Value = null;
                        }
                        orderedPagedTypedItems[orderedPagedTypedItems.Length - 1].Cell.Value = null;
                    }
                }
            }
        }
    }
}
