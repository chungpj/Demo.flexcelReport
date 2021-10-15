using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace Report.AsposeHelper
{
    public static class WordsUtils
    {
        public static DocumentBuilder ReplaceMergeField(this Document document, string mergeField)
        {
            var builder = new DocumentBuilder(document);
            builder.MoveToMergeField(mergeField);
            return builder;
        }

        /// <summary>
        /// Inserts content of the external document after the specified node.
        /// Section breaks and section formatting of the inserted document are ignored.
        /// </summary>
        /// <param name="insertAfterNode">Node in the destination document after which the content 
        /// should be inserted. This node should be a block level node (paragraph or table).</param>
        /// <param name="srcDoc">The document to insert.</param>
        public static void InsertDocument(this Node insertAfterNode, Document srcDoc)
        {
            // Make sure that the node is either a paragraph or table.
            if ((!insertAfterNode.NodeType.Equals(NodeType.Paragraph)) &
              (!insertAfterNode.NodeType.Equals(NodeType.Table)))
                throw new ArgumentException("The destination node should be either a paragraph or table.");

            // We will be inserting into the parent of the destination paragraph.
            CompositeNode dstStory = insertAfterNode.ParentNode;

            // This object will be translating styles and lists during the import.
            NodeImporter importer = new NodeImporter(srcDoc, insertAfterNode.Document, ImportFormatMode.KeepSourceFormatting);

            // Loop through all sections in the source document.
            foreach (Section srcSection in srcDoc.Sections)
            {
                // Loop through all block level nodes (paragraphs and tables) in the body of the section.
                foreach (Node srcNode in srcSection.Body)
                {
                    // Let's skip the node if it is a last empty paragraph in a section.
                    if (srcNode.NodeType.Equals(NodeType.Paragraph))
                    {
                        Paragraph para = (Paragraph)srcNode;
                        if (para.IsEndOfSection && !para.HasChildNodes)
                            continue;
                    }

                    // This creates a clone of the node, suitable for insertion into the destination document.
                    Node newNode = importer.ImportNode(srcNode, true);

                    // Insert new node after the reference node.
                    dstStory.InsertAfter(newNode, insertAfterNode);
                    insertAfterNode = newNode;
                }
            }
        }

        public static System.Drawing.Bitmap ExportFirstPage(this Document reportDocument, float xCrop = 0F, float yCrop = 0F, float? zoom = null, System.Drawing.Color? backgroundColor = null, System.Drawing.Drawing2D.InterpolationMode? interpolationMode = null, System.Drawing.Drawing2D.SmoothingMode? smoothingMode = null)
        {
            System.Drawing.Bitmap bitmap;
            System.Drawing.SizeF pageSize; // = new System.Drawing.SizeF(pageInfo.WidthInPoints * 100.0f / 72.0f, pageInfo.HeightInPoints * 100.0f / 72.0f);
            var pageInfo = reportDocument.GetPageInfo(0);

            if (zoom == null)
            {
                pageSize = new System.Drawing.SizeF(pageInfo.WidthInPoints * 100.0f / 72.0f, pageInfo.HeightInPoints * 100.0f / 72.0f);
                bitmap = global::FlexCel.Draw.GdipBitmapConstructor.CreateBitmap((int)(pageSize.Width - xCrop), (int)(pageSize.Height - xCrop));
                //bitmap.SetResolution(96, 96);
            }
            else
            {
                pageSize = new System.Drawing.SizeF(pageInfo.WidthInPoints * 100.0f / 72.0f * zoom.Value, pageInfo.HeightInPoints * 100.0f / 72.0f * zoom.Value);
                bitmap = global::FlexCel.Draw.GdipBitmapConstructor.CreateBitmap((int)(pageSize.Width - xCrop), (int)(pageSize.Height - xCrop));
                //bitmap.SetResolution(96 * zoom.Value, 96 * zoom.Value);
            }

            try
            {
                using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
                {
                    graphics.Clear(backgroundColor ?? System.Drawing.Color.White);
                    // PageUnit is measuring in 1% inches for printing
                    graphics.PageUnit = System.Drawing.GraphicsUnit.Display;
                    if (interpolationMode != null)
                        graphics.InterpolationMode = interpolationMode.Value;
                    if (smoothingMode != null)
                        graphics.SmoothingMode = smoothingMode.Value;
                    //reportDocument.RenderToSize(0, graphics, 0, 0, (pageSize.Width - xCrop), (pageSize.Height - xCrop));
                    //reportDocument.RenderToSize(0, graphics, 0, 0, (pageSize.Width - xCrop), (pageSize.Height - xCrop));
                    reportDocument.RenderToScale(0, graphics, 0, 0, zoom ?? 1.0f);
                    //, zoom ?? 1.0f
                }
                return bitmap;
            }
            catch
            {
                bitmap.Dispose();
                throw;
            }
        }
    }
}
