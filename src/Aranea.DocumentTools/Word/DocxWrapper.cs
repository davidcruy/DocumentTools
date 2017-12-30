using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentTools.Word
{
    public class DocxWrapper
    {
        private readonly WordprocessingDocument _innerDocument;

        public DocxWrapper(byte[] content)
        {
            var stream = new MemoryStream();
            stream.Write(content, 0, content.Length);
            stream.Position = 0;

            _innerDocument = WordprocessingDocument.Open(stream, true);
        }

        public DocxWrapper(Stream documentData)
        {
            _innerDocument = WordprocessingDocument.Open(documentData, true);
        }

        public void ReplaceBookmark(string bookmarkName, string text)
        {
            if (!TryFindBookmark(bookmarkName, out BookmarkStart start, out BookmarkEnd end)) return;

            var previous = start.PreviousSibling();
            var parent = start.Parent;

            RemoveBookmark(start, end);

            var toInsert = new Run(new Text(text));
            if (previous == null) parent.PrependChild(toInsert);
            else parent.InsertAfter(toInsert, previous);
        }

        public void RemoveBookmark(string name)
        {
            if (!TryFindBookmark(name, out BookmarkStart start, out BookmarkEnd end)) return;

            RemoveBookmark(start, end);
        }

        private static void RemoveBookmark(BookmarkStart start, BookmarkEnd end)
        {
            while (start.NextSibling() != null || start.Parent.NextSibling() != null)
            {
                if (start.NextSibling() != null)
                {
                    var next = start.NextSibling();
                    if (next is BookmarkEnd bookmarkEnd && bookmarkEnd.Id.Value == end.Id.Value)
                    {
                        next.Remove();
                        break; // FOUND IT!
                    }

                    next.Remove();
                }
                else
                {
                    var nextParent = start.Parent.NextSibling();
                    if (nextParent is BookmarkEnd bookmarkEnd && bookmarkEnd.Id.Value == end.Id.Value)
                    {
                        nextParent.Remove();
                        break; // FOUND IT!
                    }

                    foreach (var child in nextParent.ChildElements)
                    {
                        var cloned = child.CloneNode(true);
                        start.Parent.AppendChild(cloned);
                    }

                    nextParent.Remove();
                }
            }

            start.Remove();
        }

        private bool TryFindBookmark(string name, out BookmarkStart start, out BookmarkEnd end)
        {
            start = null;
            end = null;
            string startId = null;

            foreach (var bookmarkStart in _innerDocument.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
            {
                if (bookmarkStart.Name == name)
                {
                    start = bookmarkStart;
                    startId = bookmarkStart.Id;

                    break;
                }
            }

            if (string.IsNullOrEmpty(startId))
                return false;

            foreach (var bookmarkEnd in _innerDocument.MainDocumentPart.RootElement.Descendants<BookmarkEnd>())
            {
                if (bookmarkEnd.Id == startId)
                {
                    end = bookmarkEnd;
                    return true;
                }
            }

            return false;
        }

        public byte[] GetContent()
        {
            var outStream = new MemoryStream();

            _innerDocument.Clone(outStream);
            var docBytes = outStream.ToArray();

            return docBytes;
        }

        public bool HasMergeField(string name)
        {
            return GetMergeFields(_innerDocument.MainDocumentPart.RootElement).ContainsKey(name);
        }

        public void MergeDataSet(DataSet dataSet)
        {
            var map = GetMergeFields(_innerDocument.MainDocumentPart.RootElement);

            foreach (DataTable table in dataSet.Tables)
            {
                var startKey = $"TableStart:{table.TableName}";
                var endKey = $"TableEnd:{table.TableName}";

                IMergeField start = null;
                if (map.ContainsKey(startKey))
                {
                    start = map[startKey];
                }

                IMergeField target = null;
                if (map.ContainsKey(endKey))
                {
                    target = map[endKey];
                }

                if (start == null || target == null)
                    return; // Stop building table if start or end doesn't exist

                // Find matching parent...
                var startRow = start.Parent<TableRow>();
                var targetRow = target.Parent<TableRow>();

                if (startRow == null || targetRow == null)
                {
                    return;
                }
                if (startRow == targetRow)
                {
                    var cloneRow = startRow.CloneNode(true);

                    foreach (DataRow row in table.Rows)
                    {
                        var values = row.Table.Columns
                            .Cast<DataColumn>()
                            .ToDictionary(c => c.ColumnName, c => row[c]);

                        values.Add(startKey, "");
                        values.Add(endKey, "");

                        var isFirst = table.Rows.IndexOf(row) == 0;
                        if (isFirst)
                        {
                            MergeData(values, startRow);
                        }
                        else
                        {
                            var result = startRow.Parent.InsertAfter(cloneRow.CloneNode(true), startRow);
                            MergeData(values, result);
                        }
                    }
                }
                else
                {
                    throw new Exception("Table-merging currently is only supported when table-start and end is inside the same table row...");
                }
            }
        }

        public void MergeData<T>(T value)
        {
            var values = new Dictionary<string, object>();

            foreach (var prop in value.GetType().GetProperties())
            {
                values.Add(prop.Name, prop.GetValue(value, null));
            }

            MergeData(values, _innerDocument.MainDocumentPart.RootElement);
        }

        private void MergeData(IDictionary<string, object> values, OpenXmlElement startElement)
        {
            var mergeFieldMap = GetMergeFields(startElement);

            foreach (var value in values)
            {
                if (!mergeFieldMap.ContainsKey(value.Key)) continue;

                mergeFieldMap[value.Key].Merge(value.Value);
            }
        }

        private static IDictionary<string, IMergeField> GetMergeFields(OpenXmlElement startElement)
        {
            var map = new Dictionary<string, IMergeField>();

            foreach (var fieldStart in startElement.Descendants<FieldCode>())
            {
                if (ComplexMergeField.TryParse(fieldStart, out var field)) map.Add(field.Key, field);
            }

            foreach (var fieldStart in startElement.Descendants<SimpleField>())
            {
                if (SimpleMergeField.TryParse(fieldStart, out var field)) map.Add(field.Key, field);
            }

            return map;
        }

        public int GetNumberOfPages()
        {
            var pageCount = int.Parse(_innerDocument.ExtendedFilePropertiesPart.Properties.Pages.Text);

            return pageCount;
        }

        public void Close()
        {
            _innerDocument.Close();
        }
    }
}
