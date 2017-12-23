using System;
using System.Collections.Generic;
using System.Data;
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
            EnsureBookmarkMap();

            if (!_bookmarkMap.ContainsKey(bookmarkName)) return;

            var bookmarkStart = _bookmarkMap[bookmarkName];
            var element = bookmarkStart.NextSibling();

            while (element != null && !(element is BookmarkEnd))
            {
                var nextElem = element.NextSibling();

                element.Remove();
                element = nextElem;
            }

            bookmarkStart.Parent.InsertAfter<Run>(new Run(new Text(text)), bookmarkStart);
        }

        public void RemoveBookmark(string name)
        {
            EnsureBookmarkMap();

            if (!_bookmarkMap.ContainsKey(name)) return;

            throw new NotImplementedException();
        }

        private IDictionary<string, BookmarkStart> _bookmarkMap;
        private void EnsureBookmarkMap()
        {
            if (_bookmarkMap != null) return;

            _bookmarkMap = new Dictionary<string, BookmarkStart>();

            foreach (var bookmarkStart in _innerDocument.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
            {
                _bookmarkMap.Add(bookmarkStart.Name, bookmarkStart);
            }
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

                Console.WriteLine("start & end have been found");

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
            return int.Parse(_innerDocument.ExtendedFilePropertiesPart.Properties.Pages.Text);
        }

        public void Close()
        {
            _innerDocument.Close();
        }
    }
}
