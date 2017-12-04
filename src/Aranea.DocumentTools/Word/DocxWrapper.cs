using System;
using System.Collections.Generic;
using System.IO;
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

        public byte[] GetContentPDF()
        {
            throw new NotImplementedException();
        }

        public bool HasMergeField(string name)
        {
            EnsureMergeFieldMap();

            return _mergeFieldMap.ContainsKey(name);
        }

        public void MergeData(IDictionary<string, object> values)
        {
            EnsureMergeFieldMap();

            foreach (var value in values)
            {
                if (_mergeFieldMap.ContainsKey(value.Key))
                {
                    var field = _mergeFieldMap[value.Key];
                    var fieldStart = field.PreviousSibling();
                    var parent = field.Parent;
                    field.Remove();

                    parent.InsertAfter(new Run(new Text(value.Value.ToString())), fieldStart);
                }
                else if (_mergeFieldComplexMap.ContainsKey(value.Key))
                {
                    var field = _mergeFieldComplexMap[value.Key];

                    var rFldCode = (Run)field.Parent;
                    var rBegin = rFldCode.PreviousSibling<Run>();
                    var rSep = rFldCode.NextSibling<Run>();
                    var rText = rSep.NextSibling<Run>();
                    var rEnd = rText.NextSibling<Run>();

                    rFldCode.Remove();
                    rBegin.Remove();
                    rSep.Remove();
                    rEnd.Remove();

                    var t = rText.GetFirstChild<Text>();
                    if (t != null)
                    {
                        t.Text = value.Value.ToString();
                    }
                }
            }
        }

        public void MergeData(string key, object value)
        {
            MergeData(new Dictionary<string, object> { { key, value } });
        }

        private IDictionary<string, SimpleField> _mergeFieldMap;
        private IDictionary<string, FieldCode> _mergeFieldComplexMap;
        private void EnsureMergeFieldMap()
        {
            if (_mergeFieldMap != null) return;

            _mergeFieldMap = new Dictionary<string, SimpleField>();

            foreach (var fieldStart in _innerDocument.MainDocumentPart.RootElement.Descendants<SimpleField>())
            {
                var split = fieldStart.Instruction.Value
                    .Trim()
                    .Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                if (split.Length >= 2 && split[0] == "MERGEFIELD")
                {
                    _mergeFieldMap.Add(split[1], fieldStart);
                }
            }

            _mergeFieldComplexMap = new Dictionary<string, FieldCode>();

            foreach (var fieldStart in _innerDocument.MainDocumentPart.RootElement.Descendants<FieldCode>())
            {
                var split = fieldStart.InnerText
                    .Trim()
                    .Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                if (split.Length >= 2 && split[0] == "MERGEFIELD")
                {
                    _mergeFieldComplexMap.Add(split[1], fieldStart);
                }
            }
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
