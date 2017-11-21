using System.Collections.Generic;

namespace DocumentTools.Word
{
    public interface IDocumentWrapper
    {
        void MergeData(IDictionary<string, object> values);
        void MergeData(string key, object value);
        void ReplaceBookmark(string bookmarkName, string text);
        void RemoveBookmark(string name);
        byte[] GetContent();
        byte[] GetContentPDF();
        bool HasMergeField(string name);
        int GetNumberOfPages();
    }
}