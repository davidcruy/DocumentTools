using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentTools.Word
{
    internal class SimpleMergeField : MergeField<SimpleField>
    {
        public SimpleMergeField(string key, SimpleField field)
        {
            Key = key;
            Field = field;
        }

        public override void Merge(object value)
        {
            var fieldStart = Field.PreviousSibling();
            var parent = Field.Parent;
            Field.Remove();

            if (!string.IsNullOrEmpty(value?.ToString()))
                parent.InsertAfter(new Run(new Text(value.ToString())), fieldStart);
        }

        public static bool TryParse(SimpleField fieldStart, out SimpleMergeField field)
        {
            var split = fieldStart.Instruction.Value
                .Trim()
                .Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

            if (split.Length >= 2 && split[0] == "MERGEFIELD")
            {
                var fieldName = split[1];
                field = new SimpleMergeField(fieldName, fieldStart);
                return true;
            }

            field = null;
            return false;
        }
    }
}