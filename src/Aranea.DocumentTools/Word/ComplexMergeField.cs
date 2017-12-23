using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentTools.Word
{
    internal class ComplexMergeField : MergeField<FieldCode>
    {
        public ComplexMergeField(string key, FieldCode field)
        {
            Key = key;
            Field = field;
        }

        public override void Merge(object value)
        {
            var rFldCode = (Run) Field.Parent;
            var rBegin = rFldCode.PreviousSibling<Run>();
            var rSep = rFldCode.NextSibling<Run>();
            var rText = rSep.NextSibling<Run>();
            var rEnd = rText.NextSibling<Run>();

            rFldCode.Remove();
            rBegin.Remove();
            rSep.Remove();
            rEnd.Remove();

            if (!string.IsNullOrEmpty(value?.ToString()))
            {
                var t = rText.GetFirstChild<Text>();
                if (t != null)
                {
                    t.Text = value.ToString();
                }
            }
            else
            {
                rText.Remove();
            }
        }

        public static bool TryParse(FieldCode fieldStart, out ComplexMergeField field)
        {
            var split = fieldStart.InnerText
                .Trim()
                .Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

            if (split.Length >= 2 && split[0] == "MERGEFIELD")
            {
                var fieldKey = split[1];
                field = new ComplexMergeField(fieldKey, fieldStart);
                return true;
            }

            field = null;
            return false;
        }
    }
}