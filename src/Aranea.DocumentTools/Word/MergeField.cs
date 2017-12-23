using DocumentFormat.OpenXml;

namespace DocumentTools.Word
{
    internal interface IMergeField
    {
        string Key { get; set; }
        T Parent<T>() where T : OpenXmlElement;
        void Merge(object value);
    }

    internal abstract class MergeField<T> : IMergeField
        where T: OpenXmlElement
    {
        public string Key { get; set; }

        internal T Field { get; set; }

        public TParent Parent<TParent>() where TParent : OpenXmlElement
        {
            var currentElement = (OpenXmlElement) Field;
            while (currentElement.Parent != null)
            {
                var parent = currentElement.Parent;
                if (parent.GetType() == typeof(TParent))
                    return (TParent)parent;

                currentElement = currentElement.Parent;
            }

            return null;
        }

        public abstract void Merge(object value);
    }
}