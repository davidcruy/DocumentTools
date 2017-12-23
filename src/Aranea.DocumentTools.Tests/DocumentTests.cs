using System.IO;
using DocumentTools.Word;
using NUnit.Framework;

namespace DocumentTools.Tests
{
    [TestFixture]
    public class DocumentTests
    {
        [Test]
        public void InitWrapperTests()
        {
            var path = @"C:\Git\DocumentTools\src\Aranea.DocumentTools.Tests";
            var inputContent = File.ReadAllBytes(Path.Combine(path, "sample.docx"));
            var wrapper = new DocxWrapper(inputContent);

            Assert.AreEqual(wrapper.GetNumberOfPages(), 2);

            wrapper.ReplaceBookmark("ReplaceMe", "With me!");
            wrapper.ReplaceBookmark("ReplaceMe2", "With me too!");

            Assert.IsTrue(wrapper.HasMergeField("MergeMe"));
            Assert.IsFalse(wrapper.HasMergeField("IDontExist"));

            wrapper.MergeData(
                new
                {
                    MergeMe = "WithME!"
                });

            var content = wrapper.GetContent();

            Assert.IsTrue(content.Length > 0);
            File.WriteAllBytes(Path.Combine(path, "output.docx"), content);
        }
    }
}
