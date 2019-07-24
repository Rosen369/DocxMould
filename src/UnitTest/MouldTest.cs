using DocxMould;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace UnitTest
{
    [TestClass]
    public class MouldTest
    {
        [TestMethod]
        public void TestReplace()
        {
            using (var fs = File.Open("./template.docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
            using (var mould = new Mould(fs))
            {
                var replacement = new Replacement();
                replacement.Add("Field1", "Field1BlaBlaBla111");
                replacement.Add("Field2", "Field2BlaBlaBla222");
                replacement.Add("Field3", "Field3BlaBlaBla333");
                mould.ReplaceField(replacement);
                mould.Save();
            }
        }

        [TestMethod]
        public void TestRemovePartA()
        {
            using (var fs = File.Open("./template.docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
            using (var mould = new Mould(fs))
            {
                var removal = new Removal();
                removal.Add("PartAStart", "PartAEnd");
                mould.RemoveSection(removal);
                mould.Save();
            }
        }

        [TestMethod]
        public void TestRemovePartB()
        {
            using (var fs = File.Open("./template.docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
            using (var mould = new Mould(fs))
            {
                var removal = new Removal();
                removal.Add("PartBStart", "PartBEnd");
                mould.RemoveSection(removal);
                mould.Save();
            }
        }

        [TestMethod]
        public void TestReplaceAndRemovePartA()
        {
            using (var fs = File.Open("./template.docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
            using (var mould = new Mould(fs))
            {
                var removal = new Removal();
                removal.Add("PartAStart", "PartAEnd");
                mould.RemoveSection(removal);

                var replacement = new Replacement();
                replacement.Add("Field1", "Field1BlaBlaBla111");
                replacement.Add("Field2", "Field2BlaBlaBla222");
                replacement.Add("Field3", "Field3BlaBlaBla333");
                mould.ReplaceField(replacement);
                mould.Save();
            }
        }

        [TestMethod]
        public void TestReplaceAndRemovePartB()
        {
            using (var fs = File.Open("./template.docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
            using (var mould = new Mould(fs))
            {
                var removal = new Removal();
                removal.Add("PartBStart", "PartBEnd");
                mould.RemoveSection(removal);

                var replacement = new Replacement();
                replacement.Add("Field1", "Field1BlaBlaBla111");
                replacement.Add("Field2", "Field2BlaBlaBla222");
                replacement.Add("Field3", "Field3BlaBlaBla333");
                mould.ReplaceField(replacement);
                mould.Save();
            }
        }
    }
}
