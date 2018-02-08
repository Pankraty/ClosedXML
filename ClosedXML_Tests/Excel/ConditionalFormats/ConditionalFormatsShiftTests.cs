using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML_Tests.Excel.ConditionalFormats
{
    [TestFixture]
    public class ConditionalFormatsShiftTests
    {
        [Test]
        [TestCase("=IF(C3>A1,TRUE,FALSE)", 2, "=IF(C4>A1,TRUE,FALSE)")]
        [TestCase("=IF(C3>A1,TRUE,FALSE)", 1, "=IF(C4>A2,TRUE,FALSE)")]
        public void ConditionalFormatFormulaShiftedOnRowInsert(string originalFormula, int rowNum, string expectedFormula)
        {
            //Arrange
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("ConditionalFormat");
            var range = ws.Range("C3:E5");
            var condFormat = range.AddConditionalFormat();
            condFormat.Values.Add(new XLFormula(originalFormula));
            condFormat.Style.Fill.BackgroundColor = XLColor.AirForceBlue;

            wb.SaveAs(@"e:\cf1.xlsx");
            //Act
            //ws.Row(rowNum).InsertRowsAbove(1);
            ws.Row(rowNum).Delete();
            condFormat = ws.ConditionalFormats.First();
            wb.SaveAs(@"e:\cf2.xlsx");

            //Assert
            Assert.AreEqual(expectedFormula, condFormat.Values[1].Value);
        }
    }
}
