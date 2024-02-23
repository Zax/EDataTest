using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NUnit.Framework;

namespace EDataTest
{
	public class Tests
	{
		private readonly string _path = AppDomain.CurrentDomain.BaseDirectory + @"\";

		[Test]
		public void TestEDataFunction()
		{
			using var fileStream = new FileStream(_path + "EDATEBug.xls", FileMode.Open, FileAccess.Read, FileShare.Read);
			IWorkbook workbook = new HSSFWorkbook(fileStream, true);
			HSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);
		}

	}
}
