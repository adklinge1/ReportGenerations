namespace GenerateReportTests.TreeCalculator
{
    using System.Threading.Tasks;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using TreeCalculator = WindowsFormsApp1.Models.TreeCalculator;
    using System.Collections.Generic;
    using WindowsFormsApp1.ExcelReader;
    using WindowsFormsApp1.Models;

    [TestClass]
    public class TreeCalculatorTests
    {
        TreeCalculator _calculator = new TreeCalculator();
        private Dictionary<string, TreeSpecie> _treeSpecies;


        [TestInitialize]
        public async Task Init()
        {
            await _calculator.LoadTreePricesAsync();
            _treeSpecies = ExcelReader.TryReadTreeSpecies();
        }

        [TestMethod]
        [DataRow(10, 0.8, 0, "סבל פלמטו", 0)]
        [DataRow(10, 0.8, 0.2, "סבל פלמטו", 1464)]
        [DataRow(5.2, 0.6, 0.4, "דיפסיס משולש", 936)]
        [DataRow(0.5, 0.6, 0.4, "דיפסיס משולש", 90)]
        [DataRow(2, 0.2, 1, "בוטיה דרומית", 372)]
        public async Task TryToGetTreePrice_ForPalm_ReturnsTheRightPrice(double height, double location, double health, string species, double expectedPrice)
        {
            string scientficName = _treeSpecies[species].ScientificName;
            Assert.IsNotNull(scientficName);

            Tree tree = new Tree(1, species, height, stemDiameter: 10, healthRate: (int) (5 * health), 10, (int) (5 * location), -1, -1,scientficName);

            double ? price = _calculator.TryToGetTreePrice(tree);

            Assert.AreEqual(expectedPrice, price);
        }
    }
}
