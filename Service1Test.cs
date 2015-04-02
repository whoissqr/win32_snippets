using AutoWeeklyReport;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace XapTesterStatus
{
    
    
    /// <summary>
    ///This is a test class for Service1Test and is intended
    ///to contain all Service1Test Unit Tests
    ///</summary>
    [TestClass()]
    public class Service1Test
    {


        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        /// <summary>
        ///A test for getStartOfLastWeek
        ///</summary>
        [TestMethod()]
        public void getStartOfLastWeekTest()
        {
            Service1 target = new Service1(); // TODO: Initialize to an appropriate value
            DateTime expected = new DateTime(2014,8,17,7,0,0); // TODO: Initialize to an appropriate value
            DateTime actual;          
            actual = target.getStartOfLastWeek();
            Console.WriteLine(actual);
            Assert.AreEqual(expected, actual);
            //Assert.Inconclusive("Verify the correctness of this test method.");
        }
    }
}
