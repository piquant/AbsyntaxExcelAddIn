using System.Windows;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AbsyntaxExcelAddInTestProject.Converters
{
    [TestClass]
    public abstract class UriSchemeRegistrar
    {
        [AssemblyInitialize]
        public static void AssemblyInitialize(TestContext context)
        {
            /* The following serves to register both the "pack" and "application" schemes with the Uri parser.
             * See http://stackoverflow.com/questions/6005398/uriformatexception-invalid-uri-invalid-port-specified.
             * */
            var current = Application.Current;
        }
    }
}