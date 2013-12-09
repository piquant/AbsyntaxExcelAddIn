using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AbsyntaxExcelAddInTestProject.Converters
{
    [TestClass]
    public abstract class PngUriConverterBaseTest<T> : UriSchemeRegistrar
    {
        [TestMethod]
        public void ConvertsEveryEnumValueTest()
        {
            var values = Enum.GetValues(typeof(T)).Cast<T>().Except(Exclusions);
            IValueConverter c = GetConverter();
            foreach (object value in values) {
                Uri uri = c.Convert(value, null, null, null) as Uri;
                Assert.IsNotNull(uri);
            }
        }

        protected abstract IValueConverter GetConverter();

        protected virtual IEnumerable<T> Exclusions
        {
            get { return new T[0]; }
        }
    }
}