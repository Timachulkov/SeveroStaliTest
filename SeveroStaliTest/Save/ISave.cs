using SeveroStaliTest.Filter;
using System.Collections.Generic;

namespace SeveroStaliTest
{
    interface ISave
    {
        bool Save(string DataSource, IEnumerable<FilteredData> data);
    }
}
