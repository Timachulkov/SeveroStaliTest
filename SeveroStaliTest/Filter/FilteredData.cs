using System.Collections.Generic;

namespace SeveroStaliTest.Filter
{
    class FilteredData
    {
        public string DepartmentsName { get; set; }
        public int TaskNum { get; set; }
        public IEnumerable<FilteredName> FilteredName { get; set; }
    }
    class FilteredName
    {
        public string Name { get; set; }
        public int TaskNum { get; set; }
    }
}
