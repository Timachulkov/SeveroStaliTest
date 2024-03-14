using SeveroStaliTest.Filter;
using System.Collections.Generic;
using System.Linq;

namespace SeveroStaliTest
{
    class Filtration
    {
        public IEnumerable<FilteredData> Filter(DataCollection data)
        {
            IEnumerable<StructuredData> StructuredData = data.Employees.Join(data.Tasks,
                 e => e.TableID,
                 t => t.TableID,
                 (e, t) => new
                 {
                     FullName = e.Name.Trim() + " " + e.Surname[0] + "." + (e.Patronymic != null ? e.Patronymic[0] + "." : string.Empty),
                     e.DepartmentID,
                     t.TaskId,
                 }).Join(data.Departments,
                 e => e.DepartmentID,
                 d => d.DepartmentsId,
                 (e, d) => new StructuredData
                 {
                     DepartmentsName = d.DepartmentsName,
                     FullName = e.FullName,
                     TaskId = e.TaskId,
                 });

            return StructuredData.GroupBy(x => x.DepartmentsName)
                  .Select(x => new FilteredData
                  {
                      DepartmentsName = x.Key,
                      TaskNum = x.Count(),
                      FilteredName = x.GroupBy(n => n.FullName).Select(n => new FilteredName
                      {
                          Name = n.Key,
                          TaskNum = n.Count()
                      }).OrderByDescending(n => n.TaskNum)
                  }).OrderByDescending(n => n.TaskNum);
        }

        struct StructuredData
        {
            public string DepartmentsName;
            public string FullName;
            public ulong TaskId;
        }
    }
}
