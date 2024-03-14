using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Windows;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace SeveroStaliTest
{
    class XlsbCollector : IDataCollector
    {
        public DataCollection Read(string DataSource)
        {
            return XlsbRead(DataSource);
        }
        private DataCollection XlsbRead(string DataSource)
        {
            Application excel = new Application();
            try
            {
                Workbook wb = excel.Workbooks.Open(DataSource);
                DataCollection data = new DataCollection();

                //Заполенение Сотрудников
                Worksheet EmployeesSheet = wb.Worksheets[1];
                List<Employees> ListEmployers = new List<Employees>();
                bool breakOut = true;
                int i = 1;
                do
                {
                    i++;
                    Range Cells = EmployeesSheet.Range[$"A{i}:F{i}"];
                    var show = Cells.Value;
                    if (show[1, 1] == null)
                    {
                        breakOut = false;
                        continue;
                    }
                    Employees employe = new Employees
                    {
                        TableID = Convert.ToUInt64(show[1, 1]),
                        Name = show[1, 2],
                        Surname = show[1, 3],
                        Patronymic = show[1, 4],
                        BirthDate = (DateTime)show[1, 5],
                        DepartmentID = Convert.ToUInt64(show[1, 6])
                    };
                    ListEmployers.Add(employe);
                }
                while (breakOut);
                data.Employees = ListEmployers.ToArray();

                //Заполенение Отделов
                Worksheet DepartmentsSheet = wb.Worksheets[2];
                List<Departments> ListDepartments = new List<Departments>();
                breakOut = true;
                i = 1;
                do
                {
                    i++;
                    Range Cells = DepartmentsSheet.Range[$"A{i}:B{i}"];
                    var show = Cells.Value;
                    if (show[1, 1] == null)
                    {
                        breakOut = false;
                        continue;
                    }
                    Departments departments = new Departments
                    {
                        DepartmentsId = Convert.ToUInt64(show[1, 1]),
                        DepartmentsName = show[1, 2],
                    };
                    ListDepartments.Add(departments);
                }
                while (breakOut);
                data.Departments = ListDepartments.ToArray();

                //Заполенение Задач
                Worksheet TasksSheet = wb.Worksheets[3];
                List<Tasks> ListTasks = new List<Tasks>();
                breakOut = true;
                i = 1;
                do
                {
                    i++;
                    Range Cells = TasksSheet.Range[$"A{i}:B{i}"];
                    var show = Cells.Value;
                    if (show[1, 1] == null)
                    {
                        breakOut = false;
                        continue;
                    }
                    Tasks tasks = new Tasks
                    {
                        TaskId = Convert.ToUInt64(show[1, 1]),
                        TableID = Convert.ToUInt64(show[1, 2]),
                    };
                    ListTasks.Add(tasks);
                }
                while (breakOut);
                data.Tasks = ListTasks.ToArray();

                return data;
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
                return null;
            }
            finally
            {
                excel.Quit();
            }
        }
    }
}
