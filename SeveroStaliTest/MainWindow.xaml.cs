using Microsoft.Win32;
using SeveroStaliTest.Filter;
using System;
using System.Collections.Generic;
using System.Windows;

namespace SeveroStaliTest
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        IEnumerable<FilteredData> FiltredData;
        State AccessState = State.Open;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ChoseFile_Click(object sender, RoutedEventArgs e)
        {
            if (AccessState != State.CloseCollector)
            {
                CloseCollectorState();
            }
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = "c:\\",
                Filter = "Excel(Xlsb) files (*.Xlsb)|*.Xlsb|All files (*.*)|*.*",
                FilterIndex = 2,
                RestoreDirectory = true
            };

            try
            {
                if (openFileDialog.ShowDialog() == true)
                {
                    var filePath = openFileDialog.FileName;
                    IDataCollector collector = null;
                    switch (System.IO.Path.GetExtension(filePath))
                    {
                        case ".xlsb":
                            collector = new XlsbCollector();
                            break;
                        default:
                            throw new InvalidOperationException("Не найден возможный тип документа");
                    }
                    if (collector != null)
                    {
                        DataCollection Data = collector.Read(filePath);
                        Filtration filtre = new Filtration();
                        FiltredData = filtre.Filter(Data);

                        DemoShowGridData.ItemsSource = FiltredData;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (AccessState != State.Open)
                {
                    OpenState();
                }
            }
        }

        private void SaveAsWord_Click(object sender, RoutedEventArgs e)
        {
            if (AccessState != State.CloseSave)
            {
                CloseSaveState();
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                FileName = "Untitled",
                DefaultExt = ".docx",
                Filter = "Xml - Документ Word |*.xml|Документ Word|*.docx",
                Title = "Save as Word"
            };
            if (saveFileDialog.ShowDialog() == true)
            {
                ISave save = new SaveWord();
                string fileName = saveFileDialog.FileName;
                save.Save(fileName, FiltredData);
            }

            if (AccessState != State.Open)
            {
                OpenState();
            }
        }

        void OpenState()
        {
            AccessState = State.Open;
            ChoseFile.IsEnabled = true;
            SaveAsWord.IsEnabled = true;
            StatusLable.Content = string.Empty;
        }
        void CloseCollectorState()
        {
            AccessState = State.CloseCollector;
            ChoseFile.IsEnabled = false;
            SaveAsWord.IsEnabled = false;
            StatusLable.Content = "Происходит загрузка данных ...";
        }
        void CloseSaveState()
        {
            AccessState = State.CloseSave;
            ChoseFile.IsEnabled = false;
            SaveAsWord.IsEnabled = false;
            StatusLable.Content = "Происходит выгрузка данных ...";
        }

        enum State
        {
            Open,
            CloseCollector,
            CloseSave,
        }
    }
}
