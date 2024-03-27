using Logic;
using Logic.Utility;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;


#nullable disable
namespace IS.ViewModels
{
    /// <summary>
    /// VM для основного окна
    /// </summary>
    public class MainVewModel : INotifyPropertyChanged
    {
        DateTime month;
        public DateTime Month
        {
            get
            {
                return month;
            }
            set
            {
                month = value;
                OnPropertyChanged(nameof(Month));
            }
        }

        public string Teacher
        {
            get
            {
                return teacher;
            }
            set
            {
                teacher = value;
                OnPropertyChanged(nameof(Teacher));
            }
        }

        readonly Dictionary<int, string> kvp = new()
        {
            {1, "Январь" },
            {2, "Февраль" },
            {3, "Март" },
            {4, "Апрель" },
            {5, "Май" },
            {6, "Июнь" },
            {7, "Июль" },
            {8, "Август" },
            {9, "Сентябрь" },
            {10, "Октябрь" },
            {11, "Ноябрь" },
            {12, "Декабрь" }

        };

        string teacher;
        string calendarPath;
        string workloadPath;

        readonly string outputPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        Report Result { get; set; }

        List<Group> workloadData;
        List<Group> calendarData;

        public event PropertyChangedEventHandler? PropertyChanged;

        #region commands
        Command _load1;
        public Command Load1
        {
            get
            {
                return _load1 ??= new Command(obj => LoadCalendar());
            }
        }

        Command _load2;
        public Command Load2
        {
            get
            {
                return _load2 ??= new Command(obj => LoadWorkload());
            }
        }

        Command _start;
        public Command Start
        {
            get
            {
                return _start ??= new Command(obj => LoadCalendar());
            }
        }

        Command _save;

        public Command Save
        {
            get
            {
                return _save ??= new Command(obj => SaveFile());
            }
        }

        #endregion

        public MainVewModel()
        {
            month = new();
        }

        public void LoadCalendar()
        {
            OpenFileDialog ofd = new();
            ofd.Filter = "Excel File (*.xls)|*.xls|Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            if (ofd.ShowDialog() == true)
            {
                calendarPath = ofd.FileName;
            }

            calendarData = CalendarParser.Parse(calendarPath, kvp[month.Month]);
        }

        public void LoadWorkload()
        {
            if (Teacher != null)
            {

                OpenFileDialog ofd = new();
                ofd.Filter = "Excel File (*.xls)|*.xls|Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                if (ofd.ShowDialog() == true)
                {
                    workloadPath = ofd.FileName;
                }

                workloadData = WorkloadParser.Parse(workloadPath, Teacher);
            }
            else
            {
                MessageBox.Show("Имя преподавателя не может быть пустым", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Stop);
            }
        }

        public void DoTheWork()
        {
            // TODO: Change workloadData to preprocessedData
            Report report = ReportShaper.Shape(month, teacher, workloadData);
        }

        public void SaveFile()
        {
            FileWriter.WriteFile(Result, outputPath);
        }

        public void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
