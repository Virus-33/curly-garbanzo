using Logic;
using Logic.Utility;
using Logic.WorkloadParser;
using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace IS.ViewModels
{
    /// <summary>
    /// VM для основного окна
    /// </summary>
    public class MainVewModel : INotifyPropertyChanged
    {
        public DateTime month { get; set; }
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

        string teacher;
        string calendarPath;
        string workloadPath;

        string outputPath;

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

        }

        public void LoadCalendar()
        {
            // TODO: Add logic for path retrieving BEFORE other actions
            calendarData = CalendarParser.Parse(calendarPath);
        }

        public void LoadWorkload()
        {
            // TODO: Add logic for path retrieving BEFORE other actions
            workloadData = WorkloadParser.Parse(workloadPath, Teacher);
        }

        public void DoTheWork()
        {
            Report report = ReportShaper.Shape(month, teacher, workloadData, calendarData);
        }

        public void SaveFile()
        {
            // TODO: Add logic for path retrieving BEFORE other actions
            FileWriter.WriteFile(Result, outputPath);
        }

        public void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
