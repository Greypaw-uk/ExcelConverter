using System.Collections.ObjectModel;
using System.ComponentModel;

namespace Excel_Converter
{
    class MainWindowDataContext
    {
        public static ObservableCollection<FileContents> CSVFileContents { get; set; } = new ObservableCollection<FileContents>();

        public static ObservableCollection<FileContents> Subject { get; set; } = new ObservableCollection<FileContents>();
    }

    public class PupilInfo
    {
        public string name { get; set; }
        public string au1 { get; set; }
        public string au2 { get; set; }

        public string sp1 { get; set; }
        public string sp2 { get; set; }

        public string su1 { get; set; }
        public string su2 { get; set; }
    }


    public class YearGroup
    {
        private string _selectedYearGroup;
        public string selectedYearGroup
        {
            get { return _selectedYearGroup; }
            set
            {
                _selectedYearGroup = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("selectedYearGroup"));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }

    public class Summatives
    {
        private string _emerging;
        public string emerging
        {
            get { return _emerging; }
            set
            {
                _emerging = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("emerging"));
            }
        }

        private string _developing;
        public string developing
        {
            get { return _developing; }
            set
            {
                _developing = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("developing"));
            }
        }

        private string _secure;
        public string secure
        {
            get { return _secure; }
            set
            {
                _secure = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("secure"));
            }
        }

        private string _GD;
        public string greaterDepth
        {
            get { return _GD; }
            set
            {
                _GD = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("greaterDepth"));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
