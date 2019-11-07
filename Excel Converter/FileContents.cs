using System.ComponentModel;

namespace Excel_Converter
{
    public class FileContents : INotifyPropertyChanged
    { 
        private string _pupilName;
         public string pupilName
        {
            get { return _pupilName; }
            set
            {
                _pupilName = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("pupilName"));
            }
        }

        private string _au1;
        public string au1
        {
            get { return _au1;}
            set
            {
                _au1 = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("au1"));
            }
        }

        private string _au2;
        public string au2
        {
            get { return _au2; }
            set
            {
                _au2 = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("au2"));
            }
        }

        private string _sp1;
        public string sp1
        {
            get { return _sp1; }
            set
            {
                _sp1 = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("sp1"));
            }
        }

        private string _sp2;
        public string sp2
        {
            get { return _sp2; }
            set
            {
                _sp2 = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("sp2"));
            }
        }

        private string _su1;
        public string su1
        {
            get { return _su1; }
            set
            {
                _su1 = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("su1"));
            }
        }

        private string _su2;
        public string su2
        {
            get { return _su2; }
            set
            {
                _su2 = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("su2"));
            }
        }

        public int Length { get; internal set; }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
