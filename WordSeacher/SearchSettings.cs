using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;

namespace WordSeacher
{
    public enum Target { Text = 1, Excel = 2, Word = 3 };
    public enum SearchOption { AND = 1, OR = 2, AsiS = 3 };
    public enum CharCode { UTF8 = 1, SJIS = 2 };
  
    class SearchSettings
    {
        //シングルトンオブジェクト
        private static SearchSettings _instance = new SearchSettings();
        public static SearchSettings sharedInstance()
        {
            return _instance;
        }
        //プロパティ
        private Target _target;
        public Target target{ 
            set { _target = value;
                  Properties.Settings.Default.target = (int)_target; }  
            get { return _target; }
        }
        private SearchOption _searchOption;
        public SearchOption searchOption {
            set { _searchOption = value;
                  Properties.Settings.Default.searchOption = (int)_searchOption; }
            get { return _searchOption; } 
        }
        private CharCode _charCode;
        public CharCode charCode {
            set { _charCode = value;
                  Properties.Settings.Default.CharCode = (int)_charCode; }
            get { return _charCode; }
        }
        private bool _caseSensitive;
        public bool caseSensitive {
            set { _caseSensitive = value;
                  Properties.Settings.Default.caseSensitive = _caseSensitive; }
            get { return _caseSensitive; }
        }
        private string _baseDirectory;
        public string baseDirectory{
            set { _baseDirectory = value;
                  Properties.Settings.Default.baseDirectory = _baseDirectory; }
            get { return _baseDirectory; }
        }
        //設定値の定義
        private List<string> _includeTypeList;
        public List<string> includeTypeList { 
            get { return _includeTypeList; } 
        }
        private string _includeTypes;
        public string includeTypes{
            set { 
                _includeTypes = value;
                Properties.Settings.Default.includeTypes = _includeTypes;
                _includeTypeList = _includeTypes.Split(new string[] { " ", "　" }, StringSplitOptions.RemoveEmptyEntries).ToList();
            }
            get { return _includeTypes; }
        }
        private List<string> _excludeTypeList;
        public List<string> excludeTypeList
        {
            get { return _excludeTypeList; }
        }
        private string _excludeTypes;
        public string excludeTypes{
            set { 
                _excludeTypes = value;
                Properties.Settings.Default.excludeTypes = _excludeTypes;
                _excludeTypeList = _excludeTypes.Split(new string[] { " ", "　" }, StringSplitOptions.RemoveEmptyEntries).ToList();
            }
            get { return _excludeTypes; }
        }
        private int _numTasks;
        public int numTasks{
            set { _numTasks = value; 
                  Properties.Settings.Default.numTasks = _numTasks; }
            get { return _numTasks; }
        }
        //sizing
        private int _formWidth;
        public int formWidth{
            set { _formWidth = value;
                  Properties.Settings.Default.formWidth = _formWidth; }
            get { return _formWidth; }
        }
        private int _formHeight;
        public int formHeight{
            set { _formHeight = value;
                  Properties.Settings.Default.formHeight = _formHeight; }
            get { return _formHeight; }
        }
        private int _colFolderWidth;
        public int colFolderWidth{
            set { _colFolderWidth = value;
                  Properties.Settings.Default.colFolderWidth = _colFolderWidth; }
            get { return _colFolderWidth; }
        }
        private int _colFileWidth;
        public int colFileWidth{
            set { _colFileWidth = value;
                  Properties.Settings.Default.colFileWidth = _colFileWidth; }
            get { return _colFileWidth; }
        }
        private int _colSheetWidth;
        public int colSheetWidth{
            set { _colSheetWidth = value;
                  Properties.Settings.Default.colSheetWidth = _colSheetWidth; }
            get { return _colSheetWidth; }
        }
        private int _colCountWidth;
        public int colCountWidth{
            set { _colCountWidth = value;
                  Properties.Settings.Default.colCountWidth = _colCountWidth; }
            get { return _colCountWidth; }
        }
        //コンストラクタ
        public SearchSettings()
        {
            Debug.Print("SearchSettings init !!!!!!!!");
            this.target = (Target)Properties.Settings.Default.target;
            this.searchOption = (SearchOption)Properties.Settings.Default.searchOption;
            this.charCode = (CharCode)Properties.Settings.Default.CharCode;
            this.caseSensitive = Properties.Settings.Default.caseSensitive;
            this.baseDirectory = Properties.Settings.Default.baseDirectory;

            this.includeTypes = Properties.Settings.Default.includeTypes;
            this.excludeTypes = Properties.Settings.Default.excludeTypes;
            this.numTasks = Properties.Settings.Default.numTasks;

            this.formWidth = Properties.Settings.Default.formWidth;
            this.formHeight = Properties.Settings.Default.formHeight;
            this.colFolderWidth = Properties.Settings.Default.colFolderWidth;
            this.colFileWidth = Properties.Settings.Default.colFileWidth;
            this.colSheetWidth = Properties.Settings.Default.colSheetWidth;
            this.colCountWidth = Properties.Settings.Default.colCountWidth;     
        }
        //値の保存
        public void saveAll()
        {
            Properties.Settings.Default.Save();
        }
    }
}
