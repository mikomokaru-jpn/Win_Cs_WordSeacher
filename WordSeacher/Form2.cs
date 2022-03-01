using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace WordSeacher
{
    //結果レコード
    public struct Result
    {
        public string fullPath;
        public string folder;
        public string file;
        public string sheet;
        public int count;
        public Target type;

        //コンストラクタ
        public Result(string fullPath, int count, Target type)
        {
            this.fullPath = fullPath;
            this.folder = Path.GetDirectoryName(fullPath);
            this.file = Path.GetFileName(fullPath);
            this.sheet = "";
            this.count = count;
            this.type = type;
        }
    }
    //クラス定義
    public partial class Form2 : Form
    {
        //<<<< プロパティ >>>>
        private Form1 form1 = null;

        //処理中止フラグ
        public bool cancelFlag { get; set; }

        //対象ファイル数
        private int total { get; set; }

        //処理ファイル数
        private int accumCounter = 0;
        public int counter
        {
            get { return accumCounter; }
            set
            {
                accumCounter += value;
                if (total < 100)
                { progress.Value = accumCounter; }
                else
                {
                    if (accumCounter % (total / 100) == 0)
                    { progress.Value = accumCounter; }
                }
            }
        }
        //プログレスバー
        public ProgressBar progress = new ProgressBar()
        {
            Size = new Size(280, 20),
            Location = new Point(10, 10),
            Value = 0,
        };
        //キャンセルボタン
        Button btnCancel = new Button()
        {
            Location = new Point(150 - 40, 40),
            Size = new Size(80, 25),
            Text = "キャンセル",
        };

        //設定オブジェクト
        SearchSettings serchSet = SearchSettings.sharedInstance();

        //コンストラクタ
        public Form2()
        {
            this.ClientSize = new Size(300, 75);
            this.Text = "検索";
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.ControlBox = false;
            this.Load += Form2_Load;
            //プログレスバー
            this.Controls.Add(progress);
            //キャンセルボタン
            this.Controls.Add(btnCancel);
            btnCancel.Click += new EventHandler(btnCancel_Click);
        }
        void btnCancel_Click(object sender, EventArgs e)
        {
            cancelFlag = true;
        }
        //フォーム開始時
        private void Form2_Load(Object sender, EventArgs e)
        {
            form1 = (Form1)this.Owner;
            //表示位置の調整
            var location = this.Owner.Location;
            var size = this.Owner.Size;
            var position = new Point(location.X + (size.Width / 2) - (this.Size.Width / 2),
                                     location.Y + (size.Height / 2) - (this.Size.Height / 2));
            this.Location = position;
        }
        //検索・サブスレッド起動
        public async Task search(string baseDir, string keyword)
        {
            var sw = new Stopwatch();
            sw.Start(); //測定開始
            string[] keywordList;
            if (serchSet.searchOption == SearchOption.AsiS)
            {
                keywordList = new string[] { keyword };
            }
            else
            { 
                keywordList = keyword.Split(new string[] { " ", "　" },
                StringSplitOptions.RemoveEmptyEntries); 
            }
            //ファイル一覧の取得（再帰処理）
            var fileInfoList = new List<(string, Target)>();
            traverse(baseDir, fileInfoList);

            //ファイルのグループ化 (numTasks分割）
            var num = serchSet.target == Target.Text ? serchSet.numTasks : 1;
            var portionListArray = new List<Result>[num];
            for (int i = 0; i < num; i++)
            {
                portionListArray[i] = new List<Result>();
            }
            for (int i = 0; i < fileInfoList.Count; i++)
            {
                var index = i % num;
                var result = new Result(fileInfoList[i].Item1, 0, fileInfoList[i].Item2);
                portionListArray[index].Add(result);
            }
            //プログレスバーの初期化
            progress.Minimum = 0;
            progress.Maximum = fileInfoList.Count;
            total = fileInfoList.Count;
            //処理の並列化
            var taskList = new Task<int>[num];
            var resultList = new List<Result>();
            for (int i = 0; i < num; i++)
            {
                var portiontList = portionListArray[i];
                var task = Task.Run(() =>
                {
                    //個々の検索
                    return launcher(portiontList, keywordList, resultList);
                });
                taskList[i] = task;
            }
            int[] retCode = await Task.WhenAll<int>(taskList);
            sw.Stop();
            //ウィンドウを閉じる
            this.Close();
            //結果の表示
            form1.display(resultList, total, sw.Elapsed);
        }
        //ファイル一覧の取得（再帰処理）
        private void traverse(string dir, List<(string, Target)> fileInfoList)
        {
            var dirInfo = new DirectoryInfo(dir);
            //ファイル一覧
            FileInfo[] files = dirInfo.GetFiles();
            for (int i = 0; i < files.Length; i++)
            {
                var suffix = files[i].Extension.Replace(".", "");
                //ファイルタイプによるファイルの選択
                if (serchSet.includeTypeList.Count > 0) //Inclusionあり
                {
                    if (serchSet.includeTypeList.IndexOf(suffix) >= 0)
                    { fileInfoList.Add((files[i].FullName, Target.Text)); }
                }
                else //Inclusionなし
                {
                    if (serchSet.excludeTypeList.Count > 0) //Exclusionあり
                    {
                        if (serchSet.excludeTypeList.IndexOf(suffix) < 0)
                        { fileInfoList.Add((files[i].FullName, Target.Text)); }
                    }
                    else //Exclusionなし
                    { fileInfoList.Add((files[i].FullName, Target.Text)); }
                }
            }
            //ディレクトリ一覧
            DirectoryInfo[] dirs = dirInfo.GetDirectories();
            for (int i = 0; i < dirs.Length; i++)
            {
                try
                { traverse(dirs[i].FullName, fileInfoList); }
                catch (Exception e)
                { Debug.Print("* " + e.GetType().ToString() + " : " + dirs[i].FullName); }
            }
        }

        //検索本体
        int launcher(List<Result> portiontList, string[] keywords, List<Result> resultList)
        {
            Encoding encode = Encoding.UTF8;
            if (serchSet.target == Target.Text) {
                if (serchSet.charCode == CharCode.SJIS)
                {
                    encode = Encoding.GetEncoding("shift_jis");
                }
            }
            else //Office
            {
                encode = Encoding.GetEncoding("shift_jis");
            }
            counter = 0;
            for (int idx = 0; idx < portiontList.Count; idx++)
            {
                var preResult = portiontList[idx];
                if (cancelFlag)
                {
                    return 1;
                }
                try
                {
                    var textList = new List<(string, string)>();
                    textList.Add(("", File.ReadAllText(preResult.fullPath, encode)));
                    foreach ((string, string) record in textList)
                    {
                        var count = 0;
                        var regOpt = serchSet.caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
                        switch (serchSet.searchOption)
                        {
                            case SearchOption.AND:
                                for (int i = 0; i < keywords.Length; i++)
                                {
                                    var num = Regex.Matches(record.Item2, keywords[i], regOpt).Count;
                                    if (num > 0)
                                    {
                                        count += num;
                                    }
                                    else
                                    {
                                        count = 0;
                                        break;
                                    }
                                }
                                //var matches = Regex.Matches(text, "^(?=.*ううう)(?=.*あああ)", RegexOptions.Singleline);
                                break;
                            case SearchOption.OR:
                                var pattern = "";
                                for (int i = 0; i < keywords.Length; i++)
                                {
                                    pattern += keywords[i] + "|";
                                }
                                pattern = pattern.Remove(pattern.Length - 1, 1);
                                count = Regex.Matches(record.Item2, pattern, regOpt).Count;
                                break;
                            case SearchOption.AsiS:
                                count = Regex.Matches(record.Item2, keywords[0], regOpt).Count;
                                break;
                            default:
                                break;
                        }
                        //結果の記録・メインスレッドで実行することによりシリアル処理となる
                        this.Invoke((Action)(() =>
                        {
                            if (count > 0)
                            {
                                //結果レコードの作成
                                var result = new Result(preResult.fullPath, count, preResult.type);
                                result.sheet = record.Item1;
                                resultList.Add(result);
                            }
                        }));
                    }
                    this.Invoke((Action)(() =>
                    {
                        counter = 1;
                    }));
                }
                catch (IOException e)
                {
                    Debug.Print("+ " + e.GetType().ToString() + " : " + preResult.fullPath);
                    Debug.Print(e.Message);
                    continue;
                }
            }
            return 0;
        }
    }
}
