using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace WordSeacher
{
    public partial class Form1 : Form
    {
        //データグリッドビューWrapper
        UADataGridView dg = new UADataGridView();
        //フォルダブラウズダイアログ
        FolderBrowserDialog fbd = new FolderBrowserDialog();
        //テキストボックス・起点パス
        TextBox baseDir = new TextBox()
        {
            Location = new Point(80, 10),
            Size = new Size(255, 25),
            Anchor = (AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right),
            TabIndex = 40,
        };
        //テキストボックス・検索語
        TextBox keyword = new TextBox()
        {
            Location = new Point(80, 45),
            Size = new Size(255, 25),
            Anchor = (AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right),
            TabIndex = 10,
        };
        //テキストボックス・件数
        TextBox fileCounter = new TextBox()
        {
            Location = new Point(10, 310),
            Size = new Size(100, 25),
            ReadOnly = true,
            Anchor = (AnchorStyles.Bottom | AnchorStyles.Left),
            TabStop = false,
        };
        //テキストボックス・処理時間
        TextBox elapsTime = new TextBox()
        {
            Location = new Point(120, 310),
            Size = new Size(100, 25),
            ReadOnly = true,
            Anchor = (AnchorStyles.Bottom | AnchorStyles.Left),
            TabStop = false,
        };
        //検索条件の表示
        Label searchInfo = new Label()
        {
            Location = new Point(230, 315),
            Size = new Size(200, 25),
            Text = "",
            Anchor = (AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right),

        };
        //設定オブジェクト
        SearchSettings serchSet = SearchSettings.sharedInstance();

        //コンストラクタ
        public Form1()
        {
            this.ClientSize = new Size(400, 340);
            this.Text = "Word Searcher";     // タイトルを設定
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.Load += new EventHandler(form1_Load);
            this.FormClosing += new FormClosingEventHandler(form1_FormClosing);
            //ラベル
            Label label1 = new Label()
            {
                Location = new Point(10, 13),
                Size = new Size(70, 25),
                Text = "対象フォルダ",
            };
            this.Controls.Add(label1);
            Label label2 = new Label()
            {
                Location = new Point(10, 48),
                Size = new Size(70, 25),
                Text = "検索語",
            };
            this.Controls.Add(label2);
            //選択ボタン
            Button btnOpenFile = new Button()
            {
                Location = new Point(340, 8),
                Size = new Size(50, 22),
                Text = "選択",
                Anchor = (AnchorStyles.Top | AnchorStyles.Right),
                TabIndex = 50,
            };
            this.Controls.Add(btnOpenFile);
            btnOpenFile.Click += new EventHandler(btnOpenFile_Click);
            //検索ボタン
            Button btnExecute = new Button()
            {
                Location = new Point(340, 39),
                Size = new Size(50, 32),
                Text = "検索",
                Font = new Font(label1.Font.FontFamily, 12, label1.Font.Style),
                Anchor = (AnchorStyles.Top | AnchorStyles.Right),
                TabIndex = 20,
            };
            this.Controls.Add(btnExecute);
            btnExecute.Click += new EventHandler(btnExecute_Click);
            //設定フォームを開くボタン
            Button btnSettings = new Button()
            {
                Location = new Point(350, 305),
                Size = new Size(40, 30),
                Text = "設定",
                Anchor = (AnchorStyles.Bottom | AnchorStyles.Right),
                TabIndex = 30,
            };
            this.Controls.Add(btnSettings);
            btnSettings.Click += new EventHandler(btnSettings_Click);
            //フォルダブラウズダイアログ
            fbd.Description = "ディレクトリの選択";
            fbd.RootFolder = Environment.SpecialFolder.Desktop;
            if (serchSet.baseDirectory == "")
                fbd.SelectedPath = @"C:\Users\" + Environment.UserName + @"\Desktop\work";
            else
                fbd.SelectedPath = serchSet.baseDirectory;
            fbd.ShowNewFolderButton = false;
            //テキストボックス・パス 
            this.Controls.Add(baseDir);
            baseDir.Text = fbd.SelectedPath;
            //テキストボックス・検索語
            this.Controls.Add(keyword);

            //データグリッドビュー・結果一覧 
            this.Controls.Add(dg.resultView);
            dg.setUp();
            //デリゲートの代入
            dg.funcMenu1 = new menuOperation(openFile1);
            dg.funcMenu2 = new menuOperation(openFile2);
            dg.funcMenu3 = copyPath; //これもOK
            
            //テキストボックス・件数
            this.Controls.Add(fileCounter);
            //テキストボックス・処理時間
            this.Controls.Add(elapsTime);
            //テキストボックス・検索条件
            this.Controls.Add(searchInfo);
            //検索条件の表示
            this.DispSerachInfo();
        }
        //ファイルを開く
        void openFile1(string path)
        {
            Debug.Print("menuItem1 " + path);
            var ps = new Process();
            ps.StartInfo.FileName = path;
            ps.Start();
        }
        //Codeでファイルを開く
        void openFile2(string path)
        {
            Process.Start("code", path);
        }
        //パスをコピーする
        void copyPath(string path)
        {
            Clipboard.SetText(path);
        }
        //フォルダ選択ボタンクリック
        void btnOpenFile_Click(object sender, EventArgs e)
        {
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                baseDir.Text = fbd.SelectedPath;
            }
        }
        //検索ボタンクリック
        async void btnExecute_Click(object sender, EventArgs e)
        {
            //結果一覧のクリア
            dg.dataTable.Clear();
            //ディレクトリ存在チェック
            if (!Directory.Exists(baseDir.Text))
            {
                MessageBox.Show(string.Format("{0}が存在しません", baseDir), "Directory not found",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //検索語空白チェック
            if (keyword.Text.Trim().Length == 0)
            {
                MessageBox.Show("検索語が入力されていません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //検索処理
            var form2 = new Form2();
            form2.Show(this);
            //キーワード編集
            var keywordList = keyword.Text.Split(new string[] { " ", "　" }, StringSplitOptions.RemoveEmptyEntries);
            //++++ 検索処理の呼び出し +++++
            await form2.search(baseDir.Text, keyword.Text);
        }
        //設定フォームを開くボタンクリック
        void btnSettings_Click(object sender, EventArgs e)
        {
            var formSettings = new FormSettings();
            formSettings.ShowDialog(this);
            this.DispSerachInfo();
        }
        //結果一覧の表示
        public void display(List<Result> resultList, int total, TimeSpan elaps)
        {
            if (resultList.Count == 0)
            {
                MessageBox.Show("対象データがありません。", "結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            //結果一覧の表示本体
            dg.display(resultList);
            fileCounter.Text = string.Format("{0} / {1}  {2}%",
                resultList.Count, total, (int)((float)resultList.Count / (float)total * 100));
            elapsTime.Text = elaps.ToString();
        }
        //フォーム表示時
        private void form1_Load(object sender, EventArgs e)
        {
            //設定値の読み込み・様々なサイズ
            this.Width = serchSet.formWidth;
            this.Height = serchSet.formHeight;
            this.ActiveControl = keyword;
        }
        //フォーム終了時
        private void form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //設定値の保存・様々なサイズ
            serchSet.formWidth = this.Width;
            serchSet.formHeight = this.Height;
            serchSet.baseDirectory = baseDir.Text;
            serchSet.saveAll();
        }
        //検索条件の表示
        private void DispSerachInfo()
        {
            var text = "Searching ";
            switch (serchSet.target)
            {
                case Target.Text:
                    text += "Text";
                    break;
                case Target.Excel:
                    text += "Excel";
                    break;
                case Target.Word:
                    text += "Word";
                    break;
            }
            switch (serchSet.searchOption)
            {
                case SearchOption.AND:
                    text += " ope=AND";
                    break;
                case SearchOption.OR:
                    text += " ope=OR";
                    break;
                case SearchOption.AsiS:
                    text += " ope=asis";
                    break;
            }
            switch (serchSet.charCode)
            {
                case CharCode.UTF8:
                    text += " UTF-8";
                    break;
                case CharCode.SJIS:
                    text += " shift_jis";
                    break;
            }
            if (serchSet.caseSensitive)
            {
                text += " caseSensitive";
            }
            else
            {
                text += " NotCaseSensitive";
            }
            this.searchInfo.Text = text;
        }
    }
}
