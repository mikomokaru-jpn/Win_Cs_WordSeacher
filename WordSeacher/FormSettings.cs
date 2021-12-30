using System;
using System.Drawing;
using System.Windows.Forms;

namespace WordSeacher
{ 
    class FormSettings : Form
    {
        //設定オブジェクト
        SearchSettings serchSet = SearchSettings.sharedInstance();

        //-------------------------------------------------------------
        // パネル1
        //-------------------------------------------------------------
        Panel panel1 = new Panel()
        {
            Location = new Point(10, 10),
            Size = new Size(340, 260),
            BorderStyle = BorderStyle.FixedSingle,
        };
        //検索対象ファイルタイプ
        RadioButton radioText = new RadioButton()
        {
            Location = new Point(10, 10),
            Text = "テキスト検索",
            TextAlign = ContentAlignment.MiddleLeft,
        };
        //対象ファイルタイプ
        Label label1 = new Label
        {
            Location = new Point(10, 40),
            Size = new Size(70, 20),
            Text = "対象拡張子",
        };
        TextBox tbIncludeTypes = new TextBox()
        {
            Location = new Point(80, 35),
            Size = new Size(250, 25),
        };
        Label label3 = new Label
        {
            Location = new Point(80, 60),
            Size = new Size(300, 25),
            Text = "指定がない場合すべてのファイルタイプを対象とする",
        };
        //対象外ファイルタイプ
        Label label2 = new Label
        {
            Location = new Point(10, 90),
            Size = new Size(70, 20),
            Text = "除外拡張子",
        };
        TextBox tbExcludeTypes = new TextBox()
        {
            Location = new Point(80, 85),
            Size = new Size(250, 25),
        };
        Label label4 = new Label
        {
            Location = new Point(80, 110),
            Size = new Size(300, 25),
            Text = "対象拡張子の指定がない場合のみ有効",
        };
        //パネル1a
        Panel panel1a = new Panel()
        {
            Location = new Point(20, 130),
            Size = new Size(310, 30),
            BorderStyle = BorderStyle.None,
        };
        //文字コード
        RadioButton radioUTF8 = new RadioButton()
        {
            Location = new Point(60, 2),
            TextAlign = ContentAlignment.MiddleLeft,
            Size = new Size(70, 25),
            Text = "UTF-8"
        };
        RadioButton radioSJIS = new RadioButton()
        {
            Location = new Point(130, 2),
            TextAlign = ContentAlignment.MiddleLeft,
            Size = new Size(70, 25),
            Text = "Shift-JIS"
        };

        //並列タスク数
        Label label5 = new Label
        {
            Location = new Point(10, 160),
            Size = new Size(70, 20),
            Text = "並列タスク数",
            TextAlign = ContentAlignment.MiddleLeft,
        };
        ComboBox comNumTasks = new ComboBox()
        {
            Location = new Point(80, 160),
            Size = new Size(50, 25),
        };
        //-------------------------------------------------------------
        Label line = new Label
        {
            Location = new Point(0, 190),
            Size = new Size(340, 1),
            BorderStyle = BorderStyle.FixedSingle,
        };
        //Excelデータ
        RadioButton radioExcel = new RadioButton()
        {
            Location = new Point(10, 195),
            TextAlign = ContentAlignment.MiddleLeft,
            Size = new Size(150, 25),
            Text = "Excel データ検索"
        };
        Label line2 = new Label
        {
            Location = new Point(0, 225),
            Size = new Size(340, 1),
            BorderStyle = BorderStyle.FixedSingle,
        };
        //Wordデータ
        RadioButton radioWord = new RadioButton()
        {
            Location = new Point(10, 230),
            TextAlign = ContentAlignment.MiddleLeft,
            Size = new Size(150, 25),
            Text = "Word データ検索"
        };
        //-------------------------------------------------------------
        // パネル2
        //-------------------------------------------------------------
        Panel panel2 = new Panel()
        {
            Location = new Point(10, 280),
            Size = new Size(340, 60),
            BorderStyle = BorderStyle.FixedSingle,
        };
        Label panel2Label = new Label
        {
            Location = new Point(10, 12),
            Size = new Size(150, 20),
            Text = "検索語を空白で分割して",
        };
        //AND検索
        RadioButton radioAND = new RadioButton()
        {
            Location = new Point(10, 30),
            Size = new Size(80, 20),
            TextAlign = ContentAlignment.MiddleLeft,
            Text = "AND検索"
        };
        //OR検索
        RadioButton radioOR = new RadioButton()
        {
            Location = new Point(90, 30),
            Size = new Size(80, 20),
            TextAlign = ContentAlignment.MiddleLeft,
            Text = "OR検索"
        };
        //文字列をそのまま検索
        RadioButton radioAsIs = new RadioButton()
        {
            Location = new Point(170, 10),
            Size = new Size(150, 20),
            TextAlign = ContentAlignment.MiddleLeft,
            Text = "文字列をそのまま検索"
        };
        //-------------------------------------------------------------
        // パネル3
        //-------------------------------------------------------------
        Panel panel3 = new Panel()
        {
            Location = new Point(10, 350),
            Size = new Size(340, 40),
            BorderStyle = BorderStyle.FixedSingle,
        };
        //大文字小文字区別する
        CheckBox checkCase = new CheckBox()
        {
            Location = new Point(10, 10),
            Size = new Size(150, 20),
            TextAlign = ContentAlignment.MiddleLeft,
            Text = "大文字小文字区別する"
        };
        //コンストラクタ
        public FormSettings()
        {
            this.ClientSize = new Size(360, 430);
            this.Text = "設定";     // タイトルを設定
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.Load += form_Load;
            this.FormClosing += form_FormClosing;
            this.KeyPress += Form2_KeyPress;
            this.KeyPreview = true;
            //コントロールの追加
            this.Controls.Add(panel1);
            panel1.Controls.Add(radioText);
            panel1.Controls.Add(label1);
            panel1.Controls.Add(tbIncludeTypes);
            panel1.Controls.Add(label3);
            panel1.Controls.Add(label2);
            panel1.Controls.Add(tbExcludeTypes);
            panel1.Controls.Add(label4);
            panel1.Controls.Add(panel1a);
            panel1a.Controls.Add(radioUTF8);
            panel1a.Controls.Add(radioSJIS);
            panel1.Controls.Add(label5);
            panel1.Controls.Add(comNumTasks);
            comNumTasks.Items.AddRange(new string[] { "1", "10" });
            panel1.Controls.Add(line);
            panel1.Controls.Add(radioExcel);
            panel1.Controls.Add(line2);
            panel1.Controls.Add(radioWord);
            radioText.Checked = true;
            radioUTF8.Checked = true;
            radioText.Click += new EventHandler(radio_Click);
            radioExcel.Click += new EventHandler(radio_Click);
            radioWord.Click += new EventHandler(radio_Click);
            this.Controls.Add(panel2);
            panel2.Controls.Add(panel2Label);
            panel2.Controls.Add(radioAND);
            panel2.Controls.Add(radioOR);
            panel2.Controls.Add(radioAsIs);
            radioAND.Checked = true;
            this.Controls.Add(panel3);
            panel3.Controls.Add(checkCase);
        }
        //フォームの開始
        void form_Load(object sender, EventArgs e)
        {
            //表示位置の調整
            var form1 = (Form1)this.Owner;
            //表示位置の調整
            var location = this.Owner.Location;
            var size = this.Owner.Size;
            var position = new Point(location.X + (size.Width / 2) - (this.Size.Width / 2),
                                     location.Y + (size.Height / 2) - (this.Size.Height / 2));
            this.Location = position;
            //設定の表示
            this.setValue();
        }
        //テキスト <-> Office 切り替え
        void radio_Click(object sender, EventArgs e)
        {
            var target = Target.Text;
            if (radioExcel.Checked)
            {
                target = Target.Excel;
            }
            else if (radioWord.Checked)
            {
                target = Target.Word;
            }
            this.changeStatus(target);
        }
        //フォームの終了
        void form_FormClosing(object sender, FormClosingEventArgs e)
        {
            //値の保存
            if (radioText.Checked)
                serchSet.target = Target.Text;
            else if (radioExcel.Checked)
                serchSet.target = Target.Excel;
            else
                serchSet.target = Target.Word;
            serchSet.includeTypes = tbIncludeTypes.Text;
            serchSet.excludeTypes = tbExcludeTypes.Text;
            var w = int.Parse(comNumTasks.SelectedItem.ToString());
            serchSet.numTasks = w;
            if (radioUTF8.Checked)
                serchSet.charCode = CharCode.UTF8;
            else
                serchSet.charCode = CharCode.SJIS;
            if (radioAND.Checked)
                serchSet.searchOption = SearchOption.AND;
            else if (radioOR.Checked)
                serchSet.searchOption = SearchOption.OR;
            else
                serchSet.searchOption = SearchOption.AsiS;
            serchSet.caseSensitive = checkCase.Checked;
            serchSet.saveAll();
        }
        //press escape key 
        private void Form2_KeyPress(Object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Escape)
            {
                e.Handled = true;
                this.Close();
            }
        }
        //値の設定
        private void setValue()
        {
            this.changeStatus(serchSet.target);
            tbIncludeTypes.Text = serchSet.includeTypes;
            tbExcludeTypes.Text = serchSet.excludeTypes;
            if (serchSet.charCode == CharCode.UTF8)
                radioUTF8.Checked = true;
            else
                radioSJIS.Checked = true;
            comNumTasks.Text = serchSet.numTasks.ToString();
            var op = serchSet.searchOption;
            if (op == SearchOption.AND)
                radioAND.Checked = true;
            else if (op == SearchOption.OR)
                radioOR.Checked = true;
            else
                radioAsIs.Checked = true;
            checkCase.Checked = serchSet.caseSensitive;
        }
        //状態の切り替え
        private void changeStatus(Target target)
        {
            if (target == Target.Text)
            {
                radioText.Checked = true;
                tbIncludeTypes.Enabled = true;
                tbExcludeTypes.Enabled = true;
                panel1a.Enabled = true;
                label1.Enabled = true;
                label2.Enabled = true;
                label3.Enabled = true;
                label4.Enabled = true;
                label5.Enabled = true;
                comNumTasks.Enabled = true;
            }
            else
            {
                if (target == Target.Excel)
                    radioExcel.Checked = true;
                else
                    radioWord.Checked = true;
                panel1a.Enabled = false;
                label1.Enabled = false;
                label2.Enabled = false;
                label3.Enabled = false;
                label4.Enabled = false;
                label5.Enabled = false;
                comNumTasks.Enabled = false;
            }
        }
    }
}
