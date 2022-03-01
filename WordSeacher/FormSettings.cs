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
            Size = new Size(340, 230),
            BorderStyle = BorderStyle.FixedSingle,
        };
        //対象ファイルタイプ
        Label label1 = new Label
        {
            Location = new Point(10, 20),
            Size = new Size(70, 20),
            Text = "対象拡張子",
        };
        TextBox tbIncludeTypes = new TextBox()
        {
            Location = new Point(80, 15),
            Size = new Size(250, 25),
        };
        Label label3 = new Label
        {
            Location = new Point(80, 40),
            Size = new Size(300, 25),
            Text = "指定がない場合すべてのファイルタイプを対象とする",
        };
        //対象外ファイルタイプ
        Label label2 = new Label
        {
            Location = new Point(10, 70),
            Size = new Size(70, 20),
            Text = "除外拡張子",
        };
        TextBox tbExcludeTypes = new TextBox()
        {
            Location = new Point(80, 65),
            Size = new Size(250, 25),
        };
        Label label4 = new Label
        {
            Location = new Point(80, 90),
            Size = new Size(300, 25),
            Text = "対象拡張子の指定がない場合のみ有効",
        };
        //パネル1a
        Panel panel1a = new Panel()
        {
            Location = new Point(0, 110),
            Size = new Size(310, 30),
            //BorderStyle = BorderStyle.FixedSingle
        };
        Label label6 = new Label
        {
            Location = new Point(10, 0),
            Size = new Size(70, 30),
            TextAlign = ContentAlignment.MiddleLeft,
            Text = "文字コード",
        };
        //文字コード
        RadioButton radioUTF8 = new RadioButton()
        {
            Location = new Point(80, 2),
            TextAlign = ContentAlignment.MiddleLeft,
            Size = new Size(70, 25),
            Text = "UTF-8"
        };
        RadioButton radioSJIS = new RadioButton()
        {
            Location = new Point(150, 2),
            TextAlign = ContentAlignment.MiddleLeft,
            Size = new Size(70, 25),
            Text = "Shift-JIS"
        };

        //大文字小文字区別する
        CheckBox checkCase = new CheckBox()
        {
            Location = new Point(10, 150),
            Size = new Size(150, 20),
            TextAlign = ContentAlignment.MiddleLeft,
            Text = "大文字小文字区別する"
        };

        //並列タスク数
        Label label5 = new Label
        {
            Location = new Point(10, 180),
            Size = new Size(70, 20),
            Text = "並列タスク数",
            TextAlign = ContentAlignment.MiddleLeft,
        };
        ComboBox comNumTasks = new ComboBox()
        {
            Location = new Point(80, 180),
            Size = new Size(50, 25),
        };

       
        //-------------------------------------------------------------
        // パネル2
        //-------------------------------------------------------------
        Panel panel2 = new Panel()
        {
            Location = new Point(10, 250),
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
            Location = new Point(170, 30),
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

        //コンストラクタ
        public FormSettings()
        {
            this.ClientSize = new Size(360, 330);
            this.Text = "設定";     // タイトルを設定
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.Load += form_Load;
            this.FormClosing += form_FormClosing;
            this.KeyPress += Form2_KeyPress;
            this.KeyPreview = true;
            //コントロールの追加
            this.Controls.Add(panel1);
            panel1.Controls.AddRange(new Control[]
            {
                label1, label2, label3, label4, label5, tbIncludeTypes, tbExcludeTypes,  comNumTasks, panel1a, checkCase 
            });
            panel1a.Controls.AddRange(new Control[]
            {
               label6, radioUTF8, radioSJIS
            });
            comNumTasks.Items.AddRange(new string[] { "1", "10" });
            radioUTF8.Checked = true;
            radioAND.Checked = true;

            this.Controls.Add(panel2);
            panel2.Controls.AddRange(new Control[]
            {
                panel2Label, radioAND, radioOR, radioAsIs
            });
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
        //フォームの終了
        void form_FormClosing(object sender, FormClosingEventArgs e)
        {
            //値の保存
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
