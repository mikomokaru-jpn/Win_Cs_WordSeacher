using System;
using System.Data;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;

namespace WordSeacher
{

    public delegate void menuOperation(string path);

    class UADataGridView
    {
        //結果リスト
        public DataTable dataTable = new DataTable();
        //設定オブジェクト
        SearchSettings serchSet = SearchSettings.sharedInstance();
        //データグリッドビュー・結果一覧
        public DataGridView resultView = new DataGridView()
        {
            Location = new Point(10, 80),
            Size = new Size(380, 220),
            RowHeadersVisible = false,
            AllowUserToAddRows = false,
            ReadOnly = true,
            Anchor = (AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right),
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            AllowUserToResizeRows = false,
            AllowUserToDeleteRows = false,
            TabStop = false,
        };
        //コンテキストメニュー
        ContextMenuStrip menu = new ContextMenuStrip();
        //デリゲートの宣言
        public menuOperation funcMenu1 = null;
        public menuOperation funcMenu2 = null;
        public menuOperation funcMenu3 = null;

        //オブジェクト初期化
        public void setUp()
        {
            //列幅の変更・設定値の保存
            resultView.ColumnWidthChanged += resultView_ColumnWidthChanged;
            //列ヘッダのクリック・ソート
            resultView.ColumnHeaderMouseClick += resultView_ColumnHeaderMouseClick;
            //データテーブル
            dataTable.Columns.Add("folder", Type.GetType("System.String"));
            dataTable.Columns.Add("file", Type.GetType("System.String"));
            dataTable.Columns.Add("sheet", Type.GetType("System.String"));
            dataTable.Columns.Add("count", Type.GetType("System.Int32"));
            dataTable.Columns.Add("type", Type.GetType("System.Int32"));
            resultView.DataSource = dataTable;
            for (int i = 0; i < resultView.Columns.Count; i++)
                resultView.Columns[i].Visible = false;

            //コンテキストメニュー
            menu.Items.Add(new ToolStripMenuItem() { Name = "item1", Text = "ファイルを開く" });
            menu.Items.Add(new ToolStripMenuItem() { Name = "item2", Text = "Codeでファイルを開く" });
            menu.Items.Add(new ToolStripMenuItem() { Name = "item3", Text = "パスをコピーする" });
            menu.Items["item1"].Click += item1_Click;
            menu.Items["item2"].Click += item2_Click;
            menu.Items["item3"].Click += item3_Click;
            //セルの右クリック・ファイルを開く
            resultView.CellMouseClick += resultView_CellMouseClick;

        }
        //結果一覧表示
        public void display(List<Result> resultList)
        {
            resultView.Columns["folder"].Visible = true;
            resultView.Columns["folder"].HeaderText = "フォルダ";
            resultView.Columns["folder"].Width = serchSet.colFolderWidth;
            resultView.Columns["folder"].SortMode = DataGridViewColumnSortMode.Programmatic;

            resultView.Columns["file"].Visible = true;
            resultView.Columns["file"].HeaderText = "ファイル";
            resultView.Columns["file"].Width = serchSet.colFileWidth;
            resultView.Columns["file"].SortMode = DataGridViewColumnSortMode.Programmatic;

            resultView.Columns["sheet"].Visible = true;
            resultView.Columns["sheet"].HeaderText = "シート";
            resultView.Columns["sheet"].Width = serchSet.colSheetWidth;
            resultView.Columns["sheet"].SortMode = DataGridViewColumnSortMode.Programmatic;

            resultView.Columns["count"].Visible = true;
            resultView.Columns["count"].HeaderText = "ヒット数";
            resultView.Columns["count"].Width = serchSet.colCountWidth;
            resultView.Columns["count"].SortMode = DataGridViewColumnSortMode.Programmatic;
            
            if (serchSet.target == Target.Excel)
                resultView.Columns["sheet"].Visible = true;
            else
                resultView.Columns["sheet"].Visible = false;
          
            foreach (Result result in resultList)
            {
                var row = dataTable.NewRow();
                row.SetField("folder", result.folder);
                row.SetField("file", result.file);
                row.SetField("sheet", result.sheet);
                row.SetField("count", result.count);
                row.SetField("type", result.type);
                dataTable.Rows.Add(row);
            }
            dataTable.DefaultView.Sort = "count DESC, folder DESC, file DESC, sheet DESC";
            resultView.Rows[0].Selected = true; //先頭行を選択状態にする
        }
        //列の幅が変わった
        public void resultView_ColumnWidthChanged(object srnder, DataGridViewColumnEventArgs e)
        {
            switch (e.Column.Name)
            {
                case "folder":
                    serchSet.colFolderWidth = resultView.Columns[e.Column.Name].Width;
                    break;
                case "file":
                    serchSet.colFileWidth = resultView.Columns[e.Column.Name].Width;
                    break;
                case "sheet":
                    serchSet.colSheetWidth = resultView.Columns[e.Column.Name].Width;
                    break;
                case "count":
                    serchSet.colCountWidth = resultView.Columns[e.Column.Name].Width;
                    break;
            }
        }

        //選択行を右クリックする
        void resultView_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (e.ColumnIndex >= 0 && e.RowIndex >= 0)
                {
                    if (resultView.SelectedRows[0].Index == e.RowIndex)
                    {
                        resultView.Rows[e.RowIndex].Selected = true;
                        Point p = Control.MousePosition;
                        var flg = (Target)resultView[4, resultView.SelectedRows[0].Index].Value == Target.Text;
                        menu.Items["item2"].Visible = flg;  //Codeで開くメニュー
                        menu.Show(p);
                    }
                }
            }
        }
        //ファイルを開く
        void item1_Click(object sender, EventArgs e)
        {
            var row = resultView.SelectedRows[0].Index;
            var path = "\"" + (string)resultView[0, row].Value + @"\" + (string)resultView[1, row].Value + "\"";
            this.funcMenu1(path);   //デリゲートの呼び出し
        }
        //Codeでファイルを開く
        void item2_Click(object sender, EventArgs e)
        {
            var row = resultView.SelectedRows[0].Index;
            var path = "\"" + (string)resultView[0, row].Value + @"\" + (string)resultView[1, row].Value + "\"";
            this.funcMenu2(path);   //デリゲートの呼び出し
        }
        //パスをコピーする
        void item3_Click(object sender, EventArgs e)
        {
            var row = resultView.SelectedRows[0].Index;
            var path = "\"" + (string)resultView[0, row].Value + @"\" + (string)resultView[1, row].Value + "\"";
            this.funcMenu3(path);   //デリゲートの呼び出し
        }

        //ソート
        private void resultView_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            var column = resultView.Columns[e.ColumnIndex];
            var direct = column.HeaderCell.SortGlyphDirection;
            var script = "";
            switch (column.Name)
            {
                case "folder":
                    if (direct == SortOrder.None || direct == SortOrder.Descending)
                    { script = "folder ASC, file ASC, sheet ASC"; }
                    else
                    { script = "folder DESC, file DESC, sheet DESC"; }
                    break;
                case "file":
                    if (direct == SortOrder.None || direct == SortOrder.Descending)
                    { script = "file ASC, folder ASC, sheet ASC"; }
                    else
                    { script = "file DESC, folder DESC, sheet DESC"; }
                    break;
                case "count":
                    if (direct == SortOrder.None || direct == SortOrder.Descending)
                    { script = "count ASC, folder ASC, file ASC, sheet ASC"; }
                    else
                    { script = "count DESC, folder DESC, file DESC, sheet DESC"; }
                    break;
                default:
                    break;
            }
            dataTable.DefaultView.Sort = script;
        }
    }
}

