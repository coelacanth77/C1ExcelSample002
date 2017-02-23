using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using C1.WPF.Excel;
using System.IO;
using System.Drawing;

namespace C1ExcelSample002
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            C1XLBook book = new C1XLBook();

            // ファイルを読み込む
            book.Load(@"C:\Users\macni\Documents\mybook.xls");

            // シートを編集する
            book.Sheets[0][0, 0].Value = 5;

            // データをFlexGrid用に加工する
            List<int> list = new List<int>();

            list.Add(Convert.ToInt32(book.Sheets[0][0, 0].Value));
            list.Add(Convert.ToInt32(book.Sheets[0][0, 1].Value));
            list.Add(Convert.ToInt32(book.Sheets[0][0, 2].Value));

            // FlexGridで表示する
            FlexGrid.ItemsSource = list;

            // 別名で保存する
            book.Save(@"C:\Users\macni\Documents\mybook2.xls");

        }

        private void FlexGrid_SelectionChanged(object sender, C1.WPF.FlexGrid.CellRangeEventArgs e)
        {
            if (!Double.IsNaN(FlexGrid.Width))
            {
                var size = new System.Windows.Size(FlexGrid.Width, FlexGrid.Height);
                FlexGrid.Measure(size);
                FlexGrid.Arrange(new Rect(size));

                // RenderTargetBitmapでFlexGridをBitMapに変換する
                var renderBitmap = new RenderTargetBitmap((int)size.Width, (int)size.Height, 96.0d, 96.0d, PixelFormats.Pbgra32);
                renderBitmap.Render(FlexGrid);

                // Imageコントロールに表示する
                // C1PDFやC1Wordを利用して保存することも可能
                image.Source = BitmapFrame.Create(renderBitmap);
            }
        }
    }
}
