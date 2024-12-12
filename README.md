# kt-fkt-f32

using Word = Microsoft.Office.Interop.Word;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private void Repwo(string subst, string text, Word.Document word)
        {
            var range = word.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: subst, ReplaceWith: text);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //определение переменной для использования Word
            var WordApp = new Word.Application();
            WordApp.Visible = false;
            // делаем диалог выбора папки, в которую будет сохранятся билет
            var Worddoc = WordApp.Documents.Open(Environment.CurrentDirectory +
            @"\билет.docx");
            Repwo("{Город}", DateTime.Now.ToString(), Worddoc);
            Worddoc.SaveAs2(Environment.CurrentDirectory + @"\билет новый.docx");
            MessageBox.Show("Билет сохранен!");
        }
    }
}
