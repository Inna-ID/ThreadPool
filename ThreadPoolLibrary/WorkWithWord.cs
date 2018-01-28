using System;
using System.Text;
using System.Threading;
using Microsoft.Office.Interop.Word;

namespace ThreadPoolLibrary
{
    public class WorkWithWord
    {
        private Application app;
        private Document document;
        private bool IsApplicaionClosed = false;
        private object path;

        //с помощью конструктора устанавливаем путь до файла
        public WorkWithWord(string path)
        {
            this.path = path;
        }

        public Document OpenNewDoc()
        {
            if (app == null || IsApplicaionClosed)
                app = new Application(); //открываем новое приложение
            return app.Documents.Add(); //открываем новый докумнт
        }

        public void WriteToWord(string text)
        {
            try
            {
                document = OpenNewDoc();
                //объект диапазона текста
                Range range = document.Range();
                var time = WinAPI.СhechSystemTime();
                text += $"\n\n Edit Time: {time}";
                range.Text = text; //пишем инфо в ворд

                //сохраняем результат
                document.SaveAs(path);
                document.Close();
                CloseWord();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        //закрываем приложение
        public void CloseWord()
        {
            app.Quit();
            IsApplicaionClosed = true;
        }
    }
}
