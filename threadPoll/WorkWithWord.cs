using System;
using System.Threading;
using Microsoft.Office.Interop.Word;

namespace threadPoll
{
    class WorkWithWord
    {
        private Application app;
        private Document document;
        private bool IsApplicaionClosed = false;
        private object path;

        //
        Test test = new Test();

        public WorkWithWord(string path = null)
        {
            this.path = path ?? @"D:\Doc3.docx";
        }

        public Document OpenNewDoc()
        {
            if (app == null || IsApplicaionClosed)
                app = new Application();
            return app.Documents.Add();
        }

        public void WriteToWord(Object stateInfo)
        {
            try
            {
                document = OpenNewDoc();
                Range range = document.Range();
                var str = string.Join(", ", test.hello());
                range.Text = $"{str}";


                document.SaveAs(path);
                document.Close();
                CloseWord();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void CloseWord()
        {
            app.Quit();
            IsApplicaionClosed = true;
        }
    }
}
