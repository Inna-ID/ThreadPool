using System;
using System.Text;
using System.Runtime.InteropServices; //для использования Win API

namespace ThreadPoolLibrary
{
    [StructLayout(LayoutKind.Sequential)]
    public class SystemTime
    {
        public ushort year;
        public ushort month;
        public ushort weekday;
        public ushort day;
        public ushort hour;
        public ushort minute;
        public ushort second;
        public ushort millisecond;
    }

    public class LibWrap
    {
        // импортируем dll для работы с временем системы
        [DllImport("Kernel32.dll")]
        public static extern void GetSystemTime([In, Out] SystemTime st); //получаем метод GetSystemTime из Kernel32.dll 
    }

    public static class WinAPI
    {
        static SystemTime st = new SystemTime();
        public static string СhechSystemTime()
        {
            LibWrap.GetSystemTime(st);
            //записываем дату и время
            var date = $"{st.day}.{st.month}.{st.year} | {st.hour + 3} : {st.minute}";
            return date;
        }
    }
}
