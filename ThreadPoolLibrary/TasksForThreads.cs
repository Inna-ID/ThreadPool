using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Text;

namespace ThreadPoolLibrary
{
    public class TasksForThreads
    {
        private Thread _thread;
        int workerThreads;
        int portThreads;
 

        public void FindPrimeNumbers(Object stateInfo)
        {
            WorkWithWord www = new WorkWithWord(@"D:\Prime_numbers.docx");
            List<int> arrayOfPrimes = new List<int>();
            int min = 1;
            int max = 500;
            int j;

            for (int i = min; i <= max; i++)
            {
                for (j = 2; j <= i; j++)
                {
                    if (i % j == 0) break;
                }
                if (j == i)
                {
                    arrayOfPrimes.Add(i);
                }
            }

            foreach (int item in arrayOfPrimes)
            {
                Thread.Sleep(200);
                Console.WriteLine($"{new string(' ', 70)}{item}");
            }

            // пишем список протых чисел в ворд, разделяя элементы ,
            www.WriteToWord($"Prime numbers\n{string.Join(", ", arrayOfPrimes)}");
        }


        public void Сalculation(Object stateInfo)
        {
            WorkWithWord www = new WorkWithWord(@"D:\Сalculation.docx");
            string strArray = "";
            string str = "";
            var result = 0;
            for (int i = 10; i < 70; i++)
            {
                result = (i + 5) * i / 2;
                Thread.Sleep(1200);
                str = $"{result}\n";
                strArray += str;
                Console.WriteLine($"{new string(' ', 30)} result = {result}");
            }
            //записываем результат в ворд
            www.WriteToWord($"Result (n + 5) * n / 2\n{strArray}");
        }

        public void ToSecondDegree(Object stateInfo)
        {
            WorkWithWord www = new WorkWithWord(@"D:\Second_degree.docx");
            string strArray = "";
            string str = "";
            for (int i = 0; i < 60; i++)
            {
                Thread.Sleep(700);
                str = $"i^2 = {i * i}\n";
                strArray += str;
                Console.WriteLine($"{new string(' ', 50)} {str}");
            }
            //записываем результат в ворд
            www.WriteToWord($"Table of second degree\n{strArray}");
        }


        public void ReadFromFile(Object stateInfo)
        {
            //записываем содержимое файла в массив. 
            //здесь один элемент, содержащий в себе все содержимое файла
            string[] file = File.ReadAllLines(@"D:\cities.txt");            
            //объединяем набор данных переводим в строку
            var str = string.Concat(file).ToString();
            //разбиваем строку на массив строк
            string[] cities = str.Split();
            //сортируем по алфавиту
            Array.Sort(cities);
            
            Console.WriteLine("Cities of the Minsk region: ");

            // выводим содержимое массива в консоль
            foreach (string city in cities)
            {
                Thread.Sleep(1000);
                Console.WriteLine(city);
            }
        }
    }
}
