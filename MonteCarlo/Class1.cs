using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
namespace ConsoleApp1
{
    class Class1
    {
        Stopwatch stopWatch = new Stopwatch();
        object locker = new object();
        List<TimeSpan> Time = new List<TimeSpan>();
        List<double> accuracy = new List<double>();
        List<EventWaitHandle> handles = new List<EventWaitHandle>(); 
        public int inc = 0, circleglobal = 0, incacc = 0, xglob = 0;
        public double LeibnicPI( int pointNumber)
        {
            double aPi = 0;
            for (int i = 0; i < pointNumber; i++)
            {
                aPi += Math.Pow(-1, i) / (2 * i + 1);
            }
            aPi = 4 * aPi;
            Console.WriteLine("Pi с помощью ряда Лейбница: {0}", aPi + "\n");
            return aPi;
            
        }
        public void GlobalParamToZero()
        {
            inc = 0;
            circleglobal = 0;
            incacc = 0;
            xglob = 0;
            Time.Clear();
            accuracy.Clear();
            handles.Clear();
        }
        public int Plinq(int pointNumber, double radius, double aPi)
        {
            stopWatch.Start();
            GlobalParamToZero();
            var result = ParallelEnumerable
               .Range(0, pointNumber)
               .AsParallel()
               .WithDegreeOfParallelism(1)
               .Select(p => IsCircle(radius, aPi))
               .Count(b => b);
            return result;
        }
        public int TPL(int pointNumber, double radius, double aPi)
        {
            stopWatch.Start();
            GlobalParamToZero();
            Parallel.For(0, pointNumber, new ParallelOptions { MaxDegreeOfParallelism = 1 },i =>
            {
                IsCircle(radius, aPi);
            });
            return circleglobal;
        }
		
        public int Threads(int pointNumber, double radius, double aPi, int threadCount)
        {
            stopWatch.Start();
            GlobalParamToZero();
            var shag = pointNumber / threadCount;
            var raznica = pointNumber - shag * threadCount;
            for (int i = 0; i < threadCount; i++)
            {
                EventWaitHandle handle = new AutoResetEvent(false);
                if (i + 1 != threadCount) new Thread(delegate() { MonteCarloMethod(pointNumber, handle, radius, shag, i + 1, 0,aPi); }).Start();
                else new Thread(delegate() { MonteCarloMethod(pointNumber, handle, radius, shag, i + 1, raznica,aPi); }).Start();
                handles.Add(handle);
            }
            WaitHandle.WaitAll(handles.ToArray());
            return circleglobal;
        }
        public int SequentialWork(int pointNumber, double radius, double aPi)
        {
            stopWatch.Start();
            GlobalParamToZero();
            for (int i = 0; i < pointNumber; i++)
            {
                IsCircle(radius, aPi);
            
            }
            return circleglobal;
        }
        void MonteCarloMethod(int pointNumber, object handle,double radius,int shag, int i,int raznica,double aPi)
        {
            AutoResetEvent wh = (AutoResetEvent)handle;
            for (int y = (i-1)*shag; y < shag*i+ raznica; y++)
            {
                IsCircle(radius, aPi); 
            }
            wh.Set();
        }
        public void WriteInConsole(int pointNumber, int result, double aPi,string str)
        {
            double Pi;
            Console.WriteLine(str+":");
            Console.WriteLine("Количество попавших: {0}", result);
            Pi =  (4.0 * (double)result) / (double)pointNumber;
            Console.WriteLine("Pi: {0}", Pi); 
            Console.WriteLine("Точность расчета: {0}", Pi- aPi);
            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
            ts.Hours, ts.Minutes, ts.Seconds,
            ts.Milliseconds / 10);
            Console.WriteLine("RunTime " + elapsedTime+"\n");
            stopWatch.Restart();
            stopWatch.Stop();
            //WriteInExcel();
        }
        public void WriteInExcel()
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workBook;
            Excel.Worksheet workSheet;
            workBook = excelApp.Workbooks.Add();
            workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
            int y = 1;
            foreach (double i in accuracy)//или (TimeSpan i in Time) или  (double i in accuracy)
            {
                workSheet.Cells[y, 1] = Convert.ToString(i);
                y += 1;
            }
            excelApp.Visible = true;
            excelApp.UserControl = true;
        }
        bool IsCircle(double radius, double aPi)
        {
            int newVal = Interlocked.Increment(ref inc);
            Random rnd = new Random((int)DateTime.Now.Ticks + newVal);
            double x = rnd.NextDouble();
            double y = rnd.NextDouble();
            //ПРОВЕРКА ДЛЯ ВРЕМЕНИ
            TimeinList(newVal,x, y, radius);
            //ПРОВЕРКА ТОЧНОСТИ
            //AccuracyinList(newVal, x, y, radius, aPi);
            return ((x * x + y * y) <= radius * radius);
        }
        public void TimeinList(int newVal, double x, double y, double radius)
        {
            //ПРОВЕРКА ДЛЯ ВРЕМЕНИ
            if (newVal == 1000000) 
            {
                Time.Add(stopWatch.Elapsed);
                //
                Interlocked.Exchange(ref inc, 0);
                
            }
            if ((x * x + y * y) <= radius * radius) Interlocked.Increment(ref circleglobal);
        }
        public void AccuracyinList(int newVal, double x, double y, double radius, double aPi)
        {
            lock (locker)
            {

                if (incacc == 1000000)
                {
                    xglob += 1;
                    accuracy.Add(((4.0 * (double)circleglobal) / ((double)1000000 * xglob)) - aPi);
                    incacc = 0;
                    //Console.WriteLine(circleglobal);
                }
                incacc += 1;
                if ((x * x + y * y) <= radius * radius) circleglobal += 1;
            }
        }
    }
}
