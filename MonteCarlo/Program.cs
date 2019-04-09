using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using ConsoleApp1;
namespace MonteCarlo_Method
{
     class Program
    {

         
                        
        //double aPiGlobal = 0;
        static void Main(string[] args)
        {
            int pointNumber = 10000000;
            double radius=1;
            int threadCount = 4;
            double aPi=0;
            int result;
            Class1 clas= new Class1();
            
            aPi = clas.LeibnicPI(pointNumber);

            //result = clas.Plinq(pointNumber, radius, aPi);
            //clas.WriteInConsole(pointNumber, result, aPi,"PLINQ");

            result = clas.TPL(pointNumber, radius, aPi);
            clas.WriteInConsole(pointNumber, result, aPi, "TPL");

            //result = clas.Threads(pointNumber, radius, aPi, threadCount);
            //clas.WriteInConsole(pointNumber, result, aPi, "THREADS");

            result = clas.SequentialWork(pointNumber, radius, aPi);
            clas.WriteInConsole(pointNumber, result, aPi, "Последовательная работа");

            Console.ReadLine();
        }
         
         
    }
}