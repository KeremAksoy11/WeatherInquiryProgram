using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Text;
using wordeaktar = Microsoft.Office.Interop.Word;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Vbe.Interop;

namespace SeleniumSecondProject
{
    class Program
    {
      
        static void Main(string[] args)
        {
            string sor, bugun;
            Console.WriteLine("************* BU PROGRAM HAVA DURUMU SORGULAMA PROGRAMIDIR *************");
            Console.Write("\nLufen Sehir Ismi Giriniz: ");
            sor = Console.ReadLine();
            IWebDriver driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://www.google.com");
            Task.Delay(2000);
            IWebElement element = driver.FindElement(By.Name("q"));
            element.SendKeys(sor + " Hava Durumu");
            element.Submit();
            Task.Delay(1000);
            Console.Write(sor + " Hava Durumu");
            string santigrat = driver.FindElement(By.Id("wob_tm")).Text.Trim().ToString();
            string state = driver.FindElement(By.Id("wob_dc")).Text.Trim().ToString();
            string yagıs = driver.FindElement(By.Id("wob_pp")).Text.Trim().ToString();
            string nem = driver.FindElement(By.Id("wob_hm")).Text.Trim().ToString();
            string ruzgar = driver.FindElement(By.Id("wob_ws")).Text.Trim().ToString();
            DateTime tarih = DateTime.Now;
            bugun = (tarih.ToString("f"));
            string bilgi = (bugun + "");
            string satirlar;
            satirlar  = ( "********************************\n" + "Gün: " + bilgi + "\nAranan Sehir: " + sor + "\nKaç Derece: " + santigrat + " Santigrat Derece" + "\nHava Durumu: " + state + "\nYağış Oranı: " + yagıs + "\nNem Oranı: " + nem + "\nRüzgar Şiddeti: " + ruzgar + "\n********************************" );
            using (System.IO.StreamWriter dosya = new System.IO.StreamWriter(@"C:\Users\kerem\OneDrive\Belgeler\visual studio 2015\Projects\SeleniumSecondProject\SeleniumSecondProject\Gecmis.txt", true))
            dosya.WriteLine(satirlar);
            string dosya1 = @"C:\Users\kerem\OneDrive\Belgeler\visual studio 2015\Projects\SeleniumSecondProject\SeleniumSecondProject\Gecmis.txt";
            driver.Quit();
            Process.Start(@"C:\Users\kerem\OneDrive\Belgeler\visual studio 2015\Projects\SeleniumSecondProject\SeleniumSecondProject\Gecmis.txt");

           
        }
       }
 }

