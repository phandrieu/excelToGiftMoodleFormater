/*
First public verison published on GitHub by Paul-Henri Andrieu (@phandrieu) on 04.29.2020
*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics.SymbolStore;
using System.Runtime.InteropServices;

namespace myFormatter
{

    class Program
    {
        void print(string text) {
            Console.WriteLine(text);
        }

        static void Main(string[] args)
        {
            Question[] questions = new Question[30];

            string filePath;
            Console.WriteLine("Veuillez préciser un chemin d'accès à un classeur Excel S.V.P");
            filePath = Console.ReadLine();
            filePath = filePath.Replace("\"", "");
            Excel.Application XlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = XlApp.Workbooks.Open(@filePath, CorruptLoad: true, ReadOnly: true);
            Excel._Worksheet xlWorksheet = xlWorkBook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            for (int rowNumber = 2; rowNumber <= xlRange.Rows.Count; rowNumber++)
            {
                bool monChoixMultiple = false ;
                string maQuestion, maReponseA, mesPtsReponseA, maReponseB, mesPtsReponseB, maReponseC, mesPtsReponseC, maReponseD, mesPtsReponseD, maReponseE, mesPtsReponseE;
                maQuestion = maReponseA = maReponseB = maReponseC = maReponseD = maReponseE = mesPtsReponseA = mesPtsReponseB= mesPtsReponseC=mesPtsReponseD=mesPtsReponseE = "";
                
                if (xlRange.Cells[rowNumber, 1] != null && xlRange.Cells[rowNumber, 1].Value2 != null)
                {
                    maQuestion = xlRange.Cells[rowNumber, 1].Value2.ToString();
                }
                if(xlRange.Cells[rowNumber, 2] != null && xlRange.Cells[rowNumber, 2].Value2 != null)
                {
                    if (xlRange.Cells[rowNumber, 2].Value2.ToString() == "U" || xlRange.Cells[rowNumber, 2].Value2.ToString() == "u")
                    {
                        monChoixMultiple = false;
                    }
                    else if (xlRange.Cells[rowNumber, 2].Value2.ToString() == "M" || xlRange.Cells[rowNumber, 2].Value2.ToString() == "m") 
                    {
                        monChoixMultiple = true;
                    }
                }
                if (xlRange.Cells[rowNumber, 3] != null && xlRange.Cells[rowNumber, 3].Value2 != null)
                {
                    maReponseA = xlRange.Cells[rowNumber, 3].Value2.ToString();
                }
                if (xlRange.Cells[rowNumber, 4] != null && xlRange.Cells[rowNumber, 4].Value2 != null)
                {
                    mesPtsReponseA = xlRange.Cells[rowNumber, 4].Value2.ToString();
                }
                if (xlRange.Cells[rowNumber, 5] != null && xlRange.Cells[rowNumber, 5].Value2 != null)
                {
                    maReponseB = xlRange.Cells[rowNumber, 5].Value2.ToString();
                }
                if (xlRange.Cells[rowNumber, 6] != null && xlRange.Cells[rowNumber, 6].Value2 != null)
                {
                    mesPtsReponseB = xlRange.Cells[rowNumber, 6].Value2.ToString();
                }
                if (xlRange.Cells[rowNumber, 7] != null && xlRange.Cells[rowNumber, 7].Value2 != null)
                {
                    maReponseC = xlRange.Cells[rowNumber, 7].Value2.ToString();
                }
                if (xlRange.Cells[rowNumber, 8] != null && xlRange.Cells[rowNumber, 8].Value2 != null)
                {
                    mesPtsReponseC = xlRange.Cells[rowNumber, 8].Value2.ToString();
                }
                if (xlRange.Cells[rowNumber, 9] != null && xlRange.Cells[rowNumber, 9].Value2 != null)
                {
                    maReponseD = xlRange.Cells[rowNumber, 9].Value2.ToString();
                }
                if (xlRange.Cells[rowNumber, 10] != null && xlRange.Cells[rowNumber, 10].Value2 != null)
                {
                    mesPtsReponseD = xlRange.Cells[rowNumber, 10].Value2.ToString();
                }
                if (xlRange.Cells[rowNumber, 11] != null && xlRange.Cells[rowNumber, 11].Value2 != null)
                {
                    maReponseE = xlRange.Cells[rowNumber, 11].Value2.ToString();
                }
                if (xlRange.Cells[rowNumber, 12] != null && xlRange.Cells[rowNumber, 12].Value2 != null)
                {
                    mesPtsReponseE = xlRange.Cells[rowNumber, 12].Value2.ToString();
                }


                if (mesPtsReponseA != "")
                {
                    mesPtsReponseA.Replace(',', '.');
                    double ptsRepA100 = Convert.ToDouble(mesPtsReponseA);
                    ptsRepA100 = ptsRepA100 * 100;
                    mesPtsReponseA = ptsRepA100.ToString();
                }

                if (mesPtsReponseB != "") 
                {
                    mesPtsReponseB.Replace(',', '.');
                    double ptsRepB100 = Convert.ToDouble(mesPtsReponseB);
                    ptsRepB100 = ptsRepB100 * 100;
                    mesPtsReponseB = ptsRepB100.ToString();
                }

                if (mesPtsReponseC != "") 
                {
                    mesPtsReponseC.Replace(',', '.');
                    double ptsRepC100 = Convert.ToDouble(mesPtsReponseC);
                    ptsRepC100 = ptsRepC100 * 100;
                    mesPtsReponseC = ptsRepC100.ToString();
                }

                if (mesPtsReponseD != "")
                {
                    mesPtsReponseD.Replace(',', '.');
                    double ptsRepD100 = Convert.ToDouble(mesPtsReponseD);
                    ptsRepD100 = ptsRepD100 * 100;
                    mesPtsReponseD = ptsRepD100.ToString();
                }

                if (mesPtsReponseE != "") {
                    mesPtsReponseE.Replace(',', '.');
                    double ptsRepE100 = Convert.ToDouble(mesPtsReponseE);
                    ptsRepE100 = ptsRepE100 * 100;
                    mesPtsReponseE = ptsRepE100.ToString();
                }

                



                questions[rowNumber - 2] = new Question(maQuestion, monChoixMultiple, maReponseA, mesPtsReponseA, maReponseB, mesPtsReponseB, maReponseC, mesPtsReponseC, maReponseD, mesPtsReponseD, maReponseE,mesPtsReponseE);
            }

            Console.WriteLine("Entrez le nom du document à écrire");
            string nomDuDocument = Console.ReadLine();

            foreach (Question question in questions) {
                string filePathWtEnv = @"%userprofile%\Desktop\Fichiers_Moodlee_GIFT";
                var filePath2 = Environment.ExpandEnvironmentVariables(filePathWtEnv);
                System.IO.Directory.CreateDirectory(filePath2);
                

                using (System.IO.StreamWriter file = new System.IO.StreamWriter(filePath2+@"\"+nomDuDocument, true))
                {
                    if (question.choixMultiple == true)
                    {
                        try
                        {
                            if (question.question != "")
                            {
                                file.WriteLine(question.question + "{");
                            }
                            if (question.reponseA != "")
                            {
                                file.WriteLine("\t~%" + question.ptsReponseA + "%" + question.reponseA);
                            }
                            if (question.reponseB != "")
                            {
                                file.WriteLine("\t~%" + question.ptsReponseB + "%" + question.reponseB);
                            }
                            if (question.reponseC != "")
                            {
                                file.WriteLine("\t~%" + question.ptsReponseC + "%" + question.reponseC);
                            }
                            if (question.reponseD != "")
                            {
                                file.WriteLine("\t~%" + question.ptsReponseD + "%" + question.reponseD);
                            }
                            if (question.reponseE != "")
                            {
                                file.WriteLine("\t~%" + question.ptsReponseE + "%" + question.reponseE);
                            }
                            file.WriteLine("}\n\r");
                            if (question.question != "")
                            {
                                Console.WriteLine(question.question + "{");
                            }
                            if (question.reponseA != "")
                            {
                                Console.WriteLine("\t~%" + question.ptsReponseA + "%" + question.reponseA);
                            }
                            if (question.reponseB != "")
                            {
                                Console.WriteLine("\t~%" + question.ptsReponseB + "%" + question.reponseB);
                            }
                            if (question.reponseC != "")
                            {
                                Console.WriteLine("\t~%" + question.ptsReponseC + "%" + question.reponseC);
                            }
                            if (question.reponseD != "")
                            {
                                Console.WriteLine("\t~%" + question.ptsReponseD + "%" + question.reponseD);
                            }
                            if (question.reponseE != "")
                            {
                                Console.WriteLine("\t~%" + question.ptsReponseE + "%" + question.reponseE);
                            }
                            Console.WriteLine("}\n\r");

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Erreur : " + ex.ToString());
                            Marshal.ReleaseComObject(xlRange);
                            Marshal.ReleaseComObject(xlWorksheet);
                            xlWorkBook.Close();
                            XlApp.Quit();
                        }
                    }
                    else if (question.choixMultiple == false) 
                    {
                        if (Convert.ToInt32(question.ptsReponseA) == 100)
                        {
                            try
                            {
                                if (question.question != "")
                                {
                                    file.WriteLine(question.question + "{");
                                }
                                if (question.reponseA != "")
                                {
                                    file.WriteLine("\t=" + question.reponseA);
                                }
                                if (question.reponseB != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseB + "%" + question.reponseB);
                                }
                                if (question.reponseC != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseC + "%" + question.reponseC);
                                }
                                if (question.reponseD != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseD + "%" + question.reponseD);
                                }
                                if (question.reponseE != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseE + "%" + question.reponseE);
                                }
                                file.WriteLine("}\n\r");
                                if (question.question != "")
                                {
                                    Console.WriteLine(question.question + "{");
                                }
                                if (question.reponseA != "")
                                {
                                    Console.WriteLine("\t=" + question.reponseA);
                                }
                                if (question.reponseB != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseB + "%" + question.reponseB);
                                }
                                if (question.reponseC != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseC + "%" + question.reponseC);
                                }
                                if (question.reponseD != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseD + "%" + question.reponseD);
                                }
                                if (question.reponseE != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseE + "%" + question.reponseE);
                                }
                                Console.WriteLine("}\n\r");

                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Erreur : " + ex.ToString());
                                Marshal.ReleaseComObject(xlRange);
                                Marshal.ReleaseComObject(xlWorksheet);
                                xlWorkBook.Close();
                                XlApp.Quit();
                            }
                        }
                        else if (Convert.ToInt32(question.ptsReponseB) == 100)
                        {
                            try
                            {
                                if (question.question != "")
                                {
                                    file.WriteLine(question.question + "{");
                                }
                                if (question.reponseA != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseA + "%" + question.reponseA);
                                }
                                if (question.reponseB != "")
                                {
                                    file.WriteLine("\t=" + question.reponseB);
                                }
                                if (question.reponseC != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseC + "%" + question.reponseC);
                                }
                                if (question.reponseD != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseD + "%" + question.reponseD);
                                }
                                if (question.reponseE != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseE + "%" + question.reponseE);
                                }
                                file.WriteLine("}\n\r");
                                if (question.question != "")
                                {
                                    Console.WriteLine(question.question + "{");
                                }
                                if (question.reponseA != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseA + "%" + question.reponseA);
                                }
                                if (question.reponseB != "")
                                {
                                    Console.WriteLine("\t=" + question.reponseB);
                                }
                                if (question.reponseC != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseC + "%" + question.reponseC);
                                }
                                if (question.reponseD != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseD + "%" + question.reponseD);
                                }
                                if (question.reponseE != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseE + "%" + question.reponseE);
                                }
                                Console.WriteLine("}\n\r");

                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Erreur : " + ex.ToString());
                                Marshal.ReleaseComObject(xlRange);
                                Marshal.ReleaseComObject(xlWorksheet);
                                xlWorkBook.Close();
                                XlApp.Quit();
                            }
                        }
                        else if (Convert.ToInt32(question.reponseC) == 100) 
                        {
                            try
                            {
                                if (question.question != "")
                                {
                                    file.WriteLine(question.question + "{");
                                }
                                if (question.reponseA != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseA + "%" + question.reponseA);
                                }
                                if (question.reponseB != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseB + "%" + question.reponseB);
                                }
                                if (question.reponseC != "")
                                {
                                    file.WriteLine("\t=" + question.reponseC);
                                }
                                if (question.reponseD != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseD + "%" + question.reponseD);
                                }
                                if (question.reponseE != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseE + "%" + question.reponseE);
                                }
                                file.WriteLine("}\n\r");
                                if (question.question != "")
                                {
                                    Console.WriteLine(question.question + "{");
                                }
                                if (question.reponseA != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseA + "%" + question.reponseA);
                                }
                                if (question.reponseB != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseB + "%" + question.reponseB);
                                }
                                if (question.reponseC != "")
                                {
                                    Console.WriteLine("\t=" + question.reponseC);
                                }
                                if (question.reponseD != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseD + "%" + question.reponseD);
                                }
                                if (question.reponseE != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseE + "%" + question.reponseE);
                                }
                                Console.WriteLine("}\n\r");

                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Erreur : " + ex.ToString());
                                Marshal.ReleaseComObject(xlRange);
                                Marshal.ReleaseComObject(xlWorksheet);
                                xlWorkBook.Close();
                                XlApp.Quit();
                            }
                        }
                        else if(Convert.ToInt32(question.ptsReponseD) == 100)
                        {
                            try
                            {
                                if (question.question != "")
                                {
                                    file.WriteLine(question.question + "{");
                                }
                                if (question.reponseA != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseA + "%" + question.reponseA);
                                }
                                if (question.reponseB != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseB + "%" + question.reponseB);
                                }
                                if (question.reponseC != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseC + "%" + question.reponseC);
                                }
                                if (question.reponseD != "")
                                {
                                    file.WriteLine("\t=" + question.reponseD);
                                }
                                if (question.reponseE != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseE + "%" + question.reponseE);
                                }
                                file.WriteLine("}\n\r");
                                if (question.question != "")
                                {
                                    Console.WriteLine(question.question + "{");
                                }
                                if (question.reponseA != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseA + "%" + question.reponseA);
                                }
                                if (question.reponseB != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseB + "%" + question.reponseB);
                                }
                                if (question.reponseC != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseC + "%" + question.reponseC);
                                }
                                if (question.reponseD != "")
                                {
                                    Console.WriteLine("\t=" + question.reponseD);
                                }
                                if (question.reponseE != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseE + "%" + question.reponseE);
                                }
                                Console.WriteLine("}\n\r");

                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Erreur : " + ex.ToString());
                                Marshal.ReleaseComObject(xlRange);
                                Marshal.ReleaseComObject(xlWorksheet);
                                xlWorkBook.Close();
                                XlApp.Quit();
                            }
                        }
                        else if(Convert.ToInt32(question.ptsReponseE) == 100)
                        {
                            try
                            {
                                if (question.question != "")
                                {
                                    file.WriteLine(question.question + "{");
                                }
                                if (question.reponseA != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseA + "%" + question.reponseA);
                                }
                                if (question.reponseB != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseB + "%" + question.reponseB);
                                }
                                if (question.reponseC != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseC + "%" + question.reponseC);
                                }
                                if (question.reponseD != "")
                                {
                                    file.WriteLine("\t~%" + question.ptsReponseD + "%" + question.reponseD);
                                }
                                if (question.reponseE != "")
                                {
                                    file.WriteLine("\t=" + question.reponseE);
                                }
                                file.WriteLine("}\n\r");
                                if (question.question != "")
                                {
                                    Console.WriteLine(question.question + "{");
                                }
                                if (question.reponseA != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseA + "%" + question.reponseA);
                                }
                                if (question.reponseB != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseB + "%" + question.reponseB);
                                }
                                if (question.reponseC != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseC + "%" + question.reponseC);
                                }
                                if (question.reponseD != "")
                                {
                                    Console.WriteLine("\t~%" + question.ptsReponseD + "%" + question.reponseD);
                                }
                                if (question.reponseE != "")
                                {
                                    Console.WriteLine("\t=" + question.reponseE);
                                }
                                Console.WriteLine("}\n\r");

                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Erreur : " + ex.ToString());
                                Marshal.ReleaseComObject(xlRange);
                                Marshal.ReleaseComObject(xlWorksheet);
                                xlWorkBook.Close();
                                XlApp.Quit();
                            }
                        }
                    }
                }
                
            }

            Console.WriteLine("L'application a fonctionné correctement. Appuyez sur une touche pour quitter");
            Console.ReadKey();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkBook.Close();
            XlApp.Quit();
        }

    }
    public class Question
    {
        public string question;
        public string reponseA;
        public string ptsReponseA;
        public string reponseB;
        public string ptsReponseB;
        public string reponseC;
        public string ptsReponseC;
        public string reponseD;
        public string ptsReponseD;
        public string reponseE;
        public string ptsReponseE;
        public bool choixMultiple;
        public Question(string defQuestion, bool defChoixMultiple, string defReponseA, string defPtsReponseA, string defReponseB, string defPtsReponseB, string defReponseC, string defPtsReponseC, string defReponseD, string defPtsReponseD, string defReponseE, string defPtsReponseE) {
            question = defQuestion;
            choixMultiple = defChoixMultiple;
            reponseA = defReponseA;
            ptsReponseA = defPtsReponseA;
            reponseB = defReponseB;
            ptsReponseB = defPtsReponseB;
            reponseC = defReponseC;
            ptsReponseC = defPtsReponseC;
            reponseD = defReponseD;
            ptsReponseD = defPtsReponseD;
            reponseE = defReponseE;
            ptsReponseE = defPtsReponseE;

        }
    }
}
