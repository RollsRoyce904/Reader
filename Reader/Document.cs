using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Reader
{
    class Document
    {
        public String FilePath { get; set; }
        public String FileName { get; set; }



        public static void ExcelReader(String filepath)
        {
            Mapper tags = new Mapper();
            String path = filepath;//"C:\Users\slip4\Desktop\RemittMock.csv"
            String parsable = String.Empty;
            List<String> remittanceDetails = new List<string>();


            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            var myXL = new Excel.Application();

            Excel.Workbook xlWorkBook = myXL.Workbooks.Open(@"C:\Users\slip4\Desktop\RemittMock.csv", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            int sheetCount = xlWorkBook.Worksheets.Count;

            Excel.Worksheet myActiveSheet = xlWorkBook.Worksheets[1];

            // myActiveSheet = xlWorkBook.ActiveSheet();

            Excel.Range range = myActiveSheet.Cells[1, 1];

            string myValue = myActiveSheet.Cells[1, 1].Value();
            string myValue2 = myActiveSheet.Cells[1, 1].Value();
            string myValue3 = myActiveSheet.Cells[1, 1].Value();

            ////Here you can make a collection to get all work sheets using a loop once you get 
            ////the number of worksheets
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            range = xlWorkSheet.UsedRange; //checks which cells being used
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            string temp = range.Cells[1, 1].Value2.ToString();

            //tags.Headers.Add(temp.ToString());
            List<String> headerTags = new List<string>();

            headerTags.Add(temp.ToString());

            tags.Headers = headerTags;


            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                //make method to grab excel headers make them a class or class vars
                //then link the rest of the data to them using dictionary or something

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    //new line
                    if (cCnt > 1)
                    {
                        Console.WriteLine();
                    }


                    //write values out
                    if (range.Cells[rCnt, cCnt] != null && range.Cells[rCnt, cCnt].Value2 != null)
                    {
                        Console.WriteLine(range.Cells[rCnt, cCnt].Value2.ToString());

                        if (rCnt == 1 && cCnt == 1)
                        {
                            //rCnt 1 cCnt 1
                            parsable = range.Cells[rCnt, cCnt].Value2.ToString();
                        }
                        else
                        {
                            remittanceDetails.Add(range.Cells[rCnt, cCnt].Value2.ToString());
                        }

                    }

                    //This is a class you made, first val is the id, sec val is question, third is answer
                    //DataCell newCell = new DataCell(1, myValue2, myValue3);
                }
            }

            xlWorkBook.Close(true, null, null);
            myXL.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(myXL);


            ParseData(parsable, remittanceDetails);


        }

        public static void ParseData(String header, List<String> details)
        {
            String parseThis = header;
            List<String> remitance = details; //the data
            Dictionary<String, String> RemmitanceSection = new Dictionary<String, String>();

            string[] Headers = parseThis.Split(',');
            string line;


            String[] separatedRemit;// array for parsing each line of data


            foreach (var item in remitance)
            {
                separatedRemit = item.Split(',');

                for (int i = 0; i < separatedRemit.Length; i++)
                {
                    line = separatedRemit[i];
                    //Mapp headers to data
                    RemmitanceSection.Add(Headers[i], line);
                }

            }

        }


        //public static void MapData()

    }
}
