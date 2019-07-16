using System;
using Word = Microsoft.Office.Interop.Word;



class WordToPDF
{

    Word.Application WORD;
    Word.Document doc;


    /// </summary>
    public void ConvertDOCtoPDF(int InterviewID, string PathToWordDocWithFileName, string PathToPDFWithFileName)
    {

        object misValue = System.Reflection.Missing.Value;

        try
        {

            WORD = new Word.Application();

            PathToWordDocWithFileName = PathToWordDocWithFileName.Replace(".docx", "") + InterviewID + ".docx";
            doc = WORD.Documents.Open(PathToWordDocWithFileName);
            doc.Activate();
            //Extract fields from PDF
            
            PathToPDFWithFileName = PathToPDFWithFileName.Replace(".pdf", "") + InterviewID + ".pdf";
            doc.SaveAs2(PathToPDFWithFileName, Word.WdSaveFormat.wdFormatPDF, misValue, misValue, misValue,
            misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            System.Console.WriteLine($"{InterviewID} Report saved successfully as PDF after conversion.");
        }
        catch (Exception e) { System.Console.WriteLine("Error Occurred " + e); }
        finally
        {
            doc.Close();
            WORD.Quit();
            releaseObject(doc);
            releaseObject(WORD);
        }
    }


    private void releaseObject(object obj)
    {
        try
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            obj = null;
        }
        catch (Exception ex)
        {
            System.Console.WriteLine(ex);
        }
        finally
        {
            GC.Collect();
        }
    }

}
