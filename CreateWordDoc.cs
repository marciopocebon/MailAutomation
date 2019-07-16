
using System;

using Word = Microsoft.Office.Interop.Word;


namespace MailAutomation
{
   

    class CreateWordDoc
    {
        object OMissing;
        Word.Application oWord;
        Word.Document oDoc;

        public bool CreateWordDocFromFields(int InterviewID,
                                            string WordTemplateFilePathFileName, 
                                            string WordFilePathFileName,
                                            string[] Names, string[] Values)
        {
            //Start Word and create a new document.
            try
            {
                OMissing = System.Reflection.Missing.Value;
                //Start Word and create a new document.

                oWord = new Word.Application();
                oWord.Visible = false;

                //Creating files from Template

                System.Console.WriteLine(WordTemplateFilePathFileName);
                object oTemplate = WordTemplateFilePathFileName;
                oDoc = oWord.Documents.Add(ref oTemplate, ref OMissing,
                ref OMissing, ref OMissing);


               


                // find each field and replace it

                for (int i = 0; i < Names.Length; i++)
                {
                    FindReplace(Names[i], Values[i]);
                }



                object saveFile = (object)(WordFilePathFileName.Replace(".docx", "") + InterviewID + ".docx");

                oDoc.SaveAs2(ref saveFile, ref OMissing, ref OMissing, ref OMissing, ref OMissing,
                    ref OMissing, ref OMissing, ref OMissing, ref OMissing, ref OMissing,
                    ref OMissing, ref OMissing, ref OMissing, ref OMissing);

                oDoc.Close();
                oWord.Quit();
                releaseObject(oDoc);
                releaseObject(oWord);
                System.Console.WriteLine($"Word Document for { InterviewID } Stored With New Fields.");
                return true;
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("Error in Forming Word Document " + ex);
                return false;
            }
        }



        private void FindReplace(string FieldText, string FieldValue)
        {
            Word.Find findObject = oWord.Selection.Find;

            findObject.Text = FieldText;
            findObject.ClearFormatting();

            
            switch(FieldText)
            {

                case "Organization":
                    findObject.Text = "<Organization>";
                    break;
                case "BCRS":
                    findObject.Text ="<T_SCORES_BCRS%>" ;
                    break;
                case "AC":
                    findObject.Text = "<T_SCORES_AC>";
                    break;

                case "AAP":
                    findObject.Text = "<T_SCORES_AAP>";
                    break;

                case "IVC":
                    findObject.Text = "<" + "T_SCORES_IVC" + ">";
                    break;

                case "SI":
                    findObject.Text = "<" + "T_SCORES_SI" + ">";
                    break;

                case "DC":
                    findObject.Text = "<" + "T_SCORES_DC" + ">";
                    break;
                case "MATRIX LABEL":
                    findObject.Text = "<" + "T_SCORES_MATLABEL" + ">";
                    break;
                default:
                    System.Console.WriteLine(FieldText);
                    break;




            }

            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = FieldValue;
            object replace = Word.WdReplace.wdReplaceAll;

            findObject.Execute(ref OMissing, ref OMissing, ref OMissing, ref OMissing, ref OMissing,
            ref OMissing, ref OMissing, ref OMissing, ref OMissing, ref OMissing,
            ref replace, ref OMissing, ref OMissing, ref OMissing, ref OMissing);
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
}