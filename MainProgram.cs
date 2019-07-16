using System;
using RestSharp;
using RestSharp.Authenticators;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using S = System.Console;
using System.IO;

namespace MailAutomation

{
    class MainProgram
    {


        /// <summary>
        /// IDS CHANGE BLOCK - Change this value to the name of the Settings File.
        /// Currently Settings.ini.
    
        private const string SettingsFileName = "Settings.ini";

        private const int MAX_SETTINGS_FIELDS = 40;
        private const int MAX_INTERVIEW_IDS = 500;


        private static string PathToWordTemplate;
        private static string PathToXMLFiles;
        private static string InterviewDoobloAPIurl;
        private static string PathToWordDoc;
        private static string PathToPDFFile;
        private static string SurveyID;
        private static string RESTAPIKey;
        private static string DoobloUserName;
        private static string DoobloPassWord;
        private static int CountOfDatabaseRecordsFetched;
        private static int NumberOfFieldsFromXML; 
        private static string XMLFileName;
        private static string WordDocFileName;
        private static string PDFFileName;
        private static string WordTemplateFileName;
        private static string ClientsHistory;
 

        static List<int> ProcessedInterviewIDs;

        RestResponse resp;

        public static bool CorruptRecord;


        /// <summary>
        /// Used to Read the Settings.ini file.
        /// </summary>
        /// <param name="FileName">Contains the Settings file name with extension.</param>
        /// <param name="NumberOfFields">Number of Fields in the Settings File. (Default 9)</param>
        /// <param name="ValuePosition">The Number of Strings in One Line before the value of the fieldName. (Default 0-3)</param>
        private static void ReadSettingsFile(SendingEmail sender, int ValuePosition = 3)
        {
            string line;
            int count = 0;
            string[] values = new string[MAX_SETTINGS_FIELDS];

           

            // Read the file and display it line by line.  
            System.IO.StreamReader file =
                new System.IO.StreamReader(SettingsFileName);
            while ((line = file.ReadLine()) != null)
            {
                
                values[count] = line.Split(' ')[ValuePosition];
                System.Console.WriteLine(count + " Value: " + values[count]);
                count++;



            }
            sender = new SendingEmail(values[0], 
                                      values[1], 
                                      values[2], 
                                      values[3], 
                                      values[4], 
                                      values[16],
                                      values[21],
                                      values[22],
                                      values[23],
                                      values[24]);
            PathToXMLFiles = values[5];
            PathToWordTemplate = values[6];
            InterviewDoobloAPIurl = values[7];
            PathToWordDoc = values[8];
            PathToPDFFile = values[9];
            SurveyID = values[10];
            RESTAPIKey = values[11];
            DoobloUserName = values[12];
            DoobloPassWord = values[13];
            int.TryParse(values[14], out CountOfDatabaseRecordsFetched);
            int.TryParse(values[15], out NumberOfFieldsFromXML);
            XMLFileName = values[17];
            WordDocFileName = values[18];
            PDFFileName = values[19];
            WordTemplateFileName = values[20];
            ClientsHistory = values[25];

            Console.WriteLine($"Number of Fields Parsed: { count }");
            
            file.Close();
            System.Console.ReadLine();
            
        }


 




        /// Connects to Dooblo website, Obtains interview data in XML Format
        /// And Writes the XML To the Text File Path and Name Given Above.
        /// Calls the ExtractFieldsFromXML Method
        /// returns a boolean value, whether the operation was successful.


        private int ReadWithRESTSaveXMLReportFile()
        {
            ProcessedInterviewIDs = new List<int>();
            int interviewNum;
            bool exists;
            //------------------------------------------------------------------------------------------
            // STEP 1: Call SurveyInterviewIDs to get the interview ids of the survey (JSON format)
            //------------------------------------------------------------------------------------------

            //SubjectID and InterviewIDs are the same
            string urlInterviewIds = string.Format(
                InterviewDoobloAPIurl,
                SurveyID
            );

            RestRequest request;
            System.Console.WriteLine("urlInterviewIds:"+urlInterviewIds);
            System.Console.WriteLine("SurveyID: "+SurveyID);
            System.Console.ReadLine();

            RestClient client = new RestClient(urlInterviewIds);

            try
            {
                // Authenticate ourselves and login username and password given above in IDS Block
                client.Authenticator = new HttpBasicAuthenticator(RESTAPIKey+DoobloUserName, DoobloPassWord);
            } catch(Exception e)
            {
                S.WriteLine("Please check RESTAPIKey, DoobloUserName and DoobloPassWord in " + SettingsFileName);
                S.WriteLine("Error: " + e);
            }
            request = new RestRequest();

            request.Method = Method.GET;
            client.AddDefaultHeader("Accept", "Application/JSON");
            client.AddDefaultHeader("Accept-CharSet", "UTF-8");



            resp = (RestResponse)client.Execute(request);
            List<int> interviewIds = JsonConvert.DeserializeObject<List<int>>(resp.Content);


            System.Console.WriteLine("Pulled Subject IDs");
            foreach (int id in interviewIds)
            {
                System.Console.WriteLine(id);
            }
            System.Console.WriteLine("Pulled interview IDs from server");

            exists = false;

            //------------------------------------------------------------------------------------
            // STEP 2: Group one interview ID into one XML file.
            //	Note that the SurveyInterviewData only returns XML data, no JSON
            //------------------------------------------------------------------------------------

            try
            {


                string interviewIdsInCurPack = string.Empty;
                

                    for (int i = 0; i < interviewIds.Count(); i++)
                    {

                    //Only two API calls allowed every second. Hence we need to pause the program.
                    System.Threading.Thread.Sleep(1000);

                    interviewIdsInCurPack = string.Format("{0},{1}", interviewIdsInCurPack, interviewIds[i]);

                        interviewIdsInCurPack = interviewIdsInCurPack.Remove(0, 1);
                    exists = CheckIfInterviewIdExists(interviewIds[i]);
                    if ( exists == true)
                    {
                        S.WriteLine(interviewIds[i] + " already exists in the client history text file.");
                        continue;                        
                    }


                    ProcessedInterviewIDs.Add(interviewIds[i]);

                    string urlInterviewData = string.Format(
                    "https://api.dooblo.net/newapi/SurveyInterviewData?subjectIDs={0}&surveyID={1}&onlyHeaders=false&includeNulls=false",
                    interviewIdsInCurPack,
                    SurveyID
                );
                        client = new RestClient(urlInterviewData);
                        client.Authenticator = new HttpBasicAuthenticator(RESTAPIKey + DoobloUserName, DoobloPassWord);
                        request = new RestRequest();
                        request.Method = Method.GET;
                        client.AddDefaultHeader("Accept", "Text/XML");
                        client.AddDefaultHeader("Accept-CharSet", "UTF-8");
                        resp = (RestResponse)client.Execute(request);
                    //------------------------------------------------------------------------------------
                    // STEP 3: This is the actual XML Data as it comes as an answer, your code should 
                    //	process xmlInterviewData to process the interview data of the 99 interviews of 
                    //	this pack
                    //------------------------------------------------------------------------------------
                    interviewNum = interviewIds[i];

                    string xmlInterviewData = resp.Content.ToString();
                        System.IO.File.WriteAllText(PathToXMLFiles + XMLFileName.Replace(".xml", "") + interviewNum + ".xml", xmlInterviewData);

                        S.WriteLine($"Written XML Data of { interviewIds[i] } from Server to Disk.");

                        interviewIdsInCurPack = string.Empty;
                                                             
                    }
                if (ProcessedInterviewIDs.Count() == 0)
                {
                    System.Console.WriteLine("No new records added");
                    System.Environment.Exit(0);
                }

            } catch (Exception e)
                {
                    S.WriteLine("Error Occurred " + e);
                }
            AddProcessedInterviewIDs(ProcessedInterviewIDs);
            return ProcessedInterviewIDs.Count();

        }
        void AddProcessedInterviewIDs(List<int> ProcessedInterviewIDs)
        {
            string contents = File.ReadAllText(ClientsHistory);
            
            foreach (int ID in ProcessedInterviewIDs)
            {
                contents = contents + Environment.NewLine + ID.ToString()+Environment.NewLine;
            }
            File.WriteAllText(ClientsHistory, contents);

        }
        bool CheckIfInterviewIdExists(int InterviewId)
        {
            string IDList = File.ReadAllText(ClientsHistory);

            if (IDList.Contains(InterviewId.ToString())) return true;

            else return false;

        }

        /// The main entry point for the application.
        [STAThread]
        public static void Main(string[] args)
        {
            string[] Names, Values;
            int InterviewID = 42;
            try
            {
                MainProgram mail = new MainProgram();
                CreateWordDoc creator = new CreateWordDoc();
                SendingEmail sender = new SendingEmail();
                WordToPDF converter = new WordToPDF();
                ExtractFieldsXML job = new ExtractFieldsXML();
                
                ReadSettingsFile(sender);

                int count =  mail.ReadWithRESTSaveXMLReportFile();

                for (int i = 0; i < count; i++)
                {
                    try
                    {
                        InterviewID = ProcessedInterviewIDs.ElementAt<int>(i);

                        job.ExtractFieldsfromXML(InterviewID, PathToXMLFiles + XMLFileName,
                                                 NumberOfFieldsFromXML, out Names, out Values);

                        creator.CreateWordDocFromFields(InterviewID, PathToWordTemplate + WordTemplateFileName, PathToWordDoc + WordDocFileName, Names, Values);

                        string PDFFilePathFileName = PathToPDFFile + PDFFileName.Replace(".pdf", "") + "_" + Values[9] + ".pdf";

                        string WordFilePathFileName = PathToWordDoc + WordDocFileName;

                        converter.ConvertDOCtoPDF(InterviewID, WordFilePathFileName, PDFFilePathFileName);

                        sender.SendEmail(InterviewID, PDFFilePathFileName, Values[10]);

                    } catch (Exception e)
                    {
                        S.WriteLine($"Error in Processing Record ID: { InterviewID }.");
                        S.WriteLine("Exception " + e);
                        if (i <= count - 2) S.WriteLine("Moving to next record.");
                        else S.WriteLine("All records processed.");
                        continue;
                    }
                }

            }
            catch (Exception e)
            {
                System.Console.WriteLine(e);
            }

            finally
            {
                System.Console.ReadLine();
            }

        }

    }
}
