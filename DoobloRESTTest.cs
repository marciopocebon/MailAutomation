using Newtonsoft.Json;
using RestSharp;
using RestSharp.Authenticators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



namespace MailAutomation
{
    class DoobloRESTTest
    {        //////////////////IDS CHANGE VALUES HERE////////////////////

        private static string XMLDownloadedFilePath = @"C:\Users\ADMIN\Desktop\";
        private static string XMLDownloadedFileName = "SurveyDatabase.xml";

        private static string InterviewDoobloAPIurl = "https://api.dooblo.net/newapi/SurveyInterviewIDs?surveyIDs={0}";


        //-------------------------------------------
        // Setting the parameters for the calls
        //-------------------------------------------
        private static string SurveyID = "a334da36-2678-41ca-bc1e-1d7e4bad9618";
        private static string UserName = "0e714f59-dc62-4a7b-afb5-2b6908752835/JijoThankachan";
        private static string PassWord = "Welcome101";
      private static void OldMain2()
        {
            string urlInterviewIds = string.Format(
                InterviewDoobloAPIurl,
                SurveyID
            );
            System.Console.WriteLine("urlInterviewIds: " + urlInterviewIds);
            System.Console.WriteLine("SurveyID: " + SurveyID);
            System.Console.ReadLine();
            RestClient client = new RestClient(urlInterviewIds);

            // Authenticate ourselves and login username and password given above in IDS Block
            client.Authenticator = new HttpBasicAuthenticator(UserName, PassWord);
            RestRequest request = new RestRequest();


            request.Method = Method.GET;
            client.AddDefaultHeader("Accept", "Application/JSON");
            client.AddDefaultHeader("Accept-CharSet", "UTF-8");



            RestResponse resp = (RestResponse)client.Execute(request);
            List<int> interviewIds = JsonConvert.DeserializeObject<List<int>>(resp.Content);
            if (interviewIds == null)
            {
                System.Console.WriteLine("No data returned.");
            }

            int counter = 0;
            string interviewIdsInCurPack = string.Empty;




            for (int i = 0; i < 2; i++)
            {
                interviewIdsInCurPack = string.Format("{0},{1}", interviewIdsInCurPack, interviewIds[i]);
                counter += 1;
                bool lastInterviewID = i == interviewIds.Count - 1;
                if (counter == 10 || lastInterviewID)
                {
                    interviewIdsInCurPack = interviewIdsInCurPack.Remove(0, 1);
                    string urlInterviewData = string.Format(
                        "https://api.dooblo.net/newapi/SurveyInterviewData?subjectIDs={0}&surveyID={1}&onlyHeaders=false&includeNulls=false",
                        interviewIdsInCurPack,
                        SurveyID
                    );
                    client = new RestClient(urlInterviewData);
                    client.Authenticator = new HttpBasicAuthenticator(UserName, PassWord);
                    request = new RestRequest();
                    request.Method = Method.GET;
                    client.AddDefaultHeader("Accept", "text/xml");
                    client.AddDefaultHeader("Accept-Charset", "utf-8");
                    resp = (RestResponse)client.Execute(request);

                    //------------------------------------------------------------------------------------
                    // STEP 3: This is the actual XML Data as it comes as an answer, your code should 
                    //	process xmlInterviewData to process the interview data of the 99 interviews of 
                    //	this pack
                    //------------------------------------------------------------------------------------
                    string xmlInterviewData = resp.Content.ToString();
                    
                    System.IO.File.WriteAllText(XMLDownloadedFilePath + XMLDownloadedFileName, xmlInterviewData);

                    counter = 0;
                    interviewIdsInCurPack = string.Empty;
                    System.Console.WriteLine(xmlInterviewData);
                }
                if (lastInterviewID)
                {
                    break;
                }
            }
            
            System.Console.ReadLine();

        }


        public static void Main(string[] args)
        {
            new DoobloRESTTest().TraverseInterviewData();
        }
        private bool TraverseInterviewData()
        {
            //-------------------------------------------
            // Setting the parameters for the calls
            //-------------------------------------------
            string surveyID = "a334da36-2678-41ca-bc1e-1d7e4bad9618";

            string username = "0e714f59-dc62-4a7b-afb5-2b6908752835/JijoThankachan";
            string password = "Welcome101";

            //Thomas - need startdate and enddate in DateTime format
            string StartDate = new DateTime(2019, 6, 1).ToString();
            string EndDate = new DateTime(2019, 6, 30).ToString();

            string LastModifiedDate = new DateTime(2019, 6, 30).ToString();

    //------------------------------------------------------------------------------------------
    // STEP 1: Call SurveyInterviewIDs to get the interview ids of the survey (JSON format)
    //------------------------------------------------------------------------------------------
    string urlInterviewIds = string.Format(
        "https://api.dooblo.net/newapi/SurveyInterviewIDs?surveyIDs={0}",
        surveyID);

            //lastModifiedStart={1}lastModifiedEnd={2}
            RestClient client = new RestClient(urlInterviewIds);
            client.Authenticator = new HttpBasicAuthenticator(username, password);
            RestRequest request = new RestRequest();
            request.Method = Method.GET;
            client.AddDefaultHeader("Accept", "application/json");
            client.AddDefaultHeader("Accept-Charset", "utf-8");

            RestResponse resp = (RestResponse) client.Execute(request);

            //Thomas - the variable "resp" will contain info of interview ids with last modified date and time, Need to extract the latest subjid depending on the date.


            List<int> interviewIds = JsonConvert.DeserializeObject<List<int>>(resp.Content);

            System.Console.WriteLine("Pulled Interview IDs");
            foreach (int id in interviewIds)
            {
                System.Console.WriteLine(id);
            }
            //------------------------------------------------------------------------------------
            // STEP 2: Group interview IDs into groups of 99 interviews, then call the 
            //	SurveyinterviewData for each pack of 99 interviews to get the XML Data
            //	Note that the SurveyInterviewData only returns XML data, no JSON
            //------------------------------------------------------------------------------------
            
            string interviewIdsInCurPack = string.Empty;
            for (int i = 0; i < 10; i++)
            {
                interviewIdsInCurPack = string.Format("{0},{1}", interviewIdsInCurPack, interviewIds[i]);
                bool lastInterviewID = true;
               
                interviewIdsInCurPack = interviewIdsInCurPack.Remove(0, 1);
                string urlInterviewData = string.Format(
                    "https://api.dooblo.net/newapi/SurveyInterviewData?subjectIDs={0}&surveyID={1}&onlyHeaders=false&includeNulls=false",
                    interviewIdsInCurPack,
                    SurveyID
                );
                client = new RestClient(urlInterviewData);
                client.Authenticator = new HttpBasicAuthenticator(username, password);
                request = new RestRequest();
                request.Method = Method.GET;
                client.AddDefaultHeader("Accept", "text/xml");
                client.AddDefaultHeader("Accept-Charset", "utf-8");
                resp = (RestResponse)client.Execute(request);

                //------------------------------------------------------------------------------------
                // STEP 3: This is the actual XML Data as it comes as an answer, your code should 
                //	process xmlInterviewData to process the interview data of the 99 interviews of 
                //	this pack
                //------------------------------------------------------------------------------------
                string xmlInterviewData = resp.Content.ToString();
                System.IO.File.WriteAllText(XMLDownloadedFilePath+XMLDownloadedFileName, xmlInterviewData);

                    
                interviewIdsInCurPack = string.Empty;
                
                if (lastInterviewID)
                {
                    break;
                }
            }
            return true;
        }



    }
}
