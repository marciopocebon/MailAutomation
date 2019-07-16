using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

using S = System.Console;

namespace MailAutomation
{
    class ExtractFieldsXML
    {
        
        public void ExtractFieldsfromXML(int InterviewNum, string XMLFilePathFileName, int NumberOfFieldsFromXML, out string[] Names, out string[] Values)
        {
            int i;
            string Organization, Email;
            ///Field Names
            string[] ScoreNames= new string[NumberOfFieldsFromXML];
            string[] OutValues = new string[NumberOfFieldsFromXML];

            ScoreNames[0] = "BCRS";
            ScoreNames[1] = "AC";
            ScoreNames[2] = "AAP";
            ScoreNames[3] = "IVC";
            ScoreNames[4] = "DC";
            ScoreNames[5] = "SI";
            ScoreNames[6] = "CAS";
            ScoreNames[7] = "BCRS2";
            ScoreNames[8] = "MATRIX LABEL";
            ScoreNames[9] = "Organization";
            ScoreNames[10] = "Email Id";



            string score;

            string str;

            XMLFilePathFileName = XMLFilePathFileName.Replace(".xml", InterviewNum.ToString()) + ".xml";
            try
            {
                //The fields have to be parsed in the order that they appear on the XML file
                using (XmlReader reader = XmlReader.Create(XMLFilePathFileName))
                {


                    while (!reader.Value.Equals("Organization") && !reader.EOF)
                    {
                        reader.Read();
                    }

                    while (!reader.Name.Equals("QuestionAnswer") && !reader.EOF)
                    {
                        reader.Read();
                    }
                    if (reader.EOF)
                    {
                        Names = new string[NumberOfFieldsFromXML];
                        Values = new string[NumberOfFieldsFromXML];
                        return;
                    }
                    Organization = reader.ReadElementContentAsString();
                    OutValues[9] = Organization;


                    while (!reader.Value.Equals("Email Id") && !reader.EOF)
                    {
                        reader.Read();
                    }

                    while (!reader.Name.Equals("QuestionAnswer") && !reader.EOF)
                    {
                        reader.Read();
                    }

                    if (reader.EOF)
                    {
                        Names = new string[NumberOfFieldsFromXML];
                        Values = new string[NumberOfFieldsFromXML];
                        return;
                    }
                    Email = reader.ReadElementContentAsString();
                    OutValues[10] = Email;


                    while (!reader.EOF)
                    {


                        for (i = 0; i < NumberOfFieldsFromXML; i++)
                        {
                            if (i == 9 || i == 10) continue; //Organization and Client Email Id already stored

                            str = ScoreNames[i];

                            //For this to work the field names in the file should be in the
                            //same order as that given in the ScoreNames array.

                            while (!reader.Value.Equals(str) && !reader.EOF)
                            {
                                reader.Read();
                            }
                            while (!reader.Name.Equals("TopicAnswer") && !reader.EOF)
                            {
                                reader.Read();
                            }

                            if (reader.EOF)
                            {
                                Names = new string[NumberOfFieldsFromXML];
                                Values = new string[NumberOfFieldsFromXML];
                                return;
                            }
                            score = reader.ReadElementContentAsString();
                            OutValues[i] = score;


                        }
                    }
               }

            

        
            } catch(Exception e)
            {
                System.Console.WriteLine(e);
                System.Console.WriteLine("Erroneously formed XML for ID: " + InterviewNum);
                Names = new string[NumberOfFieldsFromXML];
                Values = new string[NumberOfFieldsFromXML];
                throw e;
            }

                for (i = 0; i < NumberOfFieldsFromXML; i++)
                {
                    System.Console.WriteLine(ScoreNames[i] + ": " + OutValues[i]);
                }
                Names = ScoreNames;
                Values = OutValues;
                

                System.Console.WriteLine("ID: " + InterviewNum + " Read XML Fields correctly.");
            }

        }
    }

 
