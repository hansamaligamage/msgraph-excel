using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using RestSharp;

namespace ProcessExcel
{
    class ExcelHelper
    {

        string accessToken = System.Environment.GetEnvironmentVariable("AccessToken", EnvironmentVariableTarget.Process);
        string baseurl = "https://graph.microsoft.com/v1.0/";

        public string RetrieveFiles (string file)
        {
            string fileId = string.Empty;

            var client = new RestClient(baseurl + "me/drive/root/children/" + file);
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", accessToken);

            IRestResponse response = client.Execute(request);
            if (response.IsSuccessful)
            {
                string content = response.Content;
                JObject filedetails = (JObject)JsonConvert.DeserializeObject(content);
                fileId = filedetails["id"].ToString();
            }
            return fileId;
        }

        public string RetrieveTable(string fileId, string worksheet)
        {
            string columnCount = string.Empty;
            string header = "!A1:H1";
           
            var client = new RestClient(baseurl + "me/drive/items/" + fileId + "/workbook/worksheets/" + worksheet + "/Range(address='" + worksheet + header + "')");
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", accessToken);

            IRestResponse response = client.Execute(request);
            if (response.IsSuccessful)
            {
                string content = response.Content;
                JObject obj = (JObject)JsonConvert.DeserializeObject(content);
                columnCount = obj["columnCount"].ToString();
            }
            return columnCount;
        }

        public string CreateSession(string fileId, TraceWriter log)
        {
            string sessionId = string.Empty;

            var client = new RestClient(baseurl + "me/drive/items/" + fileId + "/workbook/createsession");
            var request = new RestRequest(Method.POST);
            request.AddHeader("Authorization", accessToken);
            request.AddHeader("persistSession", "true");

            IRestResponse response = client.Execute(request);
            if (response.IsSuccessful)
            {
                string content = response.Content;
                JObject session = (JObject)JsonConvert.DeserializeObject(content);
                sessionId = session["id"].ToString();
            }
            else
                log.Info("ERROR : " + response.ErrorMessage + " : " + response.StatusCode);
            return sessionId;
        }
        public async Task<bool> ModifyTable(string fileId, string sessionId, string worksheet, string table)
        {
            var success = false;
            List<CourseModule> modules = ReadJsonFile();

            var client = new HttpClient();
            var request = new HttpRequestMessage(HttpMethod.Post, baseurl + "me/drive/items/" + fileId + "/workbook/worksheets('" + worksheet + "')/Tables('" + table + "')/Rows");
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            request.Headers.Add("workbook-session-id", sessionId);

            string[] module;
            string[][] modulesArray = new string[modules.Count][];

            for (int i = 0; i< modules.Count(); i++) 
            {
                module = new string[] { modules[i].Module, modules[i].Classes.ToString(), modules[i].Labs.ToString(), modules[i].Points.ToString(), modules[i].Instructor, modules[i].StartDate, modules[i].EndDate,
                    modules[i].Weekend.ToString() };
                modulesArray[i] = module;
            }

            TableRequest tableRequest = new TableRequest();
            tableRequest.index = null;
            tableRequest.values = modulesArray;

            string jsonBody = JsonConvert.SerializeObject(tableRequest);
            request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            using (var response = await client.SendAsync(request))
            {
                string statusdescription = response.ReasonPhrase;
                success = response.IsSuccessStatusCode;
            }
            return success;
        }

        public List<CourseModule> ReadJsonFile()
        {
            List<CourseModule> items = new List<CourseModule>();
            using (StreamReader reader = new StreamReader(@"document.json"))
            {
                string json = reader.ReadToEnd();
                try
                {
                    items = JsonConvert.DeserializeObject<List<CourseModule>>(json);
                }
                catch(Exception ex)
                {
                    throw ex;
                }
            }
            return items;
        }

        public int GetTotal (string fileId, string sessionId, string worksheet, string fromColumn, string toColumn)
        {
            int noOfClasses = 0;

            var client = new RestClient(baseurl + "me/drive/items/" + fileId + "/workbook/functions/sum");
            var request = new RestRequest(Method.POST);
            request.AddHeader("Authorization", accessToken);
            request.AddHeader("workbook-session-id", sessionId);
            StringBuilder classes = new StringBuilder("{\"values\" :  [{ \"address\": \"" + worksheet +"!" + fromColumn + ":" + toColumn +"\" }]}");
            request.AddParameter("undefined", classes, ParameterType.RequestBody);
            IRestResponse response = client.Execute(request);
            if (response.IsSuccessful)
            {
                string content = response.Content;
                JObject sum = (JObject)JsonConvert.DeserializeObject(content);
                noOfClasses = Convert.ToInt32(sum["value"].ToString());
            }
            return noOfClasses;
        }

        public bool CreateChart (string fileId, string worksheet, string type, string columnsrange)
        {
            var client = new RestClient(baseurl + "me/drive/items/" + fileId + "/workbook/worksheets('" + worksheet + "')/Charts/Add");
            var request = new RestRequest(Method.POST);
            request.AddHeader("Authorization", accessToken);
            request.AddHeader("Content-Type", "application/json");
            ChartRequest chartRequest = new ChartRequest { type = type, sourcedata = columnsrange, seriesby = "Auto" };
            string jsonBody = JsonConvert.SerializeObject(chartRequest);
            request.AddParameter("undefined", jsonBody, "application/json", ParameterType.RequestBody);
            IRestResponse response = client.Execute(request);
            return response.IsSuccessful;
        }

    }
}
