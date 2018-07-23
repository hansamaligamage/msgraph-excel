Microsoft Graph is the API for Microsoft 365 that provides access to all the data available in Office 365, we can connect to mail, calendar, contacts, documents, directories, users. 

Microsoft Graph exposes APIs for Azure Active Directory, Office 365 services like Sharepoint, OneDrive, Outlook, Exchange, Microsoft Team services, OneNote, Planner, Excel 

We can access to all these Office 365 products through a single REST endpoint and manage millions of data in Microsoft Cloud

This sample application is built using a Azure Function App Http trigger template, It's going to read a JSON file and writes it into an excel table using Graph API, and it generates a chart on top of that excel file. You can find the json file from the root of the project directory

You can add a row to an existing excel file in OneDrive using GraphAPI like this 
 ```
 public async Task<bool> ModifyTable(string fileId, string sessionId, string worksheet, string table) 
        { 
            var success = false; 
            //Read data from JSON file 
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
```

We can create a chart based on excel data as below

```
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
```
