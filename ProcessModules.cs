
using System.IO;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace ProcessExcel
{
    public static class ProcessModules
    {

        [FunctionName("ProcessModules")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequest req, TraceWriter log)
        {
            int classes;
            int labs;
            string file = "sldevforum.xlsx";
            string worksheet = "Year1";
            string table = "Table2";
            log.Info("C# HTTP trigger function processed a request.");
            ExcelHelper excelHelper = new ExcelHelper();
           
            string fileId = excelHelper.RetrieveFiles(file);
            log.Info("RetrieveFiles status from main : " + fileId);

            if (!string.IsNullOrEmpty(fileId))
            {
                string columns = excelHelper.RetrieveTable(fileId, worksheet);
                if (Convert.ToInt32(columns) > 0)
                {
                    log.Info("Columns count : " + columns);

                    string sessionId = excelHelper.CreateSession(fileId, log);
                    if (!string.IsNullOrEmpty(sessionId))
                    {
                        log.Info("Session ID : " + sessionId);

                        var success = await excelHelper.ModifyTable(fileId, sessionId, worksheet, table);
                        log.Info("Modifying table : " + success);

                        if (success)
                        {
                            classes = excelHelper.GetTotal(fileId, sessionId, worksheet, "B2", "B21");
                            labs = excelHelper.GetTotal(fileId, sessionId, worksheet, "C2", "C21");
                            log.Info("No of classes : " + classes + " No of labs : " + labs);

                            excelHelper.CreateChart(fileId, worksheet, "columnclustered", "A2:C21");
                            success = excelHelper.CreateChart(fileId, worksheet, "pie", "A2:B7");
                        }
                    }
                }
            }
            return new HttpResponseMessage();
        }

        

    }
}
