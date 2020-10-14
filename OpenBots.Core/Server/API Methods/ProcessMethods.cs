﻿using RestSharp;
using System.Net.Http;

namespace OpenBots.Core.Server.API_Methods
{
    public class ProcessMethods
    {
        public static void UploadProcess(RestClient client, string name, string filePath)
        {
            var request = new RestRequest("api/v1/Processes", Method.POST);
            request.AddParameter("Name", name);
            request.AddParameter("Status", "Published");
            request.RequestFormat = DataFormat.Json;

            request.AddHeader("Content-Type", "multipart/form-data"); 
            request.AddFile("File", filePath);

            var response = client.Execute(request);

            if (!response.IsSuccessful)
                throw new HttpRequestException($"Status Code: {response.StatusCode} - Error Message: {response.ErrorMessage}");
        }
    }
}