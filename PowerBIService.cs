using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using System.Text.Json;
using System.Net;
using System.IO;

namespace PowerPublisher
{
    public class PowerBiService
    {
        private readonly PowerBiSettings settings;

        public PowerBiService(PowerBiSettings pbiSettings)
        {
            this.settings = pbiSettings;
        }

        public void PublishReport(string fileName, string reportName)
        {
            PowerBiService s = new PowerBiService(settings);
            var tokens = s.GetAccessToken().Result;

            var uri = "https://api.powerbi.com/v1.0/myorg/groups/" + settings.WorkspaceId + "/imports?datasetDisplayName="+ reportName + "&nameConflict=CreateOrOverwrite";
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            request.Method = "POST";
            string boundary = "----------BOUNDARY";
            byte[] boundaryBytes = System.Text.Encoding.ASCII.GetBytes("\r\n--" + boundary + "\r\n");
            request.ContentType = "multipart/form-data; boundary=" + boundary;
            request.Headers.Add("Authorization", "Bearer " + tokens.access_token);

            string bodyTemplate = "Content-Disposition: form-data; filename=\"{0}\"\r\nContent-Type: application/octet-stream\r\n\r\n";
            string body = string.Format(bodyTemplate, reportName);
            byte[] bodyBytes = System.Text.Encoding.UTF8.GetBytes(body);

            using (Stream rs = request.GetRequestStream())
            {
                rs.Write(boundaryBytes, 0, boundaryBytes.Length);
                rs.Write(bodyBytes, 0, bodyBytes.Length);

                using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    byte[] buffer = new byte[4096];
                    int bytesRead = 0;
                    while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) != 0)
                    {
                        rs.Write(buffer, 0, bytesRead);
                    }
                }

                byte[] trailer = System.Text.Encoding.ASCII.GetBytes("\r\n--" + boundary + "--\r\n");
                rs.Write(trailer, 0, trailer.Length);
            }

            try
            {
                using (HttpWebResponse response = request.GetResponse() as System.Net.HttpWebResponse)
                {
                    if (response.StatusCode == HttpStatusCode.Accepted)
                    {
                        Console.WriteLine("Published " + fileName);
                    } 
                    else
                    {
                        Console.WriteLine("Error publishing  " + fileName + "(" + response.StatusCode +  ")");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error publishing  " + fileName + ": " + e.Message); 
            }
        }

        public async Task<GenerateAccessTokenResponse> GetAccessToken()
        {
            var httpClient = new HttpClient();

            var accessTokenRequestBody = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("grant_type", settings.GrantType),
                new KeyValuePair<string, string>("username", settings.PbiUsername),
                new KeyValuePair<string, string>("password", settings.PbiPassword),
                new KeyValuePair<string, string>("client_id", settings.ApplicationId),
                new KeyValuePair<string, string>("resource", settings.ResourceUrl),
                new KeyValuePair<string, string>("scope", settings.Scope)
            });

            // Get the access token
            HttpResponseMessage accessTokenResponse = await httpClient.PostAsync(settings.AuthorityUrl, accessTokenRequestBody);

            if (!accessTokenResponse.IsSuccessStatusCode)
            {
                throw new Exception("Exception when fetching PowerBI access token");
            }

            var contents = await accessTokenResponse.Content.ReadAsStringAsync();
            return JsonSerializer.Deserialize<GenerateAccessTokenResponse>(contents);
        }
    }

    public class GenerateAccessTokenResponse
    {
        public string access_token { get; set; }
        public string refresh_token { get; set; }
        public string id_token { get; set; }
    }
}
