using Newtonsoft.Json.Linq;
using System;
using System.Diagnostics;
using System.IO;
using System.Net;

namespace DummyClient
{
    class Program
    {
        static void Main(string[] args)
        {
            string URL = "https://dev.spectrum.pitneybowes.com/rest/GeocodeAddressBaseExpandCandidates/results.json?Data.AddressLine1=High&Data.PostalCode=AB24%203EE";
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(URL);

            request.Method = "GET";
            request.ContentLength = 0;
            request.ContentType = "text/xml";
            request.Credentials = new NetworkCredential("RI008PE", "ZS62R=8*8h");

            try
            {
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                string responseValue = string.Empty;

                if (response.StatusCode != HttpStatusCode.OK)
                {
                    string message = String.Format("Request failed. Received HTTP {0}", response.StatusCode);
                    throw new ApplicationException(message);
                }

                using (Stream responseStream = response.GetResponseStream())
                {
                    using (var reader = new StreamReader(responseStream))
                    {
                        string responseBody = reader.ReadToEnd();
                        JObject jsonResult = JObject.Parse(responseBody);
                        JArray _addresses = jsonResult["output_port"].Value<JArray>();

                        foreach (dynamic address in _addresses)
                        {
                            Debug.Write((string)address.PostalCode);
                            Debug.Write(" ");
                            Debug.Write((string)address.StreetName);
                            Debug.Write(" ");
                            Debug.Write((string)address.StreetSuffix);
                            Debug.Write(" ");
                            Debug.WriteLine((string)address.City);
                            Debug.Write(" ");
                            if (address.GBR.UPRN != null)
                                Debug.WriteLine((string)address.GBR.UPRN);
                        }
                    }
                }
            }
            catch (WebException e)
            {
                if (e.Status == WebExceptionStatus.ProtocolError)
                {
                    Debug.WriteLine("Status Code : {0}", ((HttpWebResponse)e.Response).StatusCode);
                    Debug.WriteLine("Status Description : {0}", ((HttpWebResponse)e.Response).StatusDescription);
                }
            }
        }
    }
}
