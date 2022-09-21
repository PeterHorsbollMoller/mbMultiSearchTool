using System;
using System.Net.Http;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace SearchUsingServices
{
    public class SpectrumAddressBaseUK
    {
        //static HttpClient _client = new HttpClient();
        //static HttpWebRequest _client = new HttpWebRequest(); 
        //private static bool _hasBeenInitiated = false;

        private static List<string> _arrAddressesDescription = null;
        private static List<string> _arrAddressesAddress = null;
        private static List<string> _arrAddressesCity = null;
        private static List<string> _arrAddressesPostalCode = null;
        private static List<string> _arrAddressesLatitude = null;
        private static List<string> _arrAddressesLongitude = null;
        private static List<string> _arrAddressesURPN = null;
        private static List<string> _arrAddressesBuilding = null;

        static dynamic _addresses = null;
        private static string _latestURL = "";
        private static string _serverURL = "";
        private static string _userName = "";
        private static string _passWord = "";
        private static byte[] _aditionalEntropy = { 9, 4, 2, 6, 9 };

        //********************************************************************************************************
        //********************************************************************************************************
        public static void Initiate(int timeout, string serverURL, string userName, string passWord)
        {
            //if (_hasBeenInitiated != true)
            //{
                //_client.Timeout = new TimeSpan(0, timeout, 0);
                if (serverURL == "")
                {
                    //_client.BaseAddress = new Uri("http://rau02027.pbi.global.pvt:9292/rest/AddressBaseSearch/results.json");
                    _serverURL = "http://rau02027.pbi.global.pvt:9292/rest/AddressBaseSearch/results.json";
                }
                else
                {
                    //_client.BaseAddress = new Uri(serverURL);
                    _serverURL = serverURL;
                }
                //_hasBeenInitiated = true;

                if (userName != "")
                {
                    _userName = userName;
                    _passWord = passWord;
                }
            //}
        }

        //********************************************************************************************************
        //********************************************************************************************************
        public static int DoSearch(string houseNoSearchFor, string streetSearchFor, string buildingSearchFor, string postalSearchFor, string citySearchFor)
        {
            //MessageBox.Show(String.Format("DoSearch: postal='{0}' or city='{1}'", postalSearchFor, citySearchFor));

            _latestURL = string.Empty;

            string url = string.Empty;
            if (postalSearchFor != string.Empty)
            {
                if (url == string.Empty)
                    url = string.Format("?Data.PostalCode={0}", postalSearchFor);
                else
                    url = url + string.Format("&Data.PostalCode={0}", postalSearchFor);
            }

            if (citySearchFor != string.Empty)
            {
                if (url == string.Empty)
                    url = string.Format("?Data.City={0}", citySearchFor);
                else
                    url = url + string.Format("&Data.City={0}", citySearchFor);
            }

            if (houseNoSearchFor != string.Empty)
                if (url == string.Empty)
                    url = string.Format("?Data.Housenumber={0}", houseNoSearchFor);
                else
                    url = url + string.Format("&Data.Housenumber={0}", houseNoSearchFor); ;

            if (streetSearchFor != string.Empty)
            {
                if (url == string.Empty)
                    url = string.Format("?Data.Street={0}", streetSearchFor);
                else
                    url = url + string.Format("&Data.Street={0}", streetSearchFor);
            }

            if (buildingSearchFor != string.Empty)
            {
                if (url == string.Empty)
                    url = string.Format("?Data.BuildingName={0}", buildingSearchFor);
                else
                    url = url + string.Format("&Data.BuildingName={0}", buildingSearchFor);
            }

            _latestURL = url;

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(string.Format("{0}{1}", _serverURL, url));

            request.Method = "GET";
            request.ContentLength = 0;
            request.ContentType = "text/xml";
            if (_userName != "")
                request.Credentials = new NetworkCredential(_userName, _passWord);

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
                        //MessageBox.Show(String.Format("Response: {0}", responseBody.Substring(0,500)));

                        //MessageBox.Show("A");
                        JObject jsonResult = JObject.Parse(responseBody);
                        //MessageBox.Show("B");
                        JArray _addresses = jsonResult["Output"].Value<JArray>();
                        //MessageBox.Show("C");

                        //MessageBox.Show(String.Format("_addresses holds {0} addresses", _addresses.Count()));

                        _arrAddressesDescription = new List<string>();
                        _arrAddressesAddress = new List<string>();
                        _arrAddressesCity = new List<string>();
                        _arrAddressesPostalCode = new List<string>();
                        _arrAddressesLatitude = new List<string>();
                        _arrAddressesLongitude = new List<string>();
                        _arrAddressesURPN = new List<string>();
                        _arrAddressesBuilding = new List<string>();

                        //MessageBox.Show(String.Format("Num adresses: {0}", _addresses.Count()));
                        int count = 0;
                        foreach (dynamic address in _addresses)
                        {
                            if ((string)address["Status_Code"] == "101")
                            {
                                return -1;
                            }
                            if ((string)address["Status_Code"] == "NoMatchingRecordsFound")
                            {
                                return -2;
                            }

                            _arrAddressesAddress.Add(FormatAddress(FormatAddressHouseNo((string)address["PAO_START_NUMBER"], (string)address["PAO_START_SUFFIX"]
                                                                                        , (string)address["PAO_END_NUMBER"], (string)address["PAO_END_SUFFIX"])
                                                                                        , (string)address["STREET_DESCRIPTION"]));
                            _arrAddressesCity.Add((string)address["TOWN_NAME"]);
                            _arrAddressesPostalCode.Add((string)address["POSTCODE_LOCATOR"]);
                            _arrAddressesLatitude.Add((string)address["Y_COORDINATE"]);
                            _arrAddressesLongitude.Add((string)address["X_COORDINATE"]);

                            if (String.IsNullOrEmpty((string)address["PAO_TEXT"]))
                                _arrAddressesBuilding.Add("");
                            else
                                _arrAddressesBuilding.Add((string)address["PAO_TEXT"]);

                            if (String.IsNullOrEmpty((string)address["UPRN"]))
                                _arrAddressesURPN.Add("");
                            else
                                _arrAddressesURPN.Add((string)address["UPRN"]);

                            _arrAddressesDescription.Add(FormatAddressDescription(count));
                            count++;
                        }
                        //MessageBox.Show(String.Format("Num adresses in array: {0}", _arrAddressesDescription.Count));
                        return _arrAddressesDescription.Count;
                    }
                }
            }
            catch (WebException e)
            {
                if (e.Status == WebExceptionStatus.ProtocolError)
                {
                    //Debug.WriteLine("Status Code : {0}", ((HttpWebResponse)e.Response).StatusCode);
                    MessageBox.Show("DoSearch: Status Description : {0}", ((HttpWebResponse)e.Response).StatusDescription);
                }
            }
            return 0;
        }

        //********************************************************************************************************
        //********************************************************************************************************
        public static int UKABDoSearchUPRN(string uprnSearchFor)
        {
            //MessageBox.Show(String.Format("UKABDoSearchUPRN: {0}", uprnSearchFor));
 
            _latestURL = string.Empty;
 
            string url = string.Empty;
            if (uprnSearchFor != string.Empty)
            {
                if (url == string.Empty)
                    url = string.Format("?Data.UPRN={0}", uprnSearchFor);
                else
                    url = url + string.Format("&Data.UPRN={0}", uprnSearchFor);
            }
            else
                return 0;

            //MessageBox.Show(String.Format("url: {0}", url));
            _latestURL = url;

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(string.Format("{0}{1}", _serverURL, url));

            request.Method = "GET";
            request.ContentLength = 0;
            request.ContentType = "text/xml";
            if (_userName != "")
                request.Credentials = new NetworkCredential(_userName, _passWord);

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
                        //MessageBox.Show("Reading response");
                        string responseBody = reader.ReadToEnd();
                        //MessageBox.Show(String.Format("Response: {0}", responseBody.Substring(0, 20)));

                        //MessageBox.Show("A");
                        JObject jsonResult = JObject.Parse(responseBody);
                        //MessageBox.Show("B");
                        JArray _addresses = jsonResult["Output"].Value<JArray>();
                        //MessageBox.Show("C");

                        //MessageBox.Show(String.Format("_addresses holds {0} addresses", _addresses.Count()));

                        _arrAddressesDescription = new List<string>();
                        _arrAddressesAddress = new List<string>();
                        _arrAddressesCity = new List<string>();
                        _arrAddressesPostalCode = new List<string>();
                        _arrAddressesLatitude = new List<string>();
                        _arrAddressesLongitude = new List<string>();
                        _arrAddressesURPN = new List<string>();
                        _arrAddressesBuilding = new List<string>();

                        //MessageBox.Show(String.Format("Num adresses: {0}", _addresses.Count()));
                        int count = 0;
                        foreach (dynamic address in _addresses)
                        {
                            if ((string)address["Status_Code"] == "101")
                            {
                                return -1;
                            }
                            if ((string)address["Status_Code"] == "NoMatchingRecordsFound")
                            {
                                return -2;
                            }

                            _arrAddressesAddress.Add(FormatAddress(FormatAddressHouseNo((string)address["PAO_START_NUMBER"], (string)address["PAO_START_SUFFIX"]
                                                                                        , (string)address["PAO_END_NUMBER"], (string)address["PAO_END_SUFFIX"])
                                                                                        , (string)address["STREET_DESCRIPTION"]));
                            _arrAddressesCity.Add((string)address["TOWN_NAME"]);
                            _arrAddressesPostalCode.Add((string)address["POSTCODE_LOCATOR"]);
                            _arrAddressesLatitude.Add((string)address["Y_COORDINATE"]);
                            _arrAddressesLongitude.Add((string)address["X_COORDINATE"]);

                            if (String.IsNullOrEmpty((string)address["PAO_TEXT"]))
                                _arrAddressesBuilding.Add("");
                            else
                                _arrAddressesBuilding.Add((string)address["PAO_TEXT"]);

                            if (String.IsNullOrEmpty((string)address["UPRN"]))
                                _arrAddressesURPN.Add("");
                            else
                                _arrAddressesURPN.Add((string)address["UPRN"]);

                            _arrAddressesDescription.Add(FormatAddressDescription(count));
                            count++;
                        }
                        //MessageBox.Show(String.Format("Num adresses in array: {0}", _arrAddressesDescription.Count));
                        return _arrAddressesDescription.Count;
                    }
                }
            }
            catch (WebException e)
            {
                if (e.Status == WebExceptionStatus.ProtocolError)
                {
                    //Debug.WriteLine("Status Code : {0}", ((HttpWebResponse)e.Response).StatusCode);
                    MessageBox.Show("UKABDoSearchUPRN: Status Description : {0}", ((HttpWebResponse)e.Response).StatusDescription);
                }
            }
            return 0;
        }

        //********************************************************************************************************
        //********************************************************************************************************
        public static void SaveCredentials()
        {
            // Data to protect. Convert a string to a byte[] using Encoding.UTF8.GetBytes().
            byte[] plaintext = Encoding.UTF8.GetBytes(_passWord);
            byte[] ciphertext = ProtectedData.Protect(plaintext, _aditionalEntropy, DataProtectionScope.CurrentUser);
        }

        //********************************************************************************************************
        //********************************************************************************************************
        public static void LoadCredentials()
        {
            byte[] plaintext = Encoding.UTF8.GetBytes(_passWord);
            byte[] ciphertext = ProtectedData.Unprotect(plaintext, _aditionalEntropy, DataProtectionScope.CurrentUser);
        }

        //********************************************************************************************************
        //********************************************************************************************************
        public static string GetServerURL()
        {
            return _serverURL;
        }
        //********************************************************************************************************
        //********************************************************************************************************
        public static string GetLatestURL()
        {
            return _latestURL;
        }

        //********************************************************************************************************
        //********************************************************************************************************
        public static int GetAddressDescriptions(string[] arrAddresses)
        {
            int count = 0;
            foreach (string text in _arrAddressesDescription)
            {
                arrAddresses[count] = text;
                count++;
            }

            return arrAddresses.GetLength(0);
        }

        //********************************************************************************************************
        //********************************************************************************************************
        public static string GetAddressAttribute(int element, string attributeName)
        {
              return _addresses[element][attributeName];
        }

        //********************************************************************************************************
        //********************************************************************************************************
        public static string GetAddressDetails(   int element
                                                , ref string address, ref string building
                                                , ref string postal, ref string city, ref string UPRN)
        {
            building = _arrAddressesBuilding[element];
            address = _arrAddressesAddress[element];
            postal = _arrAddressesPostalCode[element];
            city = _arrAddressesCity[element];
            UPRN = _arrAddressesURPN[element];

            return FormatAddressDescription(element);
        }

        //********************************************************************************************************
        //********************************************************************************************************
        public static string GetAddressX(int element)
        {
            return _arrAddressesLongitude[element];
        }

        //********************************************************************************************************
        //********************************************************************************************************
        public static string GetAddressY(int element)
        {
            return _arrAddressesLatitude[element];
        }

        //********************************************************************************************************
        //********************************************************************************************************
        static string FormatAddressHouseNo(string startNo, string startSuffix, string endNo, string endSuffix)
        {
            string houseNo = "";

            try
            {

                if (!String.IsNullOrEmpty(startNo + startSuffix))
                    houseNo = string.Format("{0}{1}", startNo, startSuffix);

                if (!String.IsNullOrEmpty(endNo + endSuffix))
                {
                    if (houseNo == "")
                        houseNo = endNo + endSuffix;
                    else
                        houseNo = string.Format("{0} - {1}{2}", houseNo, endNo, endSuffix);
                }
            }
            catch (WebException e)
            {
                if (e.Status == WebExceptionStatus.ProtocolError)
                {
                    //Debug.WriteLine("Status Code : {0}", ((HttpWebResponse)e.Response).StatusCode);
                    MessageBox.Show("FormatAddressHouseNo: Status Description : {0}", ((HttpWebResponse)e.Response).StatusDescription);
                }
            }

            return houseNo;
        }

        //********************************************************************************************************
        //********************************************************************************************************
        static string FormatAddress(string houseNo, string roadName)
        {
            string address = "";

            try
            {
                if (!String.IsNullOrEmpty(houseNo))
                    address = houseNo;

                if (!String.IsNullOrEmpty(roadName))
                {
                    if (address == "")
                        address = roadName;
                    else
                        address = string.Format("{0} {1}", address, roadName);
                }
            }
            catch (WebException e)
            {
                if (e.Status == WebExceptionStatus.ProtocolError)
                {
                    //Debug.WriteLine("Status Code : {0}", ((HttpWebResponse)e.Response).StatusCode);
                    MessageBox.Show("FormatAddress: Status Description : {0}", ((HttpWebResponse)e.Response).StatusDescription);
                }
            }

            return address;
        }


        //********************************************************************************************************
        //********************************************************************************************************
        static string FormatAddressDescription(int element)
        {
            string desc = "";

            try
            {
                if (!String.IsNullOrEmpty(_arrAddressesBuilding[element]))
                    desc = _arrAddressesBuilding[element];

                if (!String.IsNullOrEmpty(_arrAddressesAddress[element]))
                {
                    if (!String.IsNullOrEmpty(desc))
                        desc = string.Format("{0}, {1}", desc, _arrAddressesAddress[element]);
                    else
                        desc = string.Format("{0}", _arrAddressesAddress[element]);
                }

                if (!String.IsNullOrEmpty(_arrAddressesPostalCode[element]))
                {
                    if (!String.IsNullOrEmpty(desc))
                        desc = string.Format("{0}, {1}", desc, _arrAddressesPostalCode[element]);
                    else
                        desc = string.Format("{0}", _arrAddressesPostalCode[element]);
                }

                if (!String.IsNullOrEmpty(_arrAddressesCity[element]))
                {
                    if (!String.IsNullOrEmpty(desc))
                        desc = string.Format("{0} {1}", desc, _arrAddressesCity[element]);
                    else
                        desc = string.Format("{0}", _arrAddressesCity[element]);
                }

                if (!String.IsNullOrEmpty(_arrAddressesURPN[element]))
                {
                    if (!String.IsNullOrEmpty(desc))
                        desc = string.Format("{0}, UPRN: {1}", desc, _arrAddressesURPN[element]);
                    else
                        desc = string.Format("UPRN: {0}", _arrAddressesURPN[element]);
                }
            }
            catch (WebException e)
            {
                if (e.Status == WebExceptionStatus.ProtocolError)
                {
                    //Debug.WriteLine("Status Code : {0}", ((HttpWebResponse)e.Response).StatusCode);
                    MessageBox.Show("FormatAddressDescription: Status Description : {0}", ((HttpWebResponse)e.Response).StatusDescription);
                }
            }

            return desc;
        }
    }

    //********************************************************************************************************
    //********************************************************************************************************
    //********************************************************************************************************
    //********************************************************************************************************
    public class AWSAddressDK
    {
        private static HttpClient _client = new HttpClient();
        private static bool _hasBeenInitiated = false;
        private static List<string> _arrAddressesDescription = null;
        private static List<string> _arrRoadDescription = null;
        private static dynamic _addresses = null;
        private static dynamic _roads = null;
        private static string _latestURL = "";

        //********************************************************************************************************
        //********************************************************************************************************
        public static void Initiate(int timeout, string serverURL)
        {
            if (_hasBeenInitiated == false)
            {
                _client.Timeout = new TimeSpan(0, timeout, 0); // 20 minutter til store downloads
                if (serverURL == "")
                    _client.BaseAddress = new Uri("http://dawa.aws.dk/");
                else
                    _client.BaseAddress = new Uri(serverURL);

                _hasBeenInitiated = true;
            }
        }

        //********************************************************************************************************
        //********************************************************************************************************
        public static int FindRoads(string roadNameSearchFor, string postalSearchFor, string municipalitySearchFor)
        {
            try
            {
                string url = "";
                url = string.Format("vejnavne/autocomplete?q={0}", roadNameSearchFor);

                if (postalSearchFor != "")
                {
                    url = url + string.Format("&postnr={0}", postalSearchFor);
                }
                if (municipalitySearchFor != "")
                {
                    url = url + string.Format("&kommunenr={0}", municipalitySearchFor);
                }

                _latestURL = url;
                HttpResponseMessage response = _client.GetAsync(url).Result;
                response.EnsureSuccessStatusCode();
                string responseBody = response.Content.ReadAsStringAsync().Result;

                _roads = JArray.Parse(responseBody);

                _arrRoadDescription = new List<string>();

                foreach (dynamic road in _roads)
                {
                    _arrRoadDescription.Add(road.tekst);
                }
            }
            catch (HttpRequestException e)
            {
                MessageBox.Show(string.Format("Error :{0} ", e.Message), "Error in Assembly", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return _arrRoadDescription.Count;
        }

        //********************************************************************************************************
        //********************************************************************************************************
        public static int DoSearch(string addressSearchFor, string postalSearchFor, string municipalitySearchFor)
        {
            try
            {
                string url = "";
                //if (addressSearchFor != "")
                url = string.Format("adresser?vejnavn={0}", addressSearchFor);
                //url = string.Format("adresser?per_side=100&vejnavn={0}", addressSearchFor);
                //url = string.Format("adresser/autocomplete?per_side=100&q={0}", addressSearchFor);

                if (postalSearchFor != "")
                {
                //    if (url == "")
                //        url = string.Format("adresser?postnr={0}", postalSearchFor);
                //    else
                        url = url + string.Format("&postnr={0}", postalSearchFor);
                }
                if (municipalitySearchFor != "")
                {
                //    if (url == "")
                //        url = string.Format("adresser?kommunenr={0}", municipalitySearchFor);
                //    else
                        url = url + string.Format("&kommunenr={0}", municipalitySearchFor);
                }

                _latestURL = url;
                HttpResponseMessage response = _client.GetAsync(url).Result;
                response.EnsureSuccessStatusCode();
                string responseBody = response.Content.ReadAsStringAsync().Result;

                _addresses = JArray.Parse(responseBody);

                _arrAddressesDescription = new List<string>();

                foreach (dynamic address in _addresses)
                {
                    _arrAddressesDescription.Add(formatAddressDescription(address));
                }
            }
            catch (HttpRequestException e)
            {
                MessageBox.Show(string.Format("Error :{0} ", e.Message), "Error in Assembly", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return _arrAddressesDescription.Count;
        }
        //********************************************************************************************************
        //********************************************************************************************************
        public static string GetLatestURL()
        {
            return _latestURL;
        }

        //********************************************************************************************************
        //********************************************************************************************************
        public static int GetRoadDesriptions(string[] arrRoads)
        {
            int count = 0;
            foreach (string text in _arrRoadDescription)
            {
                arrRoads[count] = text;
                count++;
            }

            return arrRoads.GetLength(0);
        }

        //********************************************************************************************************
        //********************************************************************************************************
        public static int GetAddressDescriptions(string[] arrAddresses)
        {
            int count = 0;
            foreach (string text in _arrAddressesDescription)
            {
                arrAddresses[count] = text;
                count++;
            }

            return arrAddresses.GetLength(0);
        }

        //********************************************************************************************************
        //********************************************************************************************************
        public static string GetAddressAttribute(int element, string attributeName)
        {
            return _addresses[element].attributeName;
        }


        //********************************************************************************************************
        //********************************************************************************************************
        static string formatAddressDescription(dynamic address)
        {
            return string.Format("{0}", address.tekst);
        }    

        //********************************************************************************************************
        //********************************************************************************************************
        static string formatAccessAddressDescription(dynamic address)
        {
            if (address.etage == "")
            {
                if (address.dør == "")
                    return string.Format("{0} {1}, {2} {3}", address.adgangsadresse.vejstykke.navn, address.adgangsadresse.husnr, address.adgangsadresse.postnummer.nr, address.adgangsadresse.postnummer.navn);
                else
                    return string.Format("{0} {1} {2}, {3} {4}", address.adgangsadresse.vejstykke.navn, address.adgangsadresse.husnr, address.dør, address.adgangsadresse.postnummer.nr, address.adgangsadresse.postnummer.navn);
            }
            else
            {
                if (address.dør == "")
                    return string.Format("{0} {1} {2}, {3} {4}", address.adgangsadresse.vejstykke.navn, address.adgangsadresse.husnr, address.etage, address.adgangsadresse.postnummer.nr, address.adgangsadresse.postnummer.navn);
                else
                    return string.Format("{0} {1} {2} {3}, {4} {5}", address.adgangsadresse.vejstykke.navn, address.adgangsadresse.husnr, address.etage, address.dør, address.adgangsadresse.postnummer.nr, address.adgangsadresse.postnummer.navn);
            }
        }

        //********************************************************************************************************
        //********************************************************************************************************
        static string formatRoadDescription(dynamic road)
        {
            return string.Format("{0}", road.tekst);
            //return string.Format("{0}", road.vejnavn.navn);
        }

    }

}
