using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using HtmlAgilityPack;
using Newtonsoft.Json;
using OfficeOpenXml;

public partial class _Default : System.Web.UI.Page
{

    public List<Fiyat> ListFiyat
    {
        get
        {
            if (Session["userSession_ListFiyat"] == null)
            {
                return new List<Fiyat>();
            }
            else
            {
                return Session["userSession_ListFiyat"] as List<Fiyat>;
            }
        }
        set
        {
            Session["userSession_ListFiyat"] = value;
        }
    }

    public List<City> ListCity { get; set; }


    protected void Page_Load(object sender, EventArgs e)
    {
        ListCity = new List<City>();

        string strCity = "[{value:'1',label:'Adana'},{value:'2',label:'Adıyaman'},{value:'3',label:'Afyonkarahisar'},{value:'4',label:'Ağrı'},{value:'68',label:'Aksaray'},{value:'5',label:'Amasya'},{value:'6',label:'Ankara'},{value:'7',label:'Antalya'},{value:'75',label:'Ardahan'},{value:'8',label:'Artvin'},{value:'9',label:'Aydın'},{value:'10',label:'Balıkesir'},{value:'74',label:'Bartın'},{value:'72',label:'Batman'},{value:'69',label:'Bayburt'},{value:'11',label:'Bilecik'},{value:'12',label:'Bingöl'},{value:'13',label:'Bitlis'},{value:'14',label:'Bolu'},{value:'15',label:'Burdur'},{value:'16',label:'Bursa'},{value:'17',label:'Çanakkale'},{value:'18',label:'Çankırı'},{value:'19',label:'Çorum'},{value:'20',label:'Denizli'},{value:'21',label:'Diyarbakır'},{value:'81',label:'Düzce'},{value:'22',label:'Edirne'},{value:'23',label:'Elazığ'},{value:'24',label:'Erzincan'},{value:'25',label:'Erzurum'},{value:'26',label:'Eskişehir'},{value:'27',label:'Gaziantep'},{value:'28',label:'Giresun'},{value:'29',label:'Gümüşhane'},{value:'30',label:'Hakkari'},{value:'31',label:'Hatay'},{value:'76',label:'Iğdır'},{value:'32',label:'Isparta'},{value:'34',label:'İstanbul'},{value:'35',label:'İzmir'},{value:'46',label:'Kahramanmaraş'},{value:'78',label:'Karabük'},{value:'70',label:'Karaman'},{value:'36',label:'Kars'},{value:'37',label:'Kastamonu'},{value:'38',label:'Kayseri'},{value:'79',label:'Kilis'},{value:'71',label:'Kırıkkale'},{value:'39',label:'Kırklareli'},{value:'40',label:'Kırşehir'},{value:'41',label:'Kocaeli'},{value:'42',label:'Konya'},{value:'43',label:'Kütahya'},{value:'44',label:'Malatya'},{value:'45',label:'Manisa'},{value:'47',label:'Mardin'},{value:'33',label:'Mersin'},{value:'48',label:'Muğla'},{value:'49',label:'Muş'},{value:'50',label:'Nevşehir'},{value:'51',label:'Niğde'},{value:'52',label:'Ordu'},{value:'80',label:'Osmaniye'},{value:'53',label:'Rize'},{value:'54',label:'Sakarya'},{value:'55',label:'Samsun'},{value:'63',label:'Şanlıurfa'},{value:'56',label:'Siirt'},{value:'57',label:'Sinop'},{value:'73',label:'Şırnak'},{value:'58',label:'Sivas'},{value:'59',label:'Tekirdağ'},{value:'60',label:'Tokat'},{value:'61',label:'Trabzon'},{value:'62',label:'Tunceli'},{value:'64',label:'Uşak'},{value:'65',label:'Van'},{value:'77',label:'Yalova'},{value:'66',label:'Yozgat'},{value:'67',label:'Zonguldak'}]";

        ListCity = JsonConvert.DeserializeObject<List<City>>(strCity);

        if (!Page.IsPostBack)
        {
            ddlCity.DataTextField = "Label";
            ddlCity.DataValueField = "Value";
            ddlCity.DataSource = ListCity;
            ddlCity.DataBind();

            //ddlCity.Items.Insert(0,new ListItem("Tümü","-1"));

        }

    }

    protected void btnTumIllerdeSorgula_OnClick(object sender, EventArgs e)
    {
        EpdkSorgula("-1");

    }

    public void EpdkSorgula(string ilValue)
    {

        List<City> listCityToQuery = new List<City>();
        if (ilValue.Equals("-1"))
        {
            // tüm iller sorgulanır
            listCityToQuery.AddRange(ListCity);
        }
        else
        {
            listCityToQuery.Add(ListCity.Find(r=>r.Value==Convert.ToInt32(ilValue)));
        }

        string epdkPage = "http://www.epdk.org.tr/Themes/Theme_002/Pages/PetrolBayiFiyatlar.aspx";
        string prm1 = ""; //"1"; // il
        string prm2 = ""; // ilçe
        string prm3 = "6074"; // dağıtıcı 6074: opet sunpet
        string prm4 = ""; // bayi
        string queryAddressFormat = "http://www.epdk.org.tr/Services/Contents.ashx?Query=petrolprice&1_={0}&2_={1}&3_={2}&4_={3}";
        string queryAddress = "";
        //var html = new HtmlDocument();
        //html.LoadHtml(new WebClient().DownloadString(epdkPage));
        //WebClient webClient = new WebClient();
        ////html.Load(webClient.DownloadString(epdkPage));
        //string page1 = webClient.DownloadString(epdkPage);
        //queryAddressFormat = string.Format(queryAddressFormat, prm1, prm2, prm3, prm4);
        //string q1 = webClient.DownloadString(queryAddressFormat);
        ////WebResponse webResponse = webClient.ResponseHeaders

        CookieContainer cookieContainer = new CookieContainer();
        CookieCollection cookieCollection = new CookieCollection();

        HttpWebRequest request = WebRequest.Create(new Uri(epdkPage)) as HttpWebRequest;
        request.CookieContainer = cookieContainer;

        HttpWebResponse response = request.GetResponse() as HttpWebResponse;
        cookieCollection = response.Cookies; // capture the cookies from the response

        List<Fiyat> listFiyat = new List<Fiyat>();

        string html = "";
        foreach (City city in listCityToQuery)
        {
            Thread.Sleep(new TimeSpan(0,0,3)); // saniye arayla sorgulama

            prm1 = city.Value.ToString(); //i.ToString(); // il bilgisi
            queryAddress = string.Format(queryAddressFormat, prm1, prm2, prm3, prm4);

            request = (HttpWebRequest)WebRequest.Create(new Uri(queryAddress));
            request.CookieContainer = cookieContainer;
            request.CookieContainer.Add(cookieCollection); // add cookies from the previous response to the new request
            request.Accept = "text/html";
            request.Headers.Add("X-Requested-With", "XMLHttpRequest");
            request.UserAgent =
                "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.58 Safari/537.36";
            request.Referer = "http://www.epdk.org.tr/Themes/Theme_002/Pages/PetrolBayiFiyatlar.aspx";
            //request.Headers.Add(HttpRequestHeader.AcceptEncoding,"gzip, deflate, sdch");
            request.Headers.Add(HttpRequestHeader.AcceptLanguage, "tr-TR,tr;q=0.8,en-US;q=0.6,en;q=0.4");


            response = (HttpWebResponse)request.GetResponse();
            //cookieCollection = response.Cookies; // capture the cookies from the response

            string responseString = String.Empty;
            using (StreamReader streamReader = new StreamReader(response.GetResponseStream()))
            {
                responseString = streamReader.ReadToEnd();
            }

            html += (responseString);

            HtmlDocument htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(string.Format("<html><body>{0}</body></html>", responseString));
            HtmlNode tableHtmlNode = htmlDocument.DocumentNode.SelectSingleNode("//table");
            HtmlNodeCollection trHtmlNodeCollection = tableHtmlNode.SelectNodes("//tr[@class='row ']");

            foreach (HtmlNode trHtmlNode in trHtmlNodeCollection)
            {
                Fiyat fiyat = new Fiyat();
                fiyat.Il = city.Label;

                List<HtmlNode> tdHtmlNodeCollection = trHtmlNode.Descendants().Where(r => r.Name.Equals("td")).ToList();
                for (int j = 0; j < 6; j++)
                {
                    HtmlNode tdHtmlNode = tdHtmlNodeCollection[j];
                    if (j == 0)
                    {
                        fiyat.Bayi = tdHtmlNode.InnerText;
                    }
                    else if (j == 1)
                    {
                        fiyat.Benzin95Oktan = string.IsNullOrEmpty(tdHtmlNode.InnerText)
                            ? (decimal?)null
                            : Convert.ToDecimal(tdHtmlNode.InnerText);
                        if (tdHtmlNode.Attributes["title"] != null)
                        {
                            fiyat.Benzin95OktanGecerlilikTarihi = string.IsNullOrEmpty(tdHtmlNode.Attributes["title"].Value.Replace("Geçerlilik Tarihi : ", ""))
                            ? (DateTime?)null
                            : Convert.ToDateTime(tdHtmlNode.Attributes["title"].Value.Replace("Geçerlilik Tarihi : ", ""));
                        }

                    }
                    else if (j == 2)
                    {
                        fiyat.Motorin = string.IsNullOrEmpty(tdHtmlNode.InnerText)
                            ? (decimal?)null
                            : Convert.ToDecimal(tdHtmlNode.InnerText);
                        if (tdHtmlNode.Attributes["title"] != null)
                        {
                            fiyat.Benzin95OktanGecerlilikTarihi = string.IsNullOrEmpty(tdHtmlNode.Attributes["title"].Value.Replace("Geçerlilik Tarihi : ", ""))
                            ? (DateTime?)null
                            : Convert.ToDateTime(tdHtmlNode.Attributes["title"].Value.Replace("Geçerlilik Tarihi : ", ""));
                        }
                    }
                    else if (j == 3)
                    {
                        fiyat.BenzinDiger = string.IsNullOrEmpty(tdHtmlNode.InnerText)
                            ? (decimal?)null
                            : Convert.ToDecimal(tdHtmlNode.InnerText);
                        if (tdHtmlNode.Attributes["title"] != null)
                        {
                            fiyat.BenzinDigerGecerlilikTarihi = string.IsNullOrEmpty(tdHtmlNode.Attributes["title"].Value.Replace("Geçerlilik Tarihi : ", ""))
                            ? (DateTime?)null
                            : Convert.ToDateTime(tdHtmlNode.Attributes["title"].Value.Replace("Geçerlilik Tarihi : ", ""));
                        }
                    }
                    else if (j == 4)
                    {
                        fiyat.MotorinDiger = string.IsNullOrEmpty(tdHtmlNode.InnerText)
                            ? (decimal?)null
                            : Convert.ToDecimal(tdHtmlNode.InnerText);
                        if (tdHtmlNode.Attributes["title"] != null)
                        {
                            fiyat.MotorinDigerGecerlilikTarihi = string.IsNullOrEmpty(tdHtmlNode.Attributes["title"].Value.Replace("Geçerlilik Tarihi : ", ""))
                            ? (DateTime?)null
                            : Convert.ToDateTime(tdHtmlNode.Attributes["title"].Value.Replace("Geçerlilik Tarihi : ", ""));
                        }
                    }
                    else if (j == 5)
                    {
                        fiyat.Dagitici = tdHtmlNode.InnerText;
                    }

                }

                listFiyat.Add(fiyat);

            } // end of foreach (HtmlNode trHtmlNode in trHtmlNodeCollection)...


        }

        //divContent1.InnerHtml = HttpUtility.HtmlEncode(html);
        divContent1.InnerHtml = html;

        lblRecordCount.Text = listFiyat.Count.ToString();
        ListFiyat = listFiyat;
        gv1.DataSource = listFiyat;
        gv1.DataBind();


        // kaynaklar:
        // How to parse HttpWebResponse.Headers.Keys for a Set-Cookie session id returned: http://stackoverflow.com/questions/1055853/how-to-parse-httpwebresponse-headers-keys-for-a-set-cookie-session-id-returned
        // Using HtmlAgilityPack to GET and POST web forms: http://refactoringaspnet.blogspot.com.tr/2010/04/using-htmlagilitypack-to-get-and-post.html
        // http://stackoverflow.com/questions/17982131/cookies-dropped-lost-on-httpwebrequest-getreponse


    }


    public class Fiyat
    {
        public string Il { get; set; }
        public string Bayi { get; set; }
        public decimal? Benzin95Oktan { get; set; }
        public DateTime? Benzin95OktanGecerlilikTarihi { get; set; }
        public decimal? Motorin { get; set; }
        public DateTime? MotorinGecerlilikTarihi { get; set; }
        public decimal? BenzinDiger { get; set; }
        public DateTime? BenzinDigerGecerlilikTarihi { get; set; }
        public decimal? MotorinDiger { get; set; }
        public DateTime? MotorinDigerGecerlilikTarihi { get; set; }
        public string Dagitici { get; set; }
    }

    public class City
    {
        public int Value { get; set; }
        public string Label { get; set; }
    }


    protected void btnSorgula_OnClick(object sender, EventArgs e)
    {
        EpdkSorgula(ddlCity.SelectedItem.Value);

    }


    protected void btnExcel_OnClick(object sender, EventArgs e)
    {
        if (ListFiyat == null || ListFiyat.Count == 0)
        {
            return;    
        }

        using (ExcelPackage excelPackage = new ExcelPackage())
        {
            ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.Add("Report");


            // başlık kısmı
            excelWorksheet.Cells[1, 1].Value = "Il"; // Il
            excelWorksheet.Cells[1, 2].Value = "Bayi"; // Bayi
            excelWorksheet.Cells[1, 3].Value = "Benzin95Oktan"; // Benzin95Oktan
            excelWorksheet.Cells[1, 4].Value = "Benzin95OktanGecerlilikTarihi"; // Benzin95OktanGecerlilikTarihi
            excelWorksheet.Cells[1, 5].Value = "Motorin"; // Motorin
            excelWorksheet.Cells[1, 6].Value = "MotorinGecerlilikTarihi"; // MotorinGecerlilikTarihi
            excelWorksheet.Cells[1, 7].Value = "BenzinDiger"; //BenzinDiger
            excelWorksheet.Cells[1, 8].Value = "BenzinDigerGecerlilikTarihi"; // BenzinDigerGecerlilikTarihi
            excelWorksheet.Cells[1, 9].Value = "MotorinDiger"; // MotorinDiger
            excelWorksheet.Cells[1, 10].Value = "MotorinDigerGecerlilikTarihi"; // MotorinDigerGecerlilikTarihi
            excelWorksheet.Cells[1, 11].Value = "Dagitici"; // Dagitici


            excelWorksheet.Cells[1, 1, 1, 100].Style.Font.Bold = true;
            //excelWorksheet.Cells[1, 1, 1, 11].Style.ShrinkToFit = false;
            excelWorksheet.Cells[excelWorksheet.Dimension.Address.ToString()].AutoFitColumns();


            //if (dt == null)
            //    return;

            // veri kaynağından doldurma
            int indexRow = 2;
            foreach (Fiyat item in ListFiyat) //foreach (DataRow item in dt.Rows)
            {

                string strRowValue = string.Empty;
                if (item.Il!=null) strRowValue = item.Il;
                excelWorksheet.Cells[indexRow, 1].Value = strRowValue;

                strRowValue = string.Empty;
                if (item.Bayi!=null) strRowValue = item.Bayi;
                excelWorksheet.Cells[indexRow, 2].Value = strRowValue;

                strRowValue = string.Empty;
                if (item.Benzin95Oktan!=null) strRowValue = item.Benzin95Oktan.Value.ToString();
                excelWorksheet.Cells[indexRow, 3].Value = strRowValue;

                strRowValue = string.Empty;
                if (item.Benzin95OktanGecerlilikTarihi!=null) strRowValue = item.Benzin95OktanGecerlilikTarihi.Value.ToString();
                excelWorksheet.Cells[indexRow, 4].Value = strRowValue;

                strRowValue = string.Empty;
                if (item.Motorin!=null) strRowValue = item.Motorin.Value.ToString();
                excelWorksheet.Cells[indexRow, 5].Value = strRowValue;

                strRowValue = string.Empty;
                if (item.MotorinGecerlilikTarihi != null) strRowValue = item.MotorinGecerlilikTarihi.Value.ToString();
                excelWorksheet.Cells[indexRow, 6].Value = strRowValue;

                strRowValue = string.Empty;
                if (item.BenzinDiger != null) strRowValue = item.BenzinDiger.Value.ToString();
                excelWorksheet.Cells[indexRow, 7].Value = strRowValue;

                strRowValue = string.Empty;
                if (item.BenzinDigerGecerlilikTarihi != null) strRowValue = item.BenzinDigerGecerlilikTarihi.Value.ToString();
                excelWorksheet.Cells[indexRow, 8].Value = strRowValue;

                strRowValue = string.Empty;
                if (item.MotorinDiger != null) strRowValue = item.MotorinDiger.Value.ToString();
                excelWorksheet.Cells[indexRow, 9].Value = strRowValue;

                strRowValue = string.Empty;
                if (item.MotorinDigerGecerlilikTarihi != null) strRowValue = item.MotorinDigerGecerlilikTarihi.Value.ToString();
                excelWorksheet.Cells[indexRow, 10].Value = strRowValue;

                strRowValue = string.Empty;
                if (item.Dagitici != null) strRowValue = item.Dagitici;
                excelWorksheet.Cells[indexRow, 11].Value = strRowValue;

                indexRow++;

            } // end of foreach


            Response.Clear();
            Response.ClearHeaders();
            Response.ClearContent();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition",
                               "attachment;filename=Report_" + DateTime.Now.ToString("yyyyMMddHHmmss") +
                               ".xlsx");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            //Response.Write(document);
            Response.BinaryWrite(excelPackage.GetAsByteArray());
            Response.End();


        } // end of using (ExcelPackage excelPackage = new ExcelPackage())

    }


} // end of class