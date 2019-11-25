using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;    
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

using Spire;
using Spire.Pdf.Actions;
using Spire.Pdf.Annotations.Appearance;
using Spire.Pdf.Attachments;
using Spire.Pdf.AutomaticFields;
using Spire.Pdf.Bookmarks;
using Spire.Pdf.ColorSpace;
using Spire.Pdf.Exceptions;
using Spire.Pdf.Fields;

using Spire.Pdf.General;
using Spire.Pdf.Graphics;
using Spire.Pdf.Grid;
using Spire.Pdf.Print;
using Spire.Pdf.IO;

using Spire.Pdf.Exporting.Text;
using Spire.Pdf.Lists;
using Spire.Pdf.Widget;
using OfficeOpenXml;



namespace Sales_Tax_Aggregator
{
    public partial class Sales_Tax_Aggregator : Form
    {
        public Sales_Tax_Aggregator()
        {
            InitializeComponent();

            Check_Sharpe();
            Init_DataTables();

            Tax_Rates_To_Grid();
            init_Date();
            Load_CA_Jurisdictions();
            Fill_Jurisdictions_With_STJs();
            Check_Files();
            Load_PP_Products();
            Load_CFN_Products();
            Load_CFN_Sites();
            Load_CFN_Counties();
            Init_Zones();
            Load_Zones_With_STJ_Rates();
            Get_All_Containers();
            Load_Calc_Settings();

            Load_CADSL_Rates();
            Load_CAMVF_Rates();
            Load_VoyCodes();

            Mail_Me();

        }

        public List<Control> containers = new List<Control>();

        // File Names

        public string salesfile = "";
        public string cpstfile = "";
        public string cardlockfile = "";
        public string purchasesfile = "";

        // Variables for Calculation

        public decimal gross_sales = 0.00m;
        public decimal diesel_Clear = 0.00m;
        public decimal diesel_SET = 0.00m;

        public decimal cpst = 0.00m;
        public decimal cpst_diesel_clear = 0.00m;
        public decimal cpst_diesel_red = 0.00m;
        public decimal cpst_farm_equip = 0.00m;
        public decimal mvf_sales = 0.00m;
        public decimal wholesale = 0.00m;

        public decimal excluded_cardlock = 0.00m;
        public decimal foreign = 0.00m;

        public float stax_rate_diesel = 0.13f;
        public float stax_rate_gasoline = 0.0225f;

        // Calculation Indicators

        public int indicator_cpst_diesel_clear = 101;
        public int indicator_cpst_diesel_dyed = 103;
        public int indicator_cpst_equip = 1;
        public int indicator_cpst_oils = 102;

        public int column_cpst_amount = 18;
        public int column_cpst_indicator = 33;


        public DataTable sales_data = new DataTable();
        public DataTable cpst_data = new DataTable();

        public int datepos1 = 0;
        public int datepos2 = 0;
        public int datepos3 = 0;

        public List<int> CSV_Bad_Rows;

        public int column_product = 12;
        public int column_pcat = 11;
        public int indicator_pcat_diesel_clear = 4;
        public int indicator_pcat_diesel_dyed = 4;
        public int indicator_pcat_87 = 3;
        public int indicator_pcat_89 = 2;
        public int indicator_pcat_91 = 1;
        public int indicator_pcat_avg = 0;
        public int indicator_pcat_last_MVF = 3;

        public string indicator_product_diesel_clear = "DF";
        public string indicator_product_diesel_dyed = "DFD";
        public string indicator_product_87 = "87";
        public string indicator_product_89 = "89";
        public string indicator_product_91 = "91";
        public string indicator_product_avg = "AVG";
        public string indicator_product_oils = "";
        public string indicator_product_other = "";
        public string indicator_Excluded_Cardlock = "VOYAGER PASS THROUGH";
        public int column_excluded_cardlock = 33;

        public string indicator_State_Not = "OUT OF STATE";

        public int column_quan = 18;
        public int column_taxable = 38;
        public int column_sales = 20;
        public int column_set = 35;
        public int column_price = 0;
        public int column_cost = 21;
        public int column_fet = 0;
        public int column_salestax_amount = 33;
        public int column_salesman = 0;
        public int column_salestax_zone = 31;
        public int column_invoice_num = 14;


        public int column_wholesale = 27;
        public string indicator_wholesale = "Y";
        public string indicator_retail = "N";

        public bool set_is_each = true;
        public bool fet_is_each = true;
        public bool sales_is_each = false;
        public bool cost_is_each = false;
        public bool price_is_each = true;
        public bool use_taxable = false;
        public bool use_sales = true;
        public bool use_stax_as_wholesale = false;

        public int column_pretax_rate = 36;
        public int column_pretax_amount = 0;
        public int column_salestax_rate = 37;

        public bool set_included_in_sales = true;
        public bool stax_included_in_sales = false;

        public int column_state = 40;
        public string indicator_state_ca = "CAL";

        public int exclude_wh_1 = 2;

        public int column_warehouse = 1;

        public string pdffile = @"C:\Sharpe\cdfta531a2rate.pdf";
        public string districtfile = "https://www.cdtfa.ca.gov/formspubs/cdtfa531a2.pdf";

        public string archive = @"C:\Sharpe\archive.xlsx";

        public StringBuilder rates = new StringBuilder();
        public string ratesfile = @"C:\Sharpe\districtrates.csv";

        public StringBuilder text = new StringBuilder();

        public StringBuilder texttest = new StringBuilder();
        public StringBuilder FL1codes = new StringBuilder();
        public List<string> FL1codelist = new List<string>();

        public List<string> CountyCodes = new List<string>();
        public List<string> CountyNames = new List<string>();
        public List<string> Countys = new List<string>();

        public StringBuilder fields = new StringBuilder();
        public StringBuilder fields2 = new StringBuilder();
        public StringBuilder fields_Test = new StringBuilder();

        public StringBuilder pub61text = new StringBuilder();

        public List<string> districtrates = new List<string>();

        //public DataTable locdata = new DataTable();
        public DataTable workTable = new DataTable("TaxRates");
        public DataTable pretax = new DataTable();

        public List<Locbill> locbil = new List<Locbill>();
        public List<Locbill> locbil_de = new List<Locbill>();
        public List<Locbill> locbil_de_ca = new List<Locbill>();
        public List<Locbill> locbil_df = new List<Locbill>();
        public List<Locbill> locbil_dd = new List<Locbill>();
        public List<Locbill> locbil_fdp = new List<Locbill>();
        public List<Locbill> locbil_fdo = new List<Locbill>();



        //public DataTable locbil = new DataTable();
        //public DataTable pp_products = new DataTable();

        public List<Products_PP> pp_products = new List<Products_PP>();
        public List<Products_CFN> cfn_products = new List<Products_CFN>();

        //  public DataTable cfn_products = new DataTable();

        //  public DataTable extended = new DataTable();
        public DataTable cadtax = new DataTable();
        public DataTable sales = new DataTable();
        public DataTable juris = new DataTable();
        public DataTable zones = new DataTable();

        //public DataTable PTfile = new DataTable();
        public List<PTFile> CFNData = new List<PTFile>();

        //public DataTable PTExtended = new DataTable();
        public List<PTFile> CFNDataExt = new List<PTFile>();
        //public DataTable CFNsites = new DataTable();

        public List<Sites_CFN> CFNSites = new List<Sites_CFN>();

        public List<Counties_CFN> CFNCounties = new List<Counties_CFN>();
        //public DataTable CFNcounties = new DataTable();

        //public DataTable Voyfile = new DataTable();
        public List<VoyagerData> VoyData = new List<VoyagerData>();

        public DataTable Voy_products = new DataTable();


        //public DataTable VoyCodes = new DataTable();
        public List<VoyagerCodes> VoyCodes = new List<VoyagerCodes>();

        public DataTable CATaxRates_DSL = new DataTable();
        public DataTable CATaxRates_MVF = new DataTable();
        public DataTable CATaxRates_AVG = new DataTable();

        public string Host = "";

        public string[] locfields = new string[79];


        public decimal salestaxpaiddiesel = 0.00m;
        public decimal salestaxpaiddyed = 0.00m;
        public decimal salestaxpaidgas = 0.00m;
        public decimal salestaxpaidother = 0.00m;
        public decimal salestaxpaidtotal = 0.00m;

        public decimal salestaxpaiddiesel_pp = 0.00m;
        public decimal salestaxpaiddiesel_cfn = 0.00m;
        public decimal salestaxpaiddiesel_voy = 0.00m;
        public decimal salestaxpaiddiesel_wex = 0.00m;
        public decimal salestaxpaiddyed_pp = 0.00m;
        public decimal salestaxpaiddyed_cfn = 0.00m;
        public decimal salestaxpaiddyed_voy = 0.00m;
        public decimal salestaxpaiddyed_wex = 0.00m;
        public decimal salestaxpaidgas_pp = 0.00m;
        public decimal salestaxpaidgas_cfn = 0.00m;
        public decimal salestaxpaidgas_voy = 0.00m;
        public decimal salestaxpaidgas_wex = 0.00m;
        public decimal salestaxpaidother_pp = 0.00m;
        public decimal salestaxpaidother_cfn = 0.00m;
        public decimal salestaxpaidother_voy = 0.00m;
        public decimal salestaxpaidother_wex = 0.00m;

        public decimal salestaxpaidtotal_pp = 0.00m;
        public decimal salestaxpaidtotal_cfn = 0.00m;
        public decimal salestaxpaidtotal_voy = 0.00m;
        public decimal salestaxpaidtotal_wex = 0.00m;

        public decimal disttaxpaiddiesel = 0.00m;
        public decimal disttaxpaidgas = 0.00m;
        public decimal disttaxpaiddyed = 0.00m;
        public decimal disttaxpaidother = 0.00m;
        public decimal disttaxpaidtotal = 0.00m;

        public decimal disttaxpaiddiesel_pp = 0.00m;
        public decimal disttaxpaidgas_pp = 0.00m;
        public decimal disttaxpaiddyed_pp = 0.00m;
        public decimal disttaxpaidother_pp = 0.00m;
        public decimal disttaxpaidtotal_pp = 0.00m;

        public decimal disttaxpaiddiesel_cfn = 0.00m;
        public decimal disttaxpaidgas_cfn = 0.00m;
        public decimal disttaxpaiddyed_cfn = 0.00m;
        public decimal disttaxpaidother_cfn = 0.00m;
        public decimal disttaxpaidtotal_cfn = 0.00m;

        public decimal disttaxpaiddiesel_voy = 0.00m;
        public decimal disttaxpaidgas_voy = 0.00m;
        public decimal disttaxpaiddyed_voy = 0.00m;
        public decimal disttaxpaidother_voy = 0.00m;
        public decimal disttaxpaidtotal_voy = 0.00m;

        public decimal disttaxpaiddiesel_wex = 0.00m;
        public decimal disttaxpaidgas_wex = 0.00m;
        public decimal disttaxpaiddyed_wex = 0.00m;
        public decimal disttaxpaidother_wex = 0.00m;
        public decimal disttaxpaidtotal_wex = 0.00m;

        public decimal extcostdiesel = 0.00m;
        public decimal extcostdiesel_pp = 0.00m;
        public decimal extcostdiesel_cfn = 0.00m;
        public decimal extcostdiesel_voy = 0.00m;
        public decimal extcostdiesel_wex = 0.00m;

        public decimal extcostdyed = 0.00m;
        public decimal extcostdyed_pp = 0.00m;
        public decimal extcostdyed_cfn = 0.00m;
        public decimal extcostdyed_voy = 0.00m;
        public decimal extcostdyed_wex = 0.00m;

        public decimal extcostgas = 0.00m;
        public decimal extcostgas_pp = 0.00m;
        public decimal extcostgas_cfn = 0.00m;
        public decimal extcostgas_voy = 0.00m;
        public decimal extcostgas_wex = 0.00m;

        public decimal extcostother = 0.00m;
        public decimal extcostother_pp = 0.00m;
        public decimal extcostother_cfn = 0.00m;
        public decimal extcostother_voy = 0.00m;
        public decimal extcostother_wex = 0.00m;

        public decimal extcosttotal = 0.00m;
        public decimal extcosttotal_pp = 0.00m;
        public decimal extcosttotal_cfn = 0.00m;
        public decimal extcosttotal_voy = 0.00m;
        public decimal extcosttotal_wex = 0.00m;

        public decimal extsalesdiesel = 0.00m;
        public decimal extsalesgas = 0.00m;
        public decimal extsalesother = 0.00m;
        public decimal extsalestotal = 0.00m;

        public decimal salestaxdiesel = 0.00m;
        public decimal salestaxgas = 0.00m;
        public decimal salestaxother = 0.00m;
        public decimal salestaxtotal = 0.00m;

        public decimal pretaxpaiddiesel = 0.00m;
        public decimal pretaxpaidgas = 0.00m;

        public decimal pretaxrecdiesel = 0.00m;
        public decimal pretaxrecgas = 0.00m;

        public decimal salestotal = 0.00m;
        public decimal salesdiesel = 0.00m;
        public decimal salesdyed = 0.00m;
        public decimal salesdieselother = 0.00m;
        public decimal sales87gas = 0.00m;
        public decimal sales89gas = 0.00m;
        public decimal sales91gas = 0.00m;
        public decimal salesAVG = 0.00m;
        public decimal salesgasall = 0.00m;
        public decimal salesdieselall = 0.00m;


        public decimal costtotal = 0.00m;
        public decimal costdiesel = 0.00m;
        public decimal costdyed = 0.00m;
        public decimal cost87gas = 0.00m;
        public decimal cost89gas = 0.00m;
        public decimal cost91gas = 0.00m;

        public decimal galstotal = 0.00m;
        public decimal galsdiesel = 0.00m;
        public decimal galsdyed = 0.00m;
        public decimal gals87gas = 0.00m;
        public decimal gals89gas = 0.00m;
        public decimal gals91gas = 0.00m;
        public decimal galsAVG = 0.00m;

        public decimal galstotalext = 0.00m;
        public decimal galsdieselext = 0.00m;
        public decimal galsdyedext = 0.00m;
        public decimal gals87gasext = 0.00m;
        public decimal gals89gasext = 0.00m;
        public decimal gals91gasext = 0.00m;

        public decimal galstotalext_pp = 0.00m;
        public decimal galsdieselext_pp = 0.00m;
        public decimal galsdyedext_pp = 0.00m;
        public decimal galsgasext_pp = 0.00m;
        public decimal galsotherext_pp = 0.00m;


        public decimal galstotalext_cfn = 0.00m;
        public decimal galsdieselext_cfn = 0.00m;
        public decimal galsdyedext_cfn = 0.00m;
        public decimal galsgasext_cfn = 0.00m;
        public decimal galsotherext_cfn = 0.00m;

        public decimal galstotalext_voy = 0.00m;
        public decimal galsdieselext_voy = 0.00m;
        public decimal galsdyedext_voy = 0.00m;
        public decimal galsgasext_voy = 0.00m;
        public decimal galsotherext_voy = 0.00m;

        public decimal galstotalext_wex = 0.00m;
        public decimal galsdieselext_wex = 0.00m;
        public decimal galsdyedext_wex = 0.00m;
        public decimal galsgasext_wex = 0.00m;
        public decimal galsotherext_wex = 0.00m;

        public decimal galsothext = 0.00m;
        public decimal galsoth = 0.00m;

        public string productfile = @"C:\Sharpe\PacPrideProducts.csv";
        public string DistrictRateFile = @"C:\Sharpe\districtrates.csv";
        public string Jurisdiction = @"C:\Sharpe\CAJurisdictions.csv";
        public string Zones = @"C:\Sharpe\Zones.csv";
        public string dir = @"C:\Sharpe";

        public string safedir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

        public string CFNprodfile = @"C:\Sharpe\CFNProducts.csv";
        public string CFNsitesfile = @"C:\Sharpe\Sitemaster.csv";
        public string CFNextendedsitescafile = @"C:\Sharpe\CFNExtendedSitesCA.csv";
        public string CFNcountiesfile = @"C:\Sharpe\CFNCounties.csv";

        public string VoyagerCodeFile = @"C:\Sharpe\VoyagerCodes.csv";
        public string CATaxDSLRates = @"C:\Sharpe\CATaxDSLRates.csv";
        public string CATaxMVFRates = @"C:\Sharpe\CATaxMVFRates.csv";




        public string SalesAggCalcSettings = @"C:\Sharpe\SalesAggCalcSettings.txt";



        void Check_Sharpe()
        {
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            // MessageBox.Show("Directory where files are stored and error logs are kept = "+safedir);

            try
            {
                DirectoryInfo dinfo = new DirectoryInfo(dir);

                DirectorySecurity dsec = dinfo.GetAccessControl();

                FileSecurity fileSecure = File.GetAccessControl(dir);
                StringBuilder acer = new StringBuilder();
                fileSecure.GetSecurityDescriptorSddlForm(AccessControlSections.All);

                foreach (FileSystemAccessRule ace in fileSecure.GetAccessRules(true, true, typeof(NTAccount)))
                {
                    acer.Append(ace.FileSystemRights + ":" + ' ' + ace.IdentityReference.Value + "\n");
                }

                MessageBox.Show("Sharpe Directory Access is: " + Environment.NewLine + acer.ToString());

            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show("You do not have access to Sharpe Directory. Abort" + ex.Message);
            }

        }

        void Check_Files()
        {
            if (!File.Exists(productfile))
            {
                MessageBox.Show("Missing Pac Pride Product File. Please copy the Pac Pride Product file to C:\\Sharpe\\PacPrideProducts.csv");
            }

            if (!File.Exists(Jurisdiction))
            {
                MessageBox.Show("Missing Jurisdiction File. Please copy jurisdiction file to C:\\Sharpe\\CAJurisdictions.csv");
            }

            if (!File.Exists(Zones))
            {
                MessageBox.Show("Missing Zones File. Please copy the Zones file to C:\\Sharpe\\Zones.csv");
            }

            if (!File.Exists(CFNprodfile))
            {
                MessageBox.Show("Missing CFN Product File. Please copy the CFN Product file to C:\\Sharpe\\CFNProducts.csv");
            }
            if (!File.Exists(CFNsitesfile))
            {
                MessageBox.Show("Missing CFN Sites File. Please copy the CFN Sitemaster file to C:\\Sharpe\\Sitemaster.csv");
            }


        }

        void Mail_Me()
        {

            MailAddress jim = new MailAddress("caprepaidtax@sharpecpa.com");
            MailAddress jimphone = new MailAddress("5305925046@vtext.com");
            MailAddress ryan = new MailAddress("greentraks@gmail.com");
            MailAddress jimlake = new MailAddress("jsharpe@lakeviewpetroleum.com");
            MailAddress ryanphone = new MailAddress("5306825448@txt.att.net");

            using (MailMessage mail = new MailMessage(jim, jimlake))
            {
                mail.Subject = "Sales Tax Aggregator Being Used By " + System.Environment.MachineName + " User = " + System.Environment.UserName;

                mail.Body = "Sales Tax Aggregator Being Used By " + System.Environment.MachineName + " User = " + System.Environment.UserName + Environment.NewLine +
                    " Domain Name = " + System.Environment.UserDomainName + " Folder Path = " + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + Environment.NewLine;

                string[] drives = Environment.GetLogicalDrives();

                foreach (string drive in drives)
                {
                    mail.Body = mail.Body + "Drive Available = " + drive;
                }

                SmtpClient smtp = new SmtpClient("mail.succeed.net", 587);
                smtp.Credentials = new System.Net.NetworkCredential("caprepaidtax@sharpecpa.com", "Affile1958!");
                smtp.EnableSsl = true;
                smtp.Send(mail);
            }

            using (MailMessage mail = new MailMessage(jim, jimphone))
            {
                mail.Subject = "Sales Tax Aggregator Being Used By " + System.Environment.MachineName + " User = " + System.Environment.UserName;

                mail.Body = "Sales Tax Aggregator Being Used By " + System.Environment.MachineName + " User = " + System.Environment.UserName;

                SmtpClient smtp = new SmtpClient("mail.succeed.net", 587);
                smtp.Credentials = new System.Net.NetworkCredential("casalestax@sharpecpa.com", "Affile1958!");
                smtp.EnableSsl = true;
                smtp.Send(mail);
            }

            using (MailMessage mail = new MailMessage(jim, ryanphone))
            {
                mail.Subject = "Sales Tax Aggregator Being Used By " + System.Environment.MachineName + " User = " + System.Environment.UserName;

                mail.Body = "Sales Tax Aggregator Being Used By " + System.Environment.MachineName + " User = " + System.Environment.UserName;

                SmtpClient smtp = new SmtpClient("mail.succeed.net", 587);
                smtp.Credentials = new System.Net.NetworkCredential("caprepaidtax@sharpecpa.com", "Affile1958!");
                smtp.EnableSsl = true;
                smtp.Send(mail);
            }
        }

        void Load_Textboxes_Calc()
        {
            int.TryParse(tb_CPST_Indicator_Clear.Text, out indicator_cpst_diesel_clear);
            int.TryParse(tb_CPST_Indicator_Red.Text, out indicator_cpst_diesel_dyed);
            int.TryParse(tb_CPST_Indicator_Oils.Text, out indicator_cpst_equip);
            int.TryParse(tb_CPST_Indicator_Oils.Text, out indicator_cpst_oils);
            int.TryParse(tb_Column_CPST.Text, out column_cpst_amount);
            int.TryParse(tb_Column_CPST_Indicator.Text, out column_cpst_indicator);
            int.TryParse(tb_Column_Product_Category.Text, out column_pcat);
            int.TryParse(tb_Product_Cat_DSL.Text, out indicator_pcat_diesel_clear);
            int.TryParse(tb_Product_Cat_Red.Text, out indicator_pcat_diesel_dyed);
            int.TryParse(tb_Product_Cat_87.Text, out indicator_pcat_87);
            int.TryParse(tb_Product_Cat_89.Text, out indicator_pcat_89);
            int.TryParse(tb_Product_Cat_91.Text, out indicator_pcat_91);
            int.TryParse(tb_Product_Cat_Last_MVF.Text, out indicator_pcat_last_MVF);
            int.TryParse(tb_Column_State.Text, out column_state);
            indicator_state_ca = tb_State_Indicator.Text;
            int.TryParse(tb_Warehouse_Exclude_1.Text, out exclude_wh_1);
            int.TryParse(tb_Column_Warehouse.Text, out column_warehouse);
            int.TryParse(tb_Column_Product.Text, out column_product);
            indicator_product_diesel_clear = tb_Product_Indicator_Diesel.Text;
            indicator_product_diesel_dyed = tb_Product_Indicator_Red.Text;
            indicator_product_87 = tb_Product_Indicator_87.Text;
            indicator_product_89 = tb_Product_Indicator_89.Text;
            indicator_product_91 = tb_Product_Indicator_91.Text;
            indicator_product_avg = tb_Product_Indicator_AVG.Text;
            indicator_product_oils = indicator_product_other;
            int.TryParse(tb_Column_Quantity.Text, out column_quan);
            int.TryParse(tb_Column_Taxable.Text, out column_taxable);
            int.TryParse(tb_Column_Sales.Text, out column_sales);
            int.TryParse(tb_Column_SET.Text, out column_set);
            int.TryParse(tb_Column_Price.Text, out column_price);
            int.TryParse(tb_Column_Cost.Text, out column_cost);
            int.TryParse(tb_Column_FET.Text, out column_fet);
            int.TryParse(tb_Column_Salestax_Zone.Text, out column_salestax_zone);
            int.TryParse(tb_Column_Salestax_Amount.Text, out column_salestax_amount);
            int.TryParse(tb_Column_Salestax_Rate.Text, out column_salestax_rate);
            int.TryParse(tb_Column_Salesman.Text, out column_salesman);
            int.TryParse(tb_Column_Wholesale.Text, out column_wholesale);
            indicator_wholesale = tb_Wholesale_Indicator.Text;
            set_is_each = checkBox_SET_is_Rate.Checked;
            fet_is_each = checkBox_FET_Rate.Checked;
            sales_is_each = checkBox_Sales_is_Each.Checked;
            cost_is_each = checkBox_Cost_is_Each.Checked;
            price_is_each = checkBox_Price_is_Each.Checked;
            use_taxable = checkBox_Use_Taxable.Checked;
            use_sales = checkBox_Use_Sales.Checked;
            use_stax_as_wholesale = checkBox_STAX_as_Wholesale.Checked;
            set_included_in_sales = checkBox_SET_Included_Sales.Checked;
            stax_included_in_sales = checkBox_STAX_in_Sales.Checked;
            int.TryParse(tb_Column_Pretax_Rate.Text, out column_pretax_rate);
            int.TryParse(tb_Column_Pretax_Amount.Text, out column_pretax_amount);
            int.TryParse(tb_Column_Excluded_CL.Text, out column_excluded_cardlock);
            indicator_Excluded_Cardlock = tb_Indicator_Ex_CL.Text;
            indicator_State_Not = tb_Indicator_State_Not.Text;
        }

        void Load_Calc_Settings()
        {

            if (File.Exists(SalesAggCalcSettings))
            {
                // MessageBox.Show("Got Here");

                string[] lines = File.ReadAllLines(SalesAggCalcSettings);

                //MessageBox.Show(lines.Length.ToString());

                foreach (string line in lines)
                {
                    int start = line.IndexOf("|");

                    int len = line.Length - start - 1;

                    string name = "";

                    if (line.Length > 0)
                    {
                        name = line.Substring(0, (start));
                    }

                    string value = "";


                    if (len > 0)
                    {
                        value = line.Substring((start + 1), len);
                    }

                    //MessageBox.Show(name + " " + value);

                    switch (name)
                    {
                        case "indicator_cpst_diesel_clear":
                            int.TryParse(value, out indicator_cpst_diesel_clear);
                            tb_CPST_Indicator_Clear.Text = value;
                            break;

                        case "indicator_cpst_diesel_dyed":
                            int.TryParse(value, out indicator_cpst_diesel_dyed);
                            tb_CPST_Indicator_Red.Text = value;
                            break;
                        case "indicator_cpst_equip":
                            int.TryParse(value, out indicator_cpst_equip);
                            tb_CPST_Indicator_Oils.Text = value;
                            break;
                        case "indicator_cpst_oils":
                            int.TryParse(value, out indicator_cpst_oils);
                            tb_CPST_Indicator_Oils.Text = value;
                            break;
                        case "column_cpst_amount":
                            int.TryParse(value, out column_cpst_amount);
                            tb_Column_CPST.Text = value;
                            break;
                        case "column_cpst_indicator":
                            int.TryParse(value, out column_cpst_indicator);
                            tb_Column_CPST_Indicator.Text = value;
                            break;
                        case "column_pcat":
                            int.TryParse(value, out column_pcat);
                            tb_Column_Product_Category.Text = value;
                            break;
                        case "indicator_pcat_diesel_clear":
                            int.TryParse(value, out indicator_pcat_diesel_clear);
                            tb_Product_Cat_DSL.Text = value;
                            break;
                        case "indicator_pcat_diesel_dyed":
                            int.TryParse(value, out indicator_pcat_diesel_dyed);
                            tb_Product_Cat_Red.Text = value;
                            break;
                        case "indicator_pcat_87":
                            int.TryParse(value, out indicator_pcat_87);
                            tb_Product_Cat_87.Text = value;
                            break;
                        case "indicator_pcat_89":
                            int.TryParse(value, out indicator_pcat_89);
                            tb_Product_Cat_89.Text = value;

                            break;
                        case "indicator_pcat_91":
                            int.TryParse(value, out indicator_pcat_91);
                            tb_Product_Cat_91.Text = value;
                            break;
                        case "indicator_pcat_last_MVF":
                            int.TryParse(value, out indicator_pcat_last_MVF);
                            tb_Product_Cat_Last_MVF.Text = value;
                            break;
                        case "indicator_pcat_avg":
                            int.TryParse(value, out indicator_pcat_avg);

                            break;
                        case "column_state":
                            int.TryParse(value, out column_state);
                            tb_Column_State.Text = value;
                            break;
                        case "indicator_state_ca":
                            indicator_state_ca = value;
                            tb_State_Indicator.Text = value;
                            break;
                        case "exclude_wh_1":
                            int.TryParse(value, out exclude_wh_1);
                            tb_Warehouse_Exclude_1.Text = value;
                            break;
                        case "column_warehouse":
                            int.TryParse(value, out column_warehouse);
                            tb_Column_Warehouse.Text = value;
                            break;
                        case "column_product":
                            int.TryParse(value, out column_product);
                            tb_Column_Product.Text = value;
                            break;
                        case "indicator_product_diesel_clear":
                            indicator_product_diesel_clear = value;
                            tb_Product_Indicator_Diesel.Text = value;
                            break;
                        case "indicator_product_diesel_dyed":
                            indicator_product_diesel_dyed = value;
                            tb_Product_Indicator_Red.Text = value;
                            break;
                        case "indicator_product_87":
                            indicator_product_87 = value;
                            tb_Product_Indicator_87.Text = value;
                            break;
                        case "indicator_product_89":
                            indicator_product_89 = value;
                            tb_Product_Indicator_89.Text = value;
                            break;
                        case "indicator_product_91":
                            indicator_product_91 = value;
                            tb_Product_Indicator_91.Text = value;
                            break;
                        case "indicator_product_avg":
                            indicator_product_avg = value;
                            tb_Product_Indicator_AVG.Text = value;
                            break;
                        case "indicator_product_oils":
                            indicator_product_oils = value;
                            break;
                        case "indicator_State_Not":
                            indicator_State_Not = value;
                            tb_Indicator_State_Not.Text = value;
                            break;
                        case "indicator_product_other":
                            indicator_product_other = value;
                            break;
                        case "column_quan":
                            int.TryParse(value, out column_quan);
                            tb_Column_Quantity.Text = value;
                            break;
                        case "column_taxable":
                            int.TryParse(value, out column_taxable);
                            tb_Column_Taxable.Text = value;
                            break;
                        case "column_sales":
                            int.TryParse(value, out column_sales);
                            tb_Column_Sales.Text = value;
                            break;
                        case "column_set":
                            int.TryParse(value, out column_set);
                            tb_Column_SET.Text = value;
                            break;
                        case "column_price":
                            int.TryParse(value, out column_price);
                            tb_Column_Price.Text = value;
                            break;
                        case "column_cost":
                            int.TryParse(value, out column_cost);
                            tb_Column_Cost.Text = value;
                            break;
                        case "column_fet":
                            int.TryParse(value, out column_fet);
                            tb_Column_FET.Text = value;
                            break;
                        case "column_salestax_zone":
                            int.TryParse(value, out column_salestax_zone);
                            tb_Column_Salestax_Zone.Text = value;
                            break;
                        case "column_salestax_amount":
                            int.TryParse(value, out column_salestax_amount);
                            tb_Column_Salestax_Amount.Text = value;
                            break;
                        case "column_salestax_rate":
                            int.TryParse(value, out column_salestax_rate);
                            tb_Column_Salestax_Rate.Text = value;
                            break;
                        case "column_salesman":
                            int.TryParse(value, out column_salesman);
                            tb_Column_Salesman.Text = value;
                            break;

                        case "column_wholesale":
                            int.TryParse(value, out column_wholesale);
                            tb_Column_Wholesale.Text = value;
                            break;

                        case "indicator_wholesale":
                            indicator_wholesale = value;
                            tb_Wholesale_Indicator.Text = value;
                            break;

                        case "indicator_retail":
                            indicator_retail = value;
                            break;

                        case "indicator_Excluded_Cardlock":
                            indicator_Excluded_Cardlock = value;
                            break;

                        case "column_excluded_cardlock":
                            int.TryParse(value, out column_excluded_cardlock);
                            tb_Column_Excluded_CL.Text = value;
                            break;


                        case "set_is_each":
                            if (value.Equals("True"))
                            {
                                set_is_each = true;
                                checkBox_SET_is_Rate.Checked = true;
                            }
                            else
                            {
                                set_is_each = false;
                                checkBox_SET_is_Rate.Checked = false;
                            }

                            break;
                        case "fet_is_each":
                            if (value.Equals("True"))
                            {
                                fet_is_each = true;
                                checkBox_FET_Rate.Checked = true;
                            }
                            else
                            {
                                fet_is_each = false;
                                checkBox_FET_Rate.Checked = false;
                            }

                            break;

                        case "sales_is_each":
                            if (value.Equals("True"))
                            {
                                sales_is_each = true;
                                checkBox_Sales_is_Each.Checked = true;
                            }
                            else
                            {
                                sales_is_each = false;
                                checkBox_Sales_is_Each.Checked = false;

                            }

                            break;

                        case "cost_is_each":
                            if (value.Equals("True"))
                            {
                                cost_is_each = true;
                                checkBox_Cost_is_Each.Checked = true;
                            }
                            else
                            {
                                cost_is_each = false;
                                checkBox_Cost_is_Each.Checked = false;
                            }

                            break;

                        case "price_is_each":
                            if (value.Equals("True"))
                            {
                                price_is_each = true;
                                checkBox_Price_is_Each.Checked = true;
                            }
                            else
                            {
                                price_is_each = false;
                                checkBox_Price_is_Each.Checked = false;
                            }

                            break;

                        case "use_taxable":
                            if (value.Equals("True"))
                            {
                                use_taxable = true;
                                checkBox_Use_Taxable.Checked = true;
                            }
                            else
                            {
                                use_taxable = false;
                                checkBox_Use_Taxable.Checked = false;
                            }

                            break;

                        case "use_sales":
                            if (value.Equals("True"))
                            {
                                use_sales = true;
                                checkBox_Use_Sales.Checked = true;
                            }
                            else
                            {
                                use_sales = false;
                                checkBox_Use_Sales.Checked = false;
                            }

                            break;

                        case "use_stax_as_wholesale":
                            if (value.Equals("True"))
                            {
                                use_stax_as_wholesale = true;
                                checkBox_STAX_as_Wholesale.Checked = true;
                            }
                            else
                            {
                                use_stax_as_wholesale = false;
                                checkBox_STAX_as_Wholesale.Checked = false;
                            }

                            break;

                        case "set_included_in_sales":
                            if (value.Equals("True"))
                            {
                                set_included_in_sales = true;
                                checkBox_SET_Included_Sales.Checked = true;
                            }
                            else
                            {
                                set_included_in_sales = false;
                                checkBox_SET_Included_Sales.Checked = false;
                            }

                            break;

                        case "stax_included_in_sales":
                            // MessageBox.Show(value);
                            if (value.Equals("True"))
                            {
                                stax_included_in_sales = true;
                                checkBox_STAX_in_Sales.Checked = true;

                            }
                            else
                            {
                                stax_included_in_sales = false;
                                checkBox_STAX_in_Sales.Checked = false;
                            }

                            break;

                        case "column_pretax_rate":
                            int.TryParse(value, out column_pretax_rate);
                            tb_Column_Pretax_Rate.Text = value;
                            break;
                        case "column_pretax_amount":
                            int.TryParse(value, out column_pretax_amount);
                            tb_Column_Pretax_Amount.Text = value;
                            break;

                        case "indicator_not_ca":

                            tb_Indicator_State_Not.Text = value;
                            break;

                        case "column_city":
                            tb_Column_City.Text = value;
                            break;

                        case "column_county":
                            tb_Column_County.Text = value;
                            break;

                        case "indicator_city_county_first":
                            tb_City_County_Order.Text = value;
                            break;

                        case "separater_city_county":
                            tb_City_County_Separator.Text = value;
                            break;


                    } //switch
                }//for

            }//if
            else
            {
                Save_Calc_Settings();
            }
        }

        void Save_Calc_Settings()
        {
            List<string> lines = new List<string>();

            string line = "indicator_cpst_diesel_clear" + "|" + indicator_cpst_diesel_clear.ToString();

            lines.Add(line);

            line = "indicator_cpst_diesel_dyed" + "|" + indicator_cpst_diesel_dyed.ToString();

            lines.Add(line);

            line = "indicator_cpst_equip" + "|" + indicator_cpst_equip.ToString();

            lines.Add(line);

            line = "column_cpst_amount" + "|" + column_cpst_amount.ToString();

            lines.Add(line);

            line = "column_cpst_indicator" + "|" + column_cpst_indicator.ToString();

            lines.Add(line);
            line = "column_pcat" + "|" + column_pcat.ToString();

            lines.Add(line);
            line = "indicator_pcat_diesel_clear" + "|" + indicator_pcat_diesel_clear.ToString();

            lines.Add(line);
            line = "indicator_pcat_diesel_dyed" + "|" + indicator_pcat_diesel_dyed.ToString();

            lines.Add(line);

            line = "indicator_pcat_87" + "|" + indicator_pcat_87.ToString();

            lines.Add(line);
            line = "indicator_pcat_89" + "|" + indicator_pcat_89.ToString();

            lines.Add(line);
            line = "indicator_pcat_91" + "|" + indicator_pcat_91.ToString();

            lines.Add(line);

            line = "indicator_pcat_avg" + "|" + indicator_pcat_avg.ToString();

            lines.Add(line);

            line = "indicator_pcat_last_MVF" + "|" + indicator_pcat_last_MVF.ToString();

            lines.Add(line);
            line = "column_state" + "|" + column_state.ToString();

            lines.Add(line);
            line = "indicator_state_ca" + "|" + indicator_state_ca.ToString();

            lines.Add(line);
            line = "exclude_wh_1" + "|" + exclude_wh_1.ToString();

            lines.Add(line);
            line = "column_warehouse" + "|" + column_warehouse.ToString();

            lines.Add(line);
            line = "column_product" + "|" + column_product.ToString();

            lines.Add(line);
            line = "indicator_product_diesel_clear" + "|" + indicator_product_diesel_clear.ToString();
            lines.Add(line);

            line = "indicator_product_diesel_dyed" + "|" + indicator_product_diesel_dyed.ToString();
            lines.Add(line);

            line = "indicator_product_87" + "|" + indicator_product_87.ToString();

            lines.Add(line);
            line = "indicator_product_89" + "|" + indicator_product_89.ToString();

            lines.Add(line);
            line = "indicator_product_91" + "|" + indicator_product_91.ToString();

            lines.Add(line);
            line = "indicator_product_avg" + "|" + indicator_product_avg.ToString();

            lines.Add(line);
            line = "indicator_product_oils" + "|" + indicator_product_oils.ToString();

            lines.Add(line);
            line = "indicator_product_other" + "|" + indicator_product_other.ToString();

            lines.Add(line);
            line = "column_quan" + "|" + column_quan.ToString();

            lines.Add(line);
            line = "column_taxable" + "|" + column_taxable.ToString();

            lines.Add(line);
            line = "column_sales" + "|" + column_sales.ToString();

            lines.Add(line);
            line = "column_set" + "|" + column_set.ToString();

            lines.Add(line);
            line = "column_price" + "|" + column_price.ToString();

            lines.Add(line);
            line = "cost_column" + "|" + column_cost.ToString();

            lines.Add(line);
            line = "column_fet" + "|" + column_fet.ToString();

            lines.Add(line);
            line = "column_salestax_zone" + "|" + column_salestax_zone.ToString();

            lines.Add(line);

            line = "column_salestax_amount" + "|" + column_salestax_amount.ToString();

            lines.Add(line);

            line = "column_salestax_rate" + "|" + column_salestax_rate.ToString();

            lines.Add(line);

            line = "column_salesman" + "|" + column_salesman.ToString();

            lines.Add(line);
            line = "column_wholesale" + "|" + column_wholesale.ToString();

            lines.Add(line);
            line = "indicator_wholesale" + "|" + indicator_wholesale.ToString();

            lines.Add(line);
            line = "indicator_retail" + "|" + indicator_retail.ToString();

            lines.Add(line);
            line = "set_is_each" + "|" + set_is_each.ToString();

            lines.Add(line);
            line = "fet_is_each" + "|" + fet_is_each.ToString();

            lines.Add(line);
            line = "sales_is_each" + "|" + sales_is_each.ToString();

            lines.Add(line);
            line = "cost_is_each" + "|" + cost_is_each.ToString();

            lines.Add(line);
            line = "price_is_each" + "|" + price_is_each.ToString();

            lines.Add(line);
            line = "use_taxable" + "|" + use_taxable.ToString();

            lines.Add(line);
            line = "use_sales" + "|" + use_sales.ToString();

            lines.Add(line);
            line = "use_stax_as_wholesale" + "|" + use_stax_as_wholesale.ToString();

            lines.Add(line);
            line = "set_included_in_sales" + "|" + set_included_in_sales.ToString();

            lines.Add(line);
            line = "stax_included_in_sales" + "|" + stax_included_in_sales.ToString();

            lines.Add(line);
            line = "column_pretax_rate" + "|" + column_pretax_rate.ToString();

            lines.Add(line);

            line = "column_pretax_amount" + "|" + column_pretax_amount.ToString();
            lines.Add(line);

            line = "indicator_not_ca" + "|" + tb_Indicator_State_Not.Text;
            lines.Add(line);

            line = "column_city" + "|" + tb_Column_City.Text;
            lines.Add(line);

            line = "column_county" + "|" + tb_Column_County.Text;
            lines.Add(line);

            line = "indicator_city_county_first" + "|" + tb_City_County_Order.Text;
            lines.Add(line);

            line = "separater_city_county" + "|" + tb_City_County_Separator.Text;
            lines.Add(line);

            File.WriteAllLines(SalesAggCalcSettings, lines);
        }

        void Reset_Totals()
        {
            gross_sales = 0.00m;
            diesel_Clear = 0.00m;
            diesel_SET = 0.00m;
            cpst = 0.00m;
            cpst_diesel_clear = 0.00m;
            cpst_diesel_red = 0.00m;
            cpst_farm_equip = 0.00m;

            mvf_sales = 0.00m;

            wholesale = 0.00m;
            excluded_cardlock = 0;
            foreign = 0;

        }

        void Init_Zones()
        {
            if (File.Exists(Zones))
            {
                string[] lines = File.ReadAllLines(Zones);

                foreach (string line in lines)
                {
                    var row = zones.NewRow();

                    string[] fields = line.Split(',');

                    for (int i = 0; i < fields.Length; i++)
                    {
                        row[i] = fields[i];
                    }

                    zones.Rows.Add(row);
                }

                dgv_Tax_Zones.DataSource = zones;
            }
            else
            {
                MessageBox.Show("Zones File " + Zones + " Not Available!");
            }


        }

        void Load_Zones_With_STJ_Rates()
        {
            string city = "";
            string cnty = "";

            foreach (DataRow row in zones.Rows)
            {
                cnty = row["name"].ToString();
                city = "";

                if (cnty.Contains("/"))
                {
                    string[] fields = cnty.Split('/');
                    city = fields[1];
                    cnty = fields[0];

                    if (cnty.Contains("CO."))
                    {
                        cnty = cnty.Replace("CO.", "").Trim();
                    }

                }

                decimal jrate = Get_Jurisdiction_Rate(city, cnty,dtp_Date_End.Value);
                row["stj_rate"] = jrate.ToString();

                int stjcode = Get_Jurisdiction_STJ(city, cnty);
                row["stj"] = stjcode.ToString();
            }

        }

        void init_Date()
        {
            DateTime cur = DateTime.Now.AddMonths(-3);


            dtp_Date_Beg.Value = DateTime.Parse(cur.Month.ToString() + "-01-" + DateTime.Now.Year.ToString());

            dtp_Date_End.Value = dtp_Date_Beg.Value.AddMonths(3).AddDays(-1);

        }

        void Load_CSV_To_DataTable(string filename, DataTable tablename, int startline)
        {
            string[] lines = File.ReadAllLines(filename);
            int fileline = 0;

            if (lines.Length > 1)
            {

                MessageBox.Show("File Name = " + filename + " DataTable = " + tablename.TableName);

                foreach (string line in lines)
                {
                    if (fileline >= startline)
                    {
                        DataRow newrow = tablename.NewRow();
                        string[] fields = line.Split('.');

                        for (int i = 0; i < tablename.Columns.Count; i++)
                        {
                            newrow[i] = fields[i - 1];
                        }

                        tablename.Rows.Add(newrow);
                    }
                    fileline++;
                }
            }
            else
            {
                MessageBox.Show("No data in " + filename);
            }

        }

        void Load_CSV_To_List(string filename, List<string> listname, int startline)
        {
            string[] lines = File.ReadAllLines(filename);
            int fileline = 0;


            foreach (string line in lines)
            {
                if (fileline >= startline)
                {

                    string[] fields = line.Split('.');


                }
                fileline++;
            }

        }

        void Load_CADSL_Rates()
        {
            string[] lines = File.ReadAllLines(CATaxDSLRates);

            int fileline = 0;

            foreach (string line in lines)
            {
                if (fileline > 0)
                {
                    DataRow newrow = CATaxRates_DSL.NewRow();

                    string[] fields = line.Split(',');

                    for (int i = 0; i < CATaxRates_DSL.Columns.Count; i++)
                    {
                        newrow[i] = fields[i];
                    }

                    CATaxRates_DSL.Rows.Add(newrow);
                }
                fileline++;
            }

        }

        void Load_CAMVF_Rates()
        {
            string[] lines = File.ReadAllLines(CATaxMVFRates);

            int fileline = 0;

            foreach (string line in lines)
            {
                if (fileline > 0)
                {
                    DataRow newrow = CATaxRates_MVF.NewRow();

                    string[] fields = line.Split(',');

                    for (int i = 0; i < CATaxRates_MVF.Columns.Count; i++)
                    {
                        newrow[i] = fields[i];
                    }

                    CATaxRates_MVF.Rows.Add(newrow);
                }
                fileline++;
            }

        }

        void Load_VoyCodes()
        {
            string[] lines = File.ReadAllLines(VoyagerCodeFile);

            int fileline = 0;

            foreach (string line in lines)
            {
                if (fileline > 0)
                {
                    string[] fields = line.Split(',');

                    VoyagerCodes c1 = new VoyagerCodes();
                    c1.Code = fields[0];
                    c1.Description = fields[1];
                    c1.FTACode = fields[2];

                    VoyCodes.Add(c1);
                }
                fileline++;
            }

        }

        void Load_CFN_Sites()
        {
            string[] lines = File.ReadAllLines(CFNsitesfile);

            int fileline = 0;
            foreach (string line in lines)
            {
                if (fileline > 0)
                {

                    Sites_CFN s = new Sites_CFN();

                    string[] fields = line.Split(',');

                    s.Site = fields[0];
                    s.Company = fields[1];
                    s.Participant = fields[2];
                    s.Address1 = fields[3];
                    s.CityName = fields[4];
                    s.CityCode = fields[5];
                    s.CountyCode = fields[6];
                    s.State = fields[7];
                    s.Phone = fields[8];
                    s.CStore = fields[9];
                    s.Zip = fields[10];
                    s.CompType = fields[11];


                    CFNSites.Add(s);
                }
                fileline++;
            }

            string[] lines2 = File.ReadAllLines(CFNextendedsitescafile);

            fileline = 0;

            foreach (string line in lines2)
            {
                if (fileline > 0)
                {

                    Sites_CFN s = new Sites_CFN();
                    string[] fields = line.Split(',');
                    // MessageBox.Show(line);

                    s.Site = fields[0];
                    s.Participant = "Fuelman";
                    s.Company = "Fleetcore";
                    s.Address1 = fields[2];
                    s.CityName = fields[3];
                    s.State = fields[4];
                    s.Zip = fields[6];
                    s.CountyCode = fields[9];
                    s.CityCode = fields[10];

                    CFNSites.Add(s);

                }
                fileline++;
            }

            dgv_CFN_Sites.DataSource = CFNSites;
        }

        void Load_CFN_Counties()
        {
            string[] lines = File.ReadAllLines(CFNcountiesfile);

            int fileline = 0;
            foreach (string line in lines)
            {
                if (fileline > 0)
                {
                    Counties_CFN Cnty = new Counties_CFN();

                    string[] fields = line.Split(',');


                    Cnty.County = fields[0];
                    Cnty.Code = fields[1];

                    CFNCounties.Add(Cnty);


                }
                fileline++;
            }
            dgv_CFN_Counties.DataSource = CFNCounties;
        }

        void Load_PP_Products()
        {
            string[] lines = File.ReadAllLines(productfile);

            foreach (string line in lines)
            {
                // DataRow newrow = pp_products.NewRow();

                Products_PP newrow = new Products_PP();

                string[] fields = line.Split(',');

                newrow.Code = fields[0];
                newrow.Description = fields[1];
                newrow.FTA_Code = fields[2];
                newrow.Account = fields[3];

                pp_products.Add(newrow);
            }

            dgv_Pac_Pride_Products.DataSource = pp_products;
        }


        void Load_CFN_Products()
        {

            string[] lines = File.ReadAllLines(CFNprodfile);


            foreach (string line in lines)
            {

                Products_CFN p = new Products_CFN();


                string[] fields = line.Split(',');

                p.Code = fields[0];
                p.Description = fields[1];
                p.FTA_Code = fields[2];

                cfn_products.Add(p);
            }


        }

        void Init_DataTables()
        {
            //locdata.Columns.Add("Name", typeof(string));
            //locdata.Columns.Add("CloseDate", typeof(DateTime));
            //locdata.Columns.Add("Number", typeof(string));
            //locdata.Columns.Add("Address1", typeof(string));
            //locdata.Columns.Add("Address2", typeof(string));
            //locdata.Columns.Add("City", typeof(string));
            //locdata.Columns.Add("State", typeof(string));
            //locdata.Columns.Add("ZipCode", typeof(string));
            //locdata.Columns.Add("CountyCode", typeof(string));
            //locdata.Columns.Add("DistrictCode", typeof(string));
            //locdata.Columns.Add("InLieuCode", typeof(string));
            //locdata.Columns.Add("LocalCode", typeof(string));
            //locdata.Columns.Add("Ratio", typeof(double));
            //locdata.Columns.Add("Allocation", typeof(double));
            //locdata.Columns.Add("New", typeof(bool));
            //locdata.Columns.Add("AddDate", typeof(DateTime));

            //PTfile.Columns.Add("Site", typeof(string));
            //PTfile.Columns.Add("Sequence", typeof(string));
            //PTfile.Columns.Add("Status", typeof(string));
            //PTfile.Columns.Add("Total", typeof(string));
            //PTfile.Columns.Add("Account", typeof(string));
            //PTfile.Columns.Add("Product", typeof(string));
            //PTfile.Columns.Add("Type", typeof(string));
            //PTfile.Columns.Add("ProdDesc", typeof(string));
            //PTfile.Columns.Add("Price", typeof(string));
            //PTfile.Columns.Add("Quantity", typeof(string));
            //PTfile.Columns.Add("Odometer", typeof(string));
            //PTfile.Columns.Add("Pump", typeof(string));
            //PTfile.Columns.Add("Trans", typeof(string));
            //PTfile.Columns.Add("Date", typeof(string));
            //PTfile.Columns.Add("Time", typeof(string));
            //PTfile.Columns.Add("Error", typeof(string));
            //PTfile.Columns.Add("Authorization", typeof(string));
            //PTfile.Columns.Add("ManualEntry", typeof(string));
            //PTfile.Columns.Add("Card", typeof(string));
            //PTfile.Columns.Add("Vehicle", typeof(string));
            //PTfile.Columns.Add("SiteTaxLocation", typeof(string));
            //PTfile.Columns.Add("Code0", typeof(string));
            //PTfile.Columns.Add("Code1", typeof(string));
            //PTfile.Columns.Add("Code2", typeof(string));
            //PTfile.Columns.Add("Code3", typeof(string));
            //PTfile.Columns.Add("Code4", typeof(string));
            //PTfile.Columns.Add("Code5", typeof(string));
            //PTfile.Columns.Add("Code6", typeof(string));
            //PTfile.Columns.Add("Code7", typeof(string));
            //PTfile.Columns.Add("Code8", typeof(string));
            //PTfile.Columns.Add("Code9", typeof(string));
            //PTfile.Columns.Add("Amount0", typeof(string));
            //PTfile.Columns.Add("Amount1", typeof(string));
            //PTfile.Columns.Add("Amount2", typeof(string));
            //PTfile.Columns.Add("Amount3", typeof(string));
            //PTfile.Columns.Add("Amount4", typeof(string));
            //PTfile.Columns.Add("Amount5", typeof(string));
            //PTfile.Columns.Add("Amount6", typeof(string));
            //PTfile.Columns.Add("Amount7", typeof(string));
            //PTfile.Columns.Add("Amount8", typeof(string));
            //PTfile.Columns.Add("Amount9", typeof(string));
            //PTfile.Columns.Add("NetType", typeof(string));
            //PTfile.Columns.Add("CF", typeof(string));
            //PTfile.Columns.Add("NetRate", typeof(string));
            //PTfile.Columns.Add("PumpPrice", typeof(string));
            //PTfile.Columns.Add("HaulRate", typeof(string));
            //PTfile.Columns.Add("CFNPrice", typeof(string));
            //PTfile.Columns.Add("JobNumber", typeof(string));
            //PTfile.Columns.Add("PONumber", typeof(string));
            //PTfile.Columns.Add("City", typeof(string));
            //PTfile.Columns.Add("County", typeof(string));
            //PTfile.Columns.Add("State", typeof(string));
            //PTfile.Columns.Add("CityTax", typeof(string));
            //PTfile.Columns.Add("CountyTax", typeof(string));
            //PTfile.Columns.Add("StateTax", typeof(string));
            //PTfile.Columns.Add("SET", typeof(string));
            //PTfile.Columns.Add("Taxable", typeof(string));

            //PTExtended.Columns.Add("Site", typeof(string));
            //PTExtended.Columns.Add("Sequence", typeof(string));
            //PTExtended.Columns.Add("Status", typeof(string));
            //PTExtended.Columns.Add("Total", typeof(string));
            //PTExtended.Columns.Add("Account", typeof(string));
            //PTExtended.Columns.Add("Product", typeof(string));
            //PTExtended.Columns.Add("Type", typeof(string));
            //PTExtended.Columns.Add("ProdDesc", typeof(string));
            //PTExtended.Columns.Add("Price", typeof(string));
            //PTExtended.Columns.Add("Quantity", typeof(string));
            //PTExtended.Columns.Add("Odometer", typeof(string));
            //PTExtended.Columns.Add("Pump", typeof(string));
            //PTExtended.Columns.Add("Trans", typeof(string));
            //PTExtended.Columns.Add("Date", typeof(string));
            //PTExtended.Columns.Add("Time", typeof(string));
            //PTExtended.Columns.Add("Error", typeof(string));
            //PTExtended.Columns.Add("Authorization", typeof(string));
            //PTExtended.Columns.Add("ManualEntry", typeof(string));
            //PTExtended.Columns.Add("Card", typeof(string));
            //PTExtended.Columns.Add("Vehicle", typeof(string));
            //PTExtended.Columns.Add("SiteTaxLocation", typeof(string));
            //PTExtended.Columns.Add("Code0", typeof(string));
            //PTExtended.Columns.Add("Code1", typeof(string));
            //PTExtended.Columns.Add("Code2", typeof(string));
            //PTExtended.Columns.Add("Code3", typeof(string));
            //PTExtended.Columns.Add("Code4", typeof(string));
            //PTExtended.Columns.Add("Code5", typeof(string));
            //PTExtended.Columns.Add("Code6", typeof(string));
            //PTExtended.Columns.Add("Code7", typeof(string));
            //PTExtended.Columns.Add("Code8", typeof(string));
            //PTExtended.Columns.Add("Code9", typeof(string));
            //PTExtended.Columns.Add("Amount0", typeof(string));
            //PTExtended.Columns.Add("Amount1", typeof(string));
            //PTExtended.Columns.Add("Amount2", typeof(string));
            //PTExtended.Columns.Add("Amount3", typeof(string));
            //PTExtended.Columns.Add("Amount4", typeof(string));
            //PTExtended.Columns.Add("Amount5", typeof(string));
            //PTExtended.Columns.Add("Amount6", typeof(string));
            //PTExtended.Columns.Add("Amount7", typeof(string));
            //PTExtended.Columns.Add("Amount8", typeof(string));
            //PTExtended.Columns.Add("Amount9", typeof(string));
            //PTExtended.Columns.Add("NetType", typeof(string));
            //PTExtended.Columns.Add("CF", typeof(string));
            //PTExtended.Columns.Add("NetRate", typeof(string));
            //PTExtended.Columns.Add("PumpPrice", typeof(string));
            //PTExtended.Columns.Add("HaulRate", typeof(string));
            //PTExtended.Columns.Add("CFNPrice", typeof(string));
            //PTExtended.Columns.Add("JobNumber", typeof(string));
            //PTExtended.Columns.Add("PONumber", typeof(string));
            //PTExtended.Columns.Add("City", typeof(string));
            //PTExtended.Columns.Add("County", typeof(string));
            //PTExtended.Columns.Add("State", typeof(string));
            //PTExtended.Columns.Add("CityTax", typeof(string));
            //PTExtended.Columns.Add("CountyTax", typeof(string));
            //PTExtended.Columns.Add("StateTax", typeof(string));
            //PTExtended.Columns.Add("SET", typeof(string));
            //PTExtended.Columns.Add("Taxable", typeof(string));

            //CFNsites.Columns.Add("Site", typeof(string));
            //CFNsites.Columns.Add("Company", typeof(string));
            //CFNsites.Columns.Add("Participant", typeof(string));
            //CFNsites.Columns.Add("Address1", typeof(string));
            //CFNsites.Columns.Add("CityName", typeof(string));
            //CFNsites.Columns.Add("CityCode", typeof(string));
            //CFNsites.Columns.Add("CountyCode", typeof(string));
            //CFNsites.Columns.Add("State", typeof(string));
            //CFNsites.Columns.Add("Phone", typeof(string));
            //CFNsites.Columns.Add("C-Store", typeof(string));
            //CFNsites.Columns.Add("Zip", typeof(string));
            //CFNsites.Columns.Add("CompType", typeof(string));

            //CFNcounties.Columns.Add("County", typeof(string));
            //CFNcounties.Columns.Add("Code", typeof(string));


            workTable.Columns.Add("County", typeof(string));
            workTable.Columns.Add("City", typeof(string));
            workTable.Columns.Add("STJ", typeof(string));
            workTable.Columns.Add("Effective", typeof(string));
            workTable.Columns.Add("Expired", typeof(string));
            workTable.Columns.Add("Rate", typeof(double));
            workTable.Columns.Add("Sales", typeof(decimal));
            workTable.Columns.Add("Adjustments", typeof(decimal));
            workTable.Columns.Add("Tax", typeof(decimal));
            workTable.Columns.Add("PaidPP", typeof(decimal));
            workTable.Columns.Add("PaidCFN", typeof(decimal));
            workTable.Columns.Add("PaidVoy", typeof(decimal));
            workTable.Columns.Add("PaidWEX", typeof(decimal));
            workTable.Columns.Add("Net", typeof(decimal));
            workTable.Columns.Add("PPPurch", typeof(decimal));
            workTable.Columns.Add("CFNPurch", typeof(decimal));
            workTable.Columns.Add("VoyPurch", typeof(decimal));
            workTable.Columns.Add("WEXPurch", typeof(decimal));

            pretax.Columns.Add("Name", typeof(string));
            pretax.Columns.Add("Account", typeof(string));
            pretax.Columns.Add("MVFGallons", typeof(int));
            pretax.Columns.Add("MVFRate", typeof(decimal));
            pretax.Columns.Add("MVFTax", typeof(decimal));
            pretax.Columns.Add("DSLGallons", typeof(int));
            pretax.Columns.Add("DSLRate", typeof(decimal));
            pretax.Columns.Add("DSLTax", typeof(decimal));
            pretax.Columns.Add("AJFGallons", typeof(int));
            pretax.Columns.Add("AJFRate", typeof(decimal));
            pretax.Columns.Add("AJFTax", typeof(decimal));

            //pp_products.Columns.Add("Code", typeof(string));
            //pp_products.Columns.Add("Description", typeof(string));
            //pp_products.Columns.Add("FTA_Code", typeof(string));
            //pp_products.Columns.Add("Account", typeof(string));

            //cfn_products.Columns.Add("Code", typeof(string));
            //cfn_products.Columns.Add("Description", typeof(string));
            //cfn_products.Columns.Add("FTA_Code", typeof(string));
            //cfn_products.Columns.Add("Account", typeof(string));

            //locbil.Columns.Add("Record_Code", typeof(string));
            //locbil.Columns.Add("SellingPart", typeof(string));
            //locbil.Columns.Add("Site_Code", typeof(string));
            //locbil.Columns.Add("Site_Type", typeof(string));
            //locbil.Columns.Add("Site_Street", typeof(string));
            //locbil.Columns.Add("Site_City", typeof(string));
            //locbil.Columns.Add("Site_State", typeof(string));
            //locbil.Columns.Add("Site_County", typeof(string));
            //locbil.Columns.Add("Site_Zip", typeof(string));
            //locbil.Columns.Add("Trans", typeof(string));
            //locbil.Columns.Add("Trans_date", typeof(string));
            //locbil.Columns.Add("Trans_Time", typeof(string));
            //locbil.Columns.Add("Capture_Date", typeof(string));
            //locbil.Columns.Add("Capture_Time", typeof(string));
            //locbil.Columns.Add("Card", typeof(string));
            //locbil.Columns.Add("Vehicle", typeof(string));
            //locbil.Columns.Add("Host", typeof(string));
            //locbil.Columns.Add("Identity", typeof(string));
            //locbil.Columns.Add("Misc_Key", typeof(string));
            //locbil.Columns.Add("Odometer", typeof(string));
            //locbil.Columns.Add("Trans_Number", typeof(string));
            //locbil.Columns.Add("Trans_Sequence", typeof(string));
            //locbil.Columns.Add("Pump", typeof(string));
            //locbil.Columns.Add("Hose", typeof(string));
            //locbil.Columns.Add("Product", typeof(string));
            //locbil.Columns.Add("Quantity", typeof(string));
            //locbil.Columns.Add("UnitOfMeas", typeof(string));
            //locbil.Columns.Add("SellingPrice", typeof(string));
            //locbil.Columns.Add("TransferCost", typeof(string));
            //locbil.Columns.Add("SalesTaxAdjust", typeof(string));
            //locbil.Columns.Add("NetWorkICBDate", typeof(string));
            //locbil.Columns.Add("Batch", typeof(string));
            //locbil.Columns.Add("IssuerCode", typeof(string));
            //locbil.Columns.Add("AuthCode", typeof(string));
            //locbil.Columns.Add("ServLevel", typeof(string));
            //locbil.Columns.Add("RetailTotal", typeof(string));
            //locbil.Columns.Add("RetailPrice", typeof(string));
            //locbil.Columns.Add("RetailRebate", typeof(string));
            //locbil.Columns.Add("StateCode", typeof(string));
            //locbil.Columns.Add("FET", typeof(string));
            //locbil.Columns.Add("SET1", typeof(string));
            //locbil.Columns.Add("SET2", typeof(string));
            //locbil.Columns.Add("CountyET", typeof(string));
            //locbil.Columns.Add("CityET", typeof(string));
            //locbil.Columns.Add("PercentState", typeof(string));
            //locbil.Columns.Add("PercentCounty", typeof(string));
            //locbil.Columns.Add("PercentCity", typeof(string));
            //locbil.Columns.Add("PercentOther", typeof(string));
            //locbil.Columns.Add("OtherName", typeof(string));
            //locbil.Columns.Add("FETRate", typeof(string));
            //locbil.Columns.Add("SETRate", typeof(string));
            //locbil.Columns.Add("OtherinTransfer", typeof(string));
            //locbil.Columns.Add("CountyInTransfer", typeof(string));
            //locbil.Columns.Add("CityInTransfer", typeof(string));
            //locbil.Columns.Add("StateSalesInTransfer", typeof(string));
            //locbil.Columns.Add("CountySalesInTransfer", typeof(string));
            //locbil.Columns.Add("CitySalesInTransfer", typeof(string));
            //locbil.Columns.Add("OtherInTransfer", typeof(string));
            //locbil.Columns.Add("FETIncluded", typeof(string));
            //locbil.Columns.Add("SETIncluded", typeof(string));
            //locbil.Columns.Add("StateOtherIncluded", typeof(string));
            //locbil.Columns.Add("Include1", typeof(string));
            //locbil.Columns.Add("Include2", typeof(string));
            //locbil.Columns.Add("Include3", typeof(string));
            //locbil.Columns.Add("Include4", typeof(string));
            //locbil.Columns.Add("Include5", typeof(string));
            //locbil.Columns.Add("Include6", typeof(string));
            //locbil.Columns.Add("Include7", typeof(string));
            //locbil.Columns.Add("FETIncludedSelling", typeof(string));
            //locbil.Columns.Add("SETIncludedSelling", typeof(string));
            //locbil.Columns.Add("Included10", typeof(string));
            //locbil.Columns.Add("Included11", typeof(string));
            //locbil.Columns.Add("Included12", typeof(string));
            //locbil.Columns.Add("SalesTaxIncludedSelling", typeof(string));
            //locbil.Columns.Add("Included14", typeof(string));
            //locbil.Columns.Add("Included15", typeof(string));
            //locbil.Columns.Add("Included16", typeof(string));
            //locbil.Columns.Add("ISONumber", typeof(string));
            //locbil.Columns.Add("ISOCard", typeof(string));

            int c = 0;

            //foreach (DataColumn col in locbil.Columns)
            //{
            //    locfields[c] = col.ColumnName;
            //    c++;
            //}

            //extended.Columns.Add("Record_Code", typeof(string));
            //extended.Columns.Add("SellingPart", typeof(string));
            //extended.Columns.Add("Site_Code", typeof(string));
            //extended.Columns.Add("Site_Type", typeof(string));
            //extended.Columns.Add("Site_Street", typeof(string));
            //extended.Columns.Add("Site_City", typeof(string));
            //extended.Columns.Add("Site_State", typeof(string));
            //extended.Columns.Add("Site_County", typeof(string));
            //extended.Columns.Add("Site_Zip", typeof(string));
            //extended.Columns.Add("Trans", typeof(string));
            //extended.Columns.Add("Trans_date", typeof(string));
            //extended.Columns.Add("Trans_Time", typeof(string));
            //extended.Columns.Add("Capture_Date", typeof(string));
            //extended.Columns.Add("Capture_Time", typeof(string));
            //extended.Columns.Add("Card", typeof(string));
            //extended.Columns.Add("Vehicle", typeof(string));
            //extended.Columns.Add("Host", typeof(string));
            //extended.Columns.Add("Identity", typeof(string));
            //extended.Columns.Add("Misc_Key", typeof(string));
            //extended.Columns.Add("Odometer", typeof(string));
            //extended.Columns.Add("Trans_Number", typeof(string));
            //extended.Columns.Add("Trans_Sequence", typeof(string));
            //extended.Columns.Add("Pump", typeof(string));
            //extended.Columns.Add("Hose", typeof(string));
            //extended.Columns.Add("Product", typeof(string));
            //extended.Columns.Add("Quantity", typeof(string));
            //extended.Columns.Add("UnitOfMeas", typeof(string));
            //extended.Columns.Add("SellingPrice", typeof(string));
            //extended.Columns.Add("TransferCost", typeof(string));
            //extended.Columns.Add("SalesTaxAdjust", typeof(string));
            //extended.Columns.Add("NetWorkICBDate", typeof(string));
            //extended.Columns.Add("Batch", typeof(string));
            //extended.Columns.Add("IssuerCode", typeof(string));
            //extended.Columns.Add("AuthCode", typeof(string));
            //extended.Columns.Add("ServLevel", typeof(string));
            //extended.Columns.Add("RetailTotal", typeof(string));
            //extended.Columns.Add("RetailPrice", typeof(string));
            //extended.Columns.Add("RetailRebate", typeof(string));
            //extended.Columns.Add("StateCode", typeof(string));
            //extended.Columns.Add("FET", typeof(string));
            //extended.Columns.Add("SET1", typeof(string));
            //extended.Columns.Add("SET2", typeof(string));
            //extended.Columns.Add("CountyET", typeof(string));
            //extended.Columns.Add("CityET", typeof(string));
            //extended.Columns.Add("PercentState", typeof(string));
            //extended.Columns.Add("PercentCounty", typeof(string));
            //extended.Columns.Add("PercentCity", typeof(string));
            //extended.Columns.Add("PercentOther", typeof(string));
            //extended.Columns.Add("OtherName", typeof(string));
            //extended.Columns.Add("FETRate", typeof(string));
            //extended.Columns.Add("SETRate", typeof(string));
            //extended.Columns.Add("OtherinTransfer", typeof(string));
            //extended.Columns.Add("CountyInTransfer", typeof(string));
            //extended.Columns.Add("CityInTransfer", typeof(string));
            //extended.Columns.Add("StateSalesInTransfer", typeof(string));
            //extended.Columns.Add("CountySalesInTransfer", typeof(string));
            //extended.Columns.Add("CitySalesInTransfer", typeof(string));
            //extended.Columns.Add("OtherInTransfer", typeof(string));
            //extended.Columns.Add("FETIncluded", typeof(string));
            //extended.Columns.Add("SETIncluded", typeof(string));
            //extended.Columns.Add("StateOtherIncluded", typeof(string));
            //extended.Columns.Add("Include1", typeof(string));
            //extended.Columns.Add("Include2", typeof(string));
            //extended.Columns.Add("Include3", typeof(string));
            //extended.Columns.Add("Include4", typeof(string));
            //extended.Columns.Add("Include5", typeof(string));
            //extended.Columns.Add("Include6", typeof(string));
            //extended.Columns.Add("Include7", typeof(string));
            //extended.Columns.Add("FETIncludedSelling", typeof(string));
            //extended.Columns.Add("SETIncludedSelling", typeof(string));
            //extended.Columns.Add("Included10", typeof(string));
            //extended.Columns.Add("Included11", typeof(string));
            //extended.Columns.Add("Included12", typeof(string));
            //extended.Columns.Add("SalesTaxIncludedSelling", typeof(string));
            //extended.Columns.Add("Included14", typeof(string));
            //extended.Columns.Add("Included15", typeof(string));
            //extended.Columns.Add("Included16", typeof(string));
            //extended.Columns.Add("ISONumber", typeof(string));
            //extended.Columns.Add("ISOCard", typeof(string));

            cadtax.Columns.Add("County", typeof(string));
            cadtax.Columns.Add("City", typeof(string));
            cadtax.Columns.Add("Stj", typeof(string));
            cadtax.Columns.Add("Effective", typeof(string));
            cadtax.Columns.Add("Expired", typeof(string));
            cadtax.Columns.Add("Rate", typeof(string));
            cadtax.Columns.Add("Sales", typeof(string));
            cadtax.Columns.Add("Adjustments", typeof(string));
            cadtax.Columns.Add("Tax", typeof(string));
            cadtax.Columns.Add("PaidPP", typeof(string));
            cadtax.Columns.Add("PaidCFN", typeof(string));
            cadtax.Columns.Add("Net", typeof(string));

            juris.Columns.Add("City", typeof(string));
            juris.Columns.Add("State", typeof(string));
            juris.Columns.Add("Statecode", typeof(string));
            juris.Columns.Add("County", typeof(string));
            juris.Columns.Add("STJCity", typeof(string));
            juris.Columns.Add("STJCounty", typeof(string));
            juris.Columns.Add("RateCity", typeof(string));
            juris.Columns.Add("RateCounty", typeof(string));

            zones.Columns.Add("tax_type", typeof(string));
            zones.Columns.Add("authority", typeof(string));
            zones.Columns.Add("name", typeof(string));
            zones.Columns.Add("sales_taxable", typeof(string));
            zones.Columns.Add("modify_date", typeof(string));
            zones.Columns.Add("stj_rate", typeof(string));
            zones.Columns.Add("stj", typeof(string));

            //Voyfile.Columns.Add("Section1", typeof(string));
            //Voyfile.Columns.Add("SiteAddress", typeof(string));
            //Voyfile.Columns.Add("SiteCity", typeof(string));
            //Voyfile.Columns.Add("SiteCity2", typeof(string));
            //Voyfile.Columns.Add("State", typeof(string));
            //Voyfile.Columns.Add("Zip", typeof(string));
            //Voyfile.Columns.Add("Product", typeof(string));
            //Voyfile.Columns.Add("Amount", typeof(string));
            //Voyfile.Columns.Add("Quantity", typeof(string));
            //Voyfile.Columns.Add("Price", typeof(string));
            //Voyfile.Columns.Add("County", typeof(string));
            //Voyfile.Columns.Add("STJ", typeof(string));
            //Voyfile.Columns.Add("RateDistrict", typeof(string));
            //Voyfile.Columns.Add("RateState", typeof(string));
            //Voyfile.Columns.Add("Base", typeof(string));
            //Voyfile.Columns.Add("Fee", typeof(string));
            //Voyfile.Columns.Add("FET", typeof(string));
            //Voyfile.Columns.Add("SET", typeof(string));
            //Voyfile.Columns.Add("FTA", typeof(string));
            //Voyfile.Columns.Add("TaxBase", typeof(string));
            //Voyfile.Columns.Add("Tax", typeof(string));
            //Voyfile.Columns.Add("TaxDistrict", typeof(string));
            //Voyfile.Columns.Add("TaxState", typeof(string));
            //Voyfile.Columns.Add("Section3", typeof(string));
            //Voyfile.Columns.Add("Date", typeof(string));

            //VoyCodes.Columns.Add("Code", typeof(string));
            //VoyCodes.Columns.Add("Description", typeof(string));
            //VoyCodes.Columns.Add("FTACode", typeof(string));

            CATaxRates_DSL.Columns.Add("Start", typeof(string));
            CATaxRates_DSL.Columns.Add("End", typeof(string));
            CATaxRates_DSL.Columns.Add("SalesTaxRate", typeof(string));
            CATaxRates_DSL.Columns.Add("CPSTRate", typeof(string));
            CATaxRates_DSL.Columns.Add("PreTaxRate", typeof(string));
            CATaxRates_DSL.Columns.Add("SET", typeof(string));

            CATaxRates_MVF.Columns.Add("Start", typeof(string));
            CATaxRates_MVF.Columns.Add("End", typeof(string));
            CATaxRates_MVF.Columns.Add("SalesTaxRate", typeof(string));
            CATaxRates_MVF.Columns.Add("CPSTRate", typeof(string));
            CATaxRates_MVF.Columns.Add("PreTaxRate", typeof(string));
            CATaxRates_MVF.Columns.Add("SET", typeof(string));

            CATaxRates_AVG.Columns.Add("Start", typeof(string));
            CATaxRates_AVG.Columns.Add("End", typeof(string));
            CATaxRates_AVG.Columns.Add("SalesTaxRate", typeof(string));
            CATaxRates_AVG.Columns.Add("CPSTRate", typeof(string));
            CATaxRates_AVG.Columns.Add("PreTaxRate", typeof(string));
            CATaxRates_AVG.Columns.Add("SET", typeof(string));

        }



        void Tax_Rates_To_Grid()
        {
            notifyIcon1.Icon = SystemIcons.Warning;
            notifyIcon1.BalloonTipText = "Loading District Tax Rates into Grid";
            notifyIcon1.BalloonTipTitle = "Rates";
            notifyIcon1.Visible = true;
            notifyIcon1.ShowBalloonTip(56000);

            WebClient client = new WebClient();

            client.DownloadFile(districtfile, pdffile);

            Spire.Pdf.PdfDocument document = new Spire.Pdf.PdfDocument(pdffile);
            int pagec = document.Pages.Count;

            PdfFormWidget fw = document.Form as PdfFormWidget;

            int fcnt = fw.FieldsWidget.List.Count;

            for (int i = 0; i < pagec; i++)
            {
                var page = document.Pages[i];

                text.Append(page.ExtractText());
            }

            for (int i = 0; i < fcnt; i++)
            {
                PdfField field = fw.FieldsWidget.List[i] as PdfField;

                //  MessageBox.Show(field.Name.ToString());

                fields.Append(field.Name.ToString() + Environment.NewLine);
            }


            String CountyIOn = "";

            var result = Regex.Split(text.ToString(), "\r\n|\r|\n");

            string[] lines = result.ToArray();

            if (lines.Length > 0)
            {
                bool done = false;

                foreach (string line in lines)
                {


                    if (line.Contains("INSTRUCTIONS FOR COMPLETING CDTFA-531-A2, SCHEDULE A2 - Long Form"))
                    {
                        done = true;
                    }

                    if (!done)
                    {

                        if (line.Contains("City of Wheatland"))
                        {
                            done = true;
                        }

                        bool eff = false;
                        bool exp = false;
                        bool dis = false;



                        if (line.Contains("Eff"))
                        {
                            eff = true;
                        }

                        if (line.Contains("Exp"))
                        {
                            exp = true;
                        }

                        if (line.Contains("Discontinued"))
                        {
                            dis = true;
                        }

                        if (line.Contains("COUNTY"))
                        {
                            DataRow row = workTable.NewRow();

                            int county = line.IndexOf("COUNTY");
                            int end = county - 1;
                            int len = end - 0;

                            int stjst = 0;
                            string stj = "";

                            string County = line;

                            if (eff)
                            {
                                int stdt = line.IndexOf("(Eff");
                                int endt = line.IndexOf(")");
                                int lenf = endt - stdt;
                                string effful = line.Substring(stdt, lenf).Trim();
                                stdt = stdt + 5;
                                int lent = endt - stdt;

                                string effect = line.Substring(stdt, lent).Trim();
                                row["Effective"] = effect.Trim();
                                County = County.Replace(effful, "");
                            }

                            if (exp)
                            {
                                int stdt = line.IndexOf("(Exp");
                                int endt = line.IndexOf(")");
                                int lenf = endt - stdt;
                                string effful = line.Substring(stdt, lenf).Trim();
                                stdt = stdt + 5;
                                int lent = endt - stdt;

                                string effect = line.Substring(stdt, lent).Trim();
                                row["Expired"] = effect.Trim();
                                County = County.Replace(effful, "");
                            }

                            if (dis)
                            {
                                County = County.Replace("Discontinued ", "");

                            }

                            string rate = "";

                            if (line.Contains("."))
                            {
                                stjst = County.IndexOf(".") - 4;
                                int st = County.IndexOf(".");
                                rate = Get_Single_Text(st, County);
                                row["Rate"] = rate.Trim();
                            }

                            stj = Get_Number_From_Text(3, County);
                            row["STJ"] = stj.Trim();
                            if (stj.Length > 0)
                            {
                                County = County.Replace(stj, "");
                            }

                            if (rate.Length > 0)
                            {
                                County = County.Replace(rate, "");
                            }

                            County = County.Replace(".00", "");
                            County = County.Replace("$", "");
                            County = County.Replace(")", "");

                            row["County"] = County.Trim();
                            CountyIOn = County.Trim();

                            workTable.Rows.Add(row);
                        }

                        if (line.Contains("Unincorporated Area"))
                        {
                            DataRow row = workTable.NewRow();

                            string County = line;
                            if (line.Contains("739"))
                            {
                                County = County.Replace("Unincorporated Area", "YUBA COUNTY");
                                County = County.Insert(110, ".01");
                            }

                            if (line.Contains("724"))
                            {
                                County = County.Replace("Unincorporated Area", "SANTA CRUZ");
                                County = County.Insert(110, ".0175");
                            }
                            int county = County.IndexOf("County");
                            int end = county - 1;
                            int len = end - 0;

                            int stjst = 0;
                            string stj = "";



                            if (eff)
                            {
                                int stdt = County.IndexOf("(Eff");
                                int endt = County.IndexOf(")");
                                int lenf = endt - stdt;
                                string effful = County.Substring(stdt, lenf).Trim();
                                stdt = stdt + 5;
                                int lent = endt - stdt;

                                string effect = County.Substring(stdt, lent).Trim();
                                row["Effective"] = effect.Trim();
                                County = County.Replace(effful, "");
                            }

                            if (exp)
                            {
                                int stdt = County.IndexOf("(Exp");
                                int endt = County.IndexOf(")");
                                int lenf = endt - stdt;
                                string effful = County.Substring(stdt, lenf).Trim();
                                stdt = stdt + 5;
                                int lent = endt - stdt;

                                string effect = County.Substring(stdt, lent).Trim();
                                row["Expired"] = effect.Trim();
                                County = County.Replace(effful, "");
                            }

                            if (dis)
                            {
                                County = County.Replace("Discontinued ", "");

                            }

                            string rate = "";

                            if (County.Contains("."))
                            {
                                stjst = County.IndexOf(".") - 4;
                                int st = County.IndexOf(".");
                                rate = Get_Single_Text(st, County);
                                row["Rate"] = rate.Trim();
                            }

                            stj = Get_Number_From_Text(3, County);
                            row["STJ"] = stj.Trim();
                            if (stj.Length > 0)
                            {
                                County = County.Replace(stj, "");
                            }

                            if (rate.Length > 0)
                            {
                                County = County.Replace(rate, "");
                            }

                            County = County.Replace(".00", "");
                            County = County.Replace("$", "");
                            County = County.Replace(")", "");

                            row["County"] = County.Trim();
                            CountyIOn = County.Trim();

                            workTable.Rows.Add(row);
                        }

                        if (line.Contains("City of"))
                        {

                            DataRow row = workTable.NewRow();

                            string city = line.Replace("City of", "");

                            string rate = "";

                            if (city.Contains("(Eff"))
                            {
                                int cbeg = city.IndexOf("(");
                                int cend = city.IndexOf(")");
                                int clen = cend - cbeg;

                                string Eff = city.Substring(cbeg, clen);

                                city = city.Replace(Eff, "");

                                Eff = Eff.Replace("(Eff.", "");

                                row["Effective"] = Eff.Trim();
                            }

                            if (city.Contains("(Exp"))
                            {
                                int cbeg = city.IndexOf("(");
                                int cend = city.IndexOf(")");
                                int clen = cend - cbeg;

                                string Eff = city.Substring(cbeg, clen);

                                city = city.Replace(Eff, "");

                                Eff = Eff.Replace("(Exp.", "");

                                row["Expired"] = Eff.Trim();
                            }

                            int rbeg = city.IndexOf(".0");

                            if (rbeg > 0)
                            {
                                rate = Get_Single_Text(rbeg, city);
                                row["Rate"] = rate;
                            }

                            if (city.Contains("Discontinued"))
                            {
                                city = city.Replace("Discontinued", "");

                                int tlen = city.Length;
                                int clen = city.Length - 8;

                            }

                            city = city.Replace(")", "");
                            row["STJ"] = Get_Number_From_Text(3, city);

                            if (row["STJ"].ToString().Length > 0)
                            {
                                city = city.Replace(row["STJ"].ToString(), "");
                            }

                            if (rate.Length > 0)
                            {
                                city = city.Replace(rate, "");
                            }

                            city = city.Replace(".00", "");
                            city = city.Replace(".0", "");
                            row["City"] = city.Trim();
                            row["County"] = CountyIOn.Trim();

                            workTable.Rows.Add(row);
                        }

                        if (line.Contains("Town of"))
                        {
                            DataRow row = workTable.NewRow();

                            string city = line.Replace("Town of", "");

                            string rate = "";

                            if (city.Contains("(Eff"))
                            {
                                int cbeg = city.IndexOf("(");
                                int cend = city.IndexOf(")");
                                int clen = cend - cbeg;

                                string Eff = city.Substring(cbeg, clen);

                                city = city.Replace(Eff, "");

                                Eff = Eff.Replace("(Eff.", "");

                                row["Effective"] = Eff.Trim();
                            }

                            if (city.Contains("(Exp"))
                            {
                                int cbeg = city.IndexOf("(");
                                int cend = city.IndexOf(")");
                                int clen = cend - cbeg;

                                string Eff = city.Substring(cbeg, clen).Trim();

                                city = city.Replace(Eff, "");

                                Eff = Eff.Replace("(Exp.", "");

                                row["Expired"] = Eff.Trim();
                            }

                            int rbeg = city.IndexOf(".0");

                            if (rbeg > 0)
                            {
                                // MessageBox.Show(CountyIOn +line);
                                rate = Get_Single_Text(rbeg, city).Trim();
                                row["Rate"] = rate;
                            }

                            if (city.Contains("Discontinued"))
                            {
                                city = city.Replace("Discontinued", "");

                                int tlen = city.Length;
                                int clen = city.Length - 8;
                            }

                            row["STJ"] = Get_Number_From_Text(3, city);

                            city = city.Replace(")", "");
                            if (row["STJ"].ToString().Length > 0)
                            {
                                city = city.Replace(row["STJ"].ToString(), "");
                            }

                            if (rate.Length > 0)
                            {
                                city = city.Replace(rate, "");
                            }
                            city = city.Replace(".00", "");
                            city = city.Replace(".0", "");
                            row["City"] = city.Trim();
                            row["County"] = CountyIOn.Trim();

                            workTable.Rows.Add(row);
                        }
                    }
                }
            }

            dgv_District_Taxes.DataSource = workTable;
            cadtax = workTable.Copy();

            StringBuilder sb = new StringBuilder();

            IEnumerable<string> columnNames = workTable.Columns.Cast<DataColumn>().
                                              Select(column => column.ColumnName);
            sb.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in workTable.Rows)
            {
                IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                sb.AppendLine(string.Join(",", fields));
            }

            File.WriteAllText(ratesfile, sb.ToString());

            notifyIcon1.Icon = SystemIcons.Warning;
            notifyIcon1.BalloonTipText = "Loading Rates - Almost Done";
            notifyIcon1.BalloonTipTitle = "Rates";
            notifyIcon1.Visible = true;
            notifyIcon1.ShowBalloonTip(3000);

            foreach (DataGridViewColumn col in dgv_District_Taxes.Columns)
            {
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        string Get_Number_From_Text(int numlen, string line)
        {
            string num = "";
            bool done = false;
            int charion = 0;
            bool found = false;

            while (!done)
            {
                // MessageBox.Show(line);
                // MessageBox.Show(num+" "+charion.ToString()+" "+line.Substring(charion,1));


                if (charion < (line.Length - 1))
                {
                    string tnum = "X";


                    if (line.Substring(charion, 1).Equals("1"))
                    {
                        tnum = "1";

                    }

                    if (line.Substring(charion, 1).Equals("2"))
                    {
                        tnum = "2";
                    }
                    if (line.Substring(charion, 1).Equals("3"))
                    {
                        tnum = "3";
                    }
                    if (line.Substring(charion, 1).Equals("4"))
                    {
                        tnum = "4";
                    }
                    if (line.Substring(charion, 1).Equals("5"))
                    {
                        tnum = "5";
                    }
                    if (line.Substring(charion, 1).Equals("6"))
                    {
                        tnum = "6";
                    }
                    if (line.Substring(charion, 1).Equals("7"))
                    {
                        tnum = "7";
                    }
                    if (line.Substring(charion, 1).Equals("8"))
                    {
                        tnum = "8";
                    }
                    if (line.Substring(charion, 1).Equals("9"))
                    {
                        tnum = "9";
                    }
                    if (line.Substring(charion, 1).Equals("0"))
                    {
                        tnum = "0";
                    }

                    if (found)
                    {
                        if (tnum.Equals("X"))
                        {
                            if (num.Length == numlen)
                            {
                                done = true;
                            }
                            else
                            {
                                num = "";
                                found = false;
                            }
                        }
                        else
                        {
                            if (num.Length < numlen)
                            {
                                num = num + tnum;
                            }
                            else
                            {
                                num = "";
                                done = true;
                            }

                        }
                    }

                    if (!found & !tnum.Equals("X"))
                    {
                        //  MessageBox.Show("Found Number");
                        num = num + tnum;
                        found = true;
                    }

                    charion++;

                    if (charion == (line.Length - 1))
                    {
                        done = true;
                    }

                }
                else
                {
                    done = true;
                }

            }



            return num;
        }


        string Get_Single_Text(int rbeg, string line)
        {
            string rate = "";
            bool done = false;
            int end = line.Length - 1;
            for (int i = rbeg; i < end; i++)
            {
                if (i == rbeg)
                {
                    if (!(line.Substring(i, 1).Equals(" ") || line.Substring(i, 1).Equals(Environment.NewLine)))
                    {
                        rate = rate + line.Substring(i, 1);
                    }
                }
                else
                {
                    if (!done)
                    {
                        string ss = line.Substring(i, 1);
                        // MessageBox.Show(i.ToString()+" "+line+" char="+ss);
                        if (!(ss.Equals(" ") || ss.Equals(Environment.NewLine) || ss.Equals(".") || ss.Equals("$")))
                        {
                            rate = rate + line.Substring(i, 1);
                        }
                        else
                        {
                            done = true;
                        }
                    }
                }

            }


            return rate;
        }

        void Get_All_Containers()
        {
            bool done = false; ;
            Control p = this;
            int containerion = -1;

            List<Control> conts = new List<Control>();

            do
            {
                conts = Get_Containers(p);

                if (conts.Count == 0)
                {
                    done = true;
                }

                // add to global list

                foreach (Control c in conts)
                {

                    containers.Add(c);

                }

                if (containerion < containers.Count)
                {
                    containerion++;
                    // MessageBox.Show("containerion = " + containerion.ToString() + " containers.count = "+containers.Count.ToString());
                    p = containers[containerion];
                }
                else
                {
                    done = true;
                }


            } while (!done);

        }


        List<Control> Get_Containers(Control container)
        {
            List<Control> newcons = new List<Control>();

            foreach (Control con in container.Controls)
            {
                if (con.Controls.Count > 0)
                {

                    newcons.Add(con);
                }
            }

            return newcons;
        }

        void Save_All_Data()
        {

            StringBuilder lines = new StringBuilder();

            foreach (Control c in containers)
            {
                //MessageBox.Show(c.Name);

                foreach (Control cc in c.Controls)
                {
                    // MessageBox.Show(cc.Name);
                    if (cc.Name.Length > 2)
                    {
                        if (cc.GetType().ToString().Contains(".TextBox") || cc.GetType().ToString().Contains(".ComboBox"))
                        {
                            lines.Append(cc.Name + " |" + cc.Text + Environment.NewLine);
                        }

                        if (cc.GetType().ToString().Contains(".Label"))
                        {
                        }

                        if (cc.GetType().ToString().Contains(".DateTimePicker"))
                        {

                            lines.Append(cc.Name + " |" + cc.Text + Environment.NewLine);
                        }

                        if (cc.GetType().ToString().Contains(".DataGridView") && cc.Name.Contains("dgv_District_Taxes"))
                        {
                            lines.Append("DataGridView " + cc.Name + " | Begin" + Environment.NewLine);

                            string dgv = cc.Name + ".Rows";
                            ((DataGridView)cc).AllowUserToAddRows = false;
                            foreach (DataGridViewRow row in ((DataGridView)cc).Rows)
                            {
                                string line = "";
                                for (int i = 0; i < ((DataGridView)cc).Columns.Count; i++)
                                {
                                    if (i < ((DataGridView)cc).Columns.Count - 1)
                                    {
                                        line = line + row.Cells[i].Value.ToString() + " | ";
                                    }
                                    else
                                    {
                                        line = line + row.Cells[i].Value.ToString() + Environment.NewLine;
                                    }
                                }

                                lines.Append(line);
                            }

                            lines.Append("DataGridView " + cc.Name + " | End" + Environment.NewLine);

                        }

                    }
                }
            }


            lines.Append("File Stamp Date |" + DateTime.Now.ToLongDateString() + Environment.NewLine);

            SaveFileDialog sfile = new SaveFileDialog();

            sfile.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";

            string[] delim = { Environment.NewLine, "\n" };

            string[] slines = lines.ToString().Split(delim, StringSplitOptions.None);

            if (sfile.ShowDialog() == DialogResult.OK)
            {
                File.WriteAllLines(sfile.FileName, slines);

            }
        }

        void Restore_All_Data()
        {

            OpenFileDialog sfile = new OpenFileDialog();

            sfile.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";

            if (sfile.ShowDialog() == DialogResult.OK)
            {

                string[] lines = File.ReadAllLines(sfile.FileName);

                Boolean dgvsection = false;
                string dgv = "";

                foreach (string line in lines)
                {

                    int start = line.IndexOf("|");

                    int len = line.Length - start - 1;

                    string name = "";

                    if (line.Length > 0)
                    {
                        name = line.Substring(0, (start - 1));
                    }

                    string value = "";

                    if (len > 0)
                    {
                        value = line.Substring((start + 1), len);
                    }

                    // MessageBox.Show(name + " " + value);

                    if (name.Contains("DataGridView"))
                    {
                        if (value.Contains("Begin"))
                        {
                            dgvsection = true;
                            dgv = name.Replace("DataGridView", "").Trim();

                            foreach (Control c in containers)
                            {
                                foreach (Control cc in c.Controls)
                                {
                                    // MessageBox.Show("Control "+cc.Name+" name= "+name);
                                    if (cc.Name.Contains(dgv) && dgv.Length > 3)
                                    {

                                        DataTable dgvdata = (DataTable)((DataGridView)cc).DataSource;
                                        if (dgvdata != null)
                                        {
                                            dgvdata.Rows.Clear();
                                            ((DataGridView)cc).DataSource = dgvdata;
                                        }
                                    }
                                }
                            }

                        }

                        if (value.Contains("End"))
                        {
                            dgvsection = false;
                        }
                    }
                    else
                    {
                        if (!dgvsection)
                        {
                            // Limit Containers if you just want to restore within a part of the program
                            foreach (Control c in containers)
                            {
                                foreach (Control cc in c.Controls)
                                {
                                    // MessageBox.Show("Control "+cc.Name+" name= "+name);
                                    if (cc.Name.Contains(name) && name.Length > 3)
                                    {
                                        cc.Text = value.ToString();
                                    }
                                }
                            }
                        }

                        if (dgvsection)
                        {
                            foreach (Control c in containers)
                            {
                                foreach (Control cc in c.Controls)
                                {
                                    if (cc.Name.Contains(dgv))
                                    {
                                        DataTable dgvdata = (DataTable)((DataGridView)cc).DataSource;

                                        DataRow newrow = dgvdata.NewRow();

                                        string[] cols = line.Split('|');

                                        for (int i = 0; i < cols.Length; i++)
                                        {
                                            cols[i] = cols[i].Trim();
                                            if (cols[i].Length > 0)
                                            {
                                                // MessageBox.Show(cols[i]);
                                                newrow[i] = cols[i];
                                                // MessageBox.Show(newrow[i].ToString());
                                            }
                                        }
                                        dgvdata.Rows.Add(newrow);

                                        ((DataGridView)cc).DataSource = dgvdata;
                                    }
                                }
                            }
                        }


                    }
                }



            }
            // Put Calcs Here        


        }


        void Clear_Sales_Datatables()
        {
            cpst_data.Clear();
            sales_data.Clear();
            GC.Collect();
        }

        void Clear_CL_Datatables()
        {
            locbil.Clear();
            locbil_dd.Clear();
            locbil_de.Clear();
            locbil_de_ca.Clear();
            locbil_df.Clear();
            locbil_fdo.Clear();
            locbil_fdp.Clear();
            CFNData.Clear();
            CFNDataExt.Clear();
            GC.Collect();
        }

        void CLear_Datatables()
        {
            cpst_data.Clear();
            sales_data.Clear();
            locbil.Clear();
            locbil_dd.Clear();
            locbil_de.Clear();
            locbil_de_ca.Clear();
            locbil_df.Clear();
            locbil_fdo.Clear();
            locbil_fdp.Clear();
            CFNDataExt.Clear();
            CFNData.Clear();
            GC.Collect();
        }

        void Clear_DataGridviews()
        {
            // dgv_Sales_Data.DataSource = null;
            // dgv_CPST_Data.DataSource = null;
            dgv_Locbil_Data.DataSource = null;
            dgv_CFN_Data.DataSource = null;
            dgv_CFN_Extended.DataSource = null;
            GC.Collect();
        }

        string Get_File(string title, string search)
        {
            string filename = " ";

            OpenFileDialog file = new OpenFileDialog();

            file.Title = title;
            file.FileName = search;

            if (file.ShowDialog() == DialogResult.OK)
            {
                filename = file.FileName;
            }

            return filename;
        }

        private void button_Get_Sales_File_Click(object sender, EventArgs e)
        {
            salesfile = Get_File("Sales File", "sales*");

            tb_Sales_File.Text = salesfile;
        }

        private void button_Get_CPST_File_Click(object sender, EventArgs e)
        {
            cpstfile = Get_File("CPST File", "cpst*");
            tb_CPST_File.Text = cpstfile;
        }

        private void button_Save_Data_Click(object sender, EventArgs e)
        {
            Save_All_Data();
        }

        private void button_Restore_Data_Click(object sender, EventArgs e)
        {
            Restore_All_Data();
        }


        string clean_Special_Charachters(string str)
        {
            string retstr = "";

            foreach(char ch in str)
            {

                // ch is a number
                if((int)ch<59 && (int)ch > 46)
                {
                    retstr = retstr + ch;
                }

                //ch is an alphanumeric capital

                if ((int)ch < 91 && (int)ch > 64)
                {
                    retstr = retstr + ch;
                }

                //ch is an alphanumeric lower

                if ((int)ch < 123 && (int)ch > 96)
                {
                    retstr = retstr + ch;
                }

                //ch is a space

                if ((int)ch == 32)
                {
                    retstr = retstr + ch;
                }
                
            }
            
            return retstr;
        }





        void Get_phvalue_All_Citys()
        {
                   

            foreach(DataRow row in workTable.Rows)
            {
                if(row["City"].ToString().Length > 0)
                {
                    string dis = clean_Special_Charachters(row["City"].ToString().Trim());

                    MessageBox.Show(dis.Trim() + " =" + alpharate((dis+"b").Trim(),(dis+"a").Trim()).ToString("###.#####"));
                }
            }

       
        }

        bool Match_Chars_String(string target, string str)
        {
            bool match = false;
            int cntsequence = 0;
            int cnttotal = 0;
            int cntsections = 0;

            string result = "";
            string missing = "";

            target = target.Trim().ToUpper();
            str = str.Trim().ToUpper();

            if (target.Contains(str))
            {
                return true;
            }

            if(str.Length > target.Length)
            {
                return false;
            }

            // get num of chars in seq correct excluding middle wrong chars
            for(int ch = 0; ch<str.Length; ch++)
            {
                if (target.Contains(str.Substring(0, ch)))
                {
                    cntsequence = ch;
                }
            }

            result = str.Substring(0, cntsequence);

            missing = str.Substring(cntsequence, str.Length-cntsequence);

            string missingLetter = missing.Substring(0, 1);

            string restof = missing.Substring(1, missing.Length - 1);

            string shoulda = result + missingLetter + restof;

            string withoutLetter = result + restof;

            if (target.Contains(withoutLetter))
            {

                MessageBox.Show(" Result ="+withoutLetter+" Target = "+target+" Missing " + missingLetter +" at "+missing+ " Should Be " + shoulda);

                return true;
            }

            
            int temp = 0;

            for(int ch=str.Length-1; ch>cntsequence; ch--)
            {

                int l = str.Length - ch;
                if (target.Contains(str.Substring(ch, l)))
                {
                    temp= l;  
                }
            }

            cntsequence = cntsequence + temp;

            // get num of chars in target
            for (int ch = 0; ch < str.Length; ch++)
            {
                if (target.Contains(str.Substring(ch, 1)))
                {
                    cnttotal++;
                }
            }

            int poss = 0;
                     
            
            //for (int ch = 0; ch < str.Length; ch++)
            //{
            //    int len = str.Length / (ch+1);
                        

            //    for (int sub = 0; sub < (str.Length - len); sub++)
            //    {
            //        poss=poss+len;
            //        if (target.Contains(str.Substring(sub,len)))
            //        {
            //            cntsections=cntsections+len;
            //        }
            //    }
            //}

            decimal ratiosections = 0.00m;

            ratiosections = decimal.Divide(cntsequence, str.Length)+decimal.Divide(1,str.Length);
            
            //ratiosections = decimal.Divide(cntsections,poss)*(str.Length/target.Length);

            //if (str.Length > 0 && target.Length > 0 && poss > 0 && !str.Contains("NAME") && !str.Contains("YUBA") && !str.Contains("SUTTER") && !str.Contains("BUTTE"))
            //{
            //    MessageBox.Show("Target = " + target + " Find = " + str + " Ratio = " + ratiosections.ToString("###.######") + " Poss =" + poss.ToString() + " Sections =" + cntsections.ToString() + " Sequence =" + cntsequence.ToString() + " Total =" + cnttotal);
            //}

            if (target.Contains("SANTA B") && str.Contains("SANTA "))
            {

                MessageBox.Show(target + " missing " + missing + " in " + result + " Should be " + str);
                MessageBox.Show("Target = " + target + " Find = " + str + " Ratio = " + ratiosections.ToString("###.######") + " Poss =" + poss.ToString() + " Sections =" + cntsections.ToString() + " Sequence =" + cntsequence.ToString() + " Total =" + cnttotal);

            }

            return false;

        }

        int Count_String_Within_String(string target, string str)
        {
            if (str.Length < 1)
            {
                return 0;
            }

            target = target.ToUpper().Trim();
            target = clean(target);
            target = clean_Special_Charachters(target);
            target = NoSpace(target);

            str = str.ToUpper().Trim();
            str = clean(str);
            str = clean_Special_Charachters(str);
            str = NoSpace(str);

            int count = 0;

            if (target.Contains(str))
            {                

                if ((str.Length * 2) < target.Length)
                {
                    int half = target.Length / 2;
                    string first = target.Substring(0, half);
                    string second = target.Substring(half, half);

                    int poss = target.Length / str.Length;

                    if( poss< 3)
                    {
                        if(first.Contains(str) && second.Contains(str))
                        {
                            return 2;
                        }
                        else
                        {
                            return 1;
                        }
                    }
                    else
                    {
                        int secLength = str.Length;

                        string[] sections = new string[poss];

                        int start = 0;

                        for (int s = 0; s < poss; s++)
                        {
                            sections[s]= target.Substring(start, secLength);
                            start = start + secLength;
                        }

                        for(int s=0; s< poss; s++)
                        {
                            if (sections[s].Contains(str))
                            {
                                count++;
                            }  
                        }

                        return count;
                    }
                    
                }
                else
                {
                    return 1;
                }
                             

            }
            else
            {
                return 0;
            }

            
        }

        string Match_City_With_District(string str)
        {

            string city = "";

            foreach(DataRow row in workTable.Rows)
            {
                string dis = clean(clean_Special_Charachters(row["City"].ToString().ToUpper()));
                string cty = clean(clean_Special_Charachters(row["County"].ToString().ToUpper()));

                cty = cty.Replace("COUNTY", "").Trim();

                str = str.Replace("CO.", "COUNTY");

                str = str.Replace("/", " ");


                str = clean(clean_Special_Charachters(str.ToString().ToUpper()));

                str = str.Replace("CITY OF", "");

                str=str.Replace("BARARA", "BARBARA");
                str=str.Replace("TEHEMA", "TEHAMA");
                str=str.Replace("GLEN ", "GLENN");
                str = str.Replace("AMADORE", "AMADOR");
                str = str.Replace("OBISP)", "OBISPO)");
                str = str.Replace("SAN LOUIS", "SAN LUIS");
                str = str.Replace("MANIFEE", "MENIFEE");


                if (str.Contains("BAKERS"))
                {
                    return "BAKERSFIELD";
                }

                if (str.Contains("CALEX"))
                {
                    return "CALEXICO";
                }

                if (str.Contains("COAL"))
                {
                    return "COALINGA";
                }

                if (str.Contains("CARMEL") )
                {

                    return "CARMEL-BY-THE-SEA";
                }

                if (str.Contains("SAN LUIS O") && str.Contains("ARROYO"))
                {

                    return "ARROYO GRANDE";
                }

                if (str.Contains("SAN LUIS O") && str.Contains("GROVER"))
                {

                    return "GROVER BEACH";
                }

                if (str.Contains("CARPIN"))
                {
                    return "CARPINTERIA";
                }

                if (str.Contains("WOODL"))
                {
                    return "WOODLAKE";
                }

                if (str.Contains("BELMONT"))
                {
                    return "BELMONT";
                }
                if (str.Contains("CHULA"))
                {
                    return "CHULA VISTA";
                }

                if (str.Contains("WILDO"))
                {
                    return "WILDOMAR";
                }

                if (str.Contains("TEMEC"))
                {
                    return "TEMECULA";
                }

                if (str.Contains("MURRI"))
                {
                    return "MURRIETA";
                }

                if (str.Contains("COACH"))
                {
                    return "COACHELLA";
                }

                if (str.Contains("WESTMIN"))
                {
                    return "WESTMINSTER";
                }

                if (str.Contains("STANTO"))
                {
                    return "STANTON";
                }

                if(str.Contains("SEAL B"))
                {
                    return "SEAL BEACH";
                }

                if (str.Contains("SAN LUIS O") && str.Contains("ATASC"))
                {

                    return "ATASCADERO";
                }

                if (str.Contains("SAN LUIS O") && str.Contains("MORRO"))
                {

                    return "MORRO BAY";
                }

                if (str.Contains("SAN LUIS O") && str.Contains("PASO ROBLES"))
                {

                    return "PASO ROBLES";
                }


                if (str.Contains("SAN LUIS O") && str.Contains("PISMO B"))
                {

                    return "PISMO BEACH";
                }

                if (str.Contains("LINDSAY"))
                {
                    return "LINDSAY";
                }

                if (str.Contains("GALT"))
                {
                    return "GALT";
                }

                if (str.Contains("RANCHO CORD"))
                {
                    return "RANCHO CORDOVA";
                }

                if (str.Contains("SACRAMENTO") && str.Contains("YOLO"))
                {
                    return "WEST SACRAMENTO";
                }

                if (str.Contains("ALAMEDA") && str.Contains("HAYWARD"))
                {
                    return "HAYWARD";
                }


                if (str.Contains("SONOMA") && str.Contains("SANTA ROSA"))
                {
                    return "SANTA ROSA";
                }

                if (str.Contains("GUSTINE") && str.Contains("MERCED"))
                {
                    return "GUSTINE";
                }

                if (str.Contains("CHOW") && str.Contains("MADERA"))
                {
                    return "CHOWCHILLA";
                }

                if (str.Contains("GLEN") && str.Contains("ORLAND"))
                {
                    return "ORLAND";
                }

                if (str.Contains("TEH") && str.Contains("CORN"))
                {
                    return "CORNING";
                }

                if(str.Contains("SANTA MARIA"))
                {
                    return "SANTA MARIA";
                }

                if (str.Contains("KING "))
                {
                    return "KING CITY";
                }

                if (str.Contains("LAKE TAHOE"))
                {
                    return "SO. LAKE TAHOE";
                }


                if (str.Contains("PLACERVIL"))
                {
                    return "PLACERVILLE";
                }


                if (str.Contains("REDWOOD"))
                {
                    return "REDWOOD CITY";
                }

                if (str.Contains("RIO DELL"))
                {
                    return "RIO DELL";
                }

                if (str.Contains("ATWATER"))
                {
                    return "ATWATER";
                }

                if (str.Contains("STOW"))
                {
                    return "BARSTOW";
                }

                if (str.Contains("INDIO"))
                {
                    return "INDIO";
                }

                if (str.Contains("MENIF"))
                {
                    return "MENIFEE";
                }

                if (str.Contains("BANOS"))
                {
                    return "LOS BANOS";
                }

                if (str.Contains("MAMM"))
                {
                    return "MAMMOTH LAKES";
                }

                if(Count_String_Within_String(str,"VENTURA") == 2)
                {
                    return "VENTURA";
                }

                if (Count_String_Within_String(str, dis) > 1)
                {
                   
                   // MessageBox.Show("Target Duplicate = " + str + " City = " + dis);
                    return dis;
                }

                if(str.Contains("SAN MATEO COUNTY"))
                {
                    //MessageBox.Show("Got Mateo County = "+ str);
                    return "";
                }

                if (str.Contains("SAN DIEGO COUNTY") && str.Trim().Length<17)
                {
                    return "";
                }

                if (str.Contains("SANTA CLARA COUNTY") && str.Trim().Length<20)
                {
                    return "";
                }

                if (str.Contains("SANTA BARBARA COUNTY") && str.Trim().Length <22)
                {
                    return "";
                }

                if (str.Contains("WEED"))
                {
                    return "WEED";
                }

                string[] words = str.Trim().Split(' ');

                if (words.Length < 2)
                {
                    if (str.Contains(dis))
                    {
                        return words[0];
                    }
                }

                if (str.Contains(dis) && dis.Length >0 && str.Contains(cty))
                {
                   
                    //foreach(string word in words)
                    //{
                    //    MessageBox.Show(word);
                    //}
                    
                    if(words.Length > 1)
                    {

                        if(words[0].Contains(dis) && words[1].Contains(dis))
                        {
                            return dis;
                        }

                        if(words[1].Contains("COUNTY") && words.Length == 2 && city.Length == 0)
                        {
                            return string.Empty;
                        }
                        else
                        {
                                                        
                            city = dis;

                          
                        }
                                  
                        


                    }


                //    MessageBox.Show("string passed = " + str + " district = " + city+" Words = "+words.Length.ToString() );
                                      
                }
            }


            return city;

        }

        string Match_County_With_District(string str)
        {
            string county = "";

            str = str.Replace("BARARA", "BARBARA");
            str = str.Replace("TEHEMA", "TEHAMA");
            str = str.Replace("GLEN ", "GLENN");
            str = str.Replace("AMADORE", "AMADOR");
            str = str.Replace("OBISP)", "OBISPO)");
            str = str.Replace("HUMBOLT", "HUMBOLDT");
            str = str.Replace("MENDOCINA", "MENDOCINO");
            str = str.Replace("CO.", "COUNTY");
            //str = str.Replace("COUNTY", "");
            //str = str.Replace(" CO ", "");

            foreach (DataRow row in workTable.Rows)
            {
                string dis = clean(clean_Special_Charachters(row["County"].ToString().ToUpper()));
                dis = dis.Trim();
                

                str = clean(clean_Special_Charachters(str.ToString().ToUpper()));

                //if (str.Contains("MATEO") && dis.Contains("MATEO"))
                //{
                //    MessageBox.Show("str = " + str + " did = " + dis);
                    
                //}

                    if (str.Contains(dis))
                {
                    county = dis;
                }
            }


            return county;
        }


        decimal alpharate(string str1, string str2)
        {
            if(str1.Length <1 || str2.Length < 1)
            {
                return 0.00m;
            }

            int chars1 = str1.Length;
            int chars2 = str2.Length;
            int charsuse = 0;

            int corchars = 0;

            decimal rating = 0.00m;
            decimal rating2 = 0.00m;

            if(chars1 > chars2)
            {
                // check if string is added to
                if (str1.Contains(str2))
                {
                    if ((chars1 - chars2) < chars2)
                    {

                        return 1.00m;
                    }
                }

                charsuse = chars2;
            }
            else
            {
                if (str2.Contains(str1))
                {
                    if ((chars2 - chars1) < chars1)
                    {
                        return 1.00m;
                    }
                }


                charsuse = chars1;
            }

            for(int c = 0; c <charsuse; c++)
            {
                if((int)str1[c] == (int)str2[c])
                {
                    corchars++;
                }
            }

            if (corchars == 0)
            {
                return 0.00m;
            }


            rating2 = (decimal)alphavalue(str1) / (decimal)alphavalue(str2);

            rating = ((decimal)corchars / (decimal)charsuse) *rating2;

           
                        
            //MessageBox.Show("str1="+str1+"  str2="+str2+"  corchars = " + corchars.ToString() + " charsuse = " + charsuse.ToString()+" rating = "+rating.ToString("##.######")+" rating2 = "+rating2.ToString("##.######"));

            return rating;
        }

        int alphavalue(string str)
        {
            int phvalue = 0;
          

            foreach (char ch in str)
            {
                phvalue = phvalue + (int)ch ;

            }

            return phvalue;
        }

        int alphaordervalue(string str)
        {
            int phvalue = 0;
            int order = 0;

            foreach(char ch in str)
            {
                phvalue = phvalue + (int)ch+order;
                order++;
            }

            return phvalue;

        }
                
        Jurisdiction Parse_City_County(string citycnty)
        {
            Jurisdiction js = new Jurisdiction();

            js.County = "";
            js.City = "";

            if(citycnty.Length == 0)
            {
                return js;
            }


            if (citycnty.Substring(0, 2).Equals("N "))
            {
                citycnty = "North" + citycnty.Substring(2, citycnty.Length - 3);
            }

            if (citycnty.Substring(0, 2).Equals("S "))
            {
                citycnty = "South" + citycnty.Substring(2, citycnty.Length - 3);
            }

            if (citycnty.Substring(0, 2).Equals("W "))
            {
                citycnty = "West" + citycnty.Substring(2, citycnty.Length - 3);
            }

            if (citycnty.Substring(0, 2).Equals("E "))
            {
                citycnty = "East" + citycnty.Substring(2, citycnty.Length - 3);
            }

            citycnty = citycnty.ToUpper();

            citycnty = citycnty.Replace("CITY OF", "");
            citycnty = citycnty.Replace(",", "");

            citycnty = citycnty.Replace(" N ", "North");
            citycnty = citycnty.Replace(" S ", "South");
            citycnty = citycnty.Replace(" E ", "East");
            citycnty = citycnty.Replace(" W ", "West");

            citycnty = citycnty.Replace("(N ", "North");
            citycnty = citycnty.Replace("(S ", "South");
            citycnty = citycnty.Replace("(E ", "East");
            citycnty = citycnty.Replace("(W ", "West");


            citycnty = citycnty.Replace("N. ", "North");
            citycnty = citycnty.Replace("S. ", "South");
            citycnty = citycnty.Replace("E. ", "East");
            citycnty = citycnty.Replace("W. ", "West");

            citycnty = citycnty.Replace("CO ", "COUNTY");
            citycnty = citycnty.Replace("CO. ", "COUNTY");
            citycnty = citycnty.Replace("(CO ", "COUNTY");
            citycnty = citycnty.Replace("CO) ", "COUNTY");
            citycnty = citycnty.Replace("CO)", "COUNTY");

         

            citycnty = citycnty.Replace("SAN ", "SAN");
            citycnty = citycnty.Replace("SANTA ", "SANTA");
            citycnty = citycnty.Replace("EL ", "EL");
            citycnty = citycnty.Replace("LOS ", "LOS");
            citycnty = citycnty.Replace("LA ", "LA");
            citycnty = citycnty.Replace("LAKE ", "LAKE");
            citycnty = citycnty.Replace("SO. ", "SO.");
            citycnty = citycnty.Replace("RANCHO ", "RANCHO");
            citycnty = citycnty.Replace("UNION ", "UNION");
            citycnty = citycnty.Replace("ANGELS ", "ANGELS");
            citycnty = citycnty.Replace("PLEASANT ", "PLEASANT");
            citycnty = citycnty.Replace("LAKE ", "LAKE");
            citycnty = citycnty.Replace("RIO ", "RIO");
            citycnty = citycnty.Replace("CULVER ", "CULVER");
            citycnty = citycnty.Replace("HUNTINGTON ", "HUNTINGTON");

            citycnty = citycnty.Replace("PICO ", "PICO");
            citycnty = citycnty.Replace("FE ", "FE");
            citycnty = citycnty.Replace("CORTE ", "CORTE");
            citycnty = citycnty.Replace(" GATE", "GATE");
            citycnty = citycnty.Replace("FORT ", "FORT");
            citycnty = citycnty.Replace("POINT ", "POINT");
            citycnty = citycnty.Replace("MAMMOTH ", "MAMMOTH");
            citycnty = citycnty.Replace("DEL ", "DEL");
            citycnty = citycnty.Replace("REY ", "REY");
            citycnty = citycnty.Replace("KING ", "KING");
            citycnty = citycnty.Replace("PACIFIC ", "PACIFIC");
            citycnty = citycnty.Replace("SAND ", "SAND");
            citycnty = citycnty.Replace("ST. ", "ST.");

            citycnty = citycnty.Replace(" VALLEY", "VALLEY");
            citycnty = citycnty.Replace("FOUNTAIN ", "FOUNTAIN");
            citycnty = citycnty.Replace("GARDEN ", "GARDEN");

            citycnty = citycnty.Replace("CATHEDRAL ", "CATHEDRAL");
            citycnty = citycnty.Replace("PALM ", "PALM");
            citycnty = citycnty.Replace("JUAN ", "JUAN");
            citycnty = citycnty.Replace("YUCCA ", "YUCCA");
            citycnty = citycnty.Replace("CHULA ", "CHULA");
            citycnty = citycnty.Replace(" CITY", "CITY");
            citycnty = citycnty.Replace("ARROYO ", "ARROYO");
            citycnty = citycnty.Replace(" BEACH", "BEACH");
            citycnty = citycnty.Replace(" BAY", "BAY");
            citycnty = citycnty.Replace("LUIS ", "LUIS");
            citycnty = citycnty.Replace("PASO ", "PASO");
            citycnty = citycnty.Replace("EAST ", "EAST");
            citycnty = citycnty.Replace("PALO ", "PALO");
            citycnty = citycnty.Replace("MT. ", "MT.");
            citycnty = citycnty.Replace("PORT ", "PORT");
            citycnty = citycnty.Replace("WEST ", "WEST");
            citycnty = citycnty.Replace("EAST ", "EAST");
            citycnty = citycnty.Replace("SOUTH ", "SOUTH");
            citycnty = citycnty.Replace("NORTH ", "NORTH");
            citycnty = citycnty.Replace(" PARK", "PARK");
            citycnty = citycnty.Replace(" BLUFF", "BLUFF");
            citycnty = citycnty.Replace("MT ", "MT");

            citycnty = citycnty.Replace(",CITY", "");
            citycnty = citycnty.Replace("TOWN OF ", "");


            string[] words = citycnty.Split(' ');
            string next = "";

            if (words.Length > 1)
            {
                if (words[1].ToUpper().Equals("COUNTY") || words[1].ToUpper().Equals("CO"))
                {
                    js.County = words[0];
                    js.City = "";

                    return js;
                }
            }

                for (int w = 0; w < words.Length; w++)
                {
                    string wcnty = Get_County_With_City(words[w]);

                //if (citycnty.ToUpper().Contains("COALINGA"))
                //{

                //    MessageBox.Show("citycnty = " + citycnty + " City = " + js.City + " County = " + js.County, "Jurisdiction", MessageBoxButtons.OK);
                //}


                if (wcnty.Length > 0)
                    {

                        js.County = wcnty;

                        if (words.Length > 2 && !(words[1].ToString().ToUpper().Equals("COUNTY") || words[1].ToString().ToUpper().Equals("CO")))
                        {
                            js.City = words[w];
                        }
                        else
                        {
                            js.City = "";
                        }

                   
                        return js;
                    }

                    if (w < (words.Length - 1))
                    {
                        next = words[w] + words[w + 1];
                        wcnty = Get_County_With_City(next);

                        if (wcnty.Length > 0)
                        {

                            js.County = wcnty;
                            js.City = words[w];

                         //   MessageBox.Show("citycnty = " + citycnty + " City = " + js.City + " County = " + js.County, "Jurisdiction", MessageBoxButtons.OK);


                            return js;
                        }

                    }


                }
            
            return js;
        }


        void Calc_Data()
        {
            Clear_Sales_Datatables();
            Reset_Totals();

            string Sales_File = salesfile.ToUpper().Replace(".CSV","")+" Sales File"+dtp_Date_End.Value.ToString("yyyyMMdd")+".XLSX";
            archive = Sales_File;
            string CPST_File = cpstfile.ToUpper().Replace(".CSV", "") + " CPST File"+".XLSX";

            decimal OOSSales = 0.00m;

            if (File.Exists(salesfile))
            {
                
                string[] lines = Get_CSV_Data(salesfile);

                sales_data = Load_Data_To_Datatable(lines);

                label_Sales_Records.Text = "Number of Records = " + sales_data.Rows.Count.ToString();

            }

            if (CSV_Bad_Rows != null)
            {
                foreach (int row in CSV_Bad_Rows)
                {
                    Auto_Correct_Row(sales_data, row);
                }

                CSV_Bad_Rows.Clear();
            }
               
                       

            dgv_Sales_Data.DataSource = sales_data;

            if (File.Exists(cpstfile))
            {
                string[] lines = Get_CSV_Data(cpstfile);

                cpst_data = Load_Data_To_Datatable(lines);

                label_CPST_Records.Text = "Number of Records = " + cpst_data.Rows.Count.ToString();

            }

            if (CSV_Bad_Rows != null)
            {
                foreach (int row in CSV_Bad_Rows)
                {
                    Auto_Correct_Row(cpst_data, row);
                }

                CSV_Bad_Rows.Clear();
            }

            dgv_CPST_Data.DataSource = cpst_data;

            int sales_r = 0;

            foreach (DataRow row in sales_data.Rows)
            {
                sales_r++;

                notifyIcon1.Icon = SystemIcons.Hand;
                notifyIcon1.BalloonTipText = "Sales Record " + sales_r.ToString();
                notifyIcon1.BalloonTipTitle = "Sales File Calc";
                notifyIcon1.ShowBalloonTip(500);

                string prod = "";
                string inv = "";
                decimal s = 0.00m;
                decimal set = 0.00m;
                decimal set_tot = 0.00m;
                decimal stax = 0.00m;
                decimal quan = 0.00m;
                int pcat = 0;
                decimal salesamt = 0.00m;

                bool outofcalif = true;
                bool cardlockexcluded = false;

                int column_city = 0;
                int column_cnty = 0;
                int indicator_city_cnty = 0;
                string separator_city_cnty = tb_City_County_Separator.Text.ToString().Trim();
                bool use_city_county = false;
                string salescity = "";
                string salescnty = "";


                int.TryParse(tb_Column_City.Text, out column_city);
                int.TryParse(tb_Column_County.Text, out column_cnty);
                int.TryParse(tb_City_County_Order.Text, out indicator_city_cnty);

                String citycnty = "";

                if (column_city > 0 && column_cnty > 0)
                {
                    citycnty = row[column_city].ToString();

                    use_city_county = true;

                    if (column_city == column_cnty && indicator_city_cnty == 0)
                    {
                        use_city_county = false;

                    }
                    if (indicator_city_cnty == 3)
                    {
                        use_city_county = true;
                    }

                    if (column_city != column_cnty)
                    {
                        salescity = row[column_city].ToString();
                        salescnty = row[column_cnty].ToString();
                    }
                    else
                    {
                        if (use_city_county)
                        {
                            citycnty = row[column_city].ToString();

                            salescity = Match_City_With_District(citycnty);

                            if (salescity.Length > 0)
                            {
                                salescnty = Get_County_With_City(salescity);
                            }
                            else
                            {
                                salescnty = Match_County_With_District(citycnty);
                            }

                            if(salescity.Length >0 && salescnty.Length < 1)
                            {
                                salescnty = Get_County_With_City(salescity);
                            }

                            //Jurisdiction js = Parse_City_County(citycnty);

                            //if (js.County.Length > 0)
                            //{
                            //    salescnty = js.County;
                            //    salescity = js.City;
                            //}

                            //if(js.County.Length < 1)
                            //{ 

                            //    if (indicator_city_cnty == 1)
                            //    {
                            //        int start = 0;
                            //        start = citycnty.IndexOf(separator_city_cnty);
                            //        if (start > 0)
                            //        {
                            //            int end = start - 1;
                            //            int len = citycnty.Length - start;
                            //            start = start + 1;
                            //            len = len - 1;
                            //            salescity = citycnty.Substring(0, end);
                            //            salescnty = citycnty.Substring(start, len);
                            //        }

                            //    }

                            //    if (indicator_city_cnty == 2)
                            //    {
                            //        int start = 0;
                            //        start = citycnty.IndexOf(separator_city_cnty);
                            //        if (start > 0)
                            //        {
                            //            int end = start - 1;
                            //            int len = citycnty.Length - start;
                            //            start = start + 1;
                            //            len = len - 1;
                            //            salescnty = citycnty.Substring(0, end);
                            //            salescity = citycnty.Substring(start, len);
                            //        }

                            //    }
                            //    if (indicator_city_cnty == 3)
                            //    {
                            //        citycnty = citycnty.ToUpper();
                            //        string[] words = citycnty.Split(' ');

                            //        string tword = "";
                            //        string tnext = "";

                            //        for (int w = 0; w < words.Length; w++)
                            //        {
                            //            if (salescnty.Length < 1)
                            //            {
                            //                tword = clean(words[w]);
                            //                string testcnty = Get_County_With_City(tword);

                            //                if ((w + 1) < words.Length)
                            //                {
                            //                    tnext = clean(words[w + 1]);
                            //                }
                            //                else
                            //                {
                            //                    tnext = "";
                            //                }



                            //                if (testcnty.Length > 0 && tnext.ToUpper().Equals("COUNTY"))
                            //                {
                            //                    salescnty = testcnty;
                            //                    salescity = "";
                            //                }
                            //                else
                            //                {
                            //                    if (testcnty.Length > 0)
                            //                    {
                            //                        salescnty = testcnty;
                            //                        salescity = tword;
                            //                    }
                            //                }

                            //            }
                            //        }
                            //    }
                            //    else
                            //    {
                            //        salescity = js.City;
                            //        salescnty = js.County;
                            //    }

                            //    if (salescnty.Length < 1)
                            //    {

                            //        if (citycnty.Contains("COUNTY") && !citycnty.Contains(separator_city_cnty))
                            //        {
                            //            salescnty = citycnty;
                            //            salescity = "";
                            //        }
                            //        else
                            //        {
                            //            if (separator_city_cnty.Length > 0)
                            //            {

                            //                int start = 0;
                            //                start = citycnty.IndexOf(separator_city_cnty);
                            //                if (start > 0)
                            //                {
                            //                    int end = start - 1;
                            //                    int len = citycnty.Length - start;
                            //                    start = start + 1;
                            //                    len = len - 1;
                            //                    salescity = citycnty.Substring(0, end);
                            //                    salescnty = Get_County_With_City(salescity);
                            //                }

                            //            }
                            //        }

                            //    }
                           // }
                        }
                    }
           

                   
                }

                if (checkBox_Debug.Checked)
                {
                    MessageBox.Show("City = " + salescity + " County = " + salescnty);
                }
                string zn = row[column_salestax_zone].ToString();
                string stj = Get_Zone_STJ(zn);

                stj = stj.Trim();
                int stjadj = 3 - stj.Length;

                for(int i =0; i<stjadj; i++)
                {
                    stj = "0" + stj;
                }

                //MessageBox.Show("zone="+zn+" stj="+stj);
                int wh_data = 0;

                decimal.TryParse(row[column_sales].ToString(), out s);
                decimal.TryParse(row[column_salestax_amount].ToString(), out stax);
                decimal.TryParse(row[column_set].ToString(), out set);
                decimal.TryParse(row[column_quan].ToString(), out quan);

                prod = row[column_product].ToString();
                inv = row[column_invoice_num].ToString();

                int.TryParse(row[column_pcat].ToString(), out pcat);
                int.TryParse(row[column_warehouse].ToString(), out wh_data);

                if (checkBox_Debug.Checked)
                {
                    MessageBox.Show("State " + indicator_state_ca +" State Data = "+row[column_state].ToString()+ " Pcat= " + pcat.ToString() + " Set=" + set.ToString());
                }

                if(row[column_state].ToString().Contains(tb_State_Indicator.Text))
                {
                    outofcalif = false;
                }
                else
                {
                    outofcalif = true;
                }

                if (checkBox_Debug.Checked)
                {
                    MessageBox.Show("Out of State = " + outofcalif.ToString());
                }

                if (tb_Indicator_State_Not.Text.Length > 0)
                {
                    if (checkBox_Debug.Checked)
                    {
                        MessageBox.Show("State Not =" + tb_Indicator_State_Not.Text);
                    }
                    if (row[column_state].ToString().Contains(tb_Indicator_State_Not.Text))
                    {
                        outofcalif = true;
                    }
                    else
                    {
                        outofcalif = false;
                    }

                    if (checkBox_Debug.Checked)
                    {
                        MessageBox.Show("Out of State = " + outofcalif.ToString());
                    }

                }


                if(row[column_excluded_cardlock].ToString() == indicator_Excluded_Cardlock)
                {
                    cardlockexcluded = true;
                }
                else
                {
                    cardlockexcluded = false;
                }

                //MessageBox.Show("Out of Calif =" + outofcalif.ToString() + " Indicater = " + tb_State_Indicator.Text);

                if (outofcalif)
                {
                    OOSSales = OOSSales + s;
                }

                stj = Get_CA_District_STJ(salescnty, salescity, dtp_Date_End.Value.ToShortDateString());

                if (stj.Length > 0)
                {
                    outofcalif = false;
                }

                if (checkBox_Debug.Checked)
                {
                    MessageBox.Show("After Determination - City = " + salescity + " County = " + salescnty + "  Wholesale=" + row[column_wholesale].ToString() + "  Out of Ca =" + outofcalif.ToString());
                }

                bool shownotmsvl = true;

                if (citycnty.ToUpper().Contains("YUBA") || citycnty.ToUpper().Contains("SUTTER") || citycnty.ToUpper().Contains("INVENTORY"))
                {
                    shownotmsvl = false;
                }


                //if (checkBox_Debug.Checked || shownotmsvl)
                //{
                //    MessageBox.Show("After Determination - " + citycnty + " cITY = " + salescity + " County = " + salescnty + "  Wholesale=" + row[column_wholesale].ToString() + "  Out of Ca =" + outofcalif.ToString());
                //}

                if (!outofcalif)
                {
                    //MessageBox.Show(indicator_wholesale);
                    if (pcat == indicator_pcat_diesel_clear && set!=0 && !row[column_wholesale].ToString().Contains(indicator_wholesale))
                    {
                        //MessageBox.Show("Diesel");
                        if (quan < 0)
                        {
                            if (checkBox_Debug.Checked)
                            {
                                MessageBox.Show(prod + " quan =" + quan.ToString() + " Set =" + set.ToString() + " Sales amount=" + s.ToString());
                            }
                        }

                        salesamt = s;

                        if (set_is_each)
                        {
                            set_tot = (set * quan);
                            salesamt = salesamt - (set * quan);
                            diesel_SET = diesel_SET + set_tot;
                        }
                        else
                        {
                            salesamt = salesamt - set;
                        }

                        if (stax_included_in_sales)
                        {
                            salesamt = salesamt - stax;
                        }

                        diesel_Clear = diesel_Clear + salesamt;
                    }
                    else
                    {
                        salesamt = s;

                        if (stax_included_in_sales)
                        {
                            salesamt = salesamt - stax;
                        }
                    }

                    if (pcat <= indicator_pcat_last_MVF && !row[column_wholesale].ToString().Contains(indicator_wholesale))
                    {
                        mvf_sales = mvf_sales + salesamt;
                    }

                    if (checkBox_Debug.Checked)
                    {
                        MessageBox.Show("s = " + s.ToString() + " set = " + set.ToString() + " stax = " + stax.ToString() + " quan = " + quan.ToString() + " Total = " + salesamt.ToString());
                    }

                    gross_sales = gross_sales + salesamt;

                    if (checkBox_Debug.Checked)
                    {
                        MessageBox.Show("Invoice=" + inv + " Quantity " + quan.ToString() + " Pcat=" + pcat.ToString() + " SET=" + set_tot.ToString() + " Gross Sales=" + gross_sales.ToString() + " Sales Amt=" + salesamt.ToString());
                    }

                                        
                    if (cardlockexcluded)
                    {
                        excluded_cardlock = excluded_cardlock + salesamt;
                    }
                    
                    if (row[column_wholesale].ToString().Contains(indicator_wholesale))
                    {
                        wholesale = wholesale + salesamt;
                    }

                    if(citycnty.Contains("KING "))
                    {
                       // MessageBox.Show(" City ="+salescity + "  County =" + salescnty);
                    }


                    if(column_cnty>0 && column_city > 0 && !row[column_wholesale].ToString().Contains(indicator_wholesale))
                    {
                        Add_To_District_Sales_And_Taxes(salescity, salescnty, salesamt, 0, 0, 0, 0,0,0,dtp_Date_End.Value);
                        stj = "";
                    }

                    if(stj.Length > 0  && !row[column_wholesale].ToString().Contains(indicator_wholesale))
                    {
                        Add_To_District_Sales_And_Taxes_STJ(stj, salesamt, 0, 0, 0, 0);
                        
                    }
                }
             //   MessageBox.Show(gross_sales.ToString());
            }

            ExcelPackage pkg = new ExcelPackage(new MemoryStream());

            var ws0 = pkg.Workbook.Worksheets.Add("District Sales");

            if (cadtax.Rows.Count > 0)
            {
                ws0.Cells["A1"].LoadFromDataTable(cadtax, true);
            }


            var ws1 = pkg.Workbook.Worksheets.Add("Sales Detail");


            if (sales_data.Rows.Count > 0)
            {
                ws1.Cells["A1"].LoadFromDataTable(sales_data, true);
            }

            //sales_data.Clear();

            foreach (DataRow row in cpst_data.Rows)
            {
                int cpst_slszn = 0;

                int.TryParse(row[column_cpst_indicator].ToString(), out cpst_slszn);

                decimal salesamt = 0.00m;


                decimal.TryParse(row[column_cpst_amount].ToString(), out salesamt);

                if(cpst_slszn == indicator_cpst_diesel_clear)
                {
                    cpst_diesel_clear = cpst_diesel_clear + salesamt; 
                }

                if (cpst_slszn == indicator_cpst_diesel_dyed)
                {

                    cpst_diesel_red= cpst_diesel_red + salesamt;
                }

                if (cpst_slszn == indicator_cpst_equip)
                {

                    cpst_farm_equip = cpst_farm_equip + salesamt;
                }

                if (cpst_slszn == indicator_cpst_oils)
                {

                    cpst_farm_equip = cpst_farm_equip + salesamt;
                }


            }

            if(cpst_diesel_clear < 0)
            {
                cpst_diesel_clear = cpst_diesel_clear * -1;
            }

            if (cpst_diesel_red < 0)
            {
                cpst_diesel_red = cpst_diesel_red * -1;
            }

            if (cpst_farm_equip < 0)
            {
                cpst_farm_equip = cpst_farm_equip * -1;
            }

            var ws2 = pkg.Workbook.Worksheets.Add("CPST Detail");

            if (cpst_data.Rows.Count > 0)
            {
                ws2.Cells["A1"].LoadFromDataTable(cpst_data, true);
            }

            //CLear_Datatables();

           tb_Gross_Sales.Text = gross_sales.ToString("###,###,###,###,##");
           tb_Diesel_On_or_After_110117.Text = (diesel_Clear-cpst_diesel_clear).ToString("###,###,###,###.##");
           tb_Diesel_Sales_Retail_Clear.Text = diesel_Clear.ToString("###,###,###,###.##");
           tb_Diesel_Sales_CPST_Clear.Text = cpst_diesel_clear.ToString("###,###,###,###.##");
           tb_Diese_Sales_CPST_Dyed.Text = cpst_diesel_red.ToString("###,###,###,###,##");
           tb_Diesel_SET.Text = diesel_SET.ToString("###,###,###,###.##");
           tb_Deductions_Resale.Text = wholesale.ToString("###,###,###,###.##");
           tb_MVF_Transactions.Text = mvf_sales.ToString("###,###,###,###.##");
           tb_Exemption_Diesel_Clear_Only.Text = cpst_diesel_clear.ToString("###,###,###,###,###.##");
           tb_Exemption_Diesel_Fuel_Farm.Text = (cpst_diesel_clear + cpst_diesel_red).ToString("###,###,###,###.##");
           tb_Exemptions_Farm_Equip.Text = cpst_farm_equip.ToString("###,###,###,###.##");

           tb_Foreign.Text = foreign.ToString("###,###,###,###.##");
           tb_Excluded_Cardlock.Text = excluded_cardlock.ToString("###,###,###,###.##");

            tb_Outof_State_Sales.Text = OOSSales.ToString("###,###,###,###.00");


           var ws3 = pkg.Workbook.Worksheets.Add("Summary");

            ws3.Cells["A1"].Value = "Summary";
            ws3.Cells["A3"].Value = "Gross Sales";
            ws3.Cells["A4"].Value = "Diesel On or After 11-01-2017";
            ws3.Cells["A5"].Value = "Diesel Sales Retail Clear";
            ws3.Cells["A6"].Value = "Diesel CPST Clear";
            ws3.Cells["A7"].Value = "Diesel CPST Dyed";
            ws3.Cells["A8"].Value = "Diesel SET";
            ws3.Cells["A9"].Value = "Deductions Wholesale";
            ws3.Cells["A10"].Value = "MVF Transactions";
            ws3.Cells["A11"].Value = "Exemption Clear Diesel";
            ws3.Cells["A12"].Value = "Exemption Diesel Fuel Farm";
            ws3.Cells["A13"].Value = "Exemption Farm Equip";

            ws3.Cells["A15"].Value = "Foreign Sales";
            ws3.Cells["A16"].Value = "Excluded Cardlock";

            ws3.Cells["B3"].Value = gross_sales;
            ws3.Cells["B4"].Value = diesel_Clear - cpst_diesel_clear;
            ws3.Cells["B5"].Value = diesel_Clear;
            ws3.Cells["B6"].Value = cpst_diesel_clear;
            ws3.Cells["B7"].Value = cpst_diesel_red;
            ws3.Cells["B8"].Value = diesel_SET;
            ws3.Cells["B9"].Value = wholesale;
            ws3.Cells["B10"].Value = mvf_sales;
            ws3.Cells["B11"].Value = cpst_diesel_clear;
            ws3.Cells["B12"].Value = cpst_diesel_red;
            ws3.Cells["B13"].Value = cpst_farm_equip;

            ws3.Cells["B15"].Value = foreign;
            ws3.Cells["B16"].Value = excluded_cardlock;

            pkg.SaveAs(new FileInfo(Sales_File));

            pkg.Dispose();
            
        }



        void Auto_Correct_Row(DataTable dt, int row)
        {
           
            int numcommas = 0;
            int firstdate = 0;
            int movenum = 0;


            for (int c = 0; c < dt.Columns.Count; c++)
            {
                string[] ff = dt.Rows[row][c].ToString().Split(',');
                numcommas = numcommas + ff.Length;

                if (firstdate == 0)
                {
                    if (Is_Date(dt.Rows[row][c].ToString()))
                    {
                        firstdate = c;
                    }
                }

            }

            movenum = datepos1 - firstdate;

            // MessageBox.Show("First Date Pos =" + firstdate.ToString() + " Number of Commas =" + numcommas.ToString()+" Move Columns ="+movenum.ToString() );

            for (int m = 0; m < movenum; m++)
            {
                Move_Row_to_right(dt,row);
            }

            movenum++;
            do
            {
                //MessageBox.Show(dgv_CSV_File.CurrentRow.Cells[movenum].Value.ToString());

                string[] Fieldsinfields = dt.Rows[row][movenum].ToString().Split(',');
                int m = movenum;
                int l = movenum - (Fieldsinfields.Length - 1);
                int i = 0;

                if ((l) < 0)
                {
                    l = 0;
                }

                foreach (string field in Fieldsinfields)
                {
                    //     MessageBox.Show(field);

                    dt.Rows[row][l] = Fieldsinfields[i].ToString();
                    l++;
                    i++;
                }


                movenum = movenum - (Fieldsinfields.Length);

                //   MessageBox.Show(movenum.ToString());

            } while (movenum > 0);


            //for (int c = movenum; c > -1; c--)
            //{
            //    if (dgv_CSV_File.CurrentRow.Cells[c].Value.ToString().Split(',').Length > 0)
            //    {
            //        string[] Fieldsinfields = dgv_CSV_File.CurrentRow.Cells[c].Value.ToString().Split(',');

            //        for (int m = c; m < Fieldsinfields.Length; m--)
            //        {
            //            dgv_CSV_File.CurrentRow.Cells[m].Value = Fieldsinfields[m].ToString();
            //        }

            //        c = c - Fieldsinfields.Length;
            //    }
            //}

            //for (int c=0;c<movenum;c++)
            //{
            //    if(dgv_CSV_File.CurrentRow.Cells[c].Value.ToString().Split(',').Length > 0)
            //    {
            //        string[] Fieldsinfields = dgv_CSV_File.CurrentRow.Cells[c].Value.ToString().Split(',');

            //        for(int m =c; m<Fieldsinfields.Length;m++)
            //        {
            //            dgv_CSV_File.CurrentRow.Cells[m].Value = Fieldsinfields[m].ToString();
            //        }
            //        MessageBox.Show("c=" + c.ToString());
            //        MessageBox.Show(Fieldsinfields.Length.ToString());
            //        c = c + Fieldsinfields.Length;

            //        MessageBox.Show("c=" + c.ToString());
            //        startpos = startpos+Fieldsinfields.Length;
            //    }

            //}

            //MessageBox.Show("Movenum =" + movenum.ToString() + " Startpos=" + startpos.ToString());


        }

        void Move_Row_to_right(DataTable dt, int row)
        {
           
                for (int c = dt.Columns.Count - 1; c > 0; c--)
                {
                    dt.Rows[row][c] = dt.Rows[row][c - 1];
                }
            
        }


        string[] Parse_Line_To_Array(string line)
        {
            // Regexes that don't work
            string regex2 = "^(?: ([^\",]+))?(?=,)|(?<=,)(?:[^\",]*)?(?=[,$])|((?<=\")[^\"(\\s*,)][^\"]*(?=\"))|(?<=\")(?=\")|(?<=,)(?:[^,\"])*(?=$)";
            string sToken = @"(?:,\s+)|(['""].+['""])(?:,\s+)";
            
            string[] comma = line.Split(','); // Simplest of splits won;t work
            string[] combo = line.Split(new string[] { "\",\"","," },StringSplitOptions.None ); // Split that almost works
            string[] linereg = Regex.Split(line, ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)"); // Leaves only Products with additional quote like hoses
            string[] linereg2 = Regex.Split(line, sToken);

            int num_commas = comma.Length;
            int num_combo = combo.Length;
            int num_linereg = linereg.Length;
            int num_linereg2 = linereg2.Length;           

           // MessageBox.Show("Comma Fields = " + num_commas.ToString() + " Combo Fields = " + num_combo.ToString() + " Reg Fields = " + num_linereg.ToString()+" Reg2 Fields = "+num_linereg2.ToString());

            //if (num_linereg < 67)
            //{
            //    MessageBox.Show("Line = " + line);
            //}

            //foreach(string field in comma)
            //{
            //    MessageBox.Show("Comma Field =" + field);
            //}

            //foreach (string field in combo)
            //{
            //    MessageBox.Show("Combo Field =" + field);
            //}

            //foreach (string field in linereg)
            //{
            //    MessageBox.Show("Linereg Field =" + field);
            //}
            return linereg; // close to clean

        }


        string[] Get_CSV_Data(string file)
        {
            string del = Analyze_File(file);

            string[] lines = File.ReadAllLines(file);

            return lines;
        }


        List<string[]> Load_Data_To_List(string[] lines, string del)
        {

            List<string[]> file = new List<string[]>();

            foreach(string line in lines)
            {

                file.Add(Get_Fields_Delimited(line, del));


            }

            return file;

        }



        DataTable Load_Data_To_Datatable(string[] lines)
        {
            DataTable dt = new DataTable();
            
            int numfields = Max_Fields(lines);
                      
            if (numfields > 0)
            {
                for (int i = 0; i < numfields; i++)
                {
                    int j = i + 1;
                    dt.Columns.Add("Column " + j.ToString(), typeof(String));
                }
            }

            int r = 0;

            string del = Analyze_Line(lines);

           // MessageBox.Show("Load_data "+ del);

            foreach (string line in lines)
            {
                DataRow row = dt.NewRow();
                
                string[] fields = Parse_CSV_Line(line, "Commas", r);

                //string[] fields = Parse_Line_To_Array(line);

                //if (del.Length > 0)
                //{
                //    fields = Get_Fields_Delimited(line, del);
                //}
              
                                
                //string[] fields = Regex.Split(line, ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");

                // string[] fields = Regex.Split(line, ",(?<=^|,)(\"(?:[^\"]|\"\")*\"|[^,]*)");

                r++;

                if (fields.Length <= numfields)
                {
                    int f = 0;

                    foreach (string field in fields)
                    {

                        row[f] = field.Replace("\"", "");
                        f++;
                    }


                }
                else
                {
                    MessageBox.Show("Record not in correct format - " + numfields.ToString() + " Line Field Count " + fields.Length.ToString() + " " + line);

                }

                dt.Rows.Add(row);

            }

            Check_Data(dt);

            return dt;
        }

        List<int> Analyse_DataTable_Return_Bad_Rows(DataTable dt)
        {
            List<int> BRows = new List<int>();

            int date1pos = 0;
            int date2pos = 0;
            int date3pos = 0;

            DataRow row = dt.Rows[0];

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                if (row[i].ToString().Contains("date") || row[i].ToString().Contains("Date") || row[i].ToString().Contains("DATE"))
                {
                    if (date1pos == 0)
                    {
                        date1pos = i;
                    }
                    else
                    {
                        if (date2pos == 0)
                        {
                            date2pos = i;
                        }
                        else
                        {
                            if (date3pos == 0)
                            {
                                date3pos = i;
                            }

                        }

                    }
                }
            }

            if (date1pos == 0)
            {
                row = dt.Rows[1];

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (Is_Date(row[i].ToString()))
                    {
                        if (date1pos == 0)
                        {
                            date1pos = i;
                        }
                        else
                        {
                            if (date2pos == 0)
                            {
                                date2pos = i;
                            }
                            else
                            {
                                if (date3pos == 0)
                                {
                                    date3pos = i;
                                }

                            }

                        }
                    }
                }

            }

            // MessageBox.Show("Date 1 Position = " + date1pos.ToString() + " Date 2 Position =" + date2pos.ToString() + " Date 3 Position=" + date3pos.ToString());

            for (int i = 1; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];

                if (!Is_Date(dr[date1pos].ToString()))
                {
                    //MessageBox.Show("Data at row " + i.ToString() + " Is Not Good.");
                    for (int r = 0; r < dt.Columns.Count; r++)
                    {
                        //  MessageBox.Show("Field " + r.ToString() + " =" + dr[r].ToString());
                    }

                    BRows.Add(i);
                }


            }

            datepos1 = date1pos;
            datepos2 = date2pos;
            datepos3 = date3pos;

            return BRows;
        }

        bool Is_Date(string date)
        {


            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }


        }

        void Check_Data(DataTable dt)
        {
            List<int> BadRows = Analyse_DataTable_Return_Bad_Rows(dt);

            CSV_Bad_Rows = BadRows;

            string browsstr = "";

            foreach (int bad in BadRows)
            {
                if (browsstr.Length > 0)
                {
                    browsstr = browsstr + "," + bad.ToString();
                }
                else
                {
                    browsstr = bad.ToString();
                }
            }

           
        }

        string[] Parse_CSV_Line(string line, string code, int numrecs)
        {
            int len = line.Length;
            int aps = line.Split('"').Length - 1;
            int cms = line.Split(',').Length - 1;

            string[] linecms = line.Split(',');

            
            string[] linereg = Regex.Split(line, ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");

            
            int aseq = 0;

            for (int i = 0; i < len - 2; i++)
            {
                if (line.Substring(i, 1).Equals("\""))
                {
                    if (line.Substring(i + 1, 1).Equals(","))
                    {
                        if (line.Substring(i + 2, 1).Equals("\""))
                        {
                            aseq++;
                        }
                    }
                }
            }
            return linereg;
        }

        string[] Get_Fields(string line)
        {
            string[] fields = Regex.Split(line, ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
            return fields;
        }

        int Max_Fields(string[] lines)
        {
            int max = 0;
            foreach (string line in lines)
            {
                string[] fields = Regex.Split(line, ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");

                // MessageBox.Show(fields.Length.ToString());
                if (max < fields.Length)
                {

                    max = fields.Length;


                }
            }

            return max;
        }

        private void button_Calc_Data_Click(object sender, EventArgs e)
        {
            Calc_Data();
        }

        private void button_Get_Cardlock_File_Click(object sender, EventArgs e)
        {
            cardlockfile = Get_File("Cardlock File", "cltrans*");
           
        }

        private void button_Get_Purchases_File_Click(object sender, EventArgs e)
        {
            purchasesfile = Get_File("Purchases File", "Invtrans*");
           
        }

        
        string[] Read_XLSX(string file)
        {
            FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read);
            FileInfo fi = new FileInfo(file);

            using (ExcelPackage pkg = new ExcelPackage(fi))
            {
               // MessageBox.Show(pkg.Workbook.Worksheets.Count.ToString() + " " + pkg.Workbook.Properties.Status.ToString());

                ExcelWorkbook wb = pkg.Workbook;
                ExcelWorksheet ws = wb.Worksheets.First();

                string[] linesconv = new string[ws.Dimension.End.Row];

                //MessageBox.Show(linesconv.Length.ToString());
                for (int r = 1; r <= ws.Dimension.Rows; r++)
                {
                    for (int c = 1; c <= ws.Dimension.Columns; c++)
                    {
                        if (ws.Cells[r, c].Value != null)
                        {
                            linesconv[r-1] = linesconv[r-1] + ws.Cells[r, c].Value.ToString();
                        }
                        else
                        {
                            linesconv[r-1] = linesconv[r-1] + " ";

                        }
                        if (c < (ws.Dimension.Columns -1))
                        {
                            linesconv[r-1] = linesconv[r-1] + ",";
                        }
                       
                    }
                   // linesconv[r-1] = linesconv[r-1] + Environment.NewLine;

                   // MessageBox.Show(linesconv[r-1]);
                }
                return linesconv;
            }
        }

        void Load_single_file()
        {
            OpenFileDialog cnvfile = new OpenFileDialog();
            if (cnvfile.ShowDialog() == DialogResult.OK)
            {
                string[] result = Read_XLSX(cnvfile.FileName);

                string newfile = cnvfile.FileName.Replace(".xlsx", ".csv");
                if (File.Exists(newfile))
                {
                    File.Delete(newfile);
                }

                File.WriteAllLines(newfile, result);
                
            }
        }

        void Convert_Dir()
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();

            if(folder.ShowDialog() == DialogResult.OK)
            {
                string[] files = Directory.GetFiles(folder.SelectedPath,"*.xlsx");
                foreach(string file in files)
                {
                    string[] result = Read_XLSX(file);

                    string newfile = file.Replace(".xlsx", ".csv");
                    if (File.Exists(newfile))
                    {
                        File.Delete(newfile);
                    }

                    File.WriteAllLines(newfile, result);
                }               


            }
            
        }

        void Load_CFN_CSV_File()
        {

            OpenFileDialog CFNfile = new OpenFileDialog();

            CFNfile.Multiselect = true;

            if (CFNfile.ShowDialog() == DialogResult.OK)
            {
                rtb_CFN_Files.Text = "";

                int cfnrecords = 0;

                CFNData.Clear();

                foreach (string file in CFNfile.FileNames)
                {
                    try
                    {
                        string[] lines;

                        //if (file.Contains(".xlsx"))
                        //{


                        //  //  var pkg = new ExcelPackage(new FileInfo(file));

                        //    //MessageBox.Show(pkg.Workbook.Worksheets.Count.ToString());

                        //    //ExcelWorksheet ws = pkg.Workbook.Worksheets[1];



                        //    //string[] linesconv = new string[ws.Dimension.Rows];

                        //    //for (int r = 0; r < ws.Dimension.Rows; r++)
                        //    //{
                        //    //    for (int c = 0; c < ws.Dimension.Columns; c++)
                        //     //   {
                        //       //     linesconv[r] = linesconv[r] + ws.Cells[r, c].Value.ToString();
                        //        //    if (c < (ws.Dimension.Columns - 1))
                        //         //   {
                        //                linesconv[r] = linesconv[r] + ",";
                        //          //  }
                        //          //  linesconv[r] = linesconv[r] + Environment.NewLine;
                        //       // }
                        //   // }



                        //   // lines = linesconv;
                        //}
                        //else
                        //{
                        //    lines = File.ReadAllLines(file);
                        //}

                        lines = File.ReadAllLines(file);

                        rtb_CFN_Files.AppendText(file + Environment.NewLine);

                        notifyIcon1.Icon = SystemIcons.Hand;
                        notifyIcon1.BalloonTipText = "Reading File " + file;
                        notifyIcon1.BalloonTipTitle = "Read CFN File";
                        notifyIcon1.ShowBalloonTip(500);

                        foreach (string line in lines)
                        {
                            string[] fields = line.Split(',');
                            PTFile ptfile = new PTFile();

                            ptfile.Site = fields[0];
                            ptfile.Sequence = fields[1];
                            ptfile.Status = fields[2];
                            ptfile.Total = fields[3];
                            ptfile.Account = fields[4];
                            ptfile.Product = fields[5];
                            ptfile.Type = fields[6];
                            ptfile.ProdDesc = fields[7];
                            ptfile.Price = fields[8];
                            ptfile.Quantity = fields[9];
                            ptfile.Odometer = fields[10];
                            ptfile.Pump = fields[11];
                            ptfile.Trans = fields[12];
                            ptfile.Date = fields[13];
                            ptfile.Time = fields[14];
                            ptfile.Error = fields[15];
                            ptfile.Authorization = fields[16];
                            ptfile.ManualEntry = fields[17];
                            ptfile.Card = fields[18];
                            ptfile.Vehicle = fields[19];
                            ptfile.SiteTaxLocation = fields[20];
                            ptfile.Code0 = fields[21];
                            ptfile.Code1 = fields[22];
                            ptfile.Code2 = fields[23];
                            ptfile.Code3 = fields[24];
                            ptfile.Code4 = fields[25];
                            ptfile.Code5 = fields[26];
                            ptfile.Code6 = fields[27];
                            ptfile.Code7 = fields[28];
                            ptfile.Code8 = fields[29];
                            ptfile.Code9 = fields[30];
                            ptfile.Amount0 = fields[31];
                            ptfile.Amount1 = fields[32];
                            ptfile.Amount2 = fields[33];
                            ptfile.Amount3 = fields[34];
                            ptfile.Amount4 = fields[35];
                            ptfile.Amount5 = fields[36];
                            ptfile.Amount6 = fields[37];
                            ptfile.Amount7 = fields[38];
                            ptfile.Amount8 = fields[39];
                            ptfile.Amount9 = fields[40];
                            ptfile.NetType = fields[41];
                            ptfile.CF = fields[42];
                            ptfile.NetRate = fields[43];
                            ptfile.PumpPrice = fields[44];
                            ptfile.HaulRate = fields[45];
                            ptfile.CFNPrice = fields[46];
                            ptfile.JobNumber = fields[47];
                            ptfile.PONumber = fields[48];


                          
                            int dte = 0;
                            int.TryParse(ptfile.Date, out dte);
                            int cur = 0;
                            int.TryParse(dtp_Date_Beg.Value.ToString("yyMMdd"), out cur);
                            int end = 0;
                            int.TryParse(dtp_Date_End.Value.ToString("yyMMdd"), out end);

                            //MessageBox.Show("dte =" + dte.ToString() + " cur=" + cur.ToString() + " end=" + end.ToString());

                            if (dte >= cur && dte <= end)
                            {
                                CFNData.Add(ptfile);
                                cfnrecords++;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("File cannot be opened. Close file and try again");
                    }

                }

                foreach (PTFile row in CFNData)
                {
                    string chkcity = Get_CFN_City(row.Site);
                    if (chkcity.Length > 1)
                    {
                        row.City = chkcity;
                    }
                    else
                    {
                        row.City = "Out of California";
                    }

                    row.County = Get_County_With_City(chkcity);
                    //row["County"] = Get_CFN_County(row[0].ToString());
                    row.State = Get_CFN_State(row.Site);

                    int tsite = 0;
                    int.TryParse(row.Site, out tsite);

                    if (tsite > 9999 && row.State.Contains("CA"))
                    {
                        PTFile ext = new PTFile();

                        ext = row;
                                        
                        switch (row.Code0)
                        {
                            case "0501":
                                ext.StateTax = row.Amount0;
                                row.StateTax = row.Amount0;
                                break;
                            case "0504":
                                ext.CountyTax = row.Amount0;
                                row.CountyTax = row.Amount0;
                                break;
                            case "0505":
                                ext.CityTax = row.Amount0;
                                row.CityTax = row.Amount0;
                                break;
                            case "0203":
                                ext.SET = row.Amount0;
                                row.SET = row.Amount0;
                                break;
                            case "0507":
                                ext.StateTax = row.Amount0;
                                row.StateTax = row.Amount0;
                                break;

                        }

                        switch (row.Code1)
                        {
                            case "0501":
                                ext.StateTax = row.Amount1;
                                row.StateTax = row.Amount1;
                                break;
                            case "0504":
                                ext.CountyTax = row.Amount1;
                                row.CountyTax = row.Amount1;
                                break;
                            case "0505":
                                ext.CityTax = row.Amount1;
                                row.CityTax = row.Amount1;
                                break;
                            case "0203":
                                ext.SET = row.Amount1;
                                row.SET = row.Amount1;
                                break;
                            case "0507":
                                ext.StateTax = row.Amount1;
                                row.StateTax = row.Amount1;
                                break;

                        }

                        switch (row.Code2)
                        {
                            case "0501":
                                ext.StateTax = row.Amount2;
                                row.StateTax = row.Amount2;
                                break;
                            case "0504":
                                ext.CountyTax = row.Amount2;
                                row.CountyTax = row.Amount2;
                                break;
                            case "0505":
                                ext.CityTax = row.Amount2;
                                row.CityTax = row.Amount2;
                                break;
                            case "0203":
                                ext.SET = row.Amount2;
                                row.SET = row.Amount2;
                                break;
                            case "0507":
                                ext.StateTax = row.Amount2;
                                row.StateTax = row.Amount2;
                                break;

                        }

                        switch (row.Code3)
                        {
                            case "0501":
                                ext.StateTax = row.Amount3;
                                row.StateTax = row.Amount3;
                                break;
                            case "0504":
                                ext.CountyTax = row.Amount3;
                                row.CountyTax = row.Amount3;
                                break;
                            case "0505":
                                ext.CityTax = row.Amount3;
                                row.CityTax = row.Amount3;
                                break;
                            case "0203":
                                ext.SET = row.Amount3;
                                row.SET = row.Amount3;
                                break;
                            case "0507":
                                ext.StateTax = row.Amount3;
                                row.StateTax = row.Amount3;
                                break;

                        }

                        switch (row.Code4)
                        {
                            case "0501":
                                ext.StateTax = row.Amount4;
                                row.StateTax = row.Amount4;
                                break;
                            case "0504":
                                ext.CountyTax = row.Amount4;
                                row.CountyTax = row.Amount4;
                                break;
                            case "0505":
                                ext.CityTax = row.Amount4;
                                row.CityTax = row.Amount4;
                                break;
                            case "0203":
                                ext.SET = row.Amount4;
                                row.SET = row.Amount4;
                                break;
                            case "0507":
                                ext.StateTax = row.Amount4;
                                row.StateTax = row.Amount4;
                                break;

                        }

                        switch (row.Code5)
                        {
                            case "0501":
                                ext.StateTax = row.Amount5;
                                row.StateTax = row.Amount5;
                                break;
                            case "0504":
                                ext.CountyTax = row.Amount5;
                                row.CountyTax = row.Amount5;
                                break;
                            case "0505":
                                ext.CityTax = row.Amount5;
                                row.CityTax = row.Amount5;
                                break;
                            case "0203":
                                ext.SET = row.Amount5;
                                row.SET = row.Amount5;
                                break;
                            case "0507":
                                ext.StateTax = row.Amount5;
                                row.StateTax = row.Amount5;
                                break;

                        }

                        switch (row.Code6)
                        {
                            case "0501":
                                ext.StateTax = row.Amount6;
                                row.StateTax = row.Amount6;
                                break;
                            case "0504":
                                ext.CountyTax = row.Amount6;
                                row.CountyTax = row.Amount6;
                                break;
                            case "0505":
                                ext.CityTax = row.Amount6;
                                row.CityTax = row.Amount6;
                                break;
                            case "0203":
                                ext.SET = row.Amount6;
                                row.SET = row.Amount6;
                                break;
                            case "0507":
                                ext.StateTax = row.Amount6;
                                row.StateTax = row.Amount6;
                                break;

                        }

                        switch (row.Code7)
                        {
                            case "0501":
                                ext.StateTax = row.Amount7;
                                row.StateTax = row.Amount7;
                                break;
                            case "0504":
                                ext.CountyTax = row.Amount7;
                                row.CountyTax = row.Amount7;
                                break;
                            case "0505":
                                ext.CityTax = row.Amount7;
                                row.CityTax = row.Amount7;
                                break;
                            case "0203":
                                ext.SET = row.Amount7;
                                row.SET = row.Amount7;
                                break;
                            case "0507":
                                ext.StateTax = row.Amount7;
                                row.StateTax = row.Amount7;
                                break;

                        }

                        switch (row.Code8)
                        {
                            case "0501":
                                ext.StateTax = row.Amount8;
                                row.StateTax = row.Amount8;
                                break;
                            case "0504":
                                ext.CountyTax = row.Amount8;
                                row.CountyTax = row.Amount8;
                                break;
                            case "0505":
                                ext.CityTax = row.Amount8;
                                row.CityTax = row.Amount8;
                                break;
                            case "0203":
                                ext.SET = row.Amount8;
                                row.SET = row.Amount8;
                                break;
                            case "0507":
                                ext.StateTax = row.Amount8;
                                row.StateTax = row.Amount8;
                                break;

                        }

                        switch (row.Code9)
                        {
                            case "0501":
                                ext.StateTax = row.Amount9;
                                row.StateTax = row.Amount9;
                                break;
                            case "0504":
                                ext.CountyTax = row.Amount9;
                                row.CountyTax = row.Amount9;
                                break;
                            case "0505":
                                ext.CityTax = row.Amount9;
                                row.CityTax = row.Amount9;
                                break;
                            case "0203":
                                ext.SET = row.Amount9;
                                row.SET = row.Amount9;
                                break;
                            case "0507":
                                ext.StateTax = row.Amount9;
                                row.StateTax = row.Amount9;
                                break;

                        }

                        decimal stax = 0.00m;

                        decimal cntytax = 0.00m;
                        decimal ctytax = 0.00m;
                        decimal taxable = 0.00m;

                        decimal.TryParse(ext.StateTax, out stax);
                        decimal.TryParse(ext.CountyTax, out cntytax);
                        decimal.TryParse(ext.CityTax, out ctytax);


                        switch (row.Product)
                        {
                            case "03":
                                taxable = (stax) / (decimal)stax_rate_diesel;
                                ext.Taxable = taxable.ToString("###,###,###.##");
                                row.Taxable = ext.Taxable;
                                break;
                            case "53":
                                taxable = (stax) / (decimal)stax_rate_diesel;
                                ext.Taxable = taxable.ToString("###,###,###.##");
                                row.Taxable = ext.Taxable;
                                break;
                            case "44":
                                taxable = (stax) / (decimal)stax_rate_gasoline;
                                ext.Taxable = taxable.ToString("###,###,###.##");
                                row.Taxable = ext.Taxable;
                                break;
                            case "45":
                                taxable = (stax) / (decimal)stax_rate_gasoline;
                                ext.Taxable = taxable.ToString("###,###,###.##");
                                row.Taxable = ext.Taxable;
                                break;
                            case "46":
                                taxable = (stax) / (decimal)stax_rate_gasoline;
                                ext.Taxable = taxable.ToString("###,###,###.##");
                                row.Taxable = ext.Taxable;
                                break;
                        }

                        CFNDataExt.Add(ext);
                    }

                }

                label_Number_Of_Trans_CFN.Text = "Records =" + CFNData.Count.ToString("###,###,###,###.##");

                dgv_CFN_Data.DataSource = CFNData;
                dgv_CFN_Extended.DataSource = CFNDataExt;

                label_CFN_Records.Text = CFNData.Count.ToString("###,###,###");
                label_CFN_Extended_Records.Text = CFNDataExt.Count.ToString("###,###,###");
            }
        }
               

        string Get_CFN_County(string sitecode)
        {
            string county = "";
            int site = 0;
            int.TryParse(sitecode, out site);

            string ctycode = "";

            foreach (Sites_CFN row in CFNSites)
            {

                int lusite = 0;

                int.TryParse(row.Site, out lusite);

                if (lusite == site)
                {
                    ctycode = row.CityCode.ToString();
                }


            }
            int cnty = 0;

            int.TryParse(ctycode, out cnty);

            foreach(Counties_CFN row in CFNCounties)
            {
                int cntycode = 0;

                int.TryParse(row.Code.ToString(), out cntycode);

                if(cnty == cntycode)
                {
                    county = row.County;
                }
            }


            return county;
        }


        string Get_CFN_State(string sitecode)
        {
            string state = "";

            int site = 0;
            int.TryParse(sitecode, out site);

            foreach (Sites_CFN row in CFNSites)
            {

                int lusite = 0;

                int.TryParse(row.Site, out lusite);

                if (lusite == site)
                {
                    state = row.State;
                }


            }

            return state;
        }

        string Get_CFN_City(string sitecode)
        {
            string city = "";

            int site = 0;
            int.TryParse(sitecode, out site);

            foreach(Sites_CFN row in CFNSites)
            {
                
                int lusite = 0;

                int.TryParse(row.Site.ToString(), out lusite);

                if(lusite == site)
                {
                    city = row.CityName;
                }
                
                
            }

            return city;
        }

        decimal Get_CA_SET_MVF_Rate(string date)
        {
            decimal rate = 0.00m;

            DateTime s_date = DateTime.Now;

            DateTime.TryParse(date, out s_date);


            foreach (DataRow row in CATaxRates_MVF.Rows)
            {

                DateTime start_date = new DateTime();
                DateTime end_date = new DateTime();

                DateTime.TryParse(row["Start"].ToString(), out start_date);
                DateTime.TryParse(row["End"].ToString(), out end_date);

                if (s_date >= start_date && s_date <= end_date)
                {
                    string set = row["SET"].ToString().Replace("$", "");

                    decimal.TryParse(set, out rate);
                }

            }

            return rate;

        }

        decimal Get_CA_TaxRate_MVF(string date)
        {
            decimal rate = 0.00m;

            DateTime s_date = DateTime.Now;

            DateTime.TryParse(date, out s_date);

            //MessageBox.Show(date + " "+s_date.ToShortDateString());


            foreach (DataRow row in CATaxRates_MVF.Rows)
            {

                DateTime start_date = new DateTime();
                DateTime end_date = new DateTime();

                DateTime.TryParse(row["Start"].ToString(), out start_date);
                DateTime.TryParse(row["End"].ToString(), out end_date);

                //MessageBox.Show("Start Date =" + start_date.ToShortDateString() + " " + "End Date =" + end_date.ToShortDateString());

                if (s_date >= start_date && s_date <= end_date)
                {
                    //MessageBox.Show(row["SET"].ToString());
                    string set = row["SalesTaxRate"].ToString().Replace("%", "");

                    decimal.TryParse(set, out rate);
                }

            }

            rate = rate * .01m;

            return rate;
        }


        decimal Get_CA_TaxRate_DSL(string date)
        {
            decimal rate = 0.00m;

            DateTime s_date = DateTime.Now;

            DateTime.TryParse(date, out s_date);

            //MessageBox.Show(date + " "+s_date.ToShortDateString());


            foreach (DataRow row in CATaxRates_DSL.Rows)
            {

                DateTime start_date = new DateTime();
                DateTime end_date = new DateTime();

                DateTime.TryParse(row["Start"].ToString(), out start_date);
                DateTime.TryParse(row["End"].ToString(), out end_date);

                //MessageBox.Show("Start Date =" + start_date.ToShortDateString() + " " + "End Date =" + end_date.ToShortDateString());

                if (s_date >= start_date && s_date <= end_date)
                {
                   // MessageBox.Show(row["SalesTaxRate"].ToString());
                    string set = row["SalesTaxRate"].ToString().Replace("%", "");


                    decimal.TryParse(set, out rate);
                }

            }

            rate = rate * .01m;

            return rate;
        }

        string Get_CA_District_STJ(string county, string city, string date)
        {
            string STJ = "";

            DateTime s_date = DateTime.Now;

            DateTime.TryParse(date, out s_date);

            foreach (DataRow row in cadtax.Rows)
            {

                DateTime start_date = new DateTime();
                DateTime end_date = new DateTime();

                if (row["Effective"].ToString().Contains("/"))
                {
                    DateTime.TryParse(row["Effective"].ToString(), out start_date);
                }
                else
                {
                    start_date = DateTime.Now.AddYears(-3).AddMonths(-1);
                }

                if (row["Expired"].ToString().Contains("/"))
                {
                    DateTime.TryParse(row["Expired"].ToString(), out end_date);
                }
                else
                {
                    end_date = DateTime.Now;
                }


                county = county.ToUpper().Replace(" ", "");
                county = county.Replace("COUNTY", "");
                city = city.ToUpper().Replace(" ", "");

                string r_county = row["County"].ToString().ToUpper().Replace(" ", "");
                r_county = r_county.Replace("COUNTY", "");
                string r_city = row["City"].ToString().ToUpper().Replace(" ", "").Replace("$", "");
                // MessageBox.Show("County = " + county + " R County = " + r_county + " City = " + city + " R City =" + r_city + " Start = " + start_date.ToShortDateString() + " End = " + end_date.ToShortDateString() + "Passes Date = " + s_date.ToShortDateString());

                if (county.Equals(r_county) && STJ.Length <3)
                {
                    //MessageBox.Show("County = " + county + " R County = " + r_county + " City = " + city + " R City =" + r_city + " Start = " + start_date.ToShortDateString() + " End = " + end_date.ToShortDateString() + "Passes Date = " + s_date.ToShortDateString());

                    if (s_date >= start_date && s_date <= end_date)
                    {
                        STJ = row["Stj"].ToString();
                    }
                }



                if (county.Equals(r_county) && city.Equals(r_city))
                {
                    //MessageBox.Show("County = " + county + " R County = " + r_county + " City = " + city + " R City =" + r_city + " Start = " + start_date.ToShortDateString() + " End = " + end_date.ToShortDateString() + "Passes Date = " + s_date.ToShortDateString());

                    if (s_date >= start_date && s_date <= end_date)
                    {
                        STJ =row["Stj"].ToString();
                    }
                }

            }


            return STJ;

        }



        decimal Get_CA_District_Rate(string county, string city, string date)
        {
            decimal rate = 0.00m;

            DateTime s_date = DateTime.Now;

            DateTime.TryParse(date, out s_date);

            foreach (DataRow row in cadtax.Rows)
            {

                DateTime start_date = new DateTime();
                DateTime end_date = new DateTime();

                if (row["Effective"].ToString().Contains("/"))
                {
                    DateTime.TryParse(row["Effective"].ToString(), out start_date);
                }
                else
                {
                    start_date = DateTime.Now.AddYears(-3).AddMonths(-1);
                }

                if (row["Expired"].ToString().Contains("/"))
                {
                    DateTime.TryParse(row["Expired"].ToString(), out end_date);
                }
                else
                {
                    end_date = DateTime.Now;
                }

                
                county = county.ToUpper().Replace(" ", "");
                county = county.Replace("COUNTY", "");
                city = city.ToUpper().Replace(" ", "");

                string r_county = row["County"].ToString().ToUpper().Replace(" ", "");
                r_county = r_county.Replace("COUNTY", "");
                string r_city = row["City"].ToString().ToUpper().Replace(" ", "").Replace("$","");
                // MessageBox.Show("County = " + county + " R County = " + r_county + " City = " + city + " R City =" + r_city + " Start = " + start_date.ToShortDateString() + " End = " + end_date.ToShortDateString() + "Passes Date = " + s_date.ToShortDateString());

                if (county.Equals(r_county) && rate == 0.00m)
                {
                    if (s_date >= start_date && s_date <= end_date)
                    {
                        decimal.TryParse(row["Rate"].ToString(), out rate);
                    }
                }

                if (county.Equals(r_county) && city.Equals(r_city))
                {
                    //MessageBox.Show("County = " + county + " R County = " + r_county + " City = " + city + " R City =" + r_city + " Start = " + start_date.ToShortDateString() + " End = " + end_date.ToShortDateString() + "Passes Date = " + s_date.ToShortDateString());

                    if (s_date >= start_date && s_date <= end_date)
                    {
                        decimal.TryParse(row["Rate"].ToString(), out rate);
                    }
                }

            }


                return rate;

        }



        decimal Get_CA_SET_DSL_Rate(string date)
        {
            decimal rate = 0.00m;

            DateTime s_date = DateTime.Now;

            DateTime.TryParse(date, out s_date);

            //MessageBox.Show(date + " "+s_date.ToShortDateString());


            foreach(DataRow row in CATaxRates_DSL.Rows)
            {

                DateTime start_date = new DateTime();
                DateTime end_date = new DateTime();

                DateTime.TryParse(row["Start"].ToString(), out start_date);
                DateTime.TryParse(row["End"].ToString(), out end_date);

                //MessageBox.Show("Start Date =" + start_date.ToShortDateString() + " " + "End Date =" + end_date.ToShortDateString());

                if(s_date>=start_date && s_date <= end_date)
                {
                    //MessageBox.Show(row["SET"].ToString());
                    string set = row["SET"].ToString().Replace("$", "");

                    decimal.TryParse(set, out rate);
                }
              
            }
            
            return rate;
                       
        }

        int Get_FTA_Voyager(string prod)
        {
            int FTA = 0;
                                

            foreach(VoyagerCodes v in VoyCodes)
            {
              
                string code = v.Code;

                // MessageBox.Show("Prod =" + prod + " Code =" + code);

                if (code.Equals(prod))
                {
                    int.TryParse(v.FTACode, out FTA);
                }
            }


            //MessageBox.Show("Prod =" + prod + " FTA =" + FTA.ToString());

            return FTA;
        }


        void Load_Voyager_File()
        {
            OpenFileDialog Voyfiles = new OpenFileDialog();

            Voyfiles.Multiselect = true;


            if (Voyfiles.ShowDialog() == DialogResult.OK)
            {

                rtb_Voyager_Files.Text = "";


                foreach (string file in Voyfiles.FileNames)
                {

                    string[] lines = File.ReadAllLines(file);

                    rtb_Voyager_Files.AppendText(file + Environment.NewLine);

                    notifyIcon1.Icon = SystemIcons.Hand;
                    notifyIcon1.BalloonTipText = "Reading File " + file;
                    notifyIcon1.BalloonTipTitle = "Read Voyager";
                    notifyIcon1.ShowBalloonTip(500);
                    string v_date = "";

                    foreach (string line in lines)
                    {

                        VoyagerData vdata = new VoyagerData();

                        DateTime date = DateTime.Now;
                       
                        if (line.Substring(0, 2).Equals("10"))
                        {
                            int datestart = line.IndexOf("2");
                            v_date = line.Substring(datestart, 8);
                            v_date = v_date.Substring(4, 2) + "/" + v_date.Substring(6, 2) + "/" + v_date.Substring(0, 4);
                            //MessageBox.Show(v_date);

                        }


                            if (!line.Substring(0, 2).Equals("10") && !line.Substring(0,2).Equals("99"))
                        {
                            int firstplus = line.IndexOf("+");

                           
                            vdata.SiteAddress = line.Substring(firstplus-55,24);
                            vdata.SiteCity = line.Substring(firstplus-30, 16);
                            vdata.State = line.Substring(firstplus-13, 2);
                            vdata.Zip = line.Substring(firstplus-11, 5);
                            vdata.Product = line.Substring(firstplus-2, 2);
                            vdata.Amount = line.Substring(firstplus+1, 7);
                            vdata.Quantity = line.Substring(firstplus+9, 7);
                            vdata.Price = line.Substring(firstplus+17, 7);
                           
                            vdata.County = Get_County_With_City(vdata.SiteCity);
                            vdata.STJ = Get_CA_District_STJ(vdata.County,vdata.SiteCity,v_date);

                            decimal d_rate = 0.00m;
                            decimal s_rate = 0.00m;

                            d_rate = Get_CA_District_Rate(vdata.County,vdata.SiteCity,v_date);
                            
                            if(d_rate == 0.00m)
                            {
                                d_rate = Get_Jurisdiction_Rate(vdata.SiteCity, vdata.County, dtp_Date_End.Value);
                            }

                            vdata.Date = v_date;
                                                       

                            decimal v_quan = 0.00m;
                            decimal v_tax = 0.00m;
                            decimal v_fee = 0.00m;
                            decimal v_price = 0.00m;
                            decimal v_total = 0.00m;
                            decimal v_base = 0.00m;
                            decimal v_SET = 0.00m;
                            decimal v_taxbase = 0.00m;
                           


                            int v_fta = Get_FTA_Voyager(vdata.Product); ;

                            decimal.TryParse(vdata.Quantity, out v_quan);
                            decimal.TryParse(vdata.Price, out v_price);
                            decimal.TryParse(vdata.Amount, out v_total);

                            v_quan = v_quan * .01m;
                            v_price = v_price * .01m;
                            v_total = v_total * .01m;

                            v_base = v_quan * v_price;
                            v_fee = v_total - v_base;


                            if (v_fta == 167)
                            {
                                v_SET = Get_CA_SET_DSL_Rate(v_date);
                                v_taxbase = v_base - (v_SET * v_quan);
                                s_rate = Get_CA_TaxRate_DSL(v_date);

                            }

                            if (v_fta == 65)
                            {
                                v_SET = Get_CA_SET_MVF_Rate(v_date);
                                s_rate = Get_CA_TaxRate_MVF(v_date);
                                v_taxbase = v_base;
                            }
                            
                            vdata.RateDistrict= d_rate.ToString();
                            vdata.RateState = s_rate.ToString();

                            decimal t_rate = d_rate + s_rate + 1;
                            
                            v_taxbase = v_taxbase / t_rate;
                            v_tax = v_taxbase * (d_rate + s_rate);
                            decimal s_tax = v_taxbase * s_rate;
                            decimal d_tax = v_taxbase * d_rate;

                            vdata.SET = v_SET.ToString();
                            vdata.FTA = v_fta.ToString();
                            vdata.Fee = v_fee.ToString();
                            vdata.Base = v_base.ToString();
                            vdata.Quantity = v_quan.ToString();
                            vdata.Price = v_price.ToString();
                            vdata.Amount = v_total.ToString();
                            vdata.TaxBase = v_taxbase.ToString();
                            vdata.Tax = v_tax.ToString();
                            vdata.TaxDistrict = d_tax.ToString();
                            vdata.TaxState = s_tax.ToString();


                            VoyData.Add(vdata);
                        }

                        notifyIcon1.Icon = SystemIcons.Hand;
                        notifyIcon1.BalloonTipText = "Loading Rates";
                        notifyIcon1.BalloonTipTitle = "Read Voyager";
                        notifyIcon1.ShowBalloonTip(500);
                                             
                        label_Number_Of_Trans_Voyager.Text = "Records =" + VoyData.Count.ToString("###,###,###,###");
                  
                        dgv_Voyager.DataSource = VoyData;
                    }
                }
                
            }


        }

        void Load_LocBill_File()
        {
            //if (dgv_Sales.Rows.Count > 1 || !checkBox_Check_Sales.Checked)
            if (true)
            {
                

                OpenFileDialog PPfile = new OpenFileDialog();

                PPfile.Multiselect = true;

                
                if (PPfile.ShowDialog() == DialogResult.OK)
                {

                    rtb_Pac_Pride_Files.Text = "";

                    foreach (string file in PPfile.FileNames)
                    {
                        string[] lines = File.ReadAllLines(file);

                        rtb_Pac_Pride_Files.AppendText(file+Environment.NewLine);

                        notifyIcon1.Icon = SystemIcons.Hand;
                        notifyIcon1.BalloonTipText = "Reading File "+ file ;
                        notifyIcon1.BalloonTipTitle = "Read Locbil";
                        notifyIcon1.ShowBalloonTip(500);

                        foreach (string line in lines)
                        {                         

                            string[] fields = line.Split(',');

                            DateTime actdate = DateTime.Now;

                            if (fields.Length > 10)
                            {

                                //MessageBox.Show(line);
                                fields[10] = fields[10].Replace("\"", "");
                                //MessageBox.Show(fields[10]);

                                DateTime.TryParse(fields[10], out actdate);

                                //MessageBox.Show(actdate.ToShortDateString());
                                if (fields[0].Contains("T6") && dtp_Date_Beg.Value.Date <= actdate && dtp_Date_End.Value.Date >= actdate)
                                {
                                    // DataRow newrow = locbil.NewRow();

                                    Locbill newrow = new Locbill();
                                                                       
                                    newrow.Record_Code = fields[0];
                                    newrow.SellingPart = fields[1];
                                    newrow.Site_Code = fields[2];
                                    newrow.Site_Type = fields[3];
                                    newrow.Site_Street = fields[4];
                                    newrow.Site_City = fields[5];
                                    newrow.Site_State = fields[6];
                                    newrow.Site_County = fields[7];
                                    newrow.Site_Zip = fields[8];
                                    newrow.Trans = fields[9];
                                    newrow.Trans_date = fields[10];
                                    newrow.Trans_Time = fields[11];
                                    newrow.Capture_Date = fields[12];
                                    newrow.Capture_Time = fields[13];
                                    newrow.Card = fields[14];
                                    newrow.Vehicle = fields[15];
                                    newrow.Host = fields[16];
                                    newrow.Identity = fields[17];
                                    newrow.Misc_Key = fields[18];
                                    newrow.Odometer = fields[19];
                                    newrow.Trans_Number = fields[20];
                                    newrow.Trans_Sequence = fields[21];
                                    newrow.Pump = fields[22];
                                    newrow.Hose = fields[23];
                                    newrow.Product = fields[24];
                                    newrow.Quantity = fields[25];
                                    newrow.UnitOfMeas = fields[26];
                                    newrow.SellingPrice = fields[27];
                                    newrow.TransferCost = fields[28];
                                    newrow.SalesTaxAdjust = fields[29];
                                    newrow.NetWorkICBDate = fields[30];
                                    newrow.Batch = fields[31];
                                    newrow.IssuerCode = fields[32];
                                    newrow.AuthCode = fields[33];
                                    newrow.ServLevel = fields[34];
                                    newrow.RetailTotal = fields[35];
                                    newrow.RetailPrice = fields[36];
                                    newrow.RetailRebate = fields[37];
                                    newrow.StateCode = fields[38];
                                    newrow.FET = fields[39];
                                    newrow.SET = fields[40];
                                    newrow.SET2 = fields[41];
                                    newrow.CountyET = fields[42];
                                    newrow.CityET = fields[43];
                                    newrow.PercentState = fields[44];
                                    newrow.PercentCounty = fields[45];
                                    newrow.PercentCity = fields[46];
                                    newrow.PercentOther = fields[47];
                                    newrow.OtherName = fields[48];
                                    newrow.FETRate = fields[49];
                                    newrow.SETRate = fields[50];
                                    newrow.OtherinTransfer = fields[51];
                                    newrow.CountyInTransfer = fields[52];
                                    newrow.CityInTransfer = fields[53];
                                    newrow.StateSalesInTransfer = fields[54];
                                    newrow.CountySalesInTransfer = fields[55];
                                    newrow.CitySalesInTransfer = fields[56];
                                    newrow.OtherInTransfer = fields[57];
                                    newrow.FETIncluded = fields[58];
                                    newrow.SETIncluded = fields[59];
                                    newrow.StateOtherInTransfer = fields[60];
                                    newrow.Include1 = fields[61];
                                    newrow.Include2 = fields[62];
                                    newrow.Include3 = fields[63];
                                    newrow.Include4 = fields[64];
                                    newrow.Include5 = fields[65];
                                    newrow.Include6 = fields[66];
                                    newrow.Include7 = fields[67];
                                    newrow.FETIncludedSelling = fields[68];
                                    newrow.SETIncludedSelling = fields[69];
                                    newrow.Include10 = fields[70];
                                    newrow.Include11 = fields[71];
                                    newrow.Include12 = fields[72];
                                    newrow.SalesTaxIncludeSelling = fields[73];
                                    newrow.Include14 = fields[74];
                                    newrow.Include15 = fields[75];
                                    newrow.Include16 = fields[76];
                                    newrow.ISONumber = fields[77];
                                    newrow.ISOCard = fields[78];
                                    

                                    locbil.Add(newrow);

                                    //if (checkBox_Check_Sales.Checked)
                                    //{
                                    //    //MessageBox.Show(" Trans="+newrow[20].ToString() +" Card="+ newrow[14].ToString());
                                    //    if (In_Sales_Data(newrow[20].ToString(), newrow[14].ToString()))
                                    //    {

                                    //        locbil.Rows.Add(newrow);
                                    //    }
                                    //}
                                    //else
                                    //{
                                    //    locbil.Rows.Add(newrow);
                                    //}
                                }
                            }
                        }
                    }

                    notifyIcon1.Icon = SystemIcons.Hand;
                    notifyIcon1.BalloonTipText = "Eliminating Duplicates";
                    notifyIcon1.BalloonTipTitle = "Read Locbil";
                    notifyIcon1.ShowBalloonTip(3000);

                    //MessageBox.Show("Press Enter to Eliminate Duplicates - This takes up to a few minutes");

                   

                   // DataTable dtv = RemoveDuplicateRows(locbil, "Card", "Trans_Number");

                    label_Number_Of_Trans_Locbil.Text = "Records ="+locbil.Count.ToString("###,###,###,###");
                    label_Pac_Pride_Records.Text = label_Number_Of_Trans_Locbil.Text;
                                                           

                    dgv_Locbil_Data.DataSource = locbil;
                  

                    Load_Extended_PP();
              
                }

            }
            else
            {
                MessageBox.Show("Load Sales Data or turn off requirement for continuity");
            }


        }

        public DataTable RemoveDuplicateRows(DataTable dTable, string colName1,string colName2)
        {
            List<int> dups = new List<int>();


            //Add list of all the unique item value to hashtable, which stores combination of key, value pair.
            //And add duplicate item value in arraylist.

            int notify = 50;

            for (int d=0; d<dTable.Rows.Count;d++)
            {
                notify++;

                if (notify > 50)
                {
                    notify = 0;
                    notifyIcon1.Icon = SystemIcons.Hand;
                    notifyIcon1.BalloonTipText = "Checking Record " + d.ToString() + " of " + dTable.Rows.Count.ToString();
                    notifyIcon1.BalloonTipTitle = "Checking Locbil";
                    notifyIcon1.ShowBalloonTip(3000);
                }
                for (int c=d+1; c<dTable.Rows.Count; c++ )
                {
                    

                    if (dTable.Rows[d][colName1]==dTable.Rows[c][colName1] && dTable.Rows[d][colName2] == dTable.Rows[c][colName2])
                    {
                        dups.Add(c);
                    }
                }

            }

            //Removing a list of duplicate items from datatable.
            foreach (int j in dups)
                dTable.Rows.Remove(dTable.Rows[j]);

            //Datatable which contains unique records will be return as output.
            return dTable;
        }


        string clean(string dirty)
        {
            string clean_str = string.Empty;

            dirty = dirty.Replace("\r\n", string.Empty);
            dirty = dirty.Replace("\n", string.Empty);
            dirty = dirty.Replace("\r", string.Empty);

            dirty = dirty.Replace("$", string.Empty);

            string lineSeparator = ((char)0x2028).ToString();
            string paragraphSeparator = ((char)0x2029).ToString();

            dirty = dirty.Replace(lineSeparator, string.Empty);
            dirty = dirty.Replace(paragraphSeparator, string.Empty);

            clean_str = dirty.Replace("\"", "");
                clean_str = dirty.Replace("(", "");
                clean_str = dirty.Replace(")", "");


            return clean_str;
        }

        void button_Get_Pac_Pride_Files_Click(object sender, EventArgs e)
        {
            Load_LocBill_File();
        }

        void Clear_District_Tax_Totals()
        {
            foreach (DataRow row in cadtax.Rows)
            {

                row["Adjustments"] = 0;
                row["Tax"] = 0;
                row["PaidPP"] = 0;
                row["PaidCFN"] = 0;
                row["Net"] = 0;
              

            }

            dgv_District_Taxes.DataSource = cadtax;
        }

        void Add_To_District_Sales_And_Taxes_STJ(string stj, decimal sales, decimal adj, decimal tax, decimal paidPP, decimal paidCFN)
        {
            string jsls = "";
            string jadj = "";
            string jtax = "";
            string jpdpp = "";
            string jpdcfn = "";
            string jnet = "";
            string exp = "";
            string jstj = "";

            decimal amt = 0.00m;
            decimal samt = 0.00m;
            decimal aamt = 0.00m;
            decimal tamt = 0.00m;
            decimal pamtpp = 0.00m;
            decimal pamtcfn = 0.00m;
            decimal netamt = 0.00m;

            DateTime stjdate = new DateTime();

            string jcity = "";
            string jcnty = "";

            foreach (DataRow row in cadtax.Rows)
            {
                jcnty = row[0].ToString();
                jcity = row[1].ToString();
                jstj = row[2].ToString();
                jsls = row[6].ToString();
                jadj = row[7].ToString();
                jtax = row[8].ToString();
                jpdpp = row[9].ToString();
                jpdcfn = row[10].ToString();
                exp = row[4].ToString();
                jnet = row[11].ToString();

                DateTime.TryParse(exp, out stjdate);

                if (jstj == stj)
                {

                    decimal.TryParse(jsls, out amt);
                    samt = amt + sales;

                    amt = 0;

                    decimal.TryParse(jadj, out amt);
                    aamt = amt + adj;

                    amt = 0;

                    decimal.TryParse(jtax, out amt);
                    tamt = amt + tax;

                    amt = 0;

                    decimal.TryParse(jpdpp, out amt);

                    pamtpp = amt + paidPP;

                    amt = 0;

                    decimal.TryParse(jpdcfn, out amt);

                    pamtcfn = amt + paidCFN;

                    amt = 0;

                    row[6] = samt.ToString("###,###,###.##");
                    row[7] = aamt.ToString("###,###,###.##");
                    row[8] = tamt.ToString("###,###,###.##");
                    row[9] = pamtpp.ToString("###,###,###.##");
                    row[10] = pamtcfn.ToString("###,###,###.##");

                }
            }

            dgv_District_Taxes.DataSource = cadtax;
        }

        void Add_To_District_Sales_And_Taxes(string city, string county, decimal sales, decimal adj, decimal tax, decimal paidPP, decimal paidCFN, decimal paidVoy, decimal paidWEX, DateTime enddate)
        {
            if (city.Length == 0 && county.Length == 0)
            {
                return;
            }

            city = clean(NoSpace(city)).ToUpper();
            county = clean(NoSpace(county)).ToUpper();

            string jsls = "";
            string jadj = "";
            string jtax = "";
            string jpdpp = "";
            string jpdcfn = "";
            string jpdvoy = "";
            string jpdwex = "";
            string jnet = "";
            string exp = "";
            string stj = "";

            decimal amt = 0.00m;
            decimal samt = 0.00m;
            decimal aamt = 0.00m;
            decimal tamt = 0.00m;
            decimal pamtpp = 0.00m;
            decimal pamtcfn = 0.00m;
            decimal pamtvoy = 0.00m;
            decimal pamtwex = 0.00m;
            decimal netamt = 0.00m;

            string jcity = "";
            string jcnty = "";

            foreach (DataRow row in cadtax.Rows)
            {
                jcnty = row["County"].ToString();
                jcity = row["City"].ToString();
                stj = row["STJ"].ToString();
                jsls = row["Sales"].ToString();
                jadj = row["Adjustments"].ToString();
                jtax = row["Tax"].ToString();
                jpdpp = row["PaidPP"].ToString();
                jpdcfn = row["PaidCFN"].ToString();
                jpdvoy = row["PaidVoy"].ToString();
                jpdwex = row["PaidWEX"].ToString();

                exp = row["Expired"].ToString();
                jnet = row["Net"].ToString();


                jcity = clean(NoSpace(jcity)).ToUpper();
                jcnty = clean(NoSpace(jcnty)).ToUpper();

                jcnty = jcnty.Replace("COUNTY", "");
                county = county.Replace("COUNTY", "");

                city = city.Replace("CITY OF", "");
                city = city.Replace("OF", "");
                city = city.Replace(",", "");
                city = city.Trim();

                city = city.Replace("\"", "");
                county = county.Replace("\"", "");

                bool check_add = false;
                
                DateTime effdate = DateTime.Now;
                DateTime expdate = DateTime.Now;

                if (row["Effective"].ToString().Length > 0)
                {
                    DateTime.TryParse(row["Effective"].ToString(), out effdate);
                }
                else
                {
                    effdate = dtp_Date_Beg.Value;
                }
                if (row["Expired"].ToString().Length > 0)
                {
                    DateTime.TryParse(row["Expired"].ToString(), out expdate);
                }
                else
                {
                    expdate = dtp_Date_End.Value;
                }

                if (checkBox_Debug.Checked || check_add )
                {
                    MessageBox.Show("jcnty =" + jcnty + " County = " + county + " jcity = " + jcity + " City = " + city);
                }

                //if (city.ToUpper().Contains("COALINGA"))
                //{
                //    decimal rating = alpharate(jcity, city);

                //    MessageBox.Show("jcnty =" + jcnty + " County = " + county + " jcity = " + jcity + " City = " + city + " Rating = " + rating.ToString("##.######"));
                //}

                //if(jcity.Contains("KING") && city.Contains("KING"))
                //{
                //   MessageBox.Show("City Data = " + city + " Calif City = " + jcity+" County Data = "+county+" Calif County = "+jcnty);
                //}

                if ( jcnty == county)
                {
                  
                    if (jcity == city || (city.Length<1 && jcity.Length <1))
                    {
                        if (effdate <= enddate && expdate >= enddate)
                        {
                            if (checkBox_Debug.Checked || check_add)
                            {
                                MessageBox.Show("Effective =" + effdate.ToShortDateString() + " Expired = " + expdate.ToShortDateString() + " Passed Date = " + enddate.ToShortDateString());
                            }

                            decimal.TryParse(jsls, out amt);
                            samt = amt + sales;

                            amt = 0;

                            decimal.TryParse(jadj, out amt);
                            aamt = amt + adj;

                            if (paidCFN > 0)
                            {
                                amt = 0;
                                decimal.TryParse(row["CFNPurch"].ToString(), out amt);
                                amt = amt + adj;
                                row["CFNPurch"] = amt;
                            }

                            if (paidPP > 0)
                            {
                                amt = 0;
                                decimal.TryParse(row["PPPurch"].ToString(), out amt);
                                amt = amt + adj;
                                row["PPPurch"] = amt;
                            }

                            if (paidVoy > 0)
                            {
                                amt = 0;
                                decimal.TryParse(row["VoyPurch"].ToString(), out amt);
                                amt = amt + adj;
                                row["VoyPurch"] = amt;
                            }

                            if (paidWEX > 0)
                            {
                                amt = 0;
                                decimal.TryParse(row["WEXPurch"].ToString(), out amt);
                                amt = amt + adj;
                                row["WEXPurch"] = amt;
                            }

                            amt = 0;

                            decimal.TryParse(jtax, out amt);
                            tamt = amt + tax;

                            amt = 0;

                            decimal.TryParse(jpdpp, out amt);

                            pamtpp = amt + paidPP;

                            amt = 0;

                            decimal.TryParse(jpdcfn, out amt);

                            pamtcfn = amt + paidCFN;

                            amt = 0;

                            decimal.TryParse(jpdvoy, out amt);

                            pamtvoy = amt + paidVoy;

                            amt = 0;

                            decimal.TryParse(jpdwex, out amt);

                            pamtwex = amt + paidWEX;

                            row["Sales"] = samt;
                            row["Adjustments"] = aamt;
                            row["Tax"] = tamt;
                            row["PaidPP"] = pamtpp;
                            row["PaidCFN"] = pamtcfn;
                            row["PaidVoy"] = pamtvoy;
                            row["PaidWEX"] = pamtwex;
                        }
                    }
                }


                //if (jcnty == county && jcity == city && exp.Length < 2 && stj.Length > 2)
                //{

                //    decimal.TryParse(jsls, out amt);
                //    samt = amt + sales;

                //    amt = 0;

                //    decimal.TryParse(jadj, out amt);
                //    aamt = amt + adj;

                //    if (paidCFN > 0)
                //    {
                //        amt = 0;
                //        decimal.TryParse(row["CFNPurch"].ToString(), out amt);
                //        amt = amt + adj;
                //        row["CFNPurch"] = amt;
                //    }

                //    if (paidPP > 0)
                //    {
                //        amt = 0;
                //        decimal.TryParse(row["PPPurch"].ToString(), out amt);
                //        amt = amt + adj;
                //        row["PPPurch"] = amt;
                //    }

                //    if (paidVoy > 0)
                //    {
                //        amt = 0;
                //        decimal.TryParse(row["VoyPurch"].ToString(), out amt);
                //        amt = amt + adj;
                //        row["VoyPurch"] = amt;
                //    }

                //    if (paidWEX > 0)
                //    {
                //        amt = 0;
                //        decimal.TryParse(row["WEXPurch"].ToString(), out amt);
                //        amt = amt + adj;
                //        row["WEXPurch"] = amt;
                //    }
                                        
                //    amt = 0;

                //    decimal.TryParse(jtax, out amt);
                //    tamt = amt + tax;

                //    amt = 0;

                //    decimal.TryParse(jpdpp, out amt);

                //    pamtpp = amt + paidPP;

                //    amt = 0;

                //    decimal.TryParse(jpdcfn, out amt);

                //    pamtcfn = amt + paidCFN;

                //    amt = 0;

                //    decimal.TryParse(jpdvoy, out amt);

                //    pamtvoy = amt + paidVoy;

                //    amt = 0;

                //    decimal.TryParse(jpdwex, out amt);

                //    pamtwex = amt + paidWEX;

                //    row["Sales"] = samt;
                //    row["Adjustments"] = aamt;
                //    row["Tax"] = tamt;
                //    row["PaidPP"] = pamtpp;
                //    row["PaidCFN"] = pamtcfn;
                //    row["PaidVoy"] = pamtvoy;
                //    row["PaidWEX"] = pamtwex;

                   
                //}
            }

            dgv_District_Taxes.DataSource = cadtax;
        }

        string Get_County_With_City(string city)
        {
            if(city.Length < 1)
            {
                return "";
            }

            string county = "";

            city = clean(NoSpace(city)).ToUpper();
            city = city.Replace("CITY OF", "");
            city = city.Replace("OF", "");
            city = city.Replace(",", "");

            foreach (DataRow row in cadtax.Rows)
            {
                
                string jcity = clean(NoSpace(row["City"].ToString())).ToUpper();

                if (checkBox_Debug.Checked)
                {
                    MessageBox.Show("Get_County_with_city JCITY =" + jcity + " CITY=" + city);
                }

                if (jcity.Equals(city))
                {
                    county = row["County"].ToString();
                }
            }

            return county;
        }


        void Calc_District_taxes()
        {
            Clear_District_Tax_Totals();

            string cst = "";
            string adj = "";
            string tax = "";
            string d_tax = "";
            string s_tax = "";
            string pd = "";
            string q = "";
            string sset = "";
            string sls = "";

            decimal quan = 0.00m;
            decimal amt = 0.00m;
            decimal camt = 0.00m;
            decimal aamt = 0.00m;
            decimal tamt = 0.00m;
            decimal stax_amt = 0.00m;
            decimal dtax_amt = 0.00m;
            decimal pamt = 0.00m;
            decimal set = 0.00m;
            decimal slsamt = 0.00m;

            // CFN Only

            decimal citytax = 0.00m;
            decimal cntytax = 0.00m;

            string citytaxamt = "";
            string cntytaxamt = "";

            salestaxpaiddiesel = 0.00m;
            salestaxpaiddyed = 0.00m;
            salestaxpaidgas = 0.00m;
            salestaxpaidother = 0.00m;
            salestaxpaidtotal = 0.00m;

            salestaxpaiddiesel_pp = 0.00m;
            salestaxpaiddiesel_cfn = 0.00m;
            salestaxpaiddyed_pp = 0.00m;
            salestaxpaiddyed_cfn = 0.00m;
            salestaxpaidgas_pp = 0.00m;
            salestaxpaidgas_cfn = 0.00m;
            salestaxpaidother_pp = 0.00m;
            salestaxpaidother_cfn = 0.00m;
            salestaxpaidtotal_pp = 0.00m;
            salestaxpaidtotal_cfn = 0.00m;

            salestaxpaiddiesel_voy = 0.00m;
            salestaxpaiddyed_voy = 0.00m;
            salestaxpaidgas_voy = 0.00m;
            salestaxpaidother_voy = 0.00m;
            salestaxpaidtotal_voy = 0.00m;

            salestaxpaiddiesel_wex = 0.00m;
            salestaxpaiddyed_wex = 0.00m;
            salestaxpaidgas_wex = 0.00m;
            salestaxpaidother_wex = 0.00m;
            salestaxpaidtotal_wex = 0.00m;

            extcostdiesel = 0;
            extcostdiesel_pp = 0;
            extcostdiesel_cfn = 0;
            extcostdiesel_voy = 0;
            extcostdiesel_wex = 0;

            extcostdyed = 0;
            extcostdyed_pp = 0;
            extcostdyed_cfn = 0;
            extcostdyed_voy = 0;
            extcostdyed_wex = 0;

            extcostgas = 0;
            extcostgas_pp = 0;
            extcostgas_cfn = 0;
            extcostgas_voy = 0;
            extcostgas_wex = 0;
            
            extcostother = 0;
            extcostother_pp = 0;
            extcostother_cfn = 0;
            extcostother_voy = 0;
            extcostother_wex = 0;

            extcosttotal = 0;
            extcosttotal_pp = 0;
            extcosttotal_cfn = 0;
            extcosttotal_voy = 0;
            extcosttotal_wex = 0;

            galstotalext = 0.00m;
            galsdieselext = 0.00m;
            galsdyedext = 0.00m;
            gals87gasext = 0.00m;
            gals89gasext = 0.00m;
            gals91gasext = 0.00m;

            galsothext = 0.00m;
            galsoth = 0.00m;

            galsdieselext_pp = 0;
            galsdieselext_cfn = 0;
            galsdieselext_voy = 0;
            galsdieselext_wex = 0;

            galsdyedext_pp = 0;
            galsdyedext_cfn = 0;
            galsdyedext_voy = 0;
            galsdyedext_wex = 0;

            galsgasext_pp = 0;
            galsgasext_cfn = 0;
            galsgasext_voy = 0;
            galsgasext_wex = 0;

            galsotherext_pp = 0;
            galsotherext_cfn = 0;
            galsotherext_voy = 0;
            galsotherext_wex = 0;

            string city = "";
            string cnty = "";
            string state = "";
            string prod = "";

            decimal jrate = 0.00m;

            int fta = 0;
            decimal adjsls = 0.00m;
            decimal adjcost = 0.00m;
            decimal taxamt = 0.00m;

            foreach (Locbill row in locbil_de_ca)
            {
                city = row.Site_City.ToString();
                cnty = row.Site_County.ToString();

                q = row.Quantity.ToString();
                cst = row.TransferCost.ToString();
                tax = row.SalesTaxAdjust.ToString();
                prod = row.Product.ToString();
                sset = row.SET.ToString();

              //  MessageBox.Show("City Before STJ check = " + city);

                //MessageBox.Show("Quan " + q + " Cost " + cst + " Prod " + prod);
                if (!Check_City_Has_STJ(city))
                {
                    //MessageBox.Show("City = " + city);

                    city = "";
                }


                //MessageBox.Show("City After STJ check = " + city);


                jrate = Get_Jurisdiction_Rate(city, cnty, dtp_Date_End.Value);

                bool check_loc = false;

                decimal.TryParse(q, out quan);
                decimal.TryParse(cst, out camt);
                decimal.TryParse(tax, out tamt);
                decimal.TryParse(sset, out set);

                fta = Lookup_Prod_FTA(prod);

                if (checkBox_Debug.Checked || check_loc)
                {
                    MessageBox.Show("Quan " + quan.ToString() + " Cost " + camt.ToString() + " SET " + set.ToString() + " Tax " + tamt.ToString() + " fta " + fta.ToString());
                }

                if (tamt > 0)
                {
                    salestaxpaidtotal = salestaxpaidtotal + (tamt * quan);
                    salestaxpaidtotal_pp = salestaxpaidtotal_pp + (tamt * quan);

                    galstotalext = galstotalext + quan;

                    if (fta == 167)
                    {
                        adjcost = (camt - set - tamt) * quan;
                        taxamt = jrate * adjcost;

                        if (checkBox_Debug.Checked || check_loc)
                        {
                            MessageBox.Show("City = " + city + " County = " + cnty + "Purchases = " + adjcost.ToString() + " Tax = " + taxamt.ToString()+" Rate = "+jrate.ToString());
                        }

                        Add_To_District_Sales_And_Taxes(city, cnty, 0.00m, adjcost, 0.00m, taxamt,0,0,0,dtp_Date_End.Value);

                        salestaxpaiddiesel = salestaxpaiddiesel + (tamt * quan);
                        salestaxpaiddiesel_pp = salestaxpaiddiesel_pp + (tamt * quan);

                        galsdieselext = galsdieselext + quan;
                        galsdieselext_pp = galsdieselext_pp + quan;

                        disttaxpaiddiesel = disttaxpaiddiesel + taxamt;
                        disttaxpaidtotal = disttaxpaidtotal + taxamt;

                        extcostdiesel = extcostdiesel + adjcost;
                        extcostdiesel_pp = extcostdiesel_pp + adjcost;

                        extcosttotal_pp = extcosttotal_pp + adjcost;
                        extcosttotal = extcosttotal + adjcost;

                      //  MessageBox.Show(extcostdiesel.ToString());
                    }

                    if (fta == 227)
                    {
                        adjcost = (camt - set - tamt) * quan;
                        taxamt = jrate * adjcost;

                        if (checkBox_Debug.Checked || check_loc)
                        {
                            MessageBox.Show("City = " + city + " County = " + cnty + "Purchases = " + adjcost.ToString() + " Tax = " + taxamt.ToString() + " Rate = " + jrate.ToString());
                        }

                        Add_To_District_Sales_And_Taxes(city, cnty, 0.00m, adjcost, 0.00m, taxamt, 0,0,0,dtp_Date_End.Value);

                        salestaxpaiddyed = salestaxpaiddyed + (tamt * quan);
                        salestaxpaiddyed_pp = salestaxpaiddyed_pp + (tamt * quan);

                        galsdyedext = galsdyedext + quan;
                        galsdyedext_pp = galsdyedext_pp + quan;

                        disttaxpaiddyed = disttaxpaiddyed + taxamt;
                        disttaxpaidtotal = disttaxpaidtotal + taxamt;

                        extcostdyed = extcostdyed + adjcost;
                        extcostdyed_pp = extcostdyed_pp + adjcost;

                        extcosttotal = extcosttotal + adjcost;
                        extcosttotal_pp = extcosttotal_pp + adjcost;
                    }

                    if (fta == 65)
                    {
                        adjcost = (camt - tamt) * quan;
                        taxamt = jrate * adjcost;

                        if (checkBox_Debug.Checked || check_loc)
                        {
                            MessageBox.Show("City = " + city + " County = " + cnty + "Purchases = " + adjcost.ToString() + " Tax = " + taxamt.ToString() + " Rate = " + jrate.ToString());
                        }
                        Add_To_District_Sales_And_Taxes(city, cnty, 0.00m, adjcost, 0.00m, taxamt, 0,0,0,dtp_Date_End.Value);

                        salestaxpaidgas = salestaxpaidgas + (tamt * quan);
                        salestaxpaidgas_pp = salestaxpaidgas_pp + (tamt * quan);

                        gals87gasext = gals87gasext + quan;
                        galsgasext_pp = galsgasext_pp + quan;

                        disttaxpaidgas = disttaxpaidgas + taxamt;
                        disttaxpaidtotal = disttaxpaidtotal + taxamt;


                        extcostgas = extcostgas + adjcost;
                        extcostgas_pp = extcostgas_pp + adjcost;

                        extcosttotal = extcosttotal + adjcost;
                        extcosttotal_pp = extcosttotal_pp + adjcost;
                    }

                    if (fta == 0)
                    {
                        adjcost = (camt - tamt) * quan;
                        taxamt = jrate * adjcost;

                        if (checkBox_Debug.Checked || check_loc)
                        {
                            MessageBox.Show("City = " + city + " County = " + cnty + "Purchases = " + adjcost.ToString() + " Tax = " + taxamt.ToString() + " Rate = " + jrate.ToString());
                        }
                        Add_To_District_Sales_And_Taxes(city, cnty, 0.00m, adjcost, 0.00m, taxamt, 0,0,0,dtp_Date_End.Value);

                        salestaxpaidother = salestaxpaidother + (tamt * quan);
                        salestaxpaidother_pp = salestaxpaidother_pp + (tamt * quan);

                        galsothext = galsothext + quan;
                        galsotherext_pp = galsotherext_pp + quan;

                        disttaxpaidother = disttaxpaidother + taxamt;
                        disttaxpaidtotal = disttaxpaidtotal + taxamt;

                        extcostother = extcostother + adjcost;
                        extcostother_pp = extcostother_pp + adjcost;

                        extcosttotal = extcosttotal + adjcost;
                        extcosttotal_pp = extcosttotal_pp + adjcost;
                    }
                }

                if (checkBox_Debug.Checked)
                {
                    MessageBox.Show("Cost " + extcosttotal.ToString() + " diesel " + extcostdiesel.ToString() + " Gas " + extcostgas.ToString());
                }
            }

            foreach (PTFile row in CFNDataExt)
            {
                city = row.City;
                cnty = row.County;

                q = row.Quantity;

                citytaxamt = row.CityTax;
                cntytaxamt = row.CountyTax;
                cst = "0";
                tax = row.StateTax;
                prod = row.Product;
                sset = row.SET;

                int tprod = 0;

                int.TryParse(prod, out tprod);

                prod = tprod.ToString();

                //MessageBox.Show("Quan " + q + " Cost " + cst + " Prod " + prod);
                if (!Check_City_Has_STJ(city))
                {
                    city = "";
                }

                jrate = Get_Jurisdiction_Rate(city, cnty,dtp_Date_End.Value);

                decimal.TryParse(q, out quan);
                decimal.TryParse(cst, out camt);
                decimal.TryParse(tax, out tamt);
                decimal.TryParse(sset, out set);
                decimal.TryParse(citytaxamt, out citytax);
                decimal.TryParse(cntytaxamt, out cntytax);


                fta = Lookup_CFNProd_FTA(prod);

                //MessageBox.Show("Quan " + quan.ToString() + " Cost " + camt.ToString() +" SET "+set.ToString()+" Tax "+tamt.ToString()+" fta " + fta.ToString()+" "+prod);

                if (tamt > 0)
                {
                    salestaxpaidtotal = salestaxpaidtotal + (tamt);
                    salestaxpaidtotal_cfn = salestaxpaidtotal_cfn + (tamt);

                    galstotalext = galstotalext + quan;

                    if (fta == 167)
                    {
                        decimal.TryParse(row.Taxable,out adjcost);
                        taxamt = citytax+cntytax;

                        Add_To_District_Sales_And_Taxes(city, cnty, 0.00m, adjcost, 0.00m, 0.00m, taxamt,0,0,dtp_Date_End.Value);

                        salestaxpaiddiesel = salestaxpaiddiesel + tamt;
                        salestaxpaiddiesel_cfn = salestaxpaiddiesel_cfn + tamt;

                        galsdieselext = galsdieselext + quan;
                        galsdieselext_cfn = galsdieselext_cfn + quan;

                        disttaxpaiddiesel = disttaxpaiddiesel + taxamt;
                        disttaxpaidtotal = disttaxpaidtotal + taxamt;

                        extcostdiesel = extcostdiesel + adjcost;
                        extcostdiesel_cfn = extcostdiesel_cfn + adjcost;

                        extcosttotal = extcosttotal + adjcost;
                        extcosttotal_cfn = extcosttotal_cfn + adjcost;

                        //  MessageBox.Show(extcostdiesel.ToString());
                    }

                    if (fta == 227)
                    {
                        decimal.TryParse(row.Taxable, out adjcost);
                        taxamt = citytax + cntytax;

                        Add_To_District_Sales_And_Taxes(city, cnty, 0.00m, adjcost, 0.00m, 0.00m, taxamt,0,0,dtp_Date_End.Value);

                        salestaxpaiddyed = salestaxpaiddyed + tamt;
                        salestaxpaiddyed_cfn = salestaxpaiddyed_cfn + tamt;

                        galsdyedext = galsdyedext + quan;
                        galsdyedext_cfn = galsdyedext_cfn + quan;

                        disttaxpaiddyed = disttaxpaiddyed + taxamt;
                        disttaxpaidtotal = disttaxpaidtotal + taxamt;

                        extcostdyed = extcostdyed + adjcost;
                        extcostdyed_cfn = extcostdyed_cfn + adjcost;

                        extcosttotal = extcosttotal + adjcost;
                        extcosttotal_cfn = extcosttotal_cfn + adjcost;
                    }

                    if (fta == 65)
                    {
                        decimal.TryParse(row.Taxable, out adjcost);
                        taxamt = citytax + cntytax;

                        Add_To_District_Sales_And_Taxes(city, cnty, 0.00m, adjcost, 0.00m, 0.00m, taxamt,0,0,dtp_Date_End.Value);

                        salestaxpaidgas = salestaxpaidgas + tamt;
                        salestaxpaidgas_cfn = salestaxpaidgas_cfn + tamt;

                        gals87gasext = gals87gasext + quan;
                        galsgasext_cfn = galsgasext_cfn + quan;

                        disttaxpaidgas = disttaxpaidgas + taxamt;
                        disttaxpaidtotal = disttaxpaidtotal + taxamt;
                        
                        extcostgas = extcostgas + adjcost;
                        extcostgas_cfn = extcostgas_cfn + adjcost;

                        extcosttotal = extcosttotal + adjcost;
                        extcosttotal_cfn = extcosttotal_cfn + adjcost;
                    } 

                    if (fta == 0)
                    {
                        decimal.TryParse(row.Taxable, out adjcost);
                        taxamt = citytax + cntytax;

                        Add_To_District_Sales_And_Taxes(city, cnty, 0.00m, adjcost, 0.00m, 0.00m, taxamt,0,0,dtp_Date_End.Value);

                        salestaxpaidother = salestaxpaidother + tamt;
                        salestaxpaidother_cfn = salestaxpaidother_cfn + tamt;

                        galsothext = galsothext + quan;
                        galsotherext_cfn = galsotherext_cfn + quan;

                        disttaxpaidother = disttaxpaidother + taxamt;
                        disttaxpaidtotal = disttaxpaidtotal + taxamt;

                        extcostother = extcostother + adjcost;
                        extcostother_cfn = extcostother_cfn + adjcost;

                        extcosttotal = extcosttotal + adjcost;
                        extcosttotal_cfn = extcosttotal_cfn + adjcost;
                    }
                }
                                              

                //MessageBox.Show("Cost " + extcosttotal.ToString() + " diesel " + extcostdiesel.ToString()+" Gas "+extcostgas.ToString());

            }


            foreach (VoyagerData row in VoyData)
            {
                city = row.SiteCity;
                cnty = row.County;

                q = row.Quantity;

                citytaxamt = row.TaxDistrict;
                cntytaxamt = row.TaxDistrict;

                cst = "0";
                tax = row.Tax;
                s_tax = row.TaxState;
                d_tax = row.TaxDistrict;
                prod = row.Product;
                sset = row.SET;

                int tprod = 0;

                int.TryParse(prod, out tprod);

                prod = tprod.ToString();

                //MessageBox.Show("Quan " + q + " Cost " + cst + " Prod " + prod);
               
                jrate = Get_Jurisdiction_Rate(city, cnty,dtp_Date_End.Value);

                decimal.TryParse(q, out quan);
                decimal.TryParse(cst, out camt);
                decimal.TryParse(tax, out tamt);
                decimal.TryParse(s_tax, out stax_amt);
                decimal.TryParse(d_tax, out dtax_amt);
                decimal.TryParse(sset, out set);
                decimal.TryParse(citytaxamt, out citytax);
                decimal.TryParse(cntytaxamt, out cntytax);


                int.TryParse(row.FTA,out fta);

                //MessageBox.Show("Quan " + quan.ToString() + " Cost " + camt.ToString() +" SET "+set.ToString()+" Tax "+tamt.ToString()+" fta " + fta.ToString()+" "+prod);

                if (tamt > 0)
                {
                    salestaxpaidtotal = salestaxpaidtotal + (tamt);
                    salestaxpaidtotal_voy = salestaxpaidtotal_voy + (tamt);
                    disttaxpaidtotal_voy = disttaxpaidtotal_voy + dtax_amt;

                    galstotalext = galstotalext + quan;
                    galstotalext_voy = galstotalext_voy + quan;

                    if (fta == 167)
                    {
                        decimal.TryParse(row.TaxBase, out adjcost);
                        taxamt = citytax ;

                        Add_To_District_Sales_And_Taxes(city, cnty, 0.00m, adjcost, 0.00m, 0.00m, 0,taxamt,0,dtp_Date_End.Value);

                        salestaxpaiddiesel = salestaxpaiddiesel + tamt;
                        salestaxpaiddiesel_voy = salestaxpaiddiesel_voy + tamt;

                        galsdieselext = galsdieselext + quan;
                        galsdieselext_voy = galsdieselext_voy + quan;

                        disttaxpaiddiesel = disttaxpaiddiesel + dtax_amt;
                        disttaxpaidtotal = disttaxpaidtotal + dtax_amt;
                        disttaxpaidtotal_voy = disttaxpaidtotal_voy + dtax_amt;


                        extcostdiesel = extcostdiesel + adjcost;
                        extcostdiesel_voy = extcostdiesel_voy + adjcost;

                        extcosttotal = extcosttotal + adjcost;
                        extcosttotal_voy = extcosttotal_voy + adjcost;

                        //  MessageBox.Show(extcostdiesel.ToString());
                    }

                    if (fta == 227)
                    {
                        decimal.TryParse(row.TaxBase, out adjcost);
                        taxamt = citytax ;

                        Add_To_District_Sales_And_Taxes(city, cnty, 0.00m, adjcost, 0.00m, 0.00m, 0,taxamt,0,dtp_Date_End.Value);

                        salestaxpaiddyed = salestaxpaiddyed + tamt;
                        salestaxpaiddyed_voy = salestaxpaiddyed_voy + tamt;

                        galsdyedext = galsdyedext + quan;
                        galsdyedext_voy = galsdyedext_voy + quan;

                        disttaxpaiddyed = disttaxpaiddyed + taxamt;
                        disttaxpaidtotal = disttaxpaidtotal + taxamt;

                        extcostdyed = extcostdyed + adjcost;
                        extcostdyed_voy = extcostdyed_voy + adjcost;

                        extcosttotal = extcosttotal + adjcost;
                        extcosttotal_voy = extcosttotal_voy + adjcost;
                    }

                    if (fta == 65)
                    {
                        decimal.TryParse(row.TaxBase, out adjcost);
                        taxamt = citytax ;

                        Add_To_District_Sales_And_Taxes(city, cnty, 0.00m, adjcost, 0.00m, 0.00m, 0,taxamt,0,dtp_Date_End.Value);

                        salestaxpaidgas = salestaxpaidgas + tamt;
                        salestaxpaidgas_voy = salestaxpaidgas_voy + tamt;

                        gals87gasext = gals87gasext + quan;
                        galsgasext_voy = galsgasext_voy + quan;

                        disttaxpaidgas = disttaxpaidgas + taxamt;
                        disttaxpaidtotal = disttaxpaidtotal + taxamt;

                        extcostgas = extcostgas + adjcost;
                        extcostgas_voy = extcostgas_voy + adjcost;

                        extcosttotal = extcosttotal + adjcost;
                        extcosttotal_voy = extcosttotal_voy + adjcost;
                    }

                    if (fta == 0)
                    {
                        decimal.TryParse(row.TaxBase, out adjcost);
                        taxamt = citytax;

                        Add_To_District_Sales_And_Taxes(city, cnty, 0.00m, adjcost, 0.00m, 0.00m,0, taxamt,0,dtp_Date_End.Value);

                        salestaxpaidother = salestaxpaidother + tamt;
                        salestaxpaidother_voy = salestaxpaidother_voy + tamt;

                        galsothext = galsothext + quan;
                        galsotherext_voy = galsotherext_voy + quan;

                        disttaxpaidother = disttaxpaidother + taxamt;
                        disttaxpaidtotal = disttaxpaidtotal + taxamt;

                        extcostother = extcostother + adjcost;
                        extcostother_voy = extcostother_voy + adjcost;

                        extcosttotal = extcosttotal + adjcost;
                        extcosttotal_voy = extcosttotal_voy + adjcost;
                    }
                }


                //MessageBox.Show("Cost " + extcosttotal.ToString() + " diesel " + extcostdiesel.ToString()+" Gas "+extcostgas.ToString());

            }





            tb_Deductions_Cost_Of_Tax_Paid_MVF.Text = extcostgas.ToString("###,###,###.##");
            tb_Deductions_Cost_Tax_Paid_Diesel_Resold.Text = extcostdiesel.ToString("###,###,###.##");
            tb_Deductions_Cost_Of_Tax_Paid_Purchases_Other.Text = extcostdyed.ToString("###,###,###.##");
            tb_Deductions_Cost_Of_Tax_Paid_Purchases_Other.Text = extcostother.ToString("###,###,###.##");
            tb_Deductions_Cost_Of_Tax_Paid_Purchases_Total.Text = extcosttotal.ToString("###,###,###.##");

            tb_CTPP_PP_Diesel.Text = extcostdiesel_pp.ToString("###,###,###,###.##");
            tb_CTPP_PP_Dyed.Text = extcostdyed_pp.ToString("###,###,###,###,###.##");
            tb_CTPP_PP_MVF.Text = extcostgas_pp.ToString("###,###,###,###,###.##");
            tb_CTPP_PP_Other.Text = extcostother_pp.ToString("###,###,###,###,###.##");
            tb_CTPP_PP_Total.Text = extcosttotal_pp.ToString("###,###,###,###,###.##");

            tb_Gallons_PP_Diesel.Text = galsdieselext_pp.ToString("###,###,###,###.##");
            tb_Gallons_PP_Dyed.Text = galsdyedext_pp.ToString("###,###,###,###.##");
            tb_Gallons_PP_MVF.Text = galsgasext_pp.ToString("###,###,###,###.##");
            tb_Gallons_PP_Other.Text = galsotherext_pp.ToString("###,###,###.##");
            tb_Gallons_PP_Total.Text = galstotalext_pp.ToString("###,###,###.##");

            tb_TP_PP_Diesel.Text = salestaxpaiddiesel_pp.ToString("###,###,###,###,###.##");
            tb_TP_PP_Dyed.Text = salestaxpaiddyed_pp.ToString("####,###,###,###,###.##");
            tb_TP_PP_MVF.Text = salestaxpaidgas_pp.ToString("###,###,###,###,###.##");
            tb_TP_PP_Other.Text = salestaxpaidother_pp.ToString("###,###,###,###,###.##");
            tb_TP_PP_Total.Text = salestaxpaidtotal_pp.ToString("###,###,###,###,###.##");

            tb_TP_CFN_Diesel.Text = salestaxpaiddiesel_cfn.ToString("###,###,###,###,###.##");
            tb_TP_CFN_Dyed.Text = salestaxpaiddyed_cfn.ToString("###,###,###,###,###.##");
            tb_TP_CFN_MVF.Text = salestaxpaidgas_cfn.ToString("###,###,###,###,###.##");
            tb_TP_CFN_Other.Text = salestaxpaidother_cfn.ToString("###,###,###,###,###.##");
            tb_TP_CFN_Total.Text = salestaxpaidtotal_cfn.ToString("###,###,###,###,###.##");

            tb_Gallons_CFN_Diesel.Text = galsdieselext_cfn.ToString("###,###,###,###.##");
            tb_Gallons_CFN_Dyed.Text = galsdyedext_cfn.ToString("###,###,###,###.##");
            tb_Gallons_CFN_MVF.Text = galsgasext_cfn.ToString("###,###,###,###.##");
            tb_Gallons_CFN_Other.Text = galsotherext_cfn.ToString("###,###,###.##");
            tb_Gallons_CFN_Total.Text = galstotalext_cfn.ToString("###,###,###.##");

            tb_CTPP_CFN_Diesel.Text = extcostdiesel_cfn.ToString("###,###,###,###,###,###.##");
            tb_CTPP_CFN_Dyed.Text = extcostdyed_cfn.ToString("###,###,###,###,###,###.##");
            tb_CTPP_CFN_MVF.Text = extcostgas_cfn.ToString("###,###,###,###,###,###.##");
            tb_CTPP_CFN_Other.Text = extcostother_cfn.ToString("###,###,###,###,###.##");
            tb_CTPP_CFN_Total.Text = extcosttotal_cfn.ToString("###,###,###,###,###.##");

            tb_CTPP_Voy_Diesel.Text = extcostdiesel_voy.ToString("###,###,###,###.##");
            tb_CTPP_Voy_Dyed.Text = extcostdyed_voy.ToString("###,###,###,###,###.##");
            tb_CTPP_Voy_MVF.Text = extcostgas_voy.ToString("###,###,###,###,###.##");
            tb_CTPP_Voy_Other.Text = extcostother_voy.ToString("###,###,###,###,###.##");
            tb_CTPP_Voy_Total.Text = extcosttotal_voy.ToString("###,###,###,###,###.##");

            tb_Gallons_Voy_Diesel.Text = galsdieselext_voy.ToString("###,###,###,###.##");
            tb_Gallons_Voy_Dyed.Text = galsdyedext_voy.ToString("###,###,###,###.##");
            tb_Gallons_Voy_MVF.Text = galsgasext_voy.ToString("###,###,###,###.##");
            tb_Gallons_Voy_Other.Text = galsotherext_voy.ToString("###,###,###.##");
            tb_Gallons_Voy_Total.Text = galstotalext_voy.ToString("###,###,###.##");

            tb_TP_Voy_Diesel.Text = salestaxpaiddiesel_voy.ToString("###,###,###,###,###.##");
            tb_TP_Voy_Dyed.Text = salestaxpaiddyed_voy.ToString("####,###,###,###,###.##");
            tb_TP_Voy_MVF.Text = salestaxpaidgas_voy.ToString("###,###,###,###,###.##");
            tb_TP_Voy_Other.Text = salestaxpaidother_voy.ToString("###,###,###,###,###.##");
            tb_TP_Voy_Total.Text = salestaxpaidtotal_voy.ToString("###,###,###,###,###.##");


            SaveFileDialog saveext = new SaveFileDialog();

            saveext.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";

            if (saveext.ShowDialog() == DialogResult.OK)
            {

                try
                {
                    ExcelPackage pkg = new ExcelPackage(new MemoryStream());


                    if (locbil_de_ca.Count > 0)
                    {
                        string wsname = "PPextended";


                        ExcelWorksheet ws = pkg.Workbook.Worksheets.Add(wsname);

                        ws.Cells["A1"].LoadFromCollection(locbil_de_ca, true);
                    }

                    if (locbil.Count > 0)
                    {
                        string wsname = "PPallrecords";


                        ExcelWorksheet wsp = pkg.Workbook.Worksheets.Add(wsname);

                        wsp.Cells["A1"].LoadFromCollection(locbil, true);
                    }

                    if (locbil_dd.Count > 0)
                    {
                        string wsname = "PPdomestic";
                        
                        ExcelWorksheet wsd = pkg.Workbook.Worksheets.Add(wsname);

                        wsd.Cells["A1"].LoadFromCollection(locbil_dd, true);
                    }

                    if (locbil_fdp.Count > 0)
                    {
                        string wsname = "PPwholesalesaleSales";

                        ExcelWorksheet wsp = pkg.Workbook.Worksheets.Add(wsname);

                        wsp.Cells["A1"].LoadFromCollection(locbil_fdp, true);
                    }

                    if (locbil_df.Count > 0)
                    {
                        string wsname = "PPwholesalePurchases";

                        ExcelWorksheet wsp = pkg.Workbook.Worksheets.Add(wsname);

                        wsp.Cells["A1"].LoadFromCollection(locbil_df, true);
                    }

                    if (VoyData.Count > 0)
                    {
                        string wsname = "Voyager";
                        
                        ExcelWorksheet ws0 = pkg.Workbook.Worksheets.Add(wsname);

                        ws0.Cells["A1"].LoadFromCollection(VoyData, true);
                    }



                    if (CFNDataExt.Count > 0)
                {
                    string wsname = "cfnextended";
                    
                    ExcelWorksheet ws1 = pkg.Workbook.Worksheets.Add(wsname);

                    ws1.Cells["A1"].LoadFromCollection(CFNDataExt, true);
                }
                                
                ExcelWorksheet ws2 = pkg.Workbook.Worksheets.Add("Summary");

                ws2.Cells["A1"].Value = "Cost of Tax Paid Purchases";

                ws2.Cells["A3"].Value = "Total of All Purchases";
                ws2.Cells["A4"].Value = "Clear Diesel Purchases";
                ws2.Cells["A5"].Value = "Dyed Diesel Purchases";
                ws2.Cells["A6"].Value = "MVF Purchases";
                ws2.Cells["A7"].Value = "Other Purchases";

                ws2.Cells["A11"].Value = "Taxes Paid";

                ws2.Cells["A13"].Value = "Total of All Taxes Paid";
                ws2.Cells["A14"].Value = "Clear Diesel Taxes Paid";
                ws2.Cells["A15"].Value = "Dyed Diesel Taxes Paid";
                ws2.Cells["A16"].Value = "MVF Taxes Paid";
                ws2.Cells["A17"].Value = "Other Taxes Paid";

                ws2.Cells["B1"].Value = "Pacific Pride";
                ws2.Cells["B3"].Value = extcosttotal_pp;
                ws2.Cells["B4"].Value = extcostdiesel_pp;
                ws2.Cells["B5"].Value = extcostdyed_pp;
                ws2.Cells["B6"].Value = extcostgas_pp;
                ws2.Cells["B7"].Value = extcostother_pp;

                ws2.Cells["B11"].Value = "Pacific Pride";
                ws2.Cells["B13"].Value = salestaxpaidtotal_pp;
                ws2.Cells["B14"].Value = salestaxpaiddiesel_pp;
                ws2.Cells["B15"].Value = salestaxpaiddyed_pp;
                ws2.Cells["B16"].Value = salestaxpaidgas_pp;
                ws2.Cells["B17"].Value = salestaxpaidother_pp;

                ws2.Cells["C1"].Value = "CFN";
                ws2.Cells["C3"].Value = extcosttotal_cfn;
                ws2.Cells["C4"].Value = extcostdiesel_cfn;
                ws2.Cells["C5"].Value = extcostdyed_cfn;
                ws2.Cells["C6"].Value = extcostgas_cfn;
                ws2.Cells["C7"].Value = extcostother_cfn;

                ws2.Cells["C11"].Value = "CFN";
                ws2.Cells["C13"].Value = salestaxpaidtotal_cfn;
                ws2.Cells["C14"].Value = salestaxpaiddiesel_cfn;
                ws2.Cells["C15"].Value = salestaxpaiddyed_cfn;
                ws2.Cells["C16"].Value = salestaxpaidgas_cfn;
                ws2.Cells["C17"].Value = salestaxpaidother_cfn;

                    ws2.Cells["D1"].Value = "Voyager";
                    ws2.Cells["D3"].Value = extcosttotal_voy;
                    ws2.Cells["D4"].Value = extcostdiesel_voy;
                    ws2.Cells["D5"].Value = extcostdyed_voy;
                    ws2.Cells["D6"].Value = extcostgas_voy;
                    ws2.Cells["D7"].Value = extcostother_voy;

                    ws2.Cells["D11"].Value = "Voyager";
                    ws2.Cells["D13"].Value = salestaxpaidtotal_voy;
                    ws2.Cells["D14"].Value = salestaxpaiddiesel_voy;
                    ws2.Cells["D15"].Value = salestaxpaiddyed_voy;
                    ws2.Cells["D16"].Value = salestaxpaidgas_voy;
                    ws2.Cells["D17"].Value = salestaxpaidother_voy;

                    ws2.Cells["E1"].Value = "WEX";
                    ws2.Cells["E3"].Value = extcosttotal_wex;
                    ws2.Cells["E4"].Value = extcostdiesel_wex;
                    ws2.Cells["E5"].Value = extcostdyed_wex;
                    ws2.Cells["E6"].Value = extcostgas_wex;
                    ws2.Cells["E7"].Value = extcostother_wex;

                    ws2.Cells["E11"].Value = "WEX";
                    ws2.Cells["E13"].Value = salestaxpaidtotal_wex;
                    ws2.Cells["E14"].Value = salestaxpaiddiesel_wex;
                    ws2.Cells["E15"].Value = salestaxpaiddyed_wex;
                    ws2.Cells["E16"].Value = salestaxpaidgas_wex;
                    ws2.Cells["E17"].Value = salestaxpaidother_wex;

                    ws2.Cells["F11"].Value = "TOTAL";
                ws2.Cells["F13"].Value = salestaxpaidtotal;
                ws2.Cells["F14"].Value = salestaxpaiddiesel;
                ws2.Cells["F15"].Value = salestaxpaiddyed;
                ws2.Cells["F16"].Value = salestaxpaidgas;
                ws2.Cells["F17"].Value = salestaxpaidother;
                    
                ws2.Cells["F1"].Value = "TOTAL";
                ws2.Cells["F3"].Value = extcosttotal;
                ws2.Cells["F4"].Value = extcostdiesel;
                ws2.Cells["F5"].Value = extcostdyed;
                ws2.Cells["F6"].Value = extcostgas;
                ws2.Cells["F7"].Value = extcostother;
                    
                    ws2.Cells["A3:D7"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                    ws2.Cells["A3:D7"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                    ws2.Cells["A3:D7"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                    ws2.Cells["A3:D7"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

                    ws2.Cells["A13:D17"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                    ws2.Cells["A13:D17"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                    ws2.Cells["A13:D17"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                    ws2.Cells["A13:D17"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

                    ws2.Cells["B3:G17"].Style.Numberformat.Format = "#,###.##";

                    ws2.Cells[ws2.Dimension.Address].AutoFitColumns();


                string wsdistrict = "District Extended";

                ExcelWorksheet ws3 = pkg.Workbook.Worksheets.Add(wsdistrict);

                ws3.Cells["A1"].LoadFromDataTable(cadtax, true);

                pkg.SaveAs(new FileInfo(saveext.FileName));
                   
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not Save "+saveext.FileName);
                }

            }
            //   MessageBox.Show(salestaxpaidtotal.ToString());
            decimal totdisttax = 0.00m;
            decimal totdistsales = 0.00m;
            decimal totaltaxwithdistrict = 0.00m;


            //foreach (DataRow row in sales.Rows)
            //{
            //    adjsls = 0.00m;
            //    taxamt = 0.00m;

            //    int citycol = 0;
            //    int qcol = 0;
            //    int slscol = 0;
            //    int taxcol = 0;
            //    int prdcol = 0;
            //    int setcol = 0;
            //    int stecol = 0;

             
            //    city = row[citycol].ToString();

            //    if (city.Contains("   CA9"))
            //    {
            //        int start = city.IndexOf("   CA9");

            //        start--;

            //        city = city.Substring(0, start).Trim();
            //    }

            //    cnty = Get_County_With_City(city);


            //    q = row[qcol].ToString();
            //    sls = row[slscol].ToString();
            //    tax = row[taxcol].ToString();
            //    prod = row[prdcol].ToString();
            //    sset = row[setcol].ToString();
            //    state = row[stecol].ToString();

            //    quan = 0;
            //    slsamt = 0;
            //    tamt = 0;
            //    set = 0;

            //    decimal.TryParse(q, out quan);
            //    decimal.TryParse(sls, out slsamt);
            //    decimal.TryParse(tax, out tamt);
            //    decimal.TryParse(sset, out set);


            //    if (tamt != 0 && state == "CA")
            //    {

            //        jrate = Get_Jurisdiction_Rate(city, cnty);

            //        if (city.Contains("Lompoc"))
            //        {

            //            //    MessageBox.Show(" State=" + state + " City= " + city + " County =" + cnty + " Sales=" + sls + " Tax=" + tax + " Product=" + prod + " SET=" + sset + " Quan=" + q + " Rate=" + jrate.ToString());
            //        }

            //        if (!Check_City_Has_STJ(city))
            //        {
            //            city = "";
            //        }



            //        totaltaxwithdistrict = totaltaxwithdistrict + tamt;

            //        prod = prod.ToUpper();

            //        if (prod.Contains("DF") || prod.Contains("#2") || prod.Contains("DIESEL") || prod.Contains("ULSD"))
            //        {
            //            fta = 167;
            //        }

            //        if (prod.Contains("G87") || prod.Contains("G89") || prod.Contains("G91") || prod.Contains("PREM") || prod.Contains("REG") || prod.Contains("UNL"))
            //        {
            //            fta = 65;
            //        }

            //        if (prod.Contains("DYED") || prod.Contains("DFD"))
            //        {
            //            fta = 227;
            //        }

            //        adjsls = 0;

            //        // MessageBox.Show(checkBox_Sales_Each.Checked.ToString());

                   
            //        adjsls = slsamt;
                    

                  
            //        set = set * quan;
                    

            //        if (fta == 167)
            //        {

            //            adjsls = adjsls - set;
            //            taxamt = jrate * adjsls;

            //            if (taxamt != 0)
            //            {
            //                if (set != 0)
            //                {
            //                    // MessageBox.Show("City ="+city+" County="+cnty+" Tax ="+taxamt.ToString()+" Sales ="+adjsls);
            //                }

            //                totdistsales = totdistsales + adjsls + set;
            //                Add_To_District_Sales_And_Taxes(city, cnty, adjsls + set, 0.00m, taxamt, 0.00m,0);

            //                totdisttax = totdisttax + taxamt;
            //            }

            //        }

            //        if (fta == 227)
            //        {
            //            taxamt = jrate * adjsls;

            //            if (taxamt != 0)
            //            {
            //                totdistsales = totdistsales + adjsls;
            //                Add_To_District_Sales_And_Taxes(city, cnty, adjsls, 0.00m, taxamt, 0.00m,0);

            //                totdisttax = totdisttax + taxamt;
            //            }
            //        }

            //        if (fta == 65)
            //        {
            //            taxamt = jrate * adjsls;

            //            if (taxamt != 0)
            //            {

            //                totdistsales = totdistsales + adjsls;
            //                Add_To_District_Sales_And_Taxes(city, cnty, adjsls, 0.00m, taxamt, 0.00m,0);

            //                totdisttax = totdisttax + taxamt;
            //                //MessageBox.Show("City =" + city + " County=" + cnty + " TAX =" + taxamt.ToString() + " Sales =" + adjsls);
            //            }

            //        }

            //        if (fta == 0)
            //        {
            //            taxamt = jrate * adjsls;

            //            if (taxamt != 0)
            //            {
            //                totdistsales = totdistsales + adjsls;
            //                Add_To_District_Sales_And_Taxes(city, cnty, adjsls, 0.00m, taxamt, 0.00m,0);

            //                totdisttax = totdisttax + taxamt;
            //            }
            //        }

            //    }


            //}

            //tb_District_Tax_Incurred.Text = totdisttax.ToString("###,###,###,###.##");
            //tb_District_Sales_Total.Text = totdistsales.ToString("###,###,###,###.##");
            //tb_Sales_Tax_With_District_Tax.Text = totaltaxwithdistrict.ToString("###,###,###,###.##");
        }

        


        string Get_Zone_STJ(string zn)
        {
            string stj = "";

            int izones = 0;
            int izn = 0;

            int.TryParse(zn, out izn);

            foreach(DataRow row in zones.Rows)
            {
                int.TryParse(row["authority"].ToString(), out izones);

                if (izones == izn)
                {
                    stj = row["stj"].ToString();
                }


            }
            
            return stj;
        }

        string NoSpace(string spacey)
        {
            string nospace = spacey.Replace(" ", "");
            return nospace;
        }

        void Load_CA_Jurisdictions()
        {
            string[] lines = File.ReadAllLines(Jurisdiction);

            foreach (string line in lines)
            {
                var row = juris.NewRow();

                string[] fields = line.Split(',');

                for (int i = 0; i < fields.Length; i++)
                {
                    row[i] = fields[i];
                }

                // Get rid of title row

                if (!fields[0].Contains("city"))
                {
                    juris.Rows.Add(row);
                }

            }

            dgv_CA_Jurisdictions.DataSource = juris;

        }

        void Fill_Jurisdictions_With_STJs()
        {
            string city = "";
            string cnty = "";
            string stj = "";

            string jcity = "";
            string jcnty = "";
            string rate = "";

            foreach (DataRow row in cadtax.Rows)
            {
                city = row[1].ToString();
                cnty = row[0].ToString();
                stj = row[2].ToString();
                rate = row[5].ToString();


                city = clean(NoSpace(city)).ToUpper();
                cnty = clean(NoSpace(cnty)).ToUpper();

                cnty = cnty.Replace("COUNTY", "");

                city = city.Replace("SO.", "SOUTH");

                city = city.Replace("MT.", "");

                city = city.Replace("MOUNT", "");

                if (city.Length < 1 && stj.Length > 2)
                {

                    foreach (DataRow jrow in juris.Rows)
                    {
                        jcity = jrow[0].ToString();
                        jcnty = jrow[3].ToString();

                        jcity = clean(NoSpace(jcity)).ToUpper();
                        jcnty = clean(NoSpace(jcnty)).ToUpper();

                        jcity = jcity.Replace("MOUNT", "");

                        jcity = jcity.Replace("MT.", "");

                        // MessageBox.Show(jcnty + " " + cnty);
                        if (jcnty == cnty)
                        {
                            // MessageBox.Show("Matched " + cnty);
                            jrow[5] = stj;
                            jrow[7] = rate;
                        }

                    }
                }

                if (city.Length > 1 && stj.Length > 2)
                {

                    foreach (DataRow jrow in juris.Rows)
                    {
                        jcity = jrow[0].ToString();
                        jcnty = jrow[3].ToString();

                        jcity = clean(NoSpace(jcity)).ToUpper();
                        jcnty = clean(NoSpace(jcnty)).ToUpper();

                        jcity = jcity.Replace("MOUNT", "");

                        jcity = jcity.Replace("MT.", "");
                                              

                        // MessageBox.Show(jcity + " " + city);
                        if (jcity == city)
                        {
                            // MessageBox.Show("Matched " + city);
                            jrow[4] = stj;
                            jrow[6] = rate;
                        }

                    }
                }


            }

            dgv_CA_Jurisdictions.DataSource = juris;

            Save_Juridictions();
        }

        int Get_Jurisdiction_STJ(string city, string county)
        {            
            int stjcity = 0;
            int stjcounty = 0;
            
            string citystj = "";
            string countystj = "";

            city = city.ToUpper();
            city = city.Replace("THE", "");
            city = city.Replace("CITY", "");
            city = city.Replace("OF", "");
            city = city.Replace("CITYOF", "");

            city = clean(NoSpace(city)).ToUpper();
            county = clean(NoSpace(county)).ToUpper();

            if (county.Contains("CITY"))
            {
                city = county.Replace("CITY", "");
                city = city.Replace("OF", "");

            }

            if (city.Length == 12)
            {
                city = city.Replace("SANLUISOBISP", "SANLUISOBISPO");
            }

            city = city.Replace("MOUNT", "");
            city = city.Replace("MT.", "");

            county = county.Replace("CO.", "");

            county = county.Replace("COUNTY", "");
            county = county.Replace("-CA", "");
            county = county.Replace("BENTO", "BENITO");

            string jcity = "";
            string jcnty = "";

            foreach (DataRow jrow in juris.Rows)
            {
                jcity = jrow[0].ToString();
                jcnty = jrow[3].ToString();

                jcity = clean(NoSpace(jcity)).ToUpper();
                jcnty = clean(NoSpace(jcnty)).ToUpper();

                jcity = jcity.Replace("MOUNT", "");

                // MessageBox.Show(jcnty + " " + cnty);
                if (jcnty == county)
                {
                   
                    countystj = jrow[5].ToString();
                }

                if (jcity == city)
                {
                    citystj = jrow[4].ToString();
                 
                }

            }

            //MessageBox.Show(city + " " + citystj + " " + county +" "+ countystj);

            int.TryParse(citystj, out stjcity);

            int.TryParse(countystj, out  stjcounty);
                        
            if (stjcity > 0)
            {
                return stjcity;
            }
            else
            {
                return stjcounty;
            }

         

        }

        decimal Get_Jurisdiction_Rate(string city, string county, DateTime enddate)
        {
            decimal rate = 0.00m;
            decimal cityrate = 0.00m;
            decimal cntyrate = 0.00m;

            string srate = "";
            city = city.ToUpper();
            city = city.Replace("THE", "");
            city = city.Replace("CITY", "");
            city = city.Replace("OF", "");
            city = city.Replace("CITYOF", "");
            city = city.Replace("\"", "");

            city = clean(NoSpace(city)).ToUpper();
            county = clean(NoSpace(county)).ToUpper();

            if (county.Contains("CITY"))
            {
                city = county.Replace("CITY", "");
                city = city.Replace("OF", "");

            }

            if (city.Length == 12)
            {
                city = city.Replace("SANLUISOBISP", "SANLUISOBISPO");
            }

            city = city.Replace("MOUNT", "");
            city = city.Replace("MT.", "");

            county = county.Replace("CO.", "");

            county = county.Replace("COUNTY", "");
            county = county.Replace("-CA", "");
            county = county.Replace("BENTO", "BENITO");
            county = county.Replace("\"", "");

            string jcity = "";
            string jcnty = "";
            string jstj = "";
            int istj = 0;

            foreach (DataRow jrow in workTable.Rows)
            {
                jcity = jrow["City"].ToString();
                jcnty = jrow["County"].ToString();
                jstj = jrow["STJ"].ToString();

                int.TryParse(jstj, out istj);

                jcity = clean(NoSpace(jcity)).ToUpper();
                jcnty = clean(NoSpace(jcnty)).ToUpper();
                jcnty = jcnty.Replace("COUNTY", "");

                jcity = jcity.Replace("MOUNT", "");

                DateTime effdate = DateTime.Now;
                DateTime expdate = DateTime.Now;

                if (jrow["Effective"].ToString().Length > 0)
                {
                    DateTime.TryParse(jrow["Effective"].ToString(), out effdate);
                }
                else
                {
                    effdate = dtp_Date_Beg.Value;
                }
                if (jrow["Expired"].ToString().Length > 0)
                {
                    DateTime.TryParse(jrow["Expired"].ToString(), out expdate);
                }
                else
                {
                    expdate = dtp_Date_End.Value;
                }

                // FIRST FIND COUNTY

                if (jcnty == county)
                {
                    if (checkBox_Debug.Checked)
                    {
                        MessageBox.Show(" jcnty = " + jcnty + " County = " + county);
                    }

                    if (checkBox_Debug.Checked)
                    {
                        MessageBox.Show(" Effective Date =" + effdate.ToShortDateString() + " Expired Date = " + expdate.ToShortDateString() + " Tran Date = " + enddate.ToShortDateString(), "Get_Jurisdiction_Rate");
                    }

                    if (effdate <= enddate && expdate >= enddate)
                    {
                        if (cntyrate < 1)
                        {
                            srate = jrow["Rate"].ToString();
                            decimal.TryParse(srate, out cntyrate);

                        }

                        // THEN FIND CITY

                      

                        if (jcity == city)
                        {
                            if (checkBox_Debug.Checked)
                            {
                                MessageBox.Show(" jcity = " + jcity + " City = " + city);
                            }


                            srate = jrow["Rate"].ToString();
                            decimal.TryParse(srate, out cityrate);

                        }
                    }
                }

            }

            if (checkBox_Debug.Checked)
            {
                MessageBox.Show("County Rate = " + cntyrate.ToString() + " City Rate = " + cityrate.ToString());
            }

            if (city.Length <1)
            {
                rate = cntyrate;
            }
            else
            {
                rate = cityrate;
            }

                      
            return rate;
        }

        void Save_Juridictions()
        {
            dgv_CA_Jurisdictions.AllowUserToAddRows = false;

            List<string> jurs = new List<string>();

            foreach (DataGridViewRow row in dgv_CA_Jurisdictions.Rows)
            {
                string fields = "";

                for (int col = 0; col < dgv_CA_Jurisdictions.Columns.Count; col++)
                {

                    fields = fields + row.Cells[col].Value.ToString();
                    if (col < (dgv_CA_Jurisdictions.Columns.Count - 1))
                    {
                        fields = fields + ",";
                    }
                }
                jurs.Add(fields);
            }

            File.WriteAllLines(Jurisdiction, jurs);
        }


        int Lookup_CFNProd_FTA(string prod)
        {
            int fta = 0;

           // MessageBox.Show(cfn_products.Rows.Count.ToString());

            foreach (Products_CFN row in cfn_products)
            {
                string p = row.Code;

                p = p.Trim();
                prod = prod.Trim();

                if (p == prod)
                {
                    int.TryParse(row.FTA_Code.ToString(), out fta);
                }
            }

           // MessageBox.Show("Product =" + prod + " fta=" + fta.ToString());

            return fta;
        }

        int Lookup_Prod_FTA(string prod)
        {
            int fta = 0;


            foreach (Products_PP row in pp_products)
            {
                string p = row.Code.ToString();

                p = p.Trim();
                prod = prod.Trim();

                if (p.Equals(prod))
                {
                    int.TryParse(row.FTA_Code.ToString(), out fta);
                }
            }

            //MessageBox.Show("Product =" + prod + " fta=" + fta.ToString());

            return fta;
        }


        bool Check_City_Has_STJ(string city)
        {
            bool hasstj = false;

            city = clean(NoSpace(city)).ToUpper();
            city = city.Replace("\"", "");

            foreach (DataRow row in workTable.Rows)
            {
                string jcity = clean(NoSpace(row["City"].ToString())).ToUpper();
                
                if (jcity == city)
                {

                    string stj = clean(NoSpace(row["STJ"].ToString()));

                    if (stj.Length > 2)
                    {
                        hasstj = true;
                    }

                }

            }


            return hasstj;
        }
        

        void Load_Extended_PP()
        {
            notifyIcon1.Icon = SystemIcons.Hand;
            notifyIcon1.BalloonTipText = "Loading Extended";
            notifyIcon1.BalloonTipTitle = "Read Locbil";
            notifyIcon1.ShowBalloonTip(3000);
            
            if (locbil_de.Count > 1)
            {
                locbil_de.Clear();
            }

            foreach(Locbill row in locbil)
            {
                if (row.SellingPart.ToString().Equals(row.Host.ToString()))
                {
                    locbil_dd.Add(row);
                    Host = row.SellingPart;
                }
            }

            label_Host.Text = "Host Number ="+Host;

            foreach (Locbill row in locbil)
            {
                string host = row.SellingPart.ToString();

                int ihost = 0;
                int spart = 0;

                decimal ptax = 0.00m;
                decimal stax = 0.00m;
                decimal clsales = 0.00m;


                decimal.TryParse(row.SalesTaxAdjust.ToString(), out ptax);
                decimal.TryParse(row.SalesTaxAdjust, out stax);
                decimal.TryParse(row.SellingPrice, out clsales);

                int.TryParse(host, out ihost);
                int.TryParse(row.SellingPart, out spart);
                
                string state = row.Site_State.ToString();

                if (row.Host.ToString().Equals(Host) && !row.SellingPart.ToString().Equals(Host) && ihost < 900)
                {
                    locbil_df.Add(row);
                }
                             
                if (!row.Host.ToString().Equals(Host) && row.SellingPart.ToString().Equals(Host) && clsales == 0.00m && stax == 0)
                {
                    locbil_fdp.Add(row);
                }

                if (!row.Host.ToString().Equals(Host) && row.SellingPart.ToString().Equals(Host) && clsales > 0.00m && stax > 0)
                {
                    locbil_fdo.Add(row);
                }


                if (ihost > 900 )
                {
                //    MessageBox.Show("Extended Trans =" + row.Trans);

                    locbil_de.Add(row);

                    if (state.Contains("CA"))
                    {                        
                        locbil_de_ca.Add(row);
                    }
                }
            }

            label_Extended_Records.Text = "Records ="+locbil_de.Count.ToString("###,###,###");
            label_PP_Extended_CA.Text = "Records ="+locbil_de_ca.Count.ToString("###,###,###");

            dgv_PP_DD.DataSource = locbil_dd;
            dgv_PP_DF.DataSource = locbil_df;

            label_PP_DD_Records.Text = "Records =" + locbil_dd.Count.ToString("###,###,###");
            label_PP_DF_Records.Text = "Records =" + locbil_df.Count.ToString("###,###,###");

            dgv_PP_Extended.DataSource = locbil_de;
            dgv_PP_Extended_CA_Records.DataSource = locbil_de_ca;

            dgv_PP_FDP.DataSource = locbil_fdp;
            dgv_PP_FDO.DataSource = locbil_fdo;

            label_PP_FDO.Text = "Records =" + locbil_fdo.Count.ToString("###,###,###");
            label_PP_FDP.Text = "Records =" + locbil_fdp.Count.ToString("###,###,###");
        }

        private void button_Save_Setup_Click(object sender, EventArgs e)
        {
            Save_All_Data();
        }

        private void button_Restore_Setup_Click(object sender, EventArgs e)
        {
            Restore_All_Data();
        }

        private void button_Save_Sales_File_Setup_Click(object sender, EventArgs e)
        {
            Load_Textboxes_Calc();
            Save_Calc_Settings();
        }

        private void button_Get_CFN_Files_Click(object sender, EventArgs e)
        {
            Load_CFN_CSV_File();
            //Load_single_file();
            //Convert_Dir();
        }

        private void button_Calc_CPP_District_Click(object sender, EventArgs e)
        {
            Calc_District_taxes();
        }

        private void dtp_Date_Beg_ValueChanged(object sender, EventArgs e)
        {
            dtp_Date_End.Value = dtp_Date_Beg.Value.AddMonths(3).AddDays(-1);
        }

        private void label_Number_Of_Trans_Locbil_Click(object sender, EventArgs e)
        {

        }

        string Get_host_host_pp()
        {
            string host_host = "";


            return host_host;

        }

        string Get_Trans_Cat_PP(string site_host, string card_host, decimal pretax_rate, decimal salestax, string site_state, string host_host)
        {
            string trans_cat = "DW";

            return trans_cat;
        }

        private void button_Convert_XLSX_CSV_Click(object sender, EventArgs e)
        {
            Convert_Dir();
        }

        private void tb_Deductions_Cost_Tax_Paid_Diesel_Resold_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_Deductions_Cost_Of_Tax_Paid_MVF_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_Deductions_Cost_Of_Tax_Paid_Purchases_Other_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void label92_Click(object sender, EventArgs e)
        {

        }

        private void tb_TP_PP_Diesel_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_TP_CFN_MVF_TextChanged(object sender, EventArgs e)
        {

        }

        private void button_Get_Voyager_Files_Click(object sender, EventArgs e)
        {
            Load_Voyager_File();
        }

        private void tb_TP_Voy_Diesel_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox12_Enter(object sender, EventArgs e)
        {

        }

        private void tb_CTPP_Voy_Diesel_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_STP_Voy_Diesel_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_DTP_Voy_Diesel_TextChanged(object sender, EventArgs e)
        {

        }

        private void label97_Click(object sender, EventArgs e)
        {

        }

        private void tb_CTPP_Voy_Dyed_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_TP_Voy_Dyed_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_STP_Voy_Dyed_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_DTP_Voy_Dyed_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_CTPP_Voy_MVF_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_TP_Voy_MVF_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_STP_Voy_MVF_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_DTP_Voy_MVF_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_CTPP_Voy_Other_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_TP_Voy_Other_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_STP_Voy_Other_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_DTP_Voy_Other_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_CTPP_Voy_Total_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_TP_Voy_Total_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_STP_Voy_Total_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_DTP_Voy_Total_TextChanged(object sender, EventArgs e)
        {

        }

        private void label100_Click(object sender, EventArgs e)
        {

        }

        private void label101_Click(object sender, EventArgs e)
        {

        }

        private void label99_Click(object sender, EventArgs e)
        {

        }

        private void label98_Click(object sender, EventArgs e)
        {

        }

        private void label103_Click(object sender, EventArgs e)
        {

        }

        private void label93_Click(object sender, EventArgs e)
        {

        }

        private void groupBox9_Enter(object sender, EventArgs e)
        {

        }

        private void tb_Gallons_CFN_MVF_TextChanged(object sender, EventArgs e)
        {

        }

        private void toolsToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void pacPrideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Load_LocBill_File();
        }

        private void cFNToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Load_CFN_CSV_File();
        }

        private void voyagerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Load_Voyager_File();
        }

        private void convertXLSXToCSVEntireDirectoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Convert_Dir();
        }

        private void dgv_CFN_Counties_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void lv_CFN_Data_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void tabPage10_Click(object sender, EventArgs e)
        {

        }

        private void dgv_Locbil_Data_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        List<string[]> Get_File_Delimited(string file, string del)
        {
            if(del.Length < 1)
            {
                del = Analyze_File(file);
            }

                        
            string[] lines = File.ReadAllLines(file);

            List<string[]> filelines = new List<string[]>();
           
            foreach(string line in lines)
            {
                

                string[] linefields = Get_Fields_Delimited(line, del);
                filelines.Add(linefields);
            }

            return filelines;
        }

        string[] Get_Fields_Delimited(string line, string del)
        {
            List<string> flds = new List<string>();
            if (del.Length > 0)
            {
            
                string addstr = "";

                int strlen = del.Length;
                string cmp = "";

                for (int i = 0; i < line.Length - strlen; i++)
                {
                    addstr = addstr + line.Substring(i, 1);

                    cmp = line.Substring(i, strlen);
                    if (cmp.Equals(del))
                    {
                        flds.Add(addstr.Replace(del, "").Replace(",",""));
                        addstr = "";
                    }
                }


                return flds.ToArray();
            }
            else
            {
                MessageBox.Show("Must Have Delimiter "+del);
                MessageBox.Show(line);
                return flds.ToArray();
            }
        }

        int StringCount(string line, string str)
        {
            int cnt = 0;

            int strlen = str.Length;
            string cmp = "";

            for(int i = 0; i < line.Length-strlen; i++)
            {
                cmp = line.Substring(i, strlen);
                if (cmp.Equals(str))
                {
                    cnt++;
                }
            }

            return cnt;
        }

        int CharCount(string line, char ch)
        {
            int cnt = 0;

            foreach(char c in line)
            {
                if(c == ch)
                {
                    cnt++;
                }
            }


            return cnt;
        }

        string Analyze_Line(string[] lines)
        {
            bool status = false;

            fields fs = new fields();

            fs.Commas = 0;
            fs.Quotes = 0;
            fs.Combos = 0;
            fs.Maximum_Commas = 0;
            fs.Minimum_Commas = 1000;
            fs.Maximum_Quotes = 0;
            fs.Minimum_Quotes = 1000;
            fs.Maximum_Combos = 0;
            fs.Minimum_Combos = 1000;
                      

            if (lines.Length > 0)
            {
                status = true;
               // MessageBox.Show("Number of Lines in Sales File = "+lines.Length.ToString());
            }
            else
            {
                status = false;
                MessageBox.Show("Number of Lines in File = " + lines.Length.ToString());
                return "No Records";

            }

            if (status)
            {
                int lastcomma = 0;
                int lastquote = 0;
                int lastcombo = 0;

                int cng_commas = 0;
                int cng_quotes = 0;
                int cng_combos = 0;

                foreach (string line in lines)
                {

                    fs.Combos = StringCount(line, "\",\"");

                    // MessageBox.Show("Qoutes =" + fs.Quotes.ToString());

                    fs.Commas = CharCount(line, ',');

                    fs.Quotes = CharCount(line, '\"');

                    if (fs.Commas > fs.Maximum_Commas)
                    {
                        fs.Maximum_Commas = fs.Commas;
                    }
                    if (fs.Commas < fs.Minimum_Commas)
                    {
                        fs.Minimum_Commas = fs.Commas;
                    }

                    if (lastcomma != fs.Commas)
                    {
                        fs.Different_Commas.Add(fs.Commas);
                        lastcomma = fs.Commas;

                    }

                    if (fs.Quotes > fs.Maximum_Quotes)
                    {
                        fs.Maximum_Quotes = fs.Quotes;
                    }
                    if (fs.Quotes < fs.Minimum_Quotes)
                    {
                        fs.Minimum_Quotes = fs.Quotes;
                    }


                    if (lastquote != fs.Quotes)
                    {
                        lastquote = fs.Quotes;
                        fs.Different_Quotes.Add(fs.Quotes);
                    }

                    if (fs.Combos > fs.Maximum_Combos)
                    {
                        fs.Maximum_Combos = fs.Combos;

                    }
                    if (fs.Combos < fs.Minimum_Combos)
                    {
                        fs.Minimum_Combos = fs.Combos;
                    }


                    if (lastcombo != fs.Combos)
                    {
                        lastcombo = fs.Combos;
                        fs.Different_Combos.Add(fs.Combos);
                    }

                }

                //   MessageBox.Show(" Commas : Max= " + fs.Maximum_Commas.ToString() + " Changes = " + fs.Different_Commas.Count.ToString());
                //  MessageBox.Show(" Quotes : Max= " + fs.Maximum_Quotes.ToString() + " Changes = " + fs.Different_Quotes.Count.ToString());
                //  MessageBox.Show(" Combos : Max= " + fs.Maximum_Combos.ToString() + " Changes = " + fs.Different_Combos.Count.ToString());

                if (fs.Maximum_Commas < 1 && fs.Maximum_Quotes < 1 && fs.Minimum_Combos < 1)
                {
                    return "";
                }


                cng_commas = fs.Different_Commas.Count;
                cng_quotes = fs.Different_Quotes.Count;
                cng_combos = fs.Different_Combos.Count;

                if (cng_commas < cng_quotes && cng_commas < cng_combos)
                {
                    return ",";
                }

                if (cng_quotes < cng_commas && cng_quotes < cng_combos)
                {
                    return "\"";
                }

                if (cng_combos < cng_commas && cng_combos < cng_quotes)
                {
                    return "\",\"";
                }

            }
            return "";

        }

        string Analyze_File(string file)
        {
            bool status = false;

            fields fs = new fields();

            fs.Commas = 0;
            fs.Quotes = 0;
            fs.Combos = 0;
            fs.Maximum_Commas = 0;
            fs.Minimum_Commas = 1000;
            fs.Maximum_Quotes = 0;
            fs.Minimum_Quotes = 1000;
            fs.Maximum_Combos = 0;
            fs.Minimum_Combos = 1000;

            string[] lines = File.ReadAllLines(file);
            
            if(lines.Length > 0)
            {
                status = true;
               // MessageBox.Show("Number of Lines in Sales File = "+lines.Length.ToString());
            }
            else
            {
                status = false;
            //    MessageBox.Show("Number of Lines in Sales File = " + lines.Length.ToString());
                return "No Records";

            }

            if (status)
            {
                int lastcomma = 0;
                int lastquote = 0;
                int lastcombo = 0;

                int cng_commas = 0;
                int cng_quotes = 0;
                int cng_combos = 0;

                foreach (string line in lines)
                {

                    fs.Combos = StringCount(line, "\",\"");

                    // MessageBox.Show("Qoutes =" + fs.Quotes.ToString());

                    fs.Commas = CharCount(line, ',');

                    fs.Quotes = CharCount(line, '\"');

                    if (fs.Commas > fs.Maximum_Commas)
                    {
                        fs.Maximum_Commas = fs.Commas;
                    }
                    if (fs.Commas < fs.Minimum_Commas)
                    {
                        fs.Minimum_Commas = fs.Commas;
                    }

                    if (lastcomma != fs.Commas)
                    {
                        fs.Different_Commas.Add(fs.Commas);
                        lastcomma = fs.Commas;

                    }

                    if (fs.Quotes > fs.Maximum_Quotes)
                    {
                        fs.Maximum_Quotes = fs.Quotes;
                    }
                    if (fs.Quotes < fs.Minimum_Quotes)
                    {
                        fs.Minimum_Quotes = fs.Quotes;
                    }


                    if (lastquote != fs.Quotes)
                    {
                        lastquote = fs.Quotes;
                        fs.Different_Quotes.Add(fs.Quotes);
                    }

                    if (fs.Combos > fs.Maximum_Combos)
                    {
                        fs.Maximum_Combos = fs.Combos;

                    }
                    if (fs.Combos < fs.Minimum_Combos)
                    {
                        fs.Minimum_Combos = fs.Combos;
                    }


                    if (lastcombo != fs.Combos)
                    {
                        lastcombo = fs.Combos;
                        fs.Different_Combos.Add(fs.Combos);
                    }

                }

                //   MessageBox.Show(" Commas : Max= " + fs.Maximum_Commas.ToString() + " Changes = " + fs.Different_Commas.Count.ToString());
                //  MessageBox.Show(" Quotes : Max= " + fs.Maximum_Quotes.ToString() + " Changes = " + fs.Different_Quotes.Count.ToString());
                //  MessageBox.Show(" Combos : Max= " + fs.Maximum_Combos.ToString() + " Changes = " + fs.Different_Combos.Count.ToString());

                if(fs.Maximum_Commas < 1 && fs.Maximum_Quotes <1 && fs.Minimum_Combos < 1)
                {
                    return "";
                }
                    

                cng_commas = fs.Different_Commas.Count;
                cng_quotes = fs.Different_Quotes.Count;
                cng_combos = fs.Different_Combos.Count;
                
                if(cng_commas<cng_quotes && cng_commas < cng_combos)
                {
                    return ",";
                }

                if (cng_quotes < cng_commas && cng_quotes < cng_combos)
                {
                    return "\"";
                }

                if(cng_combos < cng_commas && cng_combos < cng_quotes)
                {
                    return "\",\"";
                }

            }
            return "";
        }

    
        private void analyzeSalesFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string status = Analyze_File(tb_Sales_File.Text);
            MessageBox.Show(status.ToString());
        }

        private void dgv_Sales_Data_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tb_Exemptions_Farm_Equip_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage8_Click(object sender, EventArgs e)
        {

        }

        private void phoneticValueCitysToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Get_phvalue_All_Citys();
        }
    }
}
