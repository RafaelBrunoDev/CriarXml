using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CriarXml
{
    public partial class Layout : Form
    {
        public Layout()
        {
            InitializeComponent();

        }

        public class Utf8StreamWriter : StreamWriter
        {
            public Utf8StreamWriter(string path) : base(path, false, Encoding.UTF8)
            {
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            #region DEFINIÇÃO DOS VALORES

            #region Fixos

            string _1_Fixo_XMLVersion = "";
            string _2_Fixo_AdiXMLVersion = "";
            string _3_Fixo_MetaData = "";
            string _4_Fixo_AMS_Provider = "";
            string _5_Fixo_Product = "";
            string _6_Fixo_Version_Major = "";
            string _7_Fixo_Version_Minor = "";
            string _8_Fixo_Provider_Id = "";
            string _9_Fixo_Asset_Class = "";
            string _10__Fixo_App_Data_Metadata_Spec_Version = "";
            string _11_Fixo_MetaData = "";
            string _12_Fixo_Asset = "";
            string _13_Fixo_MetaData = "";
            string _14_Fixo_AMS_Provider = "";
            string _15_Fixo_Product = "";
            string _16_Fixo_Version_Major = "";
            string _17_Fixo_Version_Minor = "";
            string _18_Fixo_Provider_Id = "";
            string _19_Fixo_Asset_Class = "";
            string _20_Fixo_App_Data_Type_Title = "";
            string _21_Fixo_App_Data_Box_Office = "";
            string _22_Fixo_App_Data_Preview_Period = "";
            string _23_Fixo_MetaData = "";
            string _24_Fixo_Asset = "";
            string _25_Fixo_MetaData = "";
            string _26_Fixo_AMS_Provider = "";
            string _27_Fixo_Product = "";
            string _28_Fixo_Version_Major = "";
            string _29_Fixo_Version_Minor = "";
            string _30_Fixo_Provider_Id = "";
            string _31_Fixo_Asset_Class = "";
            string _32_Fixo_App_Data_Type = "";
            string _33_Fixo_App_Data_Screen_Format = "";
            string _34_Fixo_App_Data_HDContent = "";
            string _35_Fixo_App_Data_Viewing_Can_Be_Resumed = "";
            string _36_Fixo_App_Data_Watermarking = "";
            string _37_Fixo_App_Data_Copy_Protection = "";
            string _38_Fixo_MetaData = "";
            string _39_Fixo_Asset = "";
            string _40_Fixo_Asset = "";
            string _41_Fixo_MetaData = "";
            string _42_Fixo_AMS_Provider = "";
            string _43_Fixo_Product = "";
            string _44_Fixo_Version_Major = "";
            string _45_Fixo_Version_Minor = "";
            string _46_Fixo_Provider_Id = "";
            string _47_Fixo_Asset_Class = "";
            string _48_Fixo_App_Data_Type = "";
            string _49_Fixo_MetaData = "";
            string _50_Fixo_Asset = "";
            string _51_Fixo_Asset = "";
            string _52_Fixo_ADI = "";

            #endregion


            #region Variaveis

            string _1_Asset_Name = "";
            string _2_Description = "";
            string _3_Creation_Date = "";
            string _4_Asset_Id = "";
            string _5_Asset_Name = "";
            string _6_Description = "";
            string _7_Creation_Date = "";
            string _8_Asset_Id = "";
            string _9_App_Data_Title_Sort_Name = "";
            string _10_App_Data_Title_Brief = "";
            string _11_App_Data_Title = "";
            string _12_App_Data_Episode_Name = "";
            string _13_App_Data_Episode_ID = "";
            string _14_App_Data_Summary_Long = "";
            string _15_App_Data_Summary_Medium = "";
            string _16_App_Data_Summary_Short = "";
            string _17_App_Data_Rating = "";
            string _18_App_Data_Closed_Captioning = "";
            string _19_App_Data_Run_Time = "";
            string _20_App_Data_Display_Run_Time = "";
            string _21_App_Data_Year = "";
            string _22_App_Data_Country_of_Origin = "";
            string _23_App_Data_Studio_Name = "";
            string _24_App_Data_Actors = "";
            string _25_App_Data_Actors_Display = "";
            string _26_App_Data_Director = "";
            string _27_App_Data_Director_Display = "";
            string _28_App_Data_Category = "";
            string _29_App_Data_Genre = "";
            string _30_App_Data_Genre = "";
            string _31_App_Data_Billing_ID = "";
            string _32_App_Data_Licensing_Window_Start = "";
            string _33_App_Data_Licensing_Window_End = "";
            string _34_App_Data_Suggested_Price = "";
            string _35_Asset_Name = "";
            string _36_Description = "";
            string _37_Creation_Date = "";
            string _38_Asset_Id = "";
            string _39_App_Data_Audio_Type = "";
            string _40_App_Data_Languages = "";
            string _41_App_Data_Content_FileSize = "";
            string _42_App_Data_Content_CheckSum = "";
            string _43_Content = "";
            string _44_Asset_Name = "";
            string _45_Description = "";
            string _46_Creation_Date = "";
            string _47_Asset_Id = "";
            string _48_App_Data_Content_FileSize = "";
            string _49_App_Data_Content_CheckSum = "";
            string _50_Content = "";

            #endregion

            #endregion

            #region CONEXÃO COM O BANCO
            System.Data.OleDb.OleDbConnection db = new
            System.Data.OleDb.OleDbConnection();
            // TODO: Modify the connection string and include any
            // additional required properties for your database.
            //db.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;" + @"Data Source= C:\Users\TestXml\Database2.mdb";
            //Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:

            //try
            //{
            //db.Open();
            // Insert code to process data.
            //}
            /*catch (Exception ex)
            {
                MessageBox.Show("FALHA DE CONEXÃO COM O BANCO DE DADOS !!!!!!!");
            }
            finally
            {
                db.Close();
            }*/
            #endregion

            #region CONEXÃO COM O EXCEL E CRIAÇÃO DO ARQUIVO TXT
            OleDbConnection conexao = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + txtArquivo.Text + " ; Extended Properties ='Excel 12.0 Xml; HDR = YES';");
            OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [Planilha1$]", conexao);

            OleDbConnection conexaoFixos = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + txtArquivo.Text + " ; Extended Properties ='Excel 12.0 Xml; HDR = YES';");
            OleDbDataAdapter adapterFixos = new OleDbDataAdapter("SELECT * FROM [Fixos$]", conexaoFixos);

            DataSet ds = new DataSet();
            DataSet dsFixos = new DataSet();
            conexao.Open();
            conexaoFixos.Open();
            adapter.Fill(ds);
            adapterFixos.Fill(dsFixos);



            #region valores fixos GET

            foreach (DataRow linhaFixo in dsFixos.Tables[0].Rows)
            {

                try
                {
                    _1_Fixo_XMLVersion = linhaFixo["_1_Fixo_XMLVersion"].ToString();
                    _2_Fixo_AdiXMLVersion = linhaFixo["_2_Fixo_AdiXMLVersion"].ToString();
                    _3_Fixo_MetaData = linhaFixo["_3_Fixo_MetaData"].ToString();
                    _4_Fixo_AMS_Provider = linhaFixo["_4_Fixo_AMS_Provider"].ToString();
                    _5_Fixo_Product = linhaFixo["_5_Fixo_Product"].ToString();
                    _6_Fixo_Version_Major = linhaFixo["_6_Fixo_Version_Major"].ToString();
                    _7_Fixo_Version_Minor = linhaFixo["_7_Fixo_Version_Minor"].ToString();
                    _8_Fixo_Provider_Id = linhaFixo["_8_Fixo_Provider_Id"].ToString();
                    _9_Fixo_Asset_Class = linhaFixo["_9_Fixo_Asset_Class"].ToString();
                    _10__Fixo_App_Data_Metadata_Spec_Version = linhaFixo["_10__Fixo_App_Data_Metadata_Spec_Version"].ToString();
                    _11_Fixo_MetaData = linhaFixo["_11_Fixo_MetaData"].ToString();
                    _12_Fixo_Asset = linhaFixo["_12_Fixo_Asset"].ToString();
                    _13_Fixo_MetaData = linhaFixo["_13_Fixo_MetaData"].ToString();
                    _14_Fixo_AMS_Provider = linhaFixo["_14_Fixo_AMS_Provider"].ToString();
                    _15_Fixo_Product = linhaFixo["_15_Fixo_Product"].ToString();
                    _16_Fixo_Version_Major = linhaFixo["_16_Fixo_Version_Major"].ToString();
                    _17_Fixo_Version_Minor = linhaFixo["_17_Fixo_Version_Minor"].ToString();
                    _18_Fixo_Provider_Id = linhaFixo["_18_Fixo_Provider_Id"].ToString();
                    _19_Fixo_Asset_Class = linhaFixo["_19_Fixo_Asset_Class"].ToString();
                    _20_Fixo_App_Data_Type_Title = linhaFixo["_20_Fixo_App_Data_Type_Title"].ToString();
                    _21_Fixo_App_Data_Box_Office = linhaFixo["_21_Fixo_App_Data_Box_Office"].ToString();
                    _22_Fixo_App_Data_Preview_Period = linhaFixo["_22_Fixo_App_Data_Preview_Period"].ToString();
                    _23_Fixo_MetaData = linhaFixo["_23_Fixo_MetaData"].ToString();
                    _24_Fixo_Asset = linhaFixo["_24_Fixo_Asset"].ToString();
                    _25_Fixo_MetaData = linhaFixo["_25_Fixo_MetaData"].ToString();
                    _26_Fixo_AMS_Provider = linhaFixo["_26_Fixo_AMS_Provider"].ToString();
                    _27_Fixo_Product = linhaFixo["_27_Fixo_Product"].ToString();
                    _28_Fixo_Version_Major = linhaFixo["_28_Fixo_Version_Major"].ToString();
                    _29_Fixo_Version_Minor = linhaFixo["_29_Fixo_Version_Minor"].ToString();
                    _30_Fixo_Provider_Id = linhaFixo["_30_Fixo_Provider_Id"].ToString();
                    _31_Fixo_Asset_Class = linhaFixo["_31_Fixo_Asset_Class"].ToString();
                    _32_Fixo_App_Data_Type = linhaFixo["_32_Fixo_App_Data_Type"].ToString();
                    _33_Fixo_App_Data_Screen_Format = linhaFixo["_33_Fixo_App_Data_Screen_Format"].ToString();
                    _34_Fixo_App_Data_HDContent = linhaFixo["_34_Fixo_App_Data_HDContent"].ToString();
                    _35_Fixo_App_Data_Viewing_Can_Be_Resumed = linhaFixo["_35_Fixo_App_Data_Viewing_Can_Be_Resumed"].ToString();
                    _36_Fixo_App_Data_Watermarking = linhaFixo["_36_Fixo_App_Data_Watermarking"].ToString();
                    _37_Fixo_App_Data_Copy_Protection = linhaFixo["_37_Fixo_App_Data_Copy_Protection"].ToString();
                    _38_Fixo_MetaData = linhaFixo["_38_Fixo_MetaData"].ToString();
                    _39_Fixo_Asset = linhaFixo["_39_Fixo_Asset"].ToString();
                    _40_Fixo_Asset = linhaFixo["_40_Fixo_Asset"].ToString();
                    _41_Fixo_MetaData = linhaFixo["_41_Fixo_MetaData"].ToString();
                    _42_Fixo_AMS_Provider = linhaFixo["_42_Fixo_AMS_Provider"].ToString();
                    _43_Fixo_Product = linhaFixo["_43_Fixo_Product"].ToString();
                    _44_Fixo_Version_Major = linhaFixo["_44_Fixo_Version_Major"].ToString();
                    _45_Fixo_Version_Minor = linhaFixo["_45_Fixo_Version_Minor"].ToString();
                    _46_Fixo_Provider_Id = linhaFixo["_46_Fixo_Provider_Id"].ToString();
                    _47_Fixo_Asset_Class = linhaFixo["_47_Fixo_Asset_Class"].ToString();
                    _48_Fixo_App_Data_Type = linhaFixo["_48_Fixo_App_Data_Type"].ToString();
                    _49_Fixo_MetaData = linhaFixo["_49_Fixo_MetaData"].ToString();
                    _50_Fixo_Asset = linhaFixo["_50_Fixo_Asset"].ToString();
                    _51_Fixo_Asset = linhaFixo["_51_Fixo_Asset"].ToString();
                    _52_Fixo_ADI = linhaFixo["_52_Fixo_ADI"].ToString();

                }
                catch //(Exception ex)
                {
                    //MessageBox.Show("" + ex.Message);
                }
                finally
                {

                }


            }

            #endregion

            string xml_out = "";
            int i = 1;
            string nome_arquivo = "";
            // SALVAR NA MESMA PASTA DO .EXE
            string pasta = Assembly.GetExecutingAssembly().Location;
            pasta = DateTime.Now.ToString("yyyyMddHHmmss");// 2022-10-13-1518

            foreach (DataRow linha in ds.Tables[0].Rows)
            {
                i++;

                try
                {
                    xml_out = "";
                    xml_out += _1_Fixo_XMLVersion + "\n";

                    xml_out += _2_Fixo_AdiXMLVersion + "\n";

                    xml_out += _3_Fixo_MetaData + "\n";

                    xml_out += _4_Fixo_AMS_Provider + "\n";

                    xml_out += _5_Fixo_Product + "\n";

                    _1_Asset_Name = "Asset_Name=\"" + linha["_1_Asset_Name"].ToString() + "\"" + "\n";
                    xml_out += _1_Asset_Name;

                    xml_out += _6_Fixo_Version_Major + "\n";

                    xml_out += _7_Fixo_Version_Minor + "\n";

                    _2_Description = "Description=\"" + linha["_2_Description"].ToString() + "\"" + "\n";
                    xml_out += _2_Description;
                    _3_Creation_Date = "Creation_Date=\"" + linha["_3_Creation_Date"].ToString() + "\"" + "\n";
                    xml_out += _3_Creation_Date;

                    xml_out += _8_Fixo_Provider_Id + "\n";

                    _4_Asset_Id = "Asset_ID=\"" + linha["_4_Asset_Id"].ToString() + "\"" + "\n";
                    xml_out += _4_Asset_Id;

                    xml_out += _9_Fixo_Asset_Class + "\n";

                    xml_out += _10__Fixo_App_Data_Metadata_Spec_Version + "\n";

                    xml_out += _11_Fixo_MetaData + "\n";

                    xml_out += _12_Fixo_Asset + "\n";

                    xml_out += _13_Fixo_MetaData + "\n";

                    xml_out += _14_Fixo_AMS_Provider + "\n";

                    xml_out += _15_Fixo_Product + "\n";

                    _5_Asset_Name = "Asset_Name=\"" + linha["_5_Asset_Name"].ToString() + "\"" + "\n";
                    xml_out += _5_Asset_Name;

                    xml_out += _16_Fixo_Version_Major + "\n";

                    xml_out += _17_Fixo_Version_Minor + "\n";

                    _6_Description = "Description=\"" + linha["_6_Description"].ToString() + "\"" + "\n";
                    xml_out += _6_Description;
                    _7_Creation_Date = "Creation_Date=\"" + linha["_7_Creation_Date"].ToString() + "\"" + "\n";
                    xml_out += _7_Creation_Date;

                    xml_out += _18_Fixo_Provider_Id + "\n";

                    _8_Asset_Id = "Asset_ID=\"" + linha["_8_Asset_Id"].ToString() + "\"" + "\n";
                    xml_out += _8_Asset_Id;

                    xml_out += _19_Fixo_Asset_Class + "\n";

                    xml_out += _20_Fixo_App_Data_Type_Title + "\n";

                    _9_App_Data_Title_Sort_Name = "<App_Data App=\"MOD\" Name=\"Title_Sort_Name\" Value=\"" + linha["_9_Title_Sort_Name"].ToString() + "\"" + "/>" + "\n";
                    xml_out += _9_App_Data_Title_Sort_Name;
                    _10_App_Data_Title_Brief = "<App_Data App=\"MOD\" Name=\"Title_Brief\" Value=\"" + linha["_10_Title_Brief"].ToString() + "\"" + " />" + "\n";
                    xml_out += _10_App_Data_Title_Brief;
                    _11_App_Data_Title = "<App_Data App=\"MOD\" Name=\"Title\" Value=\"" + linha["_11_Title"].ToString() + "\"" + "/>" + "\n";
                    xml_out += _11_App_Data_Title;
                    _12_App_Data_Episode_Name = "<App_Data App=\"MOD\" Name=\"Episode_Name\" Value=\"" + linha["_12_Episode_Name"].ToString() + "\"" + "/>" + "\n";
                    xml_out += _12_App_Data_Episode_Name;
                    _13_App_Data_Episode_ID = "<App_Data App=\"MOD\" Name=\"Episode_ID\" Value=\"" + linha["_13_Episode_ID"].ToString() + "\"" + "/>" + "\n";
                    xml_out += _13_App_Data_Episode_ID;
                    _14_App_Data_Summary_Long = "<App_Data App=\"MOD\" Name=\"Summary_Long\" Value=\"" + linha["_14_Summary_Long"].ToString() + "\"" + "/>" + "\n";
                    xml_out += _14_App_Data_Summary_Long;
                    _15_App_Data_Summary_Medium = "<App_Data App=\"MOD\" Name=\"Summary_Medium\" Value=\"" + linha["_15_Summary_Medium"].ToString() + "\"" + "/>" + "\n";
                    xml_out += _15_App_Data_Summary_Medium;
                    _16_App_Data_Summary_Short = "<App_Data App=\"MOD\" Name=\"Summary_Short\" Value=\"" + linha["_16_Summary_Short"].ToString() + "\"" + " />" + "\n";
                    xml_out += _16_App_Data_Summary_Short;
                    _17_App_Data_Rating = "<App_Data App=\"MOD\" Name=\"Rating\" Value=\"" + linha["_17_Rating"].ToString() + "\"" + "/>" + "\n";
                    xml_out += _17_App_Data_Rating;
                    _18_App_Data_Closed_Captioning = "<App_Data App=\"MOD\" Name=\"Closed_Captioning\" Value=\"" + linha["_18_Closed_Captioning"].ToString() + "\"" + "/>" + "\n";
                    xml_out += _18_App_Data_Closed_Captioning;
                    _19_App_Data_Run_Time = "<App_Data App=\"MOD\" Name=\"Run_Time\" Value=\"" + linha["_19_Run_Time"].ToString() + "\"" + "/>" + "\n";
                    xml_out += _19_App_Data_Run_Time;
                    _20_App_Data_Display_Run_Time = "<App_Data App=\"MOD\" Name=\"Display_Run_Time\" Value=\"" + linha["_20_Display_Run_Time"].ToString() + "\"" + " />" + "\n";
                    xml_out += _20_App_Data_Display_Run_Time;
                    _21_App_Data_Year = "<App_Data App=\"MOD\" Name=\"Year\" Value=\"" + linha["_21_Year"].ToString() + "\"" + " />" + "\n";
                    xml_out += _21_App_Data_Year;
                    _22_App_Data_Country_of_Origin = "<App_Data App=\"MOD\" Name=\"Country_of_Origin\" Value=\"" + linha["_22_Country_Of_Origin"].ToString() + "\"" + " />" + "\n";
                    xml_out += _22_App_Data_Country_of_Origin;
                    _23_App_Data_Studio_Name = "<App_Data App=\"MOD\" Name=\"Studio_Name\" Value=\"" + linha["_23_Studio_Name"].ToString() + "\"" + " />" + "\n";
                    xml_out += _23_App_Data_Studio_Name;
                    _24_App_Data_Actors = "<App_Data App=\"MOD\" Name=\"Actors\" Value=\"" + linha["_24_Actors"].ToString() + "\"" + " />" + "\n";
                    xml_out += _24_App_Data_Actors;
                    _25_App_Data_Actors_Display = "<App_Data App=\"MOD\" Name=\"Actors_Display\" Value=\"" + linha["_25_Actors_Display"].ToString() + "\"" + " />" + "\n";
                    xml_out += _25_App_Data_Actors_Display;
                    _26_App_Data_Director = "<App_Data App=\"MOD\" Name=\"Director\" Value=\"" + linha["_26_Director"].ToString() + "\"" + " />" + "\n";
                    xml_out += _26_App_Data_Director;
                    _27_App_Data_Director_Display = "<App_Data App=\"MOD\" Name=\"Director_Display\" Value=\"" + linha["_27_Director_Display"].ToString() + "\"" + " />" + "\n";
                    xml_out += _27_App_Data_Director_Display;
                    _28_App_Data_Category = "<App_Data App=\"MOD\" Name=\"Category\" Value=\"" + linha["_28_Category"].ToString() + "\"" + " />" + "\n";
                    xml_out += _28_App_Data_Category;
                    _29_App_Data_Genre = "<App_Data App=\"MOD\" Name=\"Genre\" Value=\"" + linha["_29_Genre"].ToString() + "\"" + " />" + "\n";
                    xml_out += _29_App_Data_Genre;
                    _30_App_Data_Genre = "<App_Data App=\"MOD\" Name=\"Genre\" Value=\"" + linha["_30_Genre"].ToString() + "\"" + " />" + "\n";
                    xml_out += _30_App_Data_Genre;

                    xml_out += _21_Fixo_App_Data_Box_Office + "\n";

                    _31_App_Data_Billing_ID = "<App_Data App=\"MOD\" Name=\"Billing_ID\" Value=\"" + linha["_31_Billing_ID"].ToString() + "\"" + " />" + "\n";
                    xml_out += _31_App_Data_Billing_ID;
                    _32_App_Data_Licensing_Window_Start = "<App_Data App=\"MOD\" Name=\"Licensing_Window_Start\" Value=\"" + linha["_32_Licensing_Window_Start"].ToString() + "\"" + " />" + "\n";
                    xml_out += _32_App_Data_Licensing_Window_Start;
                    _33_App_Data_Licensing_Window_End = "<App_Data App=\"MOD\" Name=\"Licensing_Window_End\" Value=\"" + linha["_33_Licensing_Window_End"].ToString() + "\"" + " />" + "\n";
                    xml_out += _33_App_Data_Licensing_Window_End;

                    xml_out += _22_Fixo_App_Data_Preview_Period + "\n";

                    _34_App_Data_Suggested_Price = "<App_Data App=\"MOD\" Name=\"Suggested_Price\" Value=\"" + linha["_34_Suggested_Price"].ToString() + "\"" + " />" + "\n";
                    xml_out += _34_App_Data_Suggested_Price;

                    xml_out += _23_Fixo_MetaData + "\n";

                    xml_out += _24_Fixo_Asset + "\n";

                    xml_out += _25_Fixo_MetaData + "\n";

                    xml_out += _26_Fixo_AMS_Provider + "\n";

                    xml_out += _27_Fixo_Product + "\n";

                    _35_Asset_Name = "Asset_Name=\"" + linha["_35_Asset_Name"].ToString() + "\"" + "\n";
                    xml_out += _35_Asset_Name;

                    xml_out += _28_Fixo_Version_Major + "\n";

                    xml_out += _29_Fixo_Version_Minor + "\n";

                    _36_Description = "Description=\"" + linha["_36_Description"].ToString() + "\"" + "\n";
                    xml_out += _36_Description;
                    _37_Creation_Date = "Creation_Date=\"" + linha["_37_Creation_Date"].ToString() + "\"" + "\n";
                    xml_out += _37_Creation_Date;

                    xml_out += _30_Fixo_Provider_Id + "\n";

                    _38_Asset_Id = "Asset_ID=\"" + linha["_38_Asset_Id"].ToString() + "\"" + "\n";
                    xml_out += _38_Asset_Id;

                    xml_out += _31_Fixo_Asset_Class + "\n";

                    xml_out += _32_Fixo_App_Data_Type + "\n";

                    xml_out += _33_Fixo_App_Data_Screen_Format + "\n";

                    xml_out += _34_Fixo_App_Data_HDContent + "\n";

                    _39_App_Data_Audio_Type = "<App_Data App=\"MOD\" Name=\"Audio_Type\" Value=\"" + linha["_39_Audio_Type"].ToString() + "\"" + " />" + "\n";
                    xml_out += _39_App_Data_Audio_Type;

                    xml_out += _35_Fixo_App_Data_Viewing_Can_Be_Resumed + "\n";

                    xml_out += _36_Fixo_App_Data_Watermarking + "\n";

                    _40_App_Data_Languages = "<App_Data App=\"MOD\" Name=\"Languages\" Value=\"" + linha["_40_Languages"].ToString() + "\"" + " />" + "\n";
                    xml_out += _40_App_Data_Languages;                   

                    xml_out += _37_Fixo_App_Data_Copy_Protection + "\n";

                    _41_App_Data_Content_FileSize = "<App_Data App=\"MOD\" Name=\"Content_FileSize\" Value=\"" + linha["_41_Content_FileSize"].ToString() + "\"" + " />" + "\n";
                    xml_out += _41_App_Data_Content_FileSize;
                    _42_App_Data_Content_CheckSum = "<App_Data App=\"MOD\" Name=\"Content_CheckSum\" Value=\"" + linha["_42_Content_CheckSum"].ToString() + "\"" + " />" + "\n";
                    xml_out += _42_App_Data_Content_CheckSum;

                    xml_out += _38_Fixo_MetaData + "\n";

                    _43_Content = "<Content Value=\"" + linha["_43_Content"].ToString() + "\" />" + "\n";
                    xml_out += _43_Content;

                    xml_out += _39_Fixo_Asset + "\n";

                    xml_out += _40_Fixo_Asset + "\n";

                    xml_out += _41_Fixo_MetaData + "\n";

                    xml_out += _42_Fixo_AMS_Provider + "\n";

                    xml_out += _43_Fixo_Product + "\n";

                    _44_Asset_Name = "Asset_Name=\"" + linha["_44_Asset_Name"].ToString() + "\"" + "\n";
                    xml_out += _44_Asset_Name;

                    xml_out += _44_Fixo_Version_Major + "\n";

                    xml_out += _45_Fixo_Version_Minor + "\n";

                    _45_Description = "Description=\"" + linha["_45_Description"].ToString() + "\"" + "\n";
                    xml_out += _45_Description;
                    _46_Creation_Date = "Creation_Date=\"" + linha["_46_Creation_Date"].ToString() + "\"" + "\n";
                    xml_out += _46_Creation_Date;

                    xml_out += _46_Fixo_Provider_Id + "\n";

                    _47_Asset_Id = "Asset_ID=\"" + linha["_47_Asset_Id"].ToString() + "\"" + "\n";
                    xml_out += _47_Asset_Id;

                    xml_out += _47_Fixo_Asset_Class + "\n";

                    xml_out += _48_Fixo_App_Data_Type + "\n";

                    _48_App_Data_Content_FileSize = "<App_Data App=\"MOD\" Name=\"Content_FileSize\" Value=\"" + linha["_48_Content_FileSize"].ToString() + "\"" + " />" + "\n";
                    xml_out += _48_App_Data_Content_FileSize;
                    _49_App_Data_Content_CheckSum = "<App_Data App=\"MOD\" Name=\"Content_CheckSum\" Value=\"" + linha["_49_Content_CheckSum"].ToString() + "\"" + " />" + "\n";
                    xml_out += _49_App_Data_Content_CheckSum;

                    xml_out += _49_Fixo_MetaData + "\n";

                    _50_Content = "<Content Value=\"" + linha["_50_Content"].ToString() + "\" />" + "\n";
                    xml_out += _50_Content;

                    xml_out += _50_Fixo_Asset + "\n";

                    xml_out += _51_Fixo_Asset + "\n";

                    xml_out += _52_Fixo_ADI;


                    // SE NÃO TIVER INFORMADO O NOME NO EXCEL SERÁ O NÚMERO DO i.
                    if (string.IsNullOrEmpty(linha["NomeArquivoSaida"].ToString()) == true)
                    {
                        nome_arquivo = i.ToString();
                    }
                    else
                    {
                        nome_arquivo = linha["NomeArquivoSaida"].ToString();
                    }

                    // CRIAR ARQUIVO .TXT COM CONTEUDO.
                    string path = nome_arquivo + ".xml";
                    // CRIAR PASTA DE DIRETÓRIO.
                    if (!Directory.Exists(pasta))
                    {
                        Directory.CreateDirectory(pasta);
                        Console.WriteLine(pasta);
                    }

                    // CRIAR UM ARQUIVO PARA GRAVAR.
                    //using (StreamWriter sw = File.CreateText(pasta + "\\" + path))
                    using (Utf8StreamWriter sw = new Utf8StreamWriter(Path.Combine(pasta, path)))
                    {
                        sw.WriteLine(xml_out);
                    }

                    txtCaminho.Text = pasta;


                    //String my_querry = "INSERT INTO TestXml(Original_Title)VALUES('" + linha["ORIGINAL TITLE"].ToString() + "')";

                    //OleDbCommand cmd = new OleDbCommand(my_querry, db);
                    //cmd.ExecuteNonQuery();


                }
                catch (Exception ex)
                {
                    MessageBox.Show("" + ex.Message);
                }
                finally
                {

                }


            }
            MessageBox.Show("Dados Salvos...");
            db.Close();
            ds.Dispose();
            adapter.Dispose();
            adapterFixos.Dispose();
            conexao.Close();
            conexaoFixos.Close();
        }
        #endregion

        #region SELEÇÃO DO ARQUIVO EXCEL QUE DESEJA CONVERTER PARA TXT
        private void btn_Selecionar_Click(object sender, EventArgs e)
        {
            //Define as propriedades do controle 
            //OpenFileDialog
            this.Seleciona_Arq.Multiselect = true;
            this.Seleciona_Arq.Title = "Selecionar Arquivos";
            Seleciona_Arq.InitialDirectory = @"";
            //Filtra para exibir somente arquivos Excel
            Seleciona_Arq.Filter = "Excel (*.XLS;*.XLSX;)|*.XLS;*.XLSX;|" + "All files (*.*)|*.*";
            Seleciona_Arq.CheckFileExists = true;
            Seleciona_Arq.CheckPathExists = true;
            Seleciona_Arq.FilterIndex = 2;
            Seleciona_Arq.RestoreDirectory = true;
            Seleciona_Arq.ReadOnlyChecked = true;
            Seleciona_Arq.ShowReadOnly = true;

            DialogResult dr = this.Seleciona_Arq.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                // Lê os arquivos selecionados 
                foreach (String arquivo in Seleciona_Arq.FileNames)
                {
                    txtArquivo.Text += arquivo;

                }
            }
        }

        private void txtArquivo_TextChanged(object sender, EventArgs e)
        {

        }

        private void Layout_Load(object sender, EventArgs e)
        {

        }

        private void btnAbrir_Click(object sender, EventArgs e)
        {


            bool isfolder = System.IO.Directory.Exists(txtCaminho.Text);
            if (isfolder)
            {
                string argument = @"/select, " + txtCaminho.Text;
                System.Diagnostics.Process.Start("explorer.exe", argument);
            }

        }
        
    }
}
#endregion

