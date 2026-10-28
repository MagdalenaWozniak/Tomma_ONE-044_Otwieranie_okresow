using Hydra;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using OtworzOkres;
using System.Collections.Generic;
using static OtworzOkres.ZS;
using System.Text;
using System;
using System.Data.SqlClient;
using System.Data;
using cdn_api;

//do update: WHERE Okr_id=(SELECT MAX(OKR_ID) FROM[ERPXL_RainbowTours].[CDN].[Okresy] WHERE OKR_ZamknietyDo<>0)

[assembly: CallbackAssemblyDescription(Name: "ONE-044 Otwieranie okresów",
                                           Description: "Otwieranie okresów w Księgowości",
                                           Author: "OneSupport",
                                           Version: "1.0",
                                           SystemVersion: "2021.0.2",
                                           Created: "17-09-2024")]


namespace OtworzOkres
{
    [SubscribeProcedure(Procedures.WybierzOkres, "Otwieranie okresów")]
    public class ZSnaFS_Class : Callback
    {
        
        private ClaWindow RapButton;
        private ClaWindow zakladkaTabWewnetrzne;
        private ClaWindow filtrBtn;
        private ClaWindow listaPoz;
        
        public override void Init()
        {

            AddSubscription(true, 0, Events.JustAfterWindowOpening, new TakeEventDelegate(OnOpenWindow));
            AddSubscription(false, 0, Events.ResizeWindow, new TakeEventDelegate(Resize));
        }



        private bool OnOpenWindow(Procedures ProcId, int ControlId, Events Event)
        {

            ClaWindow Parent = GetWindow();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Rectangle WndBnd = GetWindow().Bounds;
            zakladkaTabWewnetrzne = GetWindow().AllChildren["?List"];
            //listaPoz = GetWindow().AllChildren["?ListaZam"];
            //RapButton = zakladkaTabWewnetrzne.Children.Add(ControlTypes.button);
            RapButton = Parent.Children.Add(ControlTypes.button);
            RapButton.Visible = true;
            RapButton.TextRaw = "Open";
            RapButton.ToolTipRaw = "Otwieranie ostatniego zamkniętego okresu";
            //RapButton.IconRaw = "C:\\OneSupport\\klodka.ico";

            filtrBtn = GetWindow().AllChildren["?Select"];                            //pobiera współrzędne przycisku zapisz
            RapButton.Bounds = new Rectangle(filtrBtn.Bounds.X - 36, filtrBtn.Bounds.Y, 30, 15);
            //listID = GetWindow().AllChildren["?ListaZam"].Id;      //tu sprawdzic

            //RapButton.OnBeforeAccepted += new TakeEventDelegate(zaczytaj_OnAfterAccepted);
            RapButton.OnAfterMouseDown += new TakeEventDelegate(BtnOnMouseDown);
            //pobrane z pola w zakładce Raporty
            return true;
        }

        public bool Resize(Procedures ProcId, int ControlId, Events Event)
        {
            filtrBtn = GetWindow().AllChildren["?Select"];                            //pobiera współrzędne przycisku zapisz
            RapButton.Bounds = new Rectangle(filtrBtn.Bounds.X - 36, filtrBtn.Bounds.Y, 30, 15);

            return true;
        }

        
        public bool BtnOnMouseDown(Procedures ProcId, int ControlId, Events Event)
        {

            try
            {



                string connectionString = "";
                if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 16) == "ERPXL_Tomma_test")
                {
                    connectionString = @"Data Source = COMARCHBD; Initial Catalog = ERPXL_Tomma_test; User ID = sa; Password=fr$g4i9sSa032V,_";
                }
                else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 15) == "ERPXL_TOMMA-ERP")
                {
                    connectionString = @"Data Source = COMARCHBD; Initial Catalog = ERPXL_TOMMA-ERP; User ID = sa; Password=fr$g4i9sSa032V,_";
                }
                else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 18) == "ERPXL_BONUS_2_test")
                {
                    connectionString = @"Data Source = COMARCHBD; Initial Catalog = ERPXL_BONUS_2_test; User ID = sa; Password=fr$g4i9sSa032V,_";
                }
                else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 13) == "ERPXL_BONUS_2")
                {
                    connectionString = @"Data Source = COMARCHBD; Initial Catalog = ERPXL_BONUS_2; User ID = sa; Password=fr$g4i9sSa032V,_";
                }

                if (((Runtime.ConfigurationDictionary.NumerOperatora == 2491 || Runtime.ConfigurationDictionary.NumerOperatora== 2528 || Runtime.ConfigurationDictionary.NumerOperatora== 128) 
                    && Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 7) == "ERPXL_T")
                    || ((Runtime.ConfigurationDictionary.NumerOperatora == 121 || Runtime.ConfigurationDictionary.NumerOperatora == 141 || Runtime.ConfigurationDictionary.NumerOperatora== 14) 
                    && Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 13) == "ERPXL_BONUS_2"))
                {

                    string czyZawieraString = "if ((SELECT OKR_Opis from cdn.Okresy where Okr_id = (SELECT MAX(OKR_ID) FROM[CDN].[Okresy] WHERE OKR_ZamknietyDo<>0)) like '%#%') SELECT  'TAK' else SELECT 'NIE' ";
                    string czyZawiera = ConnectToSQL(connectionString, czyZawieraString);
                    if (czyZawiera.Trim() == "NIE")
                    {

                        string updateOkres = "UPDATE cdn.Okresy set OKR_ZamknietyDo=0 WHERE Okr_id=(SELECT MAX(OKR_ID) FROM [CDN].[Okresy] WHERE OKR_ZamknietyDo<>0)";
                        Update(updateOkres, connectionString);



                        MessageBox.Show("Otwarto ostatni zamknięty okres. Proszę uruchomić ponownie moduł księgowość.", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Ostatni zamknięty okres zawiera #. Zablokowano otwieranie okresu.", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Nie masz uprawnień do otwierania okresów. Skontaktuj się z One Support", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }


            }
            catch(Exception e)
            {
                MessageBox.Show("Błąd otwierania okresu: "+e.ToString(), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            
            
            return true;
        }


        

       

        public string Update(string sql, string connection_string)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    connection.Open();
                    adapter.UpdateCommand = new SqlCommand(sql, connection);
                    adapter.UpdateCommand.ExecuteNonQuery();
                    //MessageBox.Show("Row updated !! ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return sql;
        }




        public string ConnectToSQL(string connection_string, string query)
        {
            string str = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Query");
                    if (dataSet.Tables.Count == 0)
                    {
                        return "";
                    }
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        str = row[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                str = str + "Bład SQL:" + ex.ToString() + "\n";
            }
            return str;
        }

        public override void Cleanup()
        {

        }

    }
}
