using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Xml;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Font = iTextSharp.text.Font;
using System.Threading;
using System.Text;
using System.Linq;


namespace MonEtiquettePrix
{
    class Program
    {
        public static SAPbouiCOM.Application SBO_Application { get; private set; }
        public static SAPbobsCOM.Company oCompany { get; private set; }
        public static SAPbouiCOM.Form oForm { get; set; }

        public static SAPbouiCOM.Grid oGrid;

        private static string FromItem;
        private static string To;
        private static string LisP;
        private static string Item_5;
        private static string Format;
        //private static string pathAndFile = "etiquette.pdf";
        private static string pathAndFile = "C:\\Program Files\\SAP\\SAP Business One\\AddOns\\SFC\\Etiqv9\\X64Client\\etiquette.pdf";
        private static string query = "";
        private SAPbouiCOM.ComboBox oComboBox;

        [STAThread]
        static void Main()
        {
            try
            {
                if (ConnectUI())
                {
                    AddMenuItems();
                
                }
                SBO_Application.AppEvent += new _IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                SBO_Application.ItemEvent += new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                SBO_Application.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                System.Windows.Forms.Application.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        static bool ConnectUI(string connectionString = "")
        {
            bool returnValue = false;
            //#if DEBUG
            /*if (string.IsNullOrEmpty(connectionString))
            {
                connectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
            }
#endif
            var sboGuiApi = new SboGuiApi();*/
            SboGuiApi SboGuiApi = null;
            string sConnectionString;

            SboGuiApi = new SboGuiApi();
            sConnectionString = Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            SboGuiApi.Connect(sConnectionString);
            SBO_Application = SboGuiApi.GetApplication();

            try
            {
                ConnectwithSharedMemory();
                returnValue = true;
            }
            catch (Exception exception)
            {
                var message = string.Format(CultureInfo.InvariantCulture, "{0} Initialization - Error accessing SBO: {1}", "DB_TestConnection", exception.Message);
                SBO_Application.SetStatusBarMessage("Initialisation... " + exception.Message, BoMessageTime.bmt_Short, false);
                returnValue = true;
            }
            return returnValue;
        }

        private static bool ConnectwithSSO()
        {
            oCompany = new SAPbobsCOM.Company();
            oCompany = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
            string cookie = oCompany.GetContextCookie();
            string connInfo = SBO_Application.Company.GetConnectionContext(cookie);

            if (oCompany.Connected == true)
            {
                oCompany.Disconnect();
            }
            int ret = oCompany.SetSboLoginContext(connInfo);
            if (ret != 0)
            {
                SBO_Application.MessageBox("DI Connection failed!", 0, "Ok", "", "");
                return true;
            }
            else
            {
                SBO_Application.StatusBar.SetText("SOAS connected sucessfully", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
                return true;
            }
        }

        static void AddMenuItems()
        {
            //RunDBScript(SBO_Application, oCompany);
            //RunDBScriptNWR();

            Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;


            oMenus = SBO_Application.Menus;

            MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((MenuCreationParams)(SBO_Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams)));
            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = SBO_Application.Menus.Item("3072");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "ETIQUETTE";
                oCreationPackage.String = "Les Etiquettes de prix";
                oMenus.AddEx(oCreationPackage);
               
            }
            catch (Exception)
            { //  Menu already exists
                SBO_Application.SetStatusBarMessage("Menu est déja ouvert", BoMessageTime.bmt_Short, true);
            }
        }

        static void ConnectwithSharedMemory()
        {
            oCompany = new SAPbobsCOM.Company();
            oCompany = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
            //SBO_Application.SetStatusBarMessage("Initialisation... " + oCompany.CompanyName, BoMessageTime.bmt_Short, false);
            SBO_Application.StatusBar.SetText("Addon connecté.." + oCompany.CompanyName, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
        }

        static void SBO_Application_AppEvent(BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case BoAppEventTypes.aet_CompanyChanged:
                    break;
                case BoAppEventTypes.aet_FontChanged:
                    break;
                case BoAppEventTypes.aet_LanguageChanged:
                    break;
                case BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }

        static void SBO_Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "ETIQUETTE")
                {
                    LoadXMLOrSRF("gba.srf", SBO_Application, oCompany);
                    DefaultValue(SBO_Application, oCompany);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        static void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.ItemUID == "ShowS" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
            {
                try
                {
                    SAPbouiCOM.Form oForm1;
                    oForm1 = SBO_Application.Forms.GetForm("U_Etiq", 1);
                    EditText oEditDe = (EditText)oForm1.Items.Item("ITDE").Specific;
                    FromItem = oEditDe.Value;
                    EditText oEditTO = (EditText)oForm1.Items.Item("IDTO").Specific;
                    To = oEditTO.Value;
                    SAPbouiCOM.ComboBox cbx = (SAPbouiCOM.ComboBox)oForm1.Items.Item("LisP").Specific;
                    SAPbouiCOM.ComboBox cbx3 = (SAPbouiCOM.ComboBox)oForm1.Items.Item("Item_5").Specific;
                    LisP = cbx.Selected.Value;
                    SAPbouiCOM.ComboBox cbx2 = (SAPbouiCOM.ComboBox)oForm1.Items.Item("Format").Specific;
                    Format = cbx2.Selected.Value;
                    oEditTO = (EditText)oForm1.Items.Item("Mag").Specific;
                    string mag = oEditTO.Value;
                    oEditTO = (EditText)oForm1.Items.Item("CmP").Specific;
                    string comm = oEditTO.Value;
                    string sous_famille = cbx3.Selected.Value;
                    //GenerateAndDownloadPDF();
                    if (Format == "Promotionnel")
                    {
                        BuildQueryPromo(comm, mag);
                    }
                    else
                    {
                        BuildQuery(comm, mag, sous_famille);
                    }
                   
                    LoadSelectionScreen("selection.srf", SBO_Application, oCompany); 
                }
                catch (Exception ex)
                {
                    SBO_Application.SetStatusBarMessage("Exception: "+ex.Message, BoMessageTime.bmt_Short, true);
                }
            }

            if (pVal.ItemUID == "printSel" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
            {
                SAPbouiCOM.Form oForm1;
                oForm1 = SBO_Application.Forms.GetForm("U_Etiq", 1);
                EditText oEditDe = (EditText)oForm1.Items.Item("ITDE").Specific;
                FromItem = oEditDe.Value;
                EditText oEditTO = (EditText)oForm1.Items.Item("IDTO").Specific;
                To = oEditTO.Value;
                SAPbouiCOM.ComboBox cbx = (SAPbouiCOM.ComboBox)oForm1.Items.Item("LisP").Specific;
                SAPbouiCOM.ComboBox cbx3 = (SAPbouiCOM.ComboBox)oForm1.Items.Item("Item_5").Specific;
                LisP = cbx.Selected.Value;
                string sous_famille = cbx3.Selected.Value;
                SAPbouiCOM.ComboBox cbx2 = (SAPbouiCOM.ComboBox)oForm1.Items.Item("Format").Specific;
                Format = cbx2.Selected.Value;
                oEditTO = (EditText)oForm1.Items.Item("Mag").Specific;
                string mag = oEditTO.Value;
                oEditTO = (EditText)oForm1.Items.Item("CmP").Specific;
                string comm = oEditTO.Value;
                //Standard();
                if (Format == "Promotionnel")
                {
                    BuildQueryPromo(comm, mag);
                }
                else
                {
                    BuildQuery(comm, mag,sous_famille);
                }
                PrintCodeBar(SBO_Application, oCompany, Format);
            }

            if (pVal.ItemUID == "can" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
            {
                SAPbouiCOM.Form oForm1;
                oForm1 = SBO_Application.Forms.GetForm("U_Etiq", 1);
                oForm1.Close();
            }

                if (FormUID == "Etiq" && pVal.EventType == BoEventTypes.et_FORM_CLOSE)
            {
                SAPbouiCOM.Form oForm = null;
                oForm = SBO_Application.Forms.Item(FormUID);
                oForm.Mode = BoFormMode.fm_OK_MODE;
            }

            if (pVal.ItemUID == "OKB" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
            {
                SAPbouiCOM.Form oForm = null;
                oForm = SBO_Application.Forms.Item(FormUID);
                oForm.Close();
            }

        }
        public void ChargerArticles(string formUID, string comboUID)
        {
            try
            {
                // Récupérer le formulaire actuel
                oForm = SBO_Application.Forms.Item(formUID);

                // Récupérer le ComboBox par son UID
                oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("Item_5").Specific;

                // Connexion au DI API pour exécuter une requête SQL
                SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
                Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                // Requête pour récupérer les articles
                string query = "SELECT T0.[U_sous_fam] FROM OITM T0";

                // Exécuter la requête
                oRecordSet.DoQuery(query);

                // Vider le ComboBox avant de le remplir
                //oComboBox.ValidValues.LoadSeries(0, SAPbouiCOM.BoSeriesMode.sf_Remove);

                // Ajouter les valeurs au ComboBox
                while (!oRecordSet.EoF)
                {
                    string sous_famille = oRecordSet.Fields.Item("U_sous_fam").Value.ToString();
                    

                    oComboBox.ValidValues.Add(sous_famille, sous_famille);

                    oRecordSet.MoveNext();
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Erreur: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
        /*   static void DefaultValue(SAPbouiCOM.Application SBO_Application, SAPbobsCOM.Company oCompany)
           {
               SAPbouiCOM.Form oForm1;
               oForm1 = SBO_Application.Forms.GetForm("U_Etiq", 1);
               oForm1.Freeze(true);

               Recordset orecordset;
               orecordset = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

               string query = "SELECT T0.[ListNum], T0.[ListName] FROM OPLN T0";
               orecordset.DoQuery(query);
               string query1 = "SELECT T0.[U_sous_fam] FROM OITM T0";
               orecordset.DoQuery(query1);

               SAPbouiCOM.ComboBox oCombox = (SAPbouiCOM.ComboBox)oForm1.Items.Item("LisP").Specific;
               SAPbouiCOM.ComboBox oCombox1 = (SAPbouiCOM.ComboBox)oForm1.Items.Item("Item_5").Specific;

               //List<string> nameWH = new List<string>();
               while (!orecordset.EoF)
               {
                   oCombox.ValidValues.Add(orecordset.Fields.Item(0).Value.ToString(), orecordset.Fields.Item(1).Value.ToString());

                   orecordset.MoveNext();
               }

               //oCombox.Select(1, BoSearchKey.psk_ByValue);
               oCombox = (SAPbouiCOM.ComboBox)oForm1.Items.Item("Format").Specific;
               oCombox.Select("Standard", BoSearchKey.psk_ByValue);
               oCombox = (SAPbouiCOM.ComboBox)oForm1.Items.Item("LisP").Specific;
               oCombox1 = (SAPbouiCOM.ComboBox)oForm1.Items.Item("Item_5").Specific;
               oCombox.Select("1", BoSearchKey.psk_ByValue);
               oForm1.Freeze(false);
               oForm1.Visible = true;
           }*/
        static void DefaultValue(SAPbouiCOM.Application SBO_Application, SAPbobsCOM.Company oCompany)
        {
            // Récupération du formulaire utilisateur
            SAPbouiCOM.Form oForm1;
            oForm1 = SBO_Application.Forms.GetForm("U_Etiq", 1);
            oForm1.Freeze(true);

            // Création du premier Recordset pour récupérer les listes de prix
            Recordset orecordset;
            orecordset = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = "SELECT T0.[ListNum], T0.[ListName] FROM OPLN T0";
            orecordset.DoQuery(query);

            // Création du deuxième Recordset pour récupérer les sous-familles
            Recordset orecordset1;
            orecordset1 = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query1 = "SELECT DISTINCT T0.[U_sous_fam] FROM OITM T0 where T0.[U_sous_fam] is not null order by T0.[U_sous_fam]";
            orecordset1.DoQuery(query1);

            // Récupération des ComboBox dans le formulaire
            SAPbouiCOM.ComboBox oCombox = (SAPbouiCOM.ComboBox)oForm1.Items.Item("LisP").Specific;
            SAPbouiCOM.ComboBox oCombox1 = (SAPbouiCOM.ComboBox)oForm1.Items.Item("Item_5").Specific;

            // Remplissage de la liste des prix
            while (!orecordset.EoF)
            {
                string listNum = orecordset.Fields.Item(0).Value.ToString();
                string listName = orecordset.Fields.Item(1).Value.ToString();
                oCombox.ValidValues.Add(listNum, listName);
                
                orecordset.MoveNext();
            }

            // Remplissage de la liste des sous-familles
            while (!orecordset1.EoF)
            {
                string sousFamille = orecordset1.Fields.Item(0).Value.ToString();
                oCombox1.ValidValues.Add(sousFamille, sousFamille); // Même valeur pour clé et affichage

                orecordset1.MoveNext();
            }

            // Sélection de valeurs par défaut avec vérification
            oCombox = (SAPbouiCOM.ComboBox)oForm1.Items.Item("Format").Specific;
            if (oCombox.ValidValues.Count > 0)
            {
                oCombox.Select("Standard", BoSearchKey.psk_ByValue);
            }

            oCombox = (SAPbouiCOM.ComboBox)oForm1.Items.Item("LisP").Specific;
            if (oCombox.ValidValues.Count > 0)
            {
                oCombox.Select("1", BoSearchKey.psk_ByValue);
            }

            // Déverrouillage et affichage du formulaire
            oForm1.Freeze(false);
            oForm1.Visible = true;
        }


        static void LoadXMLOrSRF(string oPath, SAPbouiCOM.Application SBO_Application, SAPbobsCOM.Company oCompany)
        {
            XmlDocument oXml = new XmlDocument();
            string ls_Xml = "";
            Assembly AppAssembly = Assembly.GetEntryAssembly();
            Stream SRFFile = AppAssembly.GetManifestResourceStream("MonEtiquettePrix.SRF." + oPath);
            oXml.Load(SRFFile);
            ls_Xml = oXml.InnerXml.ToString();
            SBO_Application.LoadBatchActions(ref ls_Xml);

        }

        static void LoadSelectionScreen(string oPath, SAPbouiCOM.Application SBO_Application, SAPbobsCOM.Company oCompany)
        {
            try
            {
                XmlDocument oXml = new XmlDocument();
                string ls_Xml = "";
                Assembly AppAssembly = Assembly.GetEntryAssembly();
                Stream SRFFile = AppAssembly.GetManifestResourceStream("MonEtiquettePrix.SRF." + oPath);
                oXml.Load(SRFFile);
                ls_Xml = oXml.InnerXml.ToString();
                SBO_Application.LoadBatchActions(ref ls_Xml);

                oForm = SBO_Application.Forms.GetForm("U_Sel", 1);
                oForm.Freeze(true);
                oForm.State = BoFormStateEnum.fs_Maximized;
                Grid oGrid;
                
                oGrid = (Grid)oForm.Items.Item("gridITM").Specific;
                oForm.DataSources.DataTables.Add("selection");
                oForm.DataSources.DataTables.Item("selection").ExecuteQuery(query);
                oGrid.DataTable = oForm.DataSources.DataTables.Item("selection");
                if (!oGrid.DataTable.IsEmpty)
                {
                    if (Format == "Promotionnel")
                    {
                        oGrid.Columns.Item("NEW PRICE").Editable = false;
                    }
                    EditTextColumn column;
                    column = (EditTextColumn)oGrid.Columns.Item("Code Article");
                    column.LinkedObjectType = "4";
                    oGrid.Columns.Item("Description article").Editable = false;
                    oGrid.Columns.Item("Code Article").Editable = false;
                    oGrid.Columns.Item("Code barres").Editable = false;
                    oGrid.Columns.Item("Prix de vente").Editable = false;
                    oGrid.Columns.Item("PCB").Editable = false;
                    oGrid.AutoResizeColumns();
                }
                else
                {
                    if (Format == "Promotionnel")
                    {
                        oGrid.Columns.Item("NEW PRICE").Editable = false;
                    }
                    oGrid.Columns.Item("Description article").Editable = false;
                    oGrid.Columns.Item("Code Article").Editable = false;
                    oGrid.Columns.Item("Code barres").Editable = false;
                    oGrid.Columns.Item("Prix de vente").Editable = false;
                    oGrid.Columns.Item("PCB").Editable = false;
                    oGrid.AutoResizeColumns();
                    SBO_Application.StatusBar.SetText("Aucune donnée trouvée.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            }
            oForm.Freeze(false);
            oForm.Visible = true;
        }

        static void PrintCodeBar(SAPbouiCOM.Application SBO_Application, SAPbobsCOM.Company oCompany, string format)
        {
            try
            {            
                CloseAndDeletePdf(pathAndFile);
            
                float widthMm = 210;  // Largeur en mm (ex: A4)
                float heightMm = 297; // Hauteur en mm (ex: A4)

                float widthPoints = MillimetersToPoints(widthMm);
                float heightPoints = MillimetersToPoints(heightMm);
               
                if (format == "Standard")
                {
                    Thread thread = new Thread(new ThreadStart(ProduitFrais));
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    thread.Join();
                }
                else if (format == "Promotionnel")
                {
                    Thread thread = new Thread(new ThreadStart(Promotionnel));
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    thread.Join();
                }
                else if (format == "Produit Frais")
                {
                    Thread thread = new Thread(new ThreadStart(ProduitFrais));
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    thread.Join();
                }
                else if (format == "Petits Produits")
                {
                    Thread thread = new Thread(new ThreadStart(PetitProduit));
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    thread.Join();
                }
                else
                {
                    Thread thread = new Thread(new ThreadStart(ProduitEmballage));
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    thread.Join();
                }
                if (File.Exists(pathAndFile))
                {
                    //Process.Start(pathAndFile);
                    SBO_Application.SendFileToBrowser(pathAndFile);
                }                
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Erreur lors de la génération: " + ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
                //SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        static void PetitProduit()
        {

            var heightMm = 20;
            var heightPoints = MillimetersToPoints(heightMm);

            Document document = new Document(PageSize.A4, 0, 0, 0, 0);
            var designation = FontFactory.GetFont("Arial", (float)6, Font.BOLD);
            var prix = FontFactory.GetFont("Arial", (float)7, Font.BOLD);
            var prixPromo = FontFactory.GetFont("Arial", (float)11, Font.BOLD, BaseColor.RED);
            var unite = FontFactory.GetFont("Arial", (float)4, Font.BOLD);
            Recordset orec;
            orec = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            orec.DoQuery(query);

            if (orec.RecordCount > 0)
            {
                using (FileStream fs = new FileStream(pathAndFile, FileMode.Create))
                {
                    PdfWriter wt = PdfWriter.GetInstance(document, fs);
                    try
                    {
                        document.Open();

                        PdfPTable table2 = new PdfPTable(210);
                        table2.WidthPercentage = 100;
                        int cpt = 0;

                        while (!orec.EoF)
                        {
                            string barcode = orec.Fields.Item(2).Value.ToString();
                            string article = orec.Fields.Item(1).Value.ToString();
                            string price = orec.Fields.Item(3).Value.ToString("### ### ###.##") + " CFA";
                            string PCBC = orec.Fields.Item(4).Value.ToString();
                            Barcode128 bc = new Barcode128();
                            bc.TextAlignment = Element.ALIGN_CENTER;
                            bc.Code = barcode;
                            bc.StartStopText = false;
                            bc.CodeType = iTextSharp.text.pdf.Barcode128.CODE128;
                            bc.Extended = true;
                            bc.BarHeight = 12f;
                            bc.Size = 6f;
                            PdfPTable table = new PdfPTable(100);
                            if (cpt == 0)
                            {
                                PdfPCell cellFirst = new PdfPCell(new Phrase(""));
                                cellFirst.Colspan = 5;
                                cellFirst.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                                cellFirst.PaddingLeft = 0f;
                                cellFirst.PaddingBottom = 0f;
                                cellFirst.PaddingTop = 0f;
                                cellFirst.PaddingRight = 0f;
                                cellFirst.Border = Rectangle.NO_BORDER;
                                cellFirst.FixedHeight = heightPoints;
                                table2.AddCell(cellFirst);
                                cpt = 0;
                            }
                            Chunk chnk = new Chunk(article, designation);

                            Phrase ph1 = new Phrase();
                            ph1.Add(chnk);
                            PdfPCell cell1 = new PdfPCell(ph1);
                            table.WidthPercentage = 100;
                            cell1.Colspan = 100;
                            cell1.PaddingBottom = 2f;
                            cell1.PaddingTop = 1f;
                            cell1.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell1.Border = Rectangle.NO_BORDER;
                            table.AddCell(cell1);
                            table.SpacingAfter = 4f;
                            Chunk ck1 = new Chunk(price, prix);
                            Phrase ph2 = new Phrase("");
                            ph2.Add(ck1);
                            Chunk ck2 = new Chunk(" ", unite);
                            ph2.Add(ck2);
                            PdfPCell cell2a = new PdfPCell(ph2);
                            cell2a.Colspan = 100;
                            cell2a.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell2a.PaddingTop = 0f;
                            cell2a.PaddingBottom = 3f;
                            cell2a.Border = Rectangle.NO_BORDER;
                            table.AddCell(cell2a);

                            table.SpacingAfter = 4f;
                            PdfContentByte cb = wt.DirectContent;
                            Image image = bc.CreateImageWithBarcode(cb, BaseColor.BLACK, BaseColor.BLACK);
                            PdfPCell cell2 = new PdfPCell(image);
                            cell2.Colspan = 100;
                            cell2.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell2.Border = Rectangle.NO_BORDER;
                            table.AddCell(cell2);

                            var textDate = FontFactory.GetFont("Arial", (float)5, Font.NORMAL);
                            var textPCB = FontFactory.GetFont("Arial", (float)5, Font.NORMAL);
         
                            Phrase PCB = new Phrase("PCB:" + PCBC + Environment.NewLine, textPCB);
                            Phrase DateImpression = new Phrase("DI: " + DateTime.Now.ToString("dd/MM/yyyy"), textDate);

                            PdfPCell cellPCB = new PdfPCell(PCB);
                            cellPCB.Colspan = 50;
                            cellPCB.HorizontalAlignment = 0; //0=Left, 1=Centre, 2=Right
                            cellPCB.PaddingTop = 1f;
                            cellPCB.PaddingBottom = 1f;
                            cellPCB.PaddingLeft = 10f;
                            cellPCB.Border = Rectangle.NO_BORDER;
                            table.AddCell(cellPCB);

                            PdfPCell cellDateImpression = new PdfPCell(DateImpression);
                            cellDateImpression.Colspan = 50;
                            cellDateImpression.HorizontalAlignment = 2; //0=Left, 1=Centre, 2=Right
                            cellDateImpression.PaddingTop = 1f;
                            cellDateImpression.PaddingBottom = 1f;
                            cellDateImpression.PaddingRight = 10f;
                            cellDateImpression.Border = Rectangle.NO_BORDER;
                            table.AddCell(cellDateImpression);

                            PdfPCell cell0 = new PdfPCell(table);
                            cell0.Colspan = 40;
                            cell0.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell0.PaddingLeft = 0f;
                            cell0.PaddingBottom = 0f;
                            cell0.PaddingTop = 0f;
                            cell0.PaddingRight = 0f;

                            cell0.FixedHeight = heightPoints;

                            table2.AddCell(cell0);
                            cpt += 1;
                            if (cpt == 5)
                            {
                                PdfPCell celllast = new PdfPCell(new Phrase(""));
                                celllast.Colspan = 5;
                                celllast.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                                celllast.PaddingLeft = 0f;
                                celllast.PaddingBottom = 0f;
                                celllast.PaddingTop = 0f;
                                celllast.PaddingRight = 0f;
                                celllast.Border = Rectangle.NO_BORDER;
                                celllast.FixedHeight = heightPoints;
                                table2.AddCell(celllast);
                                cpt = 0;
                            }

                            orec.MoveNext();
                        }

                        if (cpt > 0)
                        {
                            PdfPCell celllast = new PdfPCell(new Phrase(""));
                            celllast.Colspan = (5 - cpt) * 40 + 5;
                            celllast.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            celllast.PaddingLeft = 0f;
                            celllast.PaddingBottom = 0f;
                            celllast.PaddingTop = 0f;
                            celllast.PaddingRight = 0f;
                            celllast.Border = Rectangle.NO_BORDER;
                            celllast.FixedHeight = heightPoints;
                            table2.AddCell(celllast);
                        }
                        document.Add(table2);
                        SBO_Application.StatusBar.SetText("Opération éffectuée avec succès. Fichier sauvegardé dans l'emplacement: " + pathAndFile, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
                    }
                    catch (Exception ex)
                    {
                        SBO_Application.StatusBar.SetText("Erreur lors de la génération: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    }
                    finally
                    {
                        document.Close();
                    }
                }
             
                Process.Start(new ProcessStartInfo(pathAndFile) { UseShellExecute = true });
            }
        }

        static void Promotionnel()
        {
            var heightMm = 50;
            //widthPoints = MillimetersToPoints(widthMm);
            var heightPoints = MillimetersToPoints(heightMm);

            Document document = new Document(PageSize.A4, 0, 0, 0, 0);
            /*var path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string filename = "article.pdf";
            string pathAndFile = System.IO.Path.Combine(path, filename);*/

            var designation = FontFactory.GetFont("Arial", (float)12, Font.BOLD);
            var prix = FontFactory.GetFont("Arial", (float)15, Font.BOLD);
            var prixPromo = FontFactory.GetFont("Arial", (float)17, Font.BOLD, BaseColor.RED);
            var unite = FontFactory.GetFont("Arial", (float)9, Font.BOLD);

            Recordset orec;
            orec = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            orec.DoQuery(query);
            if (orec.RecordCount > 0)
            {
                using (FileStream fs = new FileStream(pathAndFile, FileMode.Create))
                {
                    PdfWriter wt = PdfWriter.GetInstance(document, fs);
                    try
                    {
                        document.Open();

                        PdfPTable table2 = new PdfPTable(210);
                        table2.WidthPercentage = 100;
                        int cpt = 0;

                        while (!orec.EoF)
                        {
                            string barcode = orec.Fields.Item(2).Value.ToString();
                            string article = orec.Fields.Item(0).Value.ToString();
                            string price = orec.Fields.Item(3).Value.ToString("### ### ###.##") + " CFA";
                            string PCBC = orec.Fields.Item(4).Value.ToString();
                            string newprice = orec.Fields.Item(5).Value.ToString("### ### ###.##") + " CFA";
                            Barcode128 bc = new Barcode128();
                            bc.TextAlignment = Element.ALIGN_CENTER;
                            bc.Code = barcode;
                            bc.StartStopText = false;
                            bc.CodeType = iTextSharp.text.pdf.Barcode128.CODE128;
                            bc.Extended = true;
                            bc.BarHeight = 25f;
                            bc.Size = 10f;
                            PdfPTable table = new PdfPTable(100);
                            if (cpt == 0)
                            {
                                PdfPCell cellFirst = new PdfPCell(new Phrase(""));
                                cellFirst.Colspan = 30;
                                cellFirst.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                                cellFirst.PaddingLeft = 0f;
                                cellFirst.PaddingBottom = 0f;
                                cellFirst.PaddingTop = 0f;
                                cellFirst.PaddingRight = 0f;
                                cellFirst.Border = Rectangle.NO_BORDER;
                                cellFirst.FixedHeight = heightPoints;
                                table2.AddCell(cellFirst);
                                cpt = 0;
                            }
                            Chunk chnk = new Chunk(article, designation);

                            Phrase ph1 = new Phrase();
                            ph1.Add(chnk);
                            PdfPCell cell1 = new PdfPCell(ph1);
                            table.WidthPercentage = 100;
                            cell1.Colspan = 100;
                            cell1.PaddingBottom = 5f;
                            cell1.PaddingTop = 5f;
                            cell1.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell1.Border = Rectangle.NO_BORDER;
                            table.AddCell(cell1);

                            table.SpacingAfter = 4f;

                            Chunk ck1 = new Chunk(price, prix);
                            ck1.SetUnderline(1.0F, 5.0F);
                            Chunk ck2 = new Chunk(newprice, prixPromo);


                            Phrase ph2 = new Phrase("");
                            ph2.Add(ck1);
                            Phrase ph3 = new Phrase("");
                            ph3.Add(ck2);
                            Chunk ck3 = new Chunk("", unite);
                            ph3.Add(ck3);

                            PdfPCell cell2a = new PdfPCell(ph2);
                            cell2a.Colspan = 100;
                            cell2a.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell2a.PaddingTop = 2f;
                            cell2a.Border = Rectangle.NO_BORDER;
                            table.AddCell(cell2a);

                            PdfPCell cell2b = new PdfPCell(ph3);
                            cell2b.Colspan = 100;
                            cell2b.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell2b.PaddingBottom = 20f;
                            cell2b.Border = Rectangle.NO_BORDER;
                            table.AddCell(cell2b);

                            table.SpacingAfter = 4f;
                            PdfContentByte cb = wt.DirectContent;
                            iTextSharp.text.Image image = bc.CreateImageWithBarcode(cb, iTextSharp.text.BaseColor.BLACK, iTextSharp.text.BaseColor.BLACK);
                            PdfPCell cell2 = new PdfPCell(image);
                            cell2.Colspan = 100;
                            cell2.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell2.Border = Rectangle.NO_BORDER;
                            table.AddCell(cell2);

                            var textDate = FontFactory.GetFont("Arial", (float)7, Font.NORMAL);
                            var textPCB = FontFactory.GetFont("Arial", (float)7, Font.NORMAL);
                            Phrase PCB = new Phrase("PCB: " + PCBC + Environment.NewLine, textPCB);
                            Phrase DateImpression = new Phrase("DI: " + DateTime.Now.ToString("dd/MM/yyyy"), textDate);

                            PdfPCell cellPCB = new PdfPCell(PCB);
                            cellPCB.Colspan = 50;
                            cellPCB.HorizontalAlignment = 0; //0=Left, 1=Centre, 2=Right
                            cellPCB.PaddingTop = 3f;
                            cellPCB.PaddingBottom = 1f;
                            cellPCB.PaddingLeft = 10f;
                            cellPCB.Border = Rectangle.NO_BORDER;
                            table.AddCell(cellPCB);

                            PdfPCell cellDateImpression = new PdfPCell(DateImpression);
                            cellDateImpression.Colspan = 50;
                            cellDateImpression.HorizontalAlignment = 2; //0=Left, 1=Centre, 2=Right
                            cellDateImpression.PaddingTop = 3f;
                            cellDateImpression.PaddingBottom = 1f;
                            cellDateImpression.PaddingRight = 10f;
                            cellDateImpression.Border = Rectangle.NO_BORDER;
                            table.AddCell(cellDateImpression);

                            PdfPCell cell0 = new PdfPCell(table);
                            cell0.Colspan = 75;
                            cell0.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell0.PaddingLeft = 0f;
                            cell0.PaddingBottom = 0f;
                            cell0.PaddingTop = 0f;
                            cell0.PaddingRight = 0f;

                            cell0.FixedHeight = heightPoints;

                            table2.AddCell(cell0);
                            cpt += 1;
                            if (cpt == 2)
                            {
                                PdfPCell celllast = new PdfPCell(new Phrase(""));
                                celllast.Colspan = 30;
                                celllast.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                                celllast.PaddingLeft = 0f;
                                celllast.PaddingBottom = 0f;
                                celllast.PaddingTop = 0f;
                                celllast.PaddingRight = 0f;
                                celllast.Border = Rectangle.NO_BORDER;
                                celllast.FixedHeight = heightPoints;
                                table2.AddCell(celllast);
                                cpt = 0;
                            }
                            orec.MoveNext();
                        }
                        if (cpt > 0)
                        {
                            PdfPCell celllast = new PdfPCell(new Phrase(""));
                            celllast.Colspan = (2 - cpt) * 75 + 30;
                            celllast.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            celllast.PaddingLeft = 0f;
                            celllast.PaddingBottom = 0f;
                            celllast.PaddingTop = 0f;
                            celllast.PaddingRight = 0f;
                            celllast.Border = Rectangle.NO_BORDER;
                            celllast.FixedHeight = heightPoints;
                            table2.AddCell(celllast);
                        }

                        document.Add(table2);
                        SBO_Application.StatusBar.SetText("Opération éffectuée avec succès. Fichier sauvegardé dans l'emplacement: " + pathAndFile, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);

                    }
                    catch (Exception ex)
                    {
                        SBO_Application.StatusBar.SetText("Erreur lors de la génération: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    }
                    finally
                    {
                        document.Close();
                    }
                }

                Process.Start(new ProcessStartInfo(pathAndFile) { UseShellExecute = true });
            }
            else
            {
                SBO_Application.StatusBar.SetText("Aucune donnée trouvée.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            }
            //}
            //}
        }

        static void ProduitFrais()
        {
            var heightMm = 40;
            //widthPoints = MillimetersToPoints(widthMm);
            var heightPoints = MillimetersToPoints(heightMm);

            Document document = new Document(PageSize.A4, 0, 0, 0, 0);
            /*var path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string filename = "article.pdf";
            string pathAndFile = System.IO.Path.Combine(path, filename);*/
            var designation = FontFactory.GetFont("Arial", (float)11, Font.BOLD);
            var prix = FontFactory.GetFont("Arial", (float)16, Font.BOLD);
            var prixPromo = FontFactory.GetFont("Arial", (float)18, Font.BOLD, BaseColor.RED);
            var unite = FontFactory.GetFont("Arial", (float)9, Font.BOLD);

            Recordset orec;
            orec = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            /*string query = "";
            if (!string.IsNullOrEmpty(FromItem) && !string.IsNullOrEmpty(To))
            {
                query = $@"SELECT T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB' FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                                INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum WHERE T2.[ListNum] ={LisP} AND T0.[ItemCode] BETWEEN '{FromItem}' AND '{To}'";
            }
            else
            {
                query = $@"SELECT T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB' FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                                INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum WHERE T2.[ListNum] ={LisP}";
            }*/
            orec.DoQuery(query);
            if (orec.RecordCount > 0)
            {
                using (FileStream fs = new FileStream(pathAndFile, FileMode.Create))
                {
                    PdfWriter wt = PdfWriter.GetInstance(document, fs);
                    try
                    {
                        document.Open();

                        PdfPTable table2 = new PdfPTable(210);
                        table2.WidthPercentage = 100;
                        int cpt = 0;
                        while (!orec.EoF)
                        {
                            string barcode = orec.Fields.Item(2).Value.ToString();
                            string article = orec.Fields.Item(1).Value.ToString();
                            string price = orec.Fields.Item(3).Value.ToString("### ### ###.##") + " CFA";
                            string PCBC = orec.Fields.Item(4).Value.ToString();
                            Barcode128 bc = new Barcode128();
                            bc.TextAlignment = Element.ALIGN_CENTER;
                            bc.Code = barcode;
                            bc.StartStopText = false;
                            bc.CodeType = Barcode128.CODE128;
                            bc.Extended = true;
                            bc.BarHeight = 25f;
                            bc.Size = 10f;
                            PdfPTable table = new PdfPTable(100);
                            if (cpt == 0)
                            {
                                PdfPCell cellFirst = new PdfPCell(new Phrase(""));
                                cellFirst.Colspan = 15;
                                cellFirst.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                                cellFirst.PaddingLeft = 0f;
                                cellFirst.PaddingBottom = 0f;
                                cellFirst.PaddingTop = 0f;
                                cellFirst.PaddingRight = 0f;
                                cellFirst.Border = Rectangle.NO_BORDER;
                                cellFirst.FixedHeight = heightPoints;
                                table2.AddCell(cellFirst);
                                cpt = 0;
                            }
                            Chunk chnk = new Chunk(article, designation);

                            Phrase ph1 = new Phrase();
                            ph1.Add(chnk);
                            PdfPCell cell1 = new PdfPCell(ph1);
                            table.WidthPercentage = 100;
                            cell1.Colspan = 100;
                            cell1.PaddingBottom = 5f;
                            cell1.PaddingTop = 5f;
                            cell1.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell1.Border = Rectangle.NO_BORDER;
                            table.AddCell(cell1);

                            table.SpacingAfter = 4f;

                            Chunk ck1 = new Chunk(price, prix);

                            Phrase ph2 = new Phrase("");
                            ph2.Add(ck1);

                            Chunk ck2 = new Chunk(" " + "", unite);
                            ph2.Add(ck2);

                            PdfPCell cell2a = new PdfPCell(ph2);
                            cell2a.Colspan = 100;
                            cell2a.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell2a.PaddingTop = 2f;
                            cell2a.PaddingBottom = 10f;
                            cell2a.Border = Rectangle.NO_BORDER;
                            table.AddCell(cell2a);

                            table.SpacingAfter = 4f;
                            PdfContentByte cb = wt.DirectContent;
                            iTextSharp.text.Image image = bc.CreateImageWithBarcode(cb, iTextSharp.text.BaseColor.BLACK, iTextSharp.text.BaseColor.BLACK);
                            PdfPCell cell2 = new PdfPCell(image);
                            cell2.Colspan = 100;
                            cell2.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell2.Border = Rectangle.NO_BORDER;
                            table.AddCell(cell2);

                            //table.AddCell(cell2);
                            var textDate = FontFactory.GetFont("Arial", (float)7, Font.NORMAL);
                            var textPCB = FontFactory.GetFont("Arial", (float)7, Font.NORMAL);

                            Phrase PCB = new Phrase("PCB:" + PCBC + Environment.NewLine, textPCB);
                            Phrase DateImpression = new Phrase("DI:" + DateTime.Now.ToString("dd/MM/yyyy"), textDate);

                            PdfPCell cellPCB = new PdfPCell(PCB);
                            cellPCB.Colspan = 50;
                            cellPCB.HorizontalAlignment = 0; //0=Left, 1=Centre, 2=Right
                            cellPCB.PaddingTop = 3f;
                            cellPCB.PaddingBottom = 1f;
                            cellPCB.PaddingLeft = 10f;
                            cellPCB.Border = Rectangle.NO_BORDER;
                            table.AddCell(cellPCB);

                            PdfPCell cellDateImpression = new PdfPCell(DateImpression);
                            cellDateImpression.Colspan = 50;
                            cellDateImpression.HorizontalAlignment = 2; //0=Left, 1=Centre, 2=Right
                            cellDateImpression.PaddingTop = 3f;
                            cellDateImpression.PaddingBottom = 1f;
                            cellDateImpression.PaddingRight = 10f;
                            cellDateImpression.Border = Rectangle.NO_BORDER;
                            table.AddCell(cellDateImpression);

                            PdfPCell cell0 = new PdfPCell(table);
                            cell0.Colspan = 60;
                            cell0.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell0.PaddingLeft = 0f;
                            cell0.PaddingBottom = 0f;
                            cell0.PaddingTop = 0f;
                            cell0.PaddingRight = 0f;
                            cell0.FixedHeight = heightPoints;

                            table2.AddCell(cell0);
                            cpt += 1;
                            if (cpt == 3)
                            {
                                PdfPCell celllast = new PdfPCell(new Phrase(""));
                                celllast.Colspan = 15;
                                celllast.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                                celllast.PaddingLeft = 0f;
                                celllast.PaddingBottom = 0f;
                                celllast.PaddingTop = 0f;
                                celllast.PaddingRight = 0f;
                                celllast.Border = Rectangle.NO_BORDER;
                                celllast.FixedHeight = heightPoints;
                                table2.AddCell(celllast);
                                cpt = 0;
                            }
                            orec.MoveNext();
                        }

                        if (cpt > 0)
                        {
                            PdfPCell celllast = new PdfPCell(new Phrase(""));
                            celllast.Colspan = (3 - cpt) * 60 + 15;
                            celllast.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            celllast.PaddingLeft = 0f;
                            celllast.PaddingBottom = 0f;
                            celllast.PaddingTop = 0f;
                            celllast.PaddingRight = 0f;
                            celllast.Border = Rectangle.NO_BORDER;
                            celllast.FixedHeight = heightPoints;
                            table2.AddCell(celllast);
                        }
                        document.Add(table2);
                        SBO_Application.StatusBar.SetText("Opération éffectuée avec succès. Fichier sauvegardé dans l'emplacement: " + pathAndFile, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
                    }
                    catch (Exception ex)
                    {
                        SBO_Application.StatusBar.SetText("Erreur lors de la génération: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    }
                    finally
                    {
                        document.Close();
                    }
                }

                Process.Start(new ProcessStartInfo(pathAndFile) { UseShellExecute = true });
            }
            else
            {
                SBO_Application.StatusBar.SetText("Aucune donnée trouvée" + oCompany.CompanyName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            }
            //FileStream fs = new FileStream(pathAndFile, FileMode.Create);
            //}
            //}

        }

        static void ProduitEmballage()
        {
            /* using (SaveFileDialog saveFileDialog = new SaveFileDialog())
             {
                 saveFileDialog.Filter = "PDF Files|*.pdf";
                 saveFileDialog.Title = "Etiquettes de prix";
                 saveFileDialog.DefaultExt = "pdf";
                 saveFileDialog.AddExtension = true;
                 if (saveFileDialog.ShowDialog() == DialogResult.OK)
                 {
                     string pathAndFile = saveFileDialog.FileName;*/
            //widthMm = 80;
            var heightMm = 60;
            //widthPoints = MillimetersToPoints(widthMm);
            var heightPoints = MillimetersToPoints(heightMm);

            Document document = new Document(PageSize.A4, 0, 0, 0, 0);
            /*var path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string filename = "article.pdf";
            string pathAndFile = System.IO.Path.Combine(path, filename);*/
            var designation = FontFactory.GetFont("Arial", (float)14, Font.BOLD);
            var prix = FontFactory.GetFont("Arial", (float)30, Font.BOLD);
            var prixPromo = FontFactory.GetFont("Arial", (float)30, Font.BOLD, BaseColor.RED);
            var unite = FontFactory.GetFont("Arial", (float)12, Font.BOLD);
            Recordset orec;
            orec = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

           /* string query = "";
            if (!string.IsNullOrEmpty(FromItem) && !string.IsNullOrEmpty(To))
            {
                query = $@"SELECT T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB' FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                                INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum WHERE T2.[ListNum] ={LisP} AND T0.[ItemCode] BETWEEN '{FromItem}' AND '{To}'";
            }
            else
            {
                query = $@"SELECT T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB' FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                                INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum WHERE T2.[ListNum] ={LisP}";
            }*/

            orec.DoQuery(query);
            if (orec.RecordCount > 0)
            {
                using (FileStream fs = new FileStream(pathAndFile, FileMode.Create))
                {
                    PdfWriter wt = PdfWriter.GetInstance(document, fs);
                    try
                    {
                        document.Open();
                        PdfPTable table2 = new PdfPTable(210);
                        table2.WidthPercentage = 100;
                        int cpt = 0;
                        while (!orec.EoF)
                        {
                            string barcode = orec.Fields.Item(2).Value.ToString();
                            string article = orec.Fields.Item(1).Value.ToString();
                            string price = orec.Fields.Item(3).Value.ToString("### ### ###.##") + " CFA";
                            string PCBC = orec.Fields.Item(4).Value.ToString();
                            iTextSharp.text.pdf.Barcode128 bc = new Barcode128();
                            bc.TextAlignment = Element.ALIGN_CENTER;
                            bc.Code = barcode;
                            bc.StartStopText = false;
                            bc.CodeType = iTextSharp.text.pdf.Barcode128.CODE128;
                            bc.Extended = true;
                            bc.BarHeight = 25f;
                            bc.Size = 10f;
                            PdfPTable table = new PdfPTable(100);
                            if (cpt == 0)
                            {
                                PdfPCell cellFirst = new PdfPCell(new Phrase(""));
                                cellFirst.Colspan = 25;
                                cellFirst.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                                cellFirst.PaddingLeft = 0f;
                                cellFirst.PaddingBottom = 0f;
                                cellFirst.PaddingTop = 0f;
                                cellFirst.PaddingRight = 0f;
                                cellFirst.Border = Rectangle.NO_BORDER;
                                cellFirst.FixedHeight = heightPoints;
                                table2.AddCell(cellFirst);
                                cpt = 0;
                            }
                            Chunk chnk = new Chunk(article, designation);

                            Phrase ph1 = new Phrase();
                            ph1.Add(chnk);
                            PdfPCell cell1 = new PdfPCell(ph1);
                            table.WidthPercentage = 100;
                            cell1.Colspan = 100;
                            cell1.PaddingBottom = 5f;
                            cell1.PaddingTop = 5f;
                            cell1.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell1.Border = Rectangle.NO_BORDER;
                            table.AddCell(cell1);
                            Chunk ck1 = new Chunk(price, prix);

                            Phrase ph2 = new Phrase("");
                            ph2.Add(ck1);

                            Chunk ck2 = new Chunk("", unite);
                            ph2.Add(ck2);

                            PdfPCell cell2a = new PdfPCell(ph2);
                            cell2a.Colspan = 100;
                            cell2a.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell2a.PaddingTop = 10f;
                            cell2a.PaddingBottom = 0f;
                            cell2a.Border = Rectangle.NO_BORDER;
                            table.AddCell(cell2a);
                            PdfContentByte cb = wt.DirectContent;
                            iTextSharp.text.Image image = bc.CreateImageWithBarcode(cb, iTextSharp.text.BaseColor.BLACK, iTextSharp.text.BaseColor.BLACK);
                            PdfPCell cell2 = new PdfPCell(image);
                            cell2.Colspan = 100;
                            cell2.PaddingTop = 20f;
                            cell2.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell2.Border = Rectangle.NO_BORDER;
                            table.AddCell(cell2);

                            var textDate = FontFactory.GetFont("Arial", (float)7, Font.NORMAL);
                            var textPCB = FontFactory.GetFont("Arial", (float)7, Font.NORMAL);
                            var PCB = new Phrase("PCB : " + PCBC + Environment.NewLine, textPCB);
                            Phrase DateImpression = new Phrase("Date Impr. " + DateTime.Now.ToString("dd/MM/yyyy"), textDate);

                            PdfPCell cellPCB = new PdfPCell(PCB);
                            cellPCB.Colspan = 50;
                            cellPCB.HorizontalAlignment = 0; //0=Left, 1=Centre, 2=Right
                            cellPCB.PaddingTop = 7f;
                            cellPCB.PaddingBottom = 1f;
                            cellPCB.PaddingLeft = 10f;
                            cellPCB.Border = Rectangle.NO_BORDER;
                            table.AddCell(cellPCB);

                            PdfPCell cellDateImpression = new PdfPCell(DateImpression);
                            cellDateImpression.Colspan = 50;
                            cellDateImpression.HorizontalAlignment = 2; //0=Left, 1=Centre, 2=Right
                            cellDateImpression.PaddingTop = 7f;
                            cellDateImpression.PaddingBottom = 1f;
                            cellDateImpression.PaddingRight = 10f;
                            cellDateImpression.Border = Rectangle.NO_BORDER;
                            table.AddCell(cellDateImpression);

                            PdfPCell cell0 = new PdfPCell(table);
                            cell0.Colspan = 80;
                            cell0.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            cell0.PaddingLeft = 10f;
                            cell0.PaddingBottom = 10f;
                            cell0.PaddingTop = 10f;
                            cell0.PaddingRight = 10f;


                            cell0.FixedHeight = heightPoints;

                            table2.AddCell(cell0);
                            cpt += 1;
                            if (cpt == 2)
                            {
                                PdfPCell celllast = new PdfPCell(new Phrase(""));
                                celllast.Colspan = 25;
                                celllast.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                                celllast.PaddingLeft = 0f;
                                celllast.PaddingBottom = 0f;
                                celllast.PaddingTop = 0f;
                                celllast.PaddingRight = 0f;
                                celllast.Border = Rectangle.NO_BORDER;
                                celllast.FixedHeight = heightPoints;
                                table2.AddCell(celllast);
                                cpt = 0;
                            }
                            orec.MoveNext();
                        }

                        if (cpt > 0)
                        {
                            PdfPCell celllast = new PdfPCell(new Phrase(""));
                            celllast.Colspan = (2 - cpt) * 80 + 25;
                            celllast.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            celllast.PaddingLeft = 0f;
                            celllast.PaddingBottom = 0f;
                            celllast.PaddingTop = 0f;
                            celllast.PaddingRight = 0f;
                            celllast.Border = Rectangle.NO_BORDER;
                            celllast.FixedHeight = heightPoints;
                            table2.AddCell(celllast);
                        }
                        document.Add(table2);
                        SBO_Application.StatusBar.SetText("Opération éffectuée avec succès. Fichier sauvegardé dans l'emplacement: " + pathAndFile, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
                    }
                    catch (Exception ex)
                    {
                        SBO_Application.StatusBar.SetText("Erreur lors de la génération: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    }
                    finally
                    {
                        document.Close();
                    }
                }

                Process.Start(new ProcessStartInfo(pathAndFile) { UseShellExecute = true });
            }
            else
            {
                SBO_Application.StatusBar.SetText("Aucune donnée trouvée" + oCompany.CompanyName, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
            }
            // }
            //}
        }

        static float MillimetersToPoints(float millimeters)
        {
            return (float)Math.Round(millimeters * 2.83465f, 0);
        }

        public static void CloseAndDeletePdf(string filePath)
        {
            if (File.Exists(filePath))
            {
                try
                {
                    // Fermer tous les processus utilisant le fichier PDF
                    var processes = Process.GetProcesses()
                                           .Where(p => !string.IsNullOrEmpty(p.MainWindowTitle) &&
                                                       p.MainModule.FileName.Equals(filePath, StringComparison.OrdinalIgnoreCase));

                    foreach (var process in processes)
                    {
                        Console.WriteLine($"Fermeture du processus : {process.ProcessName} (ID: {process.Id})");
                        process.CloseMainWindow(); // Ferme gracieusement la fenêtre principale
                        process.WaitForExit(3000); // Attendre 3 secondes pour permettre au processus de se fermer
                        if (!process.HasExited)
                        {
                            process.Kill(); // Force la fermeture du processus si nécessaire
                        }
                    }
                    // S'assurer que tous les processus sont fermés
                    Thread.Sleep(1000); // Attendre un peu pour s'assurer que tous les processus sont fermés

                    // Supprimer le fichier PDF
                    File.Delete(filePath);
                    Console.WriteLine("Le fichier PDF a été supprimé avec succès.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la suppression du fichier PDF : {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Le fichier PDF n'existe pas.");
            }
        }

        static void BuildQuery(string comm,string mag, string sous_famille)
        {
            if (!string.IsNullOrEmpty(FromItem) && !string.IsNullOrEmpty(To) && !string.IsNullOrEmpty(comm) && !string.IsNullOrEmpty(mag))
            {
                query = $@"SELECT T0.[ItemCode] 'Code Article', T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB' FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                                INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum 
                                INNER JOIN OITW T3 ON T3.ItemCode = T0.ItemCode
                                WHERE T2.[ListNum] ={LisP} AND T0.[ItemCode] BETWEEN '{FromItem}' AND '{To}' and T0.[ItemName] LIKE '{comm}%'
                                AND T3.[WhsCode] = '{mag}' AND T0.[U_sous_fam] = '{sous_famille}' and T3.[OnHand] >= 0 ";
            }
            else if (!string.IsNullOrEmpty(FromItem) && !string.IsNullOrEmpty(To) && !string.IsNullOrEmpty(mag))
            {
                query = $@"SELECT T0.[ItemCode] 'Code Article', T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB' FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                                INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum 
                                INNER JOIN OITW T3 ON T3.ItemCode = T0.ItemCode
                                WHERE T2.[ListNum] ={LisP} AND T0.[U_sous_fam] = '{sous_famille}' AND T0.[ItemCode] BETWEEN '{FromItem}' AND '{To}'
                                AND T3.[WhsCode] = '{mag}' and T3.[OnHand] >= 0 ";
            }
            else if (!string.IsNullOrEmpty(FromItem) && !string.IsNullOrEmpty(To) && !string.IsNullOrEmpty(comm))
            {
                query = $@"SELECT T0.[ItemCode] 'Code Article', T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB' FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                                INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum 
                                WHERE T2.[ListNum] ={LisP} AND AND T0.[U_sous_fam] = '{sous_famille}' T0.[ItemCode] BETWEEN '{FromItem}'  AND '{To}' and T0.[ItemName] LIKE '{comm}%'
                                ";
            }
            else if (!string.IsNullOrEmpty(FromItem) && !string.IsNullOrEmpty(To))
            {
                query = $@"SELECT T0.[ItemCode] 'Code Article', T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB' FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                                INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum                               
                                WHERE T2.[ListNum] ={LisP} AND T0.[U_sous_fam] = '{sous_famille}' AND T0.[ItemCode] BETWEEN '{FromItem}' AND '{To}' ";
            }
            else if (!string.IsNullOrEmpty(comm) && !string.IsNullOrEmpty(mag))
            {
                query = $@"SELECT T0.[ItemCode] 'Code Article', T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB' FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                                INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum 
                                INNER JOIN OITW T3 ON T3.ItemCode = T0.ItemCode
                                WHERE T2.[ListNum] ={LisP} and T0.[ItemName] LIKE '{comm}%'
                                AND T3.[WhsCode] = '{mag}' AND T0.[U_sous_fam] = '{sous_famille}' and T3.[OnHand] >= 0 ";
            }
            else if (!string.IsNullOrEmpty(mag))
            {
                query = $@"SELECT T0.[ItemCode] 'Code Article', T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB' FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                                INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum 
                                INNER JOIN OITW T3 ON T3.ItemCode = T0.ItemCode
                                WHERE T2.[ListNum] ={LisP} AND T0.[U_sous_fam] = '{sous_famille}' AND T3.[WhsCode] = '{mag}' and T3.[OnHand] >= 0 ";
            }
            else if (!string.IsNullOrEmpty(comm))
            {
                query = $@"SELECT T0.[ItemCode] 'Code Article', T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB'
                                FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                                INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum                               
                                WHERE T2.[ListNum] ={LisP} AND T0.[U_sous_fam] = '{sous_famille}' and T0.[ItemName] LIKE '{comm}%'";
            }
            else
            {
                query = $@"SELECT T0.[ItemCode] 'Code Article', T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB' 
                                FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                                INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum WHERE T0.[U_sous_fam] = '{sous_famille}' and T2.[ListNum] ={LisP} ";
            }
        }

        static void BuildQueryPromo(string comm, string mag)
        {
           
            if (!string.IsNullOrEmpty(FromItem) && !string.IsNullOrEmpty(To) && !string.IsNullOrEmpty(comm) && !string.IsNullOrEmpty(mag))
            {
                query = $@"SELECT T0.[ItemCode] 'Code Article', T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB', T3.[Price] 'NEW PRICE'
                     FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                     INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum inner join SPP1 T3 ON T3.ItemCode = T0.ItemCode 
                      INNER JOIN OITW T4 ON T4.ItemCode = T0.ItemCode
                      WHERE T2.[ListNum] = {LisP}
                            AND T0.[ItemCode] BETWEEN '{FromItem}' AND '{To}' and T0.[ItemName] LIKE '{comm}%'
                                AND T4.[WhsCode] = '{mag}' and T4.[OnHand] > 0 ";
            }
            else if (!string.IsNullOrEmpty(FromItem) && !string.IsNullOrEmpty(To) && !string.IsNullOrEmpty(mag))
            {
                query = $@"SELECT T0.[ItemCode] 'Code Article', T0.[ItemName] 'Description article',T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB', T3.[Price] 'NEW PRICE'
                     FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                     INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum inner join SPP1 T3 ON T3.ItemCode = T0.ItemCode 
                      INNER JOIN OITW T4 ON T4.ItemCode = T0.ItemCode
                      WHERE T2.[ListNum] = {LisP}
                            AND T0.[ItemCode] BETWEEN '{FromItem}' AND '{To}' 
                                AND T4.[WhsCode] = '{mag}' and T4.[OnHand] > 0 ";
            }
            else if (!string.IsNullOrEmpty(FromItem) && !string.IsNullOrEmpty(To) && !string.IsNullOrEmpty(comm))
            {
                query = $@"SELECT T0.[ItemCode] 'Code Article', T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB', T3.[Price] 'NEW PRICE'
                     FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                     INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum
                    inner join SPP1 T3 ON T3.ItemCode = T0.ItemCode 
                      WHERE T2.[ListNum] = {LisP}
                            AND T0.[ItemCode] BETWEEN '{FromItem}' AND '{To}' 
                               and T0.[ItemName] LIKE '{comm}%' ";
            }
            else if (!string.IsNullOrEmpty(FromItem) && !string.IsNullOrEmpty(To))
            {
                
                query = $@"SELECT T0.[ItemCode] 'Code Article', T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB', T3.[Price] 'NEW PRICE'
                     FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                     INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum inner join SPP1 T3 ON T3.ItemCode = T0.ItemCode 
                      INNER JOIN OITW T3 ON T3.ItemCode = T0.ItemCode
                      WHERE T2.[ListNum] = {LisP} AND T0.[ItemCode] BETWEEN '{FromItem}' AND '{To}' ";
            }
            else if (!string.IsNullOrEmpty(comm) && !string.IsNullOrEmpty(mag))
            {

                query = $@"SELECT T0.[ItemCode] 'Code Article', T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB', T3.[Price] 'NEW PRICE'
                     FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                     INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum inner join SPP1 T3 ON T3.ItemCode = T0.ItemCode 
                      INNER JOIN OITW T4 ON T4.ItemCode = T0.ItemCode
                      WHERE T2.[ListNum] = {LisP} AND T0.[ItemName] LIKE '{comm}%' AND T4.[WhsCode] = '{mag}' and T4.[OnHand] > 0 ";
            }
            else if (!string.IsNullOrEmpty(mag))
            {
                query = $@"SELECT T0.[ItemCode] 'Code Article', T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB', T3.[Price] 'NEW PRICE'
                     FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                     INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum inner join SPP1 T3 ON T3.ItemCode = T0.ItemCode 
                      INNER JOIN OITW T4 ON T4.ItemCode = T0.ItemCode
                      WHERE T2.[ListNum] = {LisP} AND AND T4.[WhsCode] = '{mag}' and T4.[OnHand] > 0  ";

            }
            else if (!string.IsNullOrEmpty(comm))
            {
                query = $@"SELECT T0.[ItemCode] 'Code Article', T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB', T3.[Price] 'NEW PRICE'
                     FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode 
                     INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum inner join SPP1 T3 ON T3.ItemCode = T0.ItemCode 
                      WHERE T2.[ListNum] = {LisP} AND T0.[ItemName] LIKE '{comm}%'";
            }
            else
            {
                query = $@"SELECT T0.[ItemCode] 'Code Article', T0.[ItemName] 'Description article', T0.[CodeBars] 'Code barres', T1.[Price] 'Prix de vente', T0.[U_PCB] 'PCB', T3.[Price] 'NEW PRICE'
                     FROM OITM T0 INNER JOIN ITM1 T1 ON T0.ItemCode = T1.ItemCode
                      INNER JOIN OPLN T2 ON T1.PriceList = T2.ListNum 
                      inner join SPP1 T3 ON T3.ItemCode = T0.ItemCode  WHERE T2.[ListNum] = {LisP}";
            }
        }
    }
}

