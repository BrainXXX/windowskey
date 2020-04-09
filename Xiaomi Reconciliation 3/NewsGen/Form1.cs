using System;
using System.IO;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;


using System.Runtime.InteropServices;
using System.Threading;
using System.Globalization;
using System.Windows.Forms;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Net;
using System.Text;


namespace NewsGen
{
    public partial class Form1 : Form
    {
        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;
        Excel.Range oRng;
        object m_objOpt = System.Reflection.Missing.Value;
        int max_rows = 0;
        int max_columns = 0;

        public Form1()
        {
            InitializeComponent();
            //comboBox1.SelectedItem = "Сводная Цены";
        }



        ////////// File 1 Xiaomi

        int Service_Order_Code_Column = 1;
        int ASP_Order_NO_Column = 2;
        int Service_Station_Column = 5;
        int Model_PN_Column = 9;
        int Service_Offering_Column = 24;

        int Product_Model_Column = 21;


        int Create_Time_Column = 34;
        int Service_Order_Save_Time_Column = 35;
        int Carry_in_Time_Column = 36;
        int Apply_for_Parts_Time_Column = 37;
        int CI_Received_Time_Column = 38;
        int CP_Received_Time_Column = 39;
        int Parts_Available_in_Vendor_Time_Column = 40;
        int Parts_in_Service_Center_Time_Column = 41;
        int Service_Order_Waiting_Time_Column = 42;
        int Repair_Finish_Time_Column = 43;
        int Customer_Pick_up_Time_Column = 44;
        int System_Close_Time_Column = 45;
        int Repair_Level_Column = 46;
        int DOA_DAP_Column = 47;
        int Service_Order_Status_Column = 48;
        int Recover_Method_Column = 49;
        int Problem_Category_Column = 50;


        int Defective_Part_1_PN_Column = 68;
        int Defective_Part_1_Name_Column = 69;
        int Defective_Part_2_PN_Column = 75;
        int Defective_Part_2_Name_Column = 76;
        int Defective_Part_3_PN_Column = 84;
        int Defective_Part_3_Name_Column = 85;
        int Replacement_Part_Number_Column = 91;

        int IMEI_Column = 18;
        int SN_No_Column = 19;
        int New_IMEI_Column = 51;

        int Actived_Time_Column = 59;

        int Sale_Country_Column = 58;

        ////////// File 2 DNS

        int РемонтЗаведенКомментарий_Column = 53; //+4
        int TimeZone_Column = 58;
        int РемонтЗаведен1С_Column = 54;
        int РемонтЗакрыт1С_Column = 55;
        int ОСНТ_Номер_и_Дата_Column = 2;
        int ОСНТ_Номер_Column = 59;
        int ТоварКод_Column = 60;
        int СтатусРемонтаВендора1с_Column = 16;


        int ДатаПоступленияНаФилиалАВР_Column = 29;
        int АВР_Дата_Column = 31;
        int ДатаОтгрузкиВыдачиТовара_из_ФилиалРемонта_Column = 41;
        int ДатаПродажи_или_выявления_предторга_Column = 69;

        int ИтогиРемонтаDNS_Column = 9;
        int ДатаБезВнесенияВБазуВендора_Column = 56;
        int ДатаПринято_решение_ремонтировать_товар_на_филиале_Column = 57;
        int ДействиеОперации_Column = 3;
        int СогласованныйФилиалРемонта_Column = 34;

        int ОСНТ_Дата_Column = 22;

        int ТекущийСтатусОстатка_Column = 74;

        int ТекущийОстаток_Действие_Column = 75;
        int ТекущийОстаток_ДнейОтПоследнегоДвижения_Column = 76;

        int Событие_ЗаведенияРемонта_в_базе_вендора_СрезПоследних_Комментарий_Column = 86;
        int ГарантийныйНаОснованииПлатного_Column = 106;

        int Событие_ДиагностикаДефектВыявлен_Автор_Column = 108;
        int Событие_ДиагностикаДефектВыявлен_Комментарий_Column = 109;

        int АВР_РезультатУслуг_Column = 40;
        int Исключение_DOA_Column = 147;


        /// <summary>
        /// 
        /// </summary>


        int Товар_Column = 7;



        int КолвоКолвоЗИП_в_АВР_Column = 56;

        int Дивизион_Column = 12;

        int ЗИП1_КодПроизводителя_Column = 57;
        int ЗИП2_КодПроизводителя_Column = 62;
        int ЗИП3_КодПроизводителя_Column = 67;

        int Дата_АВР_Column = 40;

        ////////// File 3 Коды товаров

        int DNS_Code_Column = 1;
        int DNS_Name_Column = 2;
        int Xiaomi_Model_Name_Column = 4;
        int Xiaomi_Model_Code_Column = 3;


        /// <summary>
        /// /////////
        /// </summary>

        int Warehouse_Name_Column = 3;
        int Goods_PN_Column = 9;
        int Goods_Name_Column = 10;
        int Quantity_Column = 11;
        int In_Transit_Column = 12;
        int Engineer_Column = 13;

        int VendorPosCode_Column = 4;
        int FilialCode_Column = 7;
        int DNS_PosCode_Column = 3;

        int Vendor_Warehouse_Name_Column = 1;
        int DNS_FilialCode_Column = 2;



        Object[,] FileTable1 = null;
        Object[,] FileTable2 = null;
        Object[,] FileTable3 = null;




        ////////// File 1 Xiaomi

        ArrayList File1_Service_Order_Code_List = new ArrayList();
        ArrayList File1_ASP_Order_NO_List = new ArrayList();
        ArrayList File1_Service_Station_List = new ArrayList();
        ArrayList File1_Model_PN_List = new ArrayList();
        ArrayList File1_Service_Offering_List = new ArrayList();
        ArrayList File1_Product_Model_List = new ArrayList();

        ArrayList File1_Create_Time_List = new ArrayList();
        ArrayList File1_Service_Order_Save_Time_List = new ArrayList();
        ArrayList File1_Carry_in_Time_List = new ArrayList();
        ArrayList File1_Apply_for_Parts_Time_List = new ArrayList();
        ArrayList File1_CI_Received_Time_List = new ArrayList();
        ArrayList File1_CP_Received_Time_List = new ArrayList();
        ArrayList File1_Parts_Available_in_Vendor_Time_List = new ArrayList();
        ArrayList File1_Parts_in_Service_Center_Time_List = new ArrayList();
        ArrayList File1_Service_Order_Waiting_Time_List = new ArrayList();
        ArrayList File1_Repair_Finish_Time_List = new ArrayList();
        ArrayList File1_Customer_Pick_up_Time_List = new ArrayList();
        ArrayList File1_System_Close_Time_List = new ArrayList();
        ArrayList File1_Repair_Level_List = new ArrayList();
        ArrayList File1_DOA_DAP_List = new ArrayList();
        ArrayList File1_Service_Order_Status_List = new ArrayList();
        ArrayList File1_Recover_Method_List = new ArrayList();

        ArrayList File1_Defective_Part_1_PN_List = new ArrayList();
        ArrayList File1_Defective_Part_1_Name_List = new ArrayList();
        ArrayList File1_Defective_Part_2_PN_List = new ArrayList();
        ArrayList File1_Defective_Part_2_Name_List = new ArrayList();
        ArrayList File1_Defective_Part_3_PN_List = new ArrayList();
        ArrayList File1_Defective_Part_3_Name_List = new ArrayList();
        ArrayList File1_Replacement_Part_Number_List = new ArrayList();

        ArrayList File1_IMEI_List = new ArrayList();
        ArrayList File1_SN_No_List = new ArrayList();
        ArrayList File1_New_IMEI_List = new ArrayList();
        ArrayList File1_Actived_Time_List = new ArrayList();
        ArrayList File1_Sale_Country_List = new ArrayList();

        ArrayList File1_Problem_Category_List = new ArrayList();


        ////////// File 2 DNS
        ArrayList File2_Index_from_File1_List = new ArrayList();

        ArrayList File2_РемонтЗаведенКомментарий_List = new ArrayList();
        ArrayList File2_TimeZone_List = new ArrayList();
        ArrayList File2_ОСНТ_Номер_List = new ArrayList();
        ArrayList File2_РемонтЗаведен1С_List = new ArrayList();
        ArrayList File2_РемонтЗакрыт1С_List = new ArrayList();
        ArrayList File2_ОСНТ_Номер_и_Дата_List = new ArrayList();
        ArrayList File2_ТоварКод_List = new ArrayList();
        ArrayList File2_СтатусРемонтаВендора1с_List = new ArrayList();
        ArrayList File2_ДатаПоступленияНаФилиалАВР_List = new ArrayList();
        ArrayList File2_АВР_Дата_List = new ArrayList();
        ArrayList File2_ДатаОтгрузкиВыдачиТовара_из_ФилиалРемонта_List = new ArrayList();
        ArrayList File2_ДатаПродажи_или_выявления_предторга_List = new ArrayList();

        ArrayList File2_ИтогиРемонтаDNS_List = new ArrayList();
        ArrayList File2_ДатаПринято_решение_ремонтировать_товар_на_филиале_List = new ArrayList();
        ArrayList File2_ДатаБезВнесенияВБазуВендора_List = new ArrayList();
        ArrayList File2_ДействиеОперации_List = new ArrayList();
        ArrayList File2_СогласованныйФилиалРемонта_List = new ArrayList();
        ArrayList File2_ОСНТ_Дата_List = new ArrayList();
        ArrayList File2_ТекущийОстаток_Действие_List = new ArrayList();
        ArrayList File2_ТекущийОстаток_ДнейОтПоследнегоДвижения_List = new ArrayList();

        ArrayList File2_Событие_ЗаведенияРемонта_в_базе_вендора_СрезПоследних_Комментарий_List = new ArrayList();
        ArrayList File2_ГарантийныйНаОснованииПлатного_List = new ArrayList();

        ArrayList File2_Событие_ДиагностикаДефектВыявлен_Автор_List = new ArrayList();
        ArrayList File2_Событие_ДиагностикаДефектВыявлен_Комментарий_List = new ArrayList();
        ArrayList File2_ТекущийСтатусОстатка_List = new ArrayList();

        ArrayList File2_АВР_РезультатУслуг_List = new ArrayList();
        ArrayList File2_Исключение_DOA_List = new ArrayList();

        //////////////////////////



        ArrayList File2_Товар_List = new ArrayList();




        ArrayList File2_КолвоКолвоЗИП_в_АВР_List = new ArrayList();

        ArrayList File2_Дивизион_Column_List = new ArrayList();

        ArrayList File2_ЗИП1_КодПроизводителя_List = new ArrayList();
        ArrayList File2_ЗИП2_КодПроизводителя_List = new ArrayList();
        ArrayList File2_ЗИП3_КодПроизводителя_List = new ArrayList();

        ArrayList File2_Дата_АВР_List = new ArrayList();

        ////////// File 3 Коды товаров

        ArrayList File3_DNS_Code_List = new ArrayList();
        ArrayList File3_DNS_Name_List = new ArrayList();
        ArrayList File3_Xiaomi_Model_Name_List = new ArrayList();
        ArrayList File3_Xiaomi_Model_Code_List = new ArrayList();



        ArrayList NotFindList = new ArrayList();
        ArrayList NotFindListCode = new ArrayList();

        ArrayList RR_List = new ArrayList();
        ArrayList RR_Days_List = new ArrayList();
        ArrayList RR_SO_List = new ArrayList();
        ArrayList RR_Prev_Date_List = new ArrayList();
        ArrayList RR_ОСНТ_List = new ArrayList();
        ArrayList RR_Prev_ОСНТ_List = new ArrayList();

        ArrayList RR_Status_List = new ArrayList();


        //    ArrayList ResultList = new ArrayList();
        //ArrayList ModelList = new ArrayList();

        public bool LoadDataFile1(string DataFilename)
        {
            // FormObj.AddLog("Чтение из экселя");

            #region /// Чтение из экселя

            Excel.Range range;

            Object CelRet;
            ////////////////////////////////

            //    System.Globalization.CultureInfo oldCulture = System.Threading.Thread.CurrentThread.CurrentCulture;
            //    Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            ArrayList BeforeList = null, AfterList = null;

            //////////////////////////////////////
            try
            {


                BeforeList = StaticFunction.GetExcelProcessID();

                oXL = new Excel.Application();

                //////////////////
                AfterList = StaticFunction.GetExcelProcessID();

                /////////////

                oXL.Interactive = false;
                oXL.EnableEvents = false;

                oXL.ScreenUpdating = false;
                oXL.Visible = false;



                //////////////////////
                /// Прайс Конкурентов


                oWB = (Excel._Workbook)(oXL.Workbooks.Open(DataFilename, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt,
                    m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt));

                oSheet = (Excel._Worksheet)oWB.Worksheets.get_Item(1);


                //    string tmp_string, tmp_string2;





                //string NameColumnString = StaticFunction.ConvertToColumnName(NameColumn);
                // string CostColumnString = StaticFunction.ConvertToColumnName(CostColumn);


                /////////////////////////////////////////////////////////////////////////////

                string start_point, end_point;

                max_rows = oSheet.UsedRange.Rows.Count;
                max_columns = oSheet.UsedRange.Columns.Count;



                string max_columnsString = StaticFunction.ConvertToColumnName(max_columns);

                start_point = "A1";
                end_point = string.Format("{0}{1}", max_columnsString, max_rows);

                range = oSheet.get_Range(start_point, end_point);
                FileTable1 = (System.Object[,])range.get_Value(Missing.Value);




                #endregion


                for (int Ri = 1; Ri <= max_rows; Ri++)
                {
                    /// Service_Order_Code_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Service_Order_Code_Column];
                    string CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Service_Order_Code_List.Add((string)CelSring);
                    #endregion



                    /// ASP_Order_NO_Column
                      #region
                    CelRet = (System.Object)FileTable1[Ri, ASP_Order_NO_Column];
                    // string CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_ASP_Order_NO_List.Add((string)CelSring);
                    #endregion


                    /// Service_Station_Column
                      #region
                    CelRet = (System.Object)FileTable1[Ri, Service_Station_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Service_Station_List.Add((string)CelSring);
                    #endregion


                    /// Model_PN_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Model_PN_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Model_PN_List.Add((string)CelSring);
                    #endregion



                    /// Service_Offering_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Service_Offering_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Service_Offering_List.Add((string)CelSring);
                    #endregion



                    /// Create_Time_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Create_Time_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Create_Time_List.Add((string)CelSring);
                    #endregion
                    ////


                    /// Product_Model_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Product_Model_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Product_Model_List.Add((string)CelSring);
                    #endregion
                    ////


                    /// Service_Order_Save_Time_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Service_Order_Save_Time_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Service_Order_Save_Time_List.Add((string)CelSring);
                    #endregion
                    ////


                    /// Carry_in_Time_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Carry_in_Time_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Carry_in_Time_List.Add((string)CelSring);
                    #endregion
                    ////


                    /// Apply_for_Parts_Time_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Apply_for_Parts_Time_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Apply_for_Parts_Time_List.Add((string)CelSring);
                    #endregion
                    ////


                    /// CI_Received_Time_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, CI_Received_Time_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_CI_Received_Time_List.Add((string)CelSring);
                    #endregion
                    ////


                    /// CP_Received_Time_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, CP_Received_Time_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_CP_Received_Time_List.Add((string)CelSring);
                    #endregion
                    ////



                    /// Parts_Available_in_Vendor_Time_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Parts_Available_in_Vendor_Time_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Parts_Available_in_Vendor_Time_List.Add((string)CelSring);
                    #endregion
                    ////


                    /// Parts_in_Service_Center_Time_Column

                    #region
                    CelRet = (System.Object)FileTable1[Ri, Parts_in_Service_Center_Time_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Parts_in_Service_Center_Time_List.Add((string)CelSring);
                    #endregion
                    ////



                    /// Service_Order_Waiting_Time_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Service_Order_Waiting_Time_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Service_Order_Waiting_Time_List.Add((string)CelSring);
                    #endregion
                    ////


                    /// Repair_Finish_Time_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Repair_Finish_Time_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Repair_Finish_Time_List.Add((string)CelSring);
                    #endregion
                    ////


                    /// Customer_Pick_up_Time_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Customer_Pick_up_Time_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Customer_Pick_up_Time_List.Add((string)CelSring);
                    #endregion
                    ////



                    /// System_Close_Time_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, System_Close_Time_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_System_Close_Time_List.Add((string)CelSring);
                    #endregion
                    ////



                    /// Repair_Level_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Repair_Level_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Repair_Level_List.Add((string)CelSring);
                    #endregion
                    ////


                    /// DOA_DAP_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, DOA_DAP_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_DOA_DAP_List.Add((string)CelSring);
                    #endregion
                    ////


                    /// Service_Order_Status_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Service_Order_Status_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Service_Order_Status_List.Add((string)CelSring);
                    #endregion
                    ////


                    /// Recover_Method_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Recover_Method_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Recover_Method_List.Add((string)CelSring);
                    #endregion
                    ////




                    /// Defective_Part_1_PN_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Defective_Part_1_PN_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Defective_Part_1_PN_List.Add((string)CelSring);
                    #endregion
                    ////



                    /// Defective_Part_1_Name_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Defective_Part_1_Name_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Defective_Part_1_Name_List.Add((string)CelSring);
                    #endregion
                    ////


                    /// Defective_Part_2_PN_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Defective_Part_2_PN_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Defective_Part_2_PN_List.Add((string)CelSring);
                    #endregion
                    ////


                    /// Defective_Part_2_Name_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Defective_Part_2_Name_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Defective_Part_2_Name_List.Add((string)CelSring);
                    #endregion
                    ////


                    /// Defective_Part_3_PN_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Defective_Part_3_PN_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Defective_Part_3_PN_List.Add((string)CelSring);
                    #endregion
                    ////



                    /// Defective_Part_3_Name_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Defective_Part_3_Name_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Defective_Part_3_Name_List.Add((string)CelSring);
                    #endregion
                    ////

                    ///  Replacement_Part_Number_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Replacement_Part_Number_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Replacement_Part_Number_List.Add((string)CelSring);
                    #endregion
                    ////


                    ///  IMEI_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, IMEI_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_IMEI_List.Add((string)CelSring);
                    #endregion
                    ////


                    ///  SN_No_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, SN_No_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_SN_No_List.Add((string)CelSring);
                    #endregion
                    ////

                    ///  New_IMEI_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, New_IMEI_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_New_IMEI_List.Add((string)CelSring);
                    #endregion
                    ////


                    ///  Actived_Time_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Actived_Time_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Actived_Time_List.Add((string)CelSring);
                    #endregion
                    ////


                    /// Problem_Category_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Problem_Category_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Problem_Category_List.Add((string)CelSring);
                    #endregion


                    /// Sale_Country_Column
                    #region
                    CelRet = (System.Object)FileTable1[Ri, Sale_Country_Column];


                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    //  CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File1_Sale_Country_List.Add((string)CelSring);
                    #endregion
                }


                ////////////////////////////////////////////////////////////////////////

                oWB.Close(true, m_objOpt, m_objOpt);
                oXL.Quit();


                Marshal.ReleaseComObject(oSheet);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oXL);
                oSheet = null;
                oWB = null;
                oXL = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //   System.Threading.Thread.CurrentThread.CurrentCulture = oldCulture;
            }

            StaticFunction.ClearExcelProcess(BeforeList, AfterList);

            ///////////////////////



            return true;
        }


        public bool LoadDataFile2(string DataFilename)
        {
            // FormObj.AddLog("Чтение из экселя");

            #region /// Чтение из экселя

            Excel.Range range;

            Object CelRet;
            ////////////////////////////////

            System.Globalization.CultureInfo oldCulture = System.Threading.Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            ArrayList BeforeList = null, AfterList = null;

            //////////////////////////////////////
            try
            {


                BeforeList = StaticFunction.GetExcelProcessID();

                oXL = new Excel.Application();

                //////////////////
                AfterList = StaticFunction.GetExcelProcessID();

                /////////////

                oXL.Interactive = false;
                oXL.EnableEvents = false;

                oXL.ScreenUpdating = false;
                oXL.Visible = false;





                oWB = (Excel._Workbook)(oXL.Workbooks.Open(DataFilename, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt,
                    m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt));

                oSheet = (Excel._Worksheet)oWB.Worksheets.get_Item(1);


                //    string tmp_string, tmp_string2;





                //string NameColumnString = StaticFunction.ConvertToColumnName(NameColumn);
                // string CostColumnString = StaticFunction.ConvertToColumnName(CostColumn);


                /////////////////////////////////////////////////////////////////////////////

                string start_point, end_point;

                max_rows = oSheet.UsedRange.Rows.Count;
                max_columns = oSheet.UsedRange.Columns.Count;



                string max_columnsString = StaticFunction.ConvertToColumnName(max_columns);

                start_point = "A1";
                end_point = string.Format("{0}{1}", max_columnsString, max_rows);

                range = oSheet.get_Range(start_point, end_point);
                FileTable2 = (System.Object[,])range.get_Value(Missing.Value);




                #endregion

                for (int Ri = 1; Ri <= max_rows; Ri++)
                {

                    File2_Index_from_File1_List.Add((int)-1);

                    ///  РемонтЗаведенКомментарий_Column
                    CelRet = (System.Object)FileTable2[Ri, РемонтЗаведенКомментарий_Column];
                    string CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_РемонтЗаведенКомментарий_List.Add((string)CelSring);



                    ///  TimeZone_Column
                    CelRet = (System.Object)FileTable2[Ri, TimeZone_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_TimeZone_List.Add((string)CelSring);


                    ///  РемонтЗаведен1С_Column
                    CelRet = (System.Object)FileTable2[Ri, РемонтЗаведен1С_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_РемонтЗаведен1С_List.Add((string)CelSring);


                    ///  РемонтЗакрыт1С_Column
                    CelRet = (System.Object)FileTable2[Ri, РемонтЗакрыт1С_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_РемонтЗакрыт1С_List.Add((string)CelSring);


                    ///  ОСНТ_Номер_и_Дата_Column
                    CelRet = (System.Object)FileTable2[Ri, ОСНТ_Номер_и_Дата_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_ОСНТ_Номер_и_Дата_List.Add((string)CelSring);


                    ///  ОСНТ_Номер_Column
                    CelRet = (System.Object)FileTable2[Ri, ОСНТ_Номер_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_ОСНТ_Номер_List.Add((string)CelSring);


                    ///  ТоварКод_Column
                    CelRet = (System.Object)FileTable2[Ri, ТоварКод_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_ТоварКод_List.Add((string)CelSring);


                    ///  СтатусРемонтаВендора1с_Column
                    CelRet = (System.Object)FileTable2[Ri, СтатусРемонтаВендора1с_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_СтатусРемонтаВендора1с_List.Add((string)CelSring);




                    ///  ДатаПоступленияНаФилиалАВР

                    CelRet = (System.Object)FileTable2[Ri, ДатаПоступленияНаФилиалАВР_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_ДатаПоступленияНаФилиалАВР_List.Add((string)CelSring);



                    ///  АВР_Дата_Column
                    CelRet = (System.Object)FileTable2[Ri, АВР_Дата_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_АВР_Дата_List.Add((string)CelSring);


                    ///  ДатаОтгрузкиВыдачиТовара_из_ФилиалРемонта_Column
                    CelRet = (System.Object)FileTable2[Ri, ДатаОтгрузкиВыдачиТовара_из_ФилиалРемонта_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_ДатаОтгрузкиВыдачиТовара_из_ФилиалРемонта_List.Add((string)CelSring);


                    ///  ДатаПродажи_или_выявления_предторга_Column
                    CelRet = (System.Object)FileTable2[Ri, ДатаПродажи_или_выявления_предторга_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_ДатаПродажи_или_выявления_предторга_List.Add((string)CelSring);


                    ///  ИтогиРемонтаDNS_Column
                    CelRet = (System.Object)FileTable2[Ri, ИтогиРемонтаDNS_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_ИтогиРемонтаDNS_List.Add((string)CelSring);

                    ///  ДатаПринято_решение_ремонтировать_товар_на_филиале_Column
                    CelRet = (System.Object)FileTable2[Ri, ДатаПринято_решение_ремонтировать_товар_на_филиале_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_ДатаПринято_решение_ремонтировать_товар_на_филиале_List.Add((string)CelSring);


                    ///  ДатаБезВнесенияВБазуВендора_Column
                    CelRet = (System.Object)FileTable2[Ri, ДатаБезВнесенияВБазуВендора_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_ДатаБезВнесенияВБазуВендора_List.Add((string)CelSring);


                    ///  ДействиеОперации_Column
                    CelRet = (System.Object)FileTable2[Ri, ДействиеОперации_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_ДействиеОперации_List.Add((string)CelSring);

                    ///  СогласованныйФилиалРемонта_Column

                    CelRet = (System.Object)FileTable2[Ri, СогласованныйФилиалРемонта_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_СогласованныйФилиалРемонта_List.Add((string)CelSring);



                    ///  ОСНТ_Дата_Column
                    CelRet = (System.Object)FileTable2[Ri, ОСНТ_Дата_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_ОСНТ_Дата_List.Add((string)CelSring);





                    ///  ТекущийОстаток_Действие_Column
                    CelRet = (System.Object)FileTable2[Ri, ТекущийОстаток_Действие_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_ТекущийОстаток_Действие_List.Add((string)CelSring);



                    ///  ТекущийОстаток_ДнейОтПоследнегоДвижения_Column
                    CelRet = (System.Object)FileTable2[Ri, ТекущийОстаток_ДнейОтПоследнегоДвижения_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_ТекущийОстаток_ДнейОтПоследнегоДвижения_List.Add((string)CelSring);





                    ///  Событие_ЗаведенияРемонта_в_базе_вендора_СрезПоследних_Комментарий_Column

                    CelRet = (System.Object)FileTable2[Ri, Событие_ЗаведенияРемонта_в_базе_вендора_СрезПоследних_Комментарий_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_Событие_ЗаведенияРемонта_в_базе_вендора_СрезПоследних_Комментарий_List.Add((string)CelSring);


                    ///  ГарантийныйНаОснованииПлатного_Column

                    CelRet = (System.Object)FileTable2[Ri, ГарантийныйНаОснованииПлатного_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_ГарантийныйНаОснованииПлатного_List.Add((string)CelSring);


                    ///  Событие_ДиагностикаДефектВыявлен_Автор_Column


                    CelRet = (System.Object)FileTable2[Ri, Событие_ДиагностикаДефектВыявлен_Автор_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_Событие_ДиагностикаДефектВыявлен_Автор_List.Add((string)CelSring);



                    ///  Событие_ДиагностикаДефектВыявлен_Комментарий_Column
                    CelRet = (System.Object)FileTable2[Ri, Событие_ДиагностикаДефектВыявлен_Комментарий_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_Событие_ДиагностикаДефектВыявлен_Комментарий_List.Add((string)CelSring);


                    ///  АВР_РезультатУслуг_Column

                    CelRet = (System.Object)FileTable2[Ri, АВР_РезультатУслуг_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_АВР_РезультатУслуг_List.Add((string)CelSring);


                    ///  ТекущийСтатусОстатка_Column
                    CelRet = (System.Object)FileTable2[Ri, ТекущийСтатусОстатка_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_ТекущийСтатусОстатка_List.Add((string)CelSring);



                    ///  Исключение_DOA_Column
                    CelRet = (System.Object)FileTable2[Ri, Исключение_DOA_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File2_Исключение_DOA_List.Add((string)CelSring);


                    /*
                            ///  ЗИП1_КодПроизводителя_Column
                            CelRet = (System.Object)FileTable2[Ri, ЗИП1_КодПроизводителя_Column];
                            CelSring = "";

                            if ((CelRet) != null)
                            {
                                CelSring = CelRet.ToString();
                            }
                            else
                            {
                                CelSring = "";
                            }

                            CelSring = CelSring.Replace("'", "");
                            CelSring = CelSring.Trim();

                            File2_ЗИП1_КодПроизводителя_List.Add((string)CelSring);



                            ///  ЗИП2_КодПроизводителя_Column
                            CelRet = (System.Object)FileTable2[Ri, ЗИП2_КодПроизводителя_Column];
                            CelSring = "";

                            if ((CelRet) != null)
                            {
                                CelSring = CelRet.ToString();
                            }
                            else
                            {
                                CelSring = "";
                            }

                            CelSring = CelSring.Replace("'", "");
                            CelSring = CelSring.Trim();

                            File2_ЗИП2_КодПроизводителя_List.Add((string)CelSring);

                            ///  ЗИП3_КодПроизводителя_Column
                            CelRet = (System.Object)FileTable2[Ri, ЗИП3_КодПроизводителя_Column];
                            CelSring = "";

                            if ((CelRet) != null)
                            {
                                CelSring = CelRet.ToString();
                            }
                            else
                            {
                                CelSring = "";
                            }

                            CelSring = CelSring.Replace("'", "");
                            CelSring = CelSring.Trim();

                            File2_ЗИП3_КодПроизводителя_List.Add((string)CelSring);




                            ///  КолвоКолвоЗИП_в_АВР_Column
                            CelRet = (System.Object)FileTable2[Ri, КолвоКолвоЗИП_в_АВР_Column];
                            CelSring = "";

                            if ((CelRet) != null)
                            {
                                CelSring = CelRet.ToString();
                            }
                            else
                            {
                                CelSring = "";
                            }

                            CelSring = CelSring.Replace("'", "");
                            CelSring = CelSring.Trim();

                            File2_КолвоКолвоЗИП_в_АВР_List.Add((string)CelSring);




                            ///  Дата_АВР_Column
                            CelRet = (System.Object)FileTable2[Ri, Дата_АВР_Column];
                            CelSring = "";

                            if ((CelRet) != null)
                            {
                                CelSring = CelRet.ToString();
                            }
                            else
                            {
                                CelSring = "";
                            }

                            CelSring = CelSring.Replace("'", "");
                            CelSring = CelSring.Trim();

                            File2_Дата_АВР_List.Add((string)CelSring);
                            */
                }


                ////////////////////////////////////////////////////////////////////////

                oWB.Close(true, m_objOpt, m_objOpt);
                oXL.Quit();


                Marshal.ReleaseComObject(oSheet);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oXL);
                oSheet = null;
                oWB = null;
                oXL = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = oldCulture;
            }

            StaticFunction.ClearExcelProcess(BeforeList, AfterList);

            ///////////////////////



            return true;
        }

        public bool LoadDataFile3(string DataFilename)
        {
            // FormObj.AddLog("Чтение из экселя");

            #region /// Чтение из экселя

            Excel.Range range;

            Object CelRet;
            ////////////////////////////////

            System.Globalization.CultureInfo oldCulture = System.Threading.Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            ArrayList BeforeList = null, AfterList = null;

            //////////////////////////////////////
            try
            {


                BeforeList = StaticFunction.GetExcelProcessID();

                oXL = new Excel.Application();

                //////////////////
                AfterList = StaticFunction.GetExcelProcessID();

                /////////////

                oXL.Interactive = false;
                oXL.EnableEvents = false;

                oXL.ScreenUpdating = false;
                oXL.Visible = false;



                //////////////////////
                /// Прайс Конкурентов


                oWB = (Excel._Workbook)(oXL.Workbooks.Open(DataFilename, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt,
                    m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt));

                oSheet = (Excel._Worksheet)oWB.Worksheets.get_Item(1);


                //    string tmp_string, tmp_string2;





                //string NameColumnString = StaticFunction.ConvertToColumnName(NameColumn);
                // string CostColumnString = StaticFunction.ConvertToColumnName(CostColumn);


                /////////////////////////////////////////////////////////////////////////////

                string start_point, end_point;

                max_rows = oSheet.UsedRange.Rows.Count;
                max_columns = oSheet.UsedRange.Columns.Count;



                string max_columnsString = StaticFunction.ConvertToColumnName(max_columns);

                start_point = "A1";
                end_point = string.Format("{0}{1}", max_columnsString, max_rows);

                range = oSheet.get_Range(start_point, end_point);
                FileTable3 = (System.Object[,])range.get_Value(Missing.Value);




                #endregion

                for (int Ri = 1; Ri <= max_rows; Ri++)
                {



                    ///  DNS_Code_Column
                    CelRet = (System.Object)FileTable3[Ri, DNS_Code_Column];
                    string CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File3_DNS_Code_List.Add((string)CelSring);


                    ///  DNS_Name_Column
                    CelRet = (System.Object)FileTable3[Ri, DNS_Name_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File3_DNS_Name_List.Add((string)CelSring);



                    ///  Xiaomi_Model_Name_Column
                    CelRet = (System.Object)FileTable3[Ri, Xiaomi_Model_Name_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File3_Xiaomi_Model_Name_List.Add((string)CelSring);


                    ///  Xiaomi_Model_Code_Column
                    CelRet = (System.Object)FileTable3[Ri, Xiaomi_Model_Code_Column];
                    CelSring = "";

                    if ((CelRet) != null)
                    {
                        CelSring = CelRet.ToString();
                    }
                    else
                    {
                        CelSring = "";
                    }

                    CelSring = CelSring.Replace("'", "");
                    CelSring = CelSring.Trim();

                    File3_Xiaomi_Model_Code_List.Add((string)CelSring);

                }

                ////////////////////////////////////////////////////////////////////////

                oWB.Close(true, m_objOpt, m_objOpt);
                oXL.Quit();


                Marshal.ReleaseComObject(oSheet);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oXL);
                oSheet = null;
                oWB = null;
                oXL = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = oldCulture;
            }

            StaticFunction.ClearExcelProcess(BeforeList, AfterList);

            ///////////////////////



            return true;
        }



        private void button1_Click(object sender, EventArgs e)
        {

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {



                Object[,] FileTable1 = null;
                Object[,] FileTable2 = null;
                Object[,] FileTable3 = null;



                ////////// File 1 Xiaomi


                File1_Service_Order_Code_List.Clear();
                File1_ASP_Order_NO_List.Clear();
                File1_Service_Station_List.Clear();
                File1_Model_PN_List.Clear();
                File1_Service_Offering_List.Clear();
                File1_Product_Model_List.Clear();



                File1_Create_Time_List.Clear();
                File1_Service_Order_Save_Time_List.Clear();
                File1_Carry_in_Time_List.Clear();
                File1_Apply_for_Parts_Time_List.Clear();
                File1_CI_Received_Time_List.Clear();
                File1_CP_Received_Time_List.Clear();
                File1_Parts_Available_in_Vendor_Time_List.Clear();
                File1_Parts_in_Service_Center_Time_List.Clear();
                File1_Service_Order_Waiting_Time_List.Clear();
                File1_Repair_Finish_Time_List.Clear();
                File1_Customer_Pick_up_Time_List.Clear();
                File1_System_Close_Time_List.Clear();
                File1_Repair_Level_List.Clear();
                File1_DOA_DAP_List.Clear();
                File1_Service_Order_Status_List.Clear();
                File1_Recover_Method_List.Clear();

                File1_Defective_Part_1_PN_List.Clear();
                File1_Defective_Part_1_Name_List.Clear();
                File1_Defective_Part_2_PN_List.Clear();
                File1_Defective_Part_2_Name_List.Clear();
                File1_Defective_Part_3_PN_List.Clear();
                File1_Defective_Part_3_Name_List.Clear();
                File1_Replacement_Part_Number_List.Clear();

                File1_IMEI_List.Clear();
                File1_SN_No_List.Clear();
                File1_New_IMEI_List.Clear();
                File1_Problem_Category_List.Clear();
                File1_Sale_Country_List.Clear();

                ////////// File 2 DNS
                File2_Index_from_File1_List.Clear();
                File2_РемонтЗаведен1С_List.Clear();
                File2_РемонтЗакрыт1С_List.Clear();

                File2_ОСНТ_Номер_List.Clear();
                File2_ТоварКод_List.Clear();
                File2_Товар_List.Clear();
                File2_РемонтЗаведенКомментарий_List.Clear();
                File2_КолвоКолвоЗИП_в_АВР_List.Clear();
                File2_ОСНТ_Номер_и_Дата_List.Clear();
                File2_Дивизион_Column_List.Clear();

                File2_ЗИП1_КодПроизводителя_List.Clear();
                File2_ЗИП2_КодПроизводителя_List.Clear();
                File2_ЗИП3_КодПроизводителя_List.Clear();

                File2_Дата_АВР_List.Clear();
                File2_ДатаПоступленияНаФилиалАВР_List.Clear();
                File2_АВР_Дата_List.Clear();
                File2_ДатаОтгрузкиВыдачиТовара_из_ФилиалРемонта_List.Clear();
                File2_ДатаПродажи_или_выявления_предторга_List.Clear();

                File2_ИтогиРемонтаDNS_List.Clear();
                File2_ДатаПринято_решение_ремонтировать_товар_на_филиале_List.Clear();
                File2_ДатаБезВнесенияВБазуВендора_List.Clear();
                File2_ДействиеОперации_List.Clear();
                File2_СогласованныйФилиалРемонта_List.Clear();
                File2_ТекущийОстаток_Действие_List.Clear();
                File2_ТекущийОстаток_ДнейОтПоследнегоДвижения_List.Clear();

                File2_Событие_ЗаведенияРемонта_в_базе_вендора_СрезПоследних_Комментарий_List.Clear();
                File2_ГарантийныйНаОснованииПлатного_List.Clear();
                File2_СтатусРемонтаВендора1с_List.Clear();

                File2_Событие_ДиагностикаДефектВыявлен_Автор_List.Clear();
                File2_Событие_ДиагностикаДефектВыявлен_Комментарий_List.Clear();
                File2_АВР_РезультатУслуг_List.Clear();
                File2_ТекущийСтатусОстатка_List.Clear();

                ////////// File 3 Коды товаров

                File3_DNS_Code_List.Clear();
                File3_DNS_Name_List.Clear();
                File3_Xiaomi_Model_Name_List.Clear();
                File3_Xiaomi_Model_Code_List.Clear();

                /////////////////////////


                NotFindList.Clear();
                NotFindListCode.Clear();

                RR_List.Clear();
                RR_Days_List.Clear();
                RR_SO_List.Clear();
                RR_Prev_Date_List.Clear();

                richTextBox1.Clear();
                //ResultList.Clear();
                //ModelList.Clear();

                /*
                string DataString = "17.12.2018 13:56:55";
                DateTime UDate = DateTime.ParseExact(DataString, "dd.MM.yyyy HH:mm:ss", null);


                int firstDayOfYear = (int)new DateTime(UDate.Year, 1, 1).DayOfWeek;
                double weekNumber = (UDate.DayOfYear + firstDayOfYear) / 7.0 + 1.0;

                double res = Math.Ceiling((UDate.DayOfYear + firstDayOfYear-1) / 7.0);
                */

                Thread ScanThread = new Thread(new ThreadStart(RunProg));
                ScanThread.Start();





            }

        }

        public double NumOfYear(DateTime UDate)
        {
            int firstDayOfYear = (int)new DateTime(UDate.Year, 1, 1).DayOfWeek;          

            double res = Math.Ceiling((UDate.DayOfYear + firstDayOfYear - 1) / 7.0);


            return res;
        }


        public void RunProg()
        {
            button2.Enabled = false;
            richTextBox1.Text = "Загрузка данных";

            
            LoadDataFile1(textBox1.Text);
            LoadDataFile2(textBox2.Text);
            LoadDataFile3(textBox3.Text);

            richTextBox1.Text = richTextBox1.Text + "\n" + "Проверка файла";

            if (CheckFiles()!=true)
            {
                richTextBox1.Text = richTextBox1.Text + "\n" + "Ошибка формата файла";
                return;

            }


            richTextBox1.Text = richTextBox1.Text + "\n" + "Обработка";



            //GetData();


            UpdateExcel1(textBox2.Text);
            UpdateExcel3(textBox1.Text);

            //  CreateRR_File(@"C:\Тест5\RR_List.xls");
            richTextBox1.Text = richTextBox1.Text + "\n" + string.Format("Нераспознанно {0} позиций", NotFindList.Count);

            richTextBox1.Text = richTextBox1.Text + "\n" + "Завершено";


        }

        public bool CheckFiles()
        {
            /// Проверка заговков в нескольких точках

            string Service_Order_Code = (string)File1_Service_Order_Code_List[0];
            string Create_Time = (string)File1_Create_Time_List[0];
            string Actived_Time = (string)File1_Actived_Time_List[0];
            string Replacement_Part_Number = (string)File1_Replacement_Part_Number_List[0];

            if ((Service_Order_Code != "Service Order Code") || (Create_Time != "Create Time") || (Actived_Time != "Actived Time") || (Replacement_Part_Number != "Replacement Part Number"))
            {
                richTextBox1.Text = richTextBox1.Text + "\n" + "Изменилось расположение данных в файле Xiaomi";
                return false;
            }

            try
            {
                string Create_Time_string = (string)File1_Create_Time_List[1];
                DateTime Create_Time_Date = DateTime.ParseExact(Create_Time_string, "dd.MM.yyyy HH:mm:ss", null);
                DateTime CurDT = DateTime.Now;


                double ДнейРасхождения = (CurDT - Create_Time_Date).TotalDays;

                if ((ДнейРасхождения > 14) || (ДнейРасхождения < -14))
                {
                    richTextBox1.Text = richTextBox1.Text + "\n" + "Ошибка даты  Create Time, возможно неправильная сортировка файла";
                    return false;

                }

            }
            catch
            {
                richTextBox1.Text = richTextBox1.Text + "\n" + "Ошибка формата даты";
                return false;

            }



            return true;
        }


        public void UpdateExcel3(string filename)
        {

            System.DateTime now_date;
            now_date = System.DateTime.Now;

            string tmp = now_date.ToShortDateString() + " " + now_date.ToLongTimeString();

            //   string savename = @"C:\Обратотка\Ненайденные позиции.xls";


            ////////////////////////////////

           System.Globalization.CultureInfo oldCulture = System.Threading.Thread.CurrentThread.CurrentCulture;
           Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            ArrayList BeforeList, AfterList;

            //////////////////////////////////////
            try
            {

                ///////////////////

                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;
                object m_objOpt = System.Reflection.Missing.Value;





                BeforeList = StaticFunction.GetExcelProcessID();

                oXL = new Excel.Application();

                //////////////////
                AfterList = StaticFunction.GetExcelProcessID();

                /////////////

                oXL.Interactive = false;
                oXL.EnableEvents = false;

                oXL.ScreenUpdating = false;
                oXL.Visible = false;




                oWB = (Excel._Workbook)(oXL.Workbooks.Open(filename, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt,
                    m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt));

                oSheet = (Excel._Worksheet)oWB.Worksheets.get_Item(1);


                ///////////////////////////////

              

               

                // подсвечиваем ненайденные

                for (int n = 0; n < NotFindList.Count; n++)
                {

                  
                    int NotFindLine = (int)NotFindList[n];

                    string sCell11 = string.Format("A{0}", NotFindLine + 1);
                    string sCell22 = string.Format("B{0}", NotFindLine + 1);

                                       
                    oSheet.get_Range(sCell11, sCell22).Interior.ColorIndex = 46;
                    







                }
                /*
                for (int n = 0; n < RR_List.Count; n++)
                {
                    int RR_Line = (int)RR_List[n];

                    string sCell11 = string.Format("C{0}", RR_Line + 1);
                    string sCell22 = string.Format("S{0}", RR_Line + 1);
                    

                    oSheet.get_Range(sCell11, sCell22).Interior.ColorIndex = 46;

                    oSheet.Cells[RR_Line + 1, 123].Formula = (string)RR_Days_List[n];
                    oSheet.Cells[RR_Line + 1, 124].Formula = (string)RR_SO_List[n];
                    oSheet.Cells[RR_Line + 1, 125].Formula = (string)RR_Prev_Date_List[n];

                   
                }

                */
                oWB.Save();


                //oWB.SaveAs(savename, Excel.XlFileFormat.xlWorkbookNormal, m_objOpt,  m_objOpt, m_objOpt, m_objOpt, Excel.XlSaveAsAccessMode.xlExclusive,   m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);


                oWB.Close(false, m_objOpt, m_objOpt);



                oXL.Quit();


                Marshal.ReleaseComObject(oSheet);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oXL);
                oSheet = null;
                oWB = null;
                oXL = null;
                GC.Collect();

            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = oldCulture;
            }

            StaticFunction.ClearExcelProcess(BeforeList, AfterList);



        }



        public void UpdateExcel1(string filename)
        {

            System.DateTime now_date;
            now_date = System.DateTime.Now;

            string tmp = now_date.ToShortDateString() + " " + now_date.ToLongTimeString();

            //   string savename = @"C:\Обратотка\Ненайденные позиции.xls";


            ////////////////////////////////

            System.Globalization.CultureInfo oldCulture = System.Threading.Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            ArrayList BeforeList, AfterList;

            //////////////////////////////////////
            try
            {

                ///////////////////

                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;
                object m_objOpt = System.Reflection.Missing.Value;





                BeforeList = StaticFunction.GetExcelProcessID();

                oXL = new Excel.Application();

                //////////////////
                AfterList = StaticFunction.GetExcelProcessID();

                /////////////

                oXL.Interactive = false;
                oXL.EnableEvents = false;

                oXL.ScreenUpdating = false;
                oXL.Visible = false;




                oWB = (Excel._Workbook)(oXL.Workbooks.Open(filename, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt,
                    m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt));

                oSheet = (Excel._Worksheet)oWB.Worksheets.get_Item(1);


                //       oSheet.Columns[79].NumberFormat = "@";
                //     oSheet.Columns[80].NumberFormat = "@";
                //   oSheet.Columns[81].NumberFormat = "@";

                //   string s1 = (string)oSheet.Columns[34].NumberFormat;
                //  oSheet.Columns[82].NumberFormat = "0000";
                //  oSheet.Columns[83].NumberFormat = "0000";

                //     string n22= oSheet.Columns[83].NumberFormat.ToString();


                // for (int x = 1; x < File1_ASP_Order_NO_List.Count; x++)

                string TextBoxString = richTextBox1.Text;

                for (int x = File1_ASP_Order_NO_List.Count-1; x >=1; x--)
                {

                    richTextBox1.Text = TextBoxString+ "\n" + string.Format("Строка {0} из {1}", File1_ASP_Order_NO_List.Count-x, File1_ASP_Order_NO_List.Count-1);

                    string ASP_Order_NO = (string)File1_ASP_Order_NO_List[x];
                    string Service_Order_Cod = (string)File1_Service_Order_Code_List[x];

                    //    string s1 = (string)oSheet.Cells[x + 1, 34].NumberFormat;

                    ////////////  Поиск повторных обращений
                    #region
                    string IMEI_string = (string)File1_IMEI_List[x];


                    DateTime RR_DT = DateTime.Now.Date;
                    bool RR_Found = false;
                    double RR_Days_double = 0;
                    int RR_Index = 0;

                    for (int f = x + 1; f < File1_IMEI_List.Count; f++)
                    {
                        string IMEI_Temp_string1 = (string)File1_IMEI_List[f];
                        string IMEI_Temp_string2 = (string)File1_New_IMEI_List[f];

                        if ((IMEI_Temp_string1.Contains(IMEI_string)) || (IMEI_string.Contains(IMEI_Temp_string1)) || ((IMEI_Temp_string2.Contains(IMEI_string)) || (IMEI_string.Contains(IMEI_Temp_string2)) && (IMEI_Temp_string2.Trim() != "")))
                        {
                            RR_Found = true;
                            RR_Index = f;

                            string System_Close_Time = (string)File1_System_Close_Time_List[f];

                            string PrevDate = "";

                            if (System_Close_Time.Trim() != "")
                            {
                                PrevDate = System_Close_Time;

                            }
                            else
                            {
                                PrevDate = (string)File1_Create_Time_List[f];
                            }

                            string CurDate = (string)File1_Create_Time_List[x];

                            //RR_DT = DateTime.ParseExact(PrevDate, "yyyy-MM-dd HH:mm:ss", null);
                            ///DateTime CurDate_DT = DateTime.ParseExact(CurDate, "yyyy-MM-dd HH:mm:ss", null);

                            RR_DT = DateTime.ParseExact(PrevDate, "dd.MM.yyyy HH:mm:ss", null);
                            DateTime CurDate_DT = DateTime.ParseExact(CurDate, "dd.MM.yyyy HH:mm:ss", null);



                            RR_Days_double = (CurDate_DT - RR_DT).TotalDays;





                            RR_List.Add(x);
                            RR_Days_List.Add(RR_Days_double.ToString());
                            RR_SO_List.Add((string)File1_Service_Order_Code_List[RR_Index]);
                            RR_Prev_Date_List.Add(PrevDate);
                            RR_Status_List.Add((string)File1_Service_Order_Status_List[RR_Index]);

                            RR_ОСНТ_List.Add("");
                            RR_Prev_ОСНТ_List.Add("");


                            break;
                        }




                    }

                    #endregion

                    bool ASP_Order_NO_Found = false;

                    #region /// Поиск по файлу DNS

                    int d = -1;
                    for (int d_n = 1; d_n < File2_ОСНТ_Номер_List.Count; d_n++)
                    {
                        string ОСНТ_Номер = (string)File2_ОСНТ_Номер_List[d_n];

                        if (ASP_Order_NO.Contains(ОСНТ_Номер.Trim()))
                        {
                            ASP_Order_NO_Found = true;
                            d = d_n;
                            break;
                        }
                    }
                  


                    if (!ASP_Order_NO_Found)
                    {
                        for (int d_so = 1; d_so < File2_РемонтЗаведенКомментарий_List.Count; d_so++)
                        {
                            string РемонтЗаведенКомментарий = (string)File2_РемонтЗаведенКомментарий_List[d_so];
                            if (РемонтЗаведенКомментарий.Contains(Service_Order_Cod))
                            {
                                ASP_Order_NO_Found = true;
                                d = d_so;
                                break;
                            }


                        }


                    }

                    if (!ASP_Order_NO_Found)
                    {
                        for (int d_so = 1; d_so < File2_Событие_ЗаведенияРемонта_в_базе_вендора_СрезПоследних_Комментарий_List.Count; d_so++)
                        {
                            string РемонтЗаведенКомментарий = (string)File2_Событие_ЗаведенияРемонта_в_базе_вендора_СрезПоследних_Комментарий_List[d_so];
                            if (РемонтЗаведенКомментарий.Contains(Service_Order_Cod))
                            {
                                ASP_Order_NO_Found = true;
                                d = d_so;
                                break;
                            }


                        }


                    }




                    // for (int d = 1; d < File2_ОСНТ_Номер_List.Count; d++)
                    if (d > 0)
                    {

                        File2_Index_from_File1_List[d] = x;

                        string Pos_Name = (string)File2_ТоварКод_List[d];


                        string Product_Model = (string)File1_Product_Model_List[x];
                        string TimeZone_string = (string)File2_TimeZone_List[d];

                        oSheet.Cells[d + 1, 17].Formula = (string)File1_Service_Order_Status_List[x];
                        //// Статусы ремонта


                        if ((string)File2_СтатусРемонтаВендора1с_List[d] == "")
                        {
                            /// Если один из статусов пустой
                           

                            string sCell11 = string.Format("P{0}", d + 1);
                            string sCell22 = string.Format("Q{0}", d + 1);


                            oSheet.get_Range(sCell11, sCell22).Interior.ColorIndex = 46;


                        }



                        oSheet.Cells[d + 1, 111].Formula = ASP_Order_NO;
                        oSheet.Cells[d + 1, 112].Formula = (string)File1_Model_PN_List[x];
                        oSheet.Cells[d + 1, 7].Formula = Product_Model;
                        oSheet.Cells[d + 1, 133].Formula = (string)File1_IMEI_List[x];
                        oSheet.Cells[d + 1, 135].Formula = (string)File1_Replacement_Part_Number_List[x];
                        oSheet.Cells[d + 1, 136].Formula = (string)File1_Defective_Part_1_Name_List[x];
                        oSheet.Cells[d + 1, 137].Formula = (string)File1_Defective_Part_2_Name_List[x];
                        oSheet.Cells[d + 1, 138].Formula = (string)File1_Defective_Part_3_Name_List[x];


                        oSheet.Cells[d + 1, 141].Formula = (string)File1_DOA_DAP_List[x];
                        oSheet.Cells[d + 1, 142].Formula = (string)File1_Recover_Method_List[x];
                        oSheet.Cells[d + 1, 143].Formula = (string)File1_Problem_Category_List[x];
                        oSheet.Cells[d + 1, 146].Formula = (string)File1_Sale_Country_List[x];

                        #region   Обработка времени Create_Time_Date


                        string Create_Time = (string)File1_Create_Time_List[x];
                        string РемонтЗаведен1С = (string)File2_РемонтЗаведен1С_List[d];
                        string ДатаПоступленияНаФилиалАВР = (string)File2_ДатаПоступленияНаФилиалАВР_List[d];
                        string АВР_Дата = (string)File2_АВР_Дата_List[d];
                        string ДатаОтгрузкиВыдачиТовара_из_ФилиалРемонта = (string)File2_ДатаОтгрузкиВыдачиТовара_из_ФилиалРемонта_List[d];


                        DateTime Create_Time_Date = DateTime.ParseExact(Create_Time, "dd.MM.yyyy HH:mm:ss", null);
                        DateTime Parts_in_Service_Center_Time_DT = DateTime.Now;
                        DateTime Apply_for_Parts_Time_DT = DateTime.Now;
                        DateTime Service_Order_Waiting_Time = DateTime.Now;
                        DateTime Carry_in_Time = DateTime.Now;
                        DateTime ДатаПоступленияНаФилиалАВР_Date = DateTime.Now;
                        DateTime Repair_Finish_Time = DateTime.Now;
                        oSheet.Cells[d + 1, 113].Formula = Create_Time_Date.ToString("dd.MM.yyyy HH:mm:ss");


                        /*
                        int firstDayOfYear = (int)new DateTime(Create_Time_Date.Year, 1, 1).DayOfWeek;
                        int weekNumber = (Create_Time_Date.DayOfYear + firstDayOfYear) / 7 + 1;
                        oSheet.Cells[d + 1, 130].Formula = weekNumber.ToString();
                        */
                        oSheet.Cells[d + 1, 130].Formula = NumOfYear(Create_Time_Date).ToString();
                        oSheet.Cells[d + 1, 144].Formula = Create_Time_Date.Year.ToString();



                        double GMT_Zone = 0;

                        try
                        {

                            GMT_Zone = double.Parse(TimeZone_string);
                        }
                        catch
                        {
                            GMT_Zone = 8;
                        }


                        GMT_Zone = GMT_Zone - 8;

                        if (РемонтЗаведен1С.Trim() != "")
                        {
                            try
                            {
                                DateTime РемонтЗаведен1С_Date = DateTime.ParseExact(РемонтЗаведен1С, "dd.MM.yyyy H:mm:ss", null);
                                Create_Time_Date = Create_Time_Date.AddHours(GMT_Zone);

                                string Create_Time_Delta = string.Format("{0:N0}", (РемонтЗаведен1С_Date - Create_Time_Date).TotalHours);

                                oSheet.Cells[d + 1, 20].Formula = Create_Time_Delta;
                            }
                            catch
                            {

                            }
                        }

                        if (ДатаПоступленияНаФилиалАВР.Trim() != "")
                        {
                            try
                            {
                                ДатаПоступленияНаФилиалАВР_Date = DateTime.ParseExact(ДатаПоступленияНаФилиалАВР, "dd.MM.yyyy H:mm:ss", null);
                                Create_Time_Date = Create_Time_Date.AddHours(GMT_Zone);

                                string Create_Time_Delta = string.Format("{0:N0}", ( Create_Time_Date- ДатаПоступленияНаФилиалАВР_Date).TotalDays);

                                oSheet.Cells[d + 1, 30].Formula = Create_Time_Delta;
                            }
                            catch
                            {

                            }
                        }

                        ////////////////////////////////////////////////////


                        DateTime UDate;
                        string DataString = "";
                        try
                        {
                            DataString = (string)File1_Service_Order_Save_Time_List[x];
                            if (DataString.Trim() != "")
                            {
                                UDate = DateTime.ParseExact(DataString, "dd.MM.yyyy HH:mm:ss", null);
                                oSheet.Cells[d + 1, 114].Formula = UDate.ToString("dd.MM.yyyy HH:mm:ss");
                            }
                        }
                        catch
                        {

                        }

                        try
                        {

                            DataString = (string)File1_Carry_in_Time_List[x];
                            if (DataString.Trim() != "")
                            {
                                UDate = DateTime.ParseExact(DataString, "dd.MM.yyyy HH:mm:ss", null);
                                Carry_in_Time = UDate;
                                oSheet.Cells[d + 1, 115].Formula = UDate.ToString("dd.MM.yyyy HH:mm:ss");

                                if (ДатаПоступленияНаФилиалАВР.Trim() != "")
                                {
                                    try
                                    {
                                        ДатаПоступленияНаФилиалАВР_Date = DateTime.ParseExact(ДатаПоступленияНаФилиалАВР, "dd.MM.yyyy H:mm:ss", null);
                                        UDate = UDate.AddHours(GMT_Zone);

                                        string Carry_in_Time_Delta = string.Format("{0:N0}", (UDate - ДатаПоступленияНаФилиалАВР_Date).TotalDays);

                                        oSheet.Cells[d + 1, 99].Formula = Carry_in_Time_Delta;
                                    }
                                    catch
                                    {

                                    }
                                }

                            }
                        }
                        catch
                        {

                        }

                        try
                        {
                            DataString = (string)File1_Apply_for_Parts_Time_List[x];
                            if (DataString.Trim() != "")
                            {
                                UDate = DateTime.ParseExact(DataString, "dd.MM.yyyy HH:mm:ss", null);
                                Apply_for_Parts_Time_DT = UDate;
                                oSheet.Cells[d + 1, 116].Formula = UDate.ToString("dd.MM.yyyy HH:mm:ss");



                            }
                        }
                        catch
                        {

                        }


                        try
                        {
                            DataString = (string)File1_CI_Received_Time_List[x];
                            if (DataString.Trim() != "")
                            {
                                UDate = DateTime.ParseExact(DataString, "dd.MM.yyyy HH:mm:ss", null);
                                oSheet.Cells[d + 1, 117].Formula = UDate.ToString("dd.MM.yyyy HH:mm:ss");
                            }
                        }
                        catch
                        {

                        }

                        try
                        {

                            DataString = (string)File1_CP_Received_Time_List[x];
                            if (DataString.Trim() != "")
                            {
                                UDate = DateTime.ParseExact(DataString, "dd.MM.yyyy HH:mm:ss", null);
                                oSheet.Cells[d + 1, 118].Formula = UDate.ToString("dd.MM.yyyy HH:mm:ss");
                            }
                        }
                        catch
                        {

                        }
                        try
                        {
                            DataString = (string)File1_Parts_Available_in_Vendor_Time_List[x];
                            if (DataString.Trim() != "")
                            {
                                UDate = DateTime.ParseExact(DataString, "dd.MM.yyyy HH:mm:ss", null);
                                oSheet.Cells[d + 1, 119].Formula = UDate.ToString("dd.MM.yyyy HH:mm:ss");
                            }
                        }
                        catch
                        {

                        }
                        try
                        {
                            DataString = (string)File1_Parts_in_Service_Center_Time_List[x];
                            if (DataString.Trim() != "")
                            {
                                UDate = DateTime.ParseExact(DataString, "dd.MM.yyyy HH:mm:ss", null);
                                Parts_in_Service_Center_Time_DT = UDate;
                                oSheet.Cells[d + 1, 120].Formula = UDate.ToString("dd.MM.yyyy HH:mm:ss");
                            }
                        }
                        catch
                        {

                        }
                        try
                        {
                            DataString = (string)File1_Service_Order_Waiting_Time_List[x];
                            if (DataString.Trim() != "")
                            {
                                UDate = DateTime.ParseExact(DataString, "dd.MM.yyyy HH:mm:ss", null);
                                Service_Order_Waiting_Time = UDate;
                                oSheet.Cells[d + 1, 121].Formula = UDate.ToString("dd.MM.yyyy HH:mm:ss");
                            }
                        }
                        catch
                        {

                        }


                        try
                        {
                            DataString = (string)File1_Repair_Finish_Time_List[x];
                            if (DataString.Trim() != "")
                            {
                                UDate = DateTime.ParseExact(DataString, "dd.MM.yyyy HH:mm:ss", null);
                                Repair_Finish_Time = UDate;
                                oSheet.Cells[d + 1, 122].Formula = UDate.ToString("dd.MM.yyyy HH:mm:ss");


                                if (АВР_Дата.Trim() != "")
                                {
                                    try
                                    {
                                        DateTime АВР_Дата_Date = DateTime.ParseExact(АВР_Дата, "dd.MM.yyyy H:mm:ss", null);
                                        UDate = UDate.AddHours(GMT_Zone);

                                        string Time_Delta = string.Format("{0:N0}", ( UDate- АВР_Дата_Date).TotalDays);

                                        oSheet.Cells[d + 1, 32].Formula = Time_Delta;
                                    }
                                    catch
                                    {

                                    }
                                }

                                ///// TAT_REP

                                string TAT_String = string.Format("{0:N0}", (UDate - Create_Time_Date).TotalDays);
                                oSheet.Cells[d + 1, 80].Formula = TAT_String;


                                double TAT_Rep_double = (UDate.Subtract(UDate.TimeOfDay) - Create_Time_Date.Subtract(Create_Time_Date.TimeOfDay)).TotalDays;

                                string TAT_Rep_string = "";

                                if (TAT_Rep_double <= 3)
                                {
                                    TAT_Rep_string = "<=3D";
                                }

                                

                                if (TAT_Rep_double <= 3)
                                {
                                    oSheet.Cells[d + 1, 82].Formula = "KPI Repair 3 Days - OK";
                                }


                                if ((TAT_Rep_double > 3) && (TAT_Rep_double <= 7))
                                {
                                    TAT_Rep_string = "4~7D";
                                }

                                if ((TAT_Rep_double > 7) && (TAT_Rep_double <= 14))
                                {
                                    TAT_Rep_string = "8~14D";
                                }

                                if ((TAT_Rep_double > 14) && (TAT_Rep_double <= 30))
                                {
                                    TAT_Rep_string = "15~30D";
                                }

                                if (TAT_Rep_double > 30)
                                {
                                    TAT_Rep_string = ">30D";
                                }

                                oSheet.Cells[d + 1, 81].Formula = TAT_Rep_string;

                            }
                        }
                        catch
                        {

                        }



                        try
                        {
                            DataString = (string)File1_Customer_Pick_up_Time_List[x];
                            if (DataString.Trim() != "")
                            {
                                UDate = DateTime.ParseExact(DataString, "dd.MM.yyyy HH:mm:ss", null);
                                oSheet.Cells[d + 1, 123].Formula = UDate.ToString("dd.MM.yyyy HH:mm:ss");
                            }
                        }
                        catch
                        {

                        }



                        try
                        {
                            DataString = (string)File1_System_Close_Time_List[x];
                            if (DataString.Trim() != "")
                            {
                                UDate = DateTime.ParseExact(DataString, "dd.MM.yyyy HH:mm:ss", null);
                                oSheet.Cells[d + 1, 124].Formula = UDate.ToString("dd.MM.yyyy HH:mm:ss");

                                oSheet.Cells[d + 1, 134].Formula = UDate.ToString("dd.MM.yyyy");

                                /*
                                firstDayOfYear = (int)new DateTime(UDate.Year, 1, 1).DayOfWeek;
                                weekNumber = (UDate.DayOfYear + firstDayOfYear) / 7 + 1;
                                oSheet.Cells[d + 1, 131].Formula = weekNumber.ToString();
                                */
                                oSheet.Cells[d + 1, 131].Formula = NumOfYear(UDate).ToString();


                                oSheet.Cells[d + 1, 145].Formula = UDate.Year.ToString();

                                /// расхождение в сроках закрытия
                                string РемонтЗакрыт1С = (string)File2_РемонтЗакрыт1С_List[d];
                                if (РемонтЗакрыт1С.Trim() != "")
                                {
                                    DateTime РемонтЗакрыт1С_Date = DateTime.ParseExact(РемонтЗакрыт1С, "dd.MM.yyyy H:mm:ss", null);



                                    UDate = UDate.AddHours(GMT_Zone);

                                    string Close_Time_Delta = string.Format("{0:N0}", (РемонтЗакрыт1С_Date - UDate).TotalHours);

                                    oSheet.Cells[d + 1, 21].Formula = Close_Time_Delta;
                                }


                                /// Дата отгрузки/выдачи и закрытие по базе вендора

                                if (ДатаОтгрузкиВыдачиТовара_из_ФилиалРемонта.Trim() != "")
                                {
                                    DateTime ДатаОтгрузкиВыдачиТовара_из_ФилиалРемонта_Date = DateTime.ParseExact(ДатаОтгрузкиВыдачиТовара_из_ФилиалРемонта, "dd.MM.yyyy H:mm:ss", null);



                                    UDate = UDate.AddHours(GMT_Zone);

                                    string Close_Time_Delta = string.Format("{0:N0}", (UDate- ДатаОтгрузкиВыдачиТовара_из_ФилиалРемонта_Date).TotalDays);

                                    oSheet.Cells[d + 1, 42].Formula = Close_Time_Delta;
                                }


                                //// Срок ремонта

                                string TAT_String = string.Format("{0:N0}", (UDate - Create_Time_Date).TotalDays);
                                oSheet.Cells[d + 1, 19].Formula = TAT_String;

                                double TAT_Close_double = (UDate.Subtract(UDate.TimeOfDay)- Create_Time_Date.Subtract(Create_Time_Date.TimeOfDay)).TotalDays;

                                string TAT_Close_string = "";
                                

                                if (TAT_Close_double <= 3)
                                {
                                    TAT_Close_string = "<=3D";
                                }

                                if (TAT_Close_double <= 7)
                                {                                     
                                    oSheet.Cells[d + 1, 79].Formula = "KPI 7 Days - OK"; 
                                }

                                if ((TAT_Close_double > 3) && (TAT_Close_double <= 7))
                                {
                                    TAT_Close_string = "4~7D";
                                }
                                
                                if ((TAT_Close_double > 7) && (TAT_Close_double <= 14))
                                {
                                    TAT_Close_string = "8~14D";
                                }

                                if ((TAT_Close_double > 14) && (TAT_Close_double <= 30))
                                {
                                    TAT_Close_string = "15~30D";
                                }

                                if (TAT_Close_double > 30)
                                {
                                    TAT_Close_string = ">30D";
                                }

                                oSheet.Cells[d + 1, 78].Formula = TAT_Close_string;


                            }
                            else
                            {
                                DateTime CurDT = DateTime.Now;

                                //// Срок незакрытого ремонта

                                string TAT_String = string.Format("{0:N0}", (CurDT - Create_Time_Date).TotalDays);
                                oSheet.Cells[d + 1, 19].Formula = TAT_String;

                            }
                        }
                        catch
                        {

                        }

                        /////////////////////


                        ///  Время диагностики CRM
                        string Service_Order_Status = (string)File1_Service_Order_Status_List[x];
                        if (Service_Order_Status == "SR Not Submit")
                        {
                            DateTime CurDT = DateTime.Now;
                            string ДнейДиагностики_String = string.Format("{0:N0}", (CurDT - Carry_in_Time).TotalDays);
                            oSheet.Cells[d + 1, 94].Formula = ДнейДиагностики_String;

                            ///  Время диагностики от даты прихода на филиал ремонта

                            ДнейДиагностики_String = string.Format("{0:N0}", (CurDT - ДатаПоступленияНаФилиалАВР_Date).TotalDays);
                            oSheet.Cells[d + 1, 95].Formula = ДнейДиагностики_String;

                        }
                        else
                        {
                            if ((string)File1_Apply_for_Parts_Time_List[x] != "")
                            {
                                DateTime CurDT = DateTime.Now;
                                string ДнейДиагностики_String = string.Format("{0:N0}", (Apply_for_Parts_Time_DT- Carry_in_Time).TotalDays);
                                oSheet.Cells[d + 1, 94].Formula = ДнейДиагностики_String;


                                ///  Время диагностики от даты прихода на филиал ремонта

                                ДнейДиагностики_String = string.Format("{0:N0}", (Apply_for_Parts_Time_DT - ДатаПоступленияНаФилиалАВР_Date).TotalDays);
                                oSheet.Cells[d + 1, 95].Formula = ДнейДиагностики_String;

                            }
                            else
                            {
                                DateTime CurDT = DateTime.Now;
                                string ДнейДиагностики_String = string.Format("{0:N0}", (Service_Order_Waiting_Time- Carry_in_Time).TotalDays);
                                oSheet.Cells[d + 1, 94].Formula = ДнейДиагностики_String;
                               
                                ///  Время диагностики от даты прихода на филиал ремонта

                                ДнейДиагностики_String = string.Format("{0:N0}", (Service_Order_Waiting_Time - ДатаПоступленияНаФилиалАВР_Date).TotalDays);
                                oSheet.Cells[d + 1, 95].Formula = ДнейДиагностики_String;

                            }

                        }

                        //// Дней ожидания ЗИП
                        /*
                        ArrayList File1_Apply_for_Parts_Time_List = new ArrayList();
                        ArrayList File1_CI_Received_Time_List = new ArrayList();
                        ArrayList File1_CP_Received_Time_List = new ArrayList();
                        ArrayList File1_Parts_Available_in_Vendor_Time_List = new ArrayList();
                        ArrayList File1_Parts_in_Service_Center_Time_List = new ArrayList();
                        */


                        if ((string)File1_Apply_for_Parts_Time_List[x] != "")
                        {
                            if ((string)File1_Parts_in_Service_Center_Time_List[x] != "")
                            {                                
                                string ДнейОжиданияЗИП_String = string.Format("{0:N0}", (Parts_in_Service_Center_Time_DT - Apply_for_Parts_Time_DT).TotalDays);
                                oSheet.Cells[d + 1, 96].Formula = ДнейОжиданияЗИП_String;
                            }
                            else
                            {
                                DateTime CurDT = DateTime.Now;
                                string ДнейОжиданияЗИП_String = string.Format("{0:N0}", (CurDT - Apply_for_Parts_Time_DT).TotalDays);
                                oSheet.Cells[d + 1, 96].Formula = ДнейОжиданияЗИП_String;

                            }

                        }
                        else
                        {
                            oSheet.Cells[d + 1, 96].Formula = "0";

                        }

                        /////////////////////////////////////
                        /// Дней от поступления ЗИП до завершения ремонта


                        if ((string)File1_Parts_in_Service_Center_Time_List[x] != "")
                        {
                            if ((string)File1_Repair_Finish_Time_List[x]!="")
                            {
                                string ДнейРемонта_String = string.Format("{0:N0}", (Repair_Finish_Time - Parts_in_Service_Center_Time_DT).TotalDays);
                                oSheet.Cells[d + 1, 97].Formula = ДнейРемонта_String;

                            }
                            else
                            {
                                DateTime CurDT = DateTime.Now;
                                string ДнейРемонта_String = string.Format("{0:N0}", (CurDT - Parts_in_Service_Center_Time_DT).TotalDays);
                                oSheet.Cells[d + 1, 97].Formula = ДнейРемонта_String;
                            }

                        }
                        else
                        {
                            oSheet.Cells[d + 1, 97].Formula = "0";

                        }

                        ///// Активация устройства

                            try
                        {
                            DataString = (string)File1_Actived_Time_List[x];
                            if (DataString.Trim() != "")
                            {
                                UDate = DateTime.ParseExact(DataString, "dd.MM.yyyy HH:mm:ss", null);
                                oSheet.Cells[d + 1, 132].Formula = UDate.ToString("dd.MM.yyyy HH:mm:ss");


                                /// расхождение в сроках активации и продажи
                                string ДатаПродажи_или_выявления_предторга = (string)File2_ДатаПродажи_или_выявления_предторга_List[d];
                                if (ДатаПродажи_или_выявления_предторга.Trim() != "")
                                {
                                    DateTime ДатаПродажи_или_выявления_предторга_Date = DateTime.ParseExact(ДатаПродажи_или_выявления_предторга, "dd.MM.yyyy H:mm:ss", null);



                                    UDate = UDate.AddHours(GMT_Zone);

                                    string Time_Delta = string.Format("{0:N0}", ( UDate- ДатаПродажи_или_выявления_предторга_Date).TotalDays);

                                    oSheet.Cells[d + 1, 70].Formula = Time_Delta;
                                }

                            }
                        }
                        catch
                        {

                        }


                        #endregion





                        /////////////////////////
                        //// Проверка повторных ремонтов 
                        if (RR_Found)
                        {

                            oSheet.Cells[d + 1, 125].Formula = RR_DT.ToString("dd.MM.yyyy HH:mm:ss");
                            oSheet.Cells[d + 1, 126].Formula = string.Format("{0:N0}", RR_Days_double);

                            if (RR_Days_double < 31)
                            {
                                oSheet.Cells[d + 1, 127].Formula = "Повторный ремонт менее 30 дней";

                            }

                            oSheet.Cells[d + 1, 128].Formula = (string)File1_Service_Order_Code_List[RR_Index];

                            string ASP_Order_NO_Prev = (string)File1_ASP_Order_NO_List[RR_Index];
                            string ОСНТ_Номер = (string)File2_ОСНТ_Номер_List[d];

                            oSheet.Cells[d + 1, 129].Formula = ASP_Order_NO_Prev;

                            if (ASP_Order_NO_Prev.Contains(ОСНТ_Номер))
                            {
                                oSheet.Cells[d + 1, 139].Formula = string.Format("Один ремонт в 1с: {0}={1}  {2}={3}", (string)File1_Service_Order_Code_List[RR_Index], ASP_Order_NO_Prev, Service_Order_Cod, ОСНТ_Номер);
                            }

                            oSheet.Cells[d + 1, 5].Formula = string.Format("Повторный ремонт {0:N0} дней, {1}  {2}  {3}", RR_Days_double, RR_DT.ToString("dd.MM.yyyy HH:mm:ss"), (string)File1_Service_Order_Code_List[RR_Index], (string)File1_ASP_Order_NO_List[RR_Index]);

                            oSheet.Cells[d + 1, 140].Formula = (string)File1_Service_Order_Status_List[RR_Index];

                            RR_ОСНТ_List[RR_List.Count - 1] = (string)File2_ОСНТ_Номер_и_Дата_List[d];
                            RR_Prev_ОСНТ_List[RR_List.Count - 1] = (string)File1_ASP_Order_NO_List[RR_Index];


                        }





                        ////////////////////
                        /// Проверка Service Order Code

                        /*

                        string РемонтЗаведенКомментарий = (string)File2_РемонтЗаведенКомментарий_List[d];

                        string sCell11 = string.Format("Q{0}", d + 1);
                        string sCell22 = string.Format("R{0}", d + 1);


                        if ((РемонтЗаведенКомментарий.Contains(Service_Order_Cod.Trim())) && (РемонтЗаведенКомментарий.Trim() != ""))
                        {
                            oSheet.get_Range(sCell11, sCell22).Interior.ColorIndex = 4;
                        }
                        else
                        {
                            oSheet.get_Range(sCell11, sCell22).Interior.ColorIndex = 46;
                        }

                        */

                        ////////////
                        // Проверка модели


                        int FoundResult = 0;

                        if (!File3_Xiaomi_Model_Name_List.Contains(Product_Model))
                        {
                            FoundResult = 0;

                        }
                        else
                        {
                            for (int c = 0; c < File3_Xiaomi_Model_Name_List.Count; c++)
                            {

                                if (((string)File3_Xiaomi_Model_Name_List[c] == Product_Model) && ((string)File3_DNS_Code_List[c] == Pos_Name))
                                {
                                    FoundResult = 1;
                                    break;
                                }
                            }


                        }


                        string sCell1 = string.Format("F{0}", d + 1);
                        string sCell2 = string.Format("G{0}", d + 1);

                        if (FoundResult == 0)
                        {
                            oSheet.get_Range(sCell1, sCell2).Interior.ColorIndex = 46;
                            oSheet.Cells[d + 1, 100].Formula = "Да";

                        }
                        else
                        {
                            oSheet.get_Range(sCell1, sCell2).Interior.ColorIndex = 4;

                        }

                       





                    }
                    #endregion



                    //// Добавляем в лист ненайденных

                    if (d == -1)
                    {
                        NotFindList.Add(x);




                    }









                }
                ///////////////////////

                richTextBox1.Text = richTextBox1.Text + "\n" + "Формирование рекомендаций";
                /// Проверка статуса повторных событий - Внесено в базу вендора

                for (int d = 1; d < File2_ОСНТ_Номер_List.Count; d++)
                {
                    ///int x = (int)File2_Index_from_File1_List[d];

                    string Событие_ЗаведенияРемонта_в_базе_вендора_СрезПоследних_Комментарий = (string)File2_Событие_ЗаведенияРемонта_в_базе_вендора_СрезПоследних_Комментарий_List[d];

                    for (int x = 1; x < File1_Service_Order_Status_List.Count; x++)
                    {                    

                        string Service_Order_Code = (string)File1_Service_Order_Code_List[x];

                        if (Событие_ЗаведенияРемонта_в_базе_вендора_СрезПоследних_Комментарий.Contains(Service_Order_Code))
                        {
                            string Service_Order_Status = (string)File1_Service_Order_Status_List[x];
                            oSheet.Cells[1+d, 87].Formula = Service_Order_Status;
                        }
                    }
                }


                /// Формирование рекомендаций

                for (int d = 1; d < File2_ОСНТ_Номер_List.Count; d++)
                {
                    int x = (int)File2_Index_from_File1_List[d];

                    ////////////////////////////////////
                    ///// Действия по авторизации

                    string ИтогиРемонтаDNS = (string)File2_ИтогиРемонтаDNS_List[d];
                    string ДатаПринято_решение_ремонтировать_товар_на_филиале = (string)File2_ДатаПринято_решение_ремонтировать_товар_на_филиале_List[d];
                    string ДатаБезВнесенияВБазуВендора = (string)File2_ДатаБезВнесенияВБазуВендора_List[d];
                    string ДействиеОперации = (string)File2_ДействиеОперации_List[d];
                    string ДатаОтгрузкиВыдачиТовара_из_ФилиалРемонта = (string)File2_ДатаОтгрузкиВыдачиТовара_из_ФилиалРемонта_List[d];
                    string СтатусРемонтаВендора1с = (string)File2_СтатусРемонтаВендора1с_List[d];
                    string ОСНТ_Номер_и_Дата = (string)File2_ОСНТ_Номер_и_Дата_List[d];
                    string ГарантийныйНаОснованииПлатного = (string)File2_ГарантийныйНаОснованииПлатного_List[d];
                    string АВР_РезультатУслуг = (string)File2_АВР_РезультатУслуг_List[d];
                    string Событие_ДиагностикаДефектВыявлен_Автор = (string)File2_Событие_ДиагностикаДефектВыявлен_Автор_List[d];
                    string ТекущийСтатусОстатка = (string)File2_ТекущийСтатусОстатка_List[d];
                    string Исключение_DOA = (string)File2_Исключение_DOA_List[d];




                    //// ОСНТ Исключения

                    if (ОСНТ_Номер_и_Дата.Trim() == "Операции с неисправным товаром БФ6-000001 от 22.06.2018 16:41:46")
                    {
                        continue;
                    }


                    string Service_Order_Status = "";
                    if (x >= 0)
                    {
                        Service_Order_Status = (string)File1_Service_Order_Status_List[x];
                    }


                    string СогласованныйФилиалРемонта = (string)File2_СогласованныйФилиалРемонта_List[d];

                    string ДатаПоступленияНаФилиалАВР = (string)File2_ДатаПоступленияНаФилиалАВР_List[d];
                    string АВР_Дата = (string)File2_АВР_Дата_List[d];
                    /*
                    ArrayList File2_СтатусРемонтаВендора1с_List = new ArrayList();
                    ArrayList File2_ДатаПоступленияНаФилиалАВР_List = new ArrayList();
                    ArrayList File2_АВР_Дата_List = new ArrayList();
                    ArrayList File2_ДатаОтгрузкиВыдачиТовара_из_ФилиалРемонта_List = new ArrayList();
                    ArrayList File2_ДатаПродажи_или_выявления_предторга_List = new ArrayList();
                    */
                    string Repair_Finish_Time = "";
                    if (x >= 0)
                    {
                        Repair_Finish_Time = (string)File1_Repair_Finish_Time_List[x];
                    }

                    string Parts_in_Service_Center_Time = "";

                    if (x >= 0)
                    {
                        Parts_in_Service_Center_Time = (string)File1_Parts_in_Service_Center_Time_List[x];
                    }



                    DateTime Create_Time_Date = DateTime.Now;

                    if (x >= 0)
                    {
                        string Create_Time = (string)File1_Create_Time_List[x];
                        if (Create_Time != "")
                        {
                            try
                            {
                                Create_Time_Date = DateTime.ParseExact(Create_Time, "dd.MM.yyyy HH:mm:ss", null);
                            }
                            catch
                            {

                            }
                        }

                    }

                    DateTime Parts_in_Service_Center_Time_DT = DateTime.Now;

                    if (x >= 0)
                    {
                        string Parts_in_Service_Center_Time_string = (string)File1_Parts_in_Service_Center_Time_List[x];
                        if (Parts_in_Service_Center_Time_string != "")
                        {
                            try
                            {
                                Parts_in_Service_Center_Time_DT = DateTime.ParseExact(Parts_in_Service_Center_Time_string, "dd.MM.yyyy HH:mm:ss", null);
                            }
                            catch
                            {

                            }

                        }

                    }


                    DateTime Apply_for_Parts_Time_DT = DateTime.Now;

                    if (x >= 0)
                    {
                        string Apply_for_Parts_Time = (string)File1_Apply_for_Parts_Time_List[x];
                        if (Apply_for_Parts_Time != "")
                        {
                            try
                            {
                                Apply_for_Parts_Time_DT = DateTime.ParseExact(Apply_for_Parts_Time, "dd.MM.yyyy HH:mm:ss", null);
                            }
                            catch
                            {

                            }
                        }

                    }

                    DateTime ДатаПоступленияНаФилиалАВР_DT = DateTime.Now;

                    if (ДатаПоступленияНаФилиалАВР != "")
                    {
                        try
                        {
                            ДатаПоступленияНаФилиалАВР_DT = DateTime.ParseExact(ДатаПоступленияНаФилиалАВР, "dd.MM.yyyy HH:mm:ss", null);
                        }
                        catch
                        {

                        }
                    }

                    


                    string ОСНТ_Дата = (string)File2_ОСНТ_Дата_List[d];
                    DateTime ОСНТ_Дата_DT = DateTime.Now;
                    /////////////////////////////////////////////////////
                    string DOA_DAP_Cancel_Date = "15.11.2018 00:00:00";
                    /////////////////////////////////////////////////////

                    DateTime DOA_DAP_Cancel_Date_DT = DateTime.ParseExact(DOA_DAP_Cancel_Date, "dd.MM.yyyy HH:mm:ss", null);

                    if (ОСНТ_Дата != "")
                    {
                        try
                        {
                            ОСНТ_Дата_DT = DateTime.ParseExact(ОСНТ_Дата, "dd.MM.yyyy HH:mm:ss", null);
                        }
                        catch
                        { }

                    }




                    string OrderResult = "";

                    //////////////////////////////////////////////////
                    string ТекущийОстаток_Действие = (string)File2_ТекущийОстаток_Действие_List[d];
                    if (ТекущийОстаток_Действие.Trim() != "")
                    {
                        string ТекущийОстаток_ДнейОтПоследнегоДвижения = (string)File2_ТекущийОстаток_ДнейОтПоследнегоДвижения_List[d];

                        OrderResult = OrderResult + ТекущийОстаток_Действие + string.Format(" ({0} дн. без движения) ", ТекущийОстаток_ДнейОтПоследнегоДвижения);
                    }


                    if ((ГарантийныйНаОснованииПлатного.Trim() != "")&&(ДатаБезВнесенияВБазуВендора == "") && (ИтогиРемонтаDNS == ""))
                    {

                        OrderResult = OrderResult + "Гарантийный ремонт на основании платного обращения!!!; ";

                    }
                    


                    /// когда есть АВР
                    if ((АВР_Дата != "") && (СогласованныйФилиалРемонта == "Да") && (Repair_Finish_Time == "") && (Service_Order_Status != "")&&(Service_Order_Status != "cancelled"))
                    {

                        OrderResult = OrderResult + "АВР проведен, внесите данные о ремонте в CRM; ";
                    }


                    //// Приход на филиал ремонта 
                    if ((ДатаПоступленияНаФилиалАВР != "") && (Service_Order_Status == "") && (ИтогиРемонтаDNS == "") && (ДатаБезВнесенияВБазуВендора == "") && (СтатусРемонтаВендора1с != "Закрыт") && (ГарантийныйНаОснованииПлатного.Trim() == ""))
                    {
                        if ((ДействиеОперации != "Предпродажный ремонт") || ((ДействиеОперации == "Предпродажный ремонт") && (ОСНТ_Дата_DT >= DOA_DAP_Cancel_Date_DT)))
                        {
                            /// Есть АВР или Есть статус подтверждения наличия дефекта
                            if ((АВР_РезультатУслуг == "Готово (Тех. обслуживание)") || (АВР_РезультатУслуг == "Отказ от ремонта (гарантия) (Диагностика)") || (АВР_РезультатУслуг == "Отремонтирован (Ремонт)") || (Событие_ДиагностикаДефектВыявлен_Автор != ""))
                            {
                                string DOA = "";
                                //  if ((ТекущийСтатусОстатка == "41.4 Товары в браке")|| (ТекущийСтатусОстатка == "41.3 Товары в ремонте"))
                                if ((ДействиеОперации == "Предпродажный ремонт") && (Исключение_DOA != "Исключение DOA"))
                                {
                                    DOA = "Выберите isDOA Repair.";
                                }

                                OrderResult = OrderResult + string.Format("Cоздайте ремонт в CRM. {0} Внесите данные Carry-in Time: {1}; ",DOA, ДатаПоступленияНаФилиалАВР);
                            }

                            /// Если нет АВР и нет статуса - ждем резултатов первичной диагностики
                            if ((АВР_РезультатУслуг == "") && (Событие_ДиагностикаДефектВыявлен_Автор == ""))                          
                            {

                                DateTime CurDT = DateTime.Now;
                                                             

                              //  double ЧасовОтПриема_double = (CurDT - ДатаПоступленияНаФилиалАВР_DT).TotalHours;
                                string TAT_String = string.Format("{0:N0}", (CurDT - ДатаПоступленияНаФилиалАВР_DT).TotalHours);


                                OrderResult = OrderResult + string.Format("Устройство получено на филиале {0} ч. - ожидание подтверждения дефекта (первичная диагностика перед внесением в CRM)", TAT_String);
                            }   

                        }
                    }


                    //// Проверка времени диагности и заказа ЗИП
                    if ((ДатаПоступленияНаФилиалАВР != "")&& (АВР_Дата == "") && (Service_Order_Status == "SR Not Submit") && (ИтогиРемонтаDNS == ""))
                    {
                        DateTime CurDT = DateTime.Now;

                        //// Срок незакрытого ремонта

                        double ЧасовОтПриема_double = (CurDT - Create_Time_Date).TotalHours;
                        string TAT_String = string.Format("{0:N0}", (CurDT - Create_Time_Date).TotalHours);

                        if (ЧасовОтПриема_double <= 24)
                        {
                            OrderResult = OrderResult + string.Format("Устройство ожидает диагностики и заказа ЗИП в базе вендора {0} ч.; ", TAT_String);
                        }
                        else
                        {
                            OrderResult = OrderResult + string.Format("Длительная диагностика по базе вендора!!! С момента создания ремонта прошло {0} ч.; ", TAT_String);
                        }
                    }

                    ////  Если ЗИП на СЦ, но нет АВР

                    if ((АВР_Дата == "") && (Service_Order_Status == "Stock In Service Center"))
                    {
                        DateTime CurDT = DateTime.Now;

                        //// Время с момента прихода ЗИП на филиал

                        //  double ДнейОтПриема_double = (CurDT - Parts_in_Service_Center_Time_DT).TotalDays;
                        string TAT_String = string.Format("{0:N0}", (CurDT - Parts_in_Service_Center_Time_DT).TotalDays);



                        OrderResult = OrderResult + string.Format("ЗИП получен и находится в СЦ {0} дн. - приступите к ремонту; ", TAT_String);

                    }


                    ////  ЗИП внесен, но находится не в СЦ

                    if ((Service_Order_Status == "Parts In Country Stock") || (Service_Order_Status == "Parts In Local Maitrox Hub"))
                    {
                        DateTime CurDT = DateTime.Now;

                        //// Время с момента прихода ЗИП на филиал

                        double ДнейОтПриема_double = (CurDT - Apply_for_Parts_Time_DT).TotalDays;
                        string TAT_String = string.Format("{0:N0}", (CurDT - Apply_for_Parts_Time_DT).TotalDays);


                        if (АВР_Дата == "")
                        {
                            OrderResult = OrderResult + string.Format("Ожидание прихода ЗИП в СЦ {0} дн. - проверьте создание документов на отрузку ЗИП; ", TAT_String);
                        }
                        else
                        {
                            OrderResult = OrderResult + string.Format("АВР проведен, но ЗИП в CRM не принят на филиале ремонта; ", TAT_String);
                        }
                    }


                    if ((Service_Order_Status == "") && (СтатусРемонтаВендора1с == "Открыт"))
                    {
                        OrderResult = OrderResult + "Разхождение статусов - ремонт не заведен в CRM; ";
                    }

                    if ((Service_Order_Status != "") && (СтатусРемонтаВендора1с == "")&&(Service_Order_Status != "Closed"))
                    {
                        OrderResult = OrderResult + "Разхождение статусов - ремонт заведен в CRM, но нет статуса в 1с; ";
                    }


                    if (((Service_Order_Status == "Closed") || (Service_Order_Status == "cancelled")) && (СтатусРемонтаВендора1с != "Закрыт"))
                    {
                        OrderResult = OrderResult + "Разхождение статусов - ремонт закрыт в CRM, но нет статуса в 1с; ";
                    }

                    if ((Service_Order_Status != "Closed") && (Service_Order_Status != "") && (Service_Order_Status != "cancelled") && (СтатусРемонтаВендора1с == "Закрыт"))
                    {
                        OrderResult = OrderResult + "Разхождение статусов - проставлен статус закрытия в 1с, но ремонт не закрыт в CRM; ";
                    }

                    // когда товар выдан, отгружен или списан в брак/отремонтирован

                    if (((ИтогиРемонтаDNS != "") || (ДатаОтгрузкиВыдачиТовара_из_ФилиалРемонта != "")) && (Service_Order_Status != "Closed") && (Service_Order_Status != "cancelled") && (Service_Order_Status != ""))
                    {
                        OrderResult = OrderResult + "Закройте ремонт в базе вендора; ";

                        if (СтатусРемонтаВендора1с != "Закрыт")
                        {
                            OrderResult = OrderResult + "Уставновите статус закрытия ремонта в 1c; ";

                        }
                    }

                    oSheet.Cells[d+1, 10].Formula = OrderResult;

                    if (OrderResult.Trim() != "")
                    {
                        oSheet.Cells[d + 1, 98].Formula = "Да";
                    }
                }


                ///////////////////////////////////////////////
                ///// Заголовки и ширина столбцов

                string sCell1a = string.Format("DG{0}", 1);
                string sCell2a = string.Format("ET{0}", 1);            
                
               oSheet.get_Range(sCell1a, sCell2a).Interior.ColorIndex = 50 ;         
              




                oSheet.Cells[1, 111].Formula = "ASP_Order_NO";
                oSheet.Cells[1, 112].Formula = "Model_PN";
                oSheet.Cells[1, 113].Formula = "Create_Time";
                oSheet.Cells[1, 114].Formula = "Service Order Save Time";
                oSheet.Cells[1, 115].Formula = "Carry-in Time";
                oSheet.Cells[1, 116].Formula = "Apply for Parts Time";
                oSheet.Cells[1, 117].Formula = "CI Received Time";
                oSheet.Cells[1, 118].Formula = "CP Received Time";
                oSheet.Cells[1, 119].Formula = "Parts Available in Vendor Time";
                oSheet.Cells[1, 120].Formula = "Parts in Service Center Time";
                oSheet.Cells[1, 121].Formula = "Service Order Waiting Time";
                oSheet.Cells[1, 122].Formula = "Repair Finish Time";
                oSheet.Cells[1, 123].Formula = "Customer Pick-up Time";
                oSheet.Cells[1, 124].Formula = "System Close Time";


                oSheet.Cells[1, 125].Formula = "Дата предыдущего ремонта CRM";
                oSheet.Cells[1, 126].Formula = "Дней от предыдущего ремонта CRM";
                oSheet.Cells[1, 127].Formula = "Повторный ремонт в 30 дней";
                oSheet.Cells[1, 128].Formula = "SO повторного ремонта";
                oSheet.Cells[1, 129].Formula = "ASP_Order повторного ремонта";

                oSheet.Cells[1, 130].Formula = "Create Time номер недели";
                oSheet.Cells[1, 131].Formula = "System Close Time номер недели";
                oSheet.Cells[1, 132].Formula = "Actived_Time";
                oSheet.Cells[1, 133].Formula = "IMEI";
                oSheet.Cells[1, 134].Formula = "System Close Time - только дата";
                oSheet.Cells[1, 135].Formula = "Replacement_Part_Number";

                oSheet.Cells[1, 136].Formula = "Defective_Part_1_Name";
                oSheet.Cells[1, 137].Formula = "Defective_Part_2_Name";
                oSheet.Cells[1, 138].Formula = "Defective_Part_3_Name";
                oSheet.Cells[1, 139].Formula = "Повторный ремонт CRM - Один ремонт в 1с";
                oSheet.Cells[1, 140].Formula = "Статус предыдущего обращения в CRM";
                oSheet.Cells[1, 141].Formula = "DOA/DAP";
                oSheet.Cells[1, 142].Formula = "Recover Method";
                oSheet.Cells[1, 143].Formula = "Problem Category";

                oSheet.Cells[1, 144].Formula = "Create Time - Год";
                oSheet.Cells[1, 145].Formula = "System Close Time - Год";
                oSheet.Cells[1, 146].Formula = "Sale Country";

                oSheet.Columns[1].ColumnWidth = 25;
                oSheet.Columns[5].ColumnWidth = 30;
                oSheet.Columns[6].ColumnWidth = 45;
                oSheet.Columns[7].ColumnWidth = 30;
                oSheet.Columns[10].ColumnWidth = 85;
                oSheet.Columns[17].ColumnWidth = 25;
                oSheet.Columns[18].ColumnWidth = 15;
                oSheet.Columns[19].ColumnWidth = 15;
                oSheet.Columns[20].ColumnWidth = 15;
                oSheet.Columns[21].ColumnWidth = 15;
                oSheet.Columns[23].ColumnWidth = 5;
                oSheet.Columns[26].ColumnWidth = 5;
                oSheet.Columns[28].ColumnWidth = 10;
                oSheet.Columns[30].ColumnWidth = 10;
                oSheet.Columns[32].ColumnWidth = 10;
                oSheet.Columns[34].ColumnWidth = 8;
                oSheet.Columns[35].ColumnWidth = 7;
                oSheet.Columns[36].ColumnWidth = 10;
                oSheet.Columns[37].ColumnWidth = 10;
                oSheet.Columns[38].ColumnWidth = 10;
                oSheet.Columns[39].ColumnWidth = 10;
                oSheet.Columns[42].ColumnWidth = 12;
                oSheet.Columns[52].ColumnWidth = 56;
                oSheet.Columns[58].ColumnWidth = 5;
                oSheet.Columns[61].ColumnWidth = 12;
                oSheet.Columns[68].ColumnWidth = 50;
                oSheet.Columns[70].ColumnWidth = 15;

                for (int c = 78; c <= 146; c++)
                {
                    oSheet.Columns[c].ColumnWidth = 18;
                }
               
                oWB.Save();


                //oWB.SaveAs(savename, Excel.XlFileFormat.xlWorkbookNormal, m_objOpt,  m_objOpt, m_objOpt, m_objOpt, Excel.XlSaveAsAccessMode.xlExclusive,   m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);


                oWB.Close(false, m_objOpt, m_objOpt);



                oXL.Quit();


                Marshal.ReleaseComObject(oSheet);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oXL);
                oSheet = null;
                oWB = null;
                oXL = null;
                GC.Collect();

            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = oldCulture;
            }

            StaticFunction.ClearExcelProcess(BeforeList, AfterList);



        }

        public void CreateRR_File(string savename)
        {
            System.Globalization.CultureInfo oldCulture = System.Threading.Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            ArrayList BeforeList, AfterList;

            //////////////////////////////////////
            try
            {

                ///////////////////

                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;
                object m_objOpt = System.Reflection.Missing.Value;



               // WorkStatusObj.TextStatus = "Сохранение прайса";

                BeforeList = StaticFunction.GetExcelProcessID();

                oXL = new Excel.Application();

                //////////////////
                AfterList = StaticFunction.GetExcelProcessID();

                /////////////

                oXL.Interactive = false;
                oXL.EnableEvents = false;

                oXL.ScreenUpdating = false;
                oXL.Visible = false;

                oWB = (Excel._Workbook)(oXL.Workbooks.Add(m_objOpt));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;


             //   oSheet.Cells[1, 1] = string.Format("Прайс {0} {1}", FirmName, RegString);

                string Cell1, Cell2;
                Cell1 = "A1";
                Cell2 = "F1";
                oSheet.get_Range(Cell1, Cell2).Interior.ColorIndex = 6;
                oSheet.get_Range(Cell1, Cell2).Font.Bold = true;



                oSheet.Cells[1, 1] = "ОСНТ";
                oSheet.Cells[1, 2] = "Предыдущее ОСНТ";
                oSheet.Cells[1, 3] = "ДнейМеждуРемонтами";
                oSheet.Cells[1, 4] = "ДатаПредыдущегоРемонта";
                oSheet.Cells[1, 5] = "SO";


                ////////////////////////
                /*
                                RR_ОСНТ_List.Add(ОСНТ_Номер_и_Дата_Column);
                                RR_Prev_ОСНТ_List.Add((string)File1_ASP_Order_NO_List[RR_Index]);

                                     ArrayList RR_List = new ArrayList();
                        ArrayList RR_Days_List = new ArrayList();
                        ArrayList RR_SO_List = new ArrayList();
                        ArrayList RR_Prev_Date_List = new ArrayList();
                        ArrayList RR_ОСНТ_List = new ArrayList();
                        ArrayList RR_Prev_ОСНТ_List = new ArrayList();

                                oSheet.Cells[RR_Line + 1, 123].Formula = (string)RR_Days_List[n];
                                oSheet.Cells[RR_Line + 1, 124].Formula = (string)RR_SO_List[n];
                                oSheet.Cells[RR_Line + 1, 125].Formula = (string)RR_Prev_Date_List[n];

                    */
                string range = "A1";
                double Width = 70;

                object Range = oSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, oSheet, new object[] { range });
                object[] args = new object[] { Width };
                Range.GetType().InvokeMember("ColumnWidth", BindingFlags.SetProperty, null, Range, args);

                ////////////////////////////////////////////////////

                int LineIns = 2;

                for (int w = 0; w < RR_SO_List.Count; w++)
                {
                    oSheet.Cells[LineIns, 1] = RR_ОСНТ_List[w];
                    oSheet.Cells[LineIns, 2] = RR_Prev_ОСНТ_List[w];
                    oSheet.Cells[LineIns, 3] = RR_Days_List[w];
                    oSheet.Cells[LineIns, 4] = RR_Prev_Date_List[w];
                    oSheet.Cells[LineIns, 5] = RR_SO_List[w];

                    LineIns++;
                }


                oWB.SaveAs(savename, Excel.XlFileFormat.xlWorkbookNormal, m_objOpt,
                    m_objOpt, m_objOpt, m_objOpt, Excel.XlSaveAsAccessMode.xlExclusive,
                    m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);

                oWB.Close(false, m_objOpt, m_objOpt);

                oXL.Quit();

                Marshal.ReleaseComObject(oSheet);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oXL);
                oSheet = null;
                oWB = null;
                oXL = null;
                GC.Collect();
            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = oldCulture;
            }

            StaticFunction.ClearExcelProcess(BeforeList, AfterList);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

        private void button4_Click(object sender, EventArgs e)
        {

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = openFileDialog1.FileName;
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = openFileDialog1.FileName;
            }
        }
    }
}
