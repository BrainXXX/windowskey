using System;
using System.IO;
using System.Net;
using System.Text;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Reflection;
using System.Diagnostics;

using System.Runtime.InteropServices;
using System.Threading;
using System.Globalization;


namespace NewsGen
{
    class StaticFunction
    {
        /*
                public static void KillExcel()
                {

                    try
                    {
                        System.Diagnostics.Process[] processArr2 = System.Diagnostics.Process.GetProcessesByName("Excel");

                        for (int process2cnt = 0; process2cnt < processArr2.Length; process2cnt++)
                        {
                            processArr2[process2cnt].Kill();
                        }
                    }
                    catch
                    {

                    }
                }
        */

        public static string ConvertToColumnName(int ColumnIndex)
        {
            /// начало с 1 (1=А)
            /// 

            string result;

            switch (ColumnIndex)
            {
                case 1: result = "A"; break;
                case 2: result = "B"; break;
                case 3: result = "C"; break;
                case 4: result = "D"; break;
                case 5: result = "E"; break;
                case 6: result = "F"; break;
                case 7: result = "G"; break;
                case 8: result = "H"; break;
                case 9: result = "I"; break;
                case 10: result = "J"; break;
                case 11: result = "K"; break;
                case 12: result = "L"; break;
                case 13: result = "M"; break;
                case 14: result = "N"; break;
                case 15: result = "O"; break;
                case 16: result = "P"; break;
                case 17: result = "Q"; break;
                case 18: result = "R"; break;
                case 19: result = "S"; break;
                case 20: result = "T"; break;
                case 21: result = "U"; break;
                case 22: result = "V"; break;
                case 23: result = "W"; break;
                case 24: result = "X"; break;
                case 25: result = "Y"; break;
                case 26: result = "Z"; break;

                case 27: result = "AA"; break;
                case 28: result = "AB"; break;
                case 29: result = "AC"; break;
                case 30: result = "AD"; break;
                case 31: result = "AE"; break;
                case 32: result = "AF"; break;
                case 33: result = "AG"; break;
                case 34: result = "AH"; break;
                case 35: result = "AI"; break;
                case 36: result = "AJ"; break;
                case 37: result = "AK"; break;
                case 38: result = "AL"; break;
                case 39: result = "AM"; break;
                case 40: result = "AN"; break;
                case 41: result = "AO"; break;
                case 42: result = "AP"; break;
                case 43: result = "AQ"; break;
                case 44: result = "AR"; break;
                case 45: result = "AS"; break;
                case 46: result = "AT"; break;
                case 47: result = "AU"; break;
                case 48: result = "AV"; break;
                case 49: result = "AW"; break;
                case 50: result = "AX"; break;
                case 51: result = "AY"; break;
                case 52: result = "AZ"; break;

                case 53: result = "BA"; break;
                case 54: result = "BB"; break;
                case 55: result = "BC"; break;
                case 56: result = "BD"; break;
                case 57: result = "BE"; break;
                case 58: result = "BF"; break;
                case 59: result = "BG"; break;
                case 60: result = "BH"; break;
                case 61: result = "BI"; break;
                case 62: result = "BJ"; break;
                case 63: result = "BK"; break;
                case 64: result = "BL"; break;
                case 65: result = "BM"; break;
                case 66: result = "BN"; break;
                case 67: result = "BO"; break;
                case 68: result = "BP"; break;
                case 69: result = "BQ"; break;
                case 70: result = "BR"; break;
                case 71: result = "BS"; break;
                case 72: result = "BT"; break;
                case 73: result = "BU"; break;
                case 74: result = "BV"; break;
                case 75: result = "BW"; break;
                case 76: result = "BX"; break;
                case 77: result = "BY"; break;
                case 78: result = "BZ"; break;

                case 79: result = "CA"; break;
                case 80: result = "CB"; break;
                case 81: result = "CC"; break;
                case 82: result = "CD"; break;
                case 83: result = "CE"; break;
                case 84: result = "CF"; break;
                case 85: result = "CG"; break;
                case 86: result = "CH"; break;
                case 87: result = "CI"; break;
                case 88: result = "CJ"; break;
                case 89: result = "CK"; break;
                case 90: result = "CL"; break;
                case 91: result = "CM"; break;
                case 92: result = "CN"; break;
                case 93: result = "CO"; break;
                case 94: result = "CP"; break;
                case 95: result = "CQ"; break;
                case 96: result = "CR"; break;
                case 97: result = "CS"; break;
                case 98: result = "CT"; break;
                case 99: result = "CU"; break;
                case 100: result = "CV"; break;
                case 101: result = "CW"; break;
                case 102: result = "CX"; break;
                case 103: result = "CY"; break;
                case 104: result = "CZ"; break;

                case 105: result = "DA"; break;
                case 106: result = "DB"; break;
                case 107: result = "DC"; break;
                case 108: result = "DD"; break;
                case 109: result = "DE"; break;
                case 110: result = "DF"; break;
                case 111: result = "DG"; break;
                case 112: result = "DH"; break;
                case 113: result = "DI"; break;
                case 114: result = "DJ"; break;
                case 115: result = "DK"; break;
                case 116: result = "DL"; break;
                case 117: result = "DM"; break;
                case 118: result = "DN"; break;
                case 119: result = "DO"; break;
                case 120: result = "DP"; break;
                case 121: result = "DQ"; break;
                case 122: result = "DR"; break;
                case 123: result = "DS"; break;
                case 124: result = "DT"; break;
                case 125: result = "DU"; break;
                case 126: result = "DV"; break;
                case 127: result = "DW"; break;
                case 128: result = "DX"; break;
                case 129: result = "DY"; break;
                case 130: result = "DZ"; break;



                case 131: result = "EA"; break;
                case 132: result = "EB"; break;
                case 133: result = "EC"; break;
                case 134: result = "ED"; break;
                case 135: result = "EE"; break;
                case 136: result = "EF"; break;
                case 137: result = "EG"; break;
                case 138: result = "EH"; break;
                case 139: result = "EI"; break;
                case 140: result = "EJ"; break;
                case 141: result = "EK"; break;
                case 142: result = "EL"; break;
                case 143: result = "EM"; break;
                case 144: result = "EN"; break;
                case 145: result = "EO"; break;
                case 146: result = "EP"; break;
                case 147: result = "EQ"; break;
                case 148: result = "ER"; break;
                case 149: result = "ES"; break;
                case 150: result = "ET"; break;
                case 151: result = "EU"; break;
                case 152: result = "EV"; break;
                case 153: result = "EW"; break;
                case 154: result = "EX"; break;
                case 155: result = "EY"; break;
                case 156: result = "EZ"; break;


                case 157: result = "FA"; break;
                case 158: result = "FB"; break;
                case 159: result = "FC"; break;
                case 160: result = "FD"; break;
                case 161: result = "FE"; break;
                case 162: result = "FF"; break;
                case 163: result = "FG"; break;
                case 164: result = "FH"; break;
                case 165: result = "FI"; break;
                case 166: result = "FJ"; break;
                case 167: result = "FK"; break;
                case 168: result = "FL"; break;
                case 169: result = "FM"; break;
                case 170: result = "FN"; break;
                case 171: result = "FO"; break;
                case 172: result = "FP"; break;
                case 173: result = "FQ"; break;
                case 174: result = "FR"; break;
                case 175: result = "FS"; break;
                case 176: result = "FT"; break;
                case 177: result = "FU"; break;
                case 178: result = "FV"; break;
                case 179: result = "FW"; break;
                case 180: result = "FX"; break;
                case 181: result = "FY"; break;
                case 182: result = "FZ"; break;

                default:


                    result = "";
                    break;
            }

            return result;

        }


        public static ArrayList GetExcelProcessID()
        {
            ArrayList IDList = new ArrayList();

            try
            {
                System.Diagnostics.Process[] processArr2 = System.Diagnostics.Process.GetProcessesByName("Excel");
                for (int process2cnt = 0; process2cnt < processArr2.Length; process2cnt++)
                {
                    int processid = processArr2[process2cnt].Id;
                    IDList.Add(processid);
                }
            }
            catch
            {
                return null;
            }

            return IDList;
        }

        public static void ClearExcelProcess(ArrayList BeforeList, ArrayList AfterList)
        {
            try
            {

                for (int k = 0; k < AfterList.Count; k++)
                {
                    int ProcessID = (int)AfterList[k];


                    if (!BeforeList.Contains(ProcessID))
                    {
                        System.Diagnostics.Process ProcToKill = System.Diagnostics.Process.GetProcessById(ProcessID);
                        ProcToKill.Kill();
                    }

                }
            }
            catch
            {
                return;
            }
        }

        public static string Translit(string source_string)
        {
            string str_tr = source_string;

            str_tr = str_tr.Replace("а", "a");
            str_tr = str_tr.Replace("б", "b");
            str_tr = str_tr.Replace("в", "v");
            str_tr = str_tr.Replace("г", "g");
            str_tr = str_tr.Replace("д", "d");
            str_tr = str_tr.Replace("е", "e");
            str_tr = str_tr.Replace("ё", "yo");
            str_tr = str_tr.Replace("ж", "zh");
            str_tr = str_tr.Replace("з", "z");
            str_tr = str_tr.Replace("и", "i");
            str_tr = str_tr.Replace("й", "j");
            str_tr = str_tr.Replace("к", "k");
            str_tr = str_tr.Replace("л", "l");
            str_tr = str_tr.Replace("м", "m");
            str_tr = str_tr.Replace("н", "n");
            str_tr = str_tr.Replace("о", "o");
            str_tr = str_tr.Replace("п", "p");
            str_tr = str_tr.Replace("р", "r");
            str_tr = str_tr.Replace("с", "s");
            str_tr = str_tr.Replace("т", "t");
            str_tr = str_tr.Replace("у", "u");
            str_tr = str_tr.Replace("ф", "f");
            str_tr = str_tr.Replace("х", "h");
            str_tr = str_tr.Replace("ц", "c");
            str_tr = str_tr.Replace("ч", "ch");
            str_tr = str_tr.Replace("ш", "sh");
            str_tr = str_tr.Replace("щ", "sch");
            str_tr = str_tr.Replace("ъ", "j");
            str_tr = str_tr.Replace("ы", "i");
            str_tr = str_tr.Replace("ь", "j");
            str_tr = str_tr.Replace("э", "e");
            str_tr = str_tr.Replace("ю", "yu");
            str_tr = str_tr.Replace("я", "ya");
            str_tr = str_tr.Replace("А", "A");
            str_tr = str_tr.Replace("Б", "B");
            str_tr = str_tr.Replace("В", "V");
            str_tr = str_tr.Replace("Г", "G");
            str_tr = str_tr.Replace("Д", "D");
            str_tr = str_tr.Replace("Е", "E");
            str_tr = str_tr.Replace("Ё", "Yo");
            str_tr = str_tr.Replace("Ж", "Zh");
            str_tr = str_tr.Replace("З", "Z");
            str_tr = str_tr.Replace("И", "I");
            str_tr = str_tr.Replace("Й", "J");
            str_tr = str_tr.Replace("К", "K");
            str_tr = str_tr.Replace("Л", "L");
            str_tr = str_tr.Replace("М", "M");
            str_tr = str_tr.Replace("Н", "N");
            str_tr = str_tr.Replace("О", "O");
            str_tr = str_tr.Replace("П", "P");
            str_tr = str_tr.Replace("Р", "R");
            str_tr = str_tr.Replace("С", "S");
            str_tr = str_tr.Replace("Т", "T");
            str_tr = str_tr.Replace("У", "U");
            str_tr = str_tr.Replace("Ф", "F");
            str_tr = str_tr.Replace("Х", "H");
            str_tr = str_tr.Replace("Ц", "C");
            str_tr = str_tr.Replace("Ч", "CH");
            str_tr = str_tr.Replace("Ш", "SH");
            str_tr = str_tr.Replace("Щ", "SCH");
            str_tr = str_tr.Replace("Ъ", "J");
            str_tr = str_tr.Replace("Ы", "I");
            str_tr = str_tr.Replace("Ь", "J");
            str_tr = str_tr.Replace("Э", "E");
            str_tr = str_tr.Replace("Ю", "YU");
            str_tr = str_tr.Replace("Я", "YA");

            return str_tr;
        }
    }
}
