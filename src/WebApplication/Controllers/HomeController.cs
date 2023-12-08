using Syncfusion.EJ2.Base;
using Syncfusion.EJ2.Linq;
using System;
using System.Collections;
using System.Linq;
using System.Web.Mvc;
using WebApplication.Models;
using ClosedXML.Excel;
using System.IO;
using System.Collections.Generic;

namespace WebApplication.Controllers
{
    public class HomeController : Controller
    {
        private ProcurementInfoEntities _databaseRepository;

        public HomeController()
        {
            _databaseRepository = new ProcurementInfoEntities();
            _databaseRepository.Configuration.LazyLoadingEnabled = false;
        }

        #region Формирование отчета.
        public ActionResult Export(DateTime dateS, DateTime dateF)
        {
            DateTime err = new DateTime(1940, 01, 31);
            if (dateS == err && dateF != err || dateS != err && dateF == err)
            {
                if(dateS != err)
                {
                    dateF = DateTime.Now;
                }
                else
                {
                    dateS = new DateTime(1995, 01, 01);
                }
            }
            else if (dateS == err && dateF == err)
            {
                dateS = new DateTime(1995, 01, 01);
                dateF = DateTime.Now;
            }
            if (dateS > dateF)
            {
                var buf = dateS;
                dateS = dateF;
                dateF = buf;
            }

            string xsltPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath(@"~/images/TemplateReport.xlsx"));
            var workbook = new XLWorkbook(xsltPath);
            var worksheet_1 = workbook.Worksheet(1);
            var worksheet_2 = workbook.Worksheet(2);
            var worksheet_3 = workbook.Worksheet(3);
            int cef1_1 = 0, cef1_2 = 0, cef3_1 = 0, cef3_2 = 0, cef4_1 = 0, cef4_2 = 0, sa=1;
            double cef2_1 = 0, cef2_2 = 0, cef5_1 = 0, cef5_2 = 0, cef6_1 = 0, cef6_2 = 0;
            ProcurementInfoEntities context = new ProcurementInfoEntities();
            var table = context.T_ProcurementInformation.Where(c => ((c.PR_APPROVAL_DATE >= dateS & c.PR_APPROVAL_DATE <= dateF) || (c.PR_APPROVAL_DATE <= dateS & c.CONTRACT_DATE_FACT <= dateF & c.CONTRACT_DATE_FACT >= dateS)) & c.ID_SUBJECTPURCHASE == 1 & c.FLAG == true);
            int i = 3, j = 2;
            foreach (var c in table)
            {
                worksheet_2.Rows(Convert.ToString(i)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet_2.Cell("A" + i).Value = c.NAME_PR;

                if (c.ID_SUBJECTPURCHASE == 1)
                {
                    worksheet_2.Cell("B" + i).Value = "Товары";
                }
                else if (c.ID_SUBJECTPURCHASE == 2)
                {
                    worksheet_2.Cell("B" + i).Value = "Работы (услуги)";
                }
                else
                {
                    worksheet_2.Cell("B" + i).Value = c.ID_SUBJECTPURCHASE;
                }
                worksheet_2.Cell("C" + i).Value = c.INVITE_NUN;
                worksheet_2.Cell("D" + i).Value = c.FIO_ISP;

                if (c.ID_LEGISLATION == 1)
                {
                    worksheet_2.Cell("E" + i).Value = "Пост.229 - Конкурс";
                }
                else if (c.ID_LEGISLATION == 2)
                {
                    worksheet_2.Cell("E" + i).Value = "Строительство";
                }
                else if (c.ID_LEGISLATION == 3)
                {
                    worksheet_2.Cell("E" + i).Value = "Биржа";
                }
                else if (c.ID_LEGISLATION == 4)
                {
                    worksheet_2.Cell("E" + i).Value = "Пост. 229 - Приложение 1 (пр.договор)";
                }
                else if (c.ID_LEGISLATION == 6)
                {
                    worksheet_2.Cell("E" + i).Value = "Пост. 229 - Закупка из 1 источника";
                }
                else
                {
                    worksheet_2.Cell("E" + i).Value = c.ID_LEGISLATION;
                }

                worksheet_2.Cell("F" + i).Value = c.PARTICIPANT__VALUE;
                worksheet_2.Cell("G" + i).Value = c.PR_APPROVAL_DATE;

                if (c.ID_RESULT == 1)
                {
                    worksheet_2.Cell("H" + i).Value = "Cостоялась";
                }
                else if (c.ID_RESULT == 2)
                {
                    worksheet_2.Cell("H" + i).Value = "Не состоялась";
                }
                else if (c.ID_RESULT == 4)
                {
                    worksheet_2.Cell("H" + i).Value = "Отменена";
                }
                else
                {
                    worksheet_2.Cell("H" + i).Value = c.ID_RESULT;
                }

                worksheet_2.Cell("I" + i).Value = c.CONTRACT_DATE_TERM;
                worksheet_2.Cell("J" + i).Value = c.CONTRACT_DATE_PROL;
                worksheet_2.Cell("K" + i).Value = c.CONTRACT_DATE_FACT;
                worksheet_2.Cell("L" + i).Value = c.INVITE_NUN;
                worksheet_2.Cell("M" + i).Value = c.RESULT_DATE_TERM;
                worksheet_2.Cell("N" + i).Value = c.RESULT_DATE_FACT;
                worksheet_2.Cell("O" + i).Value = c.WIN_NAME;

                if (c.ID_COUNTRY != null)
                {
                    var i_c_s = context.S_Country.Where(p => p.ID == c.ID_COUNTRY);
                    foreach (var x in i_c_s)
                    {
                        worksheet_2.Cell("P" + i).Value = x.NAME;
                    }
                }
                else
                {
                    worksheet_2.Cell("P" + i).Value = c.ID_COUNTRY;
                }

                worksheet_2.Cell("Q" + i).Value = c.WIN_VALUE;

                if (c.ID_CURRENCY == 1)
                {
                    worksheet_2.Cell("R" + i).Value = "BYN";
                }
                else if (c.ID_CURRENCY == 2)
                {
                    worksheet_2.Cell("R" + i).Value = "USD";
                }
                else if (c.ID_CURRENCY == 3)
                {
                    worksheet_2.Cell("R" + i).Value = "EUR";
                }
                else if (c.ID_CURRENCY == 4)
                {
                    worksheet_2.Cell("R" + i).Value = "RUB";
                }
                else
                {
                    worksheet_2.Cell("R" + i).Value = c.ID_CURRENCY;
                }

                worksheet_2.Cell("S" + i).Value = c.WIN_VALUE_BYN;

                if (c.ID_WINSTATUS == 1)
                {
                    worksheet_2.Cell("T" + i).Value = "Производитель";
                }
                else if (c.ID_WINSTATUS == 2)
                {
                    worksheet_2.Cell("T" + i).Value = "Посредник";
                }
                else if (c.ID_WINSTATUS == 3)
                {
                    worksheet_2.Cell("T" + i).Value = "Официальный представитель";
                }
                else
                {
                    worksheet_2.Cell("T" + i).Value = c.ID_WINSTATUS;
                }

                worksheet_2.Cell("U" + i).Value = c.PURCHASE_VOLUME;
                worksheet_2.Cell("V" + i).Value = c.VOLUME_UNITS;
                worksheet_2.Cell("W" + i).Value = c.PRICE_PER_ITEM;
                worksheet_2.Cell("X" + i).Value = c.DELIVERY_COND;
                worksheet_2.Cell("Y" + i).Value = c.MAIN_TEC_SPECS;

                if (c.ID_COUNTRY_ORIGIN != null)
                {
                    var i_c_s = context.S_Country.Where(p => p.ID == c.ID_COUNTRY_ORIGIN);
                    foreach (var x in i_c_s)
                    {
                        worksheet_2.Cell("Z" + i).Value = x.NAME;
                    }
                }
                else
                {
                    worksheet_2.Cell("Z" + i).Value = c.ID_COUNTRY_ORIGIN;
                }

                worksheet_2.Cell("AA" + i).Value = c.MANUFACTURER;
                i++;

                if (c.ID_SUBJECTPURCHASE == 1 & c.ID_LEGISLATION == 1 & c.ID_RESULT == 1 & c.CONTRACT_DATE_FACT != null)
                {
                    cef1_1++;
                }
                
                if (c.ID_SUBJECTPURCHASE == 1 & (c.ID_LEGISLATION == 1 | c.ID_LEGISLATION == 6) & c.ID_RESULT != 4)
                {
                    cef1_2++;
                }

                if (c.ID_SUBJECTPURCHASE == 1 & c.ID_LEGISLATION == 1 & c.ID_RESULT == 1 & c.CONTRACT_DATE_FACT != null)
                {
                    cef2_1 = cef2_1 + Convert.ToDouble(c.WIN_VALUE_BYN);
                }

                if (c.ID_SUBJECTPURCHASE == 1 & c.ID_RESULT == 1 & (c.ID_LEGISLATION == 1 | c.ID_LEGISLATION == 4| c.ID_LEGISLATION == 6))
                {
                    cef2_2 = cef2_2 + Convert.ToDouble(c.WIN_VALUE_BYN);
                }

                if (c.ID_SUBJECTPURCHASE == 1 & c.ID_RESULT == 1 & c.CONTRACT_DATE_FACT != null & c.ID_LEGISLATION == 1)
                {
                    cef3_1 = cef3_1 + Convert.ToInt32(c.PARTICIPANT__VALUE);
                }

                if (c.ID_SUBJECTPURCHASE == 1 & c.ID_RESULT == 1 & c.ID_LEGISLATION == 1 & c.CONTRACT_DATE_FACT != null)
                {
                    cef3_2++;
                }

                if (c.ID_SUBJECTPURCHASE == 1 & c.ID_RESULT == 1 & (c.ID_WINSTATUS == 3 | c.ID_WINSTATUS == 1) & (c.ID_LEGISLATION == 6 | c.ID_LEGISLATION == 1) & (c.CONTRACT_DATE_FACT >= dateS & c.CONTRACT_DATE_FACT <= dateF))
                {
                    cef4_1++;
                }

                if (c.ID_SUBJECTPURCHASE == 1 & c.ID_RESULT == 1 & c.CONTRACT_DATE_FACT != null & (c.ID_LEGISLATION == 1 | c.ID_LEGISLATION == 6) & (c.CONTRACT_DATE_FACT >= dateS & c.CONTRACT_DATE_FACT <= dateF))
                {
                    cef4_2++;
                }

                if (c.ID_SUBJECTPURCHASE == 1 & c.ID_RESULT == 1 & (c.ID_WINSTATUS == 3 | c.ID_WINSTATUS == 1) & (c.ID_LEGISLATION == 1 | c.ID_LEGISLATION == 6) & (c.CONTRACT_DATE_FACT >= dateS & c.CONTRACT_DATE_FACT <= dateF))
                {
                    cef5_1 = cef5_1 + Convert.ToDouble(c.WIN_VALUE_BYN);
                }

                if (c.ID_RESULT == 1 & c.CONTRACT_DATE_FACT != null & (c.ID_LEGISLATION == 1 | c.ID_LEGISLATION == 6) & (c.CONTRACT_DATE_FACT >= dateS & c.CONTRACT_DATE_FACT <= dateF))
                {
                    cef5_2 = cef5_2 + Convert.ToDouble(c.WIN_VALUE_BYN);
                }

                if (c.ID_RESULT == 1 & c.ID_SUBJECTPURCHASE == 1 & c.ID_COUNTRY_ORIGIN == 94 & c.CONTRACT_DATE_FACT != null)
                {
                    cef6_1 = cef6_1 + Convert.ToDouble(c.WIN_VALUE_BYN);
                }

                if (c.ID_SUBJECTPURCHASE == 1 & c.CONTRACT_DATE_FACT != null)
                {
                    cef6_2 = cef6_2 + Convert.ToDouble(c.WIN_VALUE_BYN);
                }
            }
            var tabledo1k = context.T_PrInfoDo1000.Where(c => c.DATE_CONCLUSION >= dateS & c.DATE_CONCLUSION <= dateF & c.ID_SUBJECTPURCHASE == 1 & c.FLAG == true);
            foreach (var c in tabledo1k)
            {
                worksheet_3.Row(j).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet_3.Cell("A" + j).Value = c.NAME_PR;
                if (c.ID_SUBJECTPURCHASE == 1)
                {
                    worksheet_3.Cell("B" + j).Value = "Товары";
                }
                else if (c.ID_SUBJECTPURCHASE == 2)
                {
                    worksheet_3.Cell("B" + j).Value = "Работы (услуги)";
                }
                else
                {
                    worksheet_3.Cell("B" + j).Value = c.ID_SUBJECTPURCHASE;
                }

                worksheet_3.Cell("C" + j).Value = c.CONTRACT_NUMBER;
                worksheet_3.Cell("D" + j).Value = c.FIO_ISP;

                if (c.ID_WINSTATUSDO1000 == 1)
                {
                    worksheet_3.Cell("E" + j).Value = "Производитель";
                }
                else if (c.ID_WINSTATUSDO1000 == 2)
                {
                    worksheet_3.Cell("E" + j).Value = "Посредник";
                }
                else if (c.ID_WINSTATUSDO1000 == 3)
                {
                    worksheet_3.Cell("E" + j).Value = "Официальный представитель";
                }
                else
                {
                    worksheet_3.Cell("E" + j).Value = c.ID_WINSTATUSDO1000;
                }

                if (c.ID_LEGISLDO1000 == 1)
                {
                    worksheet_3.Cell("F" + j).Value = "Прямой договор";
                }
                else if (c.ID_LEGISLDO1000 == 2)
                {
                    worksheet_3.Cell("F" + j).Value = "Сравнительный анализ";
                }
                else if (c.ID_LEGISLDO1000 == 3)
                {
                    worksheet_3.Cell("F" + j).Value = "Биржевые торги";
                }
                else
                {
                    worksheet_3.Cell("F" + j).Value = c.ID_LEGISLDO1000;
                }

                worksheet_3.Cell("G" + j).Value = c.DATE_CONCLUSION;
                worksheet_3.Cell("H" + j).Value = c.WIN_VALUE;

                if (c.ID_CURRENCY == 1)
                {
                    worksheet_3.Cell("I" + j).Value = "BYN";
                }
                else if (c.ID_CURRENCY == 2)
                {
                    worksheet_3.Cell("I" + j).Value = "USD";
                }
                else if (c.ID_CURRENCY == 3)
                {
                    worksheet_3.Cell("I" + j).Value = "EUR";
                }
                else if (c.ID_CURRENCY == 4)
                {
                    worksheet_3.Cell("I" + j).Value = "RUB";
                }
                else
                {
                    worksheet_3.Cell("I" + j).Value = c.ID_CURRENCY;
                }

                worksheet_3.Cell("J" + j).Value = c.WIN_VALUE_NDE;
                worksheet_3.Cell("K" + j).Value = c.WIN_VALUE_NNDS;
                worksheet_3.Cell("L" + j).Value = c.WIN_NAME;

                if (c.ID_COUNTRY_ORIGIN != null)
                {
                    var i_c_s = context.S_Country.Where(p => p.ID == c.ID_COUNTRY_ORIGIN);
                    foreach (var x in i_c_s)
                    {
                        worksheet_3.Cell("M" + j).Value = x.NAME;
                    }
                }
                else
                {
                    worksheet_3.Cell("M" + j).Value = c.ID_COUNTRY_ORIGIN;
                }

                j++;

                if (c.ID_SUBJECTPURCHASE == 1 & c.ID_LEGISLDO1000 == 1)
                {
                    cef2_2 = cef2_2 + Convert.ToDouble(c.WIN_VALUE_NNDS);
                }

                if (c.ID_SUBJECTPURCHASE == 1 & c.ID_COUNTRY_ORIGIN == 94)
                {
                    cef6_1 = cef6_1 + Convert.ToDouble(c.WIN_VALUE_NNDS);
                }

                if (c.ID_SUBJECTPURCHASE == 1)
                {
                    cef6_2 = cef6_2 + Convert.ToDouble(c.WIN_VALUE_NNDS);
                }

                /*worksheet_4.Cell("B2").Value = cef1_1;
                worksheet_4.Cell("B3").Value = cef1_2;
                worksheet_4.Cell("B4").Value = cef2_1;
                worksheet_4.Cell("B5").Value = cef2_2;
                worksheet_4.Cell("B6").Value = cef3_1;
                worksheet_4.Cell("B7").Value = cef3_2;
                worksheet_4.Cell("B8").Value = cef4_1;
                worksheet_4.Cell("B9").Value = cef4_2;
                worksheet_4.Cell("B10").Value = cef5_1;
                worksheet_4.Cell("B11").Value = cef5_2;
                worksheet_4.Cell("B12").Value = cef6_1;
                worksheet_4.Cell("B13").Value = cef6_2;*/
            }

            /* int col_p1_1 = 0, col_p1_2 = 0, col_p3_1 = 0, col_p3_2 = 0, col_p4_1 = 0, col_p4_2 = 0;
             double sum_p2_1 = 0, sum_p2_2 = 0, sum_p5_1 = 0, sum_p5_2 = 0, sum_p6_1 = 0, sum_p6_2 = 0;

             for (int o = 3; o < i; o++)
             {
                 var cell = worksheet_2.Cell("B" + o);
                 var cell_res = cell.GetString();
                 cell = worksheet_2.Cell("E" + o);
                 var cell_konk = cell.GetString();
                 cell = worksheet_2.Cell("H" + o);
                 var cell_sost = cell.GetString();
                 cell = worksheet_2.Cell("T" + o);
                 var cell_stat = cell.GetString();

                 if (String.Compare(cell_res, "Товары") == 0 & String.Compare(cell_konk, "Пост.229 - Конкурс") == 0 & String.Compare(cell_sost, "Cостоялась") == 0 )
                 {
                     col_p1_1++;
                 }

                 if (String.Compare(cell_res, "Товары") == 0 & (String.Compare(cell_konk, "Пост.229 - Конкурс") == 0 | String.Compare(cell_konk, "Пост. 229 - Закупка из 1 источника") == 0))
                 {
                     col_p1_2++;
                 }

                 if (String.Compare(cell_res, "Товары") == 0 & String.Compare(cell_konk, "Пост.229 - Конкурс") == 0 & String.Compare(cell_sost, "Cостоялась") == 0)
                 {
                     cell = worksheet_2.Cell("S" + o);
                     var cell_st = cell.GetString();
                     if (cell_st != "")
                     {
                         sum_p2_1 = sum_p2_1 + Convert.ToDouble(cell_st);
                     }
                 }

                 if (String.Compare(cell_sost, "Cостоялась") == 0 & String.Compare(cell_res, "Товары") == 0 & (String.Compare(cell_konk, "Пост.229 - Конкурс") == 0 | String.Compare(cell_konk, "Пост. 229 - Приложение 1 (пр.договор)") == 0))
                 {
                     cell = worksheet_2.Cell("S" + o);
                     var cell_st = cell.GetString();
                     if (cell_st != "")
                     {
                         sum_p2_2 = sum_p2_2 + Convert.ToDouble(cell_st);
                     }
                 }
                 else if (String.Compare(cell_sost, "Cостоялась") == 0 & String.Compare(cell_res, "Товары") == 0 & String.Compare(cell_konk, "Пост. 229 - Закупка из 1 источника") == 0 )
                 {
                     cell = worksheet_2.Cell("S" + o);
                     var cell_st = cell.GetString();
                     if (cell_st != "")
                     {
                         sum_p2_2 = sum_p2_2 + Convert.ToDouble(cell_st);
                     }
                 }

                 if (String.Compare(cell_sost, "Cостоялась") == 0 & String.Compare(cell_konk, "Пост.229 - Конкурс") == 0)
                 {
                     cell = worksheet_2.Cell("F" + o);
                     var cell_st = cell.GetString();
                     if (cell_st != "")
                     {
                         col_p3_1 = col_p3_1 + Convert.ToInt32(cell_st);
                     }
                 }
                 if (String.Compare(cell_sost, "Cостоялась") == 0 & String.Compare(cell_konk, "Пост.229 - Конкурс") == 0)
                 {
                     col_p3_2++;
                 }

                 if ((String.Compare(cell_sost, "Cостоялась") == 0 & String.Compare(cell_konk, "Пост.229 - Конкурс") == 0 | String.Compare(cell_konk, "Пост. 229 - Закупка из 1 источника") == 0) & (String.Compare(cell_stat, "Официальный представитель") == 0))
                 {
                     col_p4_1++;
                 }

                 cell = worksheet_2.Cell("K" + o);
                 var date = cell.GetString();
                 if ((String.Compare(cell_konk, "Пост.229 - Конкурс") == 0 | String.Compare(cell_konk, "Пост. 229 - Закупка из 1 источника") == 0) & String.Compare(cell_sost, "Cостоялась") == 0 & date != "")
                 {
                     col_p4_2++;
                 }

                 if ((String.Compare(cell_sost, "Cостоялась") == 0 & String.Compare(cell_konk, "Пост.229 - Конкурс") == 0 | String.Compare(cell_konk, "Пост. 229 - Закупка из 1 источника") == 0) & (String.Compare(cell_stat, "Официальный представитель") == 0))
                 {
                     cell = worksheet_2.Cell("S" + o);
                     var cell_st = cell.GetString();
                     if (cell_st != "")
                     {
                         sum_p5_1 = sum_p5_1 + Convert.ToDouble(cell_st);
                     }
                 }

                 if ((String.Compare(cell_konk, "Пост.229 - Конкурс") == 0 | String.Compare(cell_konk, "Пост. 229 - Закупка из 1 источника") == 0) & String.Compare(cell_sost, "Cостоялась") == 0 & date != "")
                 {
                     cell = worksheet_2.Cell("S" + o);
                     var cell_st = cell.GetString();
                     if (cell_st != "")
                     {
                         sum_p5_2 = sum_p5_2 + Convert.ToDouble(cell_st);
                     }
                 }

                 cell = worksheet_2.Cell("Z" + o);
                 var country = cell.GetString();
                 if (String.Compare(cell_sost, "Cостоялась") == 0 & String.Compare(cell_res, "Товары") == 0 & String.Compare(country, "БЕЛАРУСЬ") == 0 & date != "")
                 {
                     if (Convert.ToDateTime(date) <= dateF & Convert.ToDateTime(date) >= dateS)
                     {
                         cell = worksheet_2.Cell("S" + o);
                         var cell_st = cell.GetString();
                         if (cell_st != "")
                         {
                             sum_p6_1 = sum_p6_1 + Convert.ToDouble(cell_st);
                         }
                     }
                 }

                 if (String.Compare(cell_res, "Товары") == 0 & date != "")
                 {
                     if (Convert.ToDateTime(date) <= dateF & Convert.ToDateTime(date) >= dateS)
                     {
                         cell = worksheet_2.Cell("S" + o);
                         var cell_st = cell.GetString();
                         if (cell_st != "")
                         {
                             sum_p6_2 = sum_p6_2 + Convert.ToDouble(cell_st);
                         }
                     }
                 }
             }

             for (int k = 2; k < j; k++)
             {
                 var cell = worksheet_3.Cell("F" + k);
                 var cell_prdo1k = cell.GetString();
                 cell = worksheet_3.Cell("E" + k);
                 var cell_statdo1k = cell.GetString();
                 if (String.Compare(cell_prdo1k, "Прямой договор") == 0)
                 {
                     cell = worksheet_3.Cell("K" + k);
                     var cell_st = cell.GetString();
                     if (cell_st != "")
                     {
                         sum_p2_2 = sum_p2_2 + Convert.ToDouble(cell_st);
                     }
                 }

                 cell = worksheet_3.Cell("M" + k);
                 var countrydo1k = cell.GetString();
                 cell = worksheet_3.Cell("B" + k);
                 var preddo1k = cell.GetString();
                 if (String.Compare(countrydo1k, "БЕЛАРУСЬ") == 0 & String.Compare(preddo1k, "Товары") == 0)
                 {
                     cell = worksheet_3.Cell("K" + k);
                     var cell_st = cell.GetString();
                     if (cell_st != "")
                     {
                         sum_p6_1 = sum_p6_1 + Convert.ToDouble(cell_st);
                     }
                 }

                 if (String.Compare(preddo1k, "Товары") == 0)
                 {
                     cell = worksheet_3.Cell("K" + k);
                     var cell_st = cell.GetString();
                     if (cell_st != "")
                     {
                         sum_p6_2 = sum_p6_2 + Convert.ToDouble(cell_st);
                     }
                 }
             }

             worksheet_1.Cell("H7").Value = col_p1_1;
             worksheet_1.Cell("H8").Value = col_p1_2;
             worksheet_1.Cell("H9").Value = sum_p2_1;
             worksheet_1.Cell("H10").Value = sum_p2_2;
             worksheet_1.Cell("H11").Value = col_p3_1;
             worksheet_1.Cell("H12").Value = col_p3_2;
             worksheet_1.Cell("H13").Value = col_p4_1;
             worksheet_1.Cell("H14").Value = col_p4_2;
             worksheet_1.Cell("H15").Value = sum_p5_1;
             worksheet_1.Cell("H16").Value = sum_p5_2;
             worksheet_1.Cell("H17").Value = sum_p6_1;
             worksheet_1.Cell("H18").Value = sum_p6_2;

             worksheet_1.Cell("C4").Value = (double)col_p1_1 / col_p1_2;
             worksheet_1.Cell("D4").Value = (double)sum_p2_1 / sum_p2_2;
             worksheet_1.Cell("E4").Value = (double)col_p3_1 / col_p3_2;
             worksheet_1.Cell("F4").Value = (double)col_p4_1 / col_p4_2;
             worksheet_1.Cell("G4").Value = (double)sum_p5_1 / sum_p5_2;
             worksheet_1.Cell("H4").Value = (double)sum_p6_1 / sum_p6_2;*/

            worksheet_1.Cell("H7").Value = cef1_1;
            worksheet_1.Cell("H8").Value = cef1_2;
            worksheet_1.Cell("H9").Value = cef2_1;
            worksheet_1.Cell("H10").Value = cef2_2;
            worksheet_1.Cell("H11").Value = cef3_1;
            worksheet_1.Cell("H12").Value = cef3_2;
            worksheet_1.Cell("H13").Value = cef4_1;
            worksheet_1.Cell("H14").Value = cef4_2;
            worksheet_1.Cell("H15").Value = cef5_1;
            worksheet_1.Cell("H16").Value = cef5_2;
            worksheet_1.Cell("H17").Value = cef6_1;
            worksheet_1.Cell("H18").Value = cef6_2;

            worksheet_1.Cell("C4").Value = (double)cef1_1 / cef1_2;
            worksheet_1.Cell("D4").Value = (double)cef2_1 / cef2_2;
            worksheet_1.Cell("E4").Value = (double)cef3_1 / cef3_2;
            worksheet_1.Cell("F4").Value = (double)cef4_1 / cef4_2;
            worksheet_1.Cell("G4").Value = (double)cef5_1 / cef5_2;
            worksheet_1.Cell("H4").Value = (double)cef6_1 / cef6_2;

            workbook.SaveAs(Server.MapPath(@"~/images/Report.xlsx"));

            StreamWriter sw = new StreamWriter(Server.MapPath(@"~/images/DateUsd.txt"));
            sw.WriteLine("Скачать отчет сформированный " + DateTime.Now.ToString("dd.MM.yyyy"));
            sw.WriteLine("Последний отчет сформировал пользователь " + User.Identity.Name + " " + DateTime.Now.ToString("dd.MM.yyyy H:mm:ss") + " период с " + dateS.ToString("dd.MM.yyyy") + " по " + dateF.ToString("dd.MM.yyyy"));
            sw.Close();

            return Json("Отчет сформирован");
        }
        #endregion

        #region Сираница свыше 1000 БВ
        public ActionResult Index()
        {
            ViewBag.WinStatusList = GetAllWinStatus().ToList();
            ViewBag.LegislationList = GetAllLegislation().ToList();
            ViewBag.SubjectPurchaseList = GetAllSubjectPurchase().ToList();
            ViewBag.SubDivisionList = GetAllSubDivision().ToList();
            ViewBag.ResultList = GetAllResult().ToList();
            ViewBag.CurrencyList = GetAllCurrency().ToList();
            ViewBag.CountryList = GetAllCountry().ToList();


            int days = (int)DateTime.Now.DayOfWeek;
            DateTime lastMonth = DateTime.Now.AddMonths(-1);
            ViewBag.weekStart = DateTime.Now.AddDays(-days);
            ViewBag.weekEnd = ViewBag.weekStart.AddDays(6);
            ViewBag.monthStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            ViewBag.monthEnd = ViewBag.monthStart.AddMonths(1).AddDays(-1);
            ViewBag.YearStart = new DateTime(DateTime.Now.Year, 1, 1);
            ViewBag.YearEnd = new DateTime(DateTime.Now.Year, 12, 31);
            ViewBag.lastMonthStart = new DateTime(lastMonth.Year, lastMonth.Month, 1);
            ViewBag.lastMonthEnd = ViewBag.lastMonthStart.AddMonths(1).AddDays(-1);
            ViewBag.lastYearStart = new DateTime(DateTime.Now.Year - 1, 1, 1);
            ViewBag.lastYearEnd = new DateTime(DateTime.Now.Year - 1, 12, 31);

            List<string> Count = new List<string>();
            ProcurementInfoEntities context = new ProcurementInfoEntities();
            var table_count = context.S_Country;
            foreach(var c in table_count)
            {
                Count.Add(c.NAME);
            }
            Count.Sort();
            ViewBag.Country = Count;

            List<string> Count_leg = new List<string>();
            var table_leg = context.S_Legislation;
            foreach (var c in table_leg)
            {
                Count_leg.Add(c.NAME);
            }
            Count_leg.Sort();
            ViewBag.Legislation = Count_leg;

            List<string> Count_sub = new List<string>();
            var table_sub = context.S_SubjectPurchase;
            foreach(var c in table_sub)
            {
                Count_sub.Add(c.NAME);
            }
            Count_sub.Sort();
            ViewBag.Subject = Count_sub;

            List<string> Count_res = new List<string>();
            var table_res = context.S_Result;
            foreach (var c in table_res)
            {
                Count_res.Add(c.NAME);
            }
            Count_res.Sort();
            ViewBag.Result = Count_res;

            List<string> Count_win = new List<string>();
            var table_win = context.S_WinStatus;
            foreach (var c in table_win)
            {
                Count_win.Add(c.NAME);
            }
            Count_win.Sort();
            ViewBag.WinStatus = Count_win;

            List<string> Count_cur = new List<string>();
            var table_cur = context.S_Currency;
            foreach (var c in table_cur)
            {
                Count_cur.Add(c.LETTER_CODE);
            }
            Count_cur.Sort();
            ViewBag.Currency = Count_cur;

            List<string> SubDivision = new List<string>();
            var table_div = context.S_SubDivision;
            foreach (var c in table_div)
            {
                SubDivision.Add(c.NAME);
            }
            ViewBag.Division = SubDivision;

            return View();
        }
        #endregion

        #region Страницы до 1000 БВ
        public ActionResult Index_do1k()
        {
            ViewBag.LegislDo1000List = GetAllLegislDo1000().ToList();
            ViewBag.SubjectPurchaseList = GetAllSubjectPurchase().ToList();
            ViewBag.SubDivisionList = GetAllSubDivision().ToList();
            ViewBag.WinStatusDo1000List = GetAllWinStatusDo1000().ToList();
            ViewBag.CurrencyList = GetAllCurrency().ToList();
            ViewBag.CountryList = GetAllCountry().ToList();

            int days = (int)DateTime.Now.DayOfWeek;
            DateTime lastMonth = DateTime.Now.AddMonths(-1);
            ViewBag.weekStart = DateTime.Now.AddDays(-days);
            ViewBag.weekEnd = ViewBag.weekStart.AddDays(6);
            ViewBag.monthStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            ViewBag.monthEnd = ViewBag.monthStart.AddMonths(1).AddDays(-1);
            ViewBag.YearStart = new DateTime(DateTime.Now.Year, 1, 1);
            ViewBag.YearEnd = new DateTime(DateTime.Now.Year, 12, 31);
            ViewBag.lastMonthStart = new DateTime(lastMonth.Year, lastMonth.Month, 1);
            ViewBag.lastMonthEnd = ViewBag.lastMonthStart.AddMonths(1).AddDays(-1);
            ViewBag.lastYearStart = new DateTime(DateTime.Now.Year - 1, 1, 1);
            ViewBag.lastYearEnd = new DateTime(DateTime.Now.Year - 1, 12, 31);

            List<string> Count = new List<string>();
            ProcurementInfoEntities context = new ProcurementInfoEntities();
            var table = context.S_Country;
            foreach (var c in table)
            {
                Count.Add(c.NAME);
            }
            Count.Sort();
            ViewBag.Country = Count;

            List<string> Count_sub = new List<string>();
            var table_sub = context.S_SubjectPurchase;
            foreach (var c in table_sub)
            {
                Count_sub.Add(c.NAME);
            }
            Count_sub.Sort();
            ViewBag.Subject = Count_sub;

            List<string> Count_win = new List<string>();
            var table_win = context.S_WinStatusDo1000;
            foreach (var c in table_win)
            {
                Count_win.Add(c.NAME);
            }
            Count_win.Sort();
            ViewBag.WinStatus = Count_win;

            List<string> Count_leg = new List<string>();
            var table_leg = context.S_LegislDo1000;
            foreach (var c in table_leg)
            {
                Count_leg.Add(c.NAME);
            }
            Count_leg.Sort();
            ViewBag.Legislation = Count_leg;

            List<string> Count_cur = new List<string>();
            var table_cur = context.S_Currency;
            foreach (var c in table_cur)
            {
                Count_cur.Add(c.LETTER_CODE);
            }
            Count_cur.Sort();
            ViewBag.Currency = Count_cur;

            List<string> SubDivision = new List<string>();
            var table_div = context.S_SubDivision;
            foreach (var c in table_div)
            {
                SubDivision.Add(c.NAME);
            }
            ViewBag.Division = SubDivision;

            return View();
        }
        #endregion

        #region Заполнение DataGrid свыше 1000 БВ
        public ActionResult Datasource(
            DataManagerRequest dataManager, 
            DateTime? dateFrom = null,
            DateTime? dateTo = null)
        {
            IEnumerable DataSource = GetAllProcurementInformation(dateFrom, dateTo);
            DataOperations operation = new DataOperations();

            if (dataManager.Search != null && dataManager.Search.Count > 0)
                DataSource = operation.PerformSearching(DataSource, dataManager.Search);

            if (dataManager.Sorted != null && dataManager.Sorted.Count > 0)
                DataSource = operation.PerformSorting(DataSource, dataManager.Sorted);

            if (dataManager.Where != null && dataManager.Where.Count > 0)
                DataSource = operation.PerformFiltering(DataSource, dataManager.Where, dataManager.Where[0].Operator);

            int count = DataSource.Cast<T_ProcurementInformation>().Count();

            if (dataManager.Skip != 0)
                DataSource = operation.PerformSkip(DataSource, dataManager.Skip);

            if (dataManager.Take != 0)
                DataSource = operation.PerformTake(DataSource, dataManager.Take);

            return dataManager.RequiresCounts ? Json(new { result = DataSource, count }, JsonRequestBehavior.AllowGet) : Json(DataSource, JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region Редактирование/удаление/добавление записи в DataGrid свыше 1000 БВ
        public ActionResult CrudFunc(CRUDModel<T_ProcurementInformation> Object, string action)
        {
            if (Object.Action == "update")
            {
                var objectToUpdate = Object.Value;
                var oldObject = GetAllProcurementInformation().Where(f => f.ID == objectToUpdate.ID).FirstOrDefault();

                oldObject.NAME_PR = objectToUpdate.NAME_PR;
                oldObject.FIO_ISP = objectToUpdate.FIO_ISP;
                oldObject.ID_LEGISLATION = objectToUpdate.ID_LEGISLATION;
                oldObject.ID_SUBJECTPURCHASE = objectToUpdate.ID_SUBJECTPURCHASE;
                oldObject.ID_SUBDIVISION = objectToUpdate.ID_SUBDIVISION;
                oldObject.PR_APPROVAL_DATE = objectToUpdate.PR_APPROVAL_DATE;
                oldObject.PARTICIPANT__VALUE = objectToUpdate.PARTICIPANT__VALUE;
                oldObject.ID_RESULT = objectToUpdate.ID_RESULT;
                oldObject.USER_ENTER = Convert.ToString(User.Identity.Name) + " " + DateTime.Now.ToString("dd.MM.yyyy H:mm:ss") + " " + "Update";
                if (objectToUpdate.ID_LEGISLATION.Value == 3)
                {
                    oldObject.CONTRACT_DATE_TERM = null;
                    oldObject.RESULT_DATE_TERM = null;
                    objectToUpdate.CONTRACT_DATE_TERM = null;
                    objectToUpdate.RESULT_DATE_TERM = null;
                }
                else
                {
                    oldObject.CONTRACT_DATE_TERM = objectToUpdate.PR_APPROVAL_DATE.Value.AddDays(28);
                    oldObject.RESULT_DATE_TERM = objectToUpdate.PR_APPROVAL_DATE.Value.AddDays(31);
                    objectToUpdate.CONTRACT_DATE_TERM = oldObject.CONTRACT_DATE_TERM.Value;
                    objectToUpdate.RESULT_DATE_TERM = oldObject.RESULT_DATE_TERM.Value;
                }
                if(objectToUpdate.ID_RESULT != null)
                {
                    if (objectToUpdate.ID_RESULT.Value == 4)
                    {
                        oldObject.CONTRACT_DATE_TERM = null;
                        oldObject.RESULT_DATE_TERM = null;
                        objectToUpdate.CONTRACT_DATE_TERM = null;
                        objectToUpdate.RESULT_DATE_TERM = null;
                    }
                }
                
                oldObject.CONTRACT_DATE_FACT = objectToUpdate.CONTRACT_DATE_FACT;
                oldObject.RESULT_DATE_FACT = objectToUpdate.RESULT_DATE_FACT;
                oldObject.WIN_NAME = objectToUpdate.WIN_NAME;
                oldObject.WIN_VALUE = objectToUpdate.WIN_VALUE;
                oldObject.ID_WINSTATUS = objectToUpdate.ID_WINSTATUS;
                
                oldObject.DATE_ACTIVE = objectToUpdate.DATE_ACTIVE;
                oldObject.DATE_DEACT = objectToUpdate.DATE_DEACT;
                oldObject.FLAG = objectToUpdate.FLAG;
                oldObject.INVITE_NUN = objectToUpdate.INVITE_NUN;
                oldObject.ID_CURRENCY = objectToUpdate.ID_CURRENCY;
                oldObject.CONTRACT_DATE_PROL = objectToUpdate.CONTRACT_DATE_PROL;
                oldObject.CONTRACT_NUMBER = objectToUpdate.CONTRACT_NUMBER;
                oldObject.ID_COUNTRY = objectToUpdate.ID_COUNTRY;
                oldObject.ID_COUNTRY_ORIGIN = objectToUpdate.ID_COUNTRY_ORIGIN;
                oldObject.PURCHASE_VOLUME = objectToUpdate.PURCHASE_VOLUME;
                oldObject.VOLUME_UNITS = objectToUpdate.VOLUME_UNITS;
                oldObject.PRICE_PER_ITEM = objectToUpdate.PRICE_PER_ITEM;
                oldObject.WIN_VALUE_BYN = objectToUpdate.WIN_VALUE_BYN;
                oldObject.DELIVERY_COND = objectToUpdate.DELIVERY_COND;
                oldObject.MAIN_TEC_SPECS = objectToUpdate.MAIN_TEC_SPECS;
                oldObject.MANUFACTURER = objectToUpdate.MANUFACTURER;

                _databaseRepository.SaveChanges();
            }
            else if (Object.Action == "insert")
            {
                var objectToAdd = Object.Value;
                objectToAdd.DATE_ACTIVE = DateTime.Now;
                objectToAdd.FLAG = true;

                if (objectToAdd.PR_APPROVAL_DATE.Value != null)
                {
                    if ((objectToAdd.ID_LEGISLATION.Value == 3) || (objectToAdd.ID_LEGISLATION.Value == 4) || (objectToAdd.ID_LEGISLATION.Value == 6))
                    {
                        objectToAdd.CONTRACT_DATE_TERM = null;
                        objectToAdd.RESULT_DATE_TERM = null;
                    }
                    else
                    {
                        objectToAdd.CONTRACT_DATE_TERM = objectToAdd.PR_APPROVAL_DATE.Value.AddDays(28);
                        objectToAdd.RESULT_DATE_TERM = objectToAdd.CONTRACT_DATE_TERM.Value.AddDays(5);
                    }
                }

                /*if (objectToAdd.RESULT_DATE_FACT.Value != null && objectToAdd.ID_RESULT.Value == 2)  
                {
                    objectToAdd.KRO_DATE_TERM = objectToAdd.RESULT_DATE_FACT.Value.AddDays(1);
                }*/
                if (objectToAdd.ID_RESULT != null)
                {
                    if (objectToAdd.ID_RESULT.Value == 4)
                    {
                        objectToAdd.CONTRACT_DATE_TERM = null;
                        objectToAdd.RESULT_DATE_TERM = null;
                    }
                }
                objectToAdd.USER_ENTER = Convert.ToString(User.Identity.Name) + " " + DateTime.Now.ToString("dd.MM.yyyy H:mm:ss") + " " + "Insert";

                _databaseRepository.T_ProcurementInformation.Add(objectToAdd);
                _databaseRepository.SaveChanges();
            }
            else if (Object.Action == "remove")
            {
                var ID = Int32.Parse(Object.Key.ToString());
                var objectToRemove = _databaseRepository.T_ProcurementInformation.FirstOrDefault(f => f.ID == ID);
                if (objectToRemove != null)
                {
                    objectToRemove.DATE_DEACT = DateTime.Now;
                    objectToRemove.FLAG = false;
                    objectToRemove.USER_ENTER = Convert.ToString(User.Identity.Name) + DateTime.Now.ToString("dd.MM.yyyy H:mm:ss") + " " + "Remove";
                    _databaseRepository.SaveChanges();
                }
                return Json(Object, JsonRequestBehavior.AllowGet);
            }
            return Json(Object.Value, JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region GetAllProcurementInformation
        private IQueryable<T_ProcurementInformation> GetAllProcurementInformation(
            DateTime? dateFrom = null,
            DateTime? dateTo = null)
        {
            if ( dateFrom == null && dateTo == null)
                return _databaseRepository.T_ProcurementInformation
                                          .Where(f => f.FLAG == true)
                                          .OrderByDescending(d => d.DATE_ACTIVE);
            else
                return _databaseRepository.T_ProcurementInformation
                                          .Where(f => f.FLAG == true && f.PR_APPROVAL_DATE >= dateFrom && f.PR_APPROVAL_DATE <= dateTo)
                                          .OrderByDescending(d => d.DATE_ACTIVE);
        }
        #endregion

        #region Заполнение DataGrid до 1000 БВ
        public ActionResult DatasourceDo1000(
       DataManagerRequest dataManager,
            DateTime? dateFrom = null,
            DateTime? dateTo = null)
        {

            IEnumerable DataSource = GetAllPrInfoDo1000(dateFrom, dateTo);
            DataOperations operation = new DataOperations();

            if (dataManager.Search != null && dataManager.Search.Count > 0)
                DataSource = operation.PerformSearching(DataSource, dataManager.Search);

            if (dataManager.Sorted != null && dataManager.Sorted.Count > 0)
                DataSource = operation.PerformSorting(DataSource, dataManager.Sorted);

            if (dataManager.Where != null && dataManager.Where.Count > 0)
                DataSource = operation.PerformFiltering(DataSource, dataManager.Where, dataManager.Where[0].Operator);

            double count = DataSource.Cast<T_PrInfoDo1000>().Count();

            if (dataManager.Skip != 0)
                DataSource = operation.PerformSkip(DataSource, dataManager.Skip);

            if (dataManager.Take != 0)
                DataSource = operation.PerformTake(DataSource, dataManager.Take);

            return dataManager.RequiresCounts ? Json(new { result = DataSource, count }, JsonRequestBehavior.AllowGet) : Json(DataSource, JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region Редактирование/удаление/добавлени записи в DataGrid до 1000 БВ
        public ActionResult CrudFuncDo1000(CRUDModel<T_PrInfoDo1000> Object, string action)
        {
            if (Object.Action == "update")
            {
                var objectToUpdate = Object.Value;
                var oldObject = GetAllPrInfoDo1000().Where(f => f.ID == objectToUpdate.ID).FirstOrDefault();

                oldObject.NAME_PR = objectToUpdate.NAME_PR;
                oldObject.ID_LEGISLDO1000 = objectToUpdate.ID_LEGISLDO1000;
                oldObject.CONTRACT_NUMBER = objectToUpdate.CONTRACT_NUMBER;
                oldObject.DATE_CONCLUSION = objectToUpdate.DATE_CONCLUSION;
                oldObject.WIN_NAME = objectToUpdate.WIN_NAME;
                oldObject.WIN_VALUE = objectToUpdate.WIN_VALUE;
                oldObject.WIN_VALUE_NDE = objectToUpdate.WIN_VALUE_NDE;
                oldObject.WIN_VALUE_NNDS = objectToUpdate.WIN_VALUE_NNDS;
                oldObject.FIO_ISP = objectToUpdate.FIO_ISP;
                oldObject.ID_CURRENCY = objectToUpdate.ID_CURRENCY;
                oldObject.DATE_ACTIVE = objectToUpdate.DATE_ACTIVE;
                oldObject.DATE_DEACT = objectToUpdate.DATE_DEACT;
                oldObject.FLAG = objectToUpdate.FLAG;
                oldObject.ID_WINSTATUSDO1000 = objectToUpdate.ID_WINSTATUSDO1000;
                oldObject.ID_COUNTRY_ORIGIN = objectToUpdate.ID_COUNTRY_ORIGIN;
                oldObject.ID_SUBJECTPURCHASE = objectToUpdate.ID_SUBJECTPURCHASE;
                oldObject.ID_SUBDIVISION = objectToUpdate.ID_SUBDIVISION;
                oldObject.USER_ENTER = Convert.ToString(User.Identity.Name) + " " + DateTime.Now.ToString("dd.MM.yyyy H:mm:ss") + " " + "Update";

                _databaseRepository.SaveChanges();
            }
            else if (Object.Action == "insert")
            {
                var objectToAdd = Object.Value;
                objectToAdd.DATE_ACTIVE = DateTime.Now;
                objectToAdd.FLAG = true;
                objectToAdd.USER_ENTER = Convert.ToString(User.Identity.Name) + " " + DateTime.Now.ToString("dd.MM.yyyy H:mm:ss") + " " + "Insert";

                _databaseRepository.T_PrInfoDo1000.Add(objectToAdd);
                _databaseRepository.SaveChanges();
            }
            else if (Object.Action == "remove")
            {
                var ID = Int32.Parse(Object.Key.ToString());
                var objectToRemove = _databaseRepository.T_PrInfoDo1000.FirstOrDefault(f => f.ID == ID);
                if (objectToRemove != null)
                {
                    objectToRemove.DATE_DEACT = DateTime.Now;
                    objectToRemove.FLAG = false;
                    objectToRemove.USER_ENTER = Convert.ToString(User.Identity.Name) + " " + DateTime.Now.ToString("dd.MM.yyyy H:mm:ss") + " " + "Remove";
                    _databaseRepository.SaveChanges();
                }
                return Json(Object, JsonRequestBehavior.AllowGet);
            }
            return Json(Object.Value, JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region GetAllPrInfoDo1000
        private IQueryable<T_PrInfoDo1000> GetAllPrInfoDo1000(
            DateTime? dateFrom = null,
            DateTime? dateTo = null)
        {
            if (dateFrom == null && dateTo == null)
                return _databaseRepository.T_PrInfoDo1000
                                          .Where(f => f.FLAG == true)
                                          .OrderByDescending(d => d.DATE_ACTIVE);
            else
                return _databaseRepository.T_PrInfoDo1000
                                          .Where(f => f.FLAG == true && f.DATE_CONCLUSION >= dateFrom && f.DATE_CONCLUSION <= dateTo)
                                          .OrderByDescending(d => d.DATE_ACTIVE);
        }
        #endregion

        #region Заполнение строк формы редактирование записи в отдельном окне до 1000 БВ
        public ActionResult EnterLenghtDo1k(CRUDModel<T_PrInfoDo1000> Object, int ID_lenght)
        {
            var oldObject = GetAllPrInfoDo1000().Where(f => f.ID == ID_lenght && f.FLAG == true).FirstOrDefault();
            if (oldObject != null)
            {
                List<string> id_lenght = new List<string>();
                id_lenght.Add(oldObject.NAME_PR);
                var sub = GetAllSubjectPurchase().Where(c => c.ID == oldObject.ID_SUBJECTPURCHASE).FirstOrDefault();
                if (sub != null)
                {
                    id_lenght.Add(sub.NAME);
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.CONTRACT_NUMBER != null)
                {
                    id_lenght.Add(oldObject.CONTRACT_NUMBER);
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.FIO_ISP != null)
                {
                    id_lenght.Add(oldObject.FIO_ISP);
                }
                else
                {
                    id_lenght.Add("");
                }

                var win = GetAllWinStatusDo1000().Where(c => c.ID == oldObject.ID_WINSTATUSDO1000).FirstOrDefault();
                if (win != null)
                {
                    id_lenght.Add(win.NAME);
                }
                else
                {
                    id_lenght.Add("");
                }

                var leg = GetAllLegislDo1000().Where(c => c.ID == oldObject.ID_LEGISLDO1000).FirstOrDefault();
                if (leg != null)
                {
                    id_lenght.Add(leg.NAME);
                }
                else
                {
                    id_lenght.Add("");
                }

                if(oldObject.DATE_CONCLUSION != null)
                {
                    id_lenght.Add(Convert.ToDateTime(oldObject.DATE_CONCLUSION.Value.AddDays(1)).ToString("yyyy/MM/dd"));
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.WIN_NAME != null)
                {
                    id_lenght.Add(Convert.ToString(oldObject.WIN_VALUE));
                }
                else
                {
                    id_lenght.Add("");
                }
                
                var cur = GetAllCurrency().Where(c => c.ID == oldObject.ID_CURRENCY).FirstOrDefault();
                if (cur != null)
                {
                    id_lenght.Add(cur.LETTER_CODE);
                }
                else
                {
                    id_lenght.Add("");
                }
                id_lenght.Add(Convert.ToString(oldObject.WIN_VALUE_NDE));
                id_lenght.Add(Convert.ToString(oldObject.WIN_VALUE_NNDS));
                if (oldObject.WIN_NAME != null)
                {
                    id_lenght.Add(oldObject.WIN_NAME);
                }
                else
                {
                    id_lenght.Add("");
                }
                var coun = GetAllCountry().Where(c => c.ID == oldObject.ID_COUNTRY_ORIGIN).FirstOrDefault();
                if (coun != null)
                {
                    id_lenght.Add(coun.NAME);
                }
                else
                {
                    id_lenght.Add("");
                }

                return Json(id_lenght);
            }
            else
            {
                return Json(1);
            }
        }
        #endregion

        #region Редактирование записи в окне до 1000 БВ
        public ActionResult Update_lineDo1k(CRUDModel<T_PrInfoDo1000> Object, string line1, string line2, string line3, string line4, string line5, string line6, string line7, string line8, string line9, string line10, string line11, string line12, string line13, string line14, int id)
        {
            var oldObject = GetAllPrInfoDo1000().Where(f => f.ID == id && f.FLAG == true).FirstOrDefault();


            if (oldObject != null)
            {
                if (line1 != "")
                {
                    oldObject.NAME_PR = line1;
                }
                else
                {
                    oldObject.NAME_PR = null;
                }

                if (line2 != "")
                {
                    var sub = GetAllSubjectPurchase().Where(c => c.NAME == line2).FirstOrDefault();
                    if (sub != null)
                    {
                        oldObject.ID_SUBJECTPURCHASE = sub.ID;
                    }
                    else
                    {
                        oldObject.ID_SUBJECTPURCHASE = null;
                    }
                }
                else
                {
                    oldObject.ID_SUBJECTPURCHASE = null;
                }

                if (line6 != "")
                {
                    var leg = GetAllLegislDo1000().Where(c => c.NAME == line6).FirstOrDefault();
                    if (leg != null)
                    {
                        oldObject.ID_LEGISLDO1000 = leg.ID;
                    }
                    else
                    {
                        oldObject.ID_LEGISLDO1000 = null;
                    }
                }
                else
                {
                    oldObject.ID_LEGISLDO1000 = null;
                }

                if (line3 != "")
                {
                    oldObject.CONTRACT_NUMBER = line3;
                }
                else
                {
                    oldObject.CONTRACT_NUMBER = null;
                }

                if (line7 != "")
                {
                    oldObject.DATE_CONCLUSION = Convert.ToDateTime(line7);
                }
                else
                {
                    oldObject.DATE_CONCLUSION = null;
                }

                if (line5 != "")
                {
                    var win = GetAllWinStatusDo1000().Where(c => c.NAME == line5).FirstOrDefault();
                    if (win != null)
                    {
                        oldObject.ID_WINSTATUSDO1000 = win.ID;
                    }
                    else
                    {
                        oldObject.ID_WINSTATUSDO1000 = null;
                    }
                }
                else
                {
                    oldObject.ID_WINSTATUSDO1000 = null;
                }

                if (line12 != "")
                {
                    oldObject.WIN_NAME = line12;
                }
                else
                {
                    oldObject.WIN_NAME = null;
                }

                if (line8 != "")
                {
                    oldObject.WIN_VALUE = Convert.ToDecimal(line8.Replace(",", "."));
                }
                else
                {
                    oldObject.WIN_VALUE = null;
                }

                if (line10 != "")
                {
                    oldObject.WIN_VALUE_NDE = Math.Round(Convert.ToDecimal(line10.Replace(",", ".")), 2);
                }
                else
                {
                    oldObject.WIN_VALUE_NDE = null;
                }

                if (line11 != "")
                {
                    oldObject.WIN_VALUE_NNDS = Convert.ToDecimal(line11.Replace(",", "."));
                }
                else
                {
                    oldObject.WIN_VALUE_NNDS = null;
                }

                if (line4 != "")
                {
                    oldObject.FIO_ISP = line4;
                }
                else
                {
                    oldObject.FIO_ISP = null;
                }

                if (line9 != "")
                {
                    var cur = GetAllCurrency().Where(c => c.LETTER_CODE == line9).FirstOrDefault();
                    if (cur != null)
                    {
                        oldObject.ID_CURRENCY = cur.ID;
                    }
                    else
                    {
                        oldObject.ID_CURRENCY = null;
                    }
                }
                else
                {
                    oldObject.ID_CURRENCY = null;
                }

                if (line13 != "")
                {
                    var caun = GetAllCountry().Where(c => c.NAME == line13).FirstOrDefault();
                    if (caun != null)
                    {
                        oldObject.ID_COUNTRY_ORIGIN = caun.ID;
                    }
                    else
                    {
                        oldObject.ID_COUNTRY_ORIGIN = null;
                    }
                }
                else
                {
                    oldObject.ID_COUNTRY_ORIGIN = null;
                }

                if (line14 != "")
                {
                    ProcurementInfoEntities context = new ProcurementInfoEntities();
                    var table = context.S_SubDivision.Where(c => c.NAME == line14);
                    if (table != null)
                    {
                        foreach (var c in table)
                        {
                            oldObject.ID_SUBDIVISION = c.ID;
                        }
                    }
                }
                else
                {
                    oldObject.ID_SUBDIVISION = null;
                }

                oldObject.USER_ENTER = Convert.ToString(User.Identity.Name) + " " + DateTime.Now.ToString("dd.MM.yyyy H:mm:ss") + " " + "Edit";

                _databaseRepository.SaveChanges();

                return Json("Запись успешна обновлена!");
            }
            else
            {
                return Json("Что-то пошло не так");
            }
        }
        #endregion

        #region Редактирование записи в окне свыше 1000 БВ
        public ActionResult UpdateLine(CRUDModel<T_ProcurementInformation> Object, int ID_line, string ln1, string ln2, string ln3, string ln4, string ln5, string ln6, string ln7, string ln8, string ln9, string ln10, string ln11, string ln12, string ln13, string ln14, string ln15, string ln16, string ln17, string ln18, string ln19, string ln20, string ln21, string ln22, string ln23, string ln24, string ln25, string ln26, string ln27, string ln28)
        {
            var oldObject = GetAllProcurementInformation().Where(f => f.ID == ID_line && f.FLAG == true).FirstOrDefault();
            if (oldObject != null)
            {
                if (ln1 != "")
                {
                    oldObject.NAME_PR = Convert.ToString(ln1);
                }
                else
                {
                    oldObject.NAME_PR = null;
                }

                if (ln2 != "")
                {
                    var sub = GetAllSubjectPurchase().Where(c => c.NAME == ln2).FirstOrDefault();
                    if (sub != null)
                    {
                        oldObject.ID_SUBJECTPURCHASE = sub.ID;
                    }
                    else
                    {
                        oldObject.ID_SUBJECTPURCHASE = null;
                    }
                }
                else
                {
                    oldObject.ID_SUBJECTPURCHASE = null;
                }

                if (ln3 != "")
                {
                    oldObject.INVITE_NUN = ln3;
                }
                else
                {
                    oldObject.INVITE_NUN = null;
                }

                if (ln4 != "")
                {
                    oldObject.FIO_ISP = ln4;
                }
                else
                {
                    oldObject.FIO_ISP = null;
                }

                if (ln5 != "")
                {
                    var leg = GetAllLegislation().Where(c => c.NAME == ln5).FirstOrDefault();
                    if (leg != null)
                    {
                        oldObject.ID_LEGISLATION = leg.ID;
                    }
                    else
                    {
                        oldObject.ID_LEGISLATION = null;
                    }
                }
                else
                {
                    oldObject.ID_LEGISLATION = null;
                }

                if (ln6 != "")
                {
                    oldObject.PARTICIPANT__VALUE = Convert.ToInt32(ln6);
                }
                else
                {
                    oldObject.PARTICIPANT__VALUE = null;
                }

                if (ln7 != "")
                {
                    oldObject.PR_APPROVAL_DATE = Convert.ToDateTime(ln7);
                }
                else
                {
                    oldObject.PR_APPROVAL_DATE = DateTime.Now;
                }

                if (ln8 != "")
                {
                    var res = GetAllResult().Where(c => c.NAME == ln8).FirstOrDefault();
                    if (res != null)
                    {
                        oldObject.ID_RESULT = res.ID;
                    }
                    else
                    {
                        oldObject.ID_RESULT = null;
                    }
                }
                else
                {
                    oldObject.ID_RESULT = null;
                }

                if(oldObject.ID_LEGISLATION == 3)
                {
                    oldObject.CONTRACT_DATE_TERM = null;
                    oldObject.RESULT_DATE_TERM = null;
                }
                else
                {
                    oldObject.CONTRACT_DATE_TERM = Convert.ToDateTime(ln7).AddDays(28);
                    oldObject.RESULT_DATE_TERM = Convert.ToDateTime(ln7).AddDays(31);
                }

                if(oldObject.ID_RESULT == 4)
                {
                    oldObject.CONTRACT_DATE_TERM = null;
                    oldObject.RESULT_DATE_TERM = null;
                }
                

                if (ln10 != "")
                {
                    oldObject.CONTRACT_DATE_PROL = Convert.ToDateTime(ln10);
                }
                else
                {
                    oldObject.CONTRACT_DATE_PROL = null;
                }

                if (ln11 != "")
                {
                    oldObject.CONTRACT_DATE_FACT = Convert.ToDateTime(ln11);
                }
                else
                {
                    oldObject.CONTRACT_DATE_FACT = null;
                }

                if (ln12 != "")
                {
                    oldObject.CONTRACT_NUMBER = ln12;
                }
                else
                {
                    oldObject.CONTRACT_NUMBER = null;
                }

                if (ln15 != "")
                {
                    oldObject.WIN_NAME = ln15;
                }
                else
                {
                    oldObject.WIN_NAME = null;
                }

                ln16 = ln16.ToUpper();
                if (ln16 != "")
                {
                    var coun = GetAllCountry().Where(c => c.NAME == ln16).FirstOrDefault();
                    if (coun != null)
                    {
                        oldObject.ID_COUNTRY = coun.ID;
                    }
                    else
                    {
                        oldObject.ID_COUNTRY = null;
                    }
                }
                else
                {
                    oldObject.ID_COUNTRY = null;
                }

                if (ln17 != "")
                {
                    oldObject.WIN_VALUE = Convert.ToDecimal(ln17.Replace(",", "."));
                }
                else
                {
                    oldObject.WIN_VALUE = null;
                }

                if (ln18 != "")
                {
                    var cur = GetAllCurrency().Where(c => c.LETTER_CODE == ln18).FirstOrDefault();
                    if (cur != null)
                    {
                        oldObject.ID_CURRENCY = cur.ID;
                    }
                    else
                    {
                        oldObject.ID_CURRENCY = null;
                    }
                }
                else
                {
                    oldObject.ID_CURRENCY = null;
                }

                if (ln19 != "")
                {
                    oldObject.WIN_VALUE_BYN = float.Parse(ln19.Replace(",", "."));
                }
                else
                {
                    oldObject.WIN_VALUE_BYN = null;
                }

                if (ln20 != "")
                {
                    var win = GetAllWinStatus().Where(c => c.NAME == ln20).FirstOrDefault();
                    if (win != null)
                    {
                        oldObject.ID_WINSTATUS = win.ID;
                    }
                    else
                    {
                        oldObject.ID_WINSTATUS = null;
                    }
                }
                else
                {
                    oldObject.ID_WINSTATUS = null;
                }

                if (ln21 != "")
                {
                    oldObject.PURCHASE_VOLUME = ln21;
                }
                else
                {
                    oldObject.PURCHASE_VOLUME = null;
                }

                if (ln22 != "")
                {
                    oldObject.VOLUME_UNITS = ln22;
                }
                else
                {
                    oldObject.VOLUME_UNITS = null;
                }

                if (ln23 != "")
                {
                    oldObject.PRICE_PER_ITEM = ln23.Replace(",", ".");
                }
                else
                {
                    oldObject.PRICE_PER_ITEM = null;
                }

                if (ln24 != "")
                {
                    oldObject.DELIVERY_COND = ln24;
                }
                else
                {
                    oldObject.DELIVERY_COND = null;
                }

                if (ln25 != "")
                {
                    oldObject.MAIN_TEC_SPECS = ln25;
                }
                else
                {
                    oldObject.MAIN_TEC_SPECS = null;
                }

                ln26 = ln26.ToUpper();
                if (ln26 != "")
                {
                    var coun_o = GetAllCountry().Where(c => c.NAME == ln26).FirstOrDefault();
                    if (coun_o != null)
                    {
                        oldObject.ID_COUNTRY_ORIGIN = coun_o.ID;
                    }
                    else
                    {
                        oldObject.ID_COUNTRY_ORIGIN = null;
                    }
                }
                else
                {
                    oldObject.ID_COUNTRY_ORIGIN = null;
                }

                if (ln28 != "")
                {
                    var coun_o = GetAllSubDivision().Where(c => c.NAME == ln28).FirstOrDefault();
                    if (coun_o != null)
                    {
                        oldObject.ID_SUBDIVISION = coun_o.ID;
                    }
                    else
                    {
                        oldObject.ID_SUBDIVISION = null;
                    }
                }
                else
                {
                    oldObject.ID_SUBDIVISION = null;
                }

                if (ln27 != "")
                {
                    oldObject.MANUFACTURER = ln27;
                }
                else
                {
                    oldObject.MANUFACTURER = null;
                }
                oldObject.USER_ENTER = Convert.ToString(User.Identity.Name) + " " + DateTime.Now.ToString("dd.MM.yyyy H:mm:ss") + " " + "Edit";

                _databaseRepository.SaveChanges();

                return Json("Запись обновлена!");
            }
            else
            {
                return Json("Что-то пошло не так");
            }
        }
        #endregion

        #region Заполнение данных в редактирование записи в окне свыше 1000 БВ
        public ActionResult EnterLenght(CRUDModel<T_ProcurementInformation> Object, int id_line)
        {
            var oldObject = GetAllProcurementInformation().Where(f => f.ID == id_line && f.FLAG == true).FirstOrDefault();
            if(oldObject != null)
            {
                List<string> id_lenght = new List<string>();
                if (oldObject.NAME_PR != null)
                {
                    id_lenght.Add(oldObject.NAME_PR);
                }
                else
                {
                    id_lenght.Add("");
                }

                var sub = GetAllSubjectPurchase().Where(c => c.ID == oldObject.ID_SUBJECTPURCHASE).FirstOrDefault();
                if (sub != null)
                {
                    id_lenght.Add(sub.NAME);
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.INVITE_NUN != null)
                {
                    id_lenght.Add(Convert.ToString(oldObject.INVITE_NUN));
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.FIO_ISP != null)
                {
                    id_lenght.Add(Convert.ToString(oldObject.FIO_ISP));
                }
                else
                {
                    id_lenght.Add("");
                }

                var leg = GetAllLegislation().Where(c => c.ID == oldObject.ID_LEGISLATION).FirstOrDefault();
                if (leg != null)
                {
                    id_lenght.Add(leg.NAME);
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.PARTICIPANT__VALUE != null)
                {
                    id_lenght.Add(Convert.ToString(oldObject.PARTICIPANT__VALUE));
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.PR_APPROVAL_DATE != null)
                {
                    id_lenght.Add(Convert.ToDateTime(oldObject.PR_APPROVAL_DATE.Value.AddDays(1)).ToString("yyyy/MM/dd"));
                }
                else
                {
                    id_lenght.Add("");
                }

                var res = GetAllResult().Where(c => c.ID == oldObject.ID_RESULT).FirstOrDefault();
                if (res != null)
                {
                    id_lenght.Add(res.NAME);
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.CONTRACT_DATE_TERM != null)
                {
                    id_lenght.Add(Convert.ToDateTime(oldObject.CONTRACT_DATE_TERM.Value.AddDays(1)).ToString("yyyy/MM/dd"));
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.CONTRACT_DATE_PROL != null)
                {
                    id_lenght.Add(Convert.ToDateTime(oldObject.CONTRACT_DATE_PROL.Value.AddDays(1)).ToString("yyyy/MM/dd"));
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.CONTRACT_DATE_FACT != null)
                {
                    id_lenght.Add(Convert.ToDateTime(oldObject.CONTRACT_DATE_FACT.Value.AddDays(1)).ToString("yyyy/MM/dd"));
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.CONTRACT_NUMBER != null)
                {
                    id_lenght.Add(oldObject.CONTRACT_NUMBER);
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.RESULT_DATE_TERM != null)
                {
                    id_lenght.Add(Convert.ToDateTime(oldObject.RESULT_DATE_TERM.Value.AddDays(1)).ToString("yyyy/MM/dd"));
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.RESULT_DATE_FACT != null)
                {
                    id_lenght.Add(Convert.ToDateTime(oldObject.RESULT_DATE_FACT.Value.AddDays(1)).ToString("yyyy/MM/dd"));
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.WIN_NAME != null)
                {
                    id_lenght.Add(oldObject.WIN_NAME);
                }
                else
                {
                    id_lenght.Add("");
                }

                var coun = GetAllCountry().Where(c => c.ID == oldObject.ID_COUNTRY).FirstOrDefault();
                if (coun != null)
                {
                    id_lenght.Add(coun.NAME);
                }
                else
                {
                    id_lenght.Add("");
                }
                if (oldObject.WIN_VALUE != null)
                {
                    id_lenght.Add(Convert.ToString(oldObject.WIN_VALUE));
                }
                else
                {
                    id_lenght.Add("");
                }

                var cur = GetAllCurrency().Where(c => c.ID == oldObject.ID_CURRENCY).FirstOrDefault();
                if (cur != null)
                {
                    id_lenght.Add(cur.LETTER_CODE);
                }
                else
                {
                    id_lenght.Add("");
                }

                id_lenght.Add(Convert.ToString(oldObject.WIN_VALUE_BYN));

                var win = GetAllWinStatus().Where(c => c.ID == oldObject.ID_WINSTATUS).FirstOrDefault();
                if (win != null)
                {
                    id_lenght.Add(win.NAME);
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.PURCHASE_VOLUME != null)
                {
                    id_lenght.Add(Convert.ToString(oldObject.PURCHASE_VOLUME));
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.VOLUME_UNITS != null)
                {
                    id_lenght.Add(Convert.ToString(oldObject.VOLUME_UNITS));
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.PRICE_PER_ITEM != null)
                {
                    id_lenght.Add(Convert.ToString(oldObject.PRICE_PER_ITEM));
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.DELIVERY_COND != null)
                {
                    id_lenght.Add(Convert.ToString(oldObject.DELIVERY_COND));
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.MAIN_TEC_SPECS != null)
                {
                    id_lenght.Add(Convert.ToString(oldObject.MAIN_TEC_SPECS));
                }
                else
                {
                    id_lenght.Add("");
                }

                var coun_o = GetAllCountry().Where(c => c.ID == oldObject.ID_COUNTRY_ORIGIN).FirstOrDefault();
                if (coun_o != null)
                {
                    id_lenght.Add(coun_o.NAME);
                }
                else
                {
                    id_lenght.Add("");
                }

                if (oldObject.MANUFACTURER != null)
                {
                    id_lenght.Add(Convert.ToString(oldObject.MANUFACTURER));
                }
                else
                {
                    id_lenght.Add("");
                }

                var subdiv = GetAllSubDivision().Where(c => c.ID == oldObject.ID_SUBDIVISION).FirstOrDefault();
                if (subdiv != null)
                {
                    id_lenght.Add(subdiv.NAME);
                }
                else
                {
                    id_lenght.Add("");
                }

                return Json(id_lenght);
            }
            else
            {
                return Json(1);
            }
           
        }
        #endregion

        #region Проверка правильности заполнения поля страны (До 1000 БВ)
        public JsonResult Cauntry_CheckDo1k(string count)
        {
            ProcurementInfoEntities context = new ProcurementInfoEntities();
            var table = context.S_Country;
            List<string> country = new List<string>();
            foreach (var c in table)
            {
                country.Add(Convert.ToString(c.NAME));
            }
            if ((country.FirstOrDefault(x => x.Contains(count.ToUpper())) != null) || count == "")
            {
                return Json(1);
            }
            else
            {
                return Json(2);
            }
        }
        #endregion

        #region Проверка заполнения поля страны (Свыше 1000 БВ)
        public JsonResult Country_Creater(string count1, string count2)
        {
            ProcurementInfoEntities context = new ProcurementInfoEntities();
            var table = context.S_Country;
            List<string> country = new List<string>();
            foreach (var c in table)
            {
                country.Add(Convert.ToString(c.NAME));
            }
            if ((country.FirstOrDefault(x => x.Contains(count1.ToUpper())) != null & country.FirstOrDefault(x => x.Contains(count2.ToUpper())) != null) || (count1 == "" & country.FirstOrDefault(x => x.Contains(count2.ToUpper())) != null) || (country.FirstOrDefault(x => x.Contains(count1.ToUpper())) != null & count2 == "") || (count1 == "" & count2 == ""))
            {
                return Json(1);
            }
            else
            {
                return Json(2);
            }
        }
        #endregion

        #region Добавление записи в талицу в отдельном окне свыше 1000 БВ
        public ActionResult Ins(CRUDModel<T_ProcurementInformation> Object, string ad1, string ad2, string ad3, string ad4, string ad5, string ad6, string ad7, string ad8, string ad9, string ad10, string ad11, string ad12, string ad13, string ad14, string ad15, string ad16, string ad17, string ad18, string ad19, string ad20, string ad21, string ad22, string ad23, string ad24, string ad25, string ad26, string ad27, string ad28)
        {
            if (ad1 != "" && ad2 != "" && ad4 != "" && ad5 != "" && ad28 != "")
            {
                T_ProcurementInformation objectAdd = new T_ProcurementInformation();

                if (ad1 != "")
                {
                    objectAdd.NAME_PR = Convert.ToString(ad1);
                }
                else
                {
                    objectAdd.NAME_PR = null;
                }

                if (String.Compare(ad2, "Работы (услуги)") == 0)
                {
                    objectAdd.ID_SUBJECTPURCHASE = 2;
                }
                else if (String.Compare(ad2, "Товары") == 0)
                {
                    objectAdd.ID_SUBJECTPURCHASE = 2;
                }

                if (ad3 != "")
                {
                    objectAdd.INVITE_NUN = Convert.ToString(ad3);
                }
                else
                {
                    objectAdd.INVITE_NUN = null;
                }
                objectAdd.FIO_ISP = Convert.ToString(ad4);

                if (String.Compare(ad5, "Биржа") == 0)
                {
                    objectAdd.ID_LEGISLATION = 3;
                    objectAdd.CONTRACT_DATE_TERM = null;
                    objectAdd.RESULT_DATE_TERM = null;
                }
                else if (String.Compare(ad5, "Пост.229 - Закупка из 1 источника") == 0)
                {
                    objectAdd.ID_LEGISLATION = 6;
                    objectAdd.CONTRACT_DATE_TERM = Convert.ToDateTime(ad7).AddDays(28);
                    objectAdd.RESULT_DATE_TERM = Convert.ToDateTime(ad7).AddDays(31);
                }
                else if (String.Compare(ad5, "Пост.229 - Приложение 1") == 0)
                {
                    objectAdd.ID_LEGISLATION = 4;
                    objectAdd.CONTRACT_DATE_TERM = Convert.ToDateTime(ad7).AddDays(28);
                    objectAdd.RESULT_DATE_TERM = Convert.ToDateTime(ad7).AddDays(31);
                }
                else if (String.Compare(ad5, "Пост.229 - Конкурс") == 0)
                {
                    objectAdd.ID_LEGISLATION = 1;
                    objectAdd.CONTRACT_DATE_TERM = Convert.ToDateTime(ad7).AddDays(28);
                    objectAdd.RESULT_DATE_TERM = Convert.ToDateTime(ad7).AddDays(31);
                }
                else if (String.Compare(ad5, "Строительство") == 0)
                {
                    objectAdd.ID_LEGISLATION = 2;
                    objectAdd.CONTRACT_DATE_TERM = Convert.ToDateTime(ad7).AddDays(28);
                    objectAdd.RESULT_DATE_TERM = Convert.ToDateTime(ad7).AddDays(31);
                }

                if (ad6 != "")
                {
                    objectAdd.PARTICIPANT__VALUE = Convert.ToInt32(ad6);
                }
                else
                {
                    objectAdd.PARTICIPANT__VALUE = null;
                }

                if (ad7 != "")
                {
                    objectAdd.PR_APPROVAL_DATE = Convert.ToDateTime(ad7);
                }
                else
                {
                    objectAdd.PR_APPROVAL_DATE = DateTime.Now;
                }

                if (String.Compare(ad8, "Состоялась") == 0)
                {
                    objectAdd.ID_RESULT = 1;
                }
                else if (String.Compare(ad8, "Не состоялась") == 0)
                {
                    objectAdd.ID_RESULT = 2;
                }
                else if (String.Compare(ad8, "Отменена") == 0)
                {
                    objectAdd.ID_RESULT = 4;
                    objectAdd.CONTRACT_DATE_TERM = null;
                    objectAdd.RESULT_DATE_TERM = null;
                }
                else if (ad8 == "")
                {
                    objectAdd.ID_RESULT = null;
                }

                if (ad10 != "")
                {
                    objectAdd.CONTRACT_DATE_PROL = Convert.ToDateTime(ad10);
                }
                else
                {
                    objectAdd.CONTRACT_DATE_PROL = null;
                }

                if (ad11 != "")
                {
                    objectAdd.CONTRACT_DATE_FACT = Convert.ToDateTime(ad11);
                }
                else
                {
                    objectAdd.CONTRACT_DATE_FACT = null;
                }

                if (ad12 != "")
                {
                    objectAdd.CONTRACT_NUMBER = Convert.ToString(ad12);
                }
                else
                {
                    objectAdd.CONTRACT_NUMBER = null;
                }

                if (ad14 != "")
                {
                    objectAdd.RESULT_DATE_FACT = Convert.ToDateTime(ad14);
                }
                else
                {
                    objectAdd.RESULT_DATE_FACT = null;
                }
                if (ad15 != "")
                {
                    objectAdd.WIN_NAME = Convert.ToString(ad15);
                }
                else
                {
                    objectAdd.WIN_NAME = null;
                }

                if (ad16 != "")
                {
                    ad16 = ad16.ToUpper();
                    ProcurementInfoEntities context = new ProcurementInfoEntities();
                    var table = context.S_Country.Where(c => c.NAME == ad16);
                    if (table != null)
                    {
                        foreach (var c in table)
                        {
                            objectAdd.ID_COUNTRY = c.ID;
                        }
                    }
                    else
                    {
                        objectAdd.ID_COUNTRY = null;
                    }
                }
                else
                {
                    objectAdd.ID_COUNTRY = null;
                }

                ad17 = ad17.Replace(".", ",");
                if (ad17 != "")
                {
                    objectAdd.WIN_VALUE = Convert.ToDecimal(ad17);
                }
                else
                {
                    objectAdd.WIN_VALUE = null;
                }

                if (String.Compare(ad18, "BYN") == 0)
                {
                    objectAdd.ID_CURRENCY = 1;
                }
                else if (String.Compare(ad18, "USD") == 0)
                {
                    objectAdd.ID_CURRENCY = 2;
                }
                else if (String.Compare(ad18, "EUR") == 0)
                {
                    objectAdd.ID_CURRENCY = 3;
                }
                else if (String.Compare(ad18, "RUB") == 0)
                {
                    objectAdd.ID_CURRENCY = 4;
                }
                else
                {
                    objectAdd.ID_CURRENCY = null;
                }

                ad19 = ad19.Replace(".", ",");
                if (ad19 != "")
                {

                    objectAdd.WIN_VALUE_BYN = float.Parse(ad19);
                }
                else
                {
                    objectAdd.WIN_VALUE_BYN = null;
                }

                if (String.Compare(ad20, "Официальный представитель") == 0)
                {
                    objectAdd.ID_WINSTATUS = 3;
                }
                else if (String.Compare(ad20, "Посредник") == 0)
                {
                    objectAdd.ID_WINSTATUS = 2;
                }
                else if (String.Compare(ad20, "Производитель") == 0)
                {
                    objectAdd.ID_WINSTATUS = 1;
                }
                else
                {
                    objectAdd.ID_WINSTATUS = null;
                }
                if (ad21 != "")
                {
                    objectAdd.PURCHASE_VOLUME = Convert.ToString(ad21);
                }
                else
                {
                    objectAdd.PURCHASE_VOLUME = null;
                }

                if (ad22 != "")
                {
                    objectAdd.VOLUME_UNITS = Convert.ToString(ad22);
                }
                else
                {
                    objectAdd.VOLUME_UNITS = null;
                }

                ad23 = ad23.Replace(".", ",");
                if (ad23 != "")
                {
                    objectAdd.PRICE_PER_ITEM = Convert.ToString(ad23);
                }
                else
                {
                    objectAdd.PRICE_PER_ITEM = null;
                }

                if (ad24 != "")
                {
                    objectAdd.DELIVERY_COND = Convert.ToString(ad24);
                }
                else
                {
                    objectAdd.DELIVERY_COND = null;
                }

                if (ad25 != "")
                {
                    objectAdd.MAIN_TEC_SPECS = Convert.ToString(ad25);
                }
                else
                {
                    objectAdd.MAIN_TEC_SPECS = null;
                }

                if (ad26 != "")
                {
                    ad26 = ad26.ToUpper();
                    ProcurementInfoEntities context = new ProcurementInfoEntities();
                    var table = context.S_Country.Where(c => c.NAME == ad26);
                    if (table != null)
                    {
                        foreach (var c in table)
                        {
                            objectAdd.ID_COUNTRY_ORIGIN = c.ID;
                        }
                    }
                    else
                    {
                        objectAdd.ID_COUNTRY_ORIGIN = null;
                    }
                }
                else
                {
                    objectAdd.ID_COUNTRY_ORIGIN = null;
                }

                if (ad27 != "")
                {
                    objectAdd.MANUFACTURER = Convert.ToString(ad27);
                }
                else
                {
                    objectAdd.MANUFACTURER = null;
                }


                if (ad28 != "")
                {
                    ProcurementInfoEntities context = new ProcurementInfoEntities();
                    var table = context.S_SubDivision.Where(c => c.NAME == ad28);
                    if (table != null)
                    {
                        foreach (var c in table)
                        {
                            objectAdd.ID_SUBDIVISION = c.ID;
                        }
                    }
                    else
                    {
                        objectAdd.ID_SUBDIVISION = null;
                    }
                }
                else
                {
                    objectAdd.ID_SUBDIVISION = null;
                }

                if (ad27 != "")
                {
                    objectAdd.MANUFACTURER = Convert.ToString(ad27);
                }
                else
                {
                    objectAdd.MANUFACTURER = null;
                }


                objectAdd.DATE_ACTIVE = DateTime.Now;
                objectAdd.FLAG = true;
                objectAdd.USER_ENTER = User.Identity.Name + " " + DateTime.Now.ToString("dd.MM.yyyy H:mm:ss");

                _databaseRepository.T_ProcurementInformation.Add(objectAdd);
                _databaseRepository.SaveChanges();

                return Json(1);
            }
            else
            {
                return Json(2);
            }
        }
        #endregion

        #region Добавление записи в талицу в отдельном окне до 1000 БВ
        public ActionResult InsDo1k(CRUDModel<T_PrInfoDo1000> Object, string add1, string add2, string add3, string add4, string add5, string add6, string add7, string add8, string add9, string add10, string add11, string add12, string add13, string add14)
        {
            T_PrInfoDo1000 objectAdd = new T_PrInfoDo1000();

            if (add1 != "" && add2 != "" && add4 != "" && add6 != "")
            {
                if (add1 != "")
                {
                    objectAdd.NAME_PR = add1;
                }
                else
                {
                    objectAdd.NAME_PR = null;
                }

                if (String.Compare(add2, "Работы (услуги)") == 0)
                {
                    objectAdd.ID_SUBJECTPURCHASE = 2;
                }
                else
                {
                    objectAdd.ID_SUBJECTPURCHASE = 1;
                }

                if (add3 != "")
                {
                    objectAdd.CONTRACT_NUMBER = Convert.ToString(add3);
                }
                else
                {
                    objectAdd.CONTRACT_NUMBER = null;
                }

                objectAdd.FIO_ISP = Convert.ToString(add4);

                if (String.Compare(add5, "Официальный представитель") == 0)
                {
                    objectAdd.ID_WINSTATUSDO1000 = 1;
                }
                else if (String.Compare(add5, "Посредник") == 0)
                {
                    objectAdd.ID_WINSTATUSDO1000 = 2;
                }
                else if (String.Compare(add5, "Производитель") == 0)
                {
                    objectAdd.ID_WINSTATUSDO1000 = 3;
                }
                else
                {
                    objectAdd.ID_WINSTATUSDO1000 = null;
                }

                if (add6 != "")
                {
                    ProcurementInfoEntities context = new ProcurementInfoEntities();
                    var table = context.S_LegislDo1000.Where(c => c.NAME == add6);
                    if (table != null)
                    {
                        foreach (var c in table)
                        {
                            objectAdd.ID_LEGISLDO1000 = c.ID;
                        }
                    }
                }
                else
                {
                    objectAdd.ID_LEGISLDO1000 = null;
                }

                if (add7 != "")
                {
                    objectAdd.DATE_CONCLUSION = Convert.ToDateTime(add7);
                }
                else
                {
                    objectAdd.DATE_CONCLUSION = DateTime.Now;
                }

                if (add8 != "")
                {
                    objectAdd.WIN_VALUE = Convert.ToDecimal(add8.Replace(".", ","));
                }
                else
                {
                    objectAdd.WIN_VALUE = null;
                }

                if (add9 != "")
                {
                    var cur = GetAllCurrency().Where(c => c.LETTER_CODE == add9).FirstOrDefault();
                    if (cur != null)
                    {
                        objectAdd.ID_CURRENCY = cur.ID;                        
                    }
                    else
                    {
                        objectAdd.ID_CURRENCY = null;
                    }
                }
                else
                {
                    objectAdd.ID_CURRENCY = null;
                }

                if (add10 != "")
                {
                    objectAdd.WIN_VALUE_NDE = Convert.ToDecimal(add10.Replace(",", "."));
                }
                else
                {
                    objectAdd.WIN_VALUE_NDE = null;
                }

                if (add11 != "")
                {
                    objectAdd.WIN_VALUE_NNDS = Convert.ToDecimal(add11.Replace(",", "."));
                }
                else
                {
                    objectAdd.WIN_VALUE_NNDS = null;
                }

                if (add12 != "")
                {
                    objectAdd.WIN_NAME = Convert.ToString(add12);
                }
                else
                {
                    objectAdd.WIN_NAME = null;
                }

                add13 = add13.ToUpper();
                if (add13 != "")
                {
                    ProcurementInfoEntities context = new ProcurementInfoEntities();
                    var table = context.S_Country.Where(c => c.NAME == add13);
                    if (table != null)
                    {
                        foreach (var c in table)
                        {
                            objectAdd.ID_COUNTRY_ORIGIN = c.ID;
                        }
                    }
                }
                else
                {
                    objectAdd.ID_COUNTRY_ORIGIN = null;
                }

                if (add14 != "")
                {
                    ProcurementInfoEntities context = new ProcurementInfoEntities();
                    var table = context.S_SubDivision.Where(c => c.NAME == add14);
                    if (table != null)
                    {
                        foreach (var c in table)
                        {
                            objectAdd.ID_SUBDIVISION = c.ID;
                        }
                    }
                }
                else
                {
                    objectAdd.ID_SUBDIVISION = null;
                }

                objectAdd.DATE_ACTIVE = DateTime.Now;
                objectAdd.FLAG = true;
                objectAdd.USER_ENTER = Convert.ToString(User.Identity.Name) + " " + DateTime.Now.ToString("dd.MM.yyyy H:mm:ss") + " " + "Add";

                _databaseRepository.T_PrInfoDo1000.Add(objectAdd);
                _databaseRepository.SaveChanges();
                return Json(1);
            }
            else
            {
                return Json(2);
            }
        }
        #endregion

        #region Импорт из Excel свыше 1000 БВ
        [HttpPost]
        public JsonResult Upload()
        {
            foreach (string file in Request.Files)
            {
                var upload = Request.Files[file];
                if (upload != null)
                {
                    upload.SaveAs(Server.MapPath(@"~/images/Upload.xlsx"));
                    string xsltPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath(@"~/images/Upload.xlsx"));
                    var workbook = new XLWorkbook(xsltPath);
                    var worksheet = workbook.Worksheet(3);
                    T_ProcurementInformation objectToUpdate = new T_ProcurementInformation();
                    ProcurementInfoEntities context = new ProcurementInfoEntities();
                    var rows = worksheet.RangeUsed().RowsUsed();
                    int i = 1;
                    string buf;
                    foreach (var row in rows)
                    {
                        if(i > 2 & Convert.ToString(row.Cell(1).Value) != "" & Convert.ToString(row.Cell(2).Value) != "" & Convert.ToString(row.Cell(4).Value) != "" & Convert.ToString(row.Cell(5).Value) != "" & Convert.ToString(row.Cell(7).Value) != "")
                        {
                            if (Convert.ToString(row.Cell(1).Value) != "")
                            {
                                objectToUpdate.NAME_PR = Convert.ToString(row.Cell(1).Value);
                            }
                            else
                            {
                                objectToUpdate.NAME_PR = null;
                            }

                            if (Convert.ToString(row.Cell(2).Value) != "")
                            {
                                objectToUpdate.ID_SUBJECTPURCHASE = Convert.ToInt32(row.Cell(2).Value);
                            }
                            else
                            {
                                objectToUpdate.ID_SUBJECTPURCHASE = null;
                            }

                            if (Convert.ToString(row.Cell(3).Value) != "")
                            {
                                objectToUpdate.INVITE_NUN = Convert.ToString(row.Cell(3).Value);
                            }
                            else
                            {
                                objectToUpdate.INVITE_NUN = null;
                            }

                            if (Convert.ToString(row.Cell(4).Value) != "")
                            {
                                objectToUpdate.FIO_ISP = Convert.ToString(row.Cell(4).Value);
                            }
                            else
                            {
                                objectToUpdate.FIO_ISP = null;
                            }

                            if (Convert.ToString(row.Cell(5).Value) != "")
                            {
                                objectToUpdate.ID_LEGISLATION = Convert.ToInt32(row.Cell(5).Value);
                            }
                            else
                            {
                                objectToUpdate.ID_LEGISLATION = null;
                            }

                            if (Convert.ToString(row.Cell(6).Value) != "")
                            {
                                objectToUpdate.PARTICIPANT__VALUE = Convert.ToInt32(row.Cell(6).Value);
                            }
                            else
                            {
                                objectToUpdate.PARTICIPANT__VALUE = null;
                            }

                            if (Convert.ToString(row.Cell(7).Value) != "")
                            {
                                objectToUpdate.PR_APPROVAL_DATE = Convert.ToDateTime(row.Cell(7).Value);
                            }
                            else
                            {
                                objectToUpdate.PR_APPROVAL_DATE = DateTime.Now;
                            }

                            if (Convert.ToString(row.Cell(8).Value) != "")
                            {
                                objectToUpdate.ID_RESULT = Convert.ToInt32(row.Cell(8).Value);
                            }
                            else
                            {
                                objectToUpdate.ID_RESULT = null;
                            }

                            if (objectToUpdate.ID_LEGISLATION.Value == 3)
                            {
                                objectToUpdate.CONTRACT_DATE_TERM = null;
                                objectToUpdate.RESULT_DATE_TERM = null;
                            }
                            else
                            {
                                objectToUpdate.CONTRACT_DATE_TERM = Convert.ToDateTime(objectToUpdate.PR_APPROVAL_DATE).AddDays(28);
                                objectToUpdate.RESULT_DATE_TERM = Convert.ToDateTime(objectToUpdate.PR_APPROVAL_DATE).AddDays(31);
                            }

                            if (objectToUpdate.ID_RESULT.Value == 4)
                            {
                                objectToUpdate.CONTRACT_DATE_TERM = null;
                                objectToUpdate.RESULT_DATE_TERM = null;
                            }

                            if (Convert.ToString(row.Cell(10).Value) != "")
                            {
                                objectToUpdate.CONTRACT_DATE_PROL = Convert.ToDateTime(row.Cell(9).Value);
                            }
                            else
                            {
                                objectToUpdate.CONTRACT_DATE_PROL = null;
                            }

                            if (Convert.ToString(row.Cell(11).Value) != "")
                            {
                                objectToUpdate.CONTRACT_DATE_FACT = Convert.ToDateTime(row.Cell(10).Value);
                            }
                            else
                            {
                                objectToUpdate.CONTRACT_DATE_FACT = null;
                            }

                            if (Convert.ToString(row.Cell(12).Value) != "")
                            {
                                objectToUpdate.CONTRACT_NUMBER = Convert.ToString(row.Cell(11).Value);
                            }
                            else
                            {
                                objectToUpdate.CONTRACT_NUMBER = null;
                            }

                            if (Convert.ToString(row.Cell(14).Value) != "")
                            {
                                objectToUpdate.RESULT_DATE_FACT = Convert.ToDateTime(row.Cell(13).Value);
                            }
                            else
                            {
                                objectToUpdate.RESULT_DATE_FACT = null;
                            }

                            if (Convert.ToString(row.Cell(15).Value) != "")
                            {
                                objectToUpdate.WIN_NAME = Convert.ToString(row.Cell(13).Value);
                            }
                            else
                            {
                                objectToUpdate.WIN_NAME = null;
                            }

                            if (Convert.ToString(row.Cell(16).Value) != "")
                            {
                                objectToUpdate.ID_COUNTRY = Convert.ToInt32(row.Cell(16).Value);
                            }
                            else
                            {
                                objectToUpdate.ID_COUNTRY = null;
                            }

                            if (Convert.ToString(row.Cell(17).Value) != "")
                            {
                                objectToUpdate.WIN_VALUE = Convert.ToDecimal(row.Cell(17).Value);
                            }
                            else
                            {
                                objectToUpdate.WIN_VALUE = null;
                            }

                            if (Convert.ToString(row.Cell(18).Value) != "")
                            {
                                objectToUpdate.ID_CURRENCY = Convert.ToInt32(row.Cell(18).Value);
                            }
                            else
                            {
                                objectToUpdate.ID_CURRENCY = null;
                            }

                            buf = Convert.ToString(row.Cell(19).Value);
                            if (Convert.ToString(row.Cell(19).Value) != "")
                            {
                                objectToUpdate.WIN_VALUE_BYN = float.Parse(buf);
                            }
                            else
                            {
                                objectToUpdate.WIN_VALUE_BYN = null;
                            }

                            if (Convert.ToString(row.Cell(20).Value) != "")
                            {
                                objectToUpdate.ID_WINSTATUS = Convert.ToInt32(row.Cell(20).Value);
                            }
                            else
                            {
                                objectToUpdate.ID_WINSTATUS = null;
                            }

                            if (Convert.ToString(row.Cell(21).Value) != "")
                            {
                                objectToUpdate.PURCHASE_VOLUME = Convert.ToString(row.Cell(21).Value);
                            }
                            else
                            {
                                objectToUpdate.PURCHASE_VOLUME = null;
                            }

                            if (Convert.ToString(row.Cell(22).Value) != "")
                            {
                                objectToUpdate.VOLUME_UNITS = Convert.ToString(row.Cell(22).Value);
                            }
                            else
                            {
                                objectToUpdate.VOLUME_UNITS = null;
                            }

                            if (Convert.ToString(row.Cell(23).Value) != "")
                            {
                                objectToUpdate.PRICE_PER_ITEM = Convert.ToString(row.Cell(23).Value);
                            }
                            else
                            {
                                objectToUpdate.PRICE_PER_ITEM = null;
                            }

                            if (Convert.ToString(row.Cell(24).Value) != "")
                            {
                                objectToUpdate.DELIVERY_COND = Convert.ToString(row.Cell(24).Value);
                            }
                            else
                            {
                                objectToUpdate.DELIVERY_COND = null;
                            }

                            if (Convert.ToString(row.Cell(25).Value) != "")
                            {
                                objectToUpdate.MAIN_TEC_SPECS = Convert.ToString(row.Cell(25).Value);
                            }
                            else
                            {
                                objectToUpdate.MAIN_TEC_SPECS = null;
                            }

                            if (Convert.ToString(row.Cell(26).Value) != "")
                            {
                                objectToUpdate.ID_COUNTRY_ORIGIN = Convert.ToInt32(row.Cell(26).Value);
                            }
                            else
                            {
                                objectToUpdate.ID_COUNTRY_ORIGIN = null;
                            }

                            if (Convert.ToString(row.Cell(27).Value) != "")
                            {
                                objectToUpdate.MANUFACTURER = Convert.ToString(row.Cell(27).Value);
                            }
                            else
                            {
                                objectToUpdate.MANUFACTURER = null;
                            }

                            objectToUpdate.DATE_ACTIVE = DateTime.Now;
                            objectToUpdate.FLAG = true;
                            objectToUpdate.USER_ENTER = User.Identity.Name + " " + DateTime.Now.ToString("dd.MM.yyyy H:mm:ss");

                            _databaseRepository.T_ProcurementInformation.Add(objectToUpdate);
                            _databaseRepository.SaveChanges();
                        }
                        i++;
                    } 
                }
                System.IO.File.Delete(Server.MapPath(@"~/images/Upload.xlsx"));
            }
            return Json("Файл загружен!");
        }
        #endregion

        #region Импорт из Excel до 1000 БВ
        public JsonResult UploadDo1k()
        {
            foreach (string file in Request.Files)
            {
                var upload = Request.Files[file];
                if (upload != null)
                {
                    upload.SaveAs(Server.MapPath(@"~/images/UploadDo1k.xlsx"));
                    string xsltPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath(@"~/images/UploadDo1k.xlsx"));
                    var workbook = new XLWorkbook(xsltPath);
                    var worksheet = workbook.Worksheet(3);
                    T_PrInfoDo1000 objectToUpdate = new T_PrInfoDo1000();
                    ProcurementInfoEntities context = new ProcurementInfoEntities();
                    var rows = worksheet.RangeUsed().RowsUsed();
                    int i = 1;
                    foreach (var row in rows)
                    {
                        if ( i > 1 & Convert.ToString(row.Cell(1).Value) != "" & Convert.ToString(row.Cell(2).Value) != "" & Convert.ToString(row.Cell(4).Value) != "" & Convert.ToString(row.Cell(6).Value) != "")
                        {
                            if (Convert.ToString(row.Cell(1).Value) != "")
                            {
                                objectToUpdate.NAME_PR = Convert.ToString(row.Cell(1).Value);
                            }
                            else
                            {
                                objectToUpdate.NAME_PR = null;
                            }

                            if(Convert.ToString(row.Cell(2).Value) != "")
                            {
                                objectToUpdate.ID_SUBJECTPURCHASE = Convert.ToInt32(row.Cell(2).Value);
                            }
                            else
                            {
                                objectToUpdate.ID_SUBJECTPURCHASE = null;
                            }

                            if (Convert.ToString(row.Cell(3).Value) != "")
                            {
                                objectToUpdate.CONTRACT_NUMBER = Convert.ToString(row.Cell(3).Value);
                            }
                            else
                            {
                                objectToUpdate.CONTRACT_NUMBER = null;
                            }

                            if (Convert.ToString(row.Cell(4).Value) != "")
                            {
                                objectToUpdate.FIO_ISP = Convert.ToString(row.Cell(4).Value); 
                            }
                            else
                            {
                                objectToUpdate.FIO_ISP = null;
                            }

                            if (Convert.ToString(row.Cell(5).Value) != "")
                            {
                                objectToUpdate.ID_WINSTATUSDO1000 = Convert.ToInt32(row.Cell(5).Value);
                            }
                            else
                            {
                                objectToUpdate.ID_WINSTATUSDO1000 = null;
                            }

                            if (Convert.ToString(row.Cell(6).Value) != "")
                            {
                                objectToUpdate.ID_LEGISLDO1000 = Convert.ToInt32(row.Cell(6).Value);
                            }
                            else
                            {
                                objectToUpdate.ID_LEGISLDO1000 = null;
                            }

                            if (Convert.ToString(row.Cell(7).Value) != "")
                            {
                                objectToUpdate.DATE_CONCLUSION = Convert.ToDateTime(row.Cell(7).Value);
                            }
                            else
                            {
                                objectToUpdate.DATE_CONCLUSION = DateTime.Now;
                            }

                            if (Convert.ToString(row.Cell(8).Value) != "")
                            {
                                objectToUpdate.WIN_VALUE = Convert.ToDecimal(row.Cell(8).Value);
                            }
                            else
                            {
                                objectToUpdate.WIN_VALUE = null;
                            }

                            if (Convert.ToString(row.Cell(9).Value) != "" )
                            {
                                objectToUpdate.ID_CURRENCY = Convert.ToInt32(row.Cell(9).Value);
                            }
                            else
                            {
                                objectToUpdate.ID_CURRENCY = null;
                            }

                            if (Convert.ToString(row.Cell(10).Value) != "")
                            {
                                objectToUpdate.WIN_VALUE_NDE = Convert.ToDecimal(row.Cell(10).Value);
                            }
                            else
                            {
                                objectToUpdate.WIN_VALUE_NDE = null;
                            }

                            if (Convert.ToString(row.Cell(11).Value) != "")
                            {
                                objectToUpdate.WIN_VALUE_NNDS = Convert.ToDecimal(row.Cell(11).Value);
                            }
                            else
                            {
                                objectToUpdate.WIN_VALUE_NNDS = null;
                            }

                            if (Convert.ToString(row.Cell(12).Value) != "")
                            {
                                objectToUpdate.WIN_NAME = Convert.ToString(row.Cell(12).Value);
                            }
                            else
                            {
                                objectToUpdate.WIN_NAME = null;
                            }

                            if (Convert.ToString(row.Cell(13).Value) != "")
                            {
                                objectToUpdate.ID_COUNTRY_ORIGIN = Convert.ToInt32(row.Cell(13).Value);
                            }
                            else
                            {
                                objectToUpdate.ID_COUNTRY_ORIGIN = null;
                            }

                            objectToUpdate.DATE_ACTIVE = DateTime.Now;
                            objectToUpdate.FLAG = true;
                            objectToUpdate.USER_ENTER = User.Identity.Name + " " + DateTime.Now.ToString("dd.MM.yyyy H:mm:ss");

                            _databaseRepository.T_PrInfoDo1000.Add(objectToUpdate);
                            _databaseRepository.SaveChanges();
                        }
                        i++;
                    }
                }
                System.IO.File.Delete(Server.MapPath(@"~/images/UploadDo1k.xlsx"));
            }
            return Json("Файл загружен!");
        }
        #endregion

        #region Чтение дат формирования отчета
        public ActionResult DateCreate()
        {
            string line;
            StreamReader sr = new StreamReader(Server.MapPath(@"~/images/DateUsd.txt"));
            line = sr.ReadLine();
            sr.Close();
            return Json(line);
            
        }

        public ActionResult InfoReport()
        {
            string line;
            StreamReader sr = new StreamReader(Server.MapPath(@"~/images/DateUsd.txt"));
            line = sr.ReadLine();
            line = sr.ReadLine();
            return Json(line);
            sr.Close();
        }
        #endregion

        #region Выпадающие листы DataGrid
        private IQueryable<S_WinStatus> GetAllWinStatus()
        {
            return _databaseRepository.S_WinStatus.Where(f => f.FLAG == null || f.FLAG == true);
        }

        private IQueryable<S_Legislation> GetAllLegislation()
        {
            return _databaseRepository.S_Legislation.Where(f => f.FLAG == null || f.FLAG == true);
        }

        private IQueryable<S_SubjectPurchase> GetAllSubjectPurchase()
        {
            return _databaseRepository.S_SubjectPurchase.Where(f => f.FLAG == null || f.FLAG == true);
        }

        private IQueryable<S_Result> GetAllResult()
        {
            return _databaseRepository.S_Result;
        }

        private IQueryable<S_SubDivision> GetAllSubDivision()
        {
            return _databaseRepository.S_SubDivision;
        }

        private IQueryable<S_Currency> GetAllCurrency()
        {
            return _databaseRepository.S_Currency.Where(f => f.FLAG == null || f.FLAG == true);
        }
        
        private IQueryable<S_Country> GetAllCountry()
        {
            return _databaseRepository.S_Country;
        }
        private IQueryable<S_WinStatus> GetAllWinStatusDo1000()
        {
            return _databaseRepository.S_WinStatus.Where(f => f.FLAG == null || f.FLAG == true);
        }
        private IQueryable<S_LegislDo1000> GetAllLegislDo1000()
        {
            return _databaseRepository.S_LegislDo1000.Where(f => f.FLAG == null || f.FLAG == true);
        }
        #endregion
    }
}