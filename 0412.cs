
        public ActionResult PRINT_3011(string UrlId)
        {
           try
           {
               AYHttpParam ap = new AYHttpParam(this.HttpContext);
               int iSCH_YEAR = ap.IntParam("iSCH_YEAR");
               int iSCH_YEAR1 = iSCH_YEAR - 1;
               int iSCH_YEAR2 = iSCH_YEAR - 2;
               int iSCH_YEAR3 = iSCH_YEAR - 3;
               int iSCH_YEAR4 = iSCH_YEAR - 4;
               int iSCH_YEAR5 = iSCH_YEAR - 5;

               var dt1 = (from t1 in db.SCHOOLS
                          where t1.bSCH_CANCEL == false
                             && t1.bSCH_STATUS == true
                             && t1.iSCH_LEVEL == 1
                          orderby t1.iSCH_LEVEL, t1.cSCH_CALLED
                          select t1).ToList();

               var dt2 = (from t1 in db.SCHOOLS
                          where t1.bSCH_CANCEL == false
                             && t1.bSCH_STATUS == true
                             && t1.iSCH_LEVEL == 2
                          orderby t1.iSCH_LEVEL, t1.cSCH_CALLED
                          select t1).ToList();

               var ORAL_CAVITY_L = (from t4 in db.ORAL_CAVITY
                                    from t5 in db.SCHOOLS
                                    where t4.cGOV_FK == CurrUser.GOV_PK
                                       && t4.bORCA_CANCEL == false
                                       && t4.iSCH_YEAR <= iSCH_YEAR
                                       && t4.iSCH_YEAR >= iSCH_YEAR5
                                       && t4.iORCA_DATA2 != null
                                       && t4.iSCH_FK == t5.iSCH_PK
                                       && t5.bSCH_CANCEL == false
                                       && t5.bSCH_STATUS == true
                                       && (t5.iSCH_LEVEL == 1 || t5.iSCH_LEVEL == 2)
                                    select t4).ToList();

               var TEST_MAIN_L = (from t1 in db.TEST_MAIN
                                  where t1.cGOV_FK == CurrUser.GOV_PK
                                     && t1.iSCH_YEAR == iSCH_YEAR
                                     && t1.cSCH_TOPIC == "01"
                                     //&& t1.cTM_GROUP03 != null
                                     && t1.bTM_CANCEL == false
                                     && t1.bTM_OVER == true
                                  orderby t1.iSCH_FK, t1.iGRADE_NUM, t1.cCLS_NO, t1.cSTU_NO, t1.iTM_KIND
                                  select t1).ToList();

               var DataTEST_ToSchool = TEST_MAIN_L.Where(t => t.iTM_KIND == 1)//.Where(t => t.iTM_KIND == 2)
                                                   .GroupBy(t => new { t.iSCH_FK })
                                                   .Select(t => new
                                                   {
                                                       t.Key.iSCH_FK,
                                                       TM_GROUP02 = t.Average(y => Convert.ToDecimal(y.cTM_GROUP02))//cTM_GROUP03
                                                   }).ToList();

               DataTable dt_ANS = AYTools.GetRptTable(280);
               if (ORAL_CAVITY_L.Count > 0)
               {
                   DataRow NewRow = dt_ANS.NewRow();

                   NewRow["DataColumn1"] = "新竹市";

                   NewRow["DataColumn11"] = iSCH_YEAR;
                   //初檢齲齒
                   NewRow["DataColumn12"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA2).Average();
                   NewRow["DataColumn13"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_LEVEL == 2)
                                                         .Select(t => t.iORCA_DATA2).Average();
                   //矯治
                   NewRow["DataColumn14"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA3).Average();
                   NewRow["DataColumn15"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_LEVEL == 2)
                                                         .Select(t => t.iORCA_DATA31).Average();
                   //初檢齲齒-一、四年級
                   NewRow["DataColumn16"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA21).Average();
                   NewRow["DataColumn17"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA22).Average();

                   NewRow["DataColumn21"] = iSCH_YEAR1;
                   NewRow["DataColumn22"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR1 && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA2).Average();
                   NewRow["DataColumn23"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR1 && t.iSCH_LEVEL == 2)
                                                         .Select(t => t.iORCA_DATA2).Average();
                   NewRow["DataColumn24"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR1 && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA3).Average();
                   NewRow["DataColumn25"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR1 && t.iSCH_LEVEL == 2)
                                                         .Select(t => t.iORCA_DATA31).Average();

                   NewRow["DataColumn31"] = iSCH_YEAR2;
                   NewRow["DataColumn32"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR2 && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA2).Average();
                   NewRow["DataColumn33"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR2 && t.iSCH_LEVEL == 2)
                                                         .Select(t => t.iORCA_DATA2).Average();
                   NewRow["DataColumn34"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR2 && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA3).Average();
                   NewRow["DataColumn35"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR2 && t.iSCH_LEVEL == 2)
                                                         .Select(t => t.iORCA_DATA31).Average();

                   NewRow["DataColumn41"] = iSCH_YEAR3;
                   NewRow["DataColumn42"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR3 && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA2).Average();
                   NewRow["DataColumn43"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR3 && t.iSCH_LEVEL == 2)
                                                         .Select(t => t.iORCA_DATA2).Average();
                   NewRow["DataColumn44"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR3 && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA3).Average();
                   NewRow["DataColumn45"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR3 && t.iSCH_LEVEL == 2)
                                                         .Select(t => t.iORCA_DATA31).Average();

                   NewRow["DataColumn51"] = iSCH_YEAR4;
                   NewRow["DataColumn52"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR4 && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA2).Average();
                   NewRow["DataColumn53"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR4 && t.iSCH_LEVEL == 2)
                                                         .Select(t => t.iORCA_DATA2).Average();
                   NewRow["DataColumn54"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR4 && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA3).Average();
                   NewRow["DataColumn55"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR4 && t.iSCH_LEVEL == 2)
                                                         .Select(t => t.iORCA_DATA31).Average();

                   NewRow["DataColumn61"] = iSCH_YEAR5;
                   NewRow["DataColumn62"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR5 && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA2).Average();
                   NewRow["DataColumn63"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR5 && t.iSCH_LEVEL == 2)
                                                         .Select(t => t.iORCA_DATA2).Average();
                   NewRow["DataColumn64"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR5 && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA3).Average();
                   NewRow["DataColumn65"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR5 && t.iSCH_LEVEL == 2)
                                                         .Select(t => t.iORCA_DATA31).Average();

                   //var DataTEST_31 = TEST_MAIN_L.Where(t => t.iGRADE_NUM == 3 && t.iTM_KIND == 1).ToList();
                   //var DataTEST_32 = TEST_MAIN_L.Where(t => t.iGRADE_NUM == 3 && t.iTM_KIND == 2).ToList();
                   var DataTEST_41 = TEST_MAIN_L.Where(t => t.iGRADE_NUM == 4 && t.iTM_KIND == 1).ToList();
                   var DataTEST_42 = TEST_MAIN_L.Where(t => t.iGRADE_NUM == 4 && t.iTM_KIND == 2).ToList();
                   var DataTEST_51 = TEST_MAIN_L.Where(t => t.iGRADE_NUM == 5 && t.iTM_KIND == 1).ToList();
                   var DataTEST_52 = TEST_MAIN_L.Where(t => t.iGRADE_NUM == 5 && t.iTM_KIND == 2).ToList();
                   var DataTEST_71 = TEST_MAIN_L.Where(t => t.iGRADE_NUM == 7 && t.iTM_KIND == 1).ToList();
                   var DataTEST_72 = TEST_MAIN_L.Where(t => t.iGRADE_NUM == 7 && t.iTM_KIND == 2).ToList();

                   NewRow["DataColumn71"] = ReDataTEST(DataTEST_41, "GROUP02").Select(t => Convert.ToDecimal(t.cTM_GROUP02)).Average();//108版本-cTM_GROUP03
                   NewRow["DataColumn72"] = ReDataTEST(DataTEST_42, "GROUP02").Select(t => Convert.ToDecimal(t.cTM_GROUP02)).Average();
                   NewRow["DataColumn73"] = ReDataTEST(DataTEST_51, "GROUP02").Select(t => Convert.ToDecimal(t.cTM_GROUP02)).Average();
                   NewRow["DataColumn74"] = ReDataTEST(DataTEST_52, "GROUP02").Select(t => Convert.ToDecimal(t.cTM_GROUP02)).Average();
                   NewRow["DataColumn75"] = ReDataTEST(DataTEST_71, "GROUP02").Select(t => Convert.ToDecimal(t.cTM_GROUP02)).Average();
                   NewRow["DataColumn76"] = ReDataTEST(DataTEST_72, "GROUP02").Select(t => Convert.ToDecimal(t.cTM_GROUP02)).Average();

                   //dt_SUB6
                   //NewRow["DataColumn81"] = DataTEST_ToSchool.FirstOrDefault()?.cSCH_NAME;
                   //NewRow["DataColumn82"] = DataTEST_ToSchool.FirstOrDefault()?.TM_GROUP03;
                   //NewRow["DataColumn83"] = DataTEST_ToSchool.Skip(1).FirstOrDefault()?.cSCH_NAME;
                   //NewRow["DataColumn84"] = DataTEST_ToSchool.Skip(1).FirstOrDefault()?.TM_GROUP03;
                   //NewRow["DataColumn85"] = DataTEST_ToSchool.Skip(2).FirstOrDefault()?.cSCH_NAME;
                   //NewRow["DataColumn86"] = DataTEST_ToSchool.Skip(2).FirstOrDefault()?.TM_GROUP03;

                   NewRow["DataColumn91"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA4).Average();
                   //NewRow["DataColumn92"] = "";
                   NewRow["DataColumn93"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA5).Average();
                   //NewRow["DataColumn94"] = "";
                   NewRow["DataColumn95"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA51 == 0 && t.iORCA_DATA5 == 0 ? 0 : t.iORCA_DATA51 / t.iORCA_DATA5).Average();
                   NewRow["DataColumn96"] = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_LEVEL == 1)
                                                         .Select(t => t.iORCA_DATA51).Sum();

                   NewRow["DataColumn101"] = ReDataTEST(DataTEST_41, "GROUP03").Select(t => Convert.ToDecimal(t.cTM_GROUP03)).Average();//108版本-cTM_GROUP04
                   NewRow["DataColumn102"] = ReDataTEST(DataTEST_42, "GROUP03").Select(t => Convert.ToDecimal(t.cTM_GROUP03)).Average();
                   NewRow["DataColumn103"] = ReDataTEST(DataTEST_51, "GROUP03").Select(t => Convert.ToDecimal(t.cTM_GROUP03)).Average();
                   NewRow["DataColumn104"] = ReDataTEST(DataTEST_52, "GROUP03").Select(t => Convert.ToDecimal(t.cTM_GROUP03)).Average();
                   NewRow["DataColumn105"] = ReDataTEST(DataTEST_71, "GROUP03").Select(t => Convert.ToDecimal(t.cTM_GROUP03)).Average();
                   NewRow["DataColumn106"] = ReDataTEST(DataTEST_72, "GROUP03").Select(t => Convert.ToDecimal(t.cTM_GROUP03)).Average();

                   //108版本-cTM_GROUP05
                   decimal? GROUP04_1 = ReDataTEST(DataTEST_51, "GROUP04").Select(t => Convert.ToDecimal(t.cTM_GROUP04)).Average();
                   decimal? GROUP04_2 = ReDataTEST(DataTEST_52, "GROUP04").Select(t => Convert.ToDecimal(t.cTM_GROUP04)).Average();

                   NewRow["DataColumn111"] = GROUP04_1;
                   NewRow["DataColumn112"] = GROUP04_2;
                   //NewRow["DataColumn113"] = GROUP04_1 == null || GROUP04_2 == null ? "" :
                   //                         (GROUP04_2 - GROUP04_1) >= 5 ? "達成" : "未達成";

                   //108版本-cTM_GROUP06
                   NewRow["DataColumn121"] = ReDataTEST(DataTEST_41, "GROUP05").Select(t => Convert.ToDecimal(t.cTM_GROUP05)).Average();
                   NewRow["DataColumn122"] = ReDataTEST(DataTEST_42, "GROUP05").Select(t => Convert.ToDecimal(t.cTM_GROUP05)).Average();
                   NewRow["DataColumn123"] = ReDataTEST(DataTEST_51, "GROUP05").Select(t => Convert.ToDecimal(t.cTM_GROUP05)).Average();
                   NewRow["DataColumn124"] = ReDataTEST(DataTEST_52, "GROUP05").Select(t => Convert.ToDecimal(t.cTM_GROUP05)).Average();
                   NewRow["DataColumn125"] = ReDataTEST(DataTEST_71, "GROUP05").Select(t => Convert.ToDecimal(t.cTM_GROUP05)).Average();
                   NewRow["DataColumn126"] = ReDataTEST(DataTEST_72, "GROUP05").Select(t => Convert.ToDecimal(t.cTM_GROUP05)).Average();

                   //108版本-cTM_GROUP07
                   NewRow["DataColumn131"] = ReDataTEST(DataTEST_41, "GROUP06").Select(t => Convert.ToDecimal(t.cTM_GROUP06)).Average();
                   NewRow["DataColumn132"] = ReDataTEST(DataTEST_42, "GROUP06").Select(t => Convert.ToDecimal(t.cTM_GROUP06)).Average();
                   NewRow["DataColumn133"] = ReDataTEST(DataTEST_51, "GROUP06").Select(t => Convert.ToDecimal(t.cTM_GROUP06)).Average();
                   NewRow["DataColumn134"] = ReDataTEST(DataTEST_52, "GROUP06").Select(t => Convert.ToDecimal(t.cTM_GROUP06)).Average();
                   NewRow["DataColumn135"] = ReDataTEST(DataTEST_71, "GROUP06").Select(t => Convert.ToDecimal(t.cTM_GROUP06)).Average();
                   NewRow["DataColumn136"] = ReDataTEST(DataTEST_72, "GROUP06").Select(t => Convert.ToDecimal(t.cTM_GROUP06)).Average();

                   var Data17_0 = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_LEVEL == 1);
                   //var Data17_1 = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR1 && t.iSCH_LEVEL == 1);

                   NewRow["DataColumn141"] = (int)Data17_0.Select(t => t.iORCA_DATA81).Sum();//四
                   NewRow["DataColumn142"] = (int)Data17_0.Select(t => t.iORCA_DATA82).Sum();
                   NewRow["DataColumn143"] = Data17_0.Select(t => t.iORCA_DATA83).Average();
                   NewRow["DataColumn144"] = (int)Data17_0.Select(t => t.iORCA_DATA61).Sum();//一
                   NewRow["DataColumn145"] = (int)Data17_0.Select(t => t.iORCA_DATA62).Sum();
                   NewRow["DataColumn146"] = Data17_0.Select(t => t.iORCA_DATA63).Average();
                   NewRow["DataColumn147"] = (int)Data17_0.Select(t => t.iORCA_DATA71).Sum();//二
                   NewRow["DataColumn148"] = (int)Data17_0.Select(t => t.iORCA_DATA72).Sum();
                   NewRow["DataColumn149"] = Data17_0.Select(t => t.iORCA_DATA73).Average();

                   //NewRow["DataColumn151"] = 

                   //先前測就好
                   int iOTHER9_1_41 = DataTEST_41.Where(t => t.cTM_FRA08 == "1").Count();
                   int iOTHER9_1_51 = DataTEST_51.Where(t => t.cTM_FRA09 == "1").Count();
                   int iOTHER9_1_71 = DataTEST_71.Where(t => t.cTM_FRA09 == "1").Count();

                   int iOTHER9_2_41 = DataTEST_41.Where(t => t.cTM_FRA08 == "2").Count();
                   int iOTHER9_2_51 = DataTEST_51.Where(t => t.cTM_FRA09 == "2").Count();
                   int iOTHER9_2_71 = DataTEST_71.Where(t => t.cTM_FRA09 == "2").Count();

                   int iOTHER9_3_41 = DataTEST_41.Where(t => t.cTM_FRA08 == "3").Count();
                   int iOTHER9_3_51 = DataTEST_51.Where(t => t.cTM_FRA09 == "3").Count();
                   int iOTHER9_3_71 = DataTEST_71.Where(t => t.cTM_FRA09 == "3").Count();

                   int iOTHER9_B = iOTHER9_1_41 + iOTHER9_2_41 + iOTHER9_3_41
                                 + iOTHER9_1_51 + iOTHER9_2_51 + iOTHER9_3_51;
                   int iOTHER9_C = iOTHER9_1_71 + iOTHER9_2_71 + iOTHER9_3_71;
                   NewRow["DataColumn241"] = (decimal)(iOTHER9_1_41 + iOTHER9_1_51) / (decimal)iOTHER9_B * 100;
                   NewRow["DataColumn242"] = (decimal)(iOTHER9_2_41 + iOTHER9_2_51) / (decimal)iOTHER9_B * 100;
                   NewRow["DataColumn243"] = (decimal)(iOTHER9_3_41 + iOTHER9_3_51) / (decimal)iOTHER9_B * 100;

                   NewRow["DataColumn246"] = (decimal)iOTHER9_1_71 / (decimal)iOTHER9_C * 100;
                   NewRow["DataColumn247"] = (decimal)iOTHER9_2_71 / (decimal)iOTHER9_C * 100;
                   NewRow["DataColumn248"] = (decimal)iOTHER9_3_71 / (decimal)iOTHER9_C * 100;

                   //先前測就好
                   int iOTHER10_1_41 = DataTEST_41.Where(t => t.cTM_FRA09 == "1").Count();
                   int iOTHER10_1_51 = DataTEST_51.Where(t => t.cTM_FRA10 == "1").Count();
                   int iOTHER10_1_71 = DataTEST_71.Where(t => t.cTM_FRA10 == "1").Count();

                   int iOTHER10_2_41 = DataTEST_41.Where(t => t.cTM_FRA09 == "2").Count();
                   int iOTHER10_2_51 = DataTEST_51.Where(t => t.cTM_FRA10 == "2").Count();
                   int iOTHER10_2_71 = DataTEST_71.Where(t => t.cTM_FRA10 == "2").Count();

                   int iOTHER10_3_41 = DataTEST_41.Where(t => t.cTM_FRA09 == "3").Count();
                   int iOTHER10_3_51 = DataTEST_51.Where(t => t.cTM_FRA10 == "3").Count();
                   int iOTHER10_3_71 = DataTEST_71.Where(t => t.cTM_FRA10 == "3").Count();

                   int iOTHER10_B = iOTHER10_1_41 + iOTHER10_2_41 + iOTHER10_3_41
                                 + iOTHER10_1_51 + iOTHER10_2_51 + iOTHER10_3_51;
                   int iOTHER10_C = iOTHER10_1_71 + iOTHER10_2_71 + iOTHER10_3_71;
                   NewRow["DataColumn251"] = (decimal)(iOTHER10_1_41 + iOTHER10_1_51) / (decimal)iOTHER10_B * 100;
                   NewRow["DataColumn252"] = (decimal)(iOTHER10_2_41 + iOTHER10_2_51) / (decimal)iOTHER10_B * 100;
                   NewRow["DataColumn253"] = (decimal)(iOTHER10_3_41 + iOTHER10_3_51) / (decimal)iOTHER10_B * 100;

                   NewRow["DataColumn256"] = (decimal)iOTHER10_1_71 / (decimal)iOTHER10_C * 100;
                   NewRow["DataColumn257"] = (decimal)iOTHER10_2_71 / (decimal)iOTHER10_C * 100;
                   NewRow["DataColumn258"] = (decimal)iOTHER10_3_71 / (decimal)iOTHER10_C * 100;

                   //先前測就好
                   int iOTHER11_1_41 = DataTEST_41.Where(t => t.cTM_FRA10 == "1").Count();
                   int iOTHER11_1_51 = DataTEST_51.Where(t => t.cTM_FRA11 == "1").Count();
                   int iOTHER11_1_71 = DataTEST_71.Where(t => t.cTM_FRA11 == "1").Count();

                   int iOTHER11_2_41 = DataTEST_41.Where(t => t.cTM_FRA10 == "2").Count();
                   int iOTHER11_2_51 = DataTEST_51.Where(t => t.cTM_FRA11 == "2").Count();
                   int iOTHER11_2_71 = DataTEST_71.Where(t => t.cTM_FRA11 == "2").Count();

                   int iOTHER3_B = iOTHER11_1_41 + iOTHER11_2_41
                                 + iOTHER11_1_51 + iOTHER11_2_51;
                   int iOTHER3_C = iOTHER11_1_71 + iOTHER11_2_71;
                   NewRow["DataColumn261"] = (decimal)(iOTHER11_1_41 + iOTHER11_1_51) / (decimal)iOTHER3_B * 100;
                   NewRow["DataColumn262"] = (decimal)(iOTHER11_2_41 + iOTHER11_2_51) / (decimal)iOTHER3_B * 100;

                   NewRow["DataColumn266"] = (decimal)iOTHER11_1_71 / (decimal)iOTHER3_C * 100;
                   NewRow["DataColumn267"] = (decimal)iOTHER11_2_71 / (decimal)iOTHER3_C * 100;

                   //先前測就好



                   dt_ANS.Rows.Add(NewRow);
               }

               DataTable dt_SUB1 = AYTools.GetRptTable();  //國小-初檢齲齒
               foreach (var Loop in dt1)
               {
                   var ORAL_R0 = ORAL_CAVITY_L.FirstOrDefault(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_FK == Loop.iSCH_PK);
                   var ORAL_R1 = ORAL_CAVITY_L.FirstOrDefault(t => t.iSCH_YEAR == (iSCH_YEAR - 1) && t.iSCH_FK == Loop.iSCH_PK);

                   if (ORAL_R1 == null || ORAL_R1?.iORCA_DATA2 == null)
                       continue;

                   DataRow NewRow = dt_SUB1.NewRow();

                   NewRow["DataColumn1"] = "新竹市";

                   NewRow["DataColumn11"] = Loop.iSCH_PK;
                   NewRow["DataColumn12"] = Loop.iSCH_LEVEL;
                   NewRow["DataColumn13"] = Loop.cSCH_CALLED;

                   decimal? iORCA_1 = ORAL_R1?.iORCA_DATA2;
                   decimal? iORCA_0 = ORAL_R0?.iORCA_DATA2;

                   NewRow["DataColumn14"] = iORCA_1;
                   NewRow["DataColumn15"] = iORCA_0;
                   NewRow["DataColumn16"] = iORCA_1 == null || iORCA_0 == null ? "" : (iORCA_1 - iORCA_0).ToString();
                   NewRow["DataColumn17"] = iORCA_1 == null || iORCA_0 == null ? "" :
                                           (iORCA_1 - iORCA_0) >= (decimal)0.5 ? "達成" : "未達成";

                   dt_SUB1.Rows.Add(NewRow);
               }

               DataTable dt_SUB2 = AYTools.GetRptTable();  //國中
               foreach (var Loop in dt2)
               {
                   var ORAL_R0 = ORAL_CAVITY_L.FirstOrDefault(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_FK == Loop.iSCH_PK);
                   var ORAL_R1 = ORAL_CAVITY_L.FirstOrDefault(t => t.iSCH_YEAR == (iSCH_YEAR - 1) && t.iSCH_FK == Loop.iSCH_PK);

                   if (ORAL_R1 == null || ORAL_R1?.iORCA_DATA2 == null)
                       continue;

                   DataRow NewRow = dt_SUB2.NewRow();

                   NewRow["DataColumn1"] = "新竹市";

                   NewRow["DataColumn11"] = Loop.iSCH_PK;
                   NewRow["DataColumn12"] = Loop.iSCH_LEVEL;
                   NewRow["DataColumn13"] = Loop.cSCH_CALLED;

                   decimal? iORCA_1 = ORAL_R1?.iORCA_DATA2;
                   decimal? iORCA_0 = ORAL_R0?.iORCA_DATA2;

                   NewRow["DataColumn14"] = iORCA_1;
                   NewRow["DataColumn15"] = iORCA_0;
                   NewRow["DataColumn16"] = iORCA_1 == null || iORCA_0 == null ? "" : (iORCA_1 - iORCA_0).ToString();
                   NewRow["DataColumn17"] = iORCA_1 == null || iORCA_0 == null ? "" :
                                           (iORCA_1 - iORCA_0) >= (decimal)0.5 ? "達成" : "未達成";

                   dt_SUB2.Rows.Add(NewRow);

               }

               DataTable dt_SUB3 = AYTools.GetRptTable();  //國小-矯治率
               foreach (var Loop in dt1)
               {
                   var ORAL_R0 = ORAL_CAVITY_L.FirstOrDefault(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_FK == Loop.iSCH_PK);

                   DataRow NewRow = dt_SUB3.NewRow();

                   NewRow["DataColumn1"] = "新竹市";

                   NewRow["DataColumn11"] = Loop.iSCH_PK;
                   NewRow["DataColumn12"] = Loop.iSCH_LEVEL;
                   NewRow["DataColumn13"] = Loop.cSCH_CALLED;

                   decimal? iORCA_0 = ORAL_R0?.iORCA_DATA3;

                   NewRow["DataColumn15"] = iORCA_0;

                   dt_SUB3.Rows.Add(NewRow);

               }

               DataTable dt_SUB4 = AYTools.GetRptTable();  //國中
               foreach (var Loop in dt2)
               {
                   var ORAL_R0 = ORAL_CAVITY_L.FirstOrDefault(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_FK == Loop.iSCH_PK);

                   DataRow NewRow = dt_SUB4.NewRow();

                   NewRow["DataColumn1"] = "新竹市";

                   NewRow["DataColumn11"] = Loop.iSCH_PK;
                   NewRow["DataColumn12"] = Loop.iSCH_LEVEL;
                   NewRow["DataColumn13"] = Loop.cSCH_CALLED;

                   decimal? iORCA_0 = ORAL_R0?.iORCA_DATA31;

                   NewRow["DataColumn15"] = iORCA_0;

                   dt_SUB4.Rows.Add(NewRow);

               }

               DataTable dt_SUB5 = AYTools.GetRptTable();  //國中-DMFT
               foreach (var Loop in dt2)
               {
                   var ORAL_R0 = ORAL_CAVITY_L.FirstOrDefault(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_FK == Loop.iSCH_PK);
                   var ORAL_R1 = ORAL_CAVITY_L.FirstOrDefault(t => t.iSCH_YEAR == iSCH_YEAR1 && t.iSCH_FK == Loop.iSCH_PK);

                   DataRow NewRow = dt_SUB5.NewRow();

                   NewRow["DataColumn1"] = "新竹市";

                   NewRow["DataColumn11"] = Loop.iSCH_PK;
                   NewRow["DataColumn12"] = Loop.iSCH_LEVEL;
                   NewRow["DataColumn13"] = Loop.cSCH_CALLED;

                   decimal? iORCA_1 = ORAL_R1?.iORCA_DATA71;
                   decimal? iORCA_0 = ORAL_R0?.iORCA_DATA71;

                   NewRow["DataColumn14"] = iORCA_1;
                   NewRow["DataColumn15"] = iORCA_0;

                   dt_SUB5.Rows.Add(NewRow);

               }

               DataTable dt_SUB6 = AYTools.GetRptTable();  //國中小-含氟牙膏
               foreach (var Loop in DataTEST_ToSchool)
               {
                   var SCHOOLS_R = db.SCHOOLS.FirstOrDefault(t => t.iSCH_PK == Loop.iSCH_FK);

                   DataRow NewRow = dt_SUB6.NewRow();

                   NewRow["DataColumn1"] = "新竹市";

                   NewRow["DataColumn11"] = SCHOOLS_R.iSCH_PK;
                   NewRow["DataColumn12"] = SCHOOLS_R.iSCH_LEVEL;
                   NewRow["DataColumn13"] = SCHOOLS_R.cSCH_CALLED;

                   NewRow["DataColumn15"] = Loop.TM_GROUP02;

                   dt_SUB6.Rows.Add(NewRow);

               }

               #region dt_SUB7 //照片
               string sPicType1 = Pic_Type.學校衛生政策.GetDescription();
               string sPicType2 = Pic_Type.學校物質環境.GetDescription();
               string sPicType3 = Pic_Type.學校社會環境.GetDescription();
               string sPicType4 = Pic_Type.健康技能與教學.GetDescription();
               string sPicType5 = Pic_Type.學校健康服務.GetDescription();
               string sPicType6 = Pic_Type.社區關係.GetDescription();
               string sPicType99 = Pic_Type.其他.GetDescription();
               var dt_SUB7 = (from t1 in db.ORAL_CAVITY
                              from t2 in db.SCHOOLS
                              where t1.iSCH_YEAR == iSCH_YEAR
                                 && t1.bORCA_CANCEL == false
                                 && t1.iSCH_FK == t2.iSCH_PK
                                 && t2.bSCH_STATUS == true
                                 && (t2.iSCH_LEVEL == 1 || t2.iSCH_LEVEL == 2)
                                 && (t1.cOCRA_PHOTO1_ACT1 != null && t1.cOCRA_PHOTO1_ACT1 != "")
                              orderby t2.iSCH_LEVEL, t1.iSCH_FK
                              select new
                              {
                                  DataColumn1 = t1.iSCH_FK,
                                  DataColumn2 = t2.cSCH_CALLED,
                                  DataColumn3 = (t1.cOCRA_PHOTO1_TYPE1 == ("" + (byte)Pic_Type.學校衛生政策) ? sPicType1 :
                                                 t1.cOCRA_PHOTO1_TYPE1 == ("" + (byte)Pic_Type.學校物質環境) ? sPicType2 :
                                                 t1.cOCRA_PHOTO1_TYPE1 == ("" + (byte)Pic_Type.學校社會環境) ? sPicType3 :
                                                 t1.cOCRA_PHOTO1_TYPE1 == ("" + (byte)Pic_Type.健康技能與教學) ? sPicType4 :
                                                 t1.cOCRA_PHOTO1_TYPE1 == ("" + (byte)Pic_Type.學校健康服務) ? sPicType5 :
                                                 t1.cOCRA_PHOTO1_TYPE1 == ("" + (byte)Pic_Type.社區關係) ? sPicType6 :
                                                 t1.cOCRA_PHOTO1_TYPE1 == ("" + (byte)Pic_Type.其他) ? sPicType99 :
                                                 "") + t1.cOCRA_PHOTO1_ACT1,
                                  DataColumn4 = sDNS + (from t3 in db.FILES
                                                        from t4 in db.FILE_SUB
                                                        where t3.iFILE_CLASS_FK == t1.iORCA_PK
                                                        && t3.cFILE_CLASS == "OCRA_PHOTO1_ACT1"
                                                        && t3.bFILE_CANCEL == false
                                                        && t4.iFILE_FK == t3.iFILE_PK
                                                        && t4.bFSUB_CANCEL == false
                                                        orderby t4.dFSUB_DATE_CRE descending
                                                        select t4.cFSUB_URL).FirstOrDefault(),
                                  DataColumn5 = (t1.cOCRA_PHOTO1_TYPE2 == ("" + (byte)Pic_Type.學校衛生政策) ? sPicType1 :
                                                 t1.cOCRA_PHOTO1_TYPE2 == ("" + (byte)Pic_Type.學校物質環境) ? sPicType2 :
                                                 t1.cOCRA_PHOTO1_TYPE2 == ("" + (byte)Pic_Type.學校社會環境) ? sPicType3 :
                                                 t1.cOCRA_PHOTO1_TYPE2 == ("" + (byte)Pic_Type.健康技能與教學) ? sPicType4 :
                                                 t1.cOCRA_PHOTO1_TYPE2 == ("" + (byte)Pic_Type.學校健康服務) ? sPicType5 :
                                                 t1.cOCRA_PHOTO1_TYPE2 == ("" + (byte)Pic_Type.社區關係) ? sPicType6 :
                                                 t1.cOCRA_PHOTO1_TYPE2 == ("" + (byte)Pic_Type.其他) ? sPicType99 :
                                                 "") + t1.cOCRA_PHOTO1_ACT2,
                                  DataColumn6 = sDNS + (from t3 in db.FILES
                                                        from t4 in db.FILE_SUB
                                                        where t3.iFILE_CLASS_FK == t1.iORCA_PK
                                                        && t3.cFILE_CLASS == "OCRA_PHOTO1_ACT2"
                                                        && t3.bFILE_CANCEL == false
                                                        && t4.iFILE_FK == t3.iFILE_PK
                                                        && t4.bFSUB_CANCEL == false
                                                        orderby t4.dFSUB_DATE_CRE descending
                                                        select t4.cFSUB_URL).FirstOrDefault(),
                                  DataColumn7 = (t1.cOCRA_PHOTO1_TYPE3 == ("" + (byte)Pic_Type.學校衛生政策) ? sPicType1 :
                                                 t1.cOCRA_PHOTO1_TYPE3 == ("" + (byte)Pic_Type.學校物質環境) ? sPicType2 :
                                                 t1.cOCRA_PHOTO1_TYPE3 == ("" + (byte)Pic_Type.學校社會環境) ? sPicType3 :
                                                 t1.cOCRA_PHOTO1_TYPE3 == ("" + (byte)Pic_Type.健康技能與教學) ? sPicType4 :
                                                 t1.cOCRA_PHOTO1_TYPE3 == ("" + (byte)Pic_Type.學校健康服務) ? sPicType5 :
                                                 t1.cOCRA_PHOTO1_TYPE3 == ("" + (byte)Pic_Type.社區關係) ? sPicType6 :
                                                 t1.cOCRA_PHOTO1_TYPE3 == ("" + (byte)Pic_Type.其他) ? sPicType99 :
                                                 "") + t1.cOCRA_PHOTO1_ACT3,
                                  DataColumn8 = sDNS + (from t3 in db.FILES
                                                        from t4 in db.FILE_SUB
                                                        where t3.iFILE_CLASS_FK == t1.iORCA_PK
                                                        && t3.cFILE_CLASS == "OCRA_PHOTO1_ACT3"
                                                        && t3.bFILE_CANCEL == false
                                                        && t4.iFILE_FK == t3.iFILE_PK
                                                        && t4.bFSUB_CANCEL == false
                                                        orderby t4.dFSUB_DATE_CRE descending
                                                        select t4.cFSUB_URL).FirstOrDefault(),
                                  DataColumn9 = (t1.cOCRA_PHOTO1_TYPE4 == ("" + (byte)Pic_Type.學校衛生政策) ? sPicType1 :
                                                 t1.cOCRA_PHOTO1_TYPE4 == ("" + (byte)Pic_Type.學校物質環境) ? sPicType2 :
                                                 t1.cOCRA_PHOTO1_TYPE4 == ("" + (byte)Pic_Type.學校社會環境) ? sPicType3 :
                                                 t1.cOCRA_PHOTO1_TYPE4 == ("" + (byte)Pic_Type.健康技能與教學) ? sPicType4 :
                                                 t1.cOCRA_PHOTO1_TYPE4 == ("" + (byte)Pic_Type.學校健康服務) ? sPicType5 :
                                                 t1.cOCRA_PHOTO1_TYPE4 == ("" + (byte)Pic_Type.社區關係) ? sPicType6 :
                                                 t1.cOCRA_PHOTO1_TYPE4 == ("" + (byte)Pic_Type.其他) ? sPicType99 :
                                                 "") + t1.cOCRA_PHOTO1_ACT4,
                                  DataColumn10 = sDNS + (from t3 in db.FILES
                                                         from t4 in db.FILE_SUB
                                                         where t3.iFILE_CLASS_FK == t1.iORCA_PK
                                                         && t3.cFILE_CLASS == "OCRA_PHOTO1_ACT4"
                                                         && t3.bFILE_CANCEL == false
                                                         && t4.iFILE_FK == t3.iFILE_PK
                                                         && t4.bFSUB_CANCEL == false
                                                         orderby t4.dFSUB_DATE_CRE descending
                                                         select t4.cFSUB_URL).FirstOrDefault(),
                                  DataColumn11 = (t1.cOCRA_PHOTO1_TYPE5 == ("" + (byte)Pic_Type.學校衛生政策) ? sPicType1 :
                                                  t1.cOCRA_PHOTO1_TYPE5 == ("" + (byte)Pic_Type.學校物質環境) ? sPicType2 :
                                                  t1.cOCRA_PHOTO1_TYPE5 == ("" + (byte)Pic_Type.學校社會環境) ? sPicType3 :
                                                  t1.cOCRA_PHOTO1_TYPE5 == ("" + (byte)Pic_Type.健康技能與教學) ? sPicType4 :
                                                  t1.cOCRA_PHOTO1_TYPE5 == ("" + (byte)Pic_Type.學校健康服務) ? sPicType5 :
                                                  t1.cOCRA_PHOTO1_TYPE5 == ("" + (byte)Pic_Type.社區關係) ? sPicType6 :
                                                  t1.cOCRA_PHOTO1_TYPE5 == ("" + (byte)Pic_Type.其他) ? sPicType99 :
                                                  "") + t1.cOCRA_PHOTO1_ACT5,
                                  DataColumn12 = sDNS + (from t3 in db.FILES
                                                         from t4 in db.FILE_SUB
                                                         where t3.iFILE_CLASS_FK == t1.iORCA_PK
                                                         && t3.cFILE_CLASS == "OCRA_PHOTO1_ACT5"
                                                         && t3.bFILE_CANCEL == false
                                                         && t4.iFILE_FK == t3.iFILE_PK
                                                         && t4.bFSUB_CANCEL == false
                                                         orderby t4.dFSUB_DATE_CRE descending
                                                         select t4.cFSUB_URL).FirstOrDefault(),
                                  DataColumn13 = (t1.cOCRA_PHOTO2_TYPE1 == ("" + (byte)Pic_Type.學校衛生政策) ? sPicType1 :
                                                  t1.cOCRA_PHOTO2_TYPE1 == ("" + (byte)Pic_Type.學校物質環境) ? sPicType2 :
                                                  t1.cOCRA_PHOTO2_TYPE1 == ("" + (byte)Pic_Type.學校社會環境) ? sPicType3 :
                                                  t1.cOCRA_PHOTO2_TYPE1 == ("" + (byte)Pic_Type.健康技能與教學) ? sPicType4 :
                                                  t1.cOCRA_PHOTO2_TYPE1 == ("" + (byte)Pic_Type.學校健康服務) ? sPicType5 :
                                                  t1.cOCRA_PHOTO2_TYPE1 == ("" + (byte)Pic_Type.社區關係) ? sPicType6 :
                                                  t1.cOCRA_PHOTO2_TYPE1 == ("" + (byte)Pic_Type.其他) ? sPicType99 :
                                                  "") + t1.cOCRA_PHOTO2_ACT1,
                                  DataColumn14 = sDNS + (from t3 in db.FILES
                                                         from t4 in db.FILE_SUB
                                                         where t3.iFILE_CLASS_FK == t1.iORCA_PK
                                                         && t3.cFILE_CLASS == "OCRA_PHOTO2_ACT1"
                                                         && t3.bFILE_CANCEL == false
                                                         && t4.iFILE_FK == t3.iFILE_PK
                                                         && t4.bFSUB_CANCEL == false
                                                         orderby t4.dFSUB_DATE_CRE descending
                                                         select t4.cFSUB_URL).FirstOrDefault(),
                                  DataColumn15 = (t1.cOCRA_PHOTO2_TYPE2 == ("" + (byte)Pic_Type.學校衛生政策) ? sPicType1 :
                                                  t1.cOCRA_PHOTO2_TYPE2 == ("" + (byte)Pic_Type.學校物質環境) ? sPicType2 :
                                                  t1.cOCRA_PHOTO2_TYPE2 == ("" + (byte)Pic_Type.學校社會環境) ? sPicType3 :
                                                  t1.cOCRA_PHOTO2_TYPE2 == ("" + (byte)Pic_Type.健康技能與教學) ? sPicType4 :
                                                  t1.cOCRA_PHOTO2_TYPE2 == ("" + (byte)Pic_Type.學校健康服務) ? sPicType5 :
                                                  t1.cOCRA_PHOTO2_TYPE2 == ("" + (byte)Pic_Type.社區關係) ? sPicType6 :
                                                  t1.cOCRA_PHOTO2_TYPE2 == ("" + (byte)Pic_Type.其他) ? sPicType99 :
                                                  "") + t1.cOCRA_PHOTO2_ACT2,
                                  DataColumn16 = sDNS + (from t3 in db.FILES
                                                         from t4 in db.FILE_SUB
                                                         where t3.iFILE_CLASS_FK == t1.iORCA_PK
                                                         && t3.cFILE_CLASS == "OCRA_PHOTO2_ACT2"
                                                         && t3.bFILE_CANCEL == false
                                                         && t4.iFILE_FK == t3.iFILE_PK
                                                         && t4.bFSUB_CANCEL == false
                                                         orderby t4.dFSUB_DATE_CRE descending
                                                         select t4.cFSUB_URL).FirstOrDefault(),
                                  DataColumn17 = (t1.cOCRA_PHOTO2_TYPE3 == ("" + (byte)Pic_Type.學校衛生政策) ? sPicType1 :
                                                  t1.cOCRA_PHOTO2_TYPE3 == ("" + (byte)Pic_Type.學校物質環境) ? sPicType2 :
                                                  t1.cOCRA_PHOTO2_TYPE3 == ("" + (byte)Pic_Type.學校社會環境) ? sPicType3 :
                                                  t1.cOCRA_PHOTO2_TYPE3 == ("" + (byte)Pic_Type.健康技能與教學) ? sPicType4 :
                                                  t1.cOCRA_PHOTO2_TYPE3 == ("" + (byte)Pic_Type.學校健康服務) ? sPicType5 :
                                                  t1.cOCRA_PHOTO2_TYPE3 == ("" + (byte)Pic_Type.社區關係) ? sPicType6 :
                                                  t1.cOCRA_PHOTO2_TYPE3 == ("" + (byte)Pic_Type.其他) ? sPicType99 :
                                                  "") + t1.cOCRA_PHOTO2_ACT3,
                                  DataColumn18 = sDNS + (from t3 in db.FILES
                                                         from t4 in db.FILE_SUB
                                                         where t3.iFILE_CLASS_FK == t1.iORCA_PK
                                                         && t3.cFILE_CLASS == "OCRA_PHOTO2_ACT3"
                                                         && t3.bFILE_CANCEL == false
                                                         && t4.iFILE_FK == t3.iFILE_PK
                                                         && t4.bFSUB_CANCEL == false
                                                         orderby t4.dFSUB_DATE_CRE descending
                                                         select t4.cFSUB_URL).FirstOrDefault(),

                                  DataColumn100 = sDNS,//判斷是否有圖片用
                              }).ToList();
               #endregion

               DataTable dt_SUB8 = AYTools.GetRptTable();  //第一大臼齒窩溝封填-一、二年級
               var DataSub8 = ORAL_CAVITY_L.Where(t => t.iSCH_YEAR == iSCH_YEAR && t.iSCH_LEVEL == 1);
               int[] GRADE_NUM = new int[] { 1, 2 };
               string[] GRADE_STR = new string[] { "一年級", "二年級" };
               int[] DO_SUM = new int[] { (int)DataSub8.Select(t => t.iORCA_DATA61).Sum()
                                         ,(int)DataSub8.Select(t => t.iORCA_DATA71).Sum()};
               int[] ALL_SUM = new int[] { (int)DataSub8.Select(t => t.iORCA_DATA62).Sum()
                                          ,(int)DataSub8.Select(t => t.iORCA_DATA72).Sum()};
               decimal?[] AVG = new decimal?[] { DataSub8.Select(t => t.iORCA_DATA63).Average()
                                                ,DataSub8.Select(t => t.iORCA_DATA73).Average()};
               int iCountSub8 = 0;
               foreach (var Loop in GRADE_NUM)
               {
                   DataRow NewRow = dt_SUB8.NewRow();

                   NewRow["DataColumn1"] = GRADE_NUM[iCountSub8];
                   NewRow["DataColumn6"] = GRADE_STR[iCountSub8];

                   NewRow["DataColumn11"] = DO_SUM[iCountSub8];
                   NewRow["DataColumn12"] = ALL_SUM[iCountSub8];
                   NewRow["DataColumn13"] = AVG[iCountSub8];

                   iCountSub8++;
                   dt_SUB8.Rows.Add(NewRow);
               }

               AYRptWrapper rw = new AYRptWrapper();
               rw.Parameters.Add(new ReportParameter("cGOV_NAME", CurrUser.GOV_NAME));
               rw.Parameters.Add(new ReportParameter("cRPT_NAME", "口腔保健實施成果暨分析探究"));
               rw.ReportPath = string.Format("{0}/{1}.rdlc", AYConst.CTRL_REPORT_CENTER, "RPT_3011");
               rw.DataSources.Add(new ReportDataSource("DataSet1", dt_ANS));
               rw.DataSources.Add(new ReportDataSource("DataSet9", dt_SUB8));
               rw.DataSourcesSub.Add(new ReportDataSource("DataSet2", dt_SUB1));
               rw.DataSourcesSub.Add(new ReportDataSource("DataSet3", dt_SUB2));
               rw.DataSourcesSub.Add(new ReportDataSource("DataSet4", dt_SUB3));
               rw.DataSourcesSub.Add(new ReportDataSource("DataSet5", dt_SUB4));
               rw.DataSourcesSub.Add(new ReportDataSource("DataSet6", dt_SUB5));
               rw.DataSourcesSub.Add(new ReportDataSource("DataSet7", dt_SUB6));
               rw.DataSourcesSub.Add(new ReportDataSource("DataSet8", dt_SUB7));

               Session[AYConst.AYREPORT_SESSION] = rw;
               return RedirectToAction("RptView", AYConst.CTRL_REPORT);
           }
           catch (Exception ex)
           {
               CatchEvent(ex);
               return this.Json(jResult.Data);
           }
        }
