using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace ImportExcel
{
    class Program
    {
        static void Main(string[] args)
        {

            string path = Path.Combine(Environment.CurrentDirectory, @"staff\");

            String[] files= System.IO.Directory.GetFiles(path,"*.xlsx");

            foreach (var file in files)
            {
                FileInfo existingFile = new FileInfo(file);
                string filename = GetFileName(file);

                //Boolean isExist = existingFile.Exists;

                using (var package = new ExcelPackage(existingFile))
                {
                    foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
                    {
                        if (worksheet.Index == 0)  // Only for first worksheet
                        {
                            for (int i = worksheet.Dimension.Start.Row+1; i <= worksheet.Dimension.End.Row; i++)
                            {
                                List<SqlParameter> paras = new List<SqlParameter>();

                                SqlParameter p = new SqlParameter("@Filename", SqlDbType.NVarChar, 32);
                                p.Value = filename;
                                paras.Add(p);

                                SqlParameter pline = new SqlParameter("@LineID", SqlDbType.Int, 16);
                                pline.Value = i;
                                paras.Add(pline);


                                for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                                {

                                    SqlParameter para = GenerateParam(worksheet.Cells[i, j].Value == null? String.Empty: worksheet.Cells[i, j].Value.ToString(),j);
                                    paras.Add(para);

                                } 

                                insertIntoDatabase(paras.ToArray());
                            }
                        }
                    }
                }

                Console.WriteLine(filename + " finished !");
            }

            Console.ReadKey();

        }

        public static SqlParameter GenerateParam(string cellvalue, int columnNo)
        {
            if (columnNo == 1)
            {
                SqlParameter p1 = new SqlParameter("@ID_ORG", SqlDbType.NVarChar, 32);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 2)
            {
                SqlParameter p1 = new SqlParameter("@RECORD_DEFUNCT_RISK", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 3)
            {
                SqlParameter p1 = new SqlParameter("@BUS_DESCRIPTION_1", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 4)
            {
                SqlParameter p1 = new SqlParameter("@BUSCODE_1", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 5)
            {
                SqlParameter p1 = new SqlParameter("@ORGANISATION", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 6)
            {
                SqlParameter p1 = new SqlParameter("@ADDRESS", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 7)
            {
                SqlParameter p1 = new SqlParameter("@LOCATION", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 8)
            {
                SqlParameter p1 = new SqlParameter("@STATE", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 9)
            {
                SqlParameter p1 = new SqlParameter("@POSTCODE", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 10)
            {
                SqlParameter p1 = new SqlParameter("@REGION", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 11)
            {
                SqlParameter p1 = new SqlParameter("@HEAD_OFFICE", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            if (columnNo == 12)
            {
                SqlParameter p1 = new SqlParameter("@PHONE1", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 13)
            {
                SqlParameter p1 = new SqlParameter("@PHONE2", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 14)
            {
                SqlParameter p1 = new SqlParameter("@TOLLFREE1", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 15)
            {
                SqlParameter p1 = new SqlParameter("@TOLLFREE2", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 16)
            {
                SqlParameter p1 = new SqlParameter("@MOBILE1", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 17)
            {
                SqlParameter p1 = new SqlParameter("@MOBILE2", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 18)
            {
                SqlParameter p1 = new SqlParameter("@FAX", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 19)
            {
                SqlParameter p1 = new SqlParameter("@EMAIL1", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 20)
            {
                SqlParameter p1 = new SqlParameter("@EMAIL2", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 21)
            {
                SqlParameter p1 = new SqlParameter("@WEBSITE", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 22)
            {
                SqlParameter p1 = new SqlParameter("@TWITTER_LINK", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 23)
            {
                SqlParameter p1 = new SqlParameter("@FACEBOOK_LINK", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 24)
            {
                SqlParameter p1 = new SqlParameter("@LINKEDIN", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 25)
            {
                SqlParameter p1 = new SqlParameter("@NUMOFEMP", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 26)
            {
                SqlParameter p1 = new SqlParameter("@ABN", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 27)
            {
                SqlParameter p1 = new SqlParameter("@ENTITY_TYPE", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 28)
            {
                SqlParameter p1 = new SqlParameter("@ABN_STATUS", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 29)
            {
                SqlParameter p1 = new SqlParameter("@ABN_STATUS_DATE", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 30)
            {
                SqlParameter p1 = new SqlParameter("@YEAR_ESTABLISHED", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 31)
            {
                SqlParameter p1 = new SqlParameter("@CONTACT1_FULL_NAME", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 32)
            {
                SqlParameter p1 = new SqlParameter("@CONTACT1_JOB_DESCRIPTION", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 33)
            {
                SqlParameter p1 = new SqlParameter("@CONTACT2_FULL_NAME", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 34)
            {
                SqlParameter p1 = new SqlParameter("@CONTACT2_JOB_DESCRIPTION", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 35)
            {
                SqlParameter p1 = new SqlParameter("@ANZSIC_CODE_1", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 36)
            {
                SqlParameter p1 = new SqlParameter("@ANZSIC_DESCRIPTION_1", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 37)
            {
                SqlParameter p1 = new SqlParameter("@BUSCODE_2", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 38)
            {
                SqlParameter p1 = new SqlParameter("@BUS_DESCRIPTION_2", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 39)
            {
                SqlParameter p1 = new SqlParameter("@ANZSIC_CODE_2", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 40)
            {
                SqlParameter p1 = new SqlParameter("@ANZSIC_DESCRIPTION_2", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 41)
            {
                SqlParameter p1 = new SqlParameter("@LATITUDE", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 42)
            {
                SqlParameter p1 = new SqlParameter("@LONGITUDE", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 43)
            {
                SqlParameter p1 = new SqlParameter("@MAPLINK", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 44)
            {
                SqlParameter p1 = new SqlParameter("@ID_ORG2018", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else if (columnNo == 45)
            {
                SqlParameter p1 = new SqlParameter("@RECORD_DATE", SqlDbType.NVarChar, 64);
                p1.Value = cellvalue;
                return p1;
            }
            else
            {
                return null;
            }


        }


        public static string GetFileName(string path)
        {
            int len = path.Length;

            int index = path.LastIndexOf(@"\");

            return path.Substring(index+1);


        }

        public static void insertIntoDatabase(params SqlParameter[] paras)
        {
            String connectionString = @"server=netcrmau;uid=dev;pwd='';database=Backup";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                                        insert into [Backup].[dbo].[ContactStates]
                                        (
                                              [Filename] 
                                              ,[LineID]
                                              ,[ID_ORG]
                                              ,[RECORD_DEFUNCT_RISK]
                                              ,[BUS_DESCRIPTION_1]
                                              ,[BUSCODE_1]
                                              ,[ORGANISATION]
                                              ,[ADDRESS]
                                              ,[LOCATION]
                                              ,[STATE]
                                              ,[POSTCODE]
                                              ,[REGION]
                                              ,[HEAD_OFFICE/BRANCH]
                                              ,[PHONE1]
                                              ,[PHONE2]
                                              ,[TOLLFREE1]
                                              ,[TOLLFREE2]
                                              ,[MOBILE1]
                                              ,[MOBILE2]
                                              ,[FAX]
                                              ,[EMAIL1]
                                              ,[EMAIL2]
                                              ,[WEBSITE]
                                              ,[TWITTER_LINK]
                                              ,[FACEBOOK_LINK]
                                              ,[LINKEDIN]
                                              ,[NUMOFEMP]
                                              ,[ABN]
                                              ,[ENTITY_TYPE]
                                              ,[ABN_STATUS]
                                              ,[ABN_STATUS_DATE]
                                              ,[YEAR_ESTABLISHED]
                                              ,[CONTACT1_FULL_NAME]
                                              ,[CONTACT1_JOB_DESCRIPTION]
                                              ,[CONTACT2_FULL_NAME]
                                              ,[CONTACT2_JOB_DESCRIPTION]
                                              ,[ANZSIC_CODE_1]
                                              ,[ANZSIC_DESCRIPTION_1]
                                              ,[BUSCODE_2]
                                              ,[BUS_DESCRIPTION_2]
                                              ,[ANZSIC_CODE_2]
                                              ,[ANZSIC_DESCRIPTION_2]
                                              ,[LATITUDE]
                                              ,[LONGITUDE]
                                              ,[MAPLINK]
                                              ,[ID_ORG2018]
                                              ,[RECORD_DATE] 
                                        ) values(
                                                      @Filename
                                                      ,@LineID
                                                      ,@ID_ORG
                                                      ,@RECORD_DEFUNCT_RISK
                                                      ,@BUS_DESCRIPTION_1
                                                      ,@BUSCODE_1
                                                      ,@ORGANISATION
                                                      ,@ADDRESS
                                                      ,@LOCATION
                                                      ,@STATE
                                                      ,@POSTCODE
                                                      ,@REGION
                                                      ,@HEAD_OFFICE
                                                      ,@PHONE1
                                                      ,@PHONE2
                                                      ,@TOLLFREE1
                                                      ,@TOLLFREE2
                                                      ,@MOBILE1
                                                      ,@MOBILE2
                                                      ,@FAX
                                                      ,@EMAIL1
                                                      ,@EMAIL2
                                                      ,@WEBSITE
                                                      ,@TWITTER_LINK
                                                      ,@FACEBOOK_LINK
                                                      ,@LINKEDIN
                                                      ,@NUMOFEMP
                                                      ,@ABN
                                                      ,@ENTITY_TYPE
                                                      ,@ABN_STATUS
                                                      ,@ABN_STATUS_DATE
                                                      ,@YEAR_ESTABLISHED
                                                      ,@CONTACT1_FULL_NAME
                                                      ,@CONTACT1_JOB_DESCRIPTION
                                                      ,@CONTACT2_FULL_NAME
                                                      ,@CONTACT2_JOB_DESCRIPTION
                                                      ,@ANZSIC_CODE_1
                                                      ,@ANZSIC_DESCRIPTION_1
                                                      ,@BUSCODE_2
                                                      ,@BUS_DESCRIPTION_2
                                                      ,@ANZSIC_CODE_2
                                                      ,@ANZSIC_DESCRIPTION_2
                                                      ,@LATITUDE
                                                      ,@LONGITUDE
                                                      ,@MAPLINK
                                                      ,@ID_ORG2018
                                                      ,@RECORD_DATE
                                            )";
                    cmd.Parameters.AddRange(paras);
                    conn.Open();
                    SqlTransaction trans = conn.BeginTransaction();
                    //int rows = cmd.ExecuteNonQuery();
                    //Console.WriteLine(rows);
                    try
                    {
                        cmd.Transaction = trans;
                        cmd.ExecuteNonQuery();
                        trans.Commit();
                    }
                    catch (Exception ex)
                    {

                        trans.Rollback();
                        using (System.IO.StreamWriter file =
                            new System.IO.StreamWriter(@"C:\Users\adrian_sun\Desktop\staff\test.txt", true))
                        {
                            file.WriteLine(paras[0].Value.ToString() +"   " + paras[1].Value.ToString());
                        }
                        Console.WriteLine(paras[0].Value.ToString() + "   " + paras[1].Value.ToString());
                        throw;
                    }



                }
            }






        }








    }
}
