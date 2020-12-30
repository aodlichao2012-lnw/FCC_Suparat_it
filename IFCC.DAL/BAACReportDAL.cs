using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;


namespace IFCC.DAL
{
   public class BAACReportDAL : BaseDAL
    {

        #region + Instance +
        private static BAACReportDAL _instance;

        public static BAACReportDAL Instance
        {
            get
            {
                _instance = new BAACReportDAL();
                return _instance;
            }

        }
        #endregion

        #region GetBasketReport
        public DataTable GetBasketReport(string startdate,string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[AACCIVR].[dbo].[sp_Basket_Report_BAAC]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "BasketSelect");

                return dt;
            } catch (Exception ex) 
            { 
                throw ex; 
            }
        }
        #endregion

        #region GetHotelReport
        public DataTable GetHotelReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[AACCIVR].[dbo].[sp_Hotels_Report_BAAC]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "HotelSelect");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region GetHealthReport
        public DataTable GetHealthReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[AACCIVR].[dbo].[sp_Health_Report_BAAC]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "HealthSelect");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        
        #region GetFoodReport
        public DataTable GetFoodReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[AACCIVR].[dbo].[sp_Food_Report_BAAC]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "FoodSelect");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion


        #region GetTravalReport
        public DataTable GetTravalReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[AACCIVR].[dbo].[sp_Traval_Report_BAAC]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "TravelSelect");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion



        #region GetSpecialExperienceReport
        public DataTable GetSpecialExperienceReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[AACCIVR].[dbo].[sp_Spacial_Report_BAAC]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "ExperienceSelect");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region GetDonetReport
        public DataTable GetDonetReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[AACCIVR].[dbo].[sp_Donet_Report_BAAC]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "DonationSelect");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region GetCarServiceReport
        public DataTable GetCarServiceReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[AACCIVR].[dbo].[sp_CarService_Report_BAAC]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "CarServiceSelect");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion



        #region GetSportServiceReport
        public DataTable GetSportServiceReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[AACCIVR].[dbo].[sp_SportService_Report_BAAC]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "CarServiceSelect");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion



        #region GetSkillCourseReport
        public DataTable GetSkillCourseReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[AACCIVR].[dbo].[sp_SkillCourse_Report_BAAC]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "CarServiceSelect");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion



        #region GetHousingReport
        public DataTable GetHousingReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[AACCIVR].[dbo].[sp_HousingService_Report_BAAC]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "CarServiceSelect");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        


        #region GetPetReport
        public DataTable GetPetReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[AACCIVR].[dbo].[sp_Pet_Report_BAAC]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "CarServiceSelect");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion


        #region GetChildReport
        public DataTable GetChildReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[AACCIVR].[dbo].[sp_Child_Report_BAAC]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "CarServiceSelect");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        //New
        #region GetRoadsideReport
       
        public DataTable GetRoadsideReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[iFCC].[dbo].[SP_O_BAAC_BAACReport]".Trim();
                cmd.CommandType = CommandType.StoredProcedure;
                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "DetailSelect");

                return dt;
            }
            catch(Exception e)
            {
                throw e;
            }
        } 
        public DataTable GetTbankReportt(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[iFCC].[dbo].[SP_O_TBANK_TBANKReport]";
                cmd.CommandType = CommandType.StoredProcedure;
                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "DetailSelect");

                return dt;
            }
            catch(Exception e)
            {
                throw e;
            }
        }
        
        public DataTable GetMTLReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[iFCC].[dbo].[SP_O_FCCDB_FCCDBReport]";
                cmd.CommandType = CommandType.StoredProcedure;
                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "DetailSelect");

                return dt;
            }
            catch(Exception e)
            {
                throw e;
            }
        }
        
        public DataTable GetKMTReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[iFCC].[dbo].[SP_O_Kmotor_KmotorReport]";
                cmd.CommandType = CommandType.StoredProcedure;
                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "DetailSelect");

                return dt;
            }
            catch(Exception e)
            {
                throw e;
            }
        }
        public DataTable GetCignaReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[iFCC].[dbo].[SP_O_Cigna_CignaReport]";
                cmd.CommandType = CommandType.StoredProcedure;
                #region Parameter 
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "DetailSelect");

                return dt;
            }
            catch(Exception e)
            {
                throw e;
            }
        }


        #endregion
    }
}
