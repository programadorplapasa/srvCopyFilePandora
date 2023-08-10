using Sap.Data.Hana;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace wSrvCopyFilePandora.Clases
{
    public class GlobalsSQL
    {
        public static string Server = string.Empty;
        public static string dbName = string.Empty;
        public static string dbUser = string.Empty;
        public static string dbPassword = string.Empty;
        public static string txtMessageBox = string.Empty;

        public static DataSet GetDataSet(string ConnectionString, string sQuery, int CommandTimeout)
        {
            DataSet dtsConsulta = new DataSet(); 

            try
            {
                using (SqlCommand cmd = new SqlCommand(sQuery, new SqlConnection(ConnectionString)))
                {
                    cmd.CommandTimeout = CommandTimeout;
                    cmd.Connection.Open();
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (!reader.IsClosed)
                        {
                            DataTable dt = new DataTable();
                            // DataTable.Load automatically advances the reader to the next result set
                            dt.Load(reader);
                            dtsConsulta.Tables.Add(dt);
                        }
                        reader.Close();
                    }
                    cmd.Connection.Close();
                    cmd.Connection.Dispose();
                    cmd.Dispose();
                }
                return dtsConsulta;
            }
            catch (Exception)
            {

                throw;
            }

            return dtsConsulta;

        }
        public static DataSet GetDataSet(string ConnectionString, string sQuery)
        {
            DataSet dtsConsulta = new DataSet();

            try
            {
                using (SqlCommand cmd = new SqlCommand(sQuery, new SqlConnection(ConnectionString)))
                {
                    cmd.CommandTimeout = 0;
                    cmd.Connection.Open();
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (!reader.IsClosed)
                        {
                            DataTable dt = new DataTable();
                            // DataTable.Load automatically advances the reader to the next result set
                            dt.Load(reader);
                            dtsConsulta.Tables.Add(dt);
                        }
                        reader.Close();
                    }
                    cmd.Connection.Close();
                    cmd.Connection.Dispose();
                    cmd.Dispose();
                }
                return dtsConsulta;
            }
            catch (Exception)
            {

                throw;
            }

            return dtsConsulta;

        }
        public static DataSet GetDataSetHana(string ConnectionString, string sQuery)
        {
            DataSet dtsConsulta = new DataSet();

            try
            {
                using (HanaCommand cmd = new HanaCommand(sQuery, new HanaConnection(ConnectionString)))
                {
                    cmd.CommandTimeout = 0;
                    cmd.Connection.Open();
                    using (HanaDataReader reader = cmd.ExecuteReader())
                    {
                        while (!reader.IsClosed)
                        {
                            DataTable dt = new DataTable();
                            // DataTable.Load automatically advances the reader to the next result set
                            dt.Load(reader);
                            dtsConsulta.Tables.Add(dt);
                        }
                        reader.Close();
                    }
                    cmd.Connection.Close();
                    cmd.Connection.Dispose();
                    cmd.Dispose();
                }
                return dtsConsulta;
            }
            catch (Exception)
            {

                throw;
            }

            return dtsConsulta;

        }
        public static DataTable GetDataTable(string ConnectionString, string sQuery, int CommandTimeout)
        {
            DataSet dtsConsulta = new DataSet();
            DataTable dtConsulta = new DataTable();

            try
            {
                using (SqlCommand cmd = new SqlCommand(sQuery, new SqlConnection(ConnectionString)))
                {
                    cmd.CommandTimeout = CommandTimeout;
                    cmd.Connection.Open();
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (!reader.IsClosed)
                        {
                            DataTable dt = new DataTable();
                            // DataTable.Load automatically advances the reader to the next result set
                            dt.Load(reader);
                            dtConsulta.Load(reader);
                            // dtsConsulta.Tables.Add(dt);
                        }
                        reader.Close();
                    }
                    cmd.Connection.Close();
                    cmd.Connection.Dispose();
                    cmd.Dispose();
                }
                return dtConsulta;
            }
            catch (Exception)
            {
                dtConsulta = new DataTable();
                return dtConsulta;
            }
        }
        public static DataTable GetDataTable(string ConnectionString, string sQuery)
        {
            DataSet dtsConsulta = new DataSet();
            DataTable dtConsulta = new DataTable();

            try
            {
                using (SqlCommand cmd = new SqlCommand(sQuery, new SqlConnection(ConnectionString)))
                {
                    cmd.CommandTimeout = 0;
                    cmd.Connection.Open();
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (!reader.IsClosed)
                        {
                            DataTable dt = new DataTable();
                            // DataTable.Load automatically advances the reader to the next result set
                            dt.Load(reader);
                            dtConsulta.Load(reader);
                            // dtsConsulta.Tables.Add(dt);
                        }
                        reader.Close();
                    }
                    cmd.Connection.Close();
                    cmd.Connection.Dispose();
                    cmd.Dispose();
                }
                return dtConsulta;
            }
            catch (Exception)
            {
                dtConsulta = new DataTable();
                return dtConsulta;
            }
        }

        public static DataSet GetDefincion(string CCiDefinicion, string CCiDetDefinicion, string DbName)
        {
            string ConnectionStringSQL = string.Empty;

            if (DbName == "Hermes") { ConnectionStringSQL = ConfigurationManager.ConnectionStrings["Hermes"].ConnectionString; }
            if (DbName == "Pandora") { ConnectionStringSQL = ConfigurationManager.ConnectionStrings["Pandora"].ConnectionString; }

            DataSet dtsConsulta = new DataSet();
            string sQuery = string.Empty;

            try
            {
                sQuery = string.Concat("select T1.CCiDetDefinicion, T1.CTxValor", Environment.NewLine);
                sQuery = string.Concat(sQuery, " from TblGeCabDefinicion T0", Environment.NewLine);
                sQuery = string.Concat(sQuery, " inner join TblGeDetDefinicion T1 ", Environment.NewLine);
                sQuery = string.Concat(sQuery, "    on T0.CCiDefinicion = T1.CCiDefinicion ", Environment.NewLine);
                sQuery = string.Concat(sQuery, " where 1=1 ", Environment.NewLine);

                if (CCiDefinicion != "") { sQuery = string.Concat(sQuery, " and T0.CCiDefinicion = '", CCiDefinicion, "'", Environment.NewLine); }
                if (CCiDetDefinicion != "") { sQuery = string.Concat(sQuery, " and T1.CCiDetDefinicion = '", CCiDetDefinicion, "'", Environment.NewLine); }

                dtsConsulta = GetDataSet(ConnectionStringSQL, sQuery);
                return dtsConsulta;

            }
            catch (Exception)
            {
                dtsConsulta = new DataSet();
                return dtsConsulta;
            }
        }
        public static DataSet GetDefincion(string CCiDefinicion, string DbName)
        {
            string ConnectionStringSQL = string.Empty;
            if (DbName == "Hermes") { ConnectionStringSQL = ConfigurationManager.ConnectionStrings["Hermes"].ConnectionString; }
            if (DbName == "Pandora") { ConnectionStringSQL = ConfigurationManager.ConnectionStrings["Pandora"].ConnectionString; }

            DataSet dtsConsulta = new DataSet();
            string sQuery = string.Empty;

            try
            {
                sQuery = string.Concat("select T1.CCiDetDefinicion, T1.CTxValor", Environment.NewLine);
                sQuery = string.Concat(sQuery, " from TblGeCabDefinicion T0", Environment.NewLine);
                sQuery = string.Concat(sQuery, " inner join TblGeDetDefinicion T1 ", Environment.NewLine);
                sQuery = string.Concat(sQuery, "    on T0.CCiDefinicion = T1.CCiDefinicion ", Environment.NewLine);
                sQuery = string.Concat(sQuery, " where 1=1 ", Environment.NewLine);
                sQuery = string.Concat(sQuery, " and T1.CCiDefinicion = '", CCiDefinicion, "'", Environment.NewLine);

                dtsConsulta = GetDataSet(ConnectionStringSQL, sQuery);
                return dtsConsulta;

            }
            catch (Exception)
            {
                dtsConsulta = new DataSet();
                return dtsConsulta;
            }
        }



    }
}
