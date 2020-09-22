using System;
using System.Data;
using System.Data.SqlClient;


namespace SqlDataServer
{
    public class Services : ReleaseObj
    {
        #region Vars
        private string Stringsql;
        private string VsystemName = "";
        private string Vmodule = "";
        private string Vmethod = "";
        private string Vuser = "";
        private System.Data.SqlClient.SqlConnection mConection;
        private System.Data.SqlClient.SqlTransaction mTransaction;
        public bool InTransaction;
        private string VTimeOut = "";
        #endregion
       
        #region Vars
        public string StringConection
        {
            get
            {
                if (Stringsql.Trim().Length == 0) 
                { 
                    throw new System.Exception("No se puede establecer la StringCommand de conexión");
                }
                return Stringsql;
            }
            set
            {
                Stringsql = value;
            }
        }
        public string SystemName
        {
            get
            {
                return VsystemName;
            }
            set
            {
                if (VsystemName.Trim().Length == 0)
                    VsystemName = value.Trim();
                else if (LastValue(VsystemName) != value)
                    VsystemName = VsystemName.Trim() + "." + value.Trim();
            }
        }
        public string ModuleName
        {
            get
            {
                return Vmodule;
            }
            set
            {
                if (Vmodule.Trim().Length == 0)
                    Vmodule = value.Trim();
                else if (LastValue(Vmodule) != value)
                    Vmodule = Vmodule.Trim() + "." + value.Trim();
            }
        }
        public string Method
        {
            get
            {
                return Vmethod;
            }
            set
            {
                if (Vmethod.Trim().Length == 0)
                    Vmethod = value.Trim();
                else if (LastValue(Vmethod) != value)
                    Vmethod = Vmethod.Trim() + "." + value.Trim();
            }
        }
        public string User
        {
            get
            {
                return Vuser;
            }
            set
            {
                if (Vuser.Trim().Length == 0)
                    Vuser = value.Trim();
                else if (Vuser.Trim() != value.Trim())
                    Vuser = Vuser.Trim() + "." + value.Trim();
            }
        }
        public string TimeOut
        {
            get
            {
                return VTimeOut;
            }
            set
            {
                VTimeOut = value.Trim();
            }
        }
        #endregion
       
        #region Constructors
        public Services(string StringConection)
        {
            Stringsql = StringConection;
        }
        #endregion

        #region Users Functions
        public int DeleteWithFilter(string Table, string Filter)
        {
            string sql;
            sql = Table + "_DX_" + Filter;
            return Execute(sql);
        }
    
        public int DeleteWithFilter(string Table, string Filter, int Args)
        {
            string sql;
            sql = Table + "_DX_" + Filter;
            return Execute(sql, Args);
        }
      
        public int DeleteWithFilter(string Table, string Filter, DataTable Args)
        {
            string sql;
            sql = Table + "_DX_" + Filter;
            return Execute(sql, Args);
        }
     
        public int DeleteWithFilter(string Table, string Filter, string[] Args)
        {
            string sql;
            sql = Table + "_DX_" + Filter;
            return Execute(sql, Args);
        }
     
        public int Delete(string Table, int Id)
        {
            return Execute(Table + "_E", Id);
        }
     
        public DataTable GetAll(string Table)
        {
            string Sql;
            Sql = Table + "_TT";
            return GetRecords(Sql);
        }
   
        public DataTable GetWithFilter(string Table, string Filter, string StringCommand)
        {
            string Sql;
            Sql = Table + "_TX_" + Filter;
            return GetRecords(Sql, StringCommand);
        }
   
        public DataTable GetWithFilter(string Table, string Filter, int Args)
        {
            string Sql;
            Sql = Table + "_TX_" + Filter;
            return GetRecords(Sql, Args);
        }
       
        public DataTable GetWithFilter(string Table, string Filter)
        {
            string Sql;
            Sql = Table + "_TX_" + Filter;
            return  GetRecords(Sql);
        }
      
        public DataTable GetWithFilter(string Table, string Filter, string[] Args)
        {
            string Sql;
            Sql = Table + "_TX_" + Filter;
            return GetRecords(Sql, Args);
        }
      
        public DataTable GetWithFilter(string Table, string Filter, DataTable Args)
        {
            string Sql;
            Sql = Table + "_TX_" + Filter;
            return GetRecords(Sql, Args);
        }
      
        public string GetValue(string Table, string Filter)
        {
            string Sql;
            Sql = Table + "_TV_" + Filter;
            return ExecuteValue(Sql).ToString();
        }
      
        public string GetValue(string Table, string Filter, int Args)
        {
            string Sql;
            Sql = Table + "_TV_" + Filter;
            return ExecuteValue(Sql, Args).ToString();
        }
     
        public string GetValue(string Table, string Filter, string StringCommand)
        {
            string Sql;
            Sql = Table + "_TV_" + Filter;
            return ExecuteValue(Sql, StringCommand).ToString();
        }
      
        public string GetValue(string Table, string Filter, string[] Args)
        {
            string Sql;
            Sql = Table + "_TV_" + Filter;
            return ExecuteValue(Sql, Args).ToString();
        }
      
        public string GetValue(string Table, string Filter, DataTable Args)
        {
            string Sql;
            Sql = Table + "_TV_" + Filter;
            return ExecuteValue(Sql, Args).ToString();
        }
    
        public DataTable GetOne(string Table, int Id)
        {
            string Sql;
            Sql = Table + "_T";
            return GetRecords(Sql, Id);
        }
      
        public int Add(string Table, DataRow Args)
        {
            string sql;
            sql = Table + "_A";
            return Execute(sql, Args);
        }
       
        public int Add(string Table, DataTable Args)
        {
            string sql;
            sql = Table + "_A";
            return Execute(sql, Args);
        }
     
        public int Add(string Table, string[] Args)
        {
            string sql;
            sql = Table + "_A";
            return Execute(sql, Args);
        }
     
        public int Update(string Table, string[] Args)
        {
            string sql;
            sql = Table + "_M";
            return Execute(sql, Args);
        }
    
        public int Update(string Table, DataRow Args)
        {
            string sql;
            sql = Table + "_M";
            return Execute(sql, Args);
        }
     
        public int Update(string Table, DataTable Args)
        {
            string sql;
            sql = Table + "_M";
            return Execute(sql, Args);
        }
      
        public int Update(string Table, string Filter)
        {
            string sql;
            sql = Table + "_M_" + Filter.Trim();
            return Execute(sql);
        }
      
        public int UpdateWithFilter(string Table, string Filter, int Id)
        {
            string sql;
            sql = Table + "_MX_" + Filter;
            return Execute(sql, Id);
        }
     
        public int UpdateWithFilter(string Table, string Filter, string[] Args)
        {
            string Sql;
            Sql = Table + "_MX_" + Filter;
            return Execute(Sql, Args);
        }
     
        public int UpdateWithFilter(string Table, string Filter, DataTable Datos)
        {
            string Sql;
            Sql = Table + "_MX_" + Filter;
            return Execute(Sql, Datos);
        }
       
        public int UpdateWithFilter(string Table, string Filter)
        {
            string sql;
            sql = Table + "_MX_" + Filter;
            return Execute(sql);
        }
      
        public string CheckConection()
        {
            try
            {
                System.Data.SqlClient.SqlConnection AuxiConec = new System.Data.SqlClient.SqlConnection(Stringsql);
                AuxiConec.Open();
                AuxiConec.Close();
                return "OK";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        #endregion

        #region Private Database Functions
        private void Conection()
        {
            if (mConection == null)
                mConection = new System.Data.SqlClient.SqlConnection(Stringsql);
            {
                var withBlock = mConection;
                if (mConection.State != ConnectionState.Open)
                    withBlock.Open();
            }
        }
       
        private string ExecuteValue(string Sql, string[] Args)
        {
            string vExecuteValue = "";
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommand(Sql);
            try
            {
                SqlComm = LoadParameters(SqlComm, Args);
                {
                    var withBlock = SqlComm;
                    if (this.VTimeOut.Trim() != "")
                        withBlock.CommandTimeout = ToInt(VTimeOut);
                    withBlock.Parameters[withBlock.Parameters.Count - 1].Value = "";
                    withBlock.ExecuteNonQuery();
                    vExecuteValue = withBlock.Parameters[withBlock.Parameters.Count - 1].Value.ToString().Trim();
                    withBlock.Dispose();
                }
                SqlComm = null;
            }
            catch (Exception ex)
            {
                if (SqlComm != null) { SqlComm.Dispose();  }
                SqlComm = null;
                ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace);
                throw ex;
            }
            return vExecuteValue;
        }
     
        private string ExecuteValue(string Sql, DataTable Args)
        {
            string vExecuteValue = "";
            int lAffected;
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommand(Sql);
            try
            {
                {
                    var withBlock = SqlComm;
                    if (this.VTimeOut.Trim() != "")
                        withBlock.CommandTimeout = ToInt(VTimeOut);
                    for (lAffected = 0; lAffected <= Args.Rows.Count - 1; lAffected++)
                        withBlock.Parameters[lAffected + 1].Value = Args.Rows[0][lAffected];
                    withBlock.Parameters[withBlock.Parameters.Count - 1].Value = "";
                    withBlock.ExecuteNonQuery();
                    vExecuteValue = withBlock.Parameters[withBlock.Parameters.Count - 1].Value.ToString().Trim();
                    withBlock.Dispose();
                }
                SqlComm = null;
            }
            catch (Exception ex)
            {
                if (SqlComm != null) { SqlComm.Dispose(); }
                SqlComm = null;
                ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace);
                throw ex;
            }
            return vExecuteValue;
        }
       
        private string ExecuteValue(string Sql, int Args)
        {
            string vExecuteValue = "";
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommand(Sql);
            try
            {
                {
                    var withBlock = SqlComm;
                    if (this.VTimeOut.Trim() != "")
                        withBlock.CommandTimeout = ToInt(VTimeOut);
                    if (withBlock.Parameters.Count != 3)
                        vExecuteValue = "Stored con demasiados Parametros";
                    else
                    {
                        withBlock.Parameters[1].Value = Args;
                        int N;
                        for (N = 1; N <= withBlock.Parameters.Count - 1; N++)
                        {
                            if (withBlock.Parameters[N].Direction == ParameterDirection.InputOutput)
                            {
                                withBlock.Parameters[N].Value = "";
                                break;
                            }
                        }
                        withBlock.ExecuteNonQuery();
                        vExecuteValue = withBlock.Parameters[N].Value.ToString().Trim();
                        withBlock.Dispose();
                    }
                    withBlock.Dispose();
                }
                SqlComm = null;
            }
            catch (Exception ex)
            {
                if (SqlComm != null) { SqlComm.Dispose(); }
                SqlComm = null;
                ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace);
                throw ex;
            }
            return vExecuteValue;
        }
    
        private string ExecuteValue(string Sql, string StringCommand)
        {
            string vExecuteValue = "";
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommand(Sql);
            try
            {
                {
                    var withBlock = SqlComm;
                    if (this.VTimeOut.Trim() != "")
                        withBlock.CommandTimeout = ToInt(VTimeOut);
                    if (withBlock.Parameters.Count != 3)
                        vExecuteValue = "Stored con demasiados Parametros";
                    else
                    {
                        withBlock.Parameters[1].Value = StringCommand;
                        int N;
                        for (N = 1; N <= SqlComm.Parameters.Count - 1; N++)
                        {
                            if (withBlock.Parameters[N].Direction == ParameterDirection.InputOutput)
                            {
                                withBlock.Parameters[N].Value = "";
                                break;
                            }
                        }
                        withBlock.ExecuteNonQuery();
                        vExecuteValue = withBlock.Parameters[N].Value.ToString().Trim();
                        withBlock.Dispose();
                    }
                    withBlock.Dispose();
                }
                SqlComm = null;
            }
            catch (Exception ex)
            {
                if (SqlComm != null) { SqlComm.Dispose(); }
                 SqlComm = null;
                ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace);
                throw ex;
            }
            return vExecuteValue;
        }
       
        private string ExecuteValue(string Sql)
        {
            string vExecuteValue = "";
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommand(Sql);
            try
            {
                {
                    var withBlock = SqlComm;
                    if (this.VTimeOut.Trim() != "")
                        withBlock.CommandTimeout = ToInt(VTimeOut);
                    withBlock.ExecuteNonQuery();
                    vExecuteValue = withBlock.Parameters[1].Value.ToString().Trim();
                    withBlock.Dispose();
                }
            }
            catch (Exception ex)
            {
                if (SqlComm != null) { SqlComm.Dispose(); }
                   ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace);
                throw ex;
            }
            return vExecuteValue;
        }
       
        public int Execute(string Sql, string[] Args)
        {
            int vExecuteValue = 0;
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommand(Sql);
            {
                var withBlock = SqlComm;
                try
                {
                    if (this.VTimeOut.Trim() != "")
                        withBlock.CommandTimeout = ToInt(VTimeOut);
                    SqlComm = LoadParameters(SqlComm, Args);
                    withBlock.ExecuteNonQuery();
                    vExecuteValue = ToInt(withBlock.Parameters[0].Value.ToString());
                    withBlock.Dispose();
                    SqlComm = null;
                }
                catch (Exception ex)
                {
                    if (SqlComm != null) { SqlComm.Dispose(); }
                    if (withBlock != null) { withBlock.Dispose(); }
                    SqlComm = null;
                    ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace);
                    throw ex;
                }
                return vExecuteValue;
            }
        }
        
        public int Execute(string Sql)
        {
            int vExecuteValue = 0;
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommand(Sql);
            {
                var withBlock = SqlComm;
                try
                {
                    if (this.VTimeOut.Trim() != "")
                        withBlock.CommandTimeout = ToInt(VTimeOut);
                    withBlock.ExecuteNonQuery();
                    if (withBlock.Parameters.Count > 0)
                        vExecuteValue = ToInt(withBlock.Parameters[0].Value.ToString());
                    withBlock.Dispose();
                    SqlComm = null;
                }
                catch (Exception ex)
                {
                    if (SqlComm != null) { SqlComm.Dispose(); }
                    if (withBlock != null) { withBlock.Dispose(); }
                    SqlComm = null;
                    ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace);
                    throw ex;
                }
                return vExecuteValue;
            }
        }
       
        public int Execute(string Sql, DataRow Args)
        {
            int vExecuteValue = 0;
            int ColData;
            int IdNoVa = 0;
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommand(Sql);
            if (Sql.Substring(Sql.Trim().Length - 2, 2) == "_A")
                IdNoVa = 1;
            {
                var withBlock = SqlComm;
                try
                {
                    if (this.VTimeOut.Trim() != "")
                        withBlock.CommandTimeout = ToInt(VTimeOut);
                    for (ColData = 0 + IdNoVa; ColData <= Args.Table.Columns.Count - 1; ColData++)
                        withBlock.Parameters[ColData - IdNoVa + 1].Value = Args[ColData];
                    withBlock.ExecuteNonQuery();
                    vExecuteValue = Convert.ToInt32(withBlock.Parameters[0].Value.ToString());
                    withBlock.Dispose();
                    SqlComm = null;
                }
                catch (Exception ex)
                {
                    if (SqlComm != null) { SqlComm.Dispose(); }
                    if (withBlock != null) { withBlock.Dispose(); }
                    SqlComm = null;
                    ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace);
                    throw ex;
                }
                return vExecuteValue;
            }
        }
       
        public int Execute(string Sql, DataTable Args)
        {
            int vExecuteValue = 0;
            int Posicion;
            int ColData;
            int IdNoVa = 0;
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommand(Sql);
            if (Sql.Substring(Sql.Trim().Length - 2, 2) == "_A")
                IdNoVa = 1;
            {
                var withBlock = SqlComm;
                try
                {
                    if (this.VTimeOut.Trim() != "")
                        withBlock.CommandTimeout = ToInt(VTimeOut);
                    for (Posicion = 0; Posicion <= Args.Rows.Count - 1; Posicion++)
                    {
                        for (ColData = 0 + IdNoVa; ColData <= Args.Columns.Count - 1; ColData++)
                            withBlock.Parameters[ColData - IdNoVa + 1].Value = Args.Rows[Posicion][ColData];
                        withBlock.ExecuteNonQuery();
                    }
                    vExecuteValue = ToInt(withBlock.Parameters[0].Value.ToString());
                    withBlock.Dispose();
                    SqlComm = null;
                    Args.Dispose();
                }
                catch (Exception ex)
                {
                    if (SqlComm != null) { SqlComm.Dispose(); }
                    if (withBlock != null) { withBlock.Dispose(); }
                    Args.Dispose();
                    SqlComm = null;
                    ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace);
                    throw ex;
                }
            }
            return vExecuteValue;
        }
        
        public int Execute(string Sql, int Args)
        {
            int vExecuteValue = 0;
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommand(Sql);
            {
                var withBlock = SqlComm;
                try
                {
                    if (this.VTimeOut.Trim() != "")
                        withBlock.CommandTimeout = ToInt(VTimeOut);
                    withBlock.Parameters[1].Value = Args;
                    withBlock.ExecuteNonQuery();
                    vExecuteValue = ToInt(withBlock.Parameters[0].Value.ToString());
                    withBlock.Dispose();
                    SqlComm = null;
                }
                catch (Exception ex)
                {
                    if (SqlComm != null) { SqlComm.Dispose(); }
                    if (withBlock != null) { withBlock.Dispose(); }
                    SqlComm = null;
                    ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace);
                    throw ex;
                }
            }
            return vExecuteValue;
        }
     
        public int Execute(string Sql, string Args)
        {
            int vExecuteValue = 0;
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommand(Sql);
            {
                var withBlock = SqlComm;
                try
                {
                    if (this.VTimeOut.Trim() != "")
                        withBlock.CommandTimeout = ToInt(VTimeOut);
                    withBlock.Parameters[1].Value = Args;
                    withBlock.ExecuteNonQuery();
                    vExecuteValue = ToInt(withBlock.Parameters[0].Value.ToString());
                    withBlock.Dispose();
                    SqlComm = null;
                }
                catch (Exception ex)
                {
                    if (SqlComm != null) { SqlComm.Dispose(); }
                    if (withBlock != null) { withBlock.Dispose(); }
                    SqlComm = null;
                    ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace);
                    throw ex;
                }
            }
            return vExecuteValue;
        }
       
        public void ExecuteTxtCommand(string Sql)
        {
            Conection();
            System.Data.SqlClient.SqlCommand mCommand = new System.Data.SqlClient.SqlCommand(Sql, mConection);
            {
                var withBlock = mCommand;
                if (this.VTimeOut.Trim() != "")
                    withBlock.CommandTimeout = ToInt(VTimeOut);
                withBlock.Connection = mConection;
                withBlock.Transaction = mTransaction;
            }
            {
                var withBlock = mCommand;
                try
                {
                    withBlock.ExecuteNonQuery();
                    withBlock.Dispose();
                    mCommand = null;
                }
                catch (Exception ex)
                {
                    if (withBlock != null) { withBlock.Dispose(); }
                    mCommand = null;
                    ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"));
                    throw ex;
                }
            }
        }
    
        private System.Data.SqlClient.SqlCommand CreateCommand(string Procedimiento)
        {
            System.Data.SqlClient.SqlConnection ConectionAux;
            ConectionAux = new System.Data.SqlClient.SqlConnection(StringConection);
            ConectionAux.Open();
            System.Data.SqlClient.SqlCommand mCommand = new System.Data.SqlClient.SqlCommand(Procedimiento, ConectionAux);
            System.Data.SqlClient.SqlCommandBuilder mConstructor = new System.Data.SqlClient.SqlCommandBuilder();
            mCommand.CommandType = CommandType.StoredProcedure;
            if (ParameterCache.IsParametersCached(mCommand))
            {
                System.Data.SqlClient.SqlParameter[] param = (System.Data.SqlClient.SqlParameter[])ParameterCache.GetCachedParameters(mCommand);
                mCommand.Parameters.AddRange(param);
            }
            else
            {
                SqlCommandBuilder.DeriveParameters(mCommand);
                ParameterCache.CacheParameters(mCommand);
            }
            {
                var withBlock = mCommand;
                if (this.VTimeOut.Trim() != "")
                    withBlock.CommandTimeout = ToInt(VTimeOut);
                withBlock.Connection = mConection;
                withBlock.Transaction = mTransaction;
            }
            {
                var withBlock = ConectionAux;
                withBlock.Close();
                withBlock.Dispose();
            }
            ConectionAux = null;
            mConstructor.Dispose();
            return mCommand;
        }
      
        private System.Data.SqlClient.SqlCommand CreateCommandFromTxt(string command)
        {
            System.Data.SqlClient.SqlConnection ConectionAux = new System.Data.SqlClient.SqlConnection(StringConection);
            ConectionAux.Open();
            System.Data.SqlClient.SqlCommand mCommand = new System.Data.SqlClient.SqlCommand(command);
            System.Data.SqlClient.SqlCommandBuilder mConstructor = new System.Data.SqlClient.SqlCommandBuilder();
            mCommand.CommandType = CommandType.Text;
            {
                var withBlock = mCommand;
                if (this.VTimeOut.Trim() != "")
                    withBlock.CommandTimeout = ToInt(VTimeOut);
                withBlock.Connection = mConection;
                withBlock.Transaction = mTransaction;
            }
            {
                var withBlock = ConectionAux;
                withBlock.Close();
                withBlock.Dispose();
            }
            ConectionAux = null;
            mConstructor.Dispose();
            return mCommand;
        }
    
        private System.Data.SqlClient.SqlCommand LoadParameters(System.Data.SqlClient.SqlCommand Command, string[] Args)
        {
            System.Data.SqlClient.SqlCommand rLoadParameters;
            int Posicion;
            SqlDbType TipoParametro;
            {
                var withBlock = Command;
                if (this.VTimeOut.Trim() != "")
                    withBlock.CommandTimeout = ToInt(VTimeOut);
                for (Posicion = 0; Posicion <= Args.Length - 1; Posicion++)
                {
                    TipoParametro = withBlock.Parameters[Posicion + 1].SqlDbType;
                    if (Args[Posicion].Trim().ToUpper() == "DBNULL")
                        withBlock.Parameters[Posicion + 1].Value = System.DBNull.Value;
                    else
                        switch (TipoParametro)
                        {
                            case System.Data.SqlDbType.BigInt:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = System.Convert.ToInt64(Args[Posicion]);
                                    break;
                                }

                            case SqlDbType.Binary:
                                {
                                    break;
                                }

                            case SqlDbType.Bit:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = CLogico(Args[Posicion]).ToString().Trim();
                                    break;
                                }

                            case SqlDbType.Char:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = Args[Posicion].Trim();
                                    break;
                                }

                            case SqlDbType.DateTime:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = Args[Posicion];
                                    break;
                                }

                            case SqlDbType.Decimal:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = CDecimal(Args[Posicion]);
                                    break;
                                }

                            case SqlDbType.Float:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = CDoble(Args[Posicion]);
                                    break;
                                }

                            case SqlDbType.Image:
                                {
                                    break;
                                }

                            case SqlDbType.Int:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = ToInt(Args[Posicion]);
                                    break;
                                }

                            case SqlDbType.Money:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = CDoble(Args[Posicion]);
                                    break;
                                }

                            case SqlDbType.NChar:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = Args[Posicion].Trim();
                                    break;
                                }

                            case SqlDbType.NText:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = Args[Posicion].Trim();
                                    break;
                                }

                            case SqlDbType.NVarChar:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = Args[Posicion].Trim();
                                    break;
                                }

                            case SqlDbType.Real:
                                {
                                    break;
                                }

                            case SqlDbType.SmallDateTime:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = Args[Posicion];
                                    break;
                                }

                            case SqlDbType.SmallInt:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = ToInt(Args[Posicion]);
                                    break;
                                }

                            case SqlDbType.SmallMoney:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = CDoble(Args[Posicion]);
                                    break;
                                }

                            case SqlDbType.Text:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = Args[Posicion].Trim();
                                    break;
                                }

                            case SqlDbType.Xml:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = Args[Posicion].Trim();
                                    break;
                                }

                            case SqlDbType.Timestamp:
                                {
                                    break;
                                }

                            case SqlDbType.TinyInt:
                                {
                                    break;
                                }

                            case SqlDbType.UniqueIdentifier:
                                {
                                    break;
                                }

                            case SqlDbType.VarBinary:
                                {
                                    break;
                                }

                            case SqlDbType.VarChar:
                                {
                                    withBlock.Parameters[Posicion + 1].Value = Args[Posicion].Trim();
                                    break;
                                }

                            case SqlDbType.Variant:
                                {
                                    break;
                                }
                        }
                }
            }
            rLoadParameters = Command;
            return rLoadParameters;
        }
     
        public DataTable GetRecords(string Sql, string[] Args)
        {
            DataTable rGetRecords= new DataTable();
            DataTable Datatable = new DataTable();
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommand(Sql);
            System.Data.SqlClient.SqlDataAdapter DataAdapter = new System.Data.SqlClient.SqlDataAdapter(SqlComm);
            {
                var withBlock = SqlComm;
                if (this.VTimeOut.Trim() != "")
                    withBlock.CommandTimeout = ToInt(VTimeOut);
                try
                {
                    SqlComm = LoadParameters(SqlComm, Args);
                    DataAdapter.Fill(Datatable);
                    rGetRecords = Datatable;
                    DataAdapter.Dispose();
                    withBlock.Dispose();
                    SqlComm = null;
                    DataAdapter = null;
                }
                catch (Exception ex)
                {
                    DataAdapter.Dispose();
                    DataAdapter = null;
                    withBlock.Dispose();
                    SqlComm = null;
                    DataAdapter = null;
                    ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace);
                    throw ex;
                }
            }
            return rGetRecords;
        }

        public DataTable GetRecords(string Sql, DataTable Args)
        {
            DataTable rGetRecords = new DataTable();
            DataTable DataTable = new DataTable();
            int lAffected;
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommand(Sql);
            System.Data.SqlClient.SqlDataAdapter DataAdapter = new System.Data.SqlClient.SqlDataAdapter(SqlComm);
            {
                var withBlock = SqlComm;
                if (this.VTimeOut.Trim() != "")
                    withBlock.CommandTimeout = ToInt(VTimeOut);
                try
                {
                    for (lAffected = 0; lAffected <= Args.Rows.Count - 1; lAffected++)
                        withBlock.Parameters[lAffected + 1].Value = Args.Rows[0][lAffected];
                    DataAdapter.Fill(DataTable);
                    rGetRecords = DataTable;
                    DataAdapter.Dispose();
                    withBlock.Dispose();
                    SqlComm = null;
                    DataAdapter = null;
                    Args.Dispose();
                }
                catch (Exception ex)
                {
                    DataAdapter.Dispose();
                    DataAdapter = null;
                    withBlock.Dispose();
                    SqlComm = null;
                    Args.Dispose();
                    ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace);
                    throw ex;
                }
            }
            return rGetRecords;
        }
    
        public DataTable GetRecords(string Sql)
        {
            DataTable rGetRecords = new DataTable();
            DataTable DataTable = new DataTable();
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommand(Sql);
            System.Data.SqlClient.SqlDataAdapter DataAdapter = new System.Data.SqlClient.SqlDataAdapter(SqlComm);
            try
            {
                DataAdapter.Fill(DataTable);
                rGetRecords = DataTable;
                SqlComm.Dispose();
                DataAdapter.Dispose();
                DataTable.Dispose();
                SqlComm = null;
                DataAdapter = null;
            }
            catch (Exception ex)
            {
                SqlComm.Dispose();
                DataAdapter.Dispose();
                DataTable.Dispose();
                SqlComm = null;
                DataAdapter = null;
                ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace);
                throw ex;
            }
            return rGetRecords;
        }
       
        public DataTable GetRecords(string Sql, int Args)
        {
            DataTable rGetRecords = new DataTable();
            DataTable Datatable = new DataTable();
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommand(Sql);
            System.Data.SqlClient.SqlDataAdapter DataAdapter = new System.Data.SqlClient.SqlDataAdapter(SqlComm);
            {
                var withBlock = SqlComm;
                try
                {
                    if (this.VTimeOut.Trim() != "")
                        withBlock.CommandTimeout = ToInt(VTimeOut);
                    withBlock.Parameters[1].Value = Args;
                    DataAdapter.Fill(Datatable);
                    rGetRecords = Datatable;
                    DataAdapter.Dispose();
                    withBlock.Dispose();
                    SqlComm = null;
                    DataAdapter = null;
                }
                catch (Exception ex)
                {
                    DataAdapter.Dispose();
                    withBlock.Dispose();
                    SqlComm = null;
                    DataAdapter = null;
                    ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"));
                    throw ex;
                }
            }
            return rGetRecords;
        }
    
        public DataTable GetRecords(string Sql, string StringCommand)
        {
            DataTable rGetRecords = new DataTable();
            DataTable Datatable = new DataTable();
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommand(Sql);
            System.Data.SqlClient.SqlDataAdapter DataAdapter = new System.Data.SqlClient.SqlDataAdapter(SqlComm);
            {
                var withBlock = SqlComm;
                try
                {
                    if (this.VTimeOut.Trim() != "")
                        withBlock.CommandTimeout = ToInt(VTimeOut);
                    if (StringCommand.Trim().ToUpper() == "DBNULL")
                        withBlock.Parameters[1].Value = System.DBNull.Value;
                    else
                        withBlock.Parameters[1].Value = StringCommand.Trim();
                    DataAdapter.Fill(Datatable);
                    rGetRecords = Datatable;
                    DataAdapter.Dispose();
                    withBlock.Dispose();
                    SqlComm = null;
                    DataAdapter = null;
                }
                catch (Exception ex)
                {
                    DataAdapter.Dispose();
                    withBlock.Dispose();
                    SqlComm = null;
                    DataAdapter = null;
                    ErrsAnalyzer(Sql + "=" + ex.Message, ex.Source, ex.StackTrace);
                    throw ex;
                }
            }
            return rGetRecords;
        }
    
        public DataTable GetRecordsFromTxtCommand(string Command)
        {
            DataTable rGetRecordsFromTxtCommand = new DataTable();
            DataTable DataTable = new DataTable();
            Conection();
            System.Data.SqlClient.SqlCommand SqlComm = CreateCommandFromTxt(Command);
            System.Data.SqlClient.SqlDataAdapter DataAdapter = new System.Data.SqlClient.SqlDataAdapter(SqlComm);
            try
            {
                DataAdapter.Fill(DataTable);
                rGetRecordsFromTxtCommand = DataTable;
                SqlComm.Dispose();
                DataAdapter.Dispose();
                DataTable.Dispose();
                SqlComm = null;
                DataAdapter = null;
            }
            catch (Exception ex)
            {
                SqlComm.Dispose();
                DataAdapter.Dispose();
                DataTable.Dispose();
                SqlComm = null;
                DataAdapter = null;
                ErrsAnalyzer(Command + "=" + ex.Message, ex.Source, ex.StackTrace);
                throw ex;
            }
            return rGetRecordsFromTxtCommand;
        }
     
        private void ErrsAnalyzer(string TxtError, string Source, string StackTrace)
        {
            string Aplication = "";
            string ProductName = "SqlDataServer";
            string File = "";
            string Drive = "";
            if (System.Configuration.ConfigurationManager.AppSettings["DriveAppErr"] == null)
            {
                if (Drive.Trim().Length == 0)
                    Drive = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location).Substring(0, 2);
                if (Drive.Trim().Length == 0)
                    Drive = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location).Substring(0, 2);
            }
            else
                Drive = System.Configuration.ConfigurationManager.AppSettings["DriveAppErr"].ToString();
            if (File.Trim().Length == 0)
            {
                string Path = "";
                if (Drive.Trim().Length == 0)
                {
                    if (System.IO.Directory.Exists(@"f:\"))
                        Path = @"F:\AppErrs\";
                    else
                        Path = @"c:\AppErrs\";
                }
                else
                    Path = Drive + @"\AppErrs\";
                Path = Path + DateTime.Now.Year.ToString() + @"\" + (DateTime.Now.Month + 100).ToString().Substring(1, 2) + @"\" + (DateTime.Now.Day + 100).ToString().Substring(1, 2);
                if (!System.IO.Directory.Exists(Path))
                    System.IO.Directory.CreateDirectory(Path);
                if (Aplication.Trim().Length == 0)
                    File = Path + @"\Errs.xml";
                else
                    File = Path + @"\" + Aplication.Trim() + ".xml";
            }
            DataSet MidataSet = new DataSet();
            MidataSet.DataSetName = "MisErrores";
            if (System.IO.File.Exists(File))
                MidataSet.ReadXml(File);
            else
            {
                DataTable MisRegistros = new DataTable();
                {
                    var withBlock = MisRegistros;
                    withBlock.Columns.Add("ProductName", System.Type.GetType("System.String"));
                    withBlock.Columns.Add("Aplication", System.Type.GetType("System.String"));
                    withBlock.Columns.Add("Class", System.Type.GetType("System.String"));
                    withBlock.Columns.Add("Function", System.Type.GetType("System.String"));
                    withBlock.Columns.Add("Date", System.Type.GetType("System.String"));
                    withBlock.Columns.Add("Time", System.Type.GetType("System.String"));
                    withBlock.Columns.Add("Source", System.Type.GetType("System.String"));
                    withBlock.Columns.Add("ErrDescription", System.Type.GetType("System.String"));
                    withBlock.Columns.Add("User", System.Type.GetType("System.String"));
                    withBlock.Columns.Add("StackTrace", System.Type.GetType("System.String"));
                }
                MidataSet.Tables.Add(MisRegistros);
                MisRegistros.Dispose();
            }
            string[] Args = new string[10];
            Args[0] = ProductName;
            Args[1] = "";
            Args[2] = "";
            Args[3] = "";
            Args[4] = System.DateTime.Now.Day.ToString().Trim() + "/" + System.DateTime.Now.Month.ToString().Trim() + "/" + System.DateTime.Now.Year.ToString().Trim();
            Args[5] = System.DateTime.Now.Hour.ToString().Trim() + ":" + System.DateTime.Now.Minute.ToString().Trim() + ":" + System.DateTime.Now.Second.ToString().Trim();
            Args[6] = Source;
            Args[7] = TxtError;
            Args[8] = User;
            Args[9] = StackTrace;
            {
                var withBlock = MidataSet;
                withBlock.Tables[0].Rows.Add(Args);
                withBlock.WriteXml(File);
                withBlock.Dispose();
            }
            Args = null;
        }
     
        public string ErrCode(string TheErr, bool Code)
        {
            string ConstErr = "DELETE statement conflicted with COLUMN REFERENCE constraint";
            if (ConstErr.Trim().Length <= TheErr.Trim().Length)
            {
                if (TheErr.Substring(0, ConstErr.Length).Trim() == ConstErr)
                {
                    if (Code)
                        return "-3";
                    else
                        return "DELETE statement conflicted with COLUMN REFERENCE constraint";
                }
            }
            ConstErr = "Violation of UNIQUE KEY";
            if (ConstErr.Trim().Length <= TheErr.Trim().Length)
            {
                if (TheErr.Substring(0, ConstErr.Length).Trim() == ConstErr)
                {
                    if (Code)
                        return "-1";
                    else
                        return "Violation of UNIQUE KEY";
                }
            }
            ConstErr = "INSERT statement conflicted with COLUMN FOREIGN KEY constraint";
            if (ConstErr.Trim().Length <= TheErr.Trim().Length)
            {
                if (TheErr.Substring(0, ConstErr.Length).Trim() == ConstErr)
                {
                    if (Code)
                        return "-6";
                    else
                        return "INSERT statement conflicted with COLUMN FOREIGN KEY constraint";
                }
            }
            ConstErr = "The stored procedure";
            if (ConstErr.Trim().Length <= TheErr.Trim().Length)
            {
                if (TheErr.Substring(0, ConstErr.Length).Trim() == ConstErr)
                {
                    if (TheErr.Substring(TheErr.Trim().Length - 14, 14) == "doesn't exist.")
                    {
                        if (Code)
                            return "-5";
                        else
                            return "A stored procedure was not found";
                    }
                }
            }
            switch (TheErr)
            {
                case object _ when TheErr == "SQL Server does not exist or access denied.":
                    {
                        if (Code)
                            return "-2";
                        else
                            return TheErr;
                        break;
                    }

                case object _ when TheErr == "The ConnectionString property has not been initialized.":
                    {
                        if (Code)
                            return "-4";
                        else
                            return TheErr;
                        break;
                    }

                default:
                    {
                        if (Code)
                            return "-999999";
                        else
                            return "Unknow Err";
                        break;
                    }
            }
        }
        #endregion

        #region Funciones de Formato Privadas
        private bool EsNumero(string Numero)
        {
            try
            {
                Convert.ToDouble(Numero);
            }
            catch
            {
                return false;
            }
            return true;
        }
       
        private double CDecimal(string Numero)
        {
            double rCDecimal = 0;
            if (Numero.Trim().Length == 0)
                return rCDecimal;
            try
            {
                System.Text.StringBuilder entero = new System.Text.StringBuilder();
                System.Text.StringBuilder decimales = new System.Text.StringBuilder();
                System.Text.StringBuilder Divisor = new System.Text.StringBuilder();
                Divisor.Append("1");
                string caracter;
                bool bandera = false;
                int n;
                for (n = 0; n <= Numero.Trim().Length - 1; n++)
                {
                    caracter = Izquierda(Derecha(Numero, Numero.Length - n), 1);
                    if (!EsNumero(caracter) & caracter != "-" & caracter != "+")
                        bandera = true;
                    else if (!bandera)
                        entero.Append(caracter);
                    else
                    {
                        decimales.Append(caracter);
                        Divisor.Append("0");
                    }
                }
                rCDecimal = double.Parse(entero.ToString());
                if (decimales.ToString().Trim().Length != 0)
                {
                    if (rCDecimal >= 0)
                        rCDecimal = rCDecimal + (long.Parse(decimales.ToString()) / (double)long.Parse(Divisor.ToString()));
                    else
                        rCDecimal = rCDecimal - (long.Parse(decimales.ToString()) / (double)long.Parse(Divisor.ToString()));
                }
                if (entero.ToString() == "-0")
                    rCDecimal = rCDecimal * -1;
            }
            catch
            {
                 rCDecimal=0;
            }
            return rCDecimal;
        }
      
        private int ToInt(string Numero)
        {
            int rToInt = 0;
            if (Numero.ToLower() == "true" | Numero.ToLower() == "verdadero")
            {
                rToInt = 1;
                return rToInt;
            }
            if (Numero.ToLower() == "false" | Numero.ToLower() == "falso")
            {
                rToInt = 0;
                return rToInt;
            }
            try
            {
                rToInt = int.Parse(Numero);
            }
            catch
            {
                rToInt=0;
            }
            return rToInt;
        }
       
        private bool CLogico(string Numero)
        {
            bool rClogico;
            if (Numero.Trim() == "0")
            {
                rClogico = false;
                return rClogico;
            }
            if (Numero.Trim() == "1")
            {
                rClogico = true;
                return rClogico;
            }
            try
            {
                rClogico = bool.Parse(Numero);
            }
            catch
            {
                rClogico= false;
            }
            return rClogico;
        }
       
        private double CDoble(string Numero)
        {
            double rCDoble = 0;
            if (Numero.Trim().Length == 0)
                return 0;
            try
            {
                System.Text.StringBuilder entero = new System.Text.StringBuilder();
                System.Text.StringBuilder decimales = new System.Text.StringBuilder();
                System.Text.StringBuilder Divisor = new System.Text.StringBuilder();
                Divisor.Append("1");
                string caracter;
                bool bandera = false;
                int n;
                for (n = 0; n <= Numero.Trim().Length - 1; n++)
                {
                    caracter = Izquierda(Derecha(Numero, Numero.Length - n), 1);
                    if (!EsNumero(caracter) & caracter != "-" & caracter != "+")
                        bandera = true;
                    else if (!bandera)
                        entero.Append(caracter);
                    else
                    {
                        decimales.Append(caracter);
                        Divisor.Append("0");
                    }
                }
                rCDoble = double.Parse(entero.ToString());
                if (decimales.ToString().Trim().Length != 0)
                {
                    if (rCDoble >= 0)
                        rCDoble = rCDoble + (long.Parse(decimales.ToString()) / (double)long.Parse(Divisor.ToString()));
                    else
                        rCDoble = rCDoble - (long.Parse(decimales.ToString()) / (double)long.Parse(Divisor.ToString()));
                }
                if (entero.ToString() == "-0")
                    rCDoble = rCDoble * -1;
            }
            catch
            {
                rCDoble= 0;
            }
            return rCDoble;
        }
      
        private string Izquierda(string StringCommand, int Posiciones)
        {
            if (Posiciones > StringCommand.Trim().Length)
                return StringCommand;
            return StringCommand.Trim().Substring(0, Posiciones);
        }
      
        private string Derecha(string StringCommand, int Posiciones)
        {
            if (Posiciones > StringCommand.Trim().Length)
                return StringCommand;
            return StringCommand.Trim().Substring(StringCommand.Trim().Length - Posiciones, Posiciones);
        }
        
        private DateTime CFecha(string Fecha)
        {
            try
            {
                return DateTime.Parse(Fecha);
            }
            catch
            {
                return System.DateTime.Now;
            }
        }
      
        public static string LastValue(string Texto)
        {
            string rLastValue = "";
            Texto = Texto.Trim();
            if (Texto.Length == 0)
                return rLastValue;
            int Largo = Texto.Length;
            int N;
            for (N = Largo - 1; N >= 0; N += -1)
            {
                if (Texto.Substring(N, 1) == ".")
                {
                    rLastValue = Texto.Substring(N + 1, Largo - 1 - N);
                    break;
                }
            }
            if (rLastValue == "")
                rLastValue = Texto;
            return rLastValue;
        }

        #endregion

        #region Transactions
        public void BeginTransaccion()
        {
            Conection();
            mTransaction = mConection.BeginTransaction();
            InTransaction = true;
        }
        public void EndTransaccion()
        {
            try
            {
                {
                    var withBlock = mTransaction;
                    withBlock.Commit();
                    withBlock.Dispose();
                }
                InTransaction = false;
                mTransaction = null;
            }
            catch (Exception ex)
            {
                mTransaction.Connection.Close();
                InTransaction = false;
                mTransaction = null;
                throw ex;
            }
        }
        public void CancelTransaccion()
        {
            try
            {
                if (InTransaction)
                {
                    {
                        var withBlock = mTransaction;
                        withBlock.Rollback();
                        withBlock.Dispose();
                    }
                }
                mTransaction = null;
                InTransaction = false;
            }
            catch (Exception Ex)
            {
                if (mTransaction != null) {                    mTransaction.Connection.Close(); }
                mTransaction = null;
                InTransaction = false;
                throw Ex;
            }
        }
        #endregion

        #region Dispose
        protected override void Dispose(bool Disposing)
        {
            if (Disposing)
            {
                if (mConection != null)
                {
                    if (!InTransaction)
                    {
                        {
                            var withBlock = mConection;
                            withBlock.Close();
                            withBlock.Dispose();
                        }
                        mConection = null;
                        base.Dispose(Disposing);
                    }
                }
            }
        }
        #endregion
    }
}
