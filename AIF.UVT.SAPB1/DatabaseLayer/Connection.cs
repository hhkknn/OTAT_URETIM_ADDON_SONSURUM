using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataTables
{
    public class Connection
    {
        //committ.
        private static SqlConnection _sql;

        public static SqlConnection sql
        {
            get
            {
                {
                    try
                    {
                        if (_sql == null)
                        {
                            //_sql = new SqlConnection(@"Server=.;Database=SBODemoTR;User Id=sa;Password=123qwe;");
                            _sql = new SqlConnection(@"Server=172.55.10.6;Database=SAPKANTAR;User Id=SapKantar;Password=suttartim;");
                        }
                        else
                        {
                            return _sql;
                        }
                    }
                    catch (Exception)
                    {
                    }

                    return _sql;
                }
            }
        }
    }
}