using ExcelToDBAppModel.HR;
using ExcelToDBAppModel.HR.Designation;
using System;
using System.Data.SqlClient;

namespace ExcelToDBAppDBLib.DAL
{
    public class ImportData
    {
        public int ImportHRDesignation(HRDesignation_VM model)
        {
            // dtItems has rows for modified records
            using (AdoHelper db = new AdoHelper())
            {
                int n = 0;
                try
                {
                    db.BeginTransaction();
                    SqlParameter[] parameter =  {
                                 new SqlParameter("@DesignationId",model.DesignationId),
                                 new SqlParameter("@DesignationName", model.DesignationName),
                                 new SqlParameter("@CreatedBy", model.CreatedBy),
                                 new SqlParameter("@Option", 1)
                    };
                    n = db.ExecNonQueryProc("sp_Hr_AddUpdateListDesignation", parameter);
                    db.Commit();
                    return n;
                }
                catch (Exception ex)
                {
                    db.Rollback();
                    string s = ex.Message.ToString();
                    throw;
                }
            }
        }
        public int ImportHRDepartment(HR_Department_VM model)
        {
            // dtItems has rows for modified records
            using (AdoHelper db = new AdoHelper())
            {
                int n = 0;
                try
                {
                    db.BeginTransaction();
                    SqlParameter[] parameter =  {
                                 new SqlParameter("@DepartmentId",model.DepartmentId),
                                 new SqlParameter("@DepartmentName", model.DepartmentName),
                                 new SqlParameter("@CreatedBy", model.CreatedBy),
                                 new SqlParameter("@Option", 1)
                    };
                    n = db.ExecNonQueryProc("sp_Hr_AddUpdateListDepartment", parameter);
                    db.Commit();
                    return n;
                }
                catch (Exception ex)
                {
                    db.Rollback();
                    string s = ex.Message.ToString();
                    throw;
                }
            }
        }
    }
}
