using System;
using System.Data;

namespace ExecCobolAnalysis
{
    /// <summary>
    /// DB情報クラス
    /// </summary>
    class DbInfo : IEquatable<DbInfo>
    {
        public string Name_P { get; }
        public string Name_L { get; }
        public bool SelectFlg { get; private set; } = false;
        public bool InsertFlg { get; private set; } = false;
        public bool UpdateFlg { get; private set; } = false;
        public bool DeleteFlg { get; private set; } = false;
        public bool CreateFlg { get; private set; } = false;

        public DbInfo(string name, SqlType type, DataTable dt)
        {
            Name_P = name;
            Name_L = string.Empty;
            foreach (DataRow dr in dt.Rows)
            {
                if (name == dr[CommonConst.TABLE_P].ToString())
                    Name_L = dr[CommonConst.TABLE_L].ToString();
            }
            SetCrudFlg(type);
        }

        public void SetCrudFlg(SqlType type)
        {
            switch (type)
            {
                case SqlType.Select:
                    SelectFlg = true;
                    break;
                case SqlType.Insert:
                    InsertFlg = true;
                    break;
                case SqlType.Update:
                    UpdateFlg = true;
                    break;
                case SqlType.Delete:
                    DeleteFlg = true;
                    break;
                case SqlType.Create:
                    CreateFlg = true;
                    break;
                default:
                    break;
            }
        }

        public override int GetHashCode()
        {
            return this.Name_P.GetHashCode();
        }

        bool IEquatable<DbInfo>.Equals(DbInfo cm)
        {
            if (cm == null)
                return false;
            return (this.Name_P == cm.Name_P);
        }

    }
}
