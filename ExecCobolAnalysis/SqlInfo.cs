using Microsoft.SqlServer.Management.SqlParser.Parser;
using System.Collections.Generic;
using System.Linq;
using TransactSqlHelpers;

namespace ExecCobolAnalysis
{
    /// <summary>
    /// SQL情報クラス
    /// </summary>
    class SqlInfo
    {
        public string Value { get; }
        public IEnumerable<TokenInfo> TokenList { get; }
        public SqlType Type { get; }
        public string UseddMethodName { get; internal set; }
        public string CursorName { get; }

        public SqlInfo(string value, IEnumerable<TokenInfo> tokenList, SqlType type, string useddMethodName, string cursorName)
        {
            Value = value;
            TokenList = tokenList;
            Type = type;
            UseddMethodName = useddMethodName;
            CursorName = cursorName;
        }

        public SqlInfo()
        {
        }

        public List<string> GetDbList()
        {
            List<string> dbList = new List<string>();
            bool dbAddFlg1 = false;
            bool dbAddFlg2 = false;

            foreach (var token in TokenList)
            {
                if (token.Token == Tokens.TOKEN_FROM || token.Token == Tokens.TOKEN_JOIN)
                {
                    dbAddFlg1 = true;
                    continue;
                }
                if (dbAddFlg1 && token.Token == Tokens.TOKEN_ID)
                {
                    // テーブル名記載箇所(SELECT, DELETE)
                    dbList.Add(token.Sql);
                    continue;
                }
                if (dbAddFlg1 && token.Token != Tokens.TOKEN_ID)
                {
                    dbAddFlg1 = false;
                    continue;
                }

                if (token.Token == Tokens.TOKEN_INSERT || token.Token == Tokens.TOKEN_UPDATE || token.Token == Tokens.TOKEN_CREATE)
                {
                    dbAddFlg2 = true;
                    continue;
                }
                if (dbAddFlg2 && token.Token == Tokens.TOKEN_ID)
                {
                    // テーブル名記載箇所(INSERT, UPDATE, CREATE)
                    dbList.Add(token.Sql);
                    dbAddFlg2 = false;
                    continue;
                }
            }
            dbList = dbList.Distinct().OrderBy(x => x).ToList();

            return dbList;
        }

        public string SqlTypeToString()
        {
            string ret = string.Empty;
            switch (Type)
            {
                case SqlType.Select:
                    ret = "SELECT";
                    break;
                case SqlType.Insert:
                    ret = "INSERT";
                    break;
                case SqlType.Update:
                    ret = "UPDATE";
                    break;
                case SqlType.Delete:
                    ret = "DELETE";
                    break;
                default:
                    break;
            }
            return ret;
        }

        public SqlType StringToSqlType(string value, SqlType sqlType)
        {
            switch (value)
            {
                case "SELECT":
                    sqlType = (sqlType == SqlType.None) ? SqlType.Select : sqlType;
                    break;
                case "INSERT":
                    sqlType = (sqlType == SqlType.None) ? SqlType.Insert : sqlType;
                    break;
                case "UPDATE":
                    sqlType = (sqlType == SqlType.None) ? SqlType.Update : sqlType;
                    break;
                case "DELETE":
                    sqlType = (sqlType == SqlType.None) ? SqlType.Delete : sqlType;
                    break;
                default:
                    break;
            }

            return sqlType;
        }
    }
}
