using System;
using System.Collections.Generic;
using System.Text;

namespace ExecCobolAnalysis
{
    /// <summary>
    /// 呼出関数管理クラス　※下記参照の上、distinct用の実装箇所あり
    /// 参照：https://qiita.com/Chrowa3/items/51e7033aa687c6274ad4
    /// 参照：https://docs.microsoft.com/ja-jp/dotnet/api/system.linq.enumerable.distinct?redirectedfrom=MSDN&view=netcore-3.1#System_Linq_Enumerable_Distinct__1_System_Collections_Generic_IEnumerable___0__
    /// </summary>
    class CalledMethod : IEquatable<CalledMethod>
    {
        public string Name { get; }
        public bool ModuleFlg { get; }
        public int MethodListIndex { get; internal set; }
        public string Conditions { get; internal set; }

        public CalledMethod(string name, bool moduleFlg, List<string> conditions)
        {
            Name = name;
            ModuleFlg = moduleFlg;
            StringBuilder sb = new StringBuilder();
            foreach (string condition in conditions)
            {
                if (sb.Length > 0)
                    sb.Append(" かつ ");

                sb.Append("【" + condition + "】");
            }
            Conditions = sb.ToString();
        }

        public override int GetHashCode()
        {
            return this.Name.GetHashCode();
        }

        bool IEquatable<CalledMethod>.Equals(CalledMethod cm)
        {
            if (cm == null)
                return false;
            return (this.Name == cm.Name);
        }
    }
}
