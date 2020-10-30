using System.Collections.Generic;

namespace ExecCobolAnalysis
{
    /// <summary>
    /// 関数管理クラス
    /// </summary>
    class Method
    {
        public string MethodNameP { get; }
        public string MethodNameL { get; }
        public int StartIndex { get; }
        public int EndIndex { get; internal set; }
        public List<string> CheckPointList { get; internal set; }
        public List<CalledMethod> CalledMethodList { get; internal set; }
        public bool CalledFlg { get; internal set; }

        public Method(string methodNameP, string methodNameL, int startIndex, int endIndex)
        {
            MethodNameP = methodNameP;
            MethodNameL = methodNameL;
            StartIndex = startIndex;
            EndIndex = endIndex;
            CheckPointList = new List<string>();
            CalledMethodList = new List<CalledMethod>();
            CalledFlg = false;
        }

        public bool DetectCheckPoint(string word)
        {
            // 引数が関数物理名を文字列中に含んでいれば、チェックポイント
            if (word.Contains(MethodNameP.Replace("-PROC", "")))
                return true;

            return false;
        }
    }
}
