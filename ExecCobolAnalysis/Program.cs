using log4net;
using Microsoft.SqlServer.Management.SqlParser.Parser;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using TransactSqlHelpers;

namespace ExecCobolAnalysis
{
    class Program
    {
        #region 変数
        static string ResultFileName = ConfigurationManager.AppSettings["ResultFilePath"];
        static string DbDifineFileName = ConfigurationManager.AppSettings["DbDifineFilePath"];
        static Encoding EncShiftJis = Encoding.GetEncoding("Shift_JIS");
        static Encoding EncUtf8 = Encoding.UTF8;
        static int InitRow = 2;
        static string SheetName = string.Empty;
        static ExcelWorksheet WsPgmInfo; // プログラム情報シート
        static ExcelWorksheet WsStruct; // 構造図シート
        static ExcelWorksheet WsMethodInfo; // 関数情報シート
        static Color ColorMethod = Color.RoyalBlue;
        static Color ColorModule = Color.DeepPink;
        private static readonly ILog _logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static Regex RegexIdentification = new Regex("(IDENTIFICATION|PROGRAM-ID|AUTHOR|DATE-WRITTEN|DATE-COMPILED)", RegexOptions.Compiled);
        static Regex RegexEnvironment = new Regex("(ENVIRONMENT|CONFIGURATION|SOURCE-COMPUTER|OBJECT-COMPUTER|INPUT-OUTPUT|FILE-CONTROL)", RegexOptions.Compiled);
        static Regex RegexData = new Regex("(DATA|FILE|WORKING-STORAGE|REPORT|SCREEN)", RegexOptions.Compiled);
        static Regex RegexPROCEDURE = new Regex("DISPLAY", RegexOptions.Compiled);
        #endregion

        enum Division
        {
            NONE = -1,
            IDENTIFICATION = 1,
            ENVIRONMENT = 2,
            DATA = 3,
            PROCEDURE = 4
        }

        private static int Main(string[] args)
        {
            // 出力ファイルが使用中ではないかチェック
            try
            {
                if (File.Exists(ResultFileName))
                {
                    using (Stream st = new FileStream(ResultFileName, FileMode.Open)) { }
                }
            }
            catch (Exception)
            {
                _logger.Error($"出力ファイル({ResultFileName})が使用中です。");
                return CommonConst.RETURN_ERR_200;
            }

            // 解析対象ファイルが読み込み可能かチェック
            if (args.Length < 1)
            {
                _logger.Error("解析対象のファイルが指定されていません。");
                return CommonConst.RETURN_ERR_100;
            }
            string file = args[0];

            if (!File.Exists(file))
            {
                _logger.Error($"{file}は存在しません。");
                return CommonConst.RETURN_ERR_100;
            }

            bool ret = false;
            try
            {
                _logger.Info($"処理開始。　{file}");
                ret = Exec(file);

            }
            catch (Exception ex)
            {
                _logger.Fatal(ex.Message);
            }
            finally
            {
                if (ret)
                    _logger.Info("処理正常終了。");
                else
                    _logger.Info("処理終了。エラーが発生しています。");
            }

            return CommonConst.RETURN_OK;
        }

        private static bool Exec(string file)
        {
            string logSectionName = string.Empty;
            string logLine = string.Empty;
            try
            {
                List<Method> methodList = new List<Method>();
                Dictionary<string, string> copyList = new Dictionary<string, string>();
                List<string> calledModuleList = new List<string>();
                List<IEnumerable<TokenInfo>> SqlTokenList = new List<IEnumerable<TokenInfo>>();
                List<SqlInfo> sqlInfoList = new List<SqlInfo>();
                List<SqlInfo> cursorList = new List<SqlInfo>();

                // =================================================================
                // 前処理①（関数の一覧をリスト化）
                // =================================================================
                #region 前処理①
                using (StreamReader sr = new StreamReader(file, EncShiftJis))
                {
                    int fileIndex = 0;
                    Division division = Division.NONE;
                    int methodIndex = -1;
                    bool inMethodErea = false;
                    bool inIfConditionDefineErea = false;
                    bool thenAddFlg = false;
                    bool whenConditionAddFlg = false;
                    string ifConditionTxt = string.Empty;
                    string evaluateConditionTxt = string.Empty;
                    string whenTxt = string.Empty;
                    List<string> conditions = new List<string>();
                    List<int> conditionCountList = new List<int>();
                    bool inSqlErea = false;
                    string sql = string.Empty;
                    SqlType sqlType = SqlType.None;
                    string cursorName = string.Empty;

                    while (sr.Peek() >= 0)
                    {
                        fileIndex++;

                        // 読み込んだ行を整形する
                        string line = sr.ReadLine();
                        logLine = line;
                        string fmtLine = FormatLine(line, false);
                        if (!inSqlErea)
                            fmtLine = fmtLine.Replace(".", "");

                        // テキストをスペースで区切った配列を作成
                        string[] arrWord = fmtLine.Split(' ');

                        // プログラム部変更の判定　※すでに手続き部にいる場合は必用なし
                        if (division != Division.PROCEDURE)
                        {
                            Division ret = CheckDivisionChanged(arrWord);
                            division = (ret != Division.NONE) ? ret : division;
                        }

                        // 対象外句はスルー
                        if (!CheckExcludedWords(arrWord, division))
                            continue;

                        // COBOLのプログラム定義部分による処理分岐
                        switch (division)
                        {
                            case Division.NONE:
                                continue;
                            case Division.IDENTIFICATION:
                                continue;
                            case Division.ENVIRONMENT:
                                continue;
                            case Division.DATA:
                                // コピー句を特定（特定行の1行手前がコメント行だと仮定する）
                                if (arrWord[0] == "COPY")
                                {
                                    string copyKey = GetArrayWord(arrWord, 1);
                                    string comLine = GetComment(file, fileIndex - 2);
                                    if (!copyList.ContainsKey(copyKey))
                                        copyList.Add(copyKey, comLine);
                                }

                                // カーソル開始行を特定　※DECLARE句が次の行に記載のある想定
                                if (String.Join(" ", arrWord) == "EXEC SQL")
                                {
                                    inSqlErea = true;
                                    sql = string.Empty;
                                    sqlType = SqlType.None;
                                    cursorName = string.Empty;
                                    continue;
                                }

                                // =================================================================
                                // カーソル内を解析
                                // =================================================================
                                if(inSqlErea)
                                    inSqlErea = AddSqlList(division, arrWord, string.Empty, ref sql, ref sqlType, ref cursorList, ref cursorName);

                                continue;
                            case Division.PROCEDURE:
                                // 関数の開始行を特定
                                if (arrWord[arrWord.Length - 1] == "SECTION")
                                {
                                    methodIndex++;

                                    // 関数名の論理名の特定（関数開始行の2行手前がコメント行だと仮定する）
                                    // ※関数開始行の3行前までを読み飛ばし、次の1行（＝コメント行）を読み込む
                                    string methodNameL = GetComment(file, fileIndex - 3);
                                    logSectionName = methodNameL;

                                    Method m = new Method(arrWord[0], methodNameL, fileIndex, -1);
                                    methodList.Add(m);
                                    inMethodErea = true;
                                    conditions.Clear();

                                    continue;
                                }

                                // =================================================================
                                // 関数内を解析
                                // =================================================================
                                if (!inMethodErea)
                                    continue;

                                // 呼出関数・モジュールを特定
                                if ((arrWord.Count() >= 2 && arrWord[0] == "PERFORM" && arrWord[1] != "VARYING")
                                        || arrWord[0] == "CALL")
                                {
                                    bool moduleFlg = (arrWord[0] == "CALL") ? true : false;
                                    string name = arrWord[1].Replace("'", "");

                                    if (moduleFlg)
                                        calledModuleList.Add(name);

                                    methodList[methodIndex].CalledMethodList.Add(new CalledMethod(name, moduleFlg, conditions));
                                    continue;
                                }

                                // 関数の終了行を特定
                                if (arrWord[0] == "EXIT" || arrWord[0] == "GOBACK")
                                {
                                    methodList[methodIndex].CalledMethodList = methodList[methodIndex].CalledMethodList.Distinct().ToList();
                                    methodList[methodIndex].EndIndex = fileIndex;
                                    inMethodErea = false;
                                    continue;
                                }

                                // SQL開始行を特定
                                if (String.Join(" ", arrWord) == "EXEC SQL")
                                {
                                    inSqlErea = true;
                                    sql = string.Empty;
                                    sqlType = SqlType.None;
                                    cursorName = string.Empty;
                                    continue;
                                }

                                // =================================================================
                                // SQL内を解析
                                // =================================================================
                                if (inSqlErea)
                                {
                                    inSqlErea = AddSqlList(division, arrWord, methodList[methodIndex].MethodNameP, ref sql, ref sqlType, ref sqlInfoList, ref cursorName);

                                    // カーソルを呼び出している場合、SQL情報クラスに使用関数名をセット
                                    if(arrWord.Length > 1 && arrWord[0] == "FETCH")
                                    {
                                        foreach (var cursor in cursorList)
                                        {
                                            if (cursor.CursorName == arrWord[1])
                                                cursor.UseddMethodName = methodList[methodIndex].MethodNameP;
                                        }
                                    }
                                    continue;
                                }

                                // =================================================================
                                // IF文による条件分岐を解析
                                // =================================================================
                                // IF文による条件分岐を検出
                                if (arrWord[0] == "IF")
                                {
                                    if(inIfConditionDefineErea && !thenAddFlg)
                                    {
                                        // まだスコープ内で分岐条件が条件リストに追加されていないのに次の"IF"句が来た場合、
                                        // "THEN"を省略しているのでここで条件リストに追加する
                                        conditions.Add(ifConditionTxt.Replace("IF", "").Trim());
                                    }
                                    inIfConditionDefineErea = true;
                                    thenAddFlg = false;
                                    ifConditionTxt = string.Empty;
                                }

                                // IF文の分岐条件が2行以上にわたっていた場合
                                if (inIfConditionDefineErea && !arrWord.Contains("THEN")
                                        && !arrWord.Contains("ELSE") && !arrWord.Contains("CONTINUE"))
                                {
                                    ifConditionTxt += " " + String.Join(" ", arrWord);
                                }

                                // ※"THEN"は分岐条件記載部分の下に改行されている前提
                                // ※疑似的な"ELSE IF"文は動作対象外
                                if (arrWord.Contains("THEN"))
                                {
                                    inIfConditionDefineErea = false;
                                    ifConditionTxt += " " + String.Join(" ", arrWord);
                                    conditions.Add(ifConditionTxt.Replace(" THEN", "").Replace("IF", "").Trim());
                                    thenAddFlg = true;
                                    continue;
                                }
                                else if (inIfConditionDefineErea && !thenAddFlg && arrWord.Contains("CONTINUE"))
                                {
                                    // まだスコープ内で分岐条件が条件リストに追加されていないのに"CONTINUE"句が来た場合、
                                    // "THEN"を省略しているのでここで条件リストに追加する
                                    inIfConditionDefineErea = false;
                                    conditions.Add(ifConditionTxt.Replace("IF", "").Trim());
                                    thenAddFlg = true;
                                    continue;
                                }
                                else if (inIfConditionDefineErea && !thenAddFlg && arrWord.Contains("ELSE"))
                                {
                                    // まだスコープ内で分岐条件が条件リストに追加されていないのに"ELSE"句が来た場合、
                                    // "THEN"を省略しているのでここで条件リストに追加する
                                    inIfConditionDefineErea = false;
                                    conditions.Add(ifConditionTxt.Replace("IF", "").Trim() + " 以外");
                                    thenAddFlg = true;
                                    continue;
                                }
                                else if (arrWord.Contains("ELSE"))
                                {
                                    // スコープがELSEに移ったので、条件リストを上書き
                                    conditions[conditions.Count - 1] = conditions.Last() + " 以外";
                                    continue;
                                }

                                if (arrWord.Contains("END-IF"))
                                {
                                    if(inIfConditionDefineErea && !thenAddFlg)
                                    {
                                        // まだスコープ内で分岐条件が条件リストに追加されていないのに"END=IF"句が来た場合、
                                        // "THEN"を省略しているのでここでリセットする
                                        ifConditionTxt = string.Empty;
                                        inIfConditionDefineErea = false;
                                        thenAddFlg = true;
                                        continue;
                                    }
                                    // スコープを外れた分岐条件を条件リストから除外
                                    conditions.RemoveAt(conditions.Count - 1);
                                    continue;
                                }

                                // IF文の分岐条件定義エリア内の場合、これ以降は処理不要のため次の行を読み込む
                                if (inIfConditionDefineErea)
                                    continue;

                                // =================================================================
                                // EVALUATE文による条件分岐を解析
                                // =================================================================
                                // EVALUATE文による条件分岐を検出
                                if (arrWord[0] == "EVALUATE")
                                {
                                    evaluateConditionTxt = "[" + String.Join(" ", arrWord).Replace("EVALUATE ", "") + "]";
                                    whenConditionAddFlg = false;
                                    continue;
                                }

                                // EVALUATE文の分岐条件を抽出
                                // ※1つのWHEN句が2行以上の場合は動作対象外
                                if (arrWord.Contains("WHEN"))
                                {
                                    // WHEN句の内容を取得
                                    whenTxt += arrWord.Contains("OTHER")
                                        ? string.Empty
                                        : "," + String.Join(" ", arrWord).Replace("WHEN", "");
                                    whenTxt = whenTxt.Trim(',').Trim();

                                    // WHEN句が続いていないか確認するため、1行前を取得
                                    string preLine = File.ReadAllLines(file, EncShiftJis).Skip(fileIndex - 2).Take(1).First();
                                    string preWhenConditionTxt = string.Empty;
                                    preLine = FormatLine(preLine, false);
                                    // WHEN句が続いていた場合、分岐条件を取得
                                    if (Left(preLine, 4) == "WHEN")
                                        preWhenConditionTxt = preLine;
                                    // スコープを外れた分岐条件を条件リストから除外
                                    if (whenConditionAddFlg)
                                        conditions.RemoveAt(conditions.Count - 1);

                                    string whenConditionTxt =
                                        String.IsNullOrEmpty(preWhenConditionTxt)
                                        ? "[" + String.Join(" ", arrWord) + "]"
                                        : "[" + preWhenConditionTxt + "]" + " または " + "[" + String.Join(" ", arrWord) + "]";

                                    string otherConditionTxt = "[" + whenTxt + "]" + " 以外";
                                    if (whenConditionTxt.Contains("OTHER"))
                                        conditions.Add(evaluateConditionTxt + " = " + otherConditionTxt);
                                    else
                                        conditions.Add(evaluateConditionTxt + " = " + whenConditionTxt.Replace("WHEN ", ""));

                                    whenConditionAddFlg = true;
                                    continue;
                                }

                                if (arrWord[0].Contains("END-EVALUATE"))
                                {
                                    // スコープを外れた分岐条件を条件リストから除外
                                    conditions.RemoveAt(conditions.Count - 1);
                                    whenTxt = string.Empty;
                                    continue;
                                }

                                continue;
                            default:
                                continue;
                        }
                    }
                }

                // モジュール内に関数がない場合は処理終了
                if (methodList.Count <= 0)
                {
                    _logger.Error("ファイル内に関数が一つもありません。");
                    return false;
                }
                #endregion

                // =================================================================
                // 前処理②（未使用の関数を特定）
                // =================================================================
                #region 前処理②
                // 関数名がリスト内のその他の関数から呼び出されているかチェックする
                foreach (var method1 in methodList)
                {
                    int index = methodList.IndexOf(method1);
                    if (index == 0) { method1.CalledFlg = true; }
                    foreach (var method2 in methodList)
                    {
                        if (index == methodList.IndexOf(method2)) { continue; }

                        foreach (var cm in method2.CalledMethodList)
                        {
                            if (method1.MethodNameP == cm.Name)
                            {
                                method1.CalledFlg = true;

                                // 呼出関数の関数リスト内でのindexをセット
                                cm.MethodListIndex = index;
                                continue;
                            }
                        }
                    }
                }

                // 未使用関数から呼び出される関数のリストを作成
                List<Method> calledMethodList = new List<Method>();
                foreach (var method1 in methodList)
                {
                    // 呼出フラグがfalseかつ、呼出先がある関数が対象
                    if (method1.CalledFlg) { continue; }
                    if (method1.CalledMethodList.Count < 1) { continue; }

                    foreach (var cm in method1.CalledMethodList)
                    {
                        if (cm.ModuleFlg)
                            continue;

                        calledMethodList.Add(methodList[cm.MethodListIndex]);
                    }
                }

                // 呼出先リストにある関数が、呼出フラグfalseの関数からのみ呼び出されていれば呼出フラグfalseにする
                foreach (var method1 in calledMethodList)
                {
                    // 呼出フラグをいったんfalseに
                    method1.CalledFlg = false;
                    int index = methodList.IndexOf(method1);
                    foreach (var method2 in methodList)
                    {
                        if (index == methodList.IndexOf(method2)) { continue; }
                        foreach (var cm in method2.CalledMethodList)
                        {
                            if (method1.MethodNameP == cm.Name && method2.CalledFlg)
                            {
                                // 呼出元関数の呼出フラグが一つでもtrueなら呼出先関数もtrueとする
                                method1.CalledFlg = true;
                                break;
                            }
                        }
                        if (method1.CalledFlg) { break; }
                    }
                }
                #endregion

                // =================================================================
                // Excelファイル作成
                // =================================================================
                // 出力ファイル準備（実行ファイルと同じフォルダに出力される）
                FileInfo newFile = new FileInfo(ResultFileName);

                // Excelファイル作成
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    bool ret = false;

                    // プログラム情報シート作成・編集（シート名：PGM情報_{読込ファイル名}）
                    WsPgmInfo = AddSheet(package, CommonConst.SHEET_NAME_PGMINFO, file);
                    ret = EditPgmInfoSheet(copyList, sqlInfoList, cursorList, calledModuleList);

                    if (!ret) { return false; }

                    // 関数情報シート作成・編集（シート名：関数情報_{読込ファイル名}）
                    WsMethodInfo = AddSheet(package, CommonConst.SHEET_NAME_METHODINFO, file);
                    ret = EditMethodInfoSheet(methodList, sqlInfoList, cursorList);

                    if (!ret) { return false; }

                    // 構造図シート作成・編集（シート名：構造図_{読込ファイル名}）
                    WsStruct = AddSheet(package, CommonConst.SHEET_NAME_STRUCT, file);
                    ret = EditStructSheet(methodList);

                    if (!ret) { return false; }

                    // 保存
                    package.Save();
                }
                return true;
            }
            catch (Exception ex)
            {
                _logger.Fatal(ex.Message + Environment.NewLine
                    + "処理セクション：[" + logSectionName + "]" + Environment.NewLine
                    + "処理行：" + logLine);
                return false;
            }
        }

        #region 前処理
        /// <summary>
        /// テキストを整形する
        /// </summary>
        /// <param name="line">指定行のテキスト</param>
        /// <param name="exceptComFlg">整形時に標準領域を除くフラグ</param>
        /// <returns></returns>
        private static string FormatLine(string line, bool exceptComFlg)
        {
            int start = exceptComFlg ? 8 : 7;
            // 一連番号領域、プログラム識別領域を取り除く（"000001  hoge   fuga " ⇒ "hoge   fuga "）
            string frmLine = Mid(line, start, 73 - start);
            // テキスト前後のスペースを詰め、テキスト内にある複数のスペースを取り除く（"hoge   fuga" ⇒ "hoge fuga"）
            frmLine = Regex.Replace(frmLine.Trim(), @"\s{2,}", " ");
            // コメント行の見出し領域より手前に修正履歴などがある場合、「A.」「D.」以降を取り除く
            if (exceptComFlg)
            {
                int index = -1;
                index = frmLine.IndexOf("A.");
                index = (index < 0) ? frmLine.IndexOf("D.") : index;
                frmLine = (index < 0) ? frmLine : Left(frmLine, index);
            }

            // 予約語の置換
            frmLine = ReplaceReservedWord(frmLine);

            return frmLine;
        }

        /// <summary>
        /// プログラム部の判定
        /// </summary>
        /// <param name="arrWord"></param>
        /// <returns></returns>
        private static Division CheckDivisionChanged(string[] arrWord)
        {
            if (arrWord.Length < 2 || arrWord[1].Replace(".", "") != CommonConst.WORD_DIVISION)
                return Division.NONE;

            switch (arrWord[0])
            {
                // 見出し部
                case CommonConst.WORD_IDENTIFICATION:
                    return Division.IDENTIFICATION;
                // 環境部
                case CommonConst.WORD_ENVIRONMENT:
                    return Division.ENVIRONMENT;
                // データ部
                case CommonConst.WORD_DATA:
                    return Division.DATA;
                // 手続き部
                case CommonConst.WORD_PROCEDURE:
                    return Division.PROCEDURE;
                default:
                    return Division.NONE;
            }
        }

        private static string ReplaceReservedWord(string line)
        {
            string[] arrWord = line.Split(' ');
            int i = -1;
            foreach (string word in arrWord)
            {
                i++;
                string replaceWord = word;
                if (word == "ZERO")
                    replaceWord = "0";
                if (word == "SPACE")
                    replaceWord = "''";
                if (word == "ALSO")
                    replaceWord = "と";

                arrWord[i] = replaceWord;
            }

            string retLine = String.Join(" ", arrWord);
            return retLine;
        }

        /// <summary>
        /// 対象外句の判定
        /// </summary>
        /// <param name="arrWord"></param>
        /// <param name="division"></param>
        /// <returns></returns>
        private static bool CheckExcludedWords(string[] arrWord, Division division)
        {
            string checkText = arrWord[0].Replace(".", "");

            // 共通
            if (checkText == string.Empty || Left(checkText, 1) == CommonConst.COM_PREFIX)
                return false;

            switch (division)
            {
                // 見出し部
                case Division.IDENTIFICATION:
                    if (RegexIdentification.IsMatch(checkText))
                        return false;
                    break;
                // 環境部
                case Division.ENVIRONMENT:
                    if (RegexEnvironment.IsMatch(checkText))
                        return false;
                    break;
                // データ部
                case Division.DATA:
                    if (RegexData.IsMatch(checkText))
                        return false;
                    break;
                // 手続き部
                case Division.PROCEDURE:
                    if (checkText == "DISPLAY")
                        return false;
                    break;
                default:
                    break;
            }

            return true;

        }

        /// <summary>
        /// コメント行の取得
        /// </summary>
        /// <param name="file"></param>
        /// <param name="comIndex"></param>
        /// <returns></returns>
        private static string GetComment(string file, int comIndex)
        {
            string comLine = File.ReadAllLines(file, EncShiftJis).Skip(comIndex).Take(1).First();
            if (Mid(comLine, 7, 1) == CommonConst.COM_PREFIX)
                return FormatLine(comLine, true);
            else
                return string.Empty;
        }

        private static bool AddSqlList(Division division , string[] arrWord , string usedMethodName
                , ref string sql, ref SqlType sqlType, ref List<SqlInfo> sqlList, ref string cursorName)
        {
            SqlInfo _sqlInfo = new SqlInfo();

            // カーソル物理名を特定
            if(division == Division.DATA && arrWord.Length > 1 && arrWord[0] == "DECLARE")
                cursorName = arrWord[1];

            // SQLの処理区分を特定
            sqlType = _sqlInfo.StringToSqlType(arrWord[0], sqlType);

            // SQL終了行を特定
            if (arrWord[0].Replace(".", "") == "END-EXEC")
            {
                IEnumerable<TokenInfo> tokens;
                if (!string.IsNullOrEmpty(sql))
                {
                    tokens = TransactSqlHelpers.Parser.ParseSql(sql);
                    SqlInfo sqlInfo = new SqlInfo(sql, tokens, sqlType, usedMethodName, cursorName);
                    sqlList.Add(sqlInfo);
                }

                return false;
            }

            // SQL文の取得
            StringBuilder str = new StringBuilder();
            str.Append(sql);
            foreach (string val in arrWord)
            {
                str.Append(val + " ");
            }
            sql = str.ToString();
            return true;
        }
        #endregion

        // todo
        #region SQL情報描画
        private static bool EditSqlInfoSheet(IEnumerable<SqlInfo> sqlInfoList)
        {
            string errorMsg = "(" + SheetName + "シート作成時エラー)";
            if (WsPgmInfo == null)
            {
                _logger.Fatal(errorMsg + "Excelシート変数に値が割り当てられませんでした。");
                return false;
            }

            try
            {
                foreach (SqlInfo sqlInfo in sqlInfoList)
                {
                    string[] sql = sqlInfo.Value.Split(' ');
                    StringBuilder sb = new StringBuilder();
                    int index = 0;

                    switch (sqlInfo.Type)
                    {
                        case SqlType.Select:
                            foreach (TokenInfo token in sqlInfo.TokenList)
                            {
                                sb.Append(sql[index]);

                            }
                            break;
                        case SqlType.Insert:
                            break;
                        case SqlType.Update:
                            break;
                        case SqlType.Delete:
                            break;
                        case SqlType.Create:
                            break;
                        default:
                            break;
                    }
                }

            }
            catch (Exception ex)
            {
                _logger.Fatal(errorMsg + ex.Message);
                return false;
            }

            return true;
        }
        #endregion
        
        #region PGM情報描画
        /// <summary>
        /// PGM情報シート編集
        /// </summary>
        /// <param name="copyList"></param>
        /// <param name="sqlInfoList"></param>
        /// <param name="calledModuleList"></param>
        /// <returns></returns>
        private static bool EditPgmInfoSheet(IReadOnlyDictionary<string, string> copyList
                                    , IEnumerable<SqlInfo> sqlInfoList
                                    , IEnumerable<SqlInfo> cursorList
                                    , IEnumerable<string> calledModuleList)
        {
            string errorMsg = "(" + SheetName + "シート作成時エラー)";
            if (WsPgmInfo == null)
            {
                _logger.Fatal(errorMsg + "Excelシート変数に値が割り当てられませんでした。");
                return false;
            }

            try
            {
                int row = 2; // 行番号

                // =================================================================
                // 各項目のタイトル部分を書き込む
                // =================================================================
                // コピー句のタイトルセット
                SetStyleOfTitle(CommonConst.SHEET_NAME_PGMINFO, "B2:C2", Color.SpringGreen);
                WsPgmInfo.Cells[row, 2].Value = "COPY句";
                // 呼出モジュールのタイトルセット
                SetStyleOfTitle(CommonConst.SHEET_NAME_PGMINFO, "D2:D2", Color.Pink);
                WsPgmInfo.Cells[row, 4].Value = "呼出モジュール";
                // カーソルのタイトルセット
                SetStyleOfTitle(CommonConst.SHEET_NAME_PGMINFO, "E2:F2", Color.Bisque);
                WsPgmInfo.Cells[row, 5].Value = "カーソル";
                WsPgmInfo.Cells[row, 6].Value = "[使用DB]";
                WsPgmInfo.Cells[row, 6].Style.Font.Size = 9;
                // SQL変数宣言部のタイトルセット
                SetStyleOfTitle(CommonConst.SHEET_NAME_PGMINFO, "G2:M2", Color.LightSteelBlue);
                WsPgmInfo.Cells[row, 7].Value = "DB情報";
                WsPgmInfo.Cells[row, 9].Value = "[SELECT]";
                WsPgmInfo.Cells[row, 10].Value = "[INSERT]";
                WsPgmInfo.Cells[row, 11].Value = "[UPDATE]";
                WsPgmInfo.Cells[row, 12].Value = "[DELETE]";
                WsPgmInfo.Cells[row, 13].Value = "[CREATE]";
                WsPgmInfo.Cells[row, 9, row, 13].Style.Font.Size = 9;

                // =================================================================
                // コピー句リストを書き込む
                // =================================================================
                foreach (var copy in copyList)
                {
                    row++;
                    WsPgmInfo.Cells[row, 2].Value = copy.Key;
                    WsPgmInfo.Cells[row, 3].Value = copy.Value;
                }

                // =================================================================
                // 呼出モジュールリストを書き込む
                // =================================================================
                row = 2;
                foreach (var module in calledModuleList)
                {
                    row++;
                    WsPgmInfo.Cells[row, 4].Value = module;
                }

                // =================================================================
                // カーソルリストを書き込む
                // =================================================================
                // テーブルの論理名取得用にDB定義一覧を取得
                DataTable dt = GetDbDefine();

                DbInfo _dbInfo;
                List<DbInfo> dbInfoList = new List<DbInfo>();

                row = 2;
                foreach (var cursor in cursorList)
                {
                    row++;
                    WsPgmInfo.Cells[row, 5].Value = cursor.CursorName; // カーソル物理名

                    // カーソル内で使用されているDBリストを取得
                    IEnumerable<string> dbList = cursor.GetDbList();
                    string dbTxt = string.Empty;
                    foreach (var dbName in dbList)
                    {
                        dbTxt += dbName + ",";
                        int i = dbInfoList.FindIndex(x => x.Name_P == dbName);
                        if (i < 0)
                        {
                            // 使用DBリストにためておく
                            _dbInfo = new DbInfo(dbName, SqlType.Select, dt);
                            dbInfoList.Add(_dbInfo);
                        }
                    }
                    WsPgmInfo.Cells[row, 6].Value = dbTxt.TrimEnd(','); // カーソル内使用DB一覧
                }

                // =================================================================
                // 使用DBリストを書き込む
                // =================================================================
                foreach (var sqlInfo in sqlInfoList)
                {
                    // SQL内で使用されているDBリストを取得
                    IEnumerable<string> dbList = sqlInfo.GetDbList();
                    // DBの使用されているCRUDをセット
                    foreach (string dbName in dbList)
                    {
                        int i = dbInfoList.FindIndex(x => x.Name_P == dbName);
                        if(i < 0)
                        {
                            // 使用DBリストにためておく
                            _dbInfo = new DbInfo(dbName, sqlInfo.Type, dt);
                            dbInfoList.Add(_dbInfo);
                        }
                        else
                        {
                            dbInfoList[i].SetCrudFlg(sqlInfo.Type);
                        }
                    }
                }
                IEnumerable<DbInfo> distinctList = dbInfoList.Distinct().OrderBy(x => x.Name_P);

                row = 2;
                foreach (var dbInfo in distinctList)
                {
                    row++;
                    WsPgmInfo.Cells[row, 7].Value = dbInfo.Name_P; // DB物理名
                    WsPgmInfo.Cells[row, 8].Value = dbInfo.Name_L; // DB論理名
                    WsPgmInfo.Cells[row, 9].Value = dbInfo.SelectFlg ? "〇" : string.Empty; // SELECT
                    WsPgmInfo.Cells[row, 10].Value = dbInfo.InsertFlg ? "〇" : string.Empty; // INSERT
                    WsPgmInfo.Cells[row, 11].Value = dbInfo.UpdateFlg ? "〇" : string.Empty; // UPDATE
                    WsPgmInfo.Cells[row, 12].Value = dbInfo.DeleteFlg ? "〇" : string.Empty; // DELETE
                    WsPgmInfo.Cells[row, 13].Value = dbInfo.CreateFlg ? "〇" : string.Empty; // CREATE
                }

                WsPgmInfo.Cells.Style.Font.Name = CommonConst.FONT_NAME_MEIRYOUI;
                WsPgmInfo.Cells[WsPgmInfo.Dimension.Address].AutoFitColumns(); // 列幅自動調整
            }
            catch (Exception ex)
            {
                _logger.Fatal(errorMsg + ex.Message);
                return false;
            }

            return true;
        }

        /// <summary>
        /// DB定義一覧を取得
        /// </summary>
        /// <returns></returns>
        private static DataTable GetDbDefine()
        {
            // データテーブルを作成
            DataTable dt = new DataTable();
            dt.Columns.Add(CommonConst.TABLE_P);
            dt.Columns.Add(CommonConst.TABLE_L);
            dt.Columns.Add(CommonConst.COLUMN_P);
            dt.Columns.Add(CommonConst.COLUMN_L);

            // DB定義一覧ファイルの存在チェック（なくても処理は止めない）
            if (!File.Exists(DbDifineFileName))
            {
                _logger.Error($"{DbDifineFileName}は存在しません。");
                return dt;
            }

            // DB定義一覧を取得
            // ※"テーブル物理名"、"テーブル論理名"、"カラム物理名"、"カラム論理名"　がカンマ区切りで並んでいるCSVファイルの想定
            using (StreamReader sr = new StreamReader(DbDifineFileName, EncUtf8))
            {
                // 読み込んだ行をデータテーブルにセット
                while (sr.Peek() >= 0)
                {
                    string[] line = sr.ReadLine().Split(',');
                    if (line.Length != 4)
                        continue;
                    
                    DataRow dr = dt.NewRow();
                    dr[CommonConst.TABLE_P] = line[0];
                    dr[CommonConst.TABLE_L] = line[1];
                    dr[CommonConst.COLUMN_P] = line[2];
                    dr[CommonConst.COLUMN_L] = line[3];
                    dt.Rows.Add(dr);
                }
            }

            return dt;
        }
        #endregion

        #region 関数情報描画
        /// <summary>
        /// 関数情報シート編集
        /// </summary>
        /// <param name="methodList"></param>
        /// <returns></returns>
        private static bool EditMethodInfoSheet(IEnumerable<Method> methodList, IReadOnlyCollection<SqlInfo> sqlInfoList, IReadOnlyCollection<SqlInfo> cursorList)
        {
            string errorMsg = "(" + SheetName + "シート作成時エラー)";
            if (WsMethodInfo == null)
            {
                _logger.Fatal(errorMsg + "Excelシート変数に値が割り当てられませんでした。");
                return false;
            }

            try
            {
                int row = 2; // 行番号
                int col = 2; // 列番号

                // 各項目のタイトル部分を書き込む
                SetStyleOfTitle(CommonConst.SHEET_NAME_METHODINFO, "B2:Z2", Color.SpringGreen);
                WsMethodInfo.Cells[row, col].Value = "関数名(物理)";
                WsMethodInfo.Cells[row, col + 1].Value = "関数名(論理)";
                WsMethodInfo.Cells[row, col + 2].Value = "開始行数";
                WsMethodInfo.Cells[row, col + 3].Value = "終了行数";
                WsMethodInfo.Cells[row, col + 4].Value = "DB操作";

                int calledMethodCol = 7;
                for (int i = 1; i <= 20; i++)
                {
                    WsMethodInfo.Cells[row, calledMethodCol].Value = "呼出関数" + i.ToString();
                    calledMethodCol++;
                }

                // 関数リストを書き込む
                foreach (var method in methodList)
                {
                    calledMethodCol = 7;
                    row++;

                    // SQLを呼んでいるかチェック
                    List<string> sqlTypeList = new List<string>();
                    foreach (var sqlInfo in sqlInfoList)
                    {
                        if (method.MethodNameP == sqlInfo.UseddMethodName)
                        {
                            string sqlType = sqlInfo.SqlTypeToString();
                            sqlTypeList.Add(sqlType);
                        }
                    }
                    IEnumerable<string> distinctList = sqlTypeList.Distinct();
                    string sqlTypeString = String.Join(",", distinctList);

                    // カーソルを呼んでいるかチェック
                    foreach (var cursor in cursorList)
                    {
                        if(method.MethodNameP == cursor.UseddMethodName)
                            sqlTypeString += ",SELECT [" + cursor.CursorName + "]";
                    }

                    // 関数物理名
                    WsMethodInfo.Cells[row, col].Value = method.MethodNameP;
                    // 関数論理名
                    WsMethodInfo.Cells[row, col + 1].Value = method.MethodNameL;
                    // 開始行数
                    WsMethodInfo.Cells[row, col + 2].Value = method.StartIndex;
                    // 終了行数
                    WsMethodInfo.Cells[row, col + 3].Value = method.EndIndex;
                    // DB操作
                    WsMethodInfo.Cells[row, col + 4].Value = sqlTypeString.Trim(',');
                    // 呼出関数
                    foreach (var cm in method.CalledMethodList)
                    {
                        Color outColor = cm.ModuleFlg ? ColorModule : ColorMethod;
                        WsMethodInfo.Cells[row, calledMethodCol].Value = cm.Name;
                        WsMethodInfo.Cells[row, calledMethodCol].Style.Font.Color.SetColor(outColor);
                        calledMethodCol++;
                    }

                    // 到達不能な関数の場合
                    if (!method.CalledFlg)
                    {
                        // コメントをセット
                        WsMethodInfo.Cells[row, col].AddComment("到達不能な関数です。", "system");
                        // 背景色変更
                        WsMethodInfo.Cells[row, col, row, col + 24].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        WsMethodInfo.Cells[row, col, row, col + 24].Style.Fill.BackgroundColor.SetColor(Color.Gray);
                    }

                }

                WsMethodInfo.Cells.Style.Font.Name = CommonConst.FONT_NAME_MEIRYOUI;
                WsMethodInfo.Cells[WsMethodInfo.Dimension.Address].AutoFitColumns(); // 列幅自動調整
            }
            catch (Exception ex)
            {
                _logger.Fatal(errorMsg + ex.Message);
                return false;
            }

            return true;
        }
        #endregion

        #region 構造図描画
        /// <summary>
        /// 構造図シート編集
        /// </summary>
        /// <param name="methodList"></param>
        /// <returns></returns>
        private static bool EditStructSheet(List<Method> methodList)
        {
            string errorMsg = "(" + SheetName + "シート作成時エラー)";
            if (WsStruct == null)
            {
                _logger.Fatal(errorMsg + "Excelシート変数に値が割り当てられませんでした。");
                return false;
            }

            try
            {
                int row = 2; // 行番号
                int col = 2; // 列番号

                // セルを方眼紙にする
                WsStruct.DefaultColWidth = 3;

                // 起点となる関数名をExcelに書き込む
                WriteMethod(methodList, 0, ref row, ref col, new CalledMethod(string.Empty, false, new List<string>()));

                // 呼び出される関数名を再帰的にExcelに書き込む
                GetCalledMethodRecursively(methodList, 0, row + 1, col + 2);

                WsStruct.Cells.Style.Font.Name = CommonConst.FONT_NAME_MEIRYOUI;
            }
            catch (Exception ex)
            {
                _logger.Fatal(errorMsg + ex.Message);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Excelに関数名を書き込む
        /// </summary>
        /// <param name="methodList"></param>
        /// <param name="index"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="cm"></param>
        private static void WriteMethod(List<Method> methodList, int index, ref int row, ref int col, CalledMethod cm)
        {
            string outText = string.Empty;
            Color outColor = ColorMethod;
            bool flg = false;
            WsStruct.Cells[row, col].Style.Numberformat.Format = "@";
            if (cm.ModuleFlg)
            {
                // モジュール呼出の場合
                outText = cm.Name;
                outColor = ColorModule;
            }
            else if (index < 0 || index > methodList.Count)
            {
                outText = "不正な関数が指定されました。(" + cm.Name + ")";
                outColor = Color.Red;
            }
            else
            {
                // 関数呼出の場合

                if (!String.IsNullOrEmpty(cm.Conditions))
                {
                    // 条件分岐がある場合
                    WsStruct.Cells[row, col].Value = cm.Conditions;
                    WsStruct.Cells[row, col].Style.Font.Size = 9;
                    WsStruct.Cells[row + 1, col].Value = "└";
                    row++;
                    col++;
                    flg = true;
                }

                // ハイパーリンクの設定
                string linkSheetName = SheetName.Replace(CommonConst.SHEET_NAME_STRUCT, CommonConst.SHEET_NAME_METHODINFO);
                outText = methodList[index].MethodNameP;
                WsStruct.Cells[row, col].Hyperlink
                    = new ExcelHyperLink("#'" + linkSheetName + "'!B" + (index + 3).ToString(), outText);
                WsStruct.Cells[row, col].Style.Font.UnderLine = true;
            }

            WsStruct.Cells[row, col].Value = outText;
            WsStruct.Cells[row, col].Style.Font.Color.SetColor(outColor);
            if (flg)
                col--;
        }

        /// <summary>
        /// 再帰的に呼出関数を取得する
        /// </summary>
        /// <param name="methodList"></param>
        /// <param name="index"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        private static void GetCalledMethodRecursively(List<Method> methodList, int index, int row, int col)
        {
            // 関数内から呼び出している他関数・モジュールを読み込む
            foreach (var calledMethod in methodList[index].CalledMethodList)
            {
                // 呼出関数名の、関数リスト内での位置を取得
                int calledMethodIndex = GetMethodIndex(methodList, calledMethod);

                // Excelに書き込む
                WriteMethod(methodList, calledMethodIndex, ref row, ref col, calledMethod);

                row++;
                InitRow = row;

                if (calledMethodIndex > -1)
                {
                    // 呼出先の関数内から呼び出している関数がないか、再帰的に探索する
                    GetCalledMethodRecursively(methodList, calledMethodIndex, row, col + 2);
                }
                row = InitRow;
            }
        }

        /// <summary>
        /// 指定された関数名の関数リスト内での位置を返す
        /// </summary>
        /// <param name="methodList"></param>
        /// <param name="method"></param>
        /// <returns></returns>
        private static int GetMethodIndex(IReadOnlyCollection<Method> methodList, CalledMethod cm)
        {
            if (cm.ModuleFlg)
                return -1;

            int i = 0;
            foreach (var method in methodList)
            {
                if (cm.Name == method.MethodNameP)
                    return i;

                i++;
            }
            return -1;
        }
        #endregion

        #region 共通
        private static ExcelWorksheet AddSheet(ExcelPackage package, string sheetType, string file)
        {
            SheetName = sheetType + "_" + Path.GetFileNameWithoutExtension(file);
            if (package.Workbook.Worksheets[SheetName] != null)
                package.Workbook.Worksheets.Delete(SheetName);

            return package.Workbook.Worksheets.Add(SheetName);

        }

        private static void SetStyleOfTitle(string sheetName, string range, Color titleColor)
        {
            switch (sheetName)
            {
                case CommonConst.SHEET_NAME_PGMINFO:
                    WsPgmInfo.Cells[range].Style.Font.Bold = true;
                    WsPgmInfo.Cells[range].Style.Fill.PatternType = ExcelFillStyle.DarkVertical;
                    WsPgmInfo.Cells[range].Style.Fill.BackgroundColor.SetColor(titleColor);
                    break;
                case CommonConst.SHEET_NAME_METHODINFO:
                    WsMethodInfo.Cells[range].Style.Font.Bold = true;
                    WsMethodInfo.Cells[range].Style.Fill.PatternType = ExcelFillStyle.DarkVertical;
                    WsMethodInfo.Cells[range].Style.Fill.BackgroundColor.SetColor(titleColor);
                    break;
                case CommonConst.SHEET_NAME_STRUCT:
                    WsStruct.Cells[range].Style.Font.Bold = true;
                    WsStruct.Cells[range].Style.Fill.PatternType = ExcelFillStyle.DarkVertical;
                    WsStruct.Cells[range].Style.Fill.BackgroundColor.SetColor(titleColor);
                    break;
                default:
                    break;
            }
        }

        private static string GetArrayWord(string[] value, int index)
        {
            if (index >= value.Length)
                return string.Empty;
            else
                return value[index];
        }

        /// <summary>
        /// 数値をExcelのカラム文字へ変換します
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        public static string ToColumnName(int source)
        {
            if (source < 1) return string.Empty;
            return ToColumnName((source - 1) / 26) + (char)('A' + ((source - 1) % 26));
        }

        /// <summary>
        /// 文字列の指定した位置から指定した長さを取得する
        /// </summary>
        /// <param name="str">文字列</param>
        /// <param name="start">開始位置</param>
        /// <param name="len">長さ</param>
        /// <returns>取得した文字列</returns>
        public static string Mid(string str, int start, int len)
        {
            if (start <= 0)
            {
                throw new ArgumentException("引数'start'は1以上でなければなりません。");
            }
            if (len < 0)
            {
                return "";
            }
            if (str == null || str.Length < start)
            {
                return "";
            }
            if (str.Length < (start + len))
            {
                return str.Substring(start - 1);
            }
            return str.Substring(start - 1, len);
        }

        /// <summary>
        /// 文字列の指定した位置から末尾までを取得する
        /// </summary>
        /// <param name="str">文字列</param>
        /// <param name="start">開始位置</param>
        /// <returns>取得した文字列</returns>
        public static string Mid(string str, int start)
        {
            return Mid(str, start, str.Length);
        }

        /// <summary>
        /// 文字列の先頭から指定した長さの文字列を取得する
        /// </summary>
        /// <param name="str">文字列</param>
        /// <param name="len">長さ</param>
        /// <returns>取得した文字列</returns>
        public static string Left(string str, int len)
        {
            if (len < 0)
            {
                throw new ArgumentException("引数'len'は0以上でなければなりません。");
            }
            if (str == null)
            {
                return "";
            }
            if (str.Length <= len)
            {
                return str;
            }
            return str.Substring(0, len);
        }

        /// <summary>
        /// 文字列の末尾から指定した長さの文字列を取得する
        /// </summary>
        /// <param name="str">文字列</param>
        /// <param name="len">長さ</param>
        /// <returns>取得した文字列</returns>
        public static string Right(string str, int len)
        {
            if (len < 0)
            {
                throw new ArgumentException("引数'len'は0以上でなければなりません。");
            }
            if (str == null)
            {
                return "";
            }
            if (str.Length <= len)
            {
                return str;
            }
            return str.Substring(str.Length - len, len);
        }
        #endregion
    }

    /// <summary>
    /// 関数管理クラス
    /// </summary>
    public class Method
    {
        public string MethodNameP { get; }
        public string MethodNameL { get; }
        public int StartIndex { get; }
        public int EndIndex { get; internal set; }
        public List<CalledMethod> CalledMethodList { get; internal set; }
        public bool CalledFlg { get; internal set; }

        public Method(string methodNameP, string methodNameL, int startIndex, int endIndex)
        {
            MethodNameP = methodNameP;
            MethodNameL = methodNameL;
            StartIndex = startIndex;
            EndIndex = endIndex;
            CalledMethodList = new List<CalledMethod>();
            CalledFlg = false;
        }
    }

    /// <summary>
    /// 呼出関数管理クラス　※下記参照の上、distinct用の実装箇所あり
    /// 参照：https://qiita.com/Chrowa3/items/51e7033aa687c6274ad4
    /// 参照：https://docs.microsoft.com/ja-jp/dotnet/api/system.linq.enumerable.distinct?redirectedfrom=MSDN&view=netcore-3.1#System_Linq_Enumerable_Distinct__1_System_Collections_Generic_IEnumerable___0__
    /// </summary>
    public class CalledMethod : IEquatable<CalledMethod>
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
                if(sb.Length > 0)
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

    /// <summary>
    /// SQL情報クラス
    /// </summary>
    public class SqlInfo
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

                if(token.Token == Tokens.TOKEN_INSERT || token.Token == Tokens.TOKEN_UPDATE || token.Token == Tokens.TOKEN_CREATE)
                {
                    dbAddFlg2 = true;
                    continue;
                }
                if(dbAddFlg2 && token.Token == Tokens.TOKEN_ID)
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

    /// <summary>
    /// DB情報クラス
    /// </summary>
    public class DbInfo : IEquatable<DbInfo>
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

public enum SqlType : int
{
    None = 0,
    Select = 1,
    Insert = 2,
    Update = 3,
    Delete = 4,
    Create = 5
}