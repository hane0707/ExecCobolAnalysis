using log4net;
using Microsoft.SqlServer.Management.SqlParser.Parser;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Configuration;
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
        #region 定数
        const string COM_PREFIX = "*";
        const string METHOD_START = "SECTION";
        const string METHOD_END = "-999";
        const string FONT_NAME = "Meiryo UI";
        const string SHEET_NAME_PGMINFO = "PGM情報";
        const string SHEET_NAME_METHODINFO = "関数情報";
        const string SHEET_NAME_STRUCT = "構造図";
        const string FLG_ON = "1";
        const string FLG_OFF = "0";
        const int RETURN_OK = 0;
        const int RETURN_ERR_100 = 100;
        const int RETURN_ERR_200 = 200;
        #endregion

        #region 変数
        static string OutFileName =
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + ConfigurationManager.AppSettings["ResultFilePath"];
        static Encoding Enc = Encoding.GetEncoding("Shift_JIS");
        static int InitRow = 2;
        static string SheetName = string.Empty;
        static ExcelWorksheet WsPgmInfo; // プログラム情報シート
        static ExcelWorksheet WsStruct; // 構造図シート
        static ExcelWorksheet WsMethodInfo; // 関数情報シート
        static Color ColorMethod = Color.RoyalBlue;
        static Color ColorModule = Color.DeepPink;
        private static readonly ILog _logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        #endregion
        private static int Main(string[] args)
        {
            // 出力ファイルが使用中ではないかチェック
            try
            {
                if (File.Exists(OutFileName))
                {
                    using (Stream st = new FileStream(OutFileName, FileMode.Open)) { }
                }
            }
            catch (Exception)
            {
                _logger.Error($"出力ファイル({OutFileName})が使用中です。");
                return RETURN_ERR_200;
            }

            // 解析対象ファイルが読み込み可能かチェック
            if (args.Length < 1)
            {
                _logger.Error("解析対象のファイルが指定されていません。");
                return RETURN_ERR_100;
            }
            string file = args[0];

            if (!File.Exists(file))
            {
                _logger.Error($"{file}は存在しません。");
                return RETURN_ERR_100;
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

            return RETURN_OK;
        }

        private static bool Exec(string file)
        {
            try
            {
                List<Method> methodList = new List<Method>();
                Dictionary<string, string> copyList = new Dictionary<string, string>();
                List<string> calledModuleList = new List<string>();
                List<IEnumerable<TokenInfo>> SqlTokenList = new List<IEnumerable<TokenInfo>>();
                List<SqlInfo> sqlInfoList = new List<SqlInfo>();

                // =================================================================
                // 前処理①（関数の一覧をリスト化）
                // =================================================================
                #region 前処理①
                using (StreamReader sr = new StreamReader(file, Enc))
                {
                    int fileIndex = 0;
                    int methodIndex = -1;
                    bool inMethodErea = false;
                    bool inSqlErea = false;
                    string sql = string.Empty;
                    SqlType sqlType = SqlType.None;

                    while (sr.Peek() >= 0)
                    {
                        fileIndex++;

                        // 読み込んだ行を整形する
                        string fmtLine = FormatLine(sr.ReadLine(), false);
                        if (!inSqlErea)
                            fmtLine = fmtLine.Replace(".", "");

                        // テキストをスペースで区切った配列を作成
                        string[] arrWord = fmtLine.Split(' ');

                        // 対象外句はスルー
                        if (!CheckLine(arrWord))
                        {
                            continue;
                        }

                        // コピー句を特定（特定行の1行手前がコメント行だと仮定する）
                        if (arrWord[0] == "COPY")
                        {
                            string copyKey = GetArrayWord(arrWord, 1);
                            string comLine = GetComment(file, fileIndex - 2);
                            if (!copyList.ContainsKey(copyKey))
                                copyList.Add(copyKey, comLine);
                            continue;
                        }

                        // 関数の開始行を特定
                        if (arrWord[arrWord.Length - 1] == METHOD_START)
                        {
                            methodIndex++;
                            Method m = new Method(arrWord[0], fileIndex, -1);
                            methodList.Add(m);
                            inMethodErea = true;

                            // 関数名の論理名の特定（関数開始行の2行手前がコメント行だと仮定する）
                            // ※関数開始行の3行前までを読み飛ばし、次の1行（＝コメント行）を読み込む
                            methodList[methodIndex].MethodNameL = GetComment(file, fileIndex - 3);
                            continue;
                        }

                        if (inMethodErea)
                        {
                            // 呼出関数・モジュールを特定
                            if ((arrWord[0] == "PERFORM" && arrWord[1] != "VARYING") || arrWord[0] == "CALL")
                            {
                                bool moduleFlg = (arrWord[0] == "CALL") ? true : false;
                                string name = arrWord[1].Replace("'", "");

                                if (moduleFlg)
                                    calledModuleList.Add(name);

                                methodList[methodIndex].CalledMethod.Add(new CalledMethod(name, moduleFlg));
                                continue;
                            }

                            // SQL開始行を特定
                            if (String.Join(" ", arrWord) == "EXEC SQL")
                            {
                                inSqlErea = true;
                                sql = string.Empty;
                                sqlType = SqlType.None;
                                continue;
                            }

                            if (inSqlErea)
                            {
                                // SQLの処理区分を特定
                                switch (arrWord[0])
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

                                // SQL終了行を特定
                                if (arrWord[0].Replace(".", "") == "END-EXEC")
                                {
                                    IEnumerable<TokenInfo> tokens;
                                    if (!string.IsNullOrEmpty(sql))
                                    {
                                        tokens = TransactSqlHelpers.Parser.ParseSql(sql);
                                        SqlInfo sqlInfo = new SqlInfo(sql, tokens, sqlType, methodList[methodIndex].MethodNameP);
                                        sqlInfoList.Add(sqlInfo);
                                    }

                                    inSqlErea = false;
                                    continue;
                                }

                                // SQL文の取得
                                StringBuilder str = new StringBuilder();
                                str.Append(sql);
                                foreach (string val in arrWord)
                                {
                                    str.Append(val + " ");
                                }
                                sql = str.ToString();
                            }

                            // 関数の終了行を特定
                            if (arrWord[0] == methodList[methodIndex].MethodNameP + METHOD_END)
                            {
                                methodList[methodIndex].CalledMethod = methodList[methodIndex].CalledMethod.Distinct().ToList();
                                methodList[methodIndex].EndIndex = fileIndex;
                                inMethodErea = false;
                                continue;
                            }
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
                        foreach (var cm in method2.CalledMethod)
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
                    if (method1.CalledMethod.Count < 1) { continue; }

                    foreach (var cm in method1.CalledMethod)
                    {
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
                        foreach (var cm in method2.CalledMethod)
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
                FileInfo newFile = new FileInfo(OutFileName);

                // Excelファイル作成
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    bool ret = false;

                    // プログラム情報シート作成・編集（シート名：PGM情報_{読込ファイル名}）
                    WsPgmInfo = AddSheet(package, SHEET_NAME_PGMINFO, file);
                    ret = EditPgmInfoSheet(copyList, sqlInfoList, calledModuleList);

                    if (!ret) { return false; }

                    // 関数情報シート作成・編集（シート名：関数情報_{読込ファイル名}）
                    WsMethodInfo = AddSheet(package, SHEET_NAME_METHODINFO, file);
                    ret = EditMethodInfoSheet(methodList, sqlInfoList);

                    if (!ret) { return false; }

                    // 構造図シート作成・編集（シート名：構造図_{読込ファイル名}）
                    WsStruct = AddSheet(package, SHEET_NAME_STRUCT, file);
                    ret = EditStructSheet(methodList);

                    if (!ret) { return false; }

                    // 保存
                    package.Save();
                }
                return true;
            }
            catch (Exception ex)
            {
                _logger.Fatal(ex.Message);
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
            string frmLine = Mid(line.Trim(), start, 66);
            // テキスト前後のスペースを詰め、テキスト内にある複数のスペースを取り除く（"hoge   fuga" ⇒ "hoge fuga"）
            frmLine = Regex.Replace(frmLine.Trim(), @"\s{2,}", " ");
            // コメント行の場合、「A.」「D.」以降を取り除く
            if (exceptComFlg)
            {
                int index = -1;
                index = frmLine.IndexOf("A.");
                index = (index < 0) ? frmLine.IndexOf("D.") : index;
                frmLine = (index < 0) ? frmLine : Left(frmLine, index);
            }

            return frmLine;
        }

        /// <summary>
        /// 対象外句の判定
        /// </summary>
        /// <param name="arrWord"></param>
        /// <returns></returns>
        private static bool CheckLine(string[] arrWord)
        {
            string checkText = arrWord[0].Replace(".", "");

            // 空行、コメント行、SQL行、その他対象外句はスルー
            if (checkText == string.Empty || Left(checkText, 1) == COM_PREFIX
                    // 見出し部
                    || checkText == "IDENTIFICATION" || checkText == "PROGRAM-ID"
                    || checkText == "AUTHOR" || checkText == "DATE-WRITTEN"
                    || checkText == "DATE-COMPILED"
                    // 環境部
                    || checkText == "ENVIRONMENT" || checkText == "CONFIGURATION"
                    || checkText == "SOURCE-COMPUTER" || checkText == "OBJECT-COMPUTER"
                    || checkText == "INPUT-OUTPUT" || checkText == "FILE-CONTROL"
                    // データ部
                    || checkText == "DATA" || checkText == "FILE"
                    || checkText == "WORKING-STORAGE" || checkText == "LINKAGE"
                    || checkText == "REPORT" || checkText == "SCREEN"
                    // 手続き部
                    || checkText == "DISPLAY")
            {
                return false;
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
            string comLine = File.ReadAllLines(file, Enc).Skip(comIndex).Take(1).First();
            if (Mid(comLine, 7, 1) == COM_PREFIX)
                return FormatLine(comLine, true);
            else
                return string.Empty;
        }
        #endregion

        #region プログラム情報描画
        private static bool EditPgmInfoSheet(IReadOnlyDictionary<string, string> copyList
                                    , IEnumerable<SqlInfo> sqlInfoList
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
                SetStyleOfTitle(SHEET_NAME_PGMINFO, "B2:C2", Color.SpringGreen);
                WsPgmInfo.Cells[row, 2].Value = "COPY句";
                // 呼出モジュールのタイトルセット
                SetStyleOfTitle(SHEET_NAME_PGMINFO, "D2:D2", Color.Pink);
                WsPgmInfo.Cells[row, 4].Value = "呼出モジュール";
                // SQL変数宣言部のタイトルセット
                SetStyleOfTitle(SHEET_NAME_PGMINFO, "E2:K2", Color.LightSteelBlue);
                WsPgmInfo.Cells[row, 5].Value = "使用DB";
                WsPgmInfo.Cells[row, 7].Value = "[SELECT]";
                WsPgmInfo.Cells[row, 8].Value = "[INSERT]";
                WsPgmInfo.Cells[row, 9].Value = "[UPDATE]";
                WsPgmInfo.Cells[row, 10].Value = "[DELETE]";
                WsPgmInfo.Cells[row, 11].Value = "[CREATE]";
                WsPgmInfo.Cells[row, 7, row, 11].Style.Font.Size = 9;

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
                // 使用DBリストを書き込む
                // =================================================================
                SqlInfo _sqlInfo = new SqlInfo();
                DbInfo _dbInfo;
                List<DbInfo> dbInfoList = new List<DbInfo>();
                foreach (var sqlInfo in sqlInfoList)
                {
                    // SQL内で使用されているDBリストを取得
                    IEnumerable<string> dbList = _sqlInfo.GetDbList(sqlInfo.TokenList);
                    // DBの使用されているCRUDをセット
                    foreach (string table in dbList)
                    {
                        int i = dbInfoList.FindIndex(x => x.Name_P == table);
                        if(i < 0)
                        {
                            _dbInfo = new DbInfo(table, sqlInfo.Type);
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
                    WsPgmInfo.Cells[row, 5].Value = dbInfo.Name_P; // DB物理名
                    //WsPgmInfo.Cells[row, 5].Value = ; // todo:DBの論理名を取得
                    WsPgmInfo.Cells[row, 7].Value = dbInfo.SelectFlg ? "〇" : string.Empty; // SELECT
                    WsPgmInfo.Cells[row, 8].Value = dbInfo.InsertFlg ? "〇" : string.Empty; // INSERT
                    WsPgmInfo.Cells[row, 9].Value = dbInfo.UpdateFlg ? "〇" : string.Empty; // UPDATE
                    WsPgmInfo.Cells[row, 10].Value = dbInfo.DeleteFlg ? "〇" : string.Empty; // DELETE
                    WsPgmInfo.Cells[row, 11].Value = dbInfo.CreateFlg ? "〇" : string.Empty; // CREATE
                }

                WsPgmInfo.Cells.Style.Font.Name = FONT_NAME;
                WsPgmInfo.Cells[WsPgmInfo.Dimension.Address].AutoFitColumns(); // 列幅自動調整
            }
            catch (Exception ex)
            {
                _logger.Fatal(errorMsg + ex.Message);
                return false;
            }

            return true;
        }
        #endregion

        #region 関数情報描画
        /// <summary>
        /// 関数情報シート編集
        /// </summary>
        /// <param name="methodList"></param>
        /// <returns></returns>
        private static bool EditMethodInfoSheet(IEnumerable<Method> methodList, IReadOnlyCollection<SqlInfo> sqlInfoList)
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
                SetStyleOfTitle(SHEET_NAME_METHODINFO, "B2:Z2", Color.SpringGreen);
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
                SqlInfo _sqlInfo = new SqlInfo();
                foreach (var method in methodList)
                {
                    calledMethodCol = 7;
                    row++;

                    // SQLを呼んでいるかチェック
                    List<string> sqlTypeList = new List<string>();
                    foreach (var sqlInfo in sqlInfoList)
                    {
                        if (method.MethodNameP == sqlInfo.CalledMethod)
                        {
                            string sqlType = SqlTypeToString(sqlInfo.Type);
                            sqlTypeList.Add(sqlType);
                        }
                    }
                    IEnumerable<string> distinctList = sqlTypeList.Distinct();
                    var sqlTypeString = String.Join(",", distinctList);

                    // 関数物理名
                    WsMethodInfo.Cells[row, col].Value = method.MethodNameP;
                    // 関数論理名
                    WsMethodInfo.Cells[row, col + 1].Value = method.MethodNameL;
                    // 開始行数
                    WsMethodInfo.Cells[row, col + 2].Value = method.StartIndex;
                    // 終了行数
                    WsMethodInfo.Cells[row, col + 3].Value = method.EndIndex;
                    // DB操作
                    WsMethodInfo.Cells[row, col + 4].Value = sqlTypeString;
                    // 呼出関数
                    foreach (var cm in method.CalledMethod)
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

                WsMethodInfo.Cells.Style.Font.Name = FONT_NAME;
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

                // 起点となる関数名をExcellに書き込む
                WriteMethod(methodList, 0, row, col, new CalledMethod(string.Empty, false));

                // 呼び出される関数名を再帰的にExcelに書き込む
                GetCalledMethodRecursively(methodList, 0, row + 1, col + 1);

                WsStruct.Cells.Style.Font.Name = FONT_NAME;
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
        private static void WriteMethod(List<Method> methodList, int index, int row, int col, CalledMethod cm)
        {
            string outText = string.Empty;
            Color outColor = ColorMethod;
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
                outText = methodList[index].MethodNameP;

                // ハイパーリンクの設定
                string linkSheetName = SheetName.Replace(SHEET_NAME_STRUCT, SHEET_NAME_METHODINFO);
                WsStruct.Cells[row, col].Hyperlink
                    = new ExcelHyperLink("#'" + linkSheetName + "'!B" + (index + 3).ToString(), outText);
                WsStruct.Cells[row, col].Style.Font.UnderLine = true;
            }

            WsStruct.Cells[row, col].Value = outText;
            WsStruct.Cells[row, col].Style.Font.Color.SetColor(outColor);

            //ExcelShape es = WsStruct.Drawings.AddShape(outText + shapeNo.ToString(), eShapeStyle.FlowChartProcess);
            //es.SetPosition(row, 0, col, 0);
            //es.SetSize(100, 50);
            //es.Text = outText;
            //es.Fill.Style = eFillStyle.SolidFill;
            //es.Fill.Color = Color.Red;
            //shapeNo++;
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
            foreach (var calledMethod in methodList[index].CalledMethod)
            {
                // 呼出関数名の、関数リスト内での位置を取得
                int calledMethodIndex = GetMethodIndex(methodList, calledMethod);

                // Excelに書き込む
                WriteMethod(methodList, calledMethodIndex, row, col, calledMethod);

                row++;
                InitRow = row;

                if (calledMethodIndex > -1)
                {
                    // 呼出先の関数内から呼び出している関数がないか、再帰的に探索する
                    GetCalledMethodRecursively(methodList, calledMethodIndex, row, col + 1);
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
            {
                return -1;
            }

            int i = 0;
            foreach (var method in methodList)
            {
                if (cm.Name == method.MethodNameP)
                {
                    return i;
                }
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
                case SHEET_NAME_PGMINFO:
                    WsPgmInfo.Cells[range].Style.Font.Bold = true;
                    WsPgmInfo.Cells[range].Style.Fill.PatternType = ExcelFillStyle.DarkVertical;
                    WsPgmInfo.Cells[range].Style.Fill.BackgroundColor.SetColor(titleColor);
                    break;
                case SHEET_NAME_METHODINFO:
                    WsMethodInfo.Cells[range].Style.Font.Bold = true;
                    WsMethodInfo.Cells[range].Style.Fill.PatternType = ExcelFillStyle.DarkVertical;
                    WsMethodInfo.Cells[range].Style.Fill.BackgroundColor.SetColor(titleColor);
                    break;
                case SHEET_NAME_STRUCT:
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

        private static string SqlTypeToString(SqlType sqlType)
        {
            string ret = string.Empty;
            switch (sqlType)
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
        public string MethodNameP { get; internal set; }
        public string MethodNameL { get; internal set; }
        public int StartIndex { get; internal set; }
        public int EndIndex { get; internal set; }
        public List<CalledMethod> CalledMethod { get; internal set; }
        public bool CalledFlg { get; internal set; }

        public Method(string methodName, int startIndex, int endIndex)
        {
            MethodNameP = methodName;
            StartIndex = startIndex;
            EndIndex = endIndex;
            CalledMethod = new List<CalledMethod>();
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

        public CalledMethod(string name, bool moduleFlg)
        {
            Name = name;
            ModuleFlg = moduleFlg;
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
        public string Value { get; internal set; }
        public IEnumerable<TokenInfo> TokenList { get; internal set; }
        public SqlType Type { get; internal set; }
        public string CalledMethod { get; internal set; }

        public SqlInfo(string value, IEnumerable<TokenInfo> tokenList, SqlType type, string calledMethod)
        {
            Value = value;
            TokenList = tokenList;
            Type = type;
            CalledMethod = calledMethod;
        }

        public SqlInfo()
        {
        }

        public List<string> GetDbList(IEnumerable<TokenInfo> tokenList)
        {
            List<string> dbList = new List<string>();
            bool dbAddFlg1 = false;
            bool dbAddFlg2 = false;

            foreach (var token in tokenList)
            {
                if (token.Token == Tokens.TOKEN_FROM || token.Token == Tokens.TOKEN_JOIN)
                {
                    dbAddFlg1 = true;
                    continue;
                }
                if (dbAddFlg1 && token.Token == Tokens.TOKEN_ID)
                {
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
                    dbList.Add(token.Sql);
                    dbAddFlg2 = false;
                    continue;
                }
            }

            return dbList;
        }
    }

    /// <summary>
    /// DB情報クラス
    /// </summary>
    public class DbInfo : IEquatable<DbInfo>
    {
        public string Name_P { get; internal set; }
        public string Name_L { get; internal set; }
        public bool SelectFlg { get; internal set; } = false;
        public bool InsertFlg { get; internal set; } = false;
        public bool UpdateFlg { get; internal set; } = false;
        public bool DeleteFlg { get; internal set; } = false;
        public bool CreateFlg { get; internal set; } = false;

        public DbInfo(string name, SqlType type)
        {
            Name_P = name;
            // todo:論理名取得ロジックを追加する
            // Name_L = "";
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