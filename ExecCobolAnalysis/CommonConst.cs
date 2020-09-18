namespace ExecCobolAnalysis
{
    class CommonConst
    {
        // 読み込みDB定義
        public const string TABLE_P = "Table_P";
        public const string TABLE_L = "Table_L";
        public const string COLUMN_P = "Column_P";
        public const string COLUMN_L = "Column_L";

        // 書き込みシート名
        public const string SHEET_NAME_PGMINFO = "PGM情報";
        public const string SHEET_NAME_METHODINFO = "関数情報";
        public const string SHEET_NAME_STRUCT = "構造図";

        // COBOL共通フレーズ
        public const string COM_PREFIX = "*";
        public const string WORD_DIVISION = "DIVISION";

        // COBOL見出し部フレーズ
        public const string WORD_IDENTIFICATION = "IDENTIFICATION";
        // COBOL環境部フレーズ
        public const string WORD_ENVIRONMENT = "ENVIRONMENT";
        // COBOLデータ部フレーズ
        public const string WORD_DATA = "DATA";
        // COBOL手続き部フレーズ
        public const string WORD_PROCEDURE = "PROCEDURE";

        // フォント名
        public const string FONT_NAME_MEIRYOUI = "Meiryo UI";

        // return値
        public const int RETURN_OK = 0;
        public const int RETURN_ERR_100 = 100;
        public const int RETURN_ERR_200 = 200;

    }
}
