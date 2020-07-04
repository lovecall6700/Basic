using System;
using System.Collections;
using System.IO;

namespace Basic
{
    /// <summary>
    /// Basic
    /// パソコン教室用BASIC
    /// </summary>
    public class Basic
    {
        /// <summary>
        /// 浮動小数点演算時の誤差
        /// </summary>
        private const double Epsilon = 0.00001;

        /// <summary>
        /// 結合方向左
        /// </summary>
        private const int LawLeft = 0;

        /// <summary>
        /// 結合方向右
        /// </summary>
        private const int LawRight = 1;

        /// <summary>
        /// 乱数モード
        /// 0以上から1未満の正の数を返す
        /// </summary>
        private const int RndMode0To1 = 0;

        /// <summary>
        /// 乱数モード
        /// 0から指定数以下の正の整数を返す
        /// </summary>
        private const int RndMode0ToArg = 1;

        /// <summary>
        /// 乱数
        /// </summary>
        private Random _rnd = new Random();

        /// <summary>
        /// RND関数の仕様切り替え
        /// デフォルトは0以上から1未満の正の数
        /// </summary>
        private int _rndMode = RndMode0To1;

        /// <summary>
        /// 前回の乱数値
        /// </summary>
        private double _lastRnd = -1f;

        /// <summary>
        /// 行番号トレースモード
        /// </summary>
        private bool _traceMode = false;

        /// <summary>
        /// ソース
        /// </summary>
        private ArrayList _source;

        /// <summary>
        /// 実行状態
        /// </summary>
        private bool _running = true;

        /// <summary>
        /// プログラムカウンタ
        /// </summary>
        private int _pc = 0;

        /// <summary>
        /// GOSUB制御
        /// </summary>
        private readonly Stack _gosubInfo = new Stack();

        /// <summary>
        /// 変数
        /// </summary>
        private Hashtable _variableData = new Hashtable();

        /// <summary>
        /// 配列の最初の添字
        /// </summary>
        private int _arrayBase = 1;

        /// <summary>
        /// 配列
        /// キー:配列名
        /// 値:配列の値(キー:添字、値:double)
        /// </summary>
        private readonly Hashtable _arrayData = new Hashtable();

        /// <summary>
        /// 配列の次元数
        /// キー:配列名
        /// 値:次元数
        /// </summary>
        private readonly Hashtable _arrayDim = new Hashtable();

        /// <summary>
        /// ユーザー定義関数名
        /// 値:関数名
        /// </summary>
        private readonly ArrayList _userFn = new ArrayList();

        /// <summary>
        /// ユーザー定義関数引数
        /// 値:引数
        /// </summary>
        private readonly ArrayList _userFnArgs = new ArrayList();

        /// <summary>
        /// ユーザー定義関数数式
        /// 値:数式
        /// </summary>
        private readonly ArrayList _userFnExp = new ArrayList();

        /// <summary>
        /// 行番号管理
        /// キー:行番号
        /// 値:ソースのインデックス
        /// </summary>
        private readonly Hashtable _lineNoToIndex = new Hashtable();

        /// <summary>
        /// 行番号管理
        /// キー:ソースのインデックス
        /// 値:行番号
        /// </summary>
        private readonly Hashtable _indexToLineNo = new Hashtable();

        /// <summary>
        /// データ
        /// </summary>
        private readonly ArrayList _data = new ArrayList();

        /// <summary>
        /// データ管理
        /// キー:行番号
        /// 値:データのインデックス
        /// </summary>
        private readonly Hashtable _lineNoToDataIndex = new Hashtable();

        /// <summary>
        /// データインデックス
        /// </summary>
        private int _dataIndex = 0;

        /// <summary>
        /// FOR-NEXT対応
        /// キー:FORのインデックス
        /// 値:NEXTのインデックス
        /// </summary>
        private readonly Hashtable _nextToForIndex = new Hashtable();

        /// <summary>
        /// NEXT-FOR対応
        /// キー:NEXTのインデックス
        /// 値:FORのインデックス
        /// </summary>
        private readonly Hashtable _forToNextIndex = new Hashtable();

        /// <summary>
        /// IFと次の行対応
        /// キー:IFのインデックス
        /// 値:次の行のインデックス
        /// </summary>
        private readonly Hashtable _ifToNextLineIndex = new Hashtable();

        /// <summary>
        /// 組み込み関数一覧
        /// #NEG#は内部処理にのみ用いられるので、ここには含まれていない
        /// </summary>
        private readonly string[] _buildinFnNames = new string[] {
            "ABS", "ACOS", "ASIN", "ATN", "COS", "EXP", "INT", "LOG", "LOG10", "LOG2", "MAX", "MIN", "MOD", "RND", "SGN", "SIN", "SQR", "TAN",
        };

        /// <summary>
        /// 比較演算子
        /// </summary>
        private readonly Hashtable _rerarionalOperators = new Hashtable();

        /// <summary>
        /// 演算子の優先順位
        /// </summary>
        private readonly Hashtable _precedence = new Hashtable();

        /// <summary>
        /// 結合法則
        /// </summary>
        private readonly Hashtable _associativeLaw = new Hashtable();

        /// <summary>
        /// 最大長コマンドの文字数
        /// </summary>
        private const int MaxCmdLength = 9;

        /// <summary>
        /// エレメント(Element)
        /// </summary>
        private readonly string[] _elements = new string[] {
            "Array", //配列の解析
			"Loop{", //ループ開始
			"}Loop", //ループ終了
			"Expression", //数式
			"Variable", //変数
			"RelationalOperator", //比較演算子
			"QuotedString", //"文字列"
			"String", //文字列
			"End", //ステートメント終了
			"EndOfLine", //行の解析を終了
			"Optional{", //オプション開始
			"}Optional", //オプション終了
			"Select{", //選択開始
			"}Select", //選択終了
			"Statement", //ステートメント
			"Group{", //グループ開始
			"}Group", //グループ終了
			"Rename", //コマンド名の読み替えを行う
			"Recursive", //再帰呼び出し
			"DelValue", //値の除去
			"AddValue", //値の追加
			"AddExpression", //数式の追加
			"AddNull", //NULLの追加
			"SetFlag", //フラグのセット
			"CheckFlag", //フラグのチェック
			"ResetFlag" //フラグのリセット
		};

        /// <summary>
        /// ステートメント
        /// キー:コマンド名
        /// 値:文法
        /// 
        /// 以下の値は文字列そのものへのマッチとする
        /// "1","#","(",")",",",":",";","=",
        /// "CLEAR","CURSOR","EXPRESSION","FOR","G.","GO",
        /// "GOS.","GOSUB","GOTO","INPUT","LET","NEXT",
        /// "PRINT","REM","RETURN","STEP","STOP","STRING",
        /// "SUB","T.","TAB","THEN","TO","OPTION","BASE","RANDOMIZE"
        /// </summary>
        private readonly Hashtable _statements = new Hashtable();

        /// <summary>
        /// Basic
        /// </summary>
        public Basic()
        {
            _arrayData.Add("@", new double[100]);
            _arrayDim.Add("@", new int[] { 100 });
            _rerarionalOperators.Add("=", "=");
            _rerarionalOperators.Add("<", "<");
            _rerarionalOperators.Add(">", ">");
            _rerarionalOperators.Add("<=", "<=");
            _rerarionalOperators.Add("=<", "<=");
            _rerarionalOperators.Add(">=", ">=");
            _rerarionalOperators.Add("=>", ">=");
            _rerarionalOperators.Add("<>", "<>");
            _rerarionalOperators.Add("><", "<>");
            _rerarionalOperators.Add("#", "<>");
            _precedence.Add("^", 4);
            _precedence.Add("*", 3);
            _precedence.Add("/", 3);
            _precedence.Add("+", 2);
            _precedence.Add("-", 2);
            _associativeLaw.Add("^", LawRight);
            _associativeLaw.Add("*", LawLeft);
            _associativeLaw.Add("/", LawLeft);
            _associativeLaw.Add("+", LawLeft);
            _associativeLaw.Add("-", LawLeft);

            _statements.Add("AFTER_LINE_NUMBER", new string[] {//cmd[0] 特殊コマンド
				"DelValue", //AFTER_LINE_NUMBERを削除する
				"Statement",
                "Loop{",
                    ":", "DelValue",
                    "Recursive",
                "}Loop",
                "End",
            });
            _statements.Add("Array", new string[] {//cmd[0] 特殊コマンド
				"DelValue", //Arrayを削除する
				"Select{",
                    "Group{",
                        "@", //cmd[1]
					"}Group",
                    "Group{",
                        "Variable", //cmd[1]
					"}Group",
                "}Select",
                "(", "DelValue",
                "Expression", //cmd[2]
				"Loop{",
                    ",", "DelValue",
                    "Expression",
                "}Loop",
                ")", "DelValue",
                "End"
            });
            _statements.Add("INPUT_LINE", new string[] { //cmd[0] 特殊コマンド
				"Expression", //cmd[1]
				"Loop{",
                    ",", "DelValue",
                    "Recursive",
                "}Loop",
                "End",
            });
            _statements.Add("CHANGE", new string[] { //cmd[0]
				"Select{",
                    "Group{",
                        "QuotedString", //cmd[1]
						"AddValue", "STRING", //cmd[2]
					"}Group",
                    "Group{",
                        "Expression", //cmd[1]
						"AddValue", "EXPRESSION", //cmd[2]
					"}Group",
                "}Select",
                "Loop{",
                    "Select{",
                        "Group{",
                            ",", "DelValue",
                            "QuotedString", //cmd[3, 5, ...]
							"AddValue", "STRING", //cmd[4, 6, ...]
						"}Group",
                        "Group{",
                            ",", "DelValue",
                            "Expression", //cmd[3, 5, ...]
							"AddValue", "EXPRESSION", //cmd[4, 6, ...]
						"}Group",
                    "}Select",
                "}Loop",
                "End",
            });
            _statements.Add("CURSOR", new string[] { //cmd[0]
				"Expression", //cmd[1]
				",", "DelValue",
                "Expression", //cmd[2]
				"End",
            });
            _statements.Add("DATA", new string[] { //cmd[0]
				"Expression", //cmd[1]
				"Loop{",
                    ",", "DelValue",
                    "Recursive",
                "}Loop",
                "End",
            });
            _statements.Add("ERASE", new string[] { //cmd[0]
				"Select{",
                    "Group{",
                        "@",
                    "}Group",
                    "Group{",
                        "Variable",
                    "}Group",
                "}Select",
                "Loop{",
                    ",", "DelValue",
                    "Recursive",
                "}Loop",
                "End",
            });
            _statements.Add("SWAP", new string[] { //cmd[0]
				"Variable", //cmd[1]
				",", "DelValue",
                "Variable", //cmd[2]
				"End",
            });
            _statements.Add("DEF", new string[] { //cmd[0]
				"String", //cmd[1]
				"(", "DelValue",
                "Variable", //cmd[2, 3, ...]
				"Loop{",
                    ",", "DelValue",
                    "Variable",
                "}Loop",
                ")", "DelValue",
                "=", "DelValue",
                "Expression", //cmd[last]
				"End",
            });
            _statements.Add("DIM", new string[] { //cmd[0]
				"Array",
                "Loop{",
                    ",", "DelValue",
                    "Recursive",
                "}Loop",
                "End",
            });
            _statements.Add("FOR", new string[] { //cmd[0]
				"Variable", //cmd[1]
				"=", "DelValue",
                "Expression", //cmd[2]
				"TO", "DelValue",
                "Expression", //cmd[3]
				"Optional{",
                    "Group{",
                        "STEP", "DelValue",
                        "Expression", //cmd[4]
					"}Group",
                    "Group{",
                        "AddExpression", "1", //cmd[4]
					"}Group",
                "}Optional",
                "End",
            });
            _statements.Add("NEXT", new string[] { //cmd[0]
				"Optional{",
                    "Group{",
                        "Variable", //cmd[1]
					"}Group",
                    "Group{",
                        "AddNull", //cmd[1]
					"}Group",
                "}Optional",
                "End",
            });
            _statements.Add("GOTO", new string[] { //cmd[0]
				"Expression", //cmd[1]
				"End",
            });
            _statements.Add("GOSUB", new string[] { //cmd[0]
				"Expression", //cmd[1]
				"End",
            });
            _statements.Add("RETURN", new string[] { //cmd[0]
				"End",
            });
            _statements.Add("IF", new string[] { //cmd[0]
				"Expression", //cmd[1]
				"RelationalOperator", //cmd[2]
				"Expression", //cmd[3]
				"Optional{",
                    "Group{",
                        "THEN", "DelValue",
                    "}Group",
                    "Group{",
                        "T.", "DelValue",
                    "}Group",
                "}Optional",
                "Select{",
                    "Group{",
                        "Statement",
                    "}Group",
                    "Group{",
                        "Rename", "GOTO",
                    "}Group",
                "}Select",
                "End",
            });
            _statements.Add("INPUT", new string[] { //cmd[0]
				"Optional{",
                    "Group{",
                        "QuotedString", //cmd[1]
						",", "DelValue",
                    "}Group",
                        "Group{",
                        "AddNull", //cmd[1]
					"}Group",
                "}Optional",
                "Variable", //cmd[2, 3, ...]
				"Loop{",
                    ",", "DelValue",
                    "Variable", //RECURSIVEで個別の命令にできない。
				"}Loop",
                "End",
            });
            _statements.Add("ON", new string[] { //cmd[0]
				"Expression", //cmd[1]
				"Select{",
                    "Group{",
                        "GOS.", "DelValue",
                        "SetFlag", "GOS.",
                    "}Group",
                    "Group{",
                        "G.", "DelValue",
                        "SetFlag", "G.",
                    "}Group",
                    "Group{",
                        "GO", "DelValue",
                    "}Group",
                "}Select",
                "Select{",
                    "Group{",
                        "CheckFlag", "G.",
                        "AddValue", "GOTO", //cmd[2]
					"}Group",
                    "Group{",
                        "CheckFlag", "GOS.",
                        "AddValue", "GOSUB", //cmd[2]
					"}Group",
                    "Group{",
                        "TO", "DelValue",
                        "AddValue", "GOTO", //cmd[2]
					"}Group",
                    "Group{",
                        "SUB", "DelValue",
                        "AddValue", "GOSUB", //cmd[2]
					"}Group",
                "}Select",
                "Expression", //cmd[3, 4, ...]
				"Loop{",
                    ",", "DelValue",
                    "Expression",//RECURSIVEで個別の命令にできない。
				"}Loop",
                "End",
            });
            _statements.Add("RANDOMIZE", new string[] { //cmd[0]
				"End",
            });
            _statements.Add("READ", new string[] { //cmd[0]
				"Variable", //cmd[1]
				"Loop{",
                    ",", "DelValue",
                    "Recursive",
                "}Loop",
                "End",
            });
            _statements.Add("REM", new string[] { //cmd[0]
				"EndOfLine",
            });
            _statements.Add("RESTORE", new string[] { //cmd[0]
				"Optional{",
                    "Group{",
                        "Expression", //cmd[1]
					"}Group",
                    "Group{",
                        "AddNull", //cmd[1]
					"}Group",
                "}Optional",
                "End",
            });
            _statements.Add("PRINT", new string[] { //cmd[0]
				"Loop{",
                    "Select{",
                        "Group{",
                            "TAB", "DelValue",
                            "(", "DelValue",
                            "Expression", //cmd[1, 4, ...]
							")", "DelValue",
                            "AddValue", "TAB", //cmd[2, 5, ...]
						"}Group",
                        "Group{",
                            "#", "DelValue",
                            "Expression", //cmd[1, 4, ...]
							"AddValue", "#", //cmd[2, 5, ...]
						"}Group",
                        "Group{",
                            "QuotedString", //cmd[1, 4, ...]
							"AddValue", "STRING", //cmd[2, 5, ...]
						"}Group",
                        "Group{",
                            "Expression", //cmd[1, 4...]
							"AddValue", "EXPRESSION", //cmd[2, 5, ...]
						"}Group",
                    "}Select",
                    "Select{",
                        "Group{",
                            ",", //cmd[3, 6, ...]
						"}Group",
                        "Group{",
                            ";", //cmd[3, 6, ...]
						"}Group",
                        "Group{",
                            "AddNull", //cmd[3]
							"End",
                        "}Group",
                    "}Select",
                "}Loop",
                "End",
            });
            _statements.Add("LET", new string[] { //cmd[0]
				"Select{",
                    "Group{",
                        "Array",
						//Arrayの展開によって以下の順で自動追加される
						//配列名 cmd[1]
						//添字 cmd[2]
						"AddValue", "ARRAY", //cmd[3]
					"}Group",
                    "Group{",
                        "Variable", //cmd[1]
						"AddNull", //cmd[2]
						"AddValue", "VARIABLE", //cmd[3]
					"}Group",
                "}Select",
                "=", "DelValue",
                "Expression", //cmd[4]
				"Loop{",
                    ",", "DelValue",
                    "Recursive",
                "}Loop",
                "End",
            });
            _statements.Add("BASE", new string[] { //cmd[0]
				"Select{",
                    "Group{",
                        "0",
                    "}Group",
                    "Group{",
                        "1",
                    "}Group",
                "}Select",
                "End",
            });
            _statements.Add("OPTION", new string[] {
                "DelValue", //OPTION BASEはBASEのみ残す
				"BASE", //cmd[0]
				"Expression", //cmd[1]
				"End",
            });
            _statements.Add("STOP", new string[] { //cmd[0]
				"End",
            });
            _statements.Add("CLEAR", new string[] { //cmd[0]
				"End",
            });
            _statements.Add("CLS", new string[] { //CLEAR
				"Rename", "CLEAR",
                "End",
            });
            _statements.Add("LC", new string[] { //CURSOR
				"Rename", "CURSOR",
                "End",
            });
            _statements.Add("LOCATE", new string[] { //CURSOR
				"Rename", "CURSOR",
                "End",
            });
            _statements.Add("F.", new string[] { //FOR
				"Rename", "FOR",
                "End",
            });
            _statements.Add("G.", new string[] { //GOTO
				"Rename", "GOTO",
                "End",
            });
            _statements.Add("GOS.", new string[] { //GOSUB
				"Rename", "GOSUB",
                "End",
            });
            _statements.Add("IN.", new string[] { //INPUT
				"Rename", "INPUT",
                "End",
            });
            _statements.Add("N.", new string[] { //NEXT
				"Rename", "NEXT",
                "End",
            });
            _statements.Add("?", new string[] { //PRINT
				"Rename", "PRINT",
                "End",
            });
            _statements.Add("P.", new string[] { //PRINT
				"Rename", "PRINT",
                "End",
            });
            _statements.Add("R.", new string[] { //RETURN
				"Rename", "RETURN",
                "End",
            });
            _statements.Add("S.", new string[] { //STOP
				"Rename", "STOP",
                "End",
            });
            _statements.Add("END", new string[] { //STOP
				"Rename", "STOP",
                "End",
            });
            _statements.Add("'", new string[] { //REM
				"Rename", "REM",
                "End",
            });
            _statements.Add("WRITE", new string[] { //PRINT
				"Rename", "PRINT",
                "End",
            });
            _statements.Add("GO", new string[] { //GOTO or GOSUB or GO TO or GO SUB
				"Select{",
                    "Group{",
                        "TO", "DelValue",
                        "Rename", "GOTO",
                    "}Group",
                    "Group{",
                        "SUB", "DelValue",
                        "Rename", "GOSUB",
                    "}Group",
                "}Select",
                "End",
            });
            _statements.Add("", new string[] { //LET
				"Rename", "LET",
                "End",
            });
        }

        /// <summary>
        /// ソースファイルの解析を行う
        /// </summary>
        /// <param name="list">ソースコードの文字配列</param>
        public void ParseSource(string[] list)
        {
            int lineNo = 0;
            Stack forNextStack = new Stack();
            ArrayList source = new ArrayList();
            bool useDim = false;
            bool atArrayReDim = false;

            //ロード時に初期化するデータ
            //変数や配列はロードされても引き継ぐので初期化しない
            Init();
            InitLoad();

            foreach (string line in list)
            {
                int len = line.Length;

                if (len == 0)
                {
                    continue;
                }
                if (len > 80)
                {
                    throw new Exception("1行が80文字を超えています。");
                }

                for (int i = 0; i < len; i++)
                {
                    if (!IsChar(line[i]) && line[i] != '\t')
                    {
                        throw new Exception("ソースファイルの読み込みに失敗しました。指定されたソースコードに改行とタブ以外のコントロールコードが含まれています。");
                    }
                }

                //行番号の抽出
                //行番号そのものを文法に含めてもいいのだが、
                //先に知っておいたほうがエラー表示時に二度手間にならない。
                int index = 0;
                Skip(line, ref index);


                if (!GetDecimal(line, ref index, true, out double r))
                {
                    throw new Exception(string.Format("文法エラーです。下記の行に行番号が必要です。\n{0}", line));
                }

                int tmpNo = (int)r;

                if (tmpNo <= lineNo)
                {
                    throw new Exception(string.Format("文法エラーです。行番号{0}は直前の行番号{1}よりも大きな数を指定してください。", tmpNo, lineNo));
                }

                lineNo = tmpNo;
                index++;
                ArrayList result = GetStatement("AFTER_LINE_NUMBER", line, ref index, false);

                if (result == null)
                {
                    throw new Exception(string.Format("文法エラーです。 行番号{0}", lineNo));
                }

                Skip(line, ref index);

                if (line.Length != index)
                {
                    throw new Exception(string.Format("文法エラーです。 行番号{0}", lineNo));
                }

                _indexToLineNo.Add(source.Count, lineNo);
                _lineNoToIndex.Add(lineNo, source.Count);

                for (int j = 0; j < result.Count; j++)
                {
                    ArrayList cmd = (ArrayList)result[j];

                    switch ((string)cmd[0])
                    {
                        case "BASE":
                            //BASEは最初にDIMが出現する前に実行しなければならない
                            if (useDim)
                            {
                                throw new Exception("文法エラーです。BASEは最初のDIMよりも前に実行してください。");
                            }
                            // 配列の最初の数を指定する
                            string baseNum = (string)cmd[1];
                            if (baseNum == "0")
                            {
                                _arrayBase = 0;
                            }
                            else if (baseNum == "1")
                            {
                                _arrayBase = 1;
                            }
                            break;
                        case "DIM":
                            useDim = true;
                            //次元数の管理を行う
                            string varName = (string)cmd[1];
                            ArrayList dimData = (ArrayList)cmd[2];
                            int dimSize = dimData.Count;

                            if (dimSize > 3)
                            {
                                throw new Exception(string.Format("文法エラーです。3次元を超える配列{0}は宣言できません。", varName));
                            }

                            if (IsArrayName(varName))
                            {
                                //@のみ、再定義を1回だけ認める
                                if (varName == "@" && !atArrayReDim)
                                {
                                    _arrayDim.Remove("@");
                                    _arrayData.Remove("@");
                                    atArrayReDim = true;
                                }
                                else
                                {
                                    throw new Exception(string.Format("文法エラーです。配列{0}の再定義は変更できません。", varName));
                                }
                            }

                            int correctBaseNum = 1 - _arrayBase;

                            if (dimSize == 1)
                            {
                                Stack s1 = (Stack)dimData[0];
                                int v1 = (int)Calc(ref s1, ref _variableData);
                                if (v1 <= 0)
                                {
                                    throw new Exception(string.Format("文法エラーです。配列{0}[{1}]を初期化できませんでした。配列の添字に0以下の数は指定できません。", varName, v1));
                                }
                                _arrayDim.Add(varName, new int[] { v1 });
                                _arrayData.Add(varName, new double[v1 + correctBaseNum]);
                            }
                            else if (dimSize == 2)
                            {
                                Stack s1 = (Stack)dimData[0];
                                Stack s2 = (Stack)dimData[1];
                                int v1 = (int)Calc(ref s1, ref _variableData);
                                int v2 = (int)Calc(ref s2, ref _variableData);
                                if (v1 <= 0 || v2 <= 0)
                                {
                                    throw new Exception(string.Format("文法エラーです。配列{0}[{1},{2}]を初期化できませんでした。配列の添字に0以下の数は指定できません。", varName, v1, v2));
                                }
                                _arrayDim.Add(varName, new int[] { v1, v2 });
                                _arrayData.Add(varName, new double[v1 + correctBaseNum, v2 + correctBaseNum]);
                            }
                            else if (dimSize == 3)
                            {
                                Stack s1 = (Stack)dimData[0];
                                Stack s2 = (Stack)dimData[1];
                                Stack s3 = (Stack)dimData[2];
                                int v1 = (int)Calc(ref s1, ref _variableData);
                                int v2 = (int)Calc(ref s2, ref _variableData);
                                int v3 = (int)Calc(ref s3, ref _variableData);
                                if (v1 <= 0 || v2 <= 0 || v3 <= 0)
                                {
                                    throw new Exception(string.Format("文法エラーです。配列{0}[{1},{2},{3}]を初期化できませんでした。配列の添字に0以下の数は指定できません。", varName, v1, v2, v3));
                                }
                                _arrayDim.Add(varName, new int[] { v1, v2, v3 });
                                _arrayData.Add(varName, new double[v1 + correctBaseNum, v2 + correctBaseNum, v3 + correctBaseNum]);
                            }
                            break;
                        case "DEF":
                            //1 2	3 4 5 6
                            //0 1	2 3 4 5
                            //DEF NAME A B C EXP
                            int argsCount = cmd.Count - 1;
                            string fnName = (string)cmd[1];
                            if (IsBuildinFnName(fnName))
                            {
                                throw new Exception(string.Format("文法エラーです。ユーザー定義関数{0}を定義できません。組み込み関数名と同じにはできません。", fnName));
                            }
                            ArrayList args = new ArrayList();
                            for (int i = 2; i < argsCount; i++)
                            {
                                args.Add((string)cmd[i]);
                            }
                            args.Reverse();
                            if (IsUserFnName(fnName))
                            {
                                throw new Exception(string.Format("文法エラーです。定義済みのユーザー定義関数{0}は再定義できません。", fnName));
                            }
                            _userFn.Add(fnName);
                            _userFnArgs.Add(args);
                            _userFnExp.Add(cmd[argsCount]);
                            break;
                        case "DATA":
                            _data.Add(cmd[1]);
                            if (!_lineNoToDataIndex.Contains(lineNo))
                            {
                                _lineNoToDataIndex.Add(lineNo, _data.Count + j - 1);
                            }
                            break;
                        case "FOR":
                            forNextStack.Push(source.Count + j);
                            break;
                        case "NEXT":
                            if (forNextStack.Count == 0)
                            {
                                throw new Exception("文法エラーです。FORとNEXTの対応が一致しません。");
                            }
                            int f = (int)forNextStack.Pop();
                            _forToNextIndex.Add(f, source.Count + j);
                            _nextToForIndex.Add(source.Count + j, f);
                            break;
                        case "IF":
                            _ifToNextLineIndex.Add(source.Count + j, result.Count + source.Count);
                            break;
                    }
                }

                source.AddRange(result);
            }

            if (source.Count == 0)
            {
                throw new Exception("指定されたソースコードに有効な行が含まれていません。");
            }

            _source = source;
        }

        /// <summary>
        /// ソースファイルの読み込みを行う
        /// </summary>
        /// <param name="fileName">ソースファイル名</param>
        public void LoadSource(string fileName)
        {
            ArrayList source = new ArrayList();

            if (fileName == null)
            {
                throw new Exception("ソースファイルの読み込みに失敗しました。ソースファイル名を指定してください。");
            }
            if (!File.Exists(fileName))
            {
                throw new Exception(string.Format("ソースファイルの読み込みに失敗しました。指定されたソースファイル{0}が存在しません。", fileName));
            }

            FileInfo fileInfo = new FileInfo(fileName);
            fileInfo.Refresh();

            if (fileInfo.Length == 0)
            {
                throw new Exception(string.Format("ソースファイルの読み込みに失敗しました。指定されたソースファイル{0}のファイルサイズがゼロです。", fileName));
            }

            StreamReader reader = null;

            try
            {
                reader = new StreamReader(fileName, System.Text.Encoding.Default);
                while (reader.Peek() > -1)
                {
                    source.Add(reader.ReadLine());
                }
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                }
            }

            ParseSource((string[])source.ToArray(typeof(string)));
        }

        /// <summary>
        /// Loadのみで初期化する項目
        /// </summary>
        private void InitLoad()
        {
            _lineNoToIndex.Clear();
            _indexToLineNo.Clear();
            _lineNoToDataIndex.Clear();
            _forToNextIndex.Clear();
            _nextToForIndex.Clear();
            _ifToNextLineIndex.Clear();
            _data.Clear();
            _source = null;
        }

        /// <summary>
        /// LoadとRUNで初期化する共通項目
        /// </summary>
        private void Init()
        {
            _pc = 0;
            _dataIndex = 0;
            _gosubInfo.Clear();
        }

        /// <summary>
        /// 浮動小数点同士の比較を行う
        /// </summary>
        /// <param name="a">値1</param>
        /// <param name="b">値2</param>
        /// <returns></returns>
        private bool AlmostEqual(double a, double b)
        {
            return Math.Abs(a - b) < Epsilon;
        }

        /// <summary>
        /// 実行を行う
        /// </summary>
        public void Run()
        {
            if (_source == null)
            {
                throw new Exception("実行時エラーです。実行データがnullです。");
            }
            if (_source.Count == 0)
            {
                throw new Exception("実行時エラーです。実行データが空です。");
            }

            //ロード時に初期化されるもの以外は実行時に初期化する
            //ロード後に実行だけ繰り返すケースも想定される
            //ソースコードと対応する行番号とインデックスの対応表、
            //データ一覧、データとインデックスのの対応表はロードで初期化される
            //CHANGEでソースファイルの切り替えが行われた時も注意
            int lineNo = 0;
            Init();

            try
            {
                while (_running)
                {
                    if (_indexToLineNo.ContainsKey(_pc))
                    {
                        lineNo = (int)_indexToLineNo[_pc];
                        if (_traceMode)
                        {
                            Console.Write(lineNo.ToString());
                            Console.Write(" ");
                        }
                    }

                    ArrayList cmd = (ArrayList)_source[_pc];
                    string name = (string)cmd[0];

                    if (name == "CHANGE")
                    {
                        //1	2 3 4 5
                        //0	1 2 3 4
                        //CHANGE A B C D
                        string fileName = "";
                        int len = cmd.Count / 2;

                        for (int i = 0; i < len; i++)
                        {
                            string type = (string)cmd[i * 2 + 2];
                            if (type == "EXPRESSION")
                            {
                                Stack stack = (Stack)cmd[i * 2 + 1];
                                double val = Calc(ref stack, ref _variableData);
                                fileName += DelZero(val.ToString());
                            }
                            else if (type == "STRING")
                            {
                                fileName += (string)cmd[i * 2 + 1];
                            }
                        }

                        LoadSource(fileName);
                        lineNo = 0;
                    }
                    else if (name == "CURSOR")
                    {
                        Stack x = (Stack)cmd[1];
                        Stack y = (Stack)cmd[2];
                        Console.CursorLeft = (int)Calc(ref x, ref _variableData);
                        Console.CursorTop = (int)Calc(ref y, ref _variableData);
                        _pc++;
                    }
                    else if (name == "FOR")
                    {
                        string varName = (string)cmd[1];
                        Stack s2 = (Stack)cmd[2];
                        Stack s3 = (Stack)cmd[3];
                        Stack s4 = (Stack)cmd[4];
                        double forVal = Calc(ref s2, ref _variableData);
                        double toVal = Calc(ref s3, ref _variableData);
                        double stepVal = Calc(ref s4, ref _variableData);
                        SetVariableVal(ref _variableData, varName, forVal);

                        // step値 > 0
                        if (stepVal > 0)
                        {
                            // FOR制御に入る段階で初期値fromが限界値toを上回っている場合はループを実行しない
                            if (toVal < forVal)
                            {
                                _pc = (int)_forToNextIndex[_pc] + 1;
                            }
                            else
                            {
                                _pc++;
                            }
                        }
                        // step値 < 0
                        else
                        {
                            // FOR制御に入る段階で初期値fromが限界値toを下回っている場合はループを実行しない
                            if (forVal < toVal)
                            {
                                _pc = (int)_forToNextIndex[_pc] + 1;
                            }
                            else
                            {
                                _pc++;
                            }
                        }
                    }
                    else if (name == "NEXT")
                    {
                        int forIndex = (int)_nextToForIndex[_pc];
                        ArrayList forCmd = (ArrayList)_source[forIndex];
                        string varName = (string)forCmd[1];

                        if (cmd[1] != null && (string)cmd[1] != varName)
                        {
                            throw new Exception("実行時エラーです。FORとNEXTの関係が一致しません。");
                        }

                        Stack s2 = (Stack)forCmd[2];
                        Stack s3 = (Stack)forCmd[3];
                        Stack s4 = (Stack)forCmd[4];
                        double forVal = Calc(ref s2, ref _variableData);
                        double toVal = Calc(ref s3, ref _variableData);
                        double stepVal = Calc(ref s4, ref _variableData);
                        double now = GetVariableVal(ref _variableData, varName) + stepVal;
                        SetVariableVal(ref _variableData, varName, now);

                        // step値 > 0
                        if (stepVal > 0)
                        {
                            // 変数の値が限界値toを上回っている場合はループの先頭に戻る
                            if (now <= toVal)
                            {
                                _pc = (int)_nextToForIndex[_pc] + 1;
                            }
                            else
                            {
                                _pc++;
                            }
                        }
                        // step値 < 0
                        else
                        {
                            // 変数の値が限界値toを下回っている場合はループの先頭に戻る
                            if (toVal <= now)
                            {
                                _pc = (int)_nextToForIndex[_pc] + 1;
                            }
                            else
                            {
                                _pc++;
                            }
                        }
                    }
                    else if (name == "GOTO")
                    {
                        Stack stack = (Stack)cmd[1];
                        int i = (int)Calc(ref stack, ref _variableData);
                        if (!_lineNoToIndex.ContainsKey(i))
                        {
                            throw new Exception(string.Format("実行時エラーです。行番号{0}が見つかりません。", i));
                        }
                        _pc = (int)_lineNoToIndex[i];
                    }
                    else if (name == "GOSUB")
                    {
                        Stack stack = (Stack)cmd[1];
                        int i = (int)Calc(ref stack, ref _variableData);
                        if (!_lineNoToIndex.ContainsKey(i))
                        {
                            throw new Exception(string.Format("実行時エラーです。行番号{0}が見つかりません。", i));
                        }
                        _gosubInfo.Push(_pc);
                        _pc = (int)_lineNoToIndex[i];
                    }
                    else if (name == "RANDOMIZE")
                    {
                        _rnd = new Random();
                        _pc++;
                    }
                    else if (name == "RETURN")
                    {
                        if (_gosubInfo.Count == 0)
                        {
                            throw new Exception("実行時エラーです。これ以上RETURNできません。");
                        }
                        _pc = (int)_gosubInfo.Pop() + 1;
                    }
                    else if (name == "IF")
                    {
                        Stack s1 = (Stack)cmd[1];
                        Stack s3 = (Stack)cmd[3];
                        double valLeft = Calc(ref s1, ref _variableData);
                        double valRight = Calc(ref s3, ref _variableData);
                        string op = (string)cmd[2];

                        switch (op)
                        {
                            case "=":
                                if (AlmostEqual(valLeft, valRight))
                                {
                                    _pc++;
                                }
                                else
                                {
                                    _pc = (int)_ifToNextLineIndex[_pc];
                                }
                                break;
                            case ">":
                                if (valLeft > valRight)
                                {
                                    _pc++;
                                }
                                else
                                {
                                    _pc = (int)_ifToNextLineIndex[_pc];
                                }
                                break;
                            case "<":
                                if (valLeft < valRight)
                                {
                                    _pc++;
                                }
                                else
                                {
                                    _pc = (int)_ifToNextLineIndex[_pc];
                                }
                                break;
                            case ">=":
                                if (AlmostEqual(valLeft, valRight) || valLeft > valRight)
                                {
                                    _pc++;
                                }
                                else
                                {
                                    _pc = (int)_ifToNextLineIndex[_pc];
                                }
                                break;
                            case "<=":
                                if (AlmostEqual(valLeft, valRight) || valLeft < valRight)
                                {
                                    _pc++;
                                }
                                else
                                {
                                    _pc = (int)_ifToNextLineIndex[_pc];
                                }
                                break;
                            case "<>":
                                if (!AlmostEqual(valLeft, valRight))
                                {
                                    _pc++;
                                }
                                else
                                {
                                    _pc = (int)_ifToNextLineIndex[_pc];
                                }
                                break;
                            default:
                                throw new Exception(string.Format("実行時エラーです。処理不可能な比較演算子{0}が出現しました。", op));
                        }
                    }
                    else if (name == "INPUT")
                    {
                        //メッセージ表示
                        string msg = "? ";

                        if (cmd[1] != null)
                        {
                            msg = (string)cmd[1];
                        }

                        //入力する変数を取得
                        //1	 2	 3 4 5
                        //0	 1	 2 3 4
                        //INPUT "MSG" A B C
                        int len = cmd.Count;
                        ArrayList vars = new ArrayList();

                        for (int i = 2; i < len; i++)
                        {
                            vars.Add(cmd[i]);
                        }

                        //入力した数式が計算可能で、変数の数と一致するまで入力を続ける
                        bool result = false;
                        len = vars.Count;

                        do
                        {
                            Console.Write(msg);
                            string input = Console.ReadLine();

                            if (input == null)
                            {
                                _running = false;
                                _pc++;
                                break;
                            }

                            input = input.Trim().ToUpper();

                            try
                            {
                                int index = 0;
                                ArrayList exp = GetStatement("INPUT_LINE", input, ref index, false);
                                if (exp == null)
                                {
                                    continue;
                                }
                                if (exp.Count != len)
                                {
                                    continue;
                                }
                                //変数の数だけ数式の取得を試みる
                                for (int i = 0; i < len; i++)
                                {
                                    ArrayList expi = (ArrayList)exp[i];
                                    Stack stack = (Stack)expi[1];
                                    double v = Calc(ref stack, ref _variableData);
                                    SetVariableVal(ref _variableData, (string)vars[i], v);
                                    result = true;
                                }
                            }
                            catch
                            {
                                result = false;
                                continue;
                            }

                        } while (!result);

                        _pc++;
                    }
                    else if (name == "ON")
                    {
                        //12 3	 4	 5	 6
                        //01 2	 3	 4	 5
                        //		 r=1	 2	 3
                        //ON EXP GOxxx EXP1, EXP2, EXP3
                        Stack s1 = (Stack)cmd[1];
                        int exp = (int)Calc(ref s1, ref _variableData);

                        if (exp <= 0)
                        {
                            _pc++;
                        }
                        else
                        {
                            if (cmd.Count - 1 < 2 + exp)
                            {
                                _pc++;
                            }
                            else
                            {
                                Stack s2 = (Stack)cmd[2 + exp];
                                int gono = (int)Calc(ref s2, ref _variableData);
                                if (_lineNoToIndex.ContainsKey(gono))
                                {
                                    string type = (string)cmd[2];
                                    int npc = (int)_lineNoToIndex[gono];
                                    if (type == "GOTO")
                                    {
                                        _pc = npc;
                                    }
                                    else
                                    {
                                        //GOSUB
                                        _gosubInfo.Push(_pc);
                                        _pc = npc;
                                    }
                                }
                                else
                                {
                                    throw new Exception(string.Format("実行時エラーです。行番号{0}が見つかりません。", gono));
                                }
                            }
                        }
                    }
                    else if (name == "READ")
                    {
                        string varName = (string)cmd[1];
                        if (_dataIndex > _data.Count - 1)
                        {
                            throw new Exception("実行時エラーです。これ以上データを読み込むことができません。");
                        }
                        Stack stack = (Stack)_data[_dataIndex];
                        double val = Calc(ref stack, ref _variableData);
                        SetVariableVal(ref _variableData, varName, val);
                        _dataIndex++;
                        _pc++;
                    }
                    else if (name == "RESTORE")
                    {
                        if (cmd[1] == null)
                        {
                            _dataIndex = 0;
                        }
                        else
                        {
                            Stack stack = (Stack)cmd[1];
                            int i = (int)Calc(ref stack, ref _variableData);
                            if (!_lineNoToDataIndex.ContainsKey(i))
                            {
                                throw new Exception(string.Format("実行時エラーです。行番号{0}が見つかりません。", i));
                            }
                            _dataIndex = (int)_lineNoToDataIndex[i];
                        }
                        _pc++;
                    }
                    else if (name == "PRINT")
                    {
                        int len = (cmd.Count - 1) / 3;
                        int format = -1;

                        if (len == 0)
                        {
                            Console.WriteLine();
                        }

                        for (int i = 0; i < len; i++)
                        {
                            object o = cmd[1 + i * 3];
                            string type = (string)cmd[2 + i * 3];
                            string terminater = (string)cmd[3 + i * 3];
                            if (type == "STRING")
                            {
                                Console.Write((string)o);
                            }
                            else
                            {
                                Stack stack = (Stack)o;
                                double d = Calc(ref stack, ref _variableData);
                                if (type == "EXPRESSION")
                                {
                                    string s = DelZero(d.ToString());
                                    if (format < 0)
                                    {
                                        Console.Write(s);
                                    }
                                    else
                                    {
                                        Console.Write(s.PadLeft(format, ' '));
                                    }
                                }
                                else if (type == "#")
                                {
                                    format = (int)d;
                                    if (format < 1)
                                    {
                                        format = -1;
                                    }
                                }
                                else if (type == "TAB")
                                {
                                    int t = (int)d;
                                    if (Console.CursorLeft < t)
                                    {
                                        Console.CursorLeft = t;
                                    }
                                }
                            }
                            if (terminater == ";")
                            {
                                //何もしない
                            }
                            else if (terminater == ",")
                            {
                                if (format < 0)
                                {
                                    Console.Write("\t");
                                }
                            }
                            else
                            {
                                Console.WriteLine();
                            }
                        }
                        _pc++;
                    }
                    else if (name == "LET")
                    {
                        //1 2	3	4	5
                        //0 1	2	3	4
                        //LET NAME args TYPE EXP

                        string varName = (string)cmd[1];
                        string varType = (string)cmd[3];
                        Stack stack = (Stack)cmd[4];
                        double val = Calc(ref stack, ref _variableData);

                        if (varType == "ARRAY")
                        {
                            ArrayList arr = (ArrayList)cmd[2];
                            int[] args = new int[arr.Count];
                            for (int i = 0; i < arr.Count; i++)
                            {
                                Stack si = (Stack)arr[i];
                                args[i] = (int)Calc(ref si, ref _variableData);
                            }
                            SetArrayVal(varName, args, val);
                        }
                        else
                        {
                            SetVariableVal(ref _variableData, varName, val);
                        }
                        _pc++;
                    }
                    else if (name == "STOP")
                    {
                        _running = false;
                        _pc++;
                    }
                    else if (name == "CLEAR")
                    {
                        Console.Clear();
                        _pc++;
                    }
                    else if (name == "SWAP")
                    {
                        string s1 = (string)cmd[1];
                        string s2 = (string)cmd[2];

                        if (IsArrayName(s1))
                        {
                            throw new Exception(string.Format("実行時エラーです。{0}は配列です。SWAPで配列を入れ替えることができません。", s1));
                        }
                        if (IsArrayName(s2))
                        {
                            throw new Exception(string.Format("実行時エラーです。{0}は配列です。SWAPで配列を入れ替えることができません。", s2));
                        }

                        double tmp = GetVariableVal(ref _variableData, s2);
                        SetVariableVal(ref _variableData, s2, GetVariableVal(ref _variableData, s1));
                        SetVariableVal(ref _variableData, s1, tmp);
                        _pc++;
                    }
                    else if (name == "ERASE")
                    {
                        string arrName = (string)cmd[1];

                        if (IsArrayName(arrName))
                        {
                            int[] dimData = (int[])_arrayDim[arrName];
                            int dimSize = dimData.Length;
                            int b = 1 - _arrayBase;

                            if (dimSize == 1)
                            {
                                int v1 = dimData[0];
                                _arrayData[arrName] = new double[v1 + b];
                            }
                            else if (dimSize == 2)
                            {
                                int v1 = dimData[0];
                                int v2 = dimData[1];
                                _arrayData[arrName] = new double[v1 + b, v2 + b];
                            }
                            else if (dimSize == 3)
                            {
                                int v1 = dimData[0];
                                int v2 = dimData[1];
                                int v3 = dimData[2];
                                _arrayData[arrName] = new double[v1 + b, v2 + b, v3 + b];
                            }
                            _pc++;
                        }
                        else
                        {
                            throw new Exception(string.Format("実行時エラーです。{0}は変数です。ERASEは配列を初期化します。", arrName));
                        }
                    }
                    else if (name == "BASE" || name == "REM" || name == "DATA" || name == "DEF" || name == "DIM")
                    {
                        //これらは実行時には何もしない
                        _pc++;
                    }
                    else
                    {
                        throw new Exception(string.Format("実行時エラーです。未知のコマンド{0}は実行できません。", name));
                    }

                    if (_source.Count - 1 < _pc)
                    {
                        _running = false;
                    }
                }
            }
            catch (Exception e)
            {
                if (lineNo == -1)
                {
                    throw new Exception(e.Message);
                }
                throw new Exception(string.Format("{0} 行番号 {1}", e.Message, lineNo));
            }
        }

        /// <summary>
        /// 数式の計算を行う
        /// </summary>
        /// <param name="original">RPN変換後の数式のスタック</param>
        /// <param name="var">変数データ</param>
        /// <returns>計算結果</returns>
        private double Calc(ref Stack original, ref Hashtable var)
        {
            if (original.Count == 1)
            {
                return GetVariableVal(ref var, original.Peek());
            }

            Stack copy = new Stack(original);
            Stack stack = new Stack();
            Hashtable local = new Hashtable();

            foreach (object token in copy)
            {
                if (copy.Count < 1)
                {
                    throw new Exception("スタックが空のため、数式を計算できません。");
                }

                string str;

                if (token == null)
                {
                    throw new Exception("トークンがnullでした。数式を計算できません。");
                }

                if (token is double)
                {
                    stack.Push(token);
                }
                else if (token is string)
                {
                    str = (string)token;
                    if (IsOp(str))
                    {
                        if (stack.Count < 2)
                        {
                            throw new Exception(string.Format("スタックの数が足りません。{0}演算子の計算には最低{1}必要です。数式を計算できません。", str, 2));
                        }

                        double v2 = (double)stack.Pop();
                        double v1 = (double)stack.Pop();

                        switch (str)
                        {
                            case "*":
                                stack.Push(v1 * v2);
                                break;
                            case "/":
                                stack.Push(v1 / v2);
                                break;
                            case "-":
                                stack.Push(v1 - v2);
                                break;
                            case "+":
                                stack.Push(v1 + v2);
                                break;
                            case "^":
                                stack.Push(Math.Pow(v1, v2));
                                break;
                        }
                    }
                    else if (IsBuildinFnName(str) || str == "#NEG#")
                    {
                        if (stack.Count < 1)
                        {
                            throw new Exception(string.Format("スタックの数が足りません。{0}関数の計算には最低{1}必要です。数式を計算できません。", str, 1));
                        }

                        double v1 = (double)stack.Pop();

                        switch (str)
                        {
                            case "ABS":
                                stack.Push(Math.Abs(v1));
                                break;
                            case "ACOS":
                                stack.Push(Math.Acos(v1));
                                break;
                            case "ASIN":
                                stack.Push(Math.Asin(v1));
                                break;
                            case "ATN":
                                stack.Push(Math.Atan(v1));
                                break;
                            case "COS":
                                stack.Push(Math.Cos(v1));
                                break;
                            case "EXP":
                                stack.Push(Math.Exp(v1));
                                break;
                            case "INT":
                                stack.Push(Math.Round(v1));
                                break;
                            case "LOG":
                                stack.Push(Math.Log(v1));
                                break;
                            case "LOG2":
                                stack.Push(Math.Log(v1, 2.0));
                                break;
                            case "LOG10":
                                stack.Push(Math.Log10(v1));
                                break;
                            case "SGN":
                                if (v1 > 0.0)
                                {
                                    stack.Push(1.0);
                                }
                                else if (v1 < 0.0)
                                {
                                    stack.Push(-1.0);
                                }
                                else
                                {
                                    stack.Push(0.0);
                                }
                                break;
                            case "SIN":
                                stack.Push(Math.Sin(v1));
                                break;
                            case "SQR":
                                stack.Push(Math.Sqrt(v1));
                                break;
                            case "RND":
                                switch (_rndMode)
                                {
                                    case RndMode0ToArg:
                                        stack.Push(1.0 * _rnd.Next(0, 1 + (int)v1));
                                        break;
                                    default:
                                        if ((int)v1 == -1)
                                        {
                                            _rnd = new Random();
                                            _lastRnd = _rnd.NextDouble();
                                        }
                                        else if ((int)v1 == 0)
                                        {
                                            if (_lastRnd < 0)
                                            {
                                                _lastRnd = _rnd.NextDouble();
                                            }
                                        }
                                        else
                                        {
                                            _lastRnd = _rnd.NextDouble();
                                        }
                                        stack.Push(_lastRnd);
                                        break;
                                }
                                break;
                            case "TAN":
                                stack.Push(Math.Tan(v1));
                                break;
                            case "#NEG#":
                                stack.Push(v1 * -1.0);
                                break;
                            default:
                                if (stack.Count < 1)
                                {
                                    throw new Exception(string.Format("スタックの数が足りません。関数{0}の計算には最低{1}必要です。数式を計算できません。", str, 2));
                                }

                                double v2 = (double)stack.Pop();

                                switch (str)
                                {
                                    case "MAX":
                                        stack.Push(Math.Max(v1, v2));
                                        break;
                                    case "MIN":
                                        stack.Push(Math.Min(v1, v2));
                                        break;
                                    case "MOD":
                                        stack.Push(1.0 * (v2 % v1));
                                        break;
                                }
                                break;
                        }
                    }
                    else if (IsArrayName(str))
                    {
                        //配列の次元数を取得する
                        int dimSize = ((int[])_arrayDim[str]).Length;

                        if (stack.Count < dimSize)
                        {
                            throw new Exception(string.Format("配列の添字の数が足りません。配列{0}には最低{1}必要です。数式を計算できません。", str, dimSize));
                        }

                        int[] args = new int[dimSize];

                        for (int i = dimSize - 1; i > -1; i--)
                        {
                            args[i] = (int)(double)stack.Pop();
                        }

                        //配列の次元数個スタックから取り出す。
                        stack.Push(GetArrayVal(str, args));
                    }
                    else if (IsUserFnName(str))
                    {
                        foreach (string key in var.Keys)
                        {
                            local[key] = var[key];
                        }

                        int fnIndex = _userFn.IndexOf(str);
                        Stack fnExp = (Stack)_userFnExp[fnIndex];
                        ArrayList fnArgs = (ArrayList)_userFnArgs[fnIndex];

                        if (stack.Count < fnArgs.Count)
                        {
                            throw new Exception(string.Format("スタックの数が足りません。ユーザー定義関数{0}の計算には最低{1}必要です。数式を計算できません。", str, fnArgs.Count));
                        }

                        foreach (string s in fnArgs)
                        {
                            local[s] = (double)stack.Pop();
                        }

                        stack.Push(Calc(ref fnExp, ref local));
                    }
                    else if (IsVariableName(str))
                    {
                        //配列を先に解決しなければ変数名は配列と同じため計算ミスが発生する
                        stack.Push(GetVariableVal(ref var, token));
                    }
                    else
                    {
                        throw new Exception(string.Format("数式に未知のトークン{0}が出現しました。数式を計算できません。", str));
                    }
                }
                else
                {
                    throw new Exception(string.Format("数式に未知のトークン{0}が出現しました。数式を計算できません。", token));
                }
            }

            if (stack.Count != 1)
            {
                throw new Exception("計算終了時のスタックに計算結果以外のデータが存在します。数式を計算できません。数式が正しくなかったり、関数の引数の数や配列の次元が違っているかもしれません。");
            }

            double result = GetVariableVal(ref var, stack.Pop());
            return result;
        }

        /// <summary>
        /// 計算式の取得を試みる。失敗した場合はnullを返す。
        /// </summary>
        /// <param name="line">ソースの1行</param>
        /// <param name="index">開始インデックス</param>
        /// <returns>数式をRPN変換した結果のスタック</returns>
        private Stack GetExpression(string line, ref int index)
        {
            // 数式トークン行頭
            const int tokenTop = 0;
            // 数式トークン数値、変数
            const int tokenNumber = 1;
            // 数式トークン左括弧(parenthesis)
            const int tokenLeftParen = 2;
            // 数式トークン右括弧
            const int tokenRightParen = 3;
            // 数式トークン関数
            const int tokenFunc = 4;
            // 数式トークン演算子
            const int tokenOperator = 5;
            // 数式トークンコンマ
            const int tokenComma = 6;

            Skip(line, ref index);
            Stack output = new Stack();
            Stack stack = new Stack();
            //Number,ParenRの後にNumber,ParenL,Funcが出現した場合は数式が終了している
            //最初は行頭
            int lastToken = tokenTop;
            int len = line.Length;

            //最初から行末以降の場合は終了
            if (len <= index)
            {
                return null;
            }

            int i;
            int parenCount = 0;

            for (i = index; i < len; i++)
            {
                char c = char.ToUpper(line[i]);
                if (c == ' ')
                {
                    continue;
                }
                if (IsNumber(c) || c == '.')
                {
                    //数値
                    if (lastToken == tokenNumber || lastToken == tokenRightParen)
                    {
                        break;
                    }
                    if (!GetDecimal(line, ref i, false, out double dec))
                    {
                        return null;
                    }
                    output.Push(dec);
                    lastToken = tokenNumber;
                }
                else if (c == '(')
                {
                    parenCount++;
                    //左括弧
                    if (lastToken == tokenNumber || lastToken == tokenRightParen)
                    {
                        break;
                    }
                    stack.Push("(");
                    lastToken = tokenLeftParen;
                }
                else if (c == ')')
                {
                    //右括弧
                    parenCount--;

                    if (parenCount < 0)
                    {
                        break;
                    }
                    if (lastToken == tokenComma || lastToken == tokenOperator)
                    {
                        return null;
                    }

                    while (stack.Peek() is string && (string)stack.Peek() != "(")
                    {
                        if (stack.Count == 0)
                        {
                            return null;
                        }
                        output.Push(stack.Pop());
                    }

                    //左括弧なので捨てる
                    stack.Pop();

                    if (stack.Count > 0)
                    {
                        object str = stack.Peek();
                        //関数の場合
                        if (str is string && IsFuncOrArrayName((string)str))
                        {
                            output.Push(stack.Pop());
                        }
                    }
                    lastToken = tokenRightParen;
                }
                else if (c == '@')
                {
                    //関数・変数
                    if (lastToken == tokenNumber || lastToken == tokenRightParen)
                    {
                        break;
                    }
                    //@は関数
                    lastToken = tokenFunc;
                    stack.Push("@");
                }
                else if (IsAlphabet(c))
                {
                    //関数・変数
                    if (lastToken == tokenNumber || lastToken == tokenRightParen)
                    {
                        break;
                    }

                    string tmp = GetString(line, ref i, false);

                    if (IsFuncOrArrayName(tmp))
                    {
                        //組み込み関数、ユーザー定義関数、配列の場合は関数として処理
                        lastToken = tokenFunc;
                        tmp = DistinctName(tmp);
                        stack.Push(tmp);
                    }
                    else if (IsVariableName(tmp))
                    {
                        //変数の場合
                        lastToken = tokenNumber;
                        output.Push(tmp);
                    }
                    else
                    {
                        //何か知らない文字列なので数式の解析をやめる
                        break;
                    }
                }
                else if (c == ',')
                {
                    //コンマ
                    if (lastToken == tokenComma)
                    {
                        break;
                    }
                    if (parenCount < 1)
                    {
                        break;
                    }
                    if (lastToken != tokenNumber && lastToken != tokenRightParen)
                    {
                        break;
                    }

                    bool pe = false;

                    while (stack.Count > 0)
                    {
                        object o1 = stack.Peek();
                        if (o1 is string && (string)o1 == "(")
                        {
                            pe = true;
                            break;
                        }
                        output.Push(stack.Pop());
                    }

                    if (!pe)
                    {
                        return null;
                    }

                    lastToken = tokenComma;
                }
                else if (IsOp(c.ToString()))
                {
                    //演算子
                    if (c == '*' || c == '/' || c == '^' || ((c == '-' || c == '+') && (lastToken == tokenNumber || lastToken == tokenRightParen || lastToken == tokenFunc)))
                    {
                        string o1 = c.ToString();
                        //**の読み替え
                        if (c == '*')
                        {
                            int j = i + 1;
                            //*の次にまだ文字がある場合**の可能性を検証する
                            Skip(line, ref j);
                            if (line[j] == '*')
                            {
                                i = j + 1;
                                o1 = "^";
                            }
                        }
                        if (stack.Count > 0)
                        {
                            object o2 = stack.Peek();
                            while (o2 is string && IsOp((string)o2) && (((int)_associativeLaw[o1] == LawLeft && ((int)_precedence[o1] <= (int)_precedence[(string)o2])) || ((int)_associativeLaw[o1] == LawRight && ((int)_precedence[o1] < (int)_precedence[(string)o2]))))
                            {
                                output.Push(o2);
                                stack.Pop();
                                if (stack.Count > 0)
                                {
                                    o2 = stack.Peek();
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }
                        stack.Push(o1);
                        lastToken = tokenOperator;
                    }
                    else if (c == '+' && (lastToken == tokenTop || lastToken == tokenOperator || lastToken == tokenLeftParen || lastToken == tokenComma))
                    {
                        //加算ではない+は何もしないため、トークンを読み飛ばす。
                    }
                    else if (c == '-' && (lastToken == tokenTop || lastToken == tokenOperator || lastToken == tokenLeftParen || lastToken == tokenComma))
                    {
                        //#NEG#と減算の識別
                        //#NEG#
                        if (lastToken == tokenNumber || lastToken == tokenRightParen)
                        {
                            break;
                        }
                        lastToken = tokenFunc;
                        stack.Push("#NEG#");
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    break;
                }
            }

            //スタックの残りがあるなら出力
            while (stack.Count > 0)
            {
                output.Push(stack.Pop());
            }

            if (output.Count == 0)
            {
                return null;
            }

            index = i;
            return output;
        }

        /// <summary>
        /// コマンド名を取得する。
        /// </summary>
        /// <param name="line">ソースの1行</param>
        /// <param name="index">解析開始インデックス</param>
        /// <returns>コマンド名</returns>
        private string GetCommandName(string line, ref int index)
        {
            Skip(line, ref index);

            for (int len = MaxCmdLength; len > -1; len--)
            {
                string str = CutStringByLen(line, index, len).ToUpper();
                if (_statements.ContainsKey(str))
                {
                    index += str.Length;
                    return str;
                }
            }

            return null;
        }

        /// <summary>
        /// コマンドを解析する。失敗した場合はnullを返す。
        /// コマンド名にnullを指定した場合はコマンド名の取得を試みる。
        /// </summary>
        /// <param name="commandName">コマンド名</param>
        /// <param name="line">ソースの1行</param>
        /// <param name="index">解析開始インデックス</param>
        /// <param name="isThen">THENか</param>
        /// <returns>解析結果</returns>
        private ArrayList GetStatement(string commandName, string line, ref int index, bool isThen)
        {
            ArrayList result = new ArrayList();
            Skip(line, ref index);

            if (commandName == null)
            {
                string name = GetCommandName(line, ref index);

                if (name != null)
                {
                    ArrayList tmp = GetStatement(name, line, ref index, false);
                    if (tmp != null)
                    {
                        result.AddRange(tmp);
                        return result;
                    }
                }

                //THENなら、再度GOTOとしてオプション解析を強行する
                if (isThen)
                {
                    ArrayList tmp = GetStatement("GOTO", line, ref index, false);
                    if (tmp != null)
                    {
                        result.AddRange(tmp);
                        return result;
                    }
                }
                return null;
            }

            //文法の解析を開始する
            ArrayList cmd = new ArrayList
            {
                commandName
            };
            string[] options = (string[])_statements[commandName];
            int optlen = options.Length;
            int iLocal = index;
            bool looped = false;
            int iLocalLoopBackup = index;
            bool select = false;
            bool selected = false;
            int iLocalSelectBackup = index;
            bool option = false;
            int iLocalOptionBackup = index;
            bool error = false;
            string str;
            ArrayList array;
            ArrayList flag = new ArrayList();

            for (int optIndex = 0; optIndex < optlen; optIndex++)
            {
                object opt = options[optIndex];
                if (0 <= Array.IndexOf(_elements, opt))
                {
                    string tmp = (string)opt;
                    if (tmp == "Array")
                    {
                        ArrayList arrayresult = GetStatement("Array", line, ref iLocal, false);
                        if (arrayresult == null || arrayresult.Count == 0)
                        {
                            error = true;
                        }
                        else
                        {
                            ArrayList a = ((ArrayList)arrayresult[0]);
                            cmd.Add(a[0]);
                            a.RemoveAt(0);
                            cmd.Add(a);
                        }
                    }
                    else if (tmp == "Loop{")
                    {
                        iLocalLoopBackup = iLocal;
                        looped = true;
                    }
                    else if (tmp == "}Loop")
                    {
                        FindOption(options, ref optIndex, "Loop{", true);
                        optIndex--;
                    }
                    else if (tmp == "Expression")
                    {
                        Stack exp = GetExpression(line, ref iLocal);
                        if (exp == null)
                        {
                            error = true;
                        }
                        else
                        {
                            cmd.Add(exp);
                        }
                    }
                    else if (tmp == "Variable")
                    {
                        str = GetVariableName(line, ref iLocal);
                        if (str == null)
                        {
                            error = true;
                        }
                        else
                        {
                            cmd.Add(str);
                        }
                    }
                    else if (tmp == "RelationalOperator")
                    {
                        //比較演算子2文字の確認
                        Skip(line, ref iLocal);
                        str = CutStringByLen(line, iLocal, 2);
                        if (_rerarionalOperators.ContainsKey(str))
                        {
                            iLocal += 2;
                            cmd.Add(_rerarionalOperators[str]);
                        }
                        else
                        {
                            //比較演算子1文字の確認
                            str = CutStringByLen(line, iLocal, 1);
                            if (_rerarionalOperators.ContainsKey(str))
                            {
                                iLocal++;
                                cmd.Add(_rerarionalOperators[str]);
                            }
                            else
                            {
                                error = true;
                            }
                        }
                    }
                    else if (tmp == "QuotedString")
                    {
                        str = GetQuotString(line, ref iLocal);
                        if (str == null)
                        {
                            error = true;
                        }
                        else
                        {
                            iLocal++;
                            cmd.Add(str);
                        }
                    }
                    else if (tmp == "String")
                    {
                        str = GetString(line, ref iLocal, false);
                        if (str == null)
                        {
                            error = true;
                        }
                        else
                        {
                            iLocal++;
                            cmd.Add(str);
                        }
                    }
                    else if (tmp == "End")
                    {
                        index = iLocal;
                        if (cmd.Count != 0)
                        {
                            result.Add(cmd);
                        }
                        return result;
                    }
                    else if (tmp == "EndOfLine")
                    {
                        index = line.Length;
                        if (cmd.Count != 0)
                        {
                            result.Add(cmd);
                        }
                        return result;
                    }
                    else if (tmp == "Optional{")
                    {
                        option = true;
                    }
                    else if (tmp == "}Optional")
                    {
                        option = false;
                    }
                    else if (tmp == "Select{")
                    {
                        select = true;
                        selected = false;
                    }
                    else if (tmp == "}Select")
                    {
                        select = false;
                        if (!selected)
                        {
                            error = true;
                        }
                    }
                    else if (tmp == "Statement")
                    {
                        array = GetStatement(null, line, ref iLocal, commandName == "IF");
                        if (array == null)
                        {
                            error = true;
                        }
                        else
                        {
                            if (cmd.Count != 0)
                            {
                                result.Add(cmd);
                                cmd = new ArrayList();
                            }
                            result.AddRange(array);
                        }
                    }
                    else if (tmp == "Group{")
                    {
                        if (select)
                        {
                            iLocalSelectBackup = iLocal;
                        }
                        else if (option)
                        {
                            iLocalOptionBackup = iLocal;
                        }
                    }
                    else if (tmp == "}Group")
                    {
                        if (select)
                        {
                            selected = true;
                            FindOption(options, ref optIndex, "}Select", false);
                            optIndex--;
                        }
                        else if (option)
                        {
                            FindOption(options, ref optIndex, "}Optional", false);
                            optIndex--;
                        }
                    }
                    else if (tmp == "Rename")
                    {
                        //次の文字列をコマンドとして再検証する
                        str = options[optIndex + 1];
                        array = GetStatement(str, line, ref iLocal, false);
                        if (array == null)
                        {
                            error = true;
                        }
                        else
                        {
                            cmd = new ArrayList();
                            optIndex++;
                            result.AddRange(array);
                        }
                    }
                    else if (tmp == "Recursive")
                    {
                        if (cmd.Count != 0)
                        {
                            result.Add(cmd);
                            cmd = new ArrayList();
                        }
                        array = GetStatement(commandName, line, ref iLocal, false);
                        if (array == null)
                        {
                            error = true;
                        }
                        else
                        {
                            optIndex++;
                            result.AddRange(array);
                        }
                    }
                    else if (tmp == "DelValue")
                    {
                        cmd.RemoveAt(cmd.Count - 1);
                    }
                    else if (tmp == "AddValue")
                    {
                        object o = options[optIndex + 1];
                        optIndex++;
                        cmd.Add(o);
                    }
                    else if (tmp == "AddExpression")
                    {
                        str = options[optIndex + 1];
                        optIndex++;
                        int x = 0;
                        Stack e = GetExpression(str, ref x);
                        cmd.Add(e);
                    }
                    else if (tmp == "AddNull")
                    {
                        cmd.Add(null);
                    }
                    else if (tmp == "SetFlag")
                    {
                        str = options[optIndex + 1];
                        optIndex++;
                        if (!flag.Contains(str))
                        {
                            flag.Add(str);
                        }
                    }
                    else if (tmp == "CheckFlag")
                    {
                        str = options[optIndex + 1];
                        optIndex++;
                        if (!flag.Contains(str))
                        {
                            error = true;
                        }
                    }
                    else if (tmp == "ResetFlag")
                    {
                        str = options[optIndex + 1];
                        optIndex++;
                        if (flag.Contains(str))
                        {
                            flag.Remove(str);
                        }
                    }
                }
                else if (opt is string)
                {
                    //文字列の一致を検証する
                    str = (string)opt;
                    Skip(line, ref iLocal);
                    string tmp = CutStringByLen(line, iLocal, str.Length).ToUpper();
                    if (str == tmp)
                    {
                        iLocal += str.Length;
                        cmd.Add(str);
                    }
                    else
                    {
                        error = true;
                    }
                }

                if (error)
                {
                    if (select)
                    {
                        iLocal = iLocalSelectBackup;
                        FindOption(options, ref optIndex, "}Group", false);
                        error = false;
                    }
                    else if (option)
                    {
                        iLocal = iLocalOptionBackup;
                        FindOption(options, ref optIndex, "}Group", false);
                        error = false;
                    }
                    else if (looped)
                    {
                        looped = false;
                        iLocal = iLocalLoopBackup;
                        FindOption(options, ref optIndex, "}Loop", false);
                        error = false;
                    }
                    else
                    {
                        return null;
                    }
                }
            }

            index = iLocal;

            if (cmd.Count != 0)
            {
                result.Add(cmd);
            }

            return result;
        }

        /// <summary>
        /// 数の取得を試みる
        /// </summary>
        /// <param name="line">ソースの1行</param>
        /// <param name="index">解析開始インデックス</param>
        /// <param name="natural">自然数か</param>
        /// <param name="result">数</param>
        /// <returns>成功・失敗</returns>
        private bool GetDecimal(string line, ref int index, bool natural, out double result)
        {
            Skip(line, ref index);
            result = 0;
            int len = line.Length;

            //最初から行末以降の場合は終了
            if (len <= index)
            {
                return false;
            }

            string tmp = "";
            int i;
            bool isDecimal = false;
            bool isPlusOrMinus = false;

            for (i = index; i < len; i++)
            {
                //最初の文字一回のみプラスとマイナスを許容する
                if (!natural && !isPlusOrMinus && line[i] == '+' && i == index)
                {
                    isPlusOrMinus = true;
                }
                else if (!natural && !isPlusOrMinus && line[i] == '-' && i == index)
                {
                    tmp += "-";
                    isPlusOrMinus = true;
                }
                else if (IsNumber(line[i]))
                {
                    //数
                    tmp += line[i];
                }
                else if (!natural && !isDecimal && line[i] == '.')
                {
                    //小数点は一回のみ許容する
                    tmp += ".";
                    isDecimal = true;
                }
                else if (line[i] == ' ')
                {
                    break;
                }
                else
                {
                    break;
                }
            }

            if (!double.TryParse(tmp, out result))
            {
                return false;
            }

            index = i - 1;
            return true;
        }

        /// <summary>
        /// ダブルクォーテーションで囲まれた文字列の取得を試みる。
        /// 返す文字列に前後のダブルクォーテーションは含まない。
        /// </summary>
        /// <param name="line">ソースの1行</param>
        /// <param name="index">解析開始インデックス</param>
        /// <returns>文字列</returns>
        private string GetQuotString(string line, ref int index)
        {
            Skip(line, ref index);
            string result = "";
            int len = line.Length;

            //最初から行末以降の場合は終了
            if (len <= index)
            {
                return null;
            }

            int start = index;
            bool escape = false;

            if (len <= start)
            {
                return null;
            }

            char c = line[start];

            if (c != '\"')
            {
                return null;
            }

            start++;

            for (int i = start; i < len; i++)
            {
                c = line[i];
                // 文字はアスキーコード表のアルファベットと記号のみ
                if (!IsChar(c))
                {
                    return null;
                }
                if (i > 1 && c == '\"' && escape)
                {
                    escape = false;
                }
                else if (c == '\\')
                {
                    escape = true;
                }
                else if (c == '\"')
                {
                    index = i;
                    return result;
                }
                result += c;
            }

            return null;
        }

        /// <summary>
        /// 変数名の取得を試みる。失敗した場合はnullを返す。
        /// </summary>
        /// <param name="line">ソースの1行</param>
        /// <param name="index">解析開始インデックス</param>
        /// <returns>変数名</returns>
        private string GetVariableName(string line, ref int index)
        {
            //最初から行末以降の場合は終了
            if (line.Length <= index)
            {
                return null;
            }

            Skip(line, ref index);

            //最初の一文字がアルファベットなら変数候補
            if (IsAlphabet(line[index]))
            {
                string s1 = CutStringByLen(line, index, 1).ToUpper();
                index++;

                //行末なら1文字だけで変数名
                if (line.Length <= index)
                {
                    return s1;
                }

                //次の文字が数字なら二文字変数
                if (IsNumber(line[index]))
                {
                    string s2 = CutStringByLen(line, index, 1);
                    index++;
                    return s1 + s2;
                }

                //一文字目がアルファベットで二文字目が数字ではないなら一文字変数
                return s1;
            }

            return null;
        }

        /// <summary>
        /// 文字列の取得を試みる。失敗した場合はnullを返す。
        /// </summary>
        /// <param name="line">ソースの1行</param>
        /// <param name="index">解析開始インデックス</param>
        /// <param name="withoutNum">数字を含むか</param>
        /// <returns>文字列</returns>
        private string GetString(string line, ref int index, bool withoutNum)
        {
            Skip(line, ref index);
            int len = line.Length;

            //最初から行末以降の場合は終了
            if (len <= index)
            {
                return null;
            }

            char c = char.ToUpper(line[index]);

            if (!IsAlphabet(c))
            {
                return null;
            }

            string result = c.ToString();
            index++;

            for (int i = index; i < len; i++)
            {
                c = char.ToUpper(line[i]);
                if (IsAlphabet(c) || (!withoutNum && IsNumber(c)))
                {
                    result += c;
                }
                else if (c == '.')
                {
                    result += '.';
                    index = i;
                    return result;
                }
                else
                {
                    index = i - 1;
                    return result;
                }
            }

            index = len - 1;
            return result;
        }

        /// <summary>
        /// 配列の値を取得する
        /// </summary>
        /// <param name="name">配列名</param>
        /// <param name="args">配列の添え字</param>
        /// <returns>値</returns>
        private double GetArrayVal(string name, int[] args)
        {
            if (!IsArrayName(name))
            {
                throw new Exception(string.Format("実行時エラーです。配列{0}は定義されていません。", name));
            }

            int[] dim = (int[])_arrayDim[name];

            if (args.Length == 0 || args.Length > 3)
            {
                throw new Exception(string.Format("実行時エラーです。配列{0}の次元数が不正です。0以下や4以上は指定できません。", name));
            }
            if (dim.Length == 0 || dim.Length != args.Length)
            {
                throw new Exception(string.Format("実行時エラーです。配列{0}の要素数が違います。", name));
            }

            for (int i = 0; i < dim.Length; i++)
            {
                //添字を補正する
                args[i] -= _arrayBase;
                if (args[i] < 0 || dim[i] < args[i])
                {
                    throw new Exception(string.Format("実行時エラーです。配列{0}の範囲外がアクセスされました。", name));
                }
            }

            try
            {
                if (dim.Length == 1)
                {
                    double[] d = (double[])_arrayData[name];
                    return d[args[0]];
                }
                if (dim.Length == 2)
                {
                    double[,] d = (double[,])_arrayData[name];
                    return d[args[0], args[1]];
                }
                if (dim.Length == 3)
                {
                    double[,,] d = (double[,,])_arrayData[name];
                    return d[args[0], args[1], args[2]];
                }
            }
            catch (IndexOutOfRangeException)
            {
                throw new Exception(string.Format("実行時エラーです。配列{0}の範囲外がアクセスされました。", name));
            }

            throw new Exception(string.Format("実行時エラーです。配列{0}でエラーが発生しました。", name));
        }

        /// <summary>
        /// 配列の値を設定する
        /// </summary>
        /// <param name="name">配列名</param>
        /// <param name="args">配列の添え字</param>
        /// <param name="val">値</param>
        private void SetArrayVal(string name, int[] args, double val)
        {
            if (!IsArrayName(name))
            {
                throw new Exception(string.Format("実行時エラーです。配列{0}は定義されていません。", name));
            }

            int[] dim = (int[])_arrayDim[name];

            if (args.Length == 0 || args.Length > 3)
            {
                throw new Exception(string.Format("実行時エラーです。配列{0}の次元数が不正です。0以下や4以上は指定できません。", name));
            }

            if (dim.Length == 0 || dim.Length != args.Length)
            {
                throw new Exception(string.Format("実行時エラーです。配列{0}の要素数が違います。", name));
            }

            for (int i = 0; i < dim.Length; i++)
            {
                //添字を補正する
                args[i] -= _arrayBase;
                if (args[i] < 0 || dim[i] < args[i])
                {
                    throw new Exception(string.Format("実行時エラーです。配列{0}の範囲外がアクセスされました。", name));
                }
            }

            try
            {
                if (dim.Length == 1)
                {
                    double[] d = (double[])_arrayData[name];
                    d[args[0]] = val;
                    _arrayData[name] = d;
                }
                else if (dim.Length == 2)
                {
                    double[,] d = (double[,])_arrayData[name];
                    d[args[0], args[1]] = val;
                    _arrayData[name] = d;
                }
                else if (dim.Length == 3)
                {
                    double[,,] d = (double[,,])_arrayData[name];
                    d[args[0], args[1], args[2]] = val;
                    _arrayData[name] = d;
                }
            }
            catch (IndexOutOfRangeException)
            {
                throw new Exception(string.Format("実行時エラーです。配列{0}の範囲外がアクセスされました。", name));
            }
        }

        /// <summary>
        /// 変数の値を取得する
        /// </summary>
        /// <param name="var">変数データ</param>
        /// <param name="v">変数名</param>
        /// <returns>値</returns>
        private double GetVariableVal(ref Hashtable var, object v)
        {
            if (v is double)
            {
                return (double)v;
            }

            string s = (string)v;

            if (IsVariableName(s))
            {
                return var.ContainsKey(s) ? (double)var[s] : 0.0;
            }

            throw new Exception(string.Format("文法エラーです。文字列{0}は変数ではありません。", s));
        }

        /// <summary>
        /// 変数の値を設定する
        /// </summary>
        /// <param name="var">変数データ</param>
        /// <param name="v">変数名</param>
        /// <param name="val">値</param>
        private void SetVariableVal(ref Hashtable var, string v, double val)
        {
            if (var.ContainsKey(v))
            {
                var[v] = val;
            }
            else
            {
                var.Add(v, val);
            }
        }

        /// <summary>
        /// 関数名の正式名称を返す
        /// </summary>
        /// <param name="name">関数名</param>
        /// <returns>関数名の正式名称</returns>
        private string DistinctName(string name)
        {
            if (name == "R.")
            {
                return "RND";
            }
            if (name == "I.")
            {
                return "INT";
            }
            return name;
        }

        /// <summary>
        /// 文字列を開始インデックスの位置から指定の長さで切断する
        /// </summary>
        /// <param name="s">文字列</param>
        /// <param name="start">開始インデックス</param>
        /// <param name="length">長さ</param>
        /// <returns>処理後の文字列</returns>
        private string CutStringByLen(string s, int start, int length)
        {
            return s.Length >= length + start ? s.Substring(start, length) : s.Substring(start);
        }

        /// <summary>
        /// 文字列を開始インデックスの位置から終了インデックスの位置で切断する
        /// </summary>
        /// <param name="s">文字列</param>
        /// <param name="start">開始インデックス</param>
        /// <param name="end">終了インデックス</param>
        /// <returns>処理後の文字列</returns>
        private string CutStrAtEndPos(string s, int start, int end)
        {
            return s.Substring(start, end - start + 1);
        }

        /// <summary>
        /// 浮動小数の文字列の前後の0を自然な形で追加・削除する。
        /// 例えば、0100.12300は100.123に変換し、.123は0.123に変更する。
        /// </summary>
        /// <param name="s">文字列</param>
        /// <returns>処理後の文字列</returns>
        private string DelZero(string s)
        {
            int start = 0;
            int end = s.Length - 1;
            int point = s.IndexOf('.');

            while (s[start] == '0')
            {
                start++;
                if (s.Length == start)
                {
                    return "0";
                }
            }

            while (-1 < point && point < end && s[end] == '0')
            {
                end--;
                if (end == 0)
                {
                    return "0";
                }
            }

            if (point == end)
            {
                end--;
                if (end < start)
                {
                    return "0";
                }
            }

            return (start == point ? "0" : "") + CutStrAtEndPos(s, start, end);
        }

        /// <summary>
        /// 文字列のインデックスを空白文字を無視した位置まで移動させる
        /// </summary>
        /// <param name="line">文字列</param>
        /// <param name="index">インデックス</param>
        private void Skip(string line, ref int index)
        {
            int len = line.Length;

            if (len <= index)
            {
                return;
            }

            while (char.IsWhiteSpace(line[index]))
            {
                index++;
                if (len <= index)
                {
                    return;
                }
            }
        }

        /// <summary>
        /// 文法の解析中位置から前後方向へ指定のエレメントを検索する
        /// </summary>
        /// <param name="options">文法</param>
        /// <param name="index">解析中のインデックス</param>
        /// <param name="target">検索対象のエレメント</param>
        /// <param name="isBackSearch">前方へ検索するか</param>
        private void FindOption(string[] options, ref int index, string target, bool isBackSearch)
        {
            int step = isBackSearch ? -1 : 1;
            int i = index;

            while (0 <= i && i < options.Length)
            {
                if (0 <= Array.IndexOf(_elements, options[i]) && options[i] == target)
                {
                    index = i;
                    return;
                }
                i += step;
            }
        }

        /// <summary>
        /// 文字列が関数や配列か調べる。
        /// </summary>
        /// <param name="name">文字列</param>
        /// <returns>関数や配列か</returns>
        private bool IsFuncOrArrayName(string name)
        {
            return IsBuildinFnName(name) || IsUserFnName(name) || IsArrayName(name);
        }

        /// <summary>
        /// 文字列が組み込み関数名か調べる。
        /// 配列の参照と省略形も対象とする。
        /// </summary>
        /// <param name="name">文字列</param>
        /// <returns>組み込み関数名か</returns>
        private bool IsBuildinFnName(string name)
        {
            return 0 <= Array.IndexOf(_buildinFnNames, name) || name == "I." || name == "R.";
        }

        /// <summary>
        /// 文字列がユーザー定義関数名か調べる。
        /// </summary>
        /// <param name="name">文字列</param>
        /// <returns>ユーザー定義関数名か</returns>
        private bool IsUserFnName(string name)
        {
            return _userFn.Contains(name);
        }

        /// <summary>
        /// 文字列が配列名か調べる
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private bool IsArrayName(string name)
        {
            return _arrayDim.ContainsKey(name);
        }

        /// <summary>
        /// 文字列が比較演算子か調べる。
        /// 比較演算子は処理上の標準形に変換される
        /// </summary>
        /// <param name="name">文字列</param>
        /// <returns>比較演算子か</returns>
        private bool IsOp(string name)
        {
            return _precedence.ContainsKey(name);
        }

        /// <summary>
        /// 文字列が変数名か調べる
        /// </summary>
        /// <param name="s">文字列</param>
        /// <returns>変数名か</returns>
        private bool IsVariableName(string s)
        {
            int len = s.Length;

            if (len == 1)
            {
                return IsAlphabet(s[0]);
            }
            if (len == 2)
            {
                return IsAlphabet(s[0]) && IsNumber(s[1]);
            }

            // 1,2ではない値または0の場合
            return false;
        }

        /// <summary>
        /// 文字がアルファベットか調べる
        /// </summary>
        /// <param name="c">文字</param>
        /// <returns>アルファベットか</returns>
        private bool IsAlphabet(char c)
        {
            return 'A' <= c && c <= 'Z';
        }

        /// <summary>
        /// 文字が数字か調べる
        /// </summary>
        /// <param name="c">文字</param>
        /// <returns>数字か</returns>
        private bool IsNumber(char c)
        {
            return '0' <= c && c <= '9';
        }

        /// <summary>
        /// 文字が英数字または記号か調べる
        /// </summary>
        /// <param name="c">文字</param>
        /// <returns>英数字または記号か</returns>
        private bool IsChar(char c)
        {
            return ' ' <= c && c <= '~';
        }

        /// <summary>
        /// Main
        /// </summary>
        /// <param name="args">コマンド引数</param>
        private static void Main(string[] args)
        {
            //-WAITオプション
            bool wait = false;
            Basic basic = new Basic();

            try
            {
                int len = args.Length;
                string fileName = null;

                if (len == 0)
                {
                    throw new Exception("コマンド引数が不正です。引数には最低でもソースファイル名が必要です。");
                }
                if (len > 4)
                {
                    throw new Exception("コマンド引数が不正です。引数は最大で4個です。");
                }

                for (int i = 0; i < len; i++)
                {
                    string temp = args[i].ToUpper().Replace('/', '-');
                    switch (temp)
                    {
                        case "-TRON":
                            basic._traceMode = true;
                            break;
                        case "-WAIT":
                            wait = true;
                            break;
                        case "-ZERO_TO_ONE":
                            basic._rndMode = RndMode0To1;
                            break;
                        case "-ZERO_TO_ARG":
                            basic._rndMode = RndMode0ToArg;
                            break;
                        default:
                            fileName = args[i];
                            break;
                    }
                }

                basic.LoadSource(fileName);
                basic.Run();
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.Message);
            }
            finally
            {
                if (wait)
                {
                    Console.WriteLine("何かキーを押せば終了します。");
                    Console.ReadKey();
                }
            }
        }
    }
}
