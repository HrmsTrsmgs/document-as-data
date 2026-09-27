using System.Text.RegularExpressions;

namespace Marimo.DocumentAsData;

public partial class DatePicker
{
    /// <summary>
    /// 日付選択のWord書式を.NET書式へ変換します。
    /// OOXML要素や表示言語の探索には関与せず、書式文字列だけを扱います。
    /// </summary>
    static partial class DisplayFormat
    {
        /// <summary>
        /// 単一引用符で囲まれた表示文字列を保持し、Wordの年・日・午前午後指定を.NET用に変換します。
        /// 一文字の書式も、.NETの標準書式ではなく指定された一項目として扱います。
        /// スラッシュとコロンは地域別の区切り文字へ変更せず、そのまま表示します。
        /// 対にならない単一引用符は文字として表示します。
        /// 二重引用符は文字として表示し、その内側も書式記号として扱います。
        /// バックスラッシュは.NET用にエスケープして表示に残します。
        /// パーセントは.NETの書式指定として解釈せず、表示に残します。
        /// f・Fも秒の小数部として解釈せず、表示に残します。
        /// K・zzz・t・ttも.NETの時差や午前午後の記号として解釈せず、表示に残します。
        /// yyyyy・YYYYYは四桁年と年の下二桁を続けて表示します。
        /// MMMMM・MMMMMMは月名と月番号を続けて表示します。
        /// ddddd・DDDDDは曜日名と日番号、dddddd・DDDDDDは曜日名と二桁の日番号を続けて表示します。
        /// hhh・HHH・mmm・sssは、二桁指定と一桁指定を続けて表示します。
        /// Wordの書式全体を.NETへ変換するものではありません。
        /// </summary>
        /// <remarks>
        /// 年はyy・yyyy（大文字のYY・YYYYも含む）を確認済みです。
        /// yyyはWordの日付選択で日時・表示が更新されなかったため、対応対象に含めません。
        /// yyy専用の変換や拒否処理は設けていません。
        /// </remarks>
        /// <param name="format">テンプレートの日付表示形式。</param>
        /// <returns>大文字の年・日指定と午前午後指定を.NETの記号へ置き換えた表示形式。</returns>
        internal static string Convert(string format)
        {
            // 引用部分を先に読み、その内部を書式記号として変換しません。午前午後指定はmやMへ分割しません。
            var convertedFormat = Tokens().Replace(
                format,
                it => ConvertToken(it.Value));

            // 一文字書式用の%は、文書に含まれる表示文字の%をエスケープした後で付けます。
            return convertedFormat.Length == 1 ? $"%{convertedFormat}" : convertedFormat;
        }

        /// <summary>
        /// Wordの書式記号を.NET用へ置き換え、文字として表示する記号はエスケープします。
        /// 引用部分は書式記号を変換せず、.NETが解釈するバックスラッシュだけを保護します。
        /// </summary>
        /// <param name="token">正規表現で切り出した書式記号または引用文字列。</param>
        /// <returns>.NETの日付書式へ組み込む文字列。</returns>
        static string ConvertToken(string token) => token switch
        {
            "yyyyy" or "YYYYY" => "yyyy''yy",
            "hhh" => "hh''h",
            "HHH" => "HH''H",
            "mmm" => "mm''m",
            "sss" => "ss''s",
            "MMMMM" => "MMMM''M", // 空の引用文字列で月名と月番号を分け、後続のMは月番号の桁数指定へ残します。
            "ddddd" or "DDDDD" => "dddd''d", // 曜日名と日番号が一つの曜日指定へ結合されないよう区切ります。
            "DDDDDD" => "dddd''dd",
            "am/pm" or "AM/PM" => "tt",
            "/" or ":" or "zzz" => $"'{token}'",
            "'" or "\"" or "\\" or "%" or "f" or "F" or "K" or "t" => $"\\{token}",
            _ => token.StartsWith('\'')
                ? token.Replace("\\", "\\\\")
                : token.ToLowerInvariant()
        };

        /// <summary>
        /// 引用文字列と、.NET用に変換するWordの日付書式記号を識別します。
        /// </summary>
        /// <returns>コンパイル時に生成される、書式記号の検索用正規表現。</returns>
        [GeneratedRegex("'[^']*'|yyyyy|MMMMM|ddddd|hhh|HHH|mmm|sss|am/pm|AM/PM|Y+|D+|/|:|'|\\\\|\"|%|f|F|K|zzz|t")]
        private static partial Regex Tokens();
    }
}
