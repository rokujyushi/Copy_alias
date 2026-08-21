//! 言語ファイル（`.aul2`）による文言の切り替え。
//!
//! SDK の `CONFIG_HANDLE::translate` を使う。参照されるセクションは
//! **プラグインのファイル名**（`CopyAlias.aux2`）になるため、言語ファイルは
//! 次の形式で記述する。
//!
//! ```ini
//! [CopyAlias.aux2]
//! エイリアスをコピー=Copy Alias
//! ```
//!
//! 未定義のキーは元の日本語がそのまま返る。

/// メニューをまとめるサブメニュー名。
///
/// プラグイン名なので翻訳対象にしない。
pub const MENU_ROOT: &str = "CopyAlias";

/// 現在の言語設定での文言を取得する。
pub fn t(text: &str) -> String {
    aviutl2::config::translate(text)
}

/// `{}` を順に置き換えながら文言を取得する。
///
/// 言語ファイル側でも `{}` の位置を入れ替えられるよう、書式は翻訳後の文字列に対して適用する。
pub fn tf(text: &str, args: &[&str]) -> String {
    let translated = t(text);
    let mut out = String::with_capacity(translated.len());
    let mut rest = translated.as_str();

    for arg in args {
        match rest.find("{}") {
            Some(pos) => {
                out.push_str(&rest[..pos]);
                out.push_str(arg);
                rest = &rest[pos + 2..];
            }
            None => break,
        }
    }

    out.push_str(rest);
    out
}

/// 設定項目メニューの名前を作る。
///
/// `grouped` が真なら `CopyAlias` サブメニュー配下へ入れる。
pub fn item_menu(text: &str, grouped: bool) -> String {
    if grouped {
        format!("{MENU_ROOT}\\{}", t(text))
    } else {
        t(text)
    }
}
