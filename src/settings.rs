//! プラグインの動作設定。
//!
//! AviUtl2 のアプリケーションデータフォルダ（通常は `ProgramData\aviutl2`）に置いた
//! `CopyAlias.ini` から読み込む。ファイルが無い場合は既定値で動作し、
//! 編集できるようテンプレートを書き出す。

const FILE_NAME: &str = "CopyAlias.ini";
const SECTION: &str = "CopyAlias";

const TEMPLATE: &str = "\
; CopyAlias の設定
; 変更を反映するには AviUtl2 の再起動が必要です。

[CopyAlias]
; オブジェクト設定の右クリックメニューを「CopyAlias」サブメニューにまとめるか
; 0 = まとめない（既定） / 1 = まとめる
submenu=0
";

/// プラグインの設定。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Settings {
    /// 設定項目メニューを `CopyAlias` サブメニュー配下へまとめるか。
    pub group_item_menus: bool,
}

impl Default for Settings {
    fn default() -> Self {
        Self {
            group_item_menus: false,
        }
    }
}

/// 設定ファイルのパス。
pub fn path() -> std::path::PathBuf {
    aviutl2::config::app_data_path().join(FILE_NAME)
}

fn parse_bool(value: &str) -> bool {
    matches!(
        value.trim().to_ascii_lowercase().as_str(),
        "1" | "true" | "yes" | "on"
    )
}

/// 設定を読み込む。
///
/// 読めない場合は既定値を返す。ファイルが存在しない場合はテンプレートを書き出す。
pub fn load() -> Settings {
    let path = path();

    let Ok(text) = std::fs::read_to_string(&path) else {
        // 書き込みに失敗しても動作には影響しないので無視する。
        if std::fs::write(&path, TEMPLATE).is_ok() {
            let _ = aviutl2::logger::write_info_log(&format!(
                "CopyAlias: 設定ファイルを作成しました: {}",
                path.display()
            ));
        }
        return Settings::default();
    };

    let text = text.strip_prefix('\u{feff}').unwrap_or(&text);
    let Ok(ini) = ini::Ini::load_from_str_noescape(text) else {
        return Settings::default();
    };
    let Some(section) = ini.section(Some(SECTION)) else {
        return Settings::default();
    };

    Settings {
        group_item_menus: section.get("submenu").map(parse_bool).unwrap_or(false),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_bool_accepts_common_forms() {
        for value in ["1", "true", "TRUE", " yes ", "on"] {
            assert!(parse_bool(value), "{value}");
        }
        for value in ["0", "false", "no", "off", "", "xxx"] {
            assert!(!parse_bool(value), "{value}");
        }
    }

    #[test]
    fn template_is_parsed_as_default() {
        let ini = ini::Ini::load_from_str_noescape(TEMPLATE).expect("parsed");
        let section = ini.section(Some(SECTION)).expect("section");
        assert!(!parse_bool(section.get("submenu").expect("submenu")));
    }
}
