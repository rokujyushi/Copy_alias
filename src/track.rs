//! エイリアス形式のトラックバー値の分解・再構築。
//!
//! トラックバーの設定値は次のフォーマットでエイリアスに保存される。
//!
//! ```ini
//! 拡大率=100.000                                        ; 移動無し（単一値）
//! X=0.00,100.00,直線移動,3                              ; 値...,移動方法,フラグ
//! X=0.00,50.00,100.00,直線移動,0                        ; 中間点1つ → 値は区間数+1個
//! X=0.0,0.0,ランダム移動,4|0.0001                       ; フラグ|パラメータ
//! X=0.000,0.000,mode@script,6|0|0,0,0,0,1,0.98,0.5,0    ; ...|時間制御データ
//! ```
//!
//! フラグは加速・減速・中間点無視・参照式のビット和になる。

/// 加速のフラグ。
pub const FLAG_ACCELERATE: u32 = 1;
/// 減速のフラグ。
pub const FLAG_DECELERATE: u32 = 2;
/// 中間点無視のフラグ。
pub const FLAG_TWOPOINT: u32 = 4;

const FLAG_SPEED: u32 = FLAG_ACCELERATE | FLAG_DECELERATE;

/// 参照式のフラグ。
///
/// パラメータ欄の右クリックメニューから参照式を設定すると立つ。
/// 実データの `直線移動(時間制御),8|X*2` がこれにあたる。
pub const FLAG_REFERENCE: u32 = 8;

/// 移動そのものに紐づくフラグのビット。
///
/// これ以外のビット（参照式と、今後追加され得る未知のビット）は
/// パラメータ欄の内容と連動するものとして、パラメータと一緒に転写する。
const FLAG_MOVE: u32 = FLAG_ACCELERATE | FLAG_DECELERATE | FLAG_TWOPOINT;

/// 推定時に時間制御データとみなす最小の要素数。
const TIMECONTROL_MIN_FIELDS: usize = 4;

/// 分解済みのトラックバー値。
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TrackParam {
    /// 各区間の値。書式（小数桁数）を保つため文字列のまま保持する。
    pub values: Vec<String>,
    /// 移動方法の名称。`None` は移動無し。
    pub mode: Option<String>,
    /// 加速・減速・中間点無視のビットフラグ。
    pub flags: u32,
    /// パラメータ欄の内容。移動方法のスクリプトパラメータや参照式が入る。
    pub param: Option<String>,
    /// 時間制御データ。
    pub timecontrol: Option<String>,
}

impl TrackParam {
    /// 加速が有効か。
    pub fn accelerate(&self) -> bool {
        self.flags & FLAG_ACCELERATE != 0
    }

    /// 減速が有効か。
    pub fn decelerate(&self) -> bool {
        self.flags & FLAG_DECELERATE != 0
    }

    /// 中間点無視が有効か。
    pub fn twopoint(&self) -> bool {
        self.flags & FLAG_TWOPOINT != 0
    }

    /// 参照式が設定されているか。
    pub fn reference(&self) -> bool {
        self.flags & FLAG_REFERENCE != 0
    }

    /// パラメータ欄に関する情報（内容またはフラグ）を持つか。
    ///
    /// 貼り付けで「パラメータ」を選べるかの判定に使う。
    pub fn has_param_info(&self) -> bool {
        self.param.is_some() || self.flags & !FLAG_MOVE != 0
    }

    /// エイリアス形式の設定値文字列へ戻す。
    pub fn to_value_string(&self) -> String {
        let mut out = self.values.join(",");

        let Some(mode) = self.mode.as_deref() else {
            return out;
        };

        out.push(',');
        out.push_str(mode);
        out.push(',');
        out.push_str(&self.flags.to_string());

        if let Some(param) = &self.param {
            out.push('|');
            out.push_str(param);
        }
        if let Some(timecontrol) = &self.timecontrol {
            // パラメータが無い場合は flags の直後が時間制御データになる。
            out.push('|');
            out.push_str(timecontrol);
        }

        out
    }
}

fn is_number(value: &str) -> bool {
    let trimmed = value.trim();
    !trimmed.is_empty() && trimmed.parse::<f64>().is_ok()
}

/// エイリアス形式の設定値文字列を分解する。
///
/// `has_param` は移動方法がパラメータを持つかどうかのヒント。
/// `None` の場合は値の形からの推定になる（カンマを含むならば時間制御データとみなす）。
pub fn parse(raw: &str, has_param: Option<bool>) -> TrackParam {
    let raw = raw.trim();
    let tokens: Vec<&str> = raw.split(',').collect();

    let mode_index = tokens.iter().position(|token| !is_number(token));

    // 移動方法が見つからない場合と、先頭が数値でない場合は移動無しの単一値として扱う。
    let Some(mode_index) = mode_index.filter(|index| *index > 0) else {
        return TrackParam {
            values: vec![raw.to_string()],
            ..TrackParam::default()
        };
    };

    let values = tokens[..mode_index]
        .iter()
        .map(|token| token.trim().to_string())
        .collect();
    let mode = tokens[mode_index].trim().to_string();

    // 時間制御データはカンマを含むので、移動方法より後ろは一旦結合してから'|'で分ける。
    let spec = tokens[mode_index + 1..].join(",");
    let mut spec_parts = spec.split('|');
    let flags = spec_parts
        .next()
        .and_then(|part| part.trim().parse::<u32>().ok())
        .unwrap_or(0);
    let rest: Vec<&str> = spec_parts.collect();

    let (param, timecontrol) = match (has_param, rest.as_slice()) {
        (_, []) => (None, None),
        (Some(true), [param]) => (Some((*param).to_string()), None),
        (Some(false), [timecontrol]) => (None, Some((*timecontrol).to_string())),
        (None, [single]) => {
            // 時間制御データはカーブを表す数値の並びなので、要素数の多いものを時間制御とみなす。
            if single.split(',').count() >= TIMECONTROL_MIN_FIELDS {
                (None, Some((*single).to_string()))
            } else {
                (Some((*single).to_string()), None)
            }
        }
        (_, [param, timecontrol, ..]) => {
            (Some((*param).to_string()), Some((*timecontrol).to_string()))
        }
    };

    TrackParam {
        values,
        mode: Some(mode),
        flags,
        param,
        timecontrol,
    }
}

/// 値の個数を調整した結果。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ValueAdjust {
    /// 調整なし。
    Keep,
    /// 多かったので先頭から切り詰めた。
    Truncated { from: usize, to: usize },
    /// 少なかったので末尾値で補完した。
    Extended { from: usize, to: usize },
}

impl ValueAdjust {
    /// 調整が発生したか。
    pub fn is_adjusted(&self) -> bool {
        !matches!(self, ValueAdjust::Keep)
    }

    /// ログ・UI表示用の説明文。
    pub fn describe(&self) -> Option<String> {
        match self {
            ValueAdjust::Keep => None,
            ValueAdjust::Truncated { from, to } => {
                Some(format!("値の数を {from} → {to} に切り詰めました"))
            }
            ValueAdjust::Extended { from, to } => {
                Some(format!("値の数を {from} → {to} に補完しました"))
            }
        }
    }
}

/// 値の個数を対象オブジェクトに合わせて調整する。
///
/// 多い場合は先頭から切り詰め、少ない場合は末尾の値で補完する。
pub fn adjust_values(values: &[String], target_len: usize) -> (Vec<String>, ValueAdjust) {
    if values.is_empty() || target_len == 0 || values.len() == target_len {
        return (values.to_vec(), ValueAdjust::Keep);
    }

    if values.len() > target_len {
        return (
            values[..target_len].to_vec(),
            ValueAdjust::Truncated {
                from: values.len(),
                to: target_len,
            },
        );
    }

    let last = values.last().expect("values is not empty").clone();
    let mut out = values.to_vec();
    out.resize(target_len, last);
    (
        out,
        ValueAdjust::Extended {
            from: values.len(),
            to: target_len,
        },
    )
}

/// ペーストで反映する要素の選択。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PasteParts {
    /// 各区間の値。
    pub values: bool,
    /// 移動方法。
    pub mode: bool,
    /// 加速・減速。
    pub speed: bool,
    /// 中間点無視。
    pub twopoint: bool,
    /// パラメータ欄の内容と、それに紐づくフラグ（参照式など）。
    pub param: bool,
    /// 時間制御データ。
    pub timecontrol: bool,
}

impl PasteParts {
    /// 何も選択されていないか。
    pub fn is_empty(&self) -> bool {
        !(self.values || self.mode || self.speed || self.twopoint || self.param || self.timecontrol)
    }
}

impl Default for PasteParts {
    fn default() -> Self {
        Self {
            values: true,
            mode: true,
            speed: true,
            twopoint: true,
            param: true,
            timecontrol: true,
        }
    }
}

/// 合成結果。
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergeResult {
    /// 適用する設定値文字列。
    pub value: String,
    /// 値の個数の調整結果。
    pub adjust: ValueAdjust,
}

/// 対象の現在値をベースに、選択された要素だけコピー元の値で差し替える。
///
/// `target_value_len` は対象オブジェクトの「区間数 + 1」。
pub fn merge(
    dst: &TrackParam,
    src: &TrackParam,
    parts: &PasteParts,
    target_value_len: usize,
) -> MergeResult {
    let mut out = dst.clone();

    if parts.values {
        out.values = src.values.clone();
    }
    if parts.mode {
        out.mode = src.mode.clone();
    }
    if parts.speed {
        out.flags = (out.flags & !FLAG_SPEED) | (src.flags & FLAG_SPEED);
    }
    if parts.twopoint {
        out.flags = (out.flags & !FLAG_TWOPOINT) | (src.flags & FLAG_TWOPOINT);
    }
    if parts.param {
        out.param = src.param.clone();
        // 参照式などパラメータ欄に紐づくフラグも一緒に差し替える。
        out.flags = (out.flags & FLAG_MOVE) | (src.flags & !FLAG_MOVE);
    }
    if parts.timecontrol {
        out.timecontrol = src.timecontrol.clone();
    }

    let adjust = if out.mode.is_some() {
        let (values, adjust) = adjust_values(&out.values, target_value_len.max(1));
        out.values = values;
        adjust
    } else {
        // 移動無しでは値は1つだけで、移動に紐づく情報は保持されない。
        let from = out.values.len();
        out.values.truncate(1);
        out.flags = 0;
        out.param = None;
        out.timecontrol = None;
        if from > 1 {
            ValueAdjust::Truncated { from, to: 1 }
        } else {
            ValueAdjust::Keep
        }
    };

    MergeResult {
        value: out.to_value_string(),
        adjust,
    }
}

/// クリップボードに書き出すトラックバーのコピーデータ。
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TrackClip {
    /// コピー元のエフェクト名。
    pub effect: String,
    /// 同名エフェクト内のインデックス（0始まり）。
    pub effect_index: usize,
    /// コピー元の設定項目名。
    pub item: String,
    /// 分解済みの値。
    pub param: TrackParam,
}

/// クリップボード用セクション名。
pub const CLIP_SECTION: &str = "CopyAlias.TrackParam";

/// クリップボードへ書き出すテキストを組み立てる。
pub fn to_clipboard_text(clip: &TrackClip) -> String {
    let param = &clip.param;
    let mut out = String::new();
    out.push_str(&format!("[{CLIP_SECTION}]\r\n"));
    out.push_str("version=1\r\n");
    out.push_str(&format!("effect={}\r\n", clip.effect));
    out.push_str(&format!("effect_index={}\r\n", clip.effect_index));
    out.push_str(&format!("item={}\r\n", clip.item));
    out.push_str(&format!("value={}\r\n", param.to_value_string()));
    out.push_str(&format!("values={}\r\n", param.values.join(",")));
    out.push_str(&format!(
        "mode={}\r\n",
        param.mode.as_deref().unwrap_or_default()
    ));
    out.push_str(&format!("flags={}\r\n", param.flags));
    if let Some(param) = &param.param {
        out.push_str(&format!("param={param}\r\n"));
    }
    if let Some(timecontrol) = &param.timecontrol {
        out.push_str(&format!("timecontrol={timecontrol}\r\n"));
    }
    out
}

/// テキストが単体のトラックバー設定値として扱えるか。
pub fn looks_like_value(text: &str) -> bool {
    let text = text.trim();
    !text.is_empty() && !text.contains('\n') && text.split(',').next().is_some_and(is_number)
}

/// クリップボードのテキストが CopyAlias 形式かどうか。
pub fn is_clipboard_text(text: &str) -> bool {
    text.contains(&format!("[{CLIP_SECTION}]"))
}

/// CopyAlias 形式のクリップボードテキストを読み取る。
pub fn from_clipboard_text(text: &str) -> Option<TrackClip> {
    let text = text.strip_prefix('\u{feff}').unwrap_or(text);
    let ini = ini::Ini::load_from_str_noescape(text).ok()?;
    let section = ini.section(Some(CLIP_SECTION))?;

    let values: Vec<String> = section
        .get("values")
        .map(|values| {
            values
                .split(',')
                .map(|value| value.trim().to_string())
                .filter(|value| !value.is_empty())
                .collect()
        })
        .unwrap_or_default();

    let mode = section
        .get("mode")
        .map(str::trim)
        .filter(|mode| !mode.is_empty())
        .map(str::to_string);

    Some(TrackClip {
        effect: section.get("effect").unwrap_or_default().trim().to_string(),
        effect_index: section
            .get("effect_index")
            .and_then(|index| index.trim().parse::<usize>().ok())
            .unwrap_or(0),
        item: section.get("item").unwrap_or_default().trim().to_string(),
        param: TrackParam {
            values,
            mode,
            flags: section
                .get("flags")
                .and_then(|flags| flags.trim().parse::<u32>().ok())
                .unwrap_or(0),
            param: section.get("param").map(str::to_string),
            timecontrol: section.get("timecontrol").map(str::to_string),
        },
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn values(param: &TrackParam) -> Vec<&str> {
        param.values.iter().map(String::as_str).collect()
    }

    #[test]
    fn parse_static_value() {
        let param = parse("100.000", None);
        assert_eq!(values(&param), ["100.000"]);
        assert_eq!(param.mode, None);
        assert_eq!(param.to_value_string(), "100.000");
    }

    #[test]
    fn parse_linear_move() {
        let param = parse("0,0,直線移動,0", None);
        assert_eq!(values(&param), ["0", "0"]);
        assert_eq!(param.mode.as_deref(), Some("直線移動"));
        assert_eq!(param.flags, 0);
        assert_eq!(param.param, None);
        assert_eq!(param.timecontrol, None);
        assert_eq!(param.to_value_string(), "0,0,直線移動,0");
    }

    #[test]
    fn parse_with_midpoint() {
        let param = parse("0,0,0,直線移動,3", None);
        assert_eq!(values(&param), ["0", "0", "0"]);
        assert!(param.accelerate());
        assert!(param.decelerate());
        assert!(!param.twopoint());
        assert_eq!(param.to_value_string(), "0,0,0,直線移動,3");
    }

    #[test]
    fn parse_param() {
        let param = parse("0.0,0.0,ランダム移動,4|0.0001", None);
        assert!(param.twopoint());
        assert_eq!(param.param.as_deref(), Some("0.0001"));
        assert_eq!(param.timecontrol, None);
        assert_eq!(param.to_value_string(), "0.0,0.0,ランダム移動,4|0.0001");
    }

    #[test]
    fn parse_timecontrol() {
        let raw = "0.000,0.000,2spt@test_track,6|0|0,0,0,0,1,0.985163,0.5,0";
        let param = parse(raw, Some(true));
        assert_eq!(values(&param), ["0.000", "0.000"]);
        assert_eq!(param.mode.as_deref(), Some("2spt@test_track"));
        assert_eq!(param.flags, 6);
        assert_eq!(param.param.as_deref(), Some("0"));
        assert_eq!(
            param.timecontrol.as_deref(),
            Some("0,0,0,0,1,0.985163,0.5,0")
        );
        assert_eq!(param.to_value_string(), raw);
    }

    #[test]
    fn parse_timecontrol_without_param() {
        let raw = "0,0,___t@test_track,0|0,0,0,0,1,0.5,0.5,0";
        let param = parse(raw, Some(false));
        assert_eq!(param.param, None);
        assert_eq!(param.timecontrol.as_deref(), Some("0,0,0,0,1,0.5,0.5,0"));
        assert_eq!(param.to_value_string(), raw);
    }

    #[test]
    fn parse_auto_hint_uses_field_count() {
        // 要素数が多いものは時間制御データとみなす。
        let param = parse("0,0,mode@script,0|0,0,0,0,1,0.5,0.5,0", None);
        assert_eq!(param.param, None);
        assert_eq!(param.timecontrol.as_deref(), Some("0,0,0,0,1,0.5,0.5,0"));

        // 要素数が少ないものはパラメータとみなす。
        let param = parse("0,0,mode@script,0|1,2", None);
        assert_eq!(param.param.as_deref(), Some("1,2"));
        assert_eq!(param.timecontrol, None);
    }

    #[test]
    fn parse_non_numeric_head_is_single_value() {
        let param = parse("通常", None);
        assert_eq!(values(&param), ["通常"]);
        assert_eq!(param.mode, None);
    }

    #[test]
    fn adjust_values_truncates_and_extends() {
        let src: Vec<String> = ["1", "2", "3"].iter().map(|s| s.to_string()).collect();

        let (adjusted, result) = adjust_values(&src, 2);
        assert_eq!(adjusted, ["1", "2"]);
        assert_eq!(result, ValueAdjust::Truncated { from: 3, to: 2 });

        let (adjusted, result) = adjust_values(&src, 5);
        assert_eq!(adjusted, ["1", "2", "3", "3", "3"]);
        assert_eq!(result, ValueAdjust::Extended { from: 3, to: 5 });

        let (adjusted, result) = adjust_values(&src, 3);
        assert_eq!(adjusted, ["1", "2", "3"]);
        assert_eq!(result, ValueAdjust::Keep);
    }

    #[test]
    fn merge_all_parts_adjusts_value_count() {
        let dst = parse("0.00,0.00,0.00,直線移動,0", None);
        let src = parse("10.00,20.00,補間移動,3", None);
        let merged = merge(&dst, &src, &PasteParts::default(), 3);

        assert_eq!(merged.value, "10.00,20.00,20.00,補間移動,3");
        assert_eq!(merged.adjust, ValueAdjust::Extended { from: 2, to: 3 });
    }

    #[test]
    fn merge_mode_only_keeps_destination_values() {
        let dst = parse("5.00,7.00,直線移動,1", None);
        let src = parse("10.00,20.00,補間移動,2", None);
        let parts = PasteParts {
            values: false,
            mode: true,
            speed: false,
            twopoint: false,
            param: false,
            timecontrol: false,
        };

        let merged = merge(&dst, &src, &parts, 2);
        assert_eq!(merged.value, "5.00,7.00,補間移動,1");
        assert_eq!(merged.adjust, ValueAdjust::Keep);
    }

    #[test]
    fn merge_carries_reference_flag_with_param() {
        // 実データ: パラメータ欄に参照式を設定すると bit3 が立つ。
        let dst = parse("0.00,0.00,直線移動(時間制御),0", None);
        let src = parse("0.00,0.00,直線移動(時間制御),8|X*2", None);
        assert_eq!(src.param.as_deref(), Some("X*2"));
        assert!(src.reference());
        assert!(!dst.reference());

        let parts = PasteParts {
            values: false,
            mode: false,
            speed: false,
            twopoint: false,
            param: true,
            timecontrol: false,
        };
        let merged = merge(&dst, &src, &parts, 2);
        assert_eq!(merged.value, "0.00,0.00,直線移動(時間制御),8|X*2");

        // 逆向き（パラメータを消す）でも参照式フラグが残らないこと。
        let merged = merge(&src, &dst, &parts, 2);
        assert_eq!(merged.value, "0.00,0.00,直線移動(時間制御),0");
    }

    #[test]
    fn merge_keeps_reference_flag_when_param_is_unchecked() {
        let dst = parse("0.00,0.00,直線移動(時間制御),8|X*2", None);
        let src = parse("0.00,0.00,補間移動(時間制御),0", None);
        let parts = PasteParts {
            values: false,
            mode: true,
            speed: true,
            twopoint: true,
            param: false,
            timecontrol: false,
        };

        let merged = merge(&dst, &src, &parts, 2);
        assert_eq!(merged.value, "0.00,0.00,補間移動(時間制御),8|X*2");
    }

    #[test]
    fn merge_speed_only_replaces_speed_bits() {
        let dst = parse("0,0,直線移動,4", None); // 中間点無視のみ
        let src = parse("0,0,直線移動,3", None); // 加速+減速
        let parts = PasteParts {
            values: false,
            mode: false,
            speed: true,
            twopoint: false,
            param: false,
            timecontrol: false,
        };

        let merged = merge(&dst, &src, &parts, 2);
        assert_eq!(merged.value, "0,0,直線移動,7");
    }

    #[test]
    fn merge_static_mode_drops_move_information() {
        let dst = parse("0.00,0.00,直線移動,3|1", None);
        let src = parse("50.00", None);
        let merged = merge(&dst, &src, &PasteParts::default(), 2);

        assert_eq!(merged.value, "50.00");
        assert_eq!(merged.adjust, ValueAdjust::Keep);
    }

    #[test]
    fn clipboard_roundtrip() {
        let clip = TrackClip {
            effect: "標準描画".to_string(),
            effect_index: 1,
            item: "拡大率".to_string(),
            param: parse("0.0,0.0,ランダム移動,4|0.0001", Some(true)),
        };

        let text = to_clipboard_text(&clip);
        assert!(is_clipboard_text(&text));

        let restored = from_clipboard_text(&text).expect("parsed");
        assert_eq!(restored, clip);
    }
}
