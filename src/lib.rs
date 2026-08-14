use aviutl2::AnyResult;
use ini::Ini;

mod gui;
mod track;

#[derive(Debug, Clone)]
struct PropertyEntry {
    key: String,
    value: String,
}

#[derive(Debug, Clone)]
struct ApplyItem {
    block_index: usize,
    effect_name: String,
    occurrence: usize,
    property_key: String,
    value: String,
}

#[derive(Debug, Default, Clone, Copy)]
struct ApplySummary {
    target_objects: usize,
    attempted: usize,
    applied: usize,
    skipped_effect_mismatch: usize,
    failed_set: usize,
    adjusted: usize,
}

#[derive(Debug, Default, Clone, Copy)]
struct PasteObjectSummary {
    attempted: usize,
    created: usize,
    failed: usize,
}

#[derive(Debug, Default, Clone, Copy)]
struct TrackPasteSummary {
    targets: usize,
    applied: usize,
    skipped: usize,
    failed: usize,
    adjusted: usize,
}

#[derive(Debug, Clone, Copy)]
struct AliasPlacement {
    layer: usize,
    frame: usize,
}

fn parse_object_section_index(section: &str) -> Option<usize> {
    let lower = section.to_ascii_lowercase();
    let prefix = "object.";
    if !lower.starts_with(prefix) {
        return None;
    }
    section[prefix.len()..].parse::<usize>().ok()
}

fn parse_first_number(value: &str) -> Option<usize> {
    value
        .split(',')
        .next()
        .map(str::trim)
        .filter(|x| !x.is_empty())
        .and_then(|x| x.parse::<usize>().ok())
}

fn extract_alias_placement(alias_text: &str) -> Option<AliasPlacement> {
    let text = alias_text.strip_prefix('\u{feff}').unwrap_or(alias_text);
    let ini = Ini::load_from_str_noescape(text).ok()?;

    let parse_section = |name: &str| {
        let props = ini.section(Some(name))?;
        let layer = props.get("layer")?.trim().parse::<usize>().ok()?;
        let frame = parse_first_number(props.get("frame")?.trim())?;
        Some(AliasPlacement { layer, frame })
    };

    parse_section("Object")
        .or_else(|| parse_section("0"))
        .or_else(|| {
            // [Object] / [0] が無い形式向けに先頭セクションをフォールバックする
            ini.iter().find_map(|(sec, props)| {
                let _ = sec?;
                let layer = props.get("layer")?.trim().parse::<usize>().ok()?;
                let frame = parse_first_number(props.get("frame")?.trim())?;
                Some(AliasPlacement { layer, frame })
            })
        })
}

fn split_aliases_from_clipboard(text: &str) -> Vec<String> {
    let normalized = text.replace("\r\n", "\n").replace('\r', "\n");
    let mut chunks = Vec::new();
    let mut current = String::new();
    let mut saw_object_header = false;

    for line in normalized.lines() {
        let trimmed = line.trim();
        if trimmed.eq_ignore_ascii_case("[Object]") {
            if !current.trim().is_empty() {
                chunks.push(current.trim().to_string());
                current.clear();
            }
            saw_object_header = true;
        }

        if !current.is_empty() {
            current.push('\n');
        }
        current.push_str(line);
    }

    if !current.trim().is_empty() {
        chunks.push(current.trim().to_string());
    }

    if saw_object_header {
        chunks
    } else if normalized.trim().is_empty() {
        Vec::new()
    } else {
        vec![normalized.trim().to_string()]
    }
}

fn apply_relative(base: usize, src: usize, src_base: usize) -> usize {
    let base_i = i64::try_from(base).unwrap_or(i64::MAX);
    let src_i = i64::try_from(src).unwrap_or(i64::MAX);
    let src_base_i = i64::try_from(src_base).unwrap_or(i64::MAX);
    let value = base_i + (src_i - src_base_i);
    value.max(0) as usize
}

fn normalize_path_candidate(value: &str) -> String {
    value
        .replace(['¥', '￥', '＼'], "\\")
        .trim()
        .trim_matches('"')
        .trim_matches('\'')
        .trim()
        .to_string()
}

fn trim_to_known_extension(path: &str) -> String {
    let lower = path.to_ascii_lowercase();
    let known_exts = [
        ".gif", ".png", ".jpg", ".jpeg", ".bmp", ".webp", ".svg", ".mp4", ".mov", ".avi", ".mkv",
        ".wav", ".mp3", ".flac", ".ogg", ".aup2", ".object", ".ini", ".json", ".txt",
    ];

    for ext in known_exts {
        if let Some(pos) = lower.find(ext) {
            let end = pos + ext.len();
            return path[..end].to_string();
        }
    }

    path.to_string()
}

fn extract_embedded_windows_path(value: &str) -> Option<String> {
    let s = normalize_path_candidate(value);
    let bytes = s.as_bytes();
    if bytes.len() < 3 {
        return None;
    }

    let mut start: Option<usize> = None;
    for i in 0..(bytes.len() - 2) {
        if bytes[i].is_ascii_alphabetic()
            && bytes[i + 1] == b':'
            && (bytes[i + 2] == b'\\' || bytes[i + 2] == b'/')
        {
            start = Some(i);
            break;
        }
    }

    let start = start?;
    let mut end = bytes.len();
    for (idx, ch) in s[start..].char_indices() {
        let c = ch;
        if c.is_whitespace() || matches!(c, '"' | '\'' | ')' | ']' | '}' | ',' | ';' | '|') {
            end = start + idx;
            break;
        }
    }

    let raw = s[start..end].trim();
    if raw.is_empty() {
        None
    } else {
        Some(trim_to_known_extension(raw))
    }
}

fn is_windows_path_like(value: &str) -> bool {
    let normalized = normalize_path_candidate(value);
    let s = normalized.trim();
    if s.len() < 3 {
        return false;
    }

    // 例: C:\foo または C:/foo
    let b = s.as_bytes();
    if b[0].is_ascii_alphabetic() && b[1] == b':' && (b[2] == b'\\' || b[2] == b'/') {
        return true;
    }

    // UNC パス: \\server\share
    if s.starts_with("\\\\") {
        return true;
    }

    false
}

fn extract_path_candidates_from_alias(alias_text: &str) -> Vec<String> {
    let text = alias_text.strip_prefix('\u{feff}').unwrap_or(alias_text);
    let ini = match Ini::load_from_str_noescape(text) {
        Ok(v) => v,
        Err(_) => return Vec::new(),
    };

    let mut out = Vec::new();
    let mut seen = std::collections::HashSet::new();

    for (_, props) in ini.iter() {
        for (_, raw_value) in props {
            let whole = normalize_path_candidate(raw_value);
            if is_windows_path_like(&whole) && seen.insert(whole.clone()) {
                out.push(whole);
            }

            if let Some(embedded) = extract_embedded_windows_path(raw_value) {
                if is_windows_path_like(&embedded) && seen.insert(embedded.clone()) {
                    out.push(embedded);
                }
            }

            // カンマ区切りに埋まっているケースも拾う
            for part in raw_value.split(',') {
                let cand = normalize_path_candidate(part);
                if is_windows_path_like(&cand) && seen.insert(cand.clone()) {
                    out.push(cand);
                }
            }
        }
    }

    out
}

fn parse_clipboard_ini_to_apply_items(text: &str) -> Vec<ApplyItem> {
    #[derive(Default)]
    struct TempEffect {
        effect_name: String,
        properties: Vec<PropertyEntry>,
    }

    let text = text.strip_prefix('\u{feff}').unwrap_or(text);
    let ini = match Ini::load_from_str_noescape(text) {
        Ok(v) => v,
        Err(_) => return Vec::new(),
    };

    let mut blocks: Vec<Vec<TempEffect>> = Vec::new();
    let mut current_block: Option<usize> = None;

    for (sec, props) in ini.iter() {
        let Some(section_raw) = sec else {
            continue;
        };
        let section = section_raw.trim();

        if section.eq_ignore_ascii_case("Object") {
            blocks.push(Vec::new());
            current_block = Some(blocks.len() - 1);
            continue;
        }

        if parse_object_section_index(section).is_none() {
            continue;
        }

        if current_block.is_none() {
            blocks.push(Vec::new());
            current_block = Some(blocks.len() - 1);
        }

        let mut temp = TempEffect::default();
        for (k, v) in props {
            let key = k.trim();
            if key.eq_ignore_ascii_case("effect.name") {
                temp.effect_name = v.trim().to_string();
            } else {
                temp.properties.push(PropertyEntry {
                    key: key.to_string(),
                    value: v.to_string(),
                });
            }
        }

        let bi = current_block.expect("block exists");
        blocks[bi].push(temp);
    }

    let mut out = Vec::new();
    for (block_index, effects) in blocks.into_iter().enumerate() {
        let mut occ: std::collections::HashMap<String, usize> = std::collections::HashMap::new();
        for eff in effects {
            if eff.effect_name.trim().is_empty() || eff.properties.is_empty() {
                continue;
            }
            let entry = occ.entry(eff.effect_name.clone()).or_insert(0);
            let occurrence = *entry;
            *entry += 1;

            for p in eff.properties {
                out.push(ApplyItem {
                    block_index,
                    effect_name: eff.effect_name.clone(),
                    occurrence,
                    property_key: p.key,
                    value: p.value,
                });
            }
        }
    }

    out
}

/// トラックバー項目の読み取り結果。
#[derive(Debug, Clone)]
struct TrackTarget {
    /// 現在の設定値。
    param: track::TrackParam,
    /// 必要な値の数（区間数 + 1）。
    value_len: usize,
}

/// 対象の設定項目をトラックバーとして読み取る。
///
/// トラックバー項目でない場合は `None` を返す。
fn read_track_target(
    edit_section: &aviutl2::generic::EditSection,
    object: aviutl2::generic::ObjectHandle,
    effect: &str,
    effect_index: usize,
    item: &str,
) -> Option<TrackTarget> {
    let obj = edit_section.object(object);

    // トラックバー以外の項目では取得自体が失敗する。移動無しの場合はOk(None)になる。
    let info = obj.get_track_info(effect, effect_index, item).ok()?;
    let raw = obj.get_effect_item(effect, effect_index, item).ok()?;

    // 時間制御が無効なら、移動方法より後ろの余りはパラメータと確定できる。
    let has_script_param = match &info {
        Some(info) if !info.timecontrol => Some(true),
        _ => None,
    };

    Some(TrackTarget {
        param: track::parse(&raw, has_script_param),
        value_len: obj.get_section_num().ok()? + 1,
    })
}

/// 適用先がトラックバー項目の場合に、値の個数を対象の区間数へ合わせた値を返す。
///
/// トラックバー項目でない場合と、調整が不要な場合は `None` を返す。
fn adjust_track_value_for_object(
    edit_section: &aviutl2::generic::EditSection,
    object: aviutl2::generic::ObjectHandle,
    item: &ApplyItem,
) -> Option<String> {
    let target = read_track_target(
        edit_section,
        object,
        &item.effect_name,
        item.occurrence,
        &item.property_key,
    )?;

    // 移動無しの値は単一値なので調整の余地が無い。
    let source = track::parse(&item.value, None);
    source.mode.as_ref()?;

    let (values, adjust) = track::adjust_values(&source.values, target.value_len.max(1));
    if !adjust.is_adjusted() {
        return None;
    }

    Some(
        track::TrackParam {
            values,
            ..source
        }
        .to_value_string(),
    )
}

/// クリップボードのテキストからコピー元のトラックバー値を解決する。
///
/// CopyAlias 形式・エイリアス形式・設定値そのものの順に解釈を試みる。
fn resolve_track_source(
    text: &str,
    effect: &str,
    effect_index: usize,
    item: &str,
) -> Option<(track::TrackParam, String)> {
    if track::is_clipboard_text(text) {
        let clip = track::from_clipboard_text(text)?;
        let label = format!("{} / {}", clip.effect, clip.item);
        return Some((clip.param, label));
    }

    // エイリアス形式の場合は同じエフェクト・項目の値を探す。
    let items = parse_clipboard_ini_to_apply_items(text);
    if !items.is_empty() {
        let found = items
            .iter()
            .find(|candidate| {
                candidate.effect_name == effect
                    && candidate.occurrence == effect_index
                    && candidate.property_key == item
            })
            .or_else(|| {
                items
                    .iter()
                    .find(|candidate| candidate.property_key == item)
            })?;

        let label = format!("エイリアス: {} / {}", found.effect_name, found.property_key);
        return Some((track::parse(&found.value, None), label));
    }

    if track::looks_like_value(text) {
        return Some((
            track::parse(text.trim(), None),
            "クリップボードの値".to_string(),
        ));
    }

    None
}

static EDIT_HANDLE: aviutl2::generic::GlobalEditHandle = aviutl2::generic::GlobalEditHandle::new();

#[aviutl2::plugin(GenericPlugin)]
struct CopyAlias;

impl aviutl2::generic::GenericPlugin for CopyAlias {
    fn new(_info: aviutl2::AviUtl2Info) -> AnyResult<Self> {
        Ok(Self)
    }

    fn plugin_info(&self) -> aviutl2::generic::GenericPluginTable {
        aviutl2::generic::GenericPluginTable {
            name: "CopyAlias".to_string(),
            information: format!(
                "CopyAlias {version} by 黒猫大福",
                version = env!("CARGO_PKG_VERSION")
            ),
        }
    }

    fn register(&mut self, registry: &mut aviutl2::generic::HostAppHandle) {
        EDIT_HANDLE.init(registry.create_edit_handle());
        registry.register_menus::<CopyAlias>();
    }
}

impl Drop for CopyAlias {
    fn drop(&mut self) {
        let _ = aviutl2::logger::write_info_log("CopyAlias: プラグインを終了します。");
        // ホスト終了やプラグインアンロード時に、先に開いているダイアログを閉じる。
        gui::close_all_plugin_dialogs();
    }
}

#[aviutl2::generic::menus]
impl CopyAlias {
    #[object(name = "エイリアスをコピー", error = "log_only")]
    fn copy_aliases() -> AnyResult<()> {
        let joined =
            EDIT_HANDLE.call_edit_section(|edit_section| -> AnyResult<Option<String>> {
                let selected_objects = edit_section.get_selected_objects()?;
                if selected_objects.is_empty() {
                    // C++版と同様: 選択が無い場合は何もしない
                    return Ok(None);
                }

                let mut aliases = Vec::new();
                for object in selected_objects {
                    if let Ok(alias) = edit_section.get_object_alias(object) {
                        if !alias.is_empty() {
                            aliases.push(alias);
                        }
                    }
                }

                if aliases.is_empty() {
                    return Ok(None);
                }

                Ok(Some(aliases.join("\r\n")))
            })??;

        let Some(joined) = joined else {
            return Ok(());
        };

        let mut clipboard = arboard::Clipboard::new()
            .map_err(|e| aviutl2::anyhow::anyhow!("クリップボードを開けませんでした: {e}"))?;
        clipboard
            .set_text(joined)
            .map_err(|e| aviutl2::anyhow::anyhow!("クリップボードに書き込めませんでした: {e}"))?;

        Ok(())
    }

    #[object(name = "エイリアスの値をペースト", error = "log_only")]
    fn paste_alias_values() -> AnyResult<()> {
        // メニュー実行時点の選択対象を保持する（ダイアログ表示で選択状態が変わる対策）
        let selected_objects =
            EDIT_HANDLE.call_edit_section(|edit_section| -> AnyResult<_> {
                Ok(edit_section.get_selected_objects()?)
            })??;

        if selected_objects.is_empty() {
            let _ = aviutl2::logger::write_info_log(
                "CopyAlias: 対象オブジェクトを選択してから実行してください。",
            );
            return Ok(());
        }

        let mut clipboard = arboard::Clipboard::new()
            .map_err(|e| aviutl2::anyhow::anyhow!("クリップボードを開けませんでした: {e}"))?;
        let clip = clipboard.get_text().map_err(|e| {
            aviutl2::anyhow::anyhow!("クリップボードのテキスト取得に失敗しました: {e}")
        })?;

        let items = parse_clipboard_ini_to_apply_items(&clip);
        if items.is_empty() {
            let _ = aviutl2::logger::write_info_log(
                "CopyAlias: INI形式を解析できませんでした。表示対象（effect.name付きセクション）がありません。",
            );
            return Ok(());
        }

        let Some(items) = gui::show_apply_dialog(items)? else {
            let _ = aviutl2::logger::write_info_log("CopyAlias: 適用をキャンセルしました。");
            return Ok(());
        };

        if items.is_empty() {
            let _ = aviutl2::logger::write_info_log(
                "CopyAlias: 適用するプロパティが選択されていません。",
            );
            return Ok(());
        }

        let summary =
            EDIT_HANDLE.call_edit_section(|edit_section| -> AnyResult<ApplySummary> {
                let mut summary = ApplySummary {
                    target_objects: selected_objects.len(),
                    ..ApplySummary::default()
                };

                for object in &selected_objects {
                    let obj = edit_section.object(*object);
                    for it in &items {
                        summary.attempted += 1;

                        let count = obj.count_effect(&it.effect_name).unwrap_or(0);
                        if count <= it.occurrence {
                            summary.skipped_effect_mismatch += 1;
                            continue;
                        }

                        // トラックバー項目は中間点数の違いで値の個数がずれるので合わせる。
                        let adjusted = adjust_track_value_for_object(edit_section, *object, it);
                        if adjusted.is_some() {
                            summary.adjusted += 1;
                        }
                        let value = adjusted.as_deref().unwrap_or(&it.value);

                        match obj.set_effect_item(
                            &it.effect_name,
                            it.occurrence,
                            &it.property_key,
                            value,
                        ) {
                            Ok(()) => summary.applied += 1,
                            Err(_) => summary.failed_set += 1,
                        }
                    }
                }

                Ok(summary)
            })??;

        let msg = format!(
            "CopyAlias: 適用対象オブジェクト: {} / 試行数: {} / 適用成功: {} / エフェクト不一致スキップ: {} / 設定失敗: {} / 値数調整: {}",
            summary.target_objects,
            summary.attempted,
            summary.applied,
            summary.skipped_effect_mismatch,
            summary.failed_set,
            summary.adjusted
        );
        let _ = aviutl2::logger::write_info_log(&msg);

        Ok(())
    }

    #[layer(name = "クリップボードからオブジェクト貼り付け", error = "log_only")]
    fn paste_objects_from_clipboard() -> AnyResult<()> {
        // 実行時点の貼り付け基準位置（選択オブジェクト優先）を先に確定する
        let base =
            EDIT_HANDLE.call_edit_section(|edit_section| -> AnyResult<AliasPlacement> {
                let selected = edit_section.get_selected_objects()?;
                if let Some(first) = selected.first() {
                    let lf = edit_section.get_object_layer_frame(*first)?;
                    return Ok(AliasPlacement {
                        layer: lf.layer,
                        frame: lf.start,
                    });
                }

                Ok(AliasPlacement {
                    layer: edit_section.info.layer,
                    frame: edit_section.info.frame,
                })
            })??;

        let mut clipboard = arboard::Clipboard::new()
            .map_err(|e| aviutl2::anyhow::anyhow!("クリップボードを開けませんでした: {e}"))?;
        let clip = clipboard.get_text().map_err(|e| {
            aviutl2::anyhow::anyhow!("クリップボードのテキスト取得に失敗しました: {e}")
        })?;

        let aliases = split_aliases_from_clipboard(&clip);
        if aliases.is_empty() {
            let _ = aviutl2::logger::write_info_log(
                "CopyAlias: クリップボードに有効なエイリアス文字列がありません。",
            );
            return Ok(());
        }

        let source_positions: Vec<Option<AliasPlacement>> = aliases
            .iter()
            .map(|alias| extract_alias_placement(alias))
            .collect();
        let source_base = source_positions
            .iter()
            .flatten()
            .next()
            .copied()
            .unwrap_or(AliasPlacement { layer: 0, frame: 0 });

        let summary =
            EDIT_HANDLE.call_edit_section(|edit_section| -> AnyResult<PasteObjectSummary> {
                let mut summary = PasteObjectSummary {
                    attempted: aliases.len(),
                    ..PasteObjectSummary::default()
                };

                for (i, alias) in aliases.iter().enumerate() {
                    let src = source_positions
                        .get(i)
                        .and_then(|x| *x)
                        .unwrap_or(source_base);

                    let target_layer = apply_relative(base.layer, src.layer, source_base.layer);
                    let target_frame = apply_relative(base.frame, src.frame, source_base.frame);

                    match edit_section.create_object_from_alias(
                        alias,
                        target_layer,
                        target_frame,
                        0,
                    ) {
                        Ok(_) => summary.created += 1,
                        Err(_) => summary.failed += 1,
                    }
                }

                Ok(summary)
            })??;

        let msg = format!(
            "CopyAlias: オブジェクト貼り付け / 試行: {} / 作成成功: {} / 作成失敗: {}",
            summary.attempted, summary.created, summary.failed
        );
        let _ = aviutl2::logger::write_info_log(&msg);

        Ok(())
    }

    #[object_item(name = "トラックバーのパラメータをコピー", error = "log_only")]
    fn copy_track_param(
        object: aviutl2::generic::ObjectHandle,
        effect: &str,
        effect_index: usize,
        item: &str,
    ) -> AnyResult<()> {
        let target = EDIT_HANDLE.call_edit_section(|edit_section| {
            read_track_target(edit_section, object, effect, effect_index, item)
        })?;

        let Some(target) = target else {
            let _ = aviutl2::logger::write_info_log(&format!(
                "CopyAlias: 「{item}」はトラックバー項目ではありません。"
            ));
            return Ok(());
        };

        let value = target.param.to_value_string();
        let text = track::to_clipboard_text(&track::TrackClip {
            effect: effect.to_string(),
            effect_index,
            item: item.to_string(),
            param: target.param,
        });

        let mut clipboard = arboard::Clipboard::new()
            .map_err(|e| aviutl2::anyhow::anyhow!("クリップボードを開けませんでした: {e}"))?;
        clipboard
            .set_text(text)
            .map_err(|e| aviutl2::anyhow::anyhow!("クリップボードに書き込めませんでした: {e}"))?;

        let _ = aviutl2::logger::write_info_log(&format!(
            "CopyAlias: トラックバーのパラメータをコピーしました: {effect} / {item} = {value}"
        ));

        Ok(())
    }

    #[object_item(name = "トラックバーのパラメータをペースト", error = "log_only")]
    fn paste_track_param(
        object: aviutl2::generic::ObjectHandle,
        effect: &str,
        effect_index: usize,
        item: &str,
    ) -> AnyResult<()> {
        // ダイアログ表示で選択状態が変わる可能性があるので、実行時点の情報を先に確定する。
        let (target, selected_objects) = EDIT_HANDLE.call_edit_section(|edit_section| {
            (
                read_track_target(edit_section, object, effect, effect_index, item),
                edit_section.get_selected_objects().unwrap_or_default(),
            )
        })?;

        let Some(target) = target else {
            let _ = aviutl2::logger::write_info_log(&format!(
                "CopyAlias: 「{item}」はトラックバー項目ではありません。"
            ));
            return Ok(());
        };

        let mut clipboard = arboard::Clipboard::new()
            .map_err(|e| aviutl2::anyhow::anyhow!("クリップボードを開けませんでした: {e}"))?;
        let clip = clipboard.get_text().map_err(|e| {
            aviutl2::anyhow::anyhow!("クリップボードのテキスト取得に失敗しました: {e}")
        })?;

        let Some((source_param, source_label)) =
            resolve_track_source(&clip, effect, effect_index, item)
        else {
            let _ = aviutl2::logger::write_info_log(
                "CopyAlias: クリップボードからトラックバーのパラメータを読み取れませんでした。",
            );
            return Ok(());
        };

        let target_label = if effect_index > 0 {
            format!("{effect} ({}) / {item}", effect_index + 1)
        } else {
            format!("{effect} / {item}")
        };

        let Some(response) = gui::show_track_paste_dialog(gui::TrackPasteRequest {
            target_label,
            source_label,
            target_param: target.param,
            source_param: source_param.clone(),
            target_value_len: target.value_len,
            selected_object_count: selected_objects.len(),
        })?
        else {
            let _ =
                aviutl2::logger::write_info_log("CopyAlias: トラックバーの貼り付けをキャンセルしました。");
            return Ok(());
        };

        // 一括適用でも、右クリックした項目のオブジェクトは必ず対象に含める。
        let mut targets = vec![object];
        if response.apply_to_selected {
            for selected in selected_objects {
                if !targets.contains(&selected) {
                    targets.push(selected);
                }
            }
        }

        let summary = EDIT_HANDLE.call_edit_section(move |edit_section| {
            let mut summary = TrackPasteSummary {
                targets: targets.len(),
                ..TrackPasteSummary::default()
            };

            for handle in targets {
                let Some(current) =
                    read_track_target(edit_section, handle, effect, effect_index, item)
                else {
                    summary.skipped += 1;
                    continue;
                };

                let merged = track::merge(
                    &current.param,
                    &source_param,
                    &response.parts,
                    current.value_len,
                );
                if merged.adjust.is_adjusted() {
                    summary.adjusted += 1;
                }

                match edit_section.object(handle).set_effect_item(
                    effect,
                    effect_index,
                    item,
                    &merged.value,
                ) {
                    Ok(()) => summary.applied += 1,
                    Err(_) => summary.failed += 1,
                }
            }

            summary
        })?;

        let _ = aviutl2::logger::write_info_log(&format!(
            "CopyAlias: トラックバー貼り付け / 対象: {} / 適用成功: {} / 項目なしスキップ: {} / 設定失敗: {} / 値数調整: {}",
            summary.targets, summary.applied, summary.skipped, summary.failed, summary.adjusted
        ));

        Ok(())
    }

    #[object(name = "エイリアスからパスをコピー", error = "log_only")]
    fn copy_path_from_alias() -> AnyResult<()> {
        let selected_objects =
            EDIT_HANDLE.call_edit_section(|edit_section| -> AnyResult<_> {
                Ok(edit_section.get_selected_objects()?)
            })??;

        if selected_objects.is_empty() {
            let _ = aviutl2::logger::write_info_log(
                "CopyAlias: 対象オブジェクトを選択してから実行してください。",
            );
            return Ok(());
        }

        let candidates =
            EDIT_HANDLE.call_edit_section(|edit_section| -> AnyResult<Vec<String>> {
                let mut out = Vec::new();
                let mut seen = std::collections::HashSet::new();

                for obj in &selected_objects {
                    if let Ok(alias) = edit_section.get_object_alias(*obj) {
                        for p in extract_path_candidates_from_alias(&alias) {
                            if seen.insert(p.clone()) {
                                out.push(p);
                            }
                        }
                    }
                }

                Ok(out)
            })??;

        if candidates.is_empty() {
            let _ = aviutl2::logger::write_info_log(
                "CopyAlias: エイリアス内にファイル/フォルダパスらしき値が見つかりませんでした。",
            );
            return Ok(());
        }

        let Some(selected_path) = gui::show_path_select_dialog(candidates)? else {
            let _ = aviutl2::logger::write_info_log("CopyAlias: パスコピーをキャンセルしました。");
            return Ok(());
        };

        let mut clipboard = arboard::Clipboard::new()
            .map_err(|e| aviutl2::anyhow::anyhow!("クリップボードを開けませんでした: {e}"))?;
        clipboard
            .set_text(selected_path.clone())
            .map_err(|e| aviutl2::anyhow::anyhow!("クリップボードに書き込めませんでした: {e}"))?;

        let _ = aviutl2::logger::write_info_log(&format!(
            "CopyAlias: パスをクリップボードへコピーしました: {}",
            selected_path
        ));

        Ok(())
    }
}

aviutl2::register_generic_plugin!(CopyAlias);
