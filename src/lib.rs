use aviutl2::AnyResult;
use eframe::egui;
use egui_ltreeview::{NodeBuilder, TreeView};
use ini::Ini;

fn try_load_japanese_font_bytes() -> Option<Vec<u8>> {
    let candidates = [
        "C:/Windows/Fonts/YuGothM.ttc",
        "C:/Windows/Fonts/YuGothR.ttc",
        "C:/Windows/Fonts/meiryo.ttc",
        "C:/Windows/Fonts/msgothic.ttc",
    ];

    for path in candidates {
        if let Ok(bytes) = std::fs::read(path) {
            return Some(bytes);
        }
    }
    None
}

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

#[derive(Debug, Clone)]
struct ApplyDialogItem {
    checked: bool,
    item: ApplyItem,
}

#[derive(Debug, Clone)]
struct TreeEffectGroup {
    effect_name: String,
    occurrence: usize,
    item_indices: Vec<usize>,
}

#[derive(Debug, Clone)]
struct TreeBlockGroup {
    block_index: usize,
    effects: Vec<TreeEffectGroup>,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
enum ApplyTreeNodeId {
    Block(usize),
    Effect { block: usize, effect_index: usize },
    Item(usize),
}

struct ApplyDialogApp {
    items: Vec<ApplyDialogItem>,
    tree: Vec<TreeBlockGroup>,
    sender: std::sync::mpsc::Sender<Option<Vec<ApplyItem>>>,
}

struct PathSelectApp {
    paths: Vec<String>,
    selected: usize,
    sender: std::sync::mpsc::Sender<Option<String>>,
}

impl eframe::App for PathSelectApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("コピーするパスを選択");
            ui.label("複数の候補が見つかりました。1つ選択してください。");
            ui.separator();

            egui::ComboBox::from_label("パス候補")
                .selected_text(
                    self.paths
                        .get(self.selected)
                        .cloned()
                        .unwrap_or_else(|| "(候補なし)".to_string()),
                )
                .show_ui(ui, |ui| {
                    for (idx, p) in self.paths.iter().enumerate() {
                        ui.selectable_value(&mut self.selected, idx, p);
                    }
                });

            ui.add_space(8.0);
            ui.horizontal(|ui| {
                if ui.button("コピー").clicked() {
                    let selected = self.paths.get(self.selected).cloned();
                    let _ = self.sender.send(selected);
                    ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                }
                if ui.button("キャンセル").clicked() {
                    let _ = self.sender.send(None);
                    ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                }
            });
        });
    }
}

impl ApplyDialogApp {
    fn new(items: Vec<ApplyItem>, sender: std::sync::mpsc::Sender<Option<Vec<ApplyItem>>>) -> Self {
        let items: Vec<ApplyDialogItem> = items
            .into_iter()
            .map(|item| ApplyDialogItem {
                checked: true,
                item,
            })
            .collect();

        let tree = Self::build_tree(&items);

        Self {
            items,
            tree,
            sender,
        }
    }

    fn build_tree(items: &[ApplyDialogItem]) -> Vec<TreeBlockGroup> {
        use std::collections::BTreeMap;

        let mut blocks: BTreeMap<usize, BTreeMap<(String, usize), Vec<usize>>> = BTreeMap::new();

        for (index, item) in items.iter().enumerate() {
            blocks
                .entry(item.item.block_index)
                .or_default()
                .entry((item.item.effect_name.clone(), item.item.occurrence))
                .or_default()
                .push(index);
        }

        blocks
            .into_iter()
            .map(|(block_index, effect_map)| {
                let effects = effect_map
                    .into_iter()
                    .map(
                        |((effect_name, occurrence), item_indices)| TreeEffectGroup {
                            effect_name,
                            occurrence,
                            item_indices,
                        },
                    )
                    .collect();

                TreeBlockGroup {
                    block_index,
                    effects,
                }
            })
            .collect()
    }

    fn set_all_checked(&mut self, checked: bool) {
        for item in &mut self.items {
            item.checked = checked;
        }
    }

    fn block_all_checked(&self, block_index: usize) -> bool {
        self.tree
            .iter()
            .find(|b| b.block_index == block_index)
            .map(|b| {
                b.effects.iter().all(|e| {
                    e.item_indices
                        .iter()
                        .all(|&idx| self.items.get(idx).map(|x| x.checked).unwrap_or(false))
                })
            })
            .unwrap_or(false)
    }

    fn set_block_checked(&mut self, block_index: usize, checked: bool) {
        if let Some(block) = self.tree.iter().find(|b| b.block_index == block_index) {
            for effect in &block.effects {
                for &idx in &effect.item_indices {
                    if let Some(item) = self.items.get_mut(idx) {
                        item.checked = checked;
                    }
                }
            }
        }
    }

    fn effect_all_checked(&self, block_index: usize, effect_index: usize) -> bool {
        self.tree
            .iter()
            .find(|b| b.block_index == block_index)
            .and_then(|b| b.effects.get(effect_index))
            .map(|e| {
                e.item_indices
                    .iter()
                    .all(|&idx| self.items.get(idx).map(|x| x.checked).unwrap_or(false))
            })
            .unwrap_or(false)
    }

    fn set_effect_checked(&mut self, block_index: usize, effect_index: usize, checked: bool) {
        if let Some(effect) = self
            .tree
            .iter()
            .find(|b| b.block_index == block_index)
            .and_then(|b| b.effects.get(effect_index))
        {
            for &idx in &effect.item_indices {
                if let Some(item) = self.items.get_mut(idx) {
                    item.checked = checked;
                }
            }
        }
    }
}

impl eframe::App for ApplyDialogApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("反映するプロパティを選択");
            ui.label("チェックを外した項目は適用しません。");
            ui.separator();

            ui.horizontal(|ui| {
                if ui.button("全選択").clicked() {
                    self.set_all_checked(true);
                }
                if ui.button("全解除").clicked() {
                    self.set_all_checked(false);
                }
            });

            ui.add_space(6.0);
            egui::ScrollArea::vertical()
                .max_height(420.0)
                .show(ui, |ui| {
                    TreeView::new(egui::Id::new("copy_alias_apply_tree"))
                        .allow_multi_selection(false)
                        .min_height(300.0)
                        .show(ui, |builder| {
                            for block in self.tree.clone() {
                                let block_id = ApplyTreeNodeId::Block(block.block_index);
                                let mut block_checked = self.block_all_checked(block.block_index);

                                let open = builder.node(
                                    NodeBuilder::dir(block_id)
                                        .default_open(true)
                                        .label_ui(|ui| {
                                            ui.checkbox(
                                                &mut block_checked,
                                                format!("ブロック {}", block.block_index + 1),
                                            );
                                        }),
                                );

                                if block_checked != self.block_all_checked(block.block_index) {
                                    self.set_block_checked(block.block_index, block_checked);
                                }

                                if open {
                                    for (effect_index, effect) in block.effects.iter().enumerate() {
                                        let effect_id = ApplyTreeNodeId::Effect {
                                            block: block.block_index,
                                            effect_index,
                                        };
                                        let mut effect_checked = self
                                            .effect_all_checked(block.block_index, effect_index);

                                        let effect_label = if effect.occurrence > 0 {
                                            format!(
                                                "{} ({})",
                                                effect.effect_name,
                                                effect.occurrence + 1
                                            )
                                        } else {
                                            effect.effect_name.clone()
                                        };

                                        let effect_open = builder.node(
                                            NodeBuilder::dir(effect_id)
                                                .default_open(true)
                                                .label_ui(|ui| {
                                                    ui.checkbox(
                                                        &mut effect_checked,
                                                        effect_label.clone(),
                                                    );
                                                }),
                                        );

                                        if effect_checked
                                            != self
                                                .effect_all_checked(block.block_index, effect_index)
                                        {
                                            self.set_effect_checked(
                                                block.block_index,
                                                effect_index,
                                                effect_checked,
                                            );
                                        }

                                        if effect_open {
                                            for &item_index in &effect.item_indices {
                                                if let Some(item) = self.items.get_mut(item_index) {
                                                    let leaf_id = ApplyTreeNodeId::Item(item_index);
                                                    let mut checked = item.checked;
                                                    let label = format!(
                                                        "{} = {}",
                                                        item.item.property_key, item.item.value
                                                    );

                                                    builder.node(
                                                        NodeBuilder::leaf(leaf_id).label_ui(|ui| {
                                                            ui.checkbox(
                                                                &mut checked,
                                                                label.clone(),
                                                            );
                                                        }),
                                                    );

                                                    item.checked = checked;
                                                }
                                            }
                                        }
                                        builder.close_dir();
                                    }
                                }

                                builder.close_dir();
                            }
                        });
                });

            ui.separator();
            ui.horizontal(|ui| {
                if ui.button("適用").clicked() {
                    let selected = self
                        .items
                        .iter()
                        .filter(|x| x.checked)
                        .map(|x| x.item.clone())
                        .collect::<Vec<_>>();
                    let _ = self.sender.send(Some(selected));
                    ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                }

                if ui.button("閉じる").clicked() {
                    let _ = self.sender.send(None);
                    ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                }
            });
        });
    }
}

#[derive(Debug, Default, Clone, Copy)]
struct ApplySummary {
    target_objects: usize,
    attempted: usize,
    applied: usize,
    skipped_effect_mismatch: usize,
    failed_set: usize,
}

#[derive(Debug, Default, Clone, Copy)]
struct PasteObjectSummary {
    attempted: usize,
    created: usize,
    failed: usize,
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
        for (raw_key, raw_value) in props {
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

fn show_path_select_dialog(paths: Vec<String>) -> AnyResult<Option<String>> {
    if paths.is_empty() {
        return Ok(None);
    }
    if paths.len() == 1 {
        return Ok(paths.first().cloned());
    }

    let (tx, rx) = std::sync::mpsc::channel::<Option<String>>();

    eframe::run_native(
        "CopyAlias - パス選択",
        eframe::NativeOptions {
            viewport: egui::ViewportBuilder::default()
                .with_inner_size([520.0, 180.0])
                .with_min_inner_size([520.0, 160.0]),
            ..Default::default()
        },
        Box::new(move |cc| {
            if let Some(font_bytes) = try_load_japanese_font_bytes() {
                let mut fonts = egui::FontDefinitions::default();
                fonts.font_data.insert(
                    "jp-ui".to_owned(),
                    std::sync::Arc::new(egui::FontData::from_owned(font_bytes)),
                );
                if let Some(family) = fonts.families.get_mut(&egui::FontFamily::Proportional) {
                    family.insert(0, "jp-ui".to_owned());
                }
                if let Some(family) = fonts.families.get_mut(&egui::FontFamily::Monospace) {
                    family.insert(0, "jp-ui".to_owned());
                }
                cc.egui_ctx.set_fonts(fonts);
            }

            Ok(Box::new(PathSelectApp {
                paths,
                selected: 0,
                sender: tx,
            }))
        }),
    )
    .map_err(|e| aviutl2::anyhow::anyhow!("パス選択ダイアログの起動に失敗しました: {e}"))?;

    match rx.try_recv() {
        Ok(result) => Ok(result),
        Err(_) => Ok(None),
    }
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

fn show_apply_dialog(items: Vec<ApplyItem>) -> AnyResult<Option<Vec<ApplyItem>>> {
    let (tx, rx) = std::sync::mpsc::channel::<Option<Vec<ApplyItem>>>();

    eframe::run_native(
        "CopyAlias - プロパティ選択",
        eframe::NativeOptions {
            viewport: egui::ViewportBuilder::default()
                .with_inner_size([420.0, 480.0])
                .with_min_inner_size([420.0, 480.0]),
            ..Default::default()
        },
        Box::new(move |cc| {
            if let Some(font_bytes) = try_load_japanese_font_bytes() {
                let mut fonts = egui::FontDefinitions::default();
                fonts.font_data.insert(
                    "jp-ui".to_owned(),
                    std::sync::Arc::new(egui::FontData::from_owned(font_bytes)),
                );
                if let Some(family) = fonts.families.get_mut(&egui::FontFamily::Proportional) {
                    family.insert(0, "jp-ui".to_owned());
                }
                if let Some(family) = fonts.families.get_mut(&egui::FontFamily::Monospace) {
                    family.insert(0, "jp-ui".to_owned());
                }
                cc.egui_ctx.set_fonts(fonts);
            }

            Ok(Box::new(ApplyDialogApp::new(items, tx)))
        }),
    )
    .map_err(|e| aviutl2::anyhow::anyhow!("ダイアログの起動に失敗しました: {e}"))?;

    match rx.try_recv() {
        Ok(result) => Ok(result),
        Err(_) => Ok(None),
    }
}

static EDIT_HANDLE: aviutl2::generic::GlobalEditHandle = aviutl2::generic::GlobalEditHandle::new();

#[aviutl2::plugin(GenericPlugin)]
struct CopyAlias;

impl aviutl2::generic::GenericPlugin for CopyAlias {
    fn new(_info: aviutl2::AviUtl2Info) -> AnyResult<Self> {
        Ok(Self)
    }

    fn register(&mut self, registry: &mut aviutl2::generic::HostAppHandle) {
        registry.set_plugin_information(&format!(
            "CopyAlias {version} by 黒猫大福",
            version = env!("CARGO_PKG_VERSION")
        ));
        EDIT_HANDLE.init(registry.create_edit_handle());
        registry.register_menus::<CopyAlias>();
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
                    if let Ok(alias) = edit_section.get_object_alias(&object) {
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

        let Some(items) = show_apply_dialog(items)? else {
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
                    let obj = edit_section.object(object);
                    for it in &items {
                        summary.attempted += 1;

                        let count = obj.count_effect(&it.effect_name).unwrap_or(0);
                        if count <= it.occurrence {
                            summary.skipped_effect_mismatch += 1;
                            continue;
                        }

                        match obj.set_effect_item(
                            &it.effect_name,
                            it.occurrence,
                            &it.property_key,
                            &it.value,
                        ) {
                            Ok(()) => summary.applied += 1,
                            Err(_) => summary.failed_set += 1,
                        }
                    }
                }

                Ok(summary)
            })??;

        let msg = format!(
            "CopyAlias: 適用対象オブジェクト: {} / 試行数: {} / 適用成功: {} / エフェクト不一致スキップ: {} / 設定失敗: {}",
            summary.target_objects,
            summary.attempted,
            summary.applied,
            summary.skipped_effect_mismatch,
            summary.failed_set
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
                    let lf = edit_section.get_object_layer_frame(first)?;
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
                    if let Ok(alias) = edit_section.get_object_alias(obj) {
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

        let Some(selected_path) = show_path_select_dialog(candidates)? else {
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
