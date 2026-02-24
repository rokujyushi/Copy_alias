use aviutl2::AnyResult;
use eframe::egui;
use egui_ltreeview::{NodeBuilder, TreeView};

use crate::ApplyItem;

static DIALOG_CONTEXTS: std::sync::OnceLock<std::sync::Mutex<Vec<egui::Context>>> =
    std::sync::OnceLock::new();

fn dialog_contexts() -> &'static std::sync::Mutex<Vec<egui::Context>> {
    DIALOG_CONTEXTS.get_or_init(|| std::sync::Mutex::new(Vec::new()))
}

fn register_dialog_context(ctx: &egui::Context) {
    if let Ok(mut contexts) = dialog_contexts().lock() {
        let _ = aviutl2::logger::write_info_log("CopyAlias: ダイアログのコンテキストを登録しました。");
        contexts.push(ctx.clone());
    }else {
        let _ = aviutl2::logger::write_warn_log("CopyAlias: ダイアログのコンテキストのロックに失敗しました。");
    }
}

pub(crate) fn close_all_plugin_dialogs() {
    if let Ok(mut contexts) = dialog_contexts().lock() {
        let _ = aviutl2::logger::write_info_log("CopyAlias: すべてのプラグインダイアログを閉じます。");
        for ctx in contexts.iter() {
            ctx.send_viewport_cmd(egui::ViewportCommand::Close);
        }
        contexts.clear();
    } else {
        let _ = aviutl2::logger::write_warn_log("CopyAlias: ダイアログのコンテキストのロックに失敗しました。");
    }
}

#[cfg(windows)]
#[repr(C)]
struct WinPoint {
    x: i32,
    y: i32,
}

#[cfg(windows)]
unsafe extern "system" {
    fn GetCursorPos(lpPoint: *mut WinPoint) -> i32;
}

fn get_cursor_screen_pos() -> Option<egui::Pos2> {
    #[cfg(windows)]
    {
        let mut p = WinPoint { x: 0, y: 0 };
        // SAFETY: GetCursorPosは有効なポインタを要求するため、スタック上のWinPointを渡す。
        let ok = unsafe { GetCursorPos(&mut p as *mut WinPoint) } != 0;
        if ok {
            return Some(egui::Pos2::new(p.x as f32, p.y as f32));
        }
    }
    None
}

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

pub(crate) fn show_path_select_dialog(paths: Vec<String>) -> AnyResult<Option<String>> {
    if paths.is_empty() {
        return Ok(None);
    }
    if paths.len() == 1 {
        return Ok(paths.first().cloned());
    }

    let (tx, rx) = std::sync::mpsc::channel::<Option<String>>();

    let mut viewport = egui::ViewportBuilder::default();
    if let Some(mut pos) = get_cursor_screen_pos() {
        pos = pos - egui::vec2(210.0, 80.0);
        viewport = viewport.with_position(pos);
    }
    viewport = viewport
        .with_inner_size([420.0, 160.0])
        .with_min_inner_size([420.0, 160.0])
        .with_max_inner_size([420.0, 160.0]);

    eframe::run_native(
        "CopyAlias - パス選択",
        eframe::NativeOptions {
            viewport,
            ..Default::default()
        },
        Box::new(move |cc| {
            register_dialog_context(&cc.egui_ctx);

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

pub(crate) fn show_apply_dialog(items: Vec<ApplyItem>) -> AnyResult<Option<Vec<ApplyItem>>> {
    let (tx, rx) = std::sync::mpsc::channel::<Option<Vec<ApplyItem>>>();

    let mut viewport = egui::ViewportBuilder::default();
    if let Some(mut pos) = get_cursor_screen_pos() {
        pos = pos - egui::vec2(210.0, 240.0);
        viewport = viewport.with_position(pos);
    }
    viewport = viewport
        .with_inner_size([420.0, 480.0])
        .with_min_inner_size([420.0, 480.0])
        .with_max_inner_size([420.0, 480.0]);

    eframe::run_native(
        "CopyAlias - プロパティ選択",
        eframe::NativeOptions {
            viewport,
            ..Default::default()
        },
        Box::new(move |cc| {
            register_dialog_context(&cc.egui_ctx);

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
