use aviutl2::AnyResult;
use eframe::egui;
use egui_ltreeview::{NodeBuilder, TreeView};

use crate::ApplyItem;
use crate::i18n::{t, tf};
use crate::track::{self, PasteParts, TrackParam};

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
#[derive(Clone, Copy)]
struct WinPoint {
    x: i32,
    y: i32,
}

#[cfg(windows)]
#[repr(C)]
#[derive(Clone, Copy, Default)]
struct WinRect {
    left: i32,
    top: i32,
    right: i32,
    bottom: i32,
}

#[cfg(windows)]
#[repr(C)]
#[derive(Clone, Copy, Default)]
struct MonitorInfo {
    cb_size: u32,
    rc_monitor: WinRect,
    rc_work: WinRect,
    dw_flags: u32,
}

#[cfg(windows)]
const MONITOR_DEFAULTTONEAREST: u32 = 2;

#[cfg(windows)]
unsafe extern "system" {
    fn GetCursorPos(lpPoint: *mut WinPoint) -> i32;
    fn MonitorFromPoint(pt: WinPoint, dwFlags: u32) -> *mut std::ffi::c_void;
    fn GetMonitorInfoW(hMonitor: *mut std::ffi::c_void, lpmi: *mut MonitorInfo) -> i32;
}

/// マウスカーソルのスクリーン座標（物理ピクセル）。
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

/// 指定座標に最も近いモニタの作業領域（タスクバーを除いた範囲、物理ピクセル）。
#[allow(unused_variables)]
fn get_work_area(point: egui::Pos2) -> Option<egui::Rect> {
    #[cfg(windows)]
    {
        let pt = WinPoint {
            x: point.x as i32,
            y: point.y as i32,
        };

        // SAFETY: MonitorFromPointは座標に最も近いモニタを返す。取得したハンドルは
        // GetMonitorInfoWへ渡すだけで、解放は不要。
        let monitor = unsafe { MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST) };
        if monitor.is_null() {
            return None;
        }

        let mut info = MonitorInfo {
            cb_size: std::mem::size_of::<MonitorInfo>() as u32,
            ..MonitorInfo::default()
        };
        // SAFETY: cb_sizeを設定した有効なMonitorInfoへのポインタを渡す。
        if unsafe { GetMonitorInfoW(monitor, &mut info as *mut MonitorInfo) } == 0 {
            return None;
        }

        return Some(egui::Rect::from_min_max(
            egui::pos2(info.rc_work.left as f32, info.rc_work.top as f32),
            egui::pos2(info.rc_work.right as f32, info.rc_work.bottom as f32),
        ));
    }

    #[cfg(not(windows))]
    None
}

/// ダイアログウィンドウの配置。
///
/// マウス位置を基準にしつつ、モニタの作業領域からはみ出さないように補正する。
/// 生成時の `with_position` は論理座標として扱われDPIを反映できないため、
/// 実際の補正は最初のフレームで [`egui::ViewportCommand::OuterPosition`]
/// （ウィンドウのDPIで物理座標へ変換される）を送って行う。
#[derive(Debug, Clone, Copy)]
struct DialogWindow {
    size: egui::Vec2,
    offset: egui::Vec2,
    placed: bool,
}

impl DialogWindow {
    fn new(size: [f32; 2], offset: [f32; 2]) -> Self {
        Self {
            size: egui::vec2(size[0], size[1]),
            offset: egui::vec2(offset[0], offset[1]),
            placed: false,
        }
    }

    /// ビューポート定義（位置は生成時の暫定値）。
    ///
    /// 内容が収まらない場合にユーザーが広げられるよう、最大サイズは固定しない。
    fn viewport(&self) -> egui::ViewportBuilder {
        let mut viewport = egui::ViewportBuilder::default();
        if let Some(pos) = self.clamped_position(1.0) {
            viewport = viewport.with_position(pos);
        }
        viewport
            .with_inner_size(self.size)
            .with_min_inner_size(egui::vec2(320.0, 200.0))
    }

    /// 作業領域に収めた表示位置をポイント単位で求める。
    fn clamped_position(&self, pixels_per_point: f32) -> Option<egui::Pos2> {
        let cursor = get_cursor_screen_pos()?;
        let size = self.size * pixels_per_point;
        let offset = self.offset * pixels_per_point;
        let pos = cursor - offset;

        let Some(work) = get_work_area(cursor) else {
            return Some(pos / pixels_per_point);
        };

        // ウィンドウが作業領域より大きい場合は左上を優先する。
        let x = (pos.x).min(work.max.x - size.x).max(work.min.x);
        let y = (pos.y).min(work.max.y - size.y).max(work.min.y);

        Some(egui::pos2(x, y) / pixels_per_point)
    }

    /// 最初のフレームで表示位置を確定する。
    fn ensure_placed(&mut self, ctx: &egui::Context) {
        if self.placed {
            return;
        }
        self.placed = true;

        if let Some(pos) = self.clamped_position(ctx.pixels_per_point()) {
            ctx.send_viewport_cmd(egui::ViewportCommand::OuterPosition(pos));
        }
    }
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

/// ダイアログ共通の初期化（コンテキスト登録と日本語フォントの適用）。
fn setup_dialog(cc: &eframe::CreationContext<'_>) {
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
    window: DialogWindow,
    items: Vec<ApplyDialogItem>,
    tree: Vec<TreeBlockGroup>,
    sender: std::sync::mpsc::Sender<Option<Vec<ApplyItem>>>,
}

struct PathSelectApp {
    window: DialogWindow,
    paths: Vec<String>,
    selected: usize,
    sender: std::sync::mpsc::Sender<Option<String>>,
}

impl eframe::App for PathSelectApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        self.window.ensure_placed(ctx);

        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading(t("コピーするパスを選択"));
            ui.label(t("複数の候補が見つかりました。1つ選択してください。"));
            ui.separator();

            egui::ComboBox::from_label(t("パス候補"))
                .selected_text(
                    self.paths
                        .get(self.selected)
                        .cloned()
                        .unwrap_or_else(|| t("(候補なし)")),
                )
                .show_ui(ui, |ui| {
                    for (idx, p) in self.paths.iter().enumerate() {
                        ui.selectable_value(&mut self.selected, idx, p);
                    }
                });

            ui.add_space(8.0);
            ui.horizontal(|ui| {
                if ui.button(t("コピー")).clicked() {
                    let selected = self.paths.get(self.selected).cloned();
                    let _ = self.sender.send(selected);
                    ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                }
                if ui.button(t("キャンセル")).clicked() {
                    let _ = self.sender.send(None);
                    ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                }
            });
        });
    }
}

impl ApplyDialogApp {
    fn new(
        window: DialogWindow,
        items: Vec<ApplyItem>,
        sender: std::sync::mpsc::Sender<Option<Vec<ApplyItem>>>,
    ) -> Self {
        let items: Vec<ApplyDialogItem> = items
            .into_iter()
            .map(|item| ApplyDialogItem {
                checked: true,
                item,
            })
            .collect();

        let tree = Self::build_tree(&items);

        Self {
            window,
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
        self.window.ensure_placed(ctx);

        // 決定ボタンは下部パネルに固定する。中央に置くと項目数が多いときに
        // ツリーへ押し出されてウィンドウ外に出てしまう。
        egui::TopBottomPanel::bottom("copy_alias_apply_actions").show(ctx, |ui| {
            ui.add_space(6.0);
            ui.horizontal(|ui| {
                if ui.button(t("適用")).clicked() {
                    let selected = self
                        .items
                        .iter()
                        .filter(|x| x.checked)
                        .map(|x| x.item.clone())
                        .collect::<Vec<_>>();
                    let _ = self.sender.send(Some(selected));
                    ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                }

                if ui.button(t("閉じる")).clicked() {
                    let _ = self.sender.send(None);
                    ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                }
            });
            ui.add_space(6.0);
        });

        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading(t("反映するプロパティを選択"));
            ui.label(t("チェックを外した項目は適用しません。"));
            ui.separator();

            ui.horizontal(|ui| {
                if ui.button(t("全選択")).clicked() {
                    self.set_all_checked(true);
                }
                if ui.button(t("全解除")).clicked() {
                    self.set_all_checked(false);
                }
            });

            ui.add_space(6.0);
            egui::ScrollArea::vertical()
                .auto_shrink([false, false])
                .show(ui, |ui| {
                    TreeView::new(egui::Id::new("copy_alias_apply_tree"))
                        .allow_multi_selection(false)
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
                                                tf("ブロック {}", &[&(block.block_index + 1).to_string()]),
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
        });
    }
}

/// トラックバーの貼り付けダイアログへ渡す情報。
#[derive(Debug, Clone)]
pub(crate) struct TrackPasteRequest {
    /// 貼り付け先の表示名（`エフェクト / 項目`）。
    pub target_label: String,
    /// コピー元の表示名。
    pub source_label: String,
    /// 貼り付け先の現在値。
    pub target_param: TrackParam,
    /// コピー元の値。
    pub source_param: TrackParam,
    /// 貼り付け先で必要な値の数（区間数 + 1）。
    pub target_value_len: usize,
    /// 選択中オブジェクトの数。
    pub selected_object_count: usize,
}

/// トラックバーの貼り付けダイアログの結果。
#[derive(Debug, Clone, Copy)]
pub(crate) struct TrackPasteResponse {
    /// 反映する要素。
    pub parts: PasteParts,
    /// 選択中の全オブジェクトへ適用するか。
    pub apply_to_selected: bool,
}

struct TrackPasteApp {
    window: DialogWindow,
    request: TrackPasteRequest,
    available: PasteParts,
    parts: PasteParts,
    apply_to_selected: bool,
    sender: std::sync::mpsc::Sender<Option<TrackPasteResponse>>,
}

impl TrackPasteApp {
    /// 反映できない要素を落とした選択肢を返す。
    ///
    /// コピー元・貼り付け先のどちらかが値を持っていれば選択できる（貼り付け先の
    /// パラメータや時間制御データを消す操作も貼り付けとして成立するため）。
    fn available_parts(source: &TrackParam, target: &TrackParam) -> PasteParts {
        PasteParts {
            values: !source.values.is_empty(),
            mode: true,
            speed: source.mode.is_some(),
            twopoint: source.mode.is_some(),
            param: source.has_param_info() || target.has_param_info(),
            timecontrol: source.timecontrol.is_some() || target.timecontrol.is_some(),
        }
    }

    /// 選択できる要素をまとめてチェック・解除する。
    fn set_all_checked(&mut self, checked: bool) {
        let available = self.available;
        self.parts = PasteParts {
            values: checked && available.values,
            mode: checked && available.mode,
            speed: checked && available.speed,
            twopoint: checked && available.twopoint,
            param: checked && available.param,
            timecontrol: checked && available.timecontrol,
        };
    }

    fn speed_label(param: &TrackParam) -> String {
        match (param.accelerate(), param.decelerate()) {
            (true, true) => t("加速 + 減速"),
            (true, false) => t("加速"),
            (false, true) => t("減速"),
            (false, false) => t("なし"),
        }
    }

    fn checkbox_row(
        ui: &mut egui::Ui,
        enabled: bool,
        checked: &mut bool,
        title: &str,
        value: &str,
    ) {
        if !enabled {
            *checked = false;
        }
        ui.add_enabled_ui(enabled, |ui| {
            ui.horizontal(|ui| {
                ui.checkbox(checked, title);
                ui.label(egui::RichText::new(value).monospace().weak());
            });
        });
    }
}

impl eframe::App for TrackPasteApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        self.window.ensure_placed(ctx);

        // 決定ボタンは下部パネルに固定する（内容が増えても隠れないようにする）。
        egui::TopBottomPanel::bottom("copy_alias_track_actions").show(ctx, |ui| {
            ui.add_space(6.0);
            ui.horizontal(|ui| {
                let can_apply = !self.parts.is_empty();
                if ui
                    .add_enabled(can_apply, egui::Button::new(t("貼り付け")))
                    .clicked()
                {
                    let _ = self.sender.send(Some(TrackPasteResponse {
                        parts: self.parts,
                        apply_to_selected: self.apply_to_selected,
                    }));
                    ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                }

                if ui.button(t("キャンセル")).clicked() {
                    let _ = self.sender.send(None);
                    ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                }
            });
            ui.add_space(6.0);
        });

        egui::CentralPanel::default().show(ctx, |ui| {
            egui::ScrollArea::vertical()
                .auto_shrink([false, false])
                .show(ui, |ui| self.content(ui));
        });
    }
}

impl TrackPasteApp {
    fn content(&mut self, ui: &mut egui::Ui) {
        let source = self.request.source_param.clone();

        {
            ui.heading(t("トラックバーのパラメータを貼り付け"));
            ui.label(tf("貼り付け先: {}", &[&self.request.target_label]));
            ui.label(tf("コピー元: {}", &[&self.request.source_label]));
            ui.separator();

            ui.horizontal(|ui| {
                ui.label(t("反映する要素:"));
                if ui.button(t("全選択")).clicked() {
                    self.set_all_checked(true);
                }
                if ui.button(t("全解除")).clicked() {
                    self.set_all_checked(false);
                }
            });
            ui.add_space(4.0);

            Self::checkbox_row(
                ui,
                self.available.values,
                &mut self.parts.values,
                "値",
                &source.values.join(", "),
            );
            Self::checkbox_row(
                ui,
                self.available.mode,
                &mut self.parts.mode,
                "移動方法",
                &source.mode.clone().unwrap_or_else(|| t("移動無し")),
            );
            Self::checkbox_row(
                ui,
                self.available.speed,
                &mut self.parts.speed,
                "加速・減速",
                &Self::speed_label(&source),
            );
            Self::checkbox_row(
                ui,
                self.available.twopoint,
                &mut self.parts.twopoint,
                "中間点無視",
                &t(if source.twopoint() { "ON" } else { "OFF" }),
            );
            let param_label = match (source.param.as_deref(), source.reference()) {
                (Some(value), true) => tf("{}（参照式）", &[value]),
                (Some(value), false) => value.to_string(),
                (None, true) => t("（参照式）"),
                (None, false) => t("(なし)"),
            };
            Self::checkbox_row(
                ui,
                self.available.param,
                &mut self.parts.param,
                "パラメータ",
                &param_label,
            );
            Self::checkbox_row(
                ui,
                self.available.timecontrol,
                &mut self.parts.timecontrol,
                "時間制御データ",
                &source.timecontrol.clone().unwrap_or_else(|| t("(なし)")),
            );

            ui.add_space(6.0);
            ui.separator();

            let merged = track::merge(
                &self.request.target_param,
                &source,
                &self.parts,
                self.request.target_value_len,
            );

            ui.label(t("適用後の値:"));
            // 外側がスクロール領域なので、ここは折り返すだけにする。
            ui.add(egui::Label::new(egui::RichText::new(&merged.value).monospace()).wrap());

            // 文言の組み立てはUI側で行い、track.rs は翻訳に依存させない。
            let adjust_message = match merged.adjust {
                track::ValueAdjust::Keep => None,
                track::ValueAdjust::Truncated { from, to } => Some(tf(
                    "中間点数が異なるため、値の数を {} → {} に切り詰めます。",
                    &[&from.to_string(), &to.to_string()],
                )),
                track::ValueAdjust::Extended { from, to } => Some(tf(
                    "中間点数が異なるため、値の数を {} → {} に補完します。",
                    &[&from.to_string(), &to.to_string()],
                )),
            };
            if let Some(message) = adjust_message {
                ui.colored_label(ui.visuals().warn_fg_color, message);
            }

            if self.request.selected_object_count > 1 {
                ui.add_space(4.0);
                ui.checkbox(
                    &mut self.apply_to_selected,
                    tf(
                        "選択中の全オブジェクト({}件)の同じ項目にも適用",
                        &[&self.request.selected_object_count.to_string()],
                    ),
                );
            }
        }
    }
}

/// 一括ペーストの対象1項目分。
#[derive(Debug, Clone)]
pub(crate) struct TrackBulkItem {
    /// 設定項目名。
    pub item: String,
    /// コピー元の値。
    pub source_param: TrackParam,
    /// 貼り付け先の現在値。
    pub target_param: TrackParam,
    /// 貼り付け先で必要な値の数（区間数 + 1）。
    pub target_value_len: usize,
}

/// 一括ペーストダイアログへ渡す情報。
#[derive(Debug, Clone)]
pub(crate) struct TrackBulkPasteRequest {
    /// 貼り付け先エフェクトの表示名。
    pub target_label: String,
    /// 対応付けできた項目。
    pub items: Vec<TrackBulkItem>,
    /// コピー元にあったが対応先が無かった項目。
    pub unmatched: Vec<String>,
    /// 選択中オブジェクトの数。
    pub selected_object_count: usize,
}

/// 一括ペーストダイアログの結果。
#[derive(Debug, Clone)]
pub(crate) struct TrackBulkPasteResponse {
    /// 反映する要素（全項目共通）。
    pub parts: PasteParts,
    /// 適用する項目名。
    pub items: Vec<String>,
    /// 選択中の全オブジェクトへ適用するか。
    pub apply_to_selected: bool,
}

struct TrackBulkPasteApp {
    window: DialogWindow,
    request: TrackBulkPasteRequest,
    available: PasteParts,
    parts: PasteParts,
    checked: Vec<bool>,
    apply_to_selected: bool,
    sender: std::sync::mpsc::Sender<Option<TrackBulkPasteResponse>>,
}

impl TrackBulkPasteApp {
    /// 全項目を通して、いずれかで反映できる要素を選択肢とする。
    fn available_parts(items: &[TrackBulkItem]) -> PasteParts {
        let mut available = PasteParts {
            values: false,
            mode: false,
            speed: false,
            twopoint: false,
            param: false,
            timecontrol: false,
        };

        for item in items {
            let source = &item.source_param;
            available.values |= !source.values.is_empty();
            available.mode = true;
            available.speed |= source.mode.is_some();
            available.twopoint |= source.mode.is_some();
            available.param |= source.has_param_info() || item.target_param.has_param_info();
            available.timecontrol |=
                source.timecontrol.is_some() || item.target_param.timecontrol.is_some();
        }

        available
    }

    fn set_all_parts(&mut self, checked: bool) {
        let available = self.available;
        self.parts = PasteParts {
            values: checked && available.values,
            mode: checked && available.mode,
            speed: checked && available.speed,
            twopoint: checked && available.twopoint,
            param: checked && available.param,
            timecontrol: checked && available.timecontrol,
        };
    }

    fn part_checkbox(ui: &mut egui::Ui, enabled: bool, checked: &mut bool, label: &str) {
        if !enabled {
            *checked = false;
        }
        ui.add_enabled(enabled, egui::Checkbox::new(checked, label));
    }

    fn response(&self) -> TrackBulkPasteResponse {
        TrackBulkPasteResponse {
            parts: self.parts,
            items: self
                .request
                .items
                .iter()
                .zip(&self.checked)
                .filter(|(_, checked)| **checked)
                .map(|(item, _)| item.item.clone())
                .collect(),
            apply_to_selected: self.apply_to_selected,
        }
    }
}

impl eframe::App for TrackBulkPasteApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        self.window.ensure_placed(ctx);

        egui::TopBottomPanel::bottom("copy_alias_bulk_actions").show(ctx, |ui| {
            ui.add_space(6.0);
            ui.horizontal(|ui| {
                let can_apply = !self.parts.is_empty() && self.checked.iter().any(|x| *x);
                if ui
                    .add_enabled(can_apply, egui::Button::new(t("貼り付け")))
                    .clicked()
                {
                    let _ = self.sender.send(Some(self.response()));
                    ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                }

                if ui.button(t("キャンセル")).clicked() {
                    let _ = self.sender.send(None);
                    ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                }
            });
            ui.add_space(6.0);
        });

        egui::CentralPanel::default().show(ctx, |ui| {
            egui::ScrollArea::vertical()
                .auto_shrink([false, false])
                .show(ui, |ui| self.content(ui));
        });
    }
}

impl TrackBulkPasteApp {
    fn content(&mut self, ui: &mut egui::Ui) {
        ui.heading(t("エフェクトのトラックバーを一括ペースト"));
        ui.label(tf("貼り付け先: {}", &[&self.request.target_label]));
        ui.separator();

        ui.horizontal(|ui| {
            ui.label(t("反映する要素:"));
            if ui.button(t("全選択")).clicked() {
                self.set_all_parts(true);
            }
            if ui.button(t("全解除")).clicked() {
                self.set_all_parts(false);
            }
        });
        ui.add_space(2.0);

        // 要素の選択は全項目で共通。
        ui.horizontal_wrapped(|ui| {
            Self::part_checkbox(ui, self.available.values, &mut self.parts.values, "値");
            Self::part_checkbox(ui, self.available.mode, &mut self.parts.mode, "移動方法");
            Self::part_checkbox(ui, self.available.speed, &mut self.parts.speed, "加速・減速");
            Self::part_checkbox(
                ui,
                self.available.twopoint,
                &mut self.parts.twopoint,
                "中間点無視",
            );
            Self::part_checkbox(ui, self.available.param, &mut self.parts.param, "パラメータ");
            Self::part_checkbox(
                ui,
                self.available.timecontrol,
                &mut self.parts.timecontrol,
                "時間制御データ",
            );
        });

        ui.add_space(6.0);
        ui.separator();

        ui.horizontal(|ui| {
            ui.label(tf("適用する項目 ({}件):", &[&self.request.items.len().to_string()]));
            if ui.button(t("全選択")).clicked() {
                self.checked.iter_mut().for_each(|x| *x = true);
            }
            if ui.button(t("全解除")).clicked() {
                self.checked.iter_mut().for_each(|x| *x = false);
            }
        });
        ui.add_space(2.0);

        let mut adjusted = 0usize;
        for (index, item) in self.request.items.iter().enumerate() {
            let merged = track::merge(
                &item.target_param,
                &item.source_param,
                &self.parts,
                item.target_value_len,
            );
            if merged.adjust.is_adjusted() {
                adjusted += 1;
            }

            if let Some(checked) = self.checked.get_mut(index) {
                ui.horizontal(|ui| {
                    ui.checkbox(checked, &item.item);
                    ui.label(egui::RichText::new(&merged.value).monospace().weak());
                });
            }
        }

        if adjusted > 0 {
            ui.add_space(4.0);
            ui.colored_label(
                ui.visuals().warn_fg_color,
                tf(
                    "中間点数が異なるため、{}件の値の数を調整します。",
                    &[&adjusted.to_string()],
                ),
            );
        }

        if !self.request.unmatched.is_empty() {
            ui.add_space(4.0);
            ui.colored_label(
                ui.visuals().weak_text_color(),
                tf(
                    "対応する項目が無いためスキップ: {}",
                    &[&self.request.unmatched.join(", ")],
                ),
            );
        }

        if self.request.selected_object_count > 1 {
            ui.add_space(6.0);
            ui.checkbox(
                &mut self.apply_to_selected,
                tf(
                    "選択中の全オブジェクト({}件)の同じエフェクトにも適用",
                    &[&self.request.selected_object_count.to_string()],
                ),
            );
        }
    }
}

pub(crate) fn show_track_bulk_paste_dialog(
    request: TrackBulkPasteRequest,
) -> AnyResult<Option<TrackBulkPasteResponse>> {
    let (tx, rx) = std::sync::mpsc::channel::<Option<TrackBulkPasteResponse>>();
    let window = DialogWindow::new([520.0, 520.0], [260.0, 260.0]);

    eframe::run_native(
        &format!("CopyAlias - {}", t("トラックバーの一括ペースト")),
        eframe::NativeOptions {
            viewport: window.viewport(),
            ..Default::default()
        },
        Box::new(move |cc| {
            setup_dialog(cc);

            let available = TrackBulkPasteApp::available_parts(&request.items);
            let checked = vec![true; request.items.len()];
            Ok(Box::new(TrackBulkPasteApp {
                window,
                request,
                available,
                parts: available,
                checked,
                apply_to_selected: false,
                sender: tx,
            }))
        }),
    )
    .map_err(|e| aviutl2::anyhow::anyhow!("一括ペーストダイアログの起動に失敗しました: {e}"))?;

    match rx.try_recv() {
        Ok(result) => Ok(result),
        Err(_) => Ok(None),
    }
}

pub(crate) fn show_track_paste_dialog(
    request: TrackPasteRequest,
) -> AnyResult<Option<TrackPasteResponse>> {
    let (tx, rx) = std::sync::mpsc::channel::<Option<TrackPasteResponse>>();
    let window = DialogWindow::new([460.0, 480.0], [230.0, 240.0]);

    eframe::run_native(
        &format!("CopyAlias - {}", t("トラックバーの貼り付け")),
        eframe::NativeOptions {
            viewport: window.viewport(),
            ..Default::default()
        },
        Box::new(move |cc| {
            setup_dialog(cc);

            let available =
                TrackPasteApp::available_parts(&request.source_param, &request.target_param);
            Ok(Box::new(TrackPasteApp {
                window,
                request,
                available,
                parts: available,
                apply_to_selected: false,
                sender: tx,
            }))
        }),
    )
    .map_err(|e| aviutl2::anyhow::anyhow!("貼り付けダイアログの起動に失敗しました: {e}"))?;

    match rx.try_recv() {
        Ok(result) => Ok(result),
        Err(_) => Ok(None),
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
    let window = DialogWindow::new([420.0, 160.0], [210.0, 80.0]);

    eframe::run_native(
        &format!("CopyAlias - {}", t("パス選択")),
        eframe::NativeOptions {
            viewport: window.viewport(),
            ..Default::default()
        },
        Box::new(move |cc| {
            setup_dialog(cc);

            Ok(Box::new(PathSelectApp {
                window,
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
    let window = DialogWindow::new([420.0, 480.0], [210.0, 240.0]);

    eframe::run_native(
        &format!("CopyAlias - {}", t("プロパティ選択")),
        eframe::NativeOptions {
            viewport: window.viewport(),
            ..Default::default()
        },
        Box::new(move |cc| {
            setup_dialog(cc);

            Ok(Box::new(ApplyDialogApp::new(window, items, tx)))
        }),
    )
    .map_err(|e| aviutl2::anyhow::anyhow!("ダイアログの起動に失敗しました: {e}"))?;

    match rx.try_recv() {
        Ok(result) => Ok(result),
        Err(_) => Ok(None),
    }
}
