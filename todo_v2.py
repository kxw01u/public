# todo_rev.py
from PyQt6.QtWidgets import (
    QApplication, QWidget, QTreeWidget, QTreeWidgetItem, QVBoxLayout,
    QPushButton, QHBoxLayout, QLineEdit, QFileDialog, QHeaderView, QFrame,
    QStyledItemDelegate, QColorDialog, QComboBox, QDateEdit, QProgressBar, QStyle
)
from PyQt6.QtGui import QPalette, QColor, QPainter, QTextDocument
from PyQt6.QtCore import Qt, QByteArray, QDate
import sys, os, json, random, datetime, re

ROLE_LEVEL = Qt.ItemDataRole.UserRole
ROLE_CASE = Qt.ItemDataRole.UserRole + 1

COL_PROJECT, COL_COLOR, COL_LEVEL = 0, 1, 2
COL_ACTION = 3
COL_DESC, COL_NATURE, COL_PRIORITY, COL_REF = 4, 5, 6, 7
COL_START, COL_END, COL_WEIGHT = 8, 9, 10
COL_LAST_UPDATE, COL_IDLE, COL_PIC, COL_REMARK = 11, 12, 13, 14
COL_STEP_START = 15
COL_STEP_END = COL_STEP_START + 9
COL_PROGRESS = COL_STEP_END + 1

HEADERS = [
    "Project", "C", "Level", "Action", "Desc", "Nature", "Priority", "Ref#",
    "Start", "End", "Weight", "Last Update", "Idle", "PIC", "Dep"
] + [f"Step{i}" for i in range(1, 11)] + ["Progress"]


NATURES = ["BA", "PM", "OA", "Infra", "DEV"]
PRIORITY_COLORS = {"1": "#ffb3b3", "2": "#ffd580", "3": "#b3d1ff", "4": "#b3ffb3"}
STEP_CYCLE = ["", "WIP", "Done", "N/A"]

class ProjectDelegate(QStyledItemDelegate):
    def __init__(self, tree):
        super().__init__(tree)
        self.tree = tree

    def paint(self, painter, option, index):
        project = index.data() or ""
        case_tag = index.data(ROLE_CASE) or ""
        col = index.column()
        item = self.tree.itemFromIndex(index)

        # detect Project & Priority columns by header text (robust to reordering)
        try:
            headers = [self.tree.headerItem().text(i) for i in range(self.tree.columnCount())]
            project_col = headers.index("Project")
            priority_col = headers.index("Priority")
        except ValueError:
            project_col, priority_col = COL_PROJECT, COL_PRIORITY

        # base background: use per-cell BackgroundRole if set; otherwise:
        bg = index.data(Qt.ItemDataRole.BackgroundRole)
        if bg:
            base_color = QColor(bg.color()) if hasattr(bg, "color") else QColor(bg)
        else:
            if col == project_col:
                proj_name = item.text(project_col).strip() if item else ""
                base_color = QColor(self.tree.window().color_map.get(proj_name, "#ffffff")) if proj_name else QColor("#ffffff")
            elif col == priority_col:
                try:
                    ptxt = item.text(priority_col).strip()
                    base_color = QColor(PRIORITY_COLORS.get(ptxt, "#ffffff"))
                except Exception:
                    base_color = QColor("#ffffff")
            else:
                base_color = option.palette.base().color()

        # paint background (no hover/selection fill here)
        painter.save()
        painter.fillRect(option.rect, base_color)
        painter.restore()

        # draw text
        if col == project_col:
            if not project and not case_tag:
                super().paint(painter, option, index)
                return
            is_parent = item is not None and item.childCount() > 0
            bold = "bold" if is_parent else "normal"
            html = f'<span style="color:#000;font-weight:{bold};">{project}</span>'
            if case_tag:
                html += f'<span style="color:#888;font-size:11px;font-style:italic;"> {case_tag}</span>'
            doc = QTextDocument()
            doc.setHtml(html)
            painter.save()
            painter.translate(option.rect.left() + 6, option.rect.top() + 4)
            doc.drawContents(painter)
            painter.restore()
        else:
            # default text rendering (selection/hover already made transparent via stylesheet)
            super().paint(painter, option, index)



class DragAwareTree(QTreeWidget):
    def __init__(self, recompute_hook=None, right_click_handler=None):
        super().__init__()
        self.recompute_hook = recompute_hook
        self.right_click_handler = right_click_handler

        # Disable mouse tracking so hover events never trigger
        self.setMouseTracking(False)
        # Remove focus rectangle and hover effects from style
        self.setStyleSheet("""
            QTreeView::item:hover { background: transparent; }
            QTreeView::item:selected:active { background: transparent; }
            QTreeView::item:selected:!active { background: transparent; }
            QTreeView::item { outline: none; }
        """)

    def dropEvent(self, event):
        super().dropEvent(event)
        self.recompute_levels()
        if self.recompute_hook:
            self.recompute_hook()

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.RightButton and self.right_click_handler:
            pos = e.pos()
            it = self.itemAt(pos)
            if it:
                col = self.columnAt(pos.x())
                self.right_click_handler(it, col)
                return
        super().mousePressEvent(e)
        
    def edit(self, index, trigger, event):
    # Disable editing for Step1…10 cells
        col = index.column()
        if COL_STEP_START <= col <= COL_STEP_END:
            return False
        return super().edit(index, trigger, event)

    def drawRow(self, painter, options, index):
        # Strip hover/active state flags completely before painting
        options.state &= ~QStyle.StateFlag.State_MouseOver
        options.state &= ~QStyle.StateFlag.State_Active
        super().drawRow(painter, options, index)

    def recompute_levels(self):
        def set_levels(it, lvl):
            it.setData(COL_PROJECT, ROLE_LEVEL, lvl)
            it.setText(COL_LEVEL, f"L{lvl}")
            for i in range(it.childCount()):
                set_levels(it.child(i), min(lvl + 1, 9))
        for i in range(self.topLevelItemCount()):
            set_levels(self.topLevelItem(i), 0)

class RecordKeeper(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Record Keeper")
        self.resize(1720, 880)
        self.data_file = os.path.join(os.path.dirname(__file__), "records.json")
        self.seq = 0
        self.color_map = {}
        self._updating = False
        self.column_order = None
        self.column_widths = None
        self.header_state = None

        layout = QVBoxLayout(self)
        top = QHBoxLayout(); layout.addLayout(top)
        self.address = QLineEdit(self.data_file); self.address.setReadOnly(True)
        self.btn_file = QPushButton("Change File")
        sep = QFrame(); sep.setFrameShape(QFrame.Shape.VLine); sep.setStyleSheet("background:#ccc;width:1px;")
        btn_style = "QPushButton{background:#0078d7;color:white;font-weight:bold;padding:6px 16px;border-radius:6px;border:none;}QPushButton:hover{background:#005fa3;}QPushButton:pressed{background:#004b82;}"
        self.btn_add = QPushButton("➕ Add Root"); self.btn_add.setStyleSheet(btn_style)
        self.btn_save = QPushButton("💾 Save"); self.btn_save.setStyleSheet(btn_style)
        self.btn_refresh = QPushButton("🔄 Refresh"); self.btn_refresh.setStyleSheet(btn_style)
        top.addWidget(self.address); top.addWidget(self.btn_file); top.addWidget(sep)
        top.addWidget(self.btn_add); top.addWidget(self.btn_save); top.addWidget(self.btn_refresh); top.addStretch(1)

        self.tree = DragAwareTree(recompute_hook=self.after_drop, right_click_handler=self.on_right_click_step)
        self.tree.setMouseTracking(False)  # disables hover tracking

        layout.addWidget(self.tree)
        self.tree.setHeaderLabels(HEADERS)
        header = self.tree.header()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setSectionsMovable(True)
        self.tree.setAlternatingRowColors(False)
        self.tree.setItemDelegate(ProjectDelegate(self.tree))
        self.apply_style()
        self.set_default_widths()
        self.set_row_height()

        self.btn_add.clicked.connect(self.add_root)
        self.btn_save.clicked.connect(self.save_all)
        self.btn_refresh.clicked.connect(self.refresh)
        self.btn_file.clicked.connect(self.change_file)
        self.tree.itemChanged.connect(self.on_edit)
        self.tree.itemClicked.connect(self.on_click)
        header.sectionMoved.connect(self.capture_layout)
        header.sectionResized.connect(self.capture_layout)

        self.load_all()
        self.restore_layout()
        self.rebind_all_row_widgets()
        self.apply_all_project_backgrounds()
        self.refresh_all_progress_idle()

    def set_row_height(self):
        self.tree.setStyleSheet(self.tree.styleSheet() + " QTreeWidget::item { min-height: 32px; } ")

    def apply_style(self):
        self.tree.setStyleSheet("""
            QTreeWidget {
                outline: none;
                border: 1px solid #ccc;
                gridline-color: #aaa;
                alternate-background-color: #f9f9f9;
                background: #ffffff;
            }
            QTreeView::item {
                border-bottom: 1px solid #ccc;
                border-right: 1px solid #ccc;
                padding: 4px;
            }
            QTreeView::item:selected,
            QTreeView::item:hover {
                background: rgba(0, 0, 0, 0);  /* 100% transparent */
                color: black;
            }
            QHeaderView::section {
                background: #e9e9e9;
                border: 1px solid #bbb;
                padding: 4px;
                font-weight: bold;
            }
        """)
        pal = self.tree.palette()
        pal.setColor(QPalette.ColorRole.Base, QColor("#ffffff"))
        pal.setColor(QPalette.ColorRole.AlternateBase, QColor("#f9f9f9"))
        pal.setColor(QPalette.ColorRole.Highlight, QColor(0, 0, 0, 0))  # transparent
        pal.setColor(QPalette.ColorRole.HighlightedText, QColor("#000000"))
        self.tree.setPalette(pal)



    def set_default_widths(self):
        widths = [
            240, 40, 40, 40, 220, 110, 60, 120,
            110, 110, 80, 160, 70, 120, 200
        ] + [45]*10 + [150]
        for i, w in enumerate(widths):
            if i < self.tree.columnCount():
                self.tree.setColumnWidth(i, w)

    def next_seq(self): self.seq += 1; return self.seq
    def rand_color(self): return "#{:06x}".format(random.randint(0, 0xFFFFFF))

    def make_item(self, project="", level=0):
        now = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")
        today = QDate.currentDate().toString("yyyy/MM/dd")
        row = [project, "🎨", f"L{level}", "", "", "", "", "", today, today, "", now, "0", "", ""] + [""]*10 + [""]
        it = QTreeWidgetItem(row)
        for c in range(COL_STEP_START, COL_STEP_END + 1):
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
            it.setTextAlignment(c, Qt.AlignmentFlag.AlignCenter)
        it.setData(COL_PROJECT, ROLE_LEVEL, level)
        it.setData(COL_PROJECT, ROLE_CASE, "")
        it.setFlags(it.flags() | Qt.ItemFlag.ItemIsEditable)
        self.setup_row_widgets(it)
        self.update_progress(it)
        return it

    def setup_row_widgets(self, it):
        # --- action buttons (+/-) ---
        box = QWidget()
        lay = QHBoxLayout(box)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(4)
        lay.setAlignment(Qt.AlignmentFlag.AlignCenter)
        addb = QPushButton("+")
        delb = QPushButton("-")
        for b in (addb, delb):
            b.setFlat(True)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.setStyleSheet("font-weight:bold;font-size:14px;color:black;background:transparent;border:none;")
        addb.clicked.connect(lambda: self.add_child(it))
        delb.clicked.connect(lambda: self.del_item(it))
        lay.addWidget(addb)
        lay.addWidget(delb)
        self.tree.setItemWidget(it, COL_ACTION, box)

        # --- start/end date pickers ---
        start_edit = QDateEdit()
        start_edit.setDisplayFormat("yyyy/MM/dd")
        start_edit.setCalendarPopup(True)
        end_edit = QDateEdit()
        end_edit.setDisplayFormat("yyyy/MM/dd")
        end_edit.setCalendarPopup(True)
        start_edit.setDate(QDate.currentDate())
        end_edit.setDate(QDate.currentDate())
        start_edit.dateChanged.connect(lambda _d, i=it: self.on_date_changed(i, True))
        end_edit.dateChanged.connect(lambda _d, i=it: self.on_date_changed(i, False))
        self.tree.setItemWidget(it, COL_START, start_edit)
        self.tree.setItemWidget(it, COL_END, end_edit)

        # --- full-block progress bar ---
        pb = QProgressBar()
        pb.setRange(0, 100)
        pb.setValue(0)
        pb.setTextVisible(True)
        pb.setAlignment(Qt.AlignmentFlag.AlignCenter)
        pb.setStyleSheet("""
            QProgressBar {
                border: none;
                text-align: center;
                margin: 0;
                padding: 0;
                min-height: 28px;
                max-height: 32px;
            }
            QProgressBar::chunk {
                margin: 0;
                border-radius: 3px;
                background-color: #3b82f6;
            }
        """)
        self.tree.setItemWidget(it, COL_PROGRESS, pb)


    def rebind_all_row_widgets(self):
        for it in self.iterate_items():
            if not self.tree.itemWidget(it, COL_ACTION):
                self.setup_row_widgets(it)


    def iterate_items(self, parent=None):
        arr = []
        if parent is None:
            for i in range(self.tree.topLevelItemCount()): arr.append(self.tree.topLevelItem(i))
        else:
            for i in range(parent.childCount()): arr.append(parent.child(i))
        for it in list(arr):
            yield it
            yield from self.iterate_items(it)

    def add_root(self):
        it = self.make_item("", 0)
        self.tree.addTopLevelItem(it)
        self.setup_row_widgets(it)          # ensure action buttons bound immediately
        self.apply_all_project_backgrounds()
        self.save_all()


    def add_child(self, parent):
        lvl = (parent.data(COL_PROJECT, ROLE_LEVEL) or 0) + 1
        proj = parent.text(COL_PROJECT).strip()
        it = self.make_item(proj, lvl)
        parent.addChild(it)
        parent.setExpanded(True)

        # inherit PIC
        pic_val = parent.text(COL_PIC).strip()
        if pic_val:
            it.setText(COL_PIC, pic_val)

        # assign case number
        if proj:
            if proj not in self.color_map:
                self.color_map[proj] = self.rand_color()
            case = f"{proj}_L{lvl}_{self.next_seq():02d}"
            self.safe_set(it, ROLE_CASE, case)
            self.apply_project_background(it, proj)

        self.mark_bold(parent, True)
        self.tree.recompute_levels()

        # 🔒 make parent’s steps read-only & refresh stage aggregation
        self.lock_parent_steps(parent)
        self.update_parent_stage_from_children(parent)

        self.save_all()

    def lock_parent_steps(self, parent):
        # make parent step columns non-editable and visually dimmed
        for c in range(COL_STEP_START, COL_STEP_END + 1):
            parent.setFlags(parent.flags() & ~Qt.ItemFlag.ItemIsEditable)
            parent.setBackground(c, QColor("#f0f0f0"))
    
    def update_parent_stage_from_children(self, parent):
        if parent.childCount() == 0:
            return  # no aggregation needed

        for c in range(COL_STEP_START, COL_STEP_END + 1):
            child_statuses = [parent.child(i).text(c) for i in range(parent.childCount())]
            if all(s in ("Done", "N/A") for s in child_statuses if s):
                parent.setText(c, "Done")
                self.color_step_cell(parent, c, "Done")
            elif all(s in ("", "WIP") for s in child_statuses):
                parent.setText(c, "")
                self.color_step_cell(parent, c, "")
            else:
                # mixed states -> keep WIP
                parent.setText(c, "WIP")
                self.color_step_cell(parent, c, "WIP")

        self.update_progress(parent)


    def del_item(self, it):
        par = it.parent()
        if par: par.removeChild(it); self.mark_bold(par, par.childCount()>0)
        else: self.tree.takeTopLevelItem(self.tree.indexOfTopLevelItem(it))
        self.save_all()

    def on_click(self, it, col):
        # 🎨 Color chooser
        if col == COL_COLOR:
            proj = it.text(COL_PROJECT).strip()
            if not proj:
                return
            cur = QColor(self.color_map.get(proj, "#ffffff"))
            color = QColorDialog.getColor(cur, self, "Select Color")
            if color.isValid():
                color_hex = color.name()
                self.color_map[proj] = color_hex
                self.apply_color_to_project(proj, color_hex)
                it.setBackground(COL_PROJECT, QColor(color_hex))
                self.tree.viewport().update()
                self.save_all()
                parent = it.parent()
                if parent:
                    self.update_parent_stage_from_children(parent)

        # 📋 Nature drill-down
        elif col == COL_NATURE:
            cb = QComboBox()
            cb.addItems(NATURES)
            cur = it.text(COL_NATURE).strip()
            if cur in NATURES:
                cb.setCurrentText(cur)
            cb.activated.connect(lambda _ix, i=it, w=cb: self.commit_combo(i, COL_NATURE, w))
            self.tree.setItemWidget(it, COL_NATURE, cb)
            cb.showPopup()

        # 🔺 Priority drill-down
        elif col == COL_PRIORITY:
            cb = QComboBox()
            cb.addItems(["1", "2", "3", "4"])
            cur = it.text(COL_PRIORITY).strip()
            if cur in {"1", "2", "3", "4"}:
                cb.setCurrentText(cur)
            cb.activated.connect(lambda _ix, i=it, w=cb: self.commit_priority(i, w))
            self.tree.setItemWidget(it, COL_PRIORITY, cb)
            cb.showPopup()

        # 🧮 Step1 … 10 cycle (left click)
        elif COL_STEP_START <= col <= COL_STEP_END:
            val = it.text(col)
            nxt = STEP_CYCLE[(STEP_CYCLE.index(val) + 1) % len(STEP_CYCLE)] if val in STEP_CYCLE else "WIP"
            it.setText(col, nxt)
            self.color_step_cell(it, col, nxt)
            self.touch_last_update(it)
            self.update_progress(it)
            self.save_all()

        # 📆 Weight update
        elif col == COL_WEIGHT:
            w = self.parse_weight_days(it.text(COL_WEIGHT))
            se = self.tree.itemWidget(it, COL_START)
            ee = self.tree.itemWidget(it, COL_END)
            if se and ee and w is not None:
                ee.setDate(self.add_workdays(se.date(), w))
                it.setText(COL_END, ee.date().toString("yyyy/MM/dd"))
                self.save_all()
        
        elif col == COL_PIC:
            # Build dropdown of ALL existing PIC values (all rows, all levels)
            cb = QComboBox()
            cb.setEditable(True)  # allow manual entry

            # Collect every non-empty PIC value from all items in the tree
            pics = set()
            for it2 in self.iterate_items():
                val = it2.text(COL_PIC).strip()
                if val:
                    pics.add(val)

            # Sort them nicely before adding to combo
            cb.addItems(sorted(pics, key=str.lower))

            # Pre-select current if exists
            cur = it.text(COL_PIC).strip()
            if cur and cur in pics:
                cb.setCurrentText(cur)

            # Save selection when changed
            cb.activated.connect(lambda _ix, i=it, w=cb: self.commit_combo(i, COL_PIC, w))
            self.tree.setItemWidget(it, COL_PIC, cb)
            cb.showPopup()


        elif col == COL_REMARK:  # now used for Dep column
            cb = QComboBox()
            cases = []
            for i in range(self.tree.topLevelItemCount()):
                for it in self.iterate_items(self.tree.topLevelItem(i)):
                    cnum = it.data(COL_PROJECT, ROLE_CASE)
                    if cnum:
                        cases.append(cnum)
            cb.addItems(sorted(set(cases)))
            cur = it.text(COL_REMARK).strip()
            if cur in cases:
                cb.setCurrentText(cur)
            cb.activated.connect(lambda _ix, i=it, w=cb: self.commit_combo(i, COL_REMARK, w))
            self.tree.setItemWidget(it, COL_REMARK, cb)
            cb.showPopup()


    def on_right_click_step(self, it, col):
        if not (COL_STEP_START <= col <= COL_STEP_END):
            return

        # first: cycle the clicked column
        cur = it.text(col)
        nxt = STEP_CYCLE[(STEP_CYCLE.index(cur) + 1) % len(STEP_CYCLE)] if cur in STEP_CYCLE else "WIP"
        it.setText(col, nxt)
        self.color_step_cell(it, col, nxt)

        # second: all previous steps follow the same state
        for c in range(COL_STEP_START, col):
            it.setText(c, nxt)
            self.color_step_cell(it, c, nxt)

        self.touch_last_update(it)
        self.update_progress(it)
        self.save_all()
        parent = it.parent()
        if parent:
            self.update_parent_stage_from_children(parent)


    def color_step_cell(self, it, col, status):
        if status == "Done":
            # green (unchanged)
            it.setBackground(col, QColor("#d0ffd0"))
        elif status == "WIP":
            # orange
            it.setBackground(col, QColor("#ffcc80"))
        elif status == "N/A":
            # gray
            it.setBackground(col, QColor("#d9d9d9"))
        else:
            # blank / reset
            it.setBackground(col, QColor("#ffffff"))


    def commit_combo(self, it, col, cb):
        txt = cb.currentText().strip()
        self.tree.removeItemWidget(it, col)
        it.setText(col, txt)
        self.touch_last_update(it)
        self.save_all()

    def commit_priority(self, it, cb):
        txt = cb.currentText().strip()
        self.tree.removeItemWidget(it, COL_PRIORITY)
        it.setText(COL_PRIORITY, txt)
        it.setBackground(COL_PRIORITY, QColor(PRIORITY_COLORS.get(txt, "#ffffff")))
        self.touch_last_update(it)
        self.save_all()

    def on_date_changed(self, it, is_start):
        se = self.tree.itemWidget(it, COL_START)
        ee = self.tree.itemWidget(it, COL_END)
        if not se or not ee: return
        if is_start:
            w = self.parse_weight_days(it.text(COL_WEIGHT))
            if w is not None:
                ee.setDate(self.add_workdays(se.date(), w))
        it.setText(COL_START, se.date().toString("yyyy/MM/dd"))
        it.setText(COL_END, ee.date().toString("yyyy/MM/dd"))
        self.touch_last_update(it)
        self.save_all()

    def parse_weight_days(self, text):
        if not text: return None
        m = re.match(r"\s*(\d+)\s*[dD]\s*$", text)
        return int(m.group(1)) if m else None

    def add_workdays(self, qdate, n):
        d = QDate(qdate); step = 1 if n >= 0 else -1; remaining = abs(n)
        while remaining > 0:
            d = d.addDays(step)
            if d.dayOfWeek() <= 5: remaining -= 1
        return d

    def on_edit(self, it, col):
        if self._updating: return
        if col == COL_PROJECT:
            proj = it.text(COL_PROJECT).strip()
            lvl = it.data(COL_PROJECT, ROLE_LEVEL) or 0
            if proj and proj not in self.color_map: self.color_map[proj] = self.rand_color()
            self.apply_project_background(it, proj)
            old = it.data(COL_PROJECT, ROLE_CASE)
            suffix = old.split("_")[-1] if old else f"{self.next_seq():02d}"
            case = f"{proj}_L{lvl}_{suffix}"
            self.safe_set(it, ROLE_CASE, case)
            self.tree.viewport().update()
            self.touch_last_update(it)
            self.save_all()
        elif col == COL_PRIORITY:
            txt = it.text(COL_PRIORITY).strip()
            it.setBackground(COL_PRIORITY, QColor(PRIORITY_COLORS.get(txt, "#ffffff")))
            self.touch_last_update(it); self.save_all()
        elif col == COL_WEIGHT:
            w = self.parse_weight_days(it.text(COL_WEIGHT))
            se = self.tree.itemWidget(it, COL_START); ee = self.tree.itemWidget(it, COL_END)
            if se and ee and w is not None:
                ee.setDate(self.add_workdays(se.date(), w))
                it.setText(COL_END, ee.date().toString("yyyy/MM/dd"))
            self.touch_last_update(it); self.save_all()
        elif COL_STEP_START <= col <= COL_STEP_END:
            self.touch_last_update(it); self.update_progress(it); self.save_all()
        else:
            self.touch_last_update(it); self.save_all()

    def touch_last_update(self, it):
        now = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")
        it.setText(COL_LAST_UPDATE, now)
        self.update_idle(it)

    def update_idle(self, it):
        s = it.text(COL_LAST_UPDATE).strip()
        try:
            dt = datetime.datetime.strptime(s, "%Y/%m/%d %H:%M")
            days = (datetime.datetime.now() - dt).days
            it.setText(COL_IDLE, str(days))
        except:
            it.setText(COL_IDLE, "")

    def refresh_all_progress_idle(self):
        for it in self.iterate_items():
            self.update_progress(it)
            self.update_idle(it)

    def update_progress(self, it):
        total_steps = COL_STEP_END - COL_STEP_START + 1
        done = 0
        for c in range(COL_STEP_START, COL_STEP_END + 1):
            val = it.text(c)
            if val in ("Done", "N/A"):
                done += 1
        pct = int((done / total_steps) * 100)

        pb = self.tree.itemWidget(it, COL_PROGRESS)
        if pb:
            pb.setValue(pct)
            if pct >= 100:
                pb.setStyleSheet("""
                    QProgressBar {
                        border: none;
                        text-align: center;
                        margin: 0;
                        padding: 0;
                        min-height: 28px;
                        max-height: 28px;
                    }
                    QProgressBar::chunk {
                        margin: 0;
                        border-radius: 3px;
                        background-color: #16a34a; /* green */
                    }
                """)
            else:
                pb.setStyleSheet("""
                    QProgressBar {
                        border: none;
                        text-align: center;
                        margin: 0;
                        padding: 0;
                        min-height: 28px;
                        max-height: 28px;
                    }
                    QProgressBar::chunk {
                        margin: 0;
                        border-radius: 3px;
                        background-color: #3b82f6; /* blue */
                    }
                """)


    def mark_bold(self, it, bold=True):
        f = it.font(COL_PROJECT); f.setBold(bold); it.setFont(COL_PROJECT, f)

    def after_drop(self):
        def update(it):
            proj = it.text(COL_PROJECT).strip()
            lvl = it.data(COL_PROJECT, ROLE_LEVEL) or 0
            case = it.data(COL_PROJECT, ROLE_CASE)
            if proj and case:
                suffix = case.split("_")[-1]
                self.safe_set(it, ROLE_CASE, f"{proj}_L{lvl}_{suffix}")
            for i in range(it.childCount()): update(it.child(i))
        for i in range(self.tree.topLevelItemCount()): update(self.tree.topLevelItem(i))
        self.save_all()

    def safe_set(self, it, role, val):
        self._updating = True; self.tree.blockSignals(True)
        it.setData(COL_PROJECT, role, val)
        self.tree.blockSignals(False); self._updating = False

    def apply_color_to_project(self, proj, color_hex):
        for it in self.iterate_items():
            if it.text(COL_PROJECT).strip() == proj:
                it.setBackground(COL_PROJECT, QColor(color_hex))
    def apply_project_background(self, it, proj):
        c = self.color_map.get(proj)
        if c: it.setBackground(COL_PROJECT, QColor(c))
    def apply_all_project_backgrounds(self):
        for it in self.iterate_items():
            self.apply_project_background(it, it.text(COL_PROJECT).strip())

    def capture_layout(self, *_):
        h = self.tree.header()
        self.column_order = [h.visualIndex(i) for i in range(self.tree.columnCount())]
        self.column_widths = [h.sectionSize(i) for i in range(self.tree.columnCount())]
        try:
            st = h.saveState()
            self.header_state = (bytes(st).hex()) if hasattr(st, "__bytes__") else st.data().hex()
        except: self.header_state = None

    def restore_layout(self):
        h = self.tree.header()
        if self.header_state:
            try:
                ba = QByteArray.fromHex(bytes(self.header_state, "utf-8"))
                h.restoreState(ba); return
            except: pass
        if self.column_order and len(self.column_order) == self.tree.columnCount():
            for logical, visual in reversed(list(enumerate(self.column_order))):
                cur = h.visualIndex(logical)
                if cur != visual: h.moveSection(cur, visual)
        if self.column_widths and len(self.column_widths) == self.tree.columnCount():
            for i, w in enumerate(self.column_widths): self.tree.setColumnWidth(i, w)

    def dump_tree(self, parent=None):
        items = [self.tree.topLevelItem(i) for i in range(self.tree.topLevelItemCount())] if not parent else [parent.child(i) for i in range(parent.childCount())]
        arr = []
        for it in items:
            arr.append({
                "values": [it.text(i) for i in range(self.tree.columnCount())],
                "level": it.data(COL_PROJECT, ROLE_LEVEL),
                "case": it.data(COL_PROJECT, ROLE_CASE)
            })
            kids = self.dump_tree(it)
            if kids: arr[-1]["children"] = kids
        return arr

    def restore_tree(self, data, parent=None):
        self.tree.blockSignals(True)
        try:
            for d in data:
                proj = d["values"][COL_PROJECT]; lvl = d.get("level", 0)
                it = self.make_item(proj, lvl)
                for i, v in enumerate(d["values"]):
                    if i < self.tree.columnCount(): it.setText(i, v)
                self.safe_set(it, ROLE_CASE, d.get("case", ""))
                if parent: parent.addChild(it)
                else: self.tree.addTopLevelItem(it)
                self.setup_row_widgets(it)
                se = self.tree.itemWidget(it, COL_START); ee = self.tree.itemWidget(it, COL_END)
                try:
                    sd = QDate.fromString(it.text(COL_START), "yyyy/MM/dd")
                    ed = QDate.fromString(it.text(COL_END), "yyyy/MM/dd")
                    if sd.isValid(): se.setDate(sd)
                    if ed.isValid(): ee.setDate(ed)
                except: pass
                ptxt = it.text(COL_PRIORITY).strip()
                it.setBackground(COL_PRIORITY, QColor(PRIORITY_COLORS.get(ptxt, "#ffffff")))
                for c in range(COL_STEP_START, COL_STEP_END+1):
                    self.color_step_cell(it, c, it.text(c))
                self.update_progress(it)
                if it.childCount() > 0: self.mark_bold(it, True)
                it.setExpanded(True)
                if "children" in d:
                    self.restore_tree(d["children"], it)
        finally:
            self.tree.blockSignals(False)
        self.tree.recompute_levels()

    def save_all(self):
        self.capture_layout()
        geom_hex = self.saveGeometry().data().hex()
        data = {
            "seq": self.seq,
            "colors": self.color_map,
            "records": self.dump_tree(),
            "column_order": self.column_order,
            "column_widths": self.column_widths,
            "header_state": self.header_state,
            "window_geometry": geom_hex
        }
        with open(self.data_file, "w", encoding="utf-8") as f: json.dump(data, f, indent=2, ensure_ascii=False)

    def load_all(self):
        if not os.path.exists(self.data_file): return
        try:
            with open(self.data_file, "r", encoding="utf-8") as f: d = json.load(f)
            self.seq = d.get("seq", 0)
            self.color_map = d.get("colors", {})
            self.column_order = d.get("column_order")
            self.column_widths = d.get("column_widths")
            self.header_state = d.get("header_state")
            geom = d.get("window_geometry")
            self.tree.clear()
            self.restore_tree(d.get("records", []))
            if geom:
                try:
                    ba = QByteArray.fromHex(bytes(geom, "utf-8"))
                    self.restoreGeometry(ba)
                except:
                    try:
                        self.restoreGeometry(QByteArray(bytes.fromhex(geom)))
                    except:
                        pass
        except Exception as e:
            print("load fail", e)

    def refresh(self):
        self.save_all()
        self.load_all()
        self.restore_layout()
        self.rebind_all_row_widgets()
        self.apply_all_project_backgrounds()
        self.refresh_all_progress_idle()
        self.tree.viewport().update()

    def change_file(self):
        fn, _ = QFileDialog.getSaveFileName(self, "Choose File", self.data_file, "JSON Files (*.json)")
        if fn:
            self.data_file = fn
            self.address.setText(fn)
            self.save_all()

    def closeEvent(self, e):
        self.save_all(); e.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = RecordKeeper(); w.show()
    sys.exit(app.exec())
