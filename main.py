import random
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, Any, List
from datetime import date, datetime

import pandas as pd


# Original Excel file (will be read AND overwritten with progress)
EXCEL_PATH = "Coto JLPT Grammar and Vocab Mega List.xlsx"
# Log file for daily progress
PROGRESS_LOG_PATH = "jlpt_mastery_progress.csv"

# Colors for JLPT levels
LEVEL_COLORS = {
    "N5": "#808080",   # gray
    "N4": "#008000",   # green
    "N3": "#0000FF",   # blue
    "N2": "#800080",   # purple
    "N1": "#FF4500",   # orange-red
}


class StudyData:
    """
    Simple wrapper around pandas DataFrame for each category.
    Handles filtering and record access.
    """

    def __init__(self, df: pd.DataFrame, key_column: str, level_column: str):
        """
        Initialize StudyData with its DataFrame and key/level columns.

        :param df: The underlying DataFrame.
        :param key_column: Column used as the main display key (for validation).
        :param level_column: Column storing JLPT level information.
        """
        self.df = df
        self.key_column = key_column
        self.level_column = level_column

        # Track which DataFrame indices have been changed this session.
        self.changed_rows: set[int] = set()

        # Normalize JLPT level string, e.g., "N5", "n4", " JLPT N3 ".
        if level_column in self.df.columns:
            self.df[level_column] = (
                self.df[level_column]
                .astype(str)
                .str.upper()
                .str.strip()
                .str.replace("JLPT", "")
                .str.strip()
            )

    def get_levels(self) -> List[str]:
        """
        Get available JLPT levels from the data.

        :return: Sorted list of unique JLPT levels.
        """
        if self.level_column not in self.df.columns:
            return []

        levels = sorted(
            lvl
            for lvl in self.df[self.level_column].dropna().unique()
            if lvl != ""
        )
        return levels

    def filtered_df(self, level: Optional[str] = None) -> pd.DataFrame:
        """
        Return a filtered DataFrame according to JLPT level.

        :param level: JLPT level string (e.g., 'N5'), or None / "All" for all.
        :return: Filtered DataFrame.
        """
        if level is None or level == "All":
            return self.df

        if self.level_column not in self.df.columns:
            return self.df

        return self.df[self.df[self.level_column] == level]

    def get_record_by_index(
        self, level: Optional[str], index: int
    ) -> Optional[pd.Series]:
        """
        Get one record from the filtered DataFrame by index.

        :param level: JLPT level for filtering.
        :param index: Row index in filtered data (0-based).
        :return: One row as Series or None if out of range.
        """
        fdf = self.filtered_df(level)
        if index < 0 or index >= len(fdf):
            return None

        return fdf.iloc[index]


class StudyTab(ttk.Frame):
    """
    Base class for a generic study tab (Vocabulary, Kanji, Grammar).
    """

    def __init__(
        self,
        master,
        study_data: StudyData,
        title: str,
        category_key: str,
        display_columns: Dict[str, str],
        key_candidates: Optional[List[str]] = None,
    ):
        """
        :param master: Parent widget.
        :param study_data: StudyData instance for this category.
        :param title: Tab name (e.g., 'Vocabulary').
        :param category_key: Identifier for this category ('vocab', 'kanji', 'grammar').
        :param display_columns: Mapping
               { 'LabelOnScreen': 'ColumnNameInDataFrame', ... }
        :param key_candidates: Ordered list of column names used for display.
               First non-null one is shown (e.g., ['Kanji', 'Hiragana']).
        """
        super().__init__(master)
        self.study_data = study_data
        self.title = title
        self.category_key = category_key
        self.display_columns = display_columns
        self.key_candidates = key_candidates or []

        # Map from listbox index -> position index in filtered DataFrame
        self.index_map: List[int] = []

        self.current_level: Optional[str] = None
        self.current_index: Optional[int] = None

        self._create_widgets()
        self._populate_level_combobox()
        self._refresh_listbox()

    def _create_widgets(self):
        """
        Create all widgets in the tab (filters, listbox, detail area, buttons).
        """
        # Top filter area
        filter_frame = ttk.Frame(self)
        filter_frame.pack(side=tk.TOP, fill=tk.X, padx=8, pady=8)

        ttk.Label(filter_frame, text=f"{self.title} JLPT level:").pack(
            side=tk.LEFT
        )

        self.level_var = tk.StringVar(value="All")
        self.level_combobox = ttk.Combobox(
            filter_frame,
            textvariable=self.level_var,
            state="readonly",
            width=10,
        )
        self.level_combobox.pack(side=tk.LEFT, padx=5)
        self.level_combobox.bind("<<ComboboxSelected>>", self.on_level_change)

        refresh_button = ttk.Button(
            filter_frame,
            text="Refresh",
            command=self._refresh_listbox,
        )
        refresh_button.pack(side=tk.LEFT, padx=5)

        quiz_button = ttk.Button(
            filter_frame,
            text="Random Quiz Item",
            command=self.on_random_item,
        )
        quiz_button.pack(side=tk.LEFT, padx=5)

        # Counter label (Mastered / Total for current filter)
        self.count_label = ttk.Label(filter_frame, text="Mastered: 0 / 0")
        self.count_label.pack(side=tk.RIGHT)

        # Main content area: listbox (left) + details (right)
        main_frame = ttk.Frame(self)
        main_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=8)

        # Listbox with scrollbar
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(side=tk.LEFT, fill=tk.Y)

        self.listbox = tk.Listbox(list_frame, width=30, activestyle="dotbox")
        self.listbox.pack(side=tk.LEFT, fill=tk.Y)

        scrollbar = ttk.Scrollbar(
            list_frame, orient=tk.VERTICAL, command=self.listbox.yview
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)
        self.listbox.bind("<<ListboxSelect>>", self.on_listbox_select)
        # Keyboard navigation
        self.listbox.bind("<Up>", self.on_key_up)
        self.listbox.bind("<Down>", self.on_key_down)
        self.listbox.bind("<Return>", self.on_key_enter)

        # Also bind Enter at the tab level (so it still works if focus is on buttons)
        self.bind_all("<Return>", self.on_key_enter)

        # Details text area
        detail_frame = ttk.Frame(main_frame)
        detail_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8)

        ttk.Label(detail_frame, text="Details:").pack(
            anchor=tk.W, pady=(0, 4)
        )

        self.detail_text = tk.Text(
            detail_frame,
            wrap=tk.WORD,
            height=18,
            state=tk.DISABLED,
        )
        self.detail_text.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Bottom action buttons
        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=8, pady=(0, 8))

        show_button = ttk.Button(
            bottom_frame, text="Show Details", command=self.show_selected
        )
        show_button.pack(side=tk.LEFT)

        self.known_button = ttk.Button(
            bottom_frame, text="I knew this", command=self.on_known
        )
        self.known_button.pack(side=tk.LEFT, padx=5)

        self.unknown_button = ttk.Button(
            bottom_frame, text="I didn't know", command=self.on_unknown
        )
        self.unknown_button.pack(side=tk.LEFT, padx=5)

    def _populate_level_combobox(self):
        """
        Fill the level combobox with available levels from the data.
        """
        levels = self.study_data.get_levels()
        values = ["All"] + levels
        self.level_combobox["values"] = values

        if "N5" in levels:
            self.level_var.set("N5")
        elif values:
            self.level_var.set(values[0])

    def _refresh_listbox(self, select_index: Optional[int] = None):
        """
        Refresh the listbox according to current JLPT level.
        Also update mastered count and row colors.

        :param select_index: If provided, re-select this list index after refresh.
        """
        self.listbox.delete(0, tk.END)
        self.index_map = []

        level = self.level_var.get()
        self.current_level = level if level != "" else None

        fdf = self.study_data.filtered_df(self.current_level)

        total_visible = 0
        mastered_visible = 0

        # Determine which column stores mastery state (if any)
        mastered_col = None
        for col in ["Mastered?", "Mastered"]:
            if col in fdf.columns:
                mastered_col = col
                break

        if self.key_candidates:
            has_any = any(col in fdf.columns for col in self.key_candidates)
            if not has_any:
                messagebox.showerror(
                    "Column Error",
                    f"No display column from {self.key_candidates} in "
                    f"{self.title} data.",
                )
                return

            for pos, (_, row) in enumerate(fdf.iterrows()):
                display_text = ""
                for col in self.key_candidates:
                    if col in row.index and not pd.isna(row[col]):
                        display_text = str(row[col])
                        break

                if display_text == "":
                    continue

                self.listbox.insert(tk.END, display_text)
                self.index_map.append(pos)
                total_visible += 1

                is_mastered = False
                if mastered_col is not None and mastered_col in row.index:
                    val = row[mastered_col]
                    if not pd.isna(val):
                        is_mastered = bool(val)

                idx = self.listbox.size() - 1
                if is_mastered:
                    self.listbox.itemconfig(idx, foreground="gray")
                    mastered_visible += 1
                else:
                    self.listbox.itemconfig(idx, foreground="black")
        else:
            key_col = self.study_data.key_column
            if key_col not in fdf.columns:
                messagebox.showerror(
                    "Column Error",
                    f"Missing key column '{key_col}' in {self.title} data.",
                )
                return

            for pos, (_, row) in enumerate(fdf.iterrows()):
                value = row.get(key_col, "")
                if pd.isna(value):
                    continue

                display_text = str(value)
                self.listbox.insert(tk.END, display_text)
                self.index_map.append(pos)
                total_visible += 1

                is_mastered = False
                if mastered_col is not None and mastered_col in row.index:
                    val = row[mastered_col]
                    if not pd.isna(val):
                        is_mastered = bool(val)

                idx = self.listbox.size() - 1
                if is_mastered:
                    self.listbox.itemconfig(idx, foreground="gray")
                    mastered_visible += 1
                else:
                    self.listbox.itemconfig(idx, foreground="black")

        self.count_label.config(text=f"Mastered: {mastered_visible} / {total_visible}")

        # Preserve selection if requested
        if select_index is not None and self.listbox.size() > 0:
            select_index = max(0, min(select_index, self.listbox.size() - 1))
            self._select_list_index(select_index)
        else:
            self._set_detail_text("")

    def on_level_change(self, _event=None):
        """
        Callback when user selects a JLPT level.
        """
        self._refresh_listbox()

    def on_listbox_select(self, _event=None):
        """
        Handle listbox selection change.
        """
        if not self.listbox.curselection():
            return

        index = self.listbox.curselection()[0]
        self.current_index = index
        self.show_selected()

    def _get_current_df_index(self) -> Optional[int]:
        """
        Translate current listbox index to DataFrame position index.
        """
        if self.current_index is None:
            if self.listbox.curselection():
                self.current_index = self.listbox.curselection()[0]
            else:
                return None

        if self.current_index < 0 or self.current_index >= len(self.index_map):
            return None

        return self.index_map[self.current_index]

    def show_selected(self):
        """
        Show details of the currently selected item in the detail text area.
        """
        df_pos = self._get_current_df_index()
        if df_pos is None:
            return

        record = self.study_data.get_record_by_index(
            self.current_level, df_pos
        )
        if record is None:
            return

        text_lines = []

        # Build detail text using the configured display_columns mapping
        for label, column_name in self.display_columns.items():
            if column_name in record.index:
                value = record[column_name]
                if pd.isna(value):
                    continue
                text_lines.append(f"{label}: {value}")

        content = "\n".join(text_lines)
        self._set_detail_text(content)

    def _set_detail_text(self, content: str):
        """
        Helper function to update detail Text widget safely.

        :param content: Text content to display.
        """
        self.detail_text.config(state=tk.NORMAL)
        self.detail_text.delete("1.0", tk.END)
        self.detail_text.insert(tk.END, content)
        self.detail_text.config(state=tk.DISABLED)

    def on_random_item(self):
        """
        Select and show a random item from the currently displayed list.
        """
        if not self.index_map:
            messagebox.showinfo(
                "No Data", f"No {self.title.lower()} available."
            )
            return

        rand_list_index = random.randrange(len(self.index_map))
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(rand_list_index)
        self.listbox.see(rand_list_index)
        self.current_index = rand_list_index
        self.show_selected()

    def _set_mastered_state(self, state: bool):
        """
        Update the mastered status for the current item.

        If the dataset has 'Mastered?' or 'Mastered', that column is used.
        Otherwise a new 'Mastered' column is created (for example in vocab).
        Also:
          - Track changed rows for faster save.
          - Log new mastery events (per category + level) to the app
            for daily progress.
          - Update listbox color for this row and counters.
        """
        df_pos = self._get_current_df_index()
        if df_pos is None:
            messagebox.showinfo("Info", "Please select an item first.")
            return

        record = self.study_data.get_record_by_index(
            self.current_level,
            df_pos,
        )
        if record is None:
            return

        df = self.study_data.df
        row_idx = record.name

        # Determine the mastered column
        mastered_col = None
        for col in ["Mastered?", "Mastered"]:
            if col in df.columns:
                mastered_col = col
                break

        # Create new mastered column if needed (for vocab)
        if mastered_col is None:
            mastered_col = "Mastered"
            df[mastered_col] = False  # default for all rows

        # Old state (before update)
        old_val = df.at[row_idx, mastered_col]
        old_bool = False
        if not pd.isna(old_val):
            old_bool = bool(old_val)

        # Update TRUE/FALSE
        df.at[row_idx, mastered_col] = bool(state)

        # Mark this row as changed for incremental save
        self.study_data.changed_rows.add(row_idx)

        # If this is a NEW mastery (False -> True), log to daily progress
        is_new_mastery = state and not old_bool
        if is_new_mastery:
            level = None
            if "JLPT" in record.index:
                level = str(record["JLPT"]).strip().upper() or None
            app = self.winfo_toplevel()
            if hasattr(app, "register_mastery_event"):
                app.register_mastery_event(
                    category=self.category_key, level=level
                )

        # Update list row color
        if self.current_index is not None:
            color = "gray" if state else "black"
            try:
                self.listbox.itemconfig(self.current_index, foreground=color)
            except tk.TclError:
                pass

        # Refresh while preserving the current list index (so we can advance reliably)
        keep_index = self.current_index
        self._refresh_listbox(select_index=keep_index)

    def on_known(self):
        """
        Mark the current item as mastered (True), then jump to the next item
        relative to what was just mastered.
        """
        # Capture current index BEFORE update/refresh
        cur = 0
        if self.listbox.curselection():
            cur = self.listbox.curselection()[0]
        elif self.current_index is not None:
            cur = self.current_index

        self.current_index = cur
        self._set_mastered_state(True)

        # Jump to the next item after the one we just mastered
        next_idx = min(cur + 1, max(0, self.listbox.size() - 1))
        self._select_list_index(next_idx)

        messagebox.showinfo("Great!", "Marked as mastered.")

    def on_unknown(self):
        """
        Mark the current item as NOT mastered (False).
        """
        self._set_mastered_state(False)
        messagebox.showinfo("Review", "Marked as NOT mastered.")

    def _select_list_index(self, list_index: int):
        """
        Select the given listbox index, scroll into view, and display details.
        """
        if list_index < 0 or list_index >= self.listbox.size():
            return

        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(list_index)
        self.listbox.activate(list_index)
        self.listbox.see(list_index)

        self.current_index = list_index
        self.show_selected()

    def select_next_item(self):
        """
        Move selection to next list item (if any), show details.
        """
        if self.listbox.size() == 0:
            return

        cur = 0
        if self.listbox.curselection():
            cur = self.listbox.curselection()[0]

        nxt = min(cur + 1, self.listbox.size() - 1)
        self._select_list_index(nxt)

    def select_prev_item(self):
        """
        Move selection to previous list item (if any), show details.
        """
        if self.listbox.size() == 0:
            return

        cur = 0
        if self.listbox.curselection():
            cur = self.listbox.curselection()[0]

        prv = max(cur - 1, 0)
        self._select_list_index(prv)

    def on_key_up(self, _event=None):
        """
        Up arrow: select previous item.
        """
        self.select_prev_item()
        return "break"

    def on_key_down(self, _event=None):
        """
        Down arrow: select next item.
        """
        self.select_next_item()
        return "break"

    def on_key_enter(self, _event=None):
        """
        Enter: mark as mastered (I knew it) and auto-advance.
        """
        self.on_known()
        return "break"


class CategoryProgressView(ttk.Frame):
    """
    One category's progress: table + stacked bar chart by JLPT level.
    """

    def __init__(self, master, app: "StudyApp", category_key: str):
        super().__init__(master)
        self.app = app
        self.category_key = category_key
        self._create_widgets()
        self.refresh()

    def _create_widgets(self):
        top_frame = ttk.Frame(self)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=8, pady=8)

        name = self.category_key.capitalize()
        ttk.Label(
            top_frame,
            text=f"{name} - Daily Mastery by JLPT Level",
        ).pack(side=tk.LEFT)

        refresh_button = ttk.Button(
            top_frame,
            text="Refresh",
            command=self.refresh,
        )
        refresh_button.pack(side=tk.RIGHT)

        # Table
        table_frame = ttk.Frame(self)
        table_frame.pack(side=tk.TOP, fill=tk.X, expand=False, padx=8)

        columns = ("date", "N5", "N4", "N3", "N2", "N1", "total")
        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            height=7,
        )
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80, anchor=tk.CENTER)
        self.tree.column("date", width=110, anchor=tk.CENTER)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(
            table_frame, orient=tk.VERTICAL, command=self.tree.yview
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Bar chart canvas
        chart_frame = ttk.Frame(self)
        chart_frame.pack(
            side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=8
        )

        self.chart_canvas = tk.Canvas(chart_frame, width=650, height=260)
        self.chart_canvas.pack(fill=tk.BOTH, expand=True)

    def refresh(self):
        """
        Refresh table and bar chart from app.daily_progress for this category.
        """
        # Build per-date per-level counts for this category
        data = self.app.daily_progress  # nested dict

        # data structure: {date_str: {category: {level: count}}}
        date_keys = sorted(data.keys())

        rows = []
        for d_str in date_keys:
            cat_dict = data[d_str].get(self.category_key, {})
            counts_by_level = {
                lv: cat_dict.get(lv, 0) for lv in ["N5", "N4", "N3", "N2", "N1"]
            }
            total = sum(counts_by_level.values())
            if total == 0:
                continue
            rows.append((d_str, counts_by_level, total))

        # Update table
        for item in self.tree.get_children():
            self.tree.delete(item)

        for d_str, counts_by_level, total in rows:
            try:
                dt = datetime.strptime(d_str, "%Y-%m-%d")
                label = dt.strftime("%m/%d")
            except ValueError:
                label = d_str

            self.tree.insert(
                "",
                tk.END,
                values=(
                    label,
                    counts_by_level["N5"],
                    counts_by_level["N4"],
                    counts_by_level["N3"],
                    counts_by_level["N2"],
                    counts_by_level["N1"],
                    total,
                ),
            )

        # Draw chart
        self._draw_chart(rows)

    def _draw_chart(self, rows: List[tuple]):
        """
        Draw a stacked bar chart given rows:
          (date_str, {N5..N1: count}, total)
        """
        self.chart_canvas.delete("all")

        if not rows:
            self.chart_canvas.create_text(
                330,
                130,
                text="No data yet for this category.",
            )
            return

        # Canvas dimensions
        width = int(self.chart_canvas["width"])
        height = int(self.chart_canvas["height"])
        padding_left = 50
        padding_bottom = 50
        padding_top = 20
        padding_right = 20

        chart_width = width - padding_left - padding_right
        chart_height = height - padding_top - padding_bottom

        totals = [total for _, _, total in rows]
        max_total = max(totals) if totals else 1
        n = len(rows)
        if n == 0:
            return

        space = chart_width / n
        bar_width = space * 0.6

        # Axes
        x0 = padding_left
        y0 = padding_top
        x_axis_y = padding_top + chart_height

        self.chart_canvas.create_line(x0, y0, x0, x_axis_y)
        self.chart_canvas.create_line(x0, x_axis_y, x0 + chart_width, x_axis_y)

        levels = ["N5", "N4", "N3", "N2", "N1"]

        for i, (d_str, counts_by_level, total) in enumerate(rows):
            x_center = padding_left + space * i + space / 2
            cum_height = 0.0

            for lv in levels:
                c = counts_by_level[lv]
                if c <= 0:
                    continue
                if max_total == 0:
                    seg_h = 0
                else:
                    seg_h = chart_height * (c / max_total)

                y1 = x_axis_y - cum_height - seg_h
                y2 = x_axis_y - cum_height
                x1 = x_center - bar_width / 2
                x2 = x_center + bar_width / 2

                color = LEVEL_COLORS.get(lv, "gray")
                self.chart_canvas.create_rectangle(
                    x1, y1, x2, y2, fill=color, outline="black"
                )

                cum_height += seg_h

            # Date label on X-axis
            try:
                dt = datetime.strptime(d_str, "%Y-%m-%d")
                label = dt.strftime("%m/%d")
            except ValueError:
                label = d_str

            self.chart_canvas.create_text(
                x_center,
                x_axis_y + 12,
                text=label,
                anchor=tk.N,
                font=("TkDefaultFont", 8),
            )

        # Legend
        legend_x = padding_left + 10
        legend_y = padding_top + 10
        for lv in levels:
            color = LEVEL_COLORS.get(lv, "gray")
            self.chart_canvas.create_rectangle(
                legend_x,
                legend_y,
                legend_x + 15,
                legend_y + 15,
                fill=color,
                outline="black",
            )
            self.chart_canvas.create_text(
                legend_x + 22,
                legend_y + 7,
                text=lv,
                anchor=tk.W,
                font=("TkDefaultFont", 8),
            )
            legend_y += 20


class ProgressTab(ttk.Frame):
    """
    Progress tab: three inner tabs (Vocab, Kanji, Grammar),
    each shows its own table and stacked bar chart.
    """

    def __init__(self, master, app: "StudyApp"):
        super().__init__(master)
        self.app = app
        self._create_widgets()

    def _create_widgets(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True)

        self.vocab_view = CategoryProgressView(self, self.app, "vocab")
        self.kanji_view = CategoryProgressView(self, self.app, "kanji")
        self.grammar_view = CategoryProgressView(self, self.app, "grammar")

        notebook.add(self.vocab_view, text="Vocabulary")
        notebook.add(self.kanji_view, text="Kanji")
        notebook.add(self.grammar_view, text="Grammar")

    def refresh(self):
        """
        Refresh all category views.
        """
        self.vocab_view.refresh()
        self.kanji_view.refresh()
        self.grammar_view.refresh()


class StudyApp(tk.Tk):
    """
    Main application window for JLPT vocabulary / kanji / grammar study.
    """

    def __init__(self):
        super().__init__()

        self.title("JLPT Study Tool - Vocabulary / Kanji / Grammar")
        self.geometry("1100x700")

        self._setup_style()

        # Load all data AND keep original sheets in memory
        try:
            self.data_sets, self.original_sheets = self._load_data()
        except Exception as exc:
            messagebox.showerror("Load Error", str(exc))
            self.destroy()
            return

        # Load historical daily progress log (if exists)
        # Structure: {date_str: {category: {level: count}}}
        self.daily_progress: Dict[str, Dict[str, Dict[str, int]]] = (
            self._load_progress_log()
        )

        # Create notebook (tabs)
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True)

        vocab_tab = StudyTab(
            notebook,
            self.data_sets["vocab"],
            title="Vocabulary",
            category_key="vocab",
            display_columns={
                "Kanji": "Kanji",
                "Hiragana": "Hiragana",
                "English": "English",
                "Percent Completed": "Percent Completed",
                "Vocabulary Mastered": "Vocabulary Mastered",
                "Mastered?": "Mastered?",
                "Mastered": "Mastered",
                "JLPT": "JLPT",
            },
            key_candidates=["Kanji", "Hiragana"],
        )

        kanji_tab = StudyTab(
            notebook,
            self.data_sets["kanji"],
            title="Kanji",
            category_key="kanji",
            display_columns={
                "Kanji": "Kanji",
                "Pronounciation": "Pronounciation",
                "English": "English",
                "Mastered?": "Mastered?",
                "Percent Completed": "Percent Completed",
                "Kanji Mastered": "KanjiMastered",
                "JLPT": "JLPT",
            },
            key_candidates=["Kanji"],
        )

        grammar_tab = StudyTab(
            notebook,
            self.data_sets["grammar"],
            title="Grammar",
            category_key="grammar",
            display_columns={
                "Grammar Point": "Grammar Point",
                "Mastered?": "Mastered?",
                "Percent Completed": "Percent Completed",
                "Grammar Points Mastered": "Grammar Points Mastered",
                "JLPT": "JLPT",
            },
            key_candidates=["Grammar Point"],
        )

        self.progress_tab = ProgressTab(notebook, self)

        notebook.add(vocab_tab, text="Vocabulary")
        notebook.add(kanji_tab, text="Kanji")
        notebook.add(grammar_tab, text="Grammar")
        notebook.add(self.progress_tab, text="Progress")

        # Save on close (to original Excel + progress log)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _setup_style(self):
        """
        Set up a basic ttk theme / style.
        """
        style = ttk.Style()
        style.theme_use(style.theme_use())

    @staticmethod
    def _load_data() -> (Dict[str, StudyData], Dict[str, pd.DataFrame]):
        """
        Load excel data for vocabulary, kanji and grammar.

        It assumes sheet names like:
          'N5 Grammar', 'N5 Vocab', 'N5 Kanji',
          'N4 Grammar', 'N4 Vocab', 'N4 Kanji', ...
        The part before the first space is treated as JLPT level.

        It also stores two internal columns:
          - '_SheetName' : original sheet name
          - '_RowIndexInSheet' : original row index inside that sheet

        Returns:
          data_sets: dict with 'vocab', 'kanji', 'grammar' => StudyData
          original_sheets: dict of sheet_name -> original DataFrame
        """
        try:
            all_sheets = pd.read_excel(EXCEL_PATH, sheet_name=None)
        except Exception as exc:
            raise RuntimeError(
                f"Failed to load Excel file '{EXCEL_PATH}': {exc}"
            ) from exc

        vocab_frames: List[pd.DataFrame] = []
        kanji_frames: List[pd.DataFrame] = []
        grammar_frames: List[pd.DataFrame] = []

        for sheet_name, df in all_sheets.items():
            if not isinstance(df, pd.DataFrame):
                continue

            parts = str(sheet_name).split()
            if not parts:
                continue

            level = parts[0].upper()

            df = df.copy()
            df["JLPT"] = level
            df["_SheetName"] = sheet_name
            df["_RowIndexInSheet"] = df.index

            name_upper = sheet_name.upper()
            if "GRAMMAR" in name_upper:
                grammar_frames.append(df)
            elif "VOCAB" in name_upper:
                vocab_frames.append(df)
            elif "KANJI" in name_upper:
                kanji_frames.append(df)

        if vocab_frames:
            vocab_df = pd.concat(vocab_frames, ignore_index=True)
        else:
            vocab_df = pd.DataFrame()

        if kanji_frames:
            kanji_df = pd.concat(kanji_frames, ignore_index=True)
        else:
            kanji_df = pd.DataFrame()

        if grammar_frames:
            grammar_df = pd.concat(grammar_frames, ignore_index=True)
        else:
            grammar_df = pd.DataFrame()

        data_sets: Dict[str, Any] = {
            "vocab": StudyData(
                vocab_df,
                key_column="Kanji",
                level_column="JLPT",
            ),
            "kanji": StudyData(
                kanji_df,
                key_column="Kanji",
                level_column="JLPT",
            ),
            "grammar": StudyData(
                grammar_df,
                key_column="Grammar Point",
                level_column="JLPT",
            ),
        }

        return data_sets, all_sheets

    def _load_progress_log(self) -> Dict[str, Dict[str, Dict[str, int]]]:
        """
        Load historical daily mastery counts from CSV, if present.

        Returns nested dict:
          {date_str: {category: {level: count}}}
        """
        try:
            df = pd.read_csv(PROGRESS_LOG_PATH)
        except FileNotFoundError:
            return {}
        except Exception:
            return {}

        required_cols = {"date", "category", "level", "count"}
        if not required_cols.issubset(df.columns):
            return {}

        result: Dict[str, Dict[str, Dict[str, int]]] = {}
        for _, row in df.iterrows():
            d_str = str(row["date"])
            cat = str(row["category"])
            lv = str(row["level"])
            cnt = int(row["count"])

            if d_str not in result:
                result[d_str] = {}
            if cat not in result[d_str]:
                result[d_str][cat] = {}
            result[d_str][cat][lv] = result[d_str][cat].get(lv, 0) + cnt

        return result

    def _save_progress_log(self):
        """
        Save self.daily_progress into CSV.
        """
        records = []
        for d_str, cat_dict in self.daily_progress.items():
            for cat, lvl_dict in cat_dict.items():
                for lv, cnt in lvl_dict.items():
                    records.append(
                        {
                            "date": d_str,
                            "category": cat,
                            "level": lv,
                            "count": cnt,
                        }
                    )
        if not records:
            return
        df = pd.DataFrame(records)
        df.to_csv(PROGRESS_LOG_PATH, index=False)

    def register_mastery_event(self, category: str, level: Optional[str]):
        """
        Called when a new mastery (False -> True) occurs.
        Increase today's count for given category and level.
        Level can be None or unknown; we still record it as 'Unknown'.
        """
        today_str = date.today().isoformat()
        lvl = level if level in LEVEL_COLORS or (level and level.startswith("N")) else "Unknown"

        if today_str not in self.daily_progress:
            self.daily_progress[today_str] = {}
        if category not in self.daily_progress[today_str]:
            self.daily_progress[today_str][category] = {}
        self.daily_progress[today_str][category][lvl] = (
            self.daily_progress[today_str][category].get(lvl, 0) + 1
        )

        if hasattr(self, "progress_tab"):
            self.progress_tab.refresh()

    def _save_progress(self):
        """
        Save current progress back into the original Excel file.

        Only rows whose mastered state changed this session are applied.
        Other rows and unrelated sheets are left untouched.
        """
        vocab_data = self.data_sets["vocab"]
        kanji_data = self.data_sets["kanji"]
        grammar_data = self.data_sets["grammar"]

        # If nothing changed in any category, skip saving entirely
        if (
            not vocab_data.changed_rows
            and not kanji_data.changed_rows
            and not grammar_data.changed_rows
        ):
            return

        vocab_df_all = vocab_data.df
        kanji_df_all = kanji_data.df
        grammar_df_all = grammar_data.df

        # Build "changed only" views
        vocab_changed = (
            vocab_df_all.loc[list(vocab_data.changed_rows)]
            if vocab_data.changed_rows
            else pd.DataFrame()
        )
        kanji_changed = (
            kanji_df_all.loc[list(kanji_data.changed_rows)]
            if kanji_data.changed_rows
            else pd.DataFrame()
        )
        grammar_changed = (
            grammar_df_all.loc[list(grammar_data.changed_rows)]
            if grammar_data.changed_rows
            else pd.DataFrame()
        )

        # Work on a copy of original sheets
        all_sheets = {
            name: df.copy() for name, df in self.original_sheets.items()
        }

        for sheet_name, orig_df in all_sheets.items():
            if not isinstance(orig_df, pd.DataFrame):
                continue

            name_upper = sheet_name.upper()

            if "GRAMMAR" in name_upper:
                category_df = grammar_changed
            elif "VOCAB" in name_upper:
                category_df = vocab_changed
            elif "KANJI" in name_upper:
                category_df = kanji_changed
            else:
                # Unrelated sheet, leave as is
                all_sheets[sheet_name] = orig_df
                continue

            if category_df.empty:
                # No changes for this category in this sheet
                all_sheets[sheet_name] = orig_df
                continue

            # Rows from this category that belong to this sheet
            sheet_rows = category_df[
                category_df["_SheetName"] == sheet_name
            ]

            if sheet_rows.empty:
                all_sheets[sheet_name] = orig_df
                continue

            updated_df = orig_df.copy()

            # Ensure any mastered-columns existing in our data also exist
            # in the original sheet.
            for col in ["Mastered?", "Mastered"]:
                if col in sheet_rows.columns and col not in updated_df.columns:
                    updated_df[col] = pd.NA

            mastered_cols = [
                col
                for col in ["Mastered?", "Mastered"]
                if col in sheet_rows.columns and col in updated_df.columns
            ]

            # Apply updates according to _RowIndexInSheet
            for _, row in sheet_rows.iterrows():
                row_idx = row["_RowIndexInSheet"]
                if row_idx not in updated_df.index:
                    continue
                for col in mastered_cols:
                    updated_df.at[row_idx, col] = row[col]

            all_sheets[sheet_name] = updated_df

        # Write everything back to the original file
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="w") as writer:
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Clear change tracking after successful save
        vocab_data.changed_rows.clear()
        kanji_data.changed_rows.clear()
        grammar_data.changed_rows.clear()

    def on_close(self):
        """
        Handle window closing: save progress back to original file
        and save daily progress log, then close.
        """
        try:
            self._save_progress()
            self._save_progress_log()
            messagebox.showinfo(
                "Progress Saved",
                f"Your progress has been saved to '{EXCEL_PATH}' "
                f"and '{PROGRESS_LOG_PATH}'.",
            )
        except Exception as exc:
            answer = messagebox.askyesno(
                "Save Error",
                f"Failed to save progress:\n{exc}\n\n"
                f"Do you want to exit without saving?",
            )
            if not answer:
                return

        self.destroy()


def main():
    """
    Entry point for the application.
    """
    app = StudyApp()
    app.mainloop()


if __name__ == "__main__":
    main()
