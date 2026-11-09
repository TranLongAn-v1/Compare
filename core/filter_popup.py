import tkinter as tk


def lay_gia_tri_loc_theo_ngu_canh(df_nguon, column_filters, target_col):
    if df_nguon is None or df_nguon.empty:
        return []

    df = df_nguon.copy()
    for col_name, selected_set in column_filters.items():
        if col_name == target_col or selected_set is None or col_name not in df.columns:
            continue
        series = df[col_name].fillna("").astype(str)
        df = df.loc[series.isin(selected_set)]

    if target_col not in df.columns:
        return []
    return sorted({str(v) for v in df[target_col].fillna("").astype(str).tolist()})


def mo_popup_loc_cot(owner, col_name, x_root, y_root, values, column_filters, on_apply, popup_cu=None):
    if popup_cu is not None and popup_cu.winfo_exists():
        popup_cu.destroy()

    popup = tk.Toplevel(owner)
    popup.title(f"Lọc cột: {col_name}")
    popup.geometry(f"340x400+{x_root}+{y_root}")
    popup.transient(owner.winfo_toplevel())

    tk.Label(popup, text=f"Giá trị của cột '{col_name}':", anchor="w").pack(fill="x", padx=10, pady=(10, 6))

    selected_set = column_filters.get(col_name)
    selected_set = set(values) if selected_set is None else set(selected_set)

    tk.Label(popup, text="Tìm kiếm:", anchor="w").pack(fill="x", padx=10, pady=(0, 4))
    search_var = tk.StringVar()
    entry_search = tk.Entry(popup, textvariable=search_var)
    entry_search.pack(fill="x", padx=10, pady=(0, 8))

    listbox = tk.Listbox(popup, selectmode="multiple", exportselection=False)
    listbox.pack(fill="both", expand=True, padx=10, pady=(0, 8))
    filtered_values = list(values)

    def render_listbox():
        listbox.delete(0, "end")
        for val in filtered_values:
            listbox.insert("end", val)
        for i, val in enumerate(filtered_values):
            if val in selected_set:
                listbox.select_set(i)

    def on_search_change(*_args):
        keyword = search_var.get().strip().lower()
        filtered_values.clear()
        if not keyword:
            filtered_values.extend(values)
        else:
            filtered_values.extend([v for v in values if keyword in v.lower()])
        render_listbox()

    search_var.trace_add("write", on_search_change)
    render_listbox()
    entry_search.focus_set()

    btn_frame = tk.Frame(popup)
    btn_frame.pack(fill="x", padx=10, pady=(0, 10))

    def ap_dung():
        selected_idx = listbox.curselection()
        selected = {filtered_values[i] for i in selected_idx}
        keyword = search_var.get().strip().lower()

        if keyword:
            # Khi đang tìm kiếm, mặc định lọc theo toàn bộ danh sách đang hiển thị
            # nếu người dùng không chọn cụ thể mục nào.
            if not selected_idx:
                selected = set(filtered_values)
            ket_qua_set = set(selected)
        else:
            if not selected_idx:
                # Giữ nguyên trạng thái chọn nếu người dùng bấm áp dụng khi listbox vừa mất focus.
                selected = {v for v in filtered_values if v in selected_set}
            selected_set.difference_update(filtered_values)
            selected_set.update(selected)
            ket_qua_set = set(selected_set)

        ket_qua = None if len(ket_qua_set) == len(values) else ket_qua_set
        on_apply(ket_qua)
        popup.destroy()

    def chon_tat_ca():
        for val in filtered_values:
            selected_set.add(val)
        render_listbox()

    def bo_chon():
        for val in filtered_values:
            if val in selected_set:
                selected_set.remove(val)
        render_listbox()

    tk.Button(btn_frame, text="Chọn tất cả", command=chon_tat_ca).pack(side="left")
    tk.Button(btn_frame, text="Bỏ chọn", command=bo_chon).pack(side="left", padx=6)
    tk.Button(btn_frame, text="Áp dụng", command=ap_dung).pack(side="right")

    return popup
