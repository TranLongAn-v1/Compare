def pre_find_module_path(hook_api):
    # Keep default module search behavior. The bundled Python runtime on this
    # machine cannot self-discover Tcl/Tk, but tkinter itself is still present.
    return
