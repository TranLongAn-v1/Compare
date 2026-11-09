from tkcalendar import DateEntry as TkCalendarDateEntry


class DateEntry(TkCalendarDateEntry):
    """DateEntry that keeps the calendar open for clicks inside its popup."""

    def _pointer_trong_lich(self):
        x, y = self._top_cal.winfo_pointerxy()
        left = self._top_cal.winfo_rootx()
        top = self._top_cal.winfo_rooty()
        right = left + self._top_cal.winfo_width()
        bottom = top + self._top_cal.winfo_height()
        return left <= x <= right and top <= y <= bottom

    def _on_focus_out_cal(self, event):
        if self._calendar.winfo_ismapped() and self._pointer_trong_lich():
            self._calendar.focus_force()
            return
        super()._on_focus_out_cal(event)
