import tkinter as tk


class DragDropListbox(tk.Listbox):
    def __init__(self, master, on_drop_callback, **kw):
        super().__init__(master, **kw)
        self.on_drop_callback = on_drop_callback
        self.bind("<Button-1>", self.start_drag)
        self.bind("<B1-Motion>", self.on_drag)
        self.bind("<ButtonRelease-1>", self.drop)
        self.drag_start_index = None

    def start_drag(self, event):
        self.drag_start_index = self.nearest(event.y)

    def on_drag(self, event):
        if self.winfo_containing(event.x_root, event.y_root) != self:
            self.drop_outside(event)
        else:
            self.on_drag_within(event)

    def on_drag_within(self, event):
        drag_end_index = self.nearest(event.y)
        if drag_end_index < self.size() and self.drag_start_index != drag_end_index:
            dragging_item = self.get(self.drag_start_index)
            self.delete(self.drag_start_index)
            self.insert(drag_end_index, dragging_item)
            self.drag_start_index = drag_end_index

    def drop(self, event):
        # Find the widget under the current mouse position
        root_window = self.master.master
        widget_under_cursor = root_window.winfo_containing(event.x_root, event.y_root)
        if self.on_drop_callback:
            self.on_drop_callback(self.drag_start_index, event, widget_under_cursor)

    def drop_outside(self, event):
        # Assuming self.master is the frame, and self.master.master is the root window
        root_window = self.master.master
        widget_under_cursor = root_window.winfo_containing(event.x_root, event.y_root)
        if self.on_drop_callback:
            self.on_drop_callback(self.drag_start_index, widget_under_cursor)
