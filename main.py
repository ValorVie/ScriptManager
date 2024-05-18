import json
import subprocess
import logging
import tkinter as tk
from tkinter import filedialog, Menu, messagebox, font, simpledialog
from drag_drop_listbox import DragDropListbox  # Import the drag and drop module

logging.basicConfig(
    filename="app.log",
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
)


class ScriptManager:
    def __init__(self, root):
        logging.info("Application started")
        self.root = root
        self.root.title("腳本管理工具")
        self.config_file = "config.json"

        # Set application icon
        self.root.iconbitmap("batch.ico")

        # Create font using the font module
        custom_font = font.Font(family="microsoft jhenghei", size=10)

        # Globally set the style to white text on black background
        self.root.option_add(
            "*Font", custom_font
        )  # You can adjust the font and size as needed
        self.root.option_add("*Background", "black")
        self.root.option_add("*Foreground", "white")
        self.root.option_add("*Listbox*SelectBackground", "gray")
        self.root.option_add("*Listbox*SelectForeground", "white")

        # Use PanedWindow to support resizing
        self.paned_window = tk.PanedWindow(
            root, orient=tk.HORIZONTAL, sashrelief=tk.RAISED, sashwidth=6
        )
        self.paned_window.pack(fill=tk.BOTH, expand=True)

        # Initialize left and right frames
        self.left_frame = tk.Frame(self.paned_window, width=200, height=400)
        self.paned_window.add(self.left_frame, stretch="always")
        self.right_frame = tk.Frame(self.paned_window, width=400, height=400)
        self.paned_window.add(self.right_frame, stretch="always")

        # Set the minimum width of the left frame
        self.min_width = 200
        self.root.bind("<Configure>", self.on_window_resize)

        # Load configuration, including sash position
        self.load_windows()

        self.load_config()
        # Left side category list
        self.category_frame = tk.Frame(self.left_frame)
        self.category_frame.pack(fill=tk.BOTH, expand=True)
        self.category_list = DragDropListbox(
            self.category_frame, self.category_drop_event
        )
        self.category_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.category_scrollbar = tk.Scrollbar(self.category_frame, orient=tk.VERTICAL)
        self.category_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.category_list.config(yscrollcommand=self.category_scrollbar.set)
        self.category_scrollbar.config(command=self.category_list.yview)
        for category in self.categories:
            self.category_list.insert(tk.END, category)
        self.category_list.bind("<Button-3>", self.show_category_context_menu)

        if not hasattr(self, "selected_category") or not self.selected_category:
            if self.categories:
                self.selected_category = next(iter(self.categories))
                # Find the index of the selected category in the Listbox and select it
                index = list(self.categories.keys()).index(self.selected_category)
                self.category_list.select_set(
                    index
                )  # This highlights the item in the Listbox
                self.category_list.event_generate(
                    "<<ListboxSelect>>"
                )  # Trigger the select event

        # Right side script list
        self.script_frame = tk.Frame(self.right_frame)
        self.script_frame.pack(fill=tk.BOTH, expand=True)
        self.script_list = DragDropListbox(self.script_frame, self.drop_event)
        self.script_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.script_scrollbar = tk.Scrollbar(self.script_frame, orient=tk.VERTICAL)
        self.script_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.script_list.config(yscrollcommand=self.script_scrollbar.set)
        self.script_scrollbar.config(command=self.script_list.yview)
        self.script_list.bind("<Double-1>", self.execute_script)
        self.script_list.bind("<Return>", self.execute_script)
        self.script_list.bind("<Button-3>", self.show_context_menu)

        # Allow sash to be controlled independently
        self.sash_dragging = False
        self.paned_window.bind("<Button-1>", self.on_sash_drag_start)
        self.paned_window.bind("<ButtonRelease-1>", self.on_sash_drag_end)

        # Import button
        self.import_button = tk.Button(
            self.left_frame, text="導入腳本", command=self.add_scripts
        )
        self.import_button.pack(fill=tk.X)

        # Add category button
        self.add_category_button = tk.Button(
            self.left_frame, text="新增類別", command=self.add_category
        )
        self.add_category_button.pack(fill=tk.X)

        # Category right-click menu
        self.category_context_menu = Menu(self.root, tearoff=0)
        self.category_context_menu.add_command(
            label="刪除", command=self.delete_category
        )

        # Right-click menu
        self.context_menu = Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="刪除", command=self.delete_script)
        self.context_menu.add_command(
            label="在 VSCode 中打開", command=self.open_in_vscode
        )
        # Ensure all event bindings are after all widgets are created
        self.category_list.bind("<<ListboxSelect>>", self.update_script_list)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_window_resize(self, event):
        if self.paned_window.sash_coord(0)[0] < self.min_width:
            self.paned_window.sash_place(0, self.min_width, 1)

    def category_drop_event(self, index, event=None, widget_under_cursor=None):
        categories = list(self.categories.keys())
        categories = list(self.category_list.get(0, tk.END))
        self.categories = {
            category: self.categories[category] for category in categories
        }
        self.save_config()

    def show_category_context_menu(self, event):
        try:
            self.category_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.category_context_menu.grab_release()

    def delete_category(self):
        selection = self.category_list.curselection()
        if selection:  # Ensure an item is selected
            selected = selection[0]
            category_name = self.category_list.get(selected)
            if messagebox.askyesno(
                "刪除類別", f"確定要刪除類別 '{category_name}' 嗎？"
            ):
                del self.categories[category_name]
                logging.info(f"Delete Category: {category_name}")
                self.save_config()
                self.category_list.delete(
                    selected
                )  # Remove the corresponding entry from self.category_list
                if self.categories:
                    self.selected_category = next(iter(self.categories))
                    self.category_list.select_set(0)
                    self.category_list.event_generate("<<ListboxSelect>>")
                else:
                    self.selected_category = None
                    self.script_list.delete(0, tk.END)
        else:
            print("沒有選擇任何類別")  # Or you can use a dialog to prompt the user

    def drop_event(self, index, event=None, widget_under_cursor=None):
        if event and widget_under_cursor and widget_under_cursor == self.category_list:
            script_name = self.categories[self.selected_category].pop(index)
            target_category = self.category_list.get(
                self.category_list.nearest(event.y)
            )
            if target_category:
                self.categories[target_category].append(script_name)
                logging.info(f"{script_name} moved to {target_category}")
                self.save_config()
                self.update_script_list()
                messagebox.showinfo(
                    "Script Moved", f"Script moved to {target_category}"
                )
        else:
            self.categories[self.selected_category] = list(
                self.script_list.get(0, tk.END)
            )
            self.save_config()

    def load_config(self):
        try:
            with open(self.config_file, "r") as f:
                config = json.load(f)
                self.categories = config.get(
                    "categories",
                    {
                        "General": [],
                        "Restart tools": [],
                        "Admin tools": [],
                        "AI tools": [],
                    },
                )
        except FileNotFoundError:
            self.categories = {
                "General": [],
                "Restart tools": [],
                "Admin tools": [],
                "AI tools": [],
            }

    def load_windows(self):
        try:
            with open(self.config_file, "r") as f:
                config = json.load(f)
                window_size = config.get("window_size", {"width": 800, "height": 600})
                geometry_string = f"{window_size['width']}x{window_size['height']}"
                self.root.geometry(geometry_string)
                sash_position = config.get("sash_position", self.min_width)
                if sash_position < self.min_width:
                    sash_position = self.min_width
                self.root.bind(
                    "<Map>", lambda event: self.set_sash_position(sash_position)
                )
        except FileNotFoundError:
            self.root.geometry("800x600")
            self.root.bind(
                "<Map>", lambda event: self.set_sash_position(self.min_width)
            )

    def set_sash_position(self, position):
        self.paned_window.sash_place(0, position, 0)

    def save_config(self):
        window_size = {
            "width": self.root.winfo_width(),
            "height": self.root.winfo_height(),
        }
        sash_position = self.paned_window.sash_coord(0)[
            0
        ]  # Get the horizontal position of the first sash
        config = {
            "categories": self.categories,
            "window_size": window_size,
            "sash_position": sash_position,
        }
        with open(self.config_file, "w") as f:
            json.dump(config, f, indent=4)

    def update_script_list(self, event=None):
        if event:
            selection = event.widget.curselection()
        else:
            selection = self.category_list.curselection()

        if selection:
            self.selected_category = self.category_list.get(selection[0])
            self.script_list.delete(0, tk.END)
            for script in self.categories[self.selected_category]:
                self.script_list.insert(tk.END, script)

    def add_scripts(self):
        logging.info("Add Script")
        file_types = [
            ("Script files", "*.bat;*.ps1"),
        ]
        files = filedialog.askopenfilenames(filetypes=file_types)
        if files:
            for file in files:
                self.categories[self.selected_category].append(file)
                logging.info(f"Add Script: {file}")
            self.save_config()
            self.update_script_list()

    def execute_script(self, event):
        logging.info("Execute Script")
        script_path = self.script_list.get(self.script_list.curselection()[0])
        if script_path.endswith(".bat"):
            subprocess.run(f'start cmd.exe /c "{script_path}"', shell=True)
            logging.info(f"Execute Script: {script_path}")
        elif script_path.endswith(".ps1"):
            subprocess.run(
                f'start powershell.exe -NoExit -ExecutionPolicy Unrestricted -File "{script_path}"',
                shell=True,
            )
            logging.info(f"Execute Script: {script_path}")

    def delete_script(self):
        if not self.selected_category:
            print("請先選擇一個類別")
            return
        selection = self.script_list.curselection()
        if selection:  # Ensure an item is selected
            selected = selection[0]
            script_name = self.script_list.get(selected)
            if messagebox.askyesno("刪除類別", f"確定要刪除腳本 '{script_name}' 嗎？"):
                del self.categories[self.selected_category][selected]
                logging.info(f"Delete Script: {selected}")
                self.save_config()
                self.script_list.delete(
                    selected
                )  # Remove the corresponding entry from self.script_list
                self.update_script_list()
        else:
            print("沒有選擇任何腳本")  # Or you can use a dialog to prompt the user

    def add_category(self):
        category_name = simpledialog.askstring("新增類別", "請輸入新的類別名稱：")
        if category_name:
            if category_name not in self.categories:
                self.categories[category_name] = []
                self.save_config()
                self.category_list.insert(tk.END, category_name)
                logging.info(f"Add Category: {category_name}")
            else:
                messagebox.showwarning("警告", "類別名稱已存在！")

    def open_in_vscode(self):
        script_path = self.script_list.get(self.script_list.curselection()[0])
        subprocess.run(["code", script_path], shell=True)

    def show_context_menu(self, event):
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def on_sash_drag_start(self, event):
        if self.paned_window.identify(event.x, event.y) == "sash":
            self.sash_dragging = True

    def on_sash_drag_end(self, event):
        if self.sash_dragging:
            self.sash_dragging = False
            self.sash_position = self.paned_window.sash_coord(0)[0]

    def on_close(self):
        logging.info("Application ended")
        self.save_config()
        self.root.destroy()  # Ensure the window is closed


if __name__ == "__main__":
    root = tk.Tk()
    app = ScriptManager(root)
    root.mainloop()
