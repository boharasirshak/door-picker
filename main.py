import os
from sys import exit

import sqlite3
import warnings
from tkinter import messagebox
from typing import Tuple, Callable

from PIL import Image
from tksheet import Sheet
import customtkinter as ctk

warnings.filterwarnings("ignore")

WIDTH = 800
HEIGHT = 600

if not os.path.exists("data/db.db"):
    messagebox.showerror(
        "Ошибка",
        "база данных data.db не найдена в папке data/, пожалуйста, свяжитесь с администраторами!",
    )
    exit(1)


conn = sqlite3.connect("data/db.db")
cur = conn.cursor()


class InputFrame(ctk.CTkFrame):
    def __init__(
        self,
        parent,
        handle_height_input: Callable,
        handle_width_input: Callable,
        handle_color_input: Callable,
        handle_handle_type_input: Callable,
        handle_profile_system_input: Callable,
    ):
        super().__init__(parent)

        # Widgets for input frame
        self.grid_columnconfigure((1, 3, 5, 7), weight=1)  # Empty columns for spacing

        self.height_label = ctk.CTkLabel(self, text="Высота:")
        self.height_label.grid(row=0, column=2, padx=10, pady=10)
        self.height_entry = ctk.CTkEntry(self, width=120)
        self.height_entry.grid(row=0, column=4, padx=10, pady=10)
        self.height_entry.bind(
            "<KeyPress>", command=lambda x: handle_height_input(x, self.height_entry)
        )

        self.width_label = ctk.CTkLabel(self, text="Ширина:")
        self.width_label.grid(row=0, column=6, padx=10, pady=10)
        self.width_entry = ctk.CTkEntry(self, width=120)
        self.width_entry.grid(row=0, column=8, padx=10, pady=10)
        self.width_entry.bind(
            "<KeyPress>", command=lambda x: handle_width_input(x, self.width_entry)
        )

        # Dropdowns for Color, Handle Type, and Profile System
        self.color_label = ctk.CTkLabel(self, text="Цвет:")
        self.color_label.grid(row=1, column=2, padx=10, pady=10)
        self.color_dropdown = ctk.CTkComboBox(
            self,
            values=["белый", "чёрный", "серебро", "без окраса"],
            command=lambda x: handle_color_input(x, self.color_dropdown),
        )
        self.color_dropdown.grid(row=1, column=4, padx=10, pady=10, sticky="ew")

        self.handle_type_label = ctk.CTkLabel(self, text="Ручка:")
        self.handle_type_label.grid(row=1, column=6, padx=10, pady=10)
        self.handle_type_dropdown = ctk.CTkComboBox(
            self,
            values=["", "2-ст руч"],
            command=lambda x: handle_handle_type_input(x, self.handle_type_dropdown),
        )
        self.handle_type_dropdown.grid(row=1, column=8, padx=10, pady=10, sticky="ew")

        self.profile_system_label = ctk.CTkLabel(self, text="Профильная система:")
        self.profile_system_label.grid(row=1, column=10, padx=10, pady=10)
        self.profile_system_dropdown = ctk.CTkComboBox(
            self,
            values=[
                "Alumark S70",
                "Alutech W62,W72",
                "Krauss KRWD64",
                "Татпроф ТПТ 65",
            ],
            command=lambda x: handle_profile_system_input(
                x, self.profile_system_dropdown
            ),
        )
        self.profile_system_dropdown.grid(
            row=1, column=12, padx=10, pady=10, sticky="ew"
        )


class ImageFrame(ctk.CTkFrame):
    def __init__(self, master, size: Tuple[int, int] = (30, 30)):
        super().__init__(master)
        self.size = size

    def set_image(self, image: Image.Image):
        self.image = ctk.CTkImage(image, size=self.size)


class ImageRowFrame(ctk.CTkScrollableFrame):
    def __init__(
        self,
        parent,
        handle_click: Callable,
        images=[],
    ):
        super().__init__(parent)
        self.img_frames: list[ctk.CTkButton] = []
        self.update_images(images)
        self.handle_click = handle_click

    def update_images(self, image_paths: list[str]):
        for img_frame in self.img_frames:
            img_frame.destroy()
        self.img_frames.clear()

        number_of_images = len(image_paths)
        self.number_of_images = number_of_images
        images_per_row = 4
        number_of_rows = (number_of_images + images_per_row - 1) // images_per_row

        for row in range(number_of_rows):
            for col in range(images_per_row):
                self.grid_columnconfigure(col, weight=1)

        img_count = 0
        for row in range(number_of_rows):
            for col in range(images_per_row):
                if img_count >= number_of_images:
                    break
                image_path = image_paths[img_count]
                img_frame = ImageFrame(
                    self,
                    size=(WIDTH / images_per_row, HEIGHT / number_of_rows),
                )
                img_frame.set_image(Image.open("data/imgs/" + image_path))
                img_btn = ctk.CTkButton(
                    self,
                    image=img_frame.image,
                    bg_color="transparent",
                    fg_color="transparent",
                    hover_color=("gray75", "gray25"),
                    text=image_path,
                )
                img_btn._command = lambda b=img_btn: self.handle_click(b)
                img_btn.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
                self.img_frames.append(img_btn)
                img_count += 1


class ExcelFrame(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)

        # Widgets for Excel display frame (placeholder)
        # self.excel_data_label = ctk.CTkLabel(self, text="Data")
        # self.excel_data_label.grid(row=0, column=0, padx=10, pady=10)
        self.sheet = Sheet(
            self,
            header=[
                "id",
                "name",
                "per unit",
                "scheme",
                "height",
                "width",
                "color",
                "profile",
            ],
            # show_x_scrollbar=False,
            zoom=150,
            data=[],
            width=WIDTH + 165,
            theme="dark",
            default_column_width=110,
        )
        self.sheet.enable_bindings()
        self.sheet.grid(row=0, column=1, sticky="nswe", pady=10, padx=10)


class ButtonFrame(ctk.CTkFrame):
    def __init__(self, parent, save_data_callback):
        super().__init__(parent)
        self.save_data_callback = save_data_callback

        # Save Data Button
        self.save_button = ctk.CTkButton(
            self, text="Сохранить данные", command=self.on_save_data
        )
        self.save_button.grid(
            row=0, column=0, padx=10, pady=10, sticky="nsew", columnspan=2
        )

        self.grid_columnconfigure(0, weight=1)

    def on_generate_data(self):
        self.generate_data_callback()

    def on_save_data(self):
        self.save_data_callback()


class App(ctk.CTk):
    height_input: str = ""
    width_input: str = ""
    color_input: str = "белый"
    handle_type_input: str = ""
    profile_system_input: str = "Alumark S70"
    img_name: str = "*"
    all_data = []

    def __init__(self, fg_color: str | Tuple[str, str] | None = None, **kwargs):
        super().__init__(fg_color, **kwargs)

        self.title("Система определения конфигурации двери")
        self.geometry(f"{WIDTH}x{HEIGHT}")
        self.resizable(False, False)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.input_frame = InputFrame(
            self,
            handle_height_input=self.handle_height_input,
            handle_width_input=self.handle_width_input,
            handle_color_input=self.handle_color_input,
            handle_handle_type_input=self.handle_handle_type_input,
            handle_profile_system_input=self.handle_profile_system_input,
        )
        self.input_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.input_frame.grid_columnconfigure(0, weight=1)
        self.input_frame.grid_columnconfigure(3, weight=1)

        self.image_frame = ImageRowFrame(
            self, images=[], handle_click=self.handle_img_click
        )

        self.image_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

        self.image_frame.grid_columnconfigure((0, 1, 2), weight=1)

        self.excel_frame = ExcelFrame(self)
        self.excel_frame.grid(
            row=2, column=0, padx=5, pady=5, sticky="nsew", columnspan=3
        )
        self.grid_rowconfigure(2, weight=3)

        self.button_frame = ButtonFrame(self, self.save_data)
        self.button_frame.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

        self.button_frame.grid_columnconfigure((0, 1), weight=1)
        self.button_frame.save_button.grid_configure(sticky="ew")

        self._search_data()

    def handle_height_input(self, event, entry: ctk.CTkEntry):
        self._check_input(event, entry)
        self.height_input = entry.get()
        self._search_data()

    def handle_width_input(self, event, entry: ctk.CTkEntry):
        self._check_input(event, entry)
        self.width_input = entry.get()
        self._search_data()

    def handle_color_input(self, event, entry: ctk.CTkComboBox):
        self.color_input = entry.get()
        self._search_data()

    def handle_handle_type_input(self, event, entry: ctk.CTkComboBox):
        self.handle_type_input = entry.get()
        self._search_data()

    def handle_profile_system_input(self, event, entry: ctk.CTkComboBox):
        self.profile_system_input = entry.get()
        self._search_data()

    # TODO: fix the char checking!
    def _check_input(self, event, entry: ctk.CTkEntry):
        if entry.get().isnumeric():
            entry.configure(border_color="gray74", text_color="white")
            return True
        else:
            entry.configure(border_color="red", text_color="red")
            return False

    def handle_img_click(self, btn: ctk.CTkButton):
        self.img_name = btn.cget("text")

        # TODO: Fix the hover color
        for img_btn in self.image_frame.img_frames:
            if img_btn.cget("text") == self.img_name:
                img_btn.configure(bg_color="black")
            else:
                img_btn.configure(bg_color="transparent")

        self._search_data()

    def filter_data(self, sanitized_data):
        filtered_data = []

        if self.height_input == "" and self.width_input == "":
            return sanitized_data

        if self.height_input.isnumeric():
            height = int(self.height_input)
            for data in sanitized_data:
                height_field = str(data[4])
                if "-" in height_field:
                    try:
                        left_bound, right_bound = map(int, height_field.split("-"))
                        if left_bound <= height <= right_bound:
                            filtered_data.append(data)
                    except ValueError:
                        # Handle error if split does not result in two integers
                        print("Error processing range:", height_field)
                else:
                    if height_field.isnumeric() and height == int(height_field):
                        filtered_data.append(data)

        # If height input is not numeric, assume all data passes the height filter
        else:
            filtered_data = sanitized_data.copy()

        # Now filter by Width, reusing the list with height filtering applied
        final_data = []
        if self.width_input.isnumeric():
            width = int(self.width_input)
            for data in filtered_data:
                width_field = str(data[5])
                if "-" in width_field:
                    try:
                        left_bound, right_bound = map(int, width_field.split("-"))
                        if left_bound <= width <= right_bound:
                            final_data.append(data)
                    except ValueError:
                        print("Error processing range:", width_field)
                else:
                    if width_field.isnumeric() and width == int(width_field):
                        final_data.append(data)
        else:
            final_data = filtered_data.copy()

        return final_data

    def _search_data(self):
        res = cur.execute(
            """
            SELECT 
            f.id,
            f.name,
            f.per_unit, 
            p.opening_scheme,
            p.height,
            p.width,
            p.color,
            p.profile_system,
            p.image_path,
            p.id
            FROM products p
            LEFT JOIN product_features pf ON p.id = pf.product_id
            LEFT JOIN features f ON pf.feature_id = f.id
            WHERE color = ? AND handle_type = ? AND profile_system = ?
        """,
            (self.color_input, self.handle_type_input, self.profile_system_input),
        )
        all_data = res.fetchall()
        if not all_data:
            return

        all_images = [data[8] for data in all_data]
        image_data = list(set(all_images))
        sanitized_data = []

        if self.img_name != "*":
            all_data = [data for data in all_data if data[8] == self.img_name]

        self.all_data = all_data

        for data in all_data:
            data = list(data)
            height = (
                str(data[4])
                .replace("высота", "")
                .replace("мм", "")
                .replace(" ", "")
                .strip()
            )
            width = (
                str(data[5])
                .replace("ширина створки ", "")
                .replace("мм", "")
                .replace("до", "")
                .replace(" ", "")
                .strip()
            )
            data[4] = height
            data[5] = width
            sanitized_data.append(data)

        final_data = self.filter_data(sanitized_data)
        del sanitized_data

        excel_data = [data[:8] for data in final_data]

        self.image_frame.update_images(image_data)
        self.excel_frame.sheet.set_sheet_data(excel_data)
        self.all_data = final_data

    def save_data(self):
        print("Data saved/exported")
        entries = {}

        for data in self.all_data:
            entries[data[-1]] = [*data]

        print(entries)


if __name__ == "__main__":
    app = App()
    app.mainloop()
