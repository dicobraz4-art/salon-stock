import sqlite3
import os
from kivymd.app import MDApp
from kivymd.uix.screen import MDScreen
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.toolbar import MDTopAppBar
from kivymd.uix.textfield import MDTextField
from kivymd.uix.list import MDList, TwoLineAvatarIconListItem, IconLeftWidget
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.button import MDRaisedButton, MDIconButton
from kivymd.uix.dialog import MDDialog
from kivymd.uix.label import MDLabel
from openpyxl import Workbook, load_workbook

# --- 1. РАБОТА С БАЗОЙ ДАННЫХ (SQLite) ---
class DBManager:
    def __init__(self, db_name="salon_stock.db"):
        self.db_name = db_name
        self.init_db()

    def get_connection(self):
        return sqlite3.connect(self.db_name)

    def init_db(self):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS products (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    brand TEXT NOT NULL,
                    shade TEXT NOT NULL,
                    barcode TEXT,
                    nominal_g REAL DEFAULT 60.0,
                    full_count INTEGER DEFAULT 0,
                    opened_weight_g REAL DEFAULT 0.0,
                    is_archived INTEGER DEFAULT 0,
                    notes TEXT
                )
            """)
            conn.commit()

    def get_products(self, query="", show_archived=False):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            sql = """
                SELECT id, brand, shade, nominal_g, full_count, opened_weight_g, is_archived, notes 
                FROM products 
                WHERE (is_archived = ?) AND (brand LIKE ? OR shade LIKE ? OR barcode = ?)
                ORDER BY brand, shade
            """
            q_like = f"%{query}%"
            cursor.execute(sql, (1 if show_archived else 0, q_like, q_like, query))
            return cursor.fetchall()

    def get_product(self, prod_id):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM products WHERE id = ?", (prod_id,))
            return cursor.fetchone()

    def add_or_update(self, brand, shade, barcode, nominal_g, full_cnt, opened_g, notes, prod_id=None):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            if prod_id:
                cursor.execute("""
                    UPDATE products 
                    SET brand=?, shade=?, barcode=?, nominal_g=?, full_count=?, opened_weight_g=?, notes=? 
                    WHERE id=?
                """, (brand, shade, barcode, nominal_g, full_cnt, opened_g, notes, prod_id))
            else:
                cursor.execute("""
                    INSERT INTO products (brand, shade, barcode, nominal_g, full_count, opened_weight_g, notes)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (brand, shade, barcode, nominal_g, full_cnt, opened_g, notes))
            conn.commit()

    def write_off_full(self, prod_id):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE products SET full_count = MAX(0, full_count - 1) WHERE id = ?", (prod_id,))
            conn.commit()

    def write_off_weight(self, prod_id, grams):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT full_count, opened_weight_g, nominal_g FROM products WHERE id = ?", (prod_id,))
            row = cursor.fetchone()
            if not row:
                return
            full_cnt, opened_g, nom = row
            total_g = (full_cnt * nom) + opened_g
            new_total = max(0.0, total_g - grams)

            new_full = int(new_total // nom)
            new_opened = round(new_total % nom, 1)

            cursor.execute("UPDATE products SET full_count = ?, opened_weight_g = ? WHERE id = ?", 
                           (new_full, new_opened, prod_id))
            conn.commit()

    def toggle_archive(self, prod_id):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE products SET is_archived = CASE WHEN is_archived = 1 THEN 0 ELSE 1 END WHERE id = ?", (prod_id,))
            conn.commit()

    def export_excel(self, filepath="salon_export.xlsx"):
        wb = Workbook()
        ws = wb.active
        ws.title = "Склад"
        ws.append(["ID", "Бренд", "Номер", "Штрихкод", "Номинал (г)", "Полных (шт)", "Начатых (г)", "Всего (г)", "Архив", "Примечания"])
        
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT id, brand, shade, barcode, nominal_g, full_count, opened_weight_g, is_archived, notes FROM products")
            for r in cursor.fetchall():
                total = (r[5] * r[4]) + r[6]
                ws.append([r[0], r[1], r[2], r[3], r[4], r[5], r[6], total, "Да" if r[7] else "Нет", r[8]])
        wb.save(filepath)

    def import_excel(self, filepath="salon_import.xlsx"):
        if not os.path.exists(filepath):
            return
        wb = load_workbook(filepath)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[1] or not row[2]:
                continue
            self.add_or_update(
                brand=str(row[1]).strip(),
                shade=str(row[2]).strip(),
                barcode=str(row[3]).strip() if row[3] else None,
                nominal_g=float(row[4] or 60.0),
                full_cnt=int(row[5] or 0),
                opened_g=float(row[6] or 0.0),
                notes=str(row[9] or "")
            )

# --- 2. ИНТЕРФЕЙС ПРИЛОЖЕНИЯ (KivyMD) ---
class SalonStockApp(MDApp):
    def build(self):
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "Pink"
        self.db = DBManager()
        self.dialog = None

        # Демо-данные при первом запуске
        if not self.db.get_products():
            self.db.add_or_update("L'Oreal", "7.11", "3474630501234", 60.0, 3, 45.0, "Холодный блонд")
            self.db.add_or_update("Wella", "9/16", "4015600123456", 60.0, 1, 15.0, "Колестон перфект")

        self.root_screen = MDScreen()
        layout = MDBoxLayout(orientation="vertical")

        # Верхняя панель
        self.toolbar = MDTopAppBar(
            title="Склад Красителей",
            elevation=4,
            right_action_items=[
                ["file-excel-outline", lambda x: self.export_to_excel()],
                ["plus-circle-outline", lambda x: self.open_add_dialog()]
            ]
        )
        layout.add_widget(self.toolbar)

        # Строка живого поиска
        self.search_field = MDTextField(
            hint_text="Поиск по бренду, оттенку или штрихкоду...",
            size_hint_x=0.95,
            pos_hint={"center_x": 0.5},
            on_text_validate=self.filter_list
        )
        layout.add_widget(self.search_field)

        # Список красителей
        scroll = MDScrollView()
        self.list_container = MDList()
        scroll.add_widget(self.list_container)
        layout.add_widget(scroll)

        self.root_screen.add_widget(layout)
        self.load_list()
        return self.root_screen

    def load_list(self, query=""):
        self.list_container.clear_widgets()
        items = self.db.get_products(query)
        for item in items:
            p_id, brand, shade, nom, full_c, opened_g, is_arch, _ = item
            total_g = (full_c * nom) + opened_g

            list_item = TwoLineAvatarIconListItem(
                text=f"{brand} — {shade}",
                secondary_text=f"Полных: {full_c} шт | Открыто: {opened_g:.0f} г | Всего: {total_g:.0f} г",
                on_release=lambda x, pid=p_id: self.open_card(pid)
            )
            icon = IconLeftWidget(icon="bottle-tonic-plus-outline" if full_c > 0 else "bottle-tonic-outline")
            list_item.add_widget(icon)
            self.list_container.add_widget(list_item)

    def filter_list(self, instance):
        self.load_list(self.search_field.text)

    # Карточка красителя и списание
    def open_card(self, prod_id):
        prod = self.db.get_product(prod_id)
        # prod: (id, brand, shade, barcode, nominal_g, full_count, opened_weight_g, is_archived, notes)
        total_g = (prod[5] * prod[4]) + prod[6]

        content = MDBoxLayout(orientation="vertical", spacing="12dp", size_hint_y=None, height="300dp")
        content.add_widget(MDLabel(text=f"Штрихкод: {prod[3] or 'Не задан'}\nНоминал: {prod[4]} г\nЗаметки: {prod[8] or '-'}", theme_text_color="Secondary"))
        content.add_widget(MDLabel(text=f"ИТОГО НА СКЛАДЕ: {total_g:.1f} г ({prod[5]} целых + {prod[6]:.1f} г в начатых)", bold=True))

        grams_input = MDTextField(hint_text="Граммы для списания (например 25)", input_filter="float")
        content.add_widget(grams_input)

        btn_box = MDBoxLayout(spacing="10dp")
        btn_box.add_widget(MDRaisedButton(
            text="-1 Целый флакон",
            on_release=lambda x: self.exec_write_off_full(prod_id)
        ))
        btn_box.add_widget(MDRaisedButton(
            text="Списать граммы",
            on_release=lambda x: self.exec_write_off_grams(prod_id, grams_input.text)
        ))
        content.add_widget(btn_box)

        self.dialog = MDDialog(
            title=f"{prod[1]} {prod[2]}",
            type="custom",
            content_cls=content,
            buttons=[
                MDRaisedButton(text="В архив", on_release=lambda x: self.archive_item(prod_id)),
                MDRaisedButton(text="Закрыть", on_release=lambda x: self.dialog.dismiss())
            ]
        )
        self.dialog.open()

    def exec_write_off_full(self, prod_id):
        self.db.write_off_full(prod_id)
        self.dialog.dismiss()
        self.load_list(self.search_field.text)

    def exec_write_off_grams(self, prod_id, grams_text):
        try:
            g = float(grams_text)
            self.db.write_off_weight(prod_id, g)
            self.dialog.dismiss()
            self.load_list(self.search_field.text)
        except ValueError:
            pass

    def archive_item(self, prod_id):
        self.db.toggle_archive(prod_id)
        self.dialog.dismiss()
        self.load_list(self.search_field.text)

    # Диалог добавления позиции
    def open_add_dialog(self):
        content = MDBoxLayout(orientation="vertical", spacing="10dp", size_hint_y=None, height="320dp")
        b_input = MDTextField(hint_text="Бренд (Wella, L'Oreal)")
        s_input = MDTextField(hint_text="Номер оттенка (9/16, 7.11)")
        bc_input = MDTextField(hint_text="Штрихкод")
        nom_input = MDTextField(hint_text="Номинальный вес (г)", text="60", input_filter="float")
        cnt_input = MDTextField(hint_text="Полных штук", text="1", input_filter="int")
        notes_input = MDTextField(hint_text="Заметки/формулы")

        for widget in [b_input, s_input, bc_input, nom_input, cnt_input, notes_input]:
            content.add_widget(widget)

        def save_action(x):
            if b_input.text and s_input.text:
                self.db.add_or_update(
                    brand=b_input.text,
                    shade=s_input.text,
                    barcode=bc_input.text or None,
                    nominal_g=float(nom_input.text or 60.0),
                    full_cnt=int(cnt_input.text or 0),
                    opened_g=0.0,
                    notes=notes_input.text
                )
                self.dialog.dismiss()
                self.load_list()

        self.dialog = MDDialog(
            title="Добавить краситель",
            type="custom",
            content_cls=content,
            buttons=[
                MDRaisedButton(text="Отмена", on_release=lambda x: self.dialog.dismiss()),
                MDRaisedButton(text="Сохранить", on_release=save_action)
            ]
        )
        self.dialog.open()

    def export_to_excel(self):
        self.db.export_excel("salon_export.xlsx")

if __name__ == "__main__":
    SalonStockApp().run()