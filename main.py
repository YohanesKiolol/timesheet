import openpyxl
import ttkbootstrap as ttk
from datetime import datetime, timedelta
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText
from ttkbootstrap.tableview import Tableview, TableColumn
from openpyxl.styles import Border, Side, Alignment, Font

path = "/Users/yohaneskiolol/Library/CloudStorage/OneDrive-SharedLibraries-ComputradeTechnologyInternational/[CSD] RnD - 2024/Timesheet Yohan.xlsx"


class TimesheetForm(ttk.Frame):

    def __init__(self, master):
        super().__init__(master, padding=(20, 10))
        self.pack(fill=BOTH, expand=YES)
        start_date = datetime.now()
        minutes = round((start_date.minute + 7.5) / 15) * 15
        if minutes > 45:
            start_date += timedelta(hours=1)
            minutes = 0

        start_date = start_date.replace(minute=minutes)
        start_date = start_date.replace(second=0)
        end_date = start_date + timedelta(hours=1)
        fstart_date = start_date.strftime("%d/%m/%Y %H:%M:%S")
        fend_date = end_date.strftime("%d/%m/%Y %H:%M:%S")

        # form variables
        self.name = ttk.StringVar(value="Yohan")
        self.start = ttk.StringVar(value=fstart_date)
        self.end = ttk.StringVar(value=fend_date)
        self.detail = ttk.StringVar(value="")

        # form entries

        form_frame = ttk.Frame(self)
        form_frame.pack(side=LEFT, fill=Y)

        # # form header
        # f_hdr = "Insert New Timesheet"
        # f_lbl = ttk.Label(form_frame, text=f_hdr, width=50)
        # f_lbl.pack(fill=X, pady=10)

        self.create_form_entry("NAME", self.name, form_frame)
        self.create_form_entry("START", self.start, form_frame, 'date')
        self.create_form_entry("END", self.end, form_frame, 'date')
        self.create_form_entry("DETAIL", self.detail, form_frame, 't_area')
        self.create_buttonbox(form_frame)

        table_frame = ttk.Frame(self)
        table_frame.pack(side=RIGHT, fill=BOTH, expand=YES)

        self.create_table()

    def create_table(self):
        l_frame = ttk.Frame(self)
        l_frame.pack(fill=X)
        # l_hdr = "Timesheet"
        # l_lbl = ttk.Label(l_frame, text=l_hdr)
        # l_lbl.pack(side=LEFT, padx=5)

        del_btn = ttk.Button(
            master=l_frame,
            text="Delete",
            command=self.on_delete,
            bootstyle=DANGER
        )
        del_btn.pack(side=RIGHT, padx=5)

        columns = [
            {"text": "NAME", "stretch": False},
            {"text": "START", "stretch": False},
            {"text": "END", "stretch": False},
            {"text": "DETAIL", "stretch": False},
            {"text": "MANDAYS", "stretch": False},
            {"text": "NONE", "stretch": False},
        ]

        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active
        list_values = list(sheet.values)

        self.treeview = Tableview(self, coldata=columns, paginated=True,
                                  searchable=True, rowdata=list_values[1:][::-1])

        TableColumn(self.treeview, 4, 'MANDAYS').hide()
        TableColumn(self.treeview, 5, 'NONE').hide()
        # self.treeview.view.bind("<<TreeviewSelect>>", self.tetete())
        # self.treeview.view.bind("<<TreeviewSelect>>", tetete)
        self.treeview.pack(fill=BOTH, expand=YES, padx=10, pady=10)

    def refresh_table(self):
        # Clear existing data in the treeview
        self.treeview.delete_rows()

        # Reload the data into the table
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active
        list_values = list(sheet.values)

        self.treeview.insert_rows(0, list_values[1:][::-1])
        self.treeview.load_table_data()

    def create_form_entry(self, label, variable, container, type=None):

        container = ttk.Frame(container)
        container.pack(fill=X, expand=YES, pady=5)

        lbl = ttk.Label(master=container, text=label.title(), width=10)
        lbl.pack(side=LEFT, padx=5)

        if type == "date":
            date_object = datetime.strptime(
                variable.get(), "%d/%m/%Y %H:%M:%S")
            l_minutes = [0, 15, 30, 45]
            fdate = ttk.DateEntry(
                master=container, dateformat='%d/%m/%Y')
            fdate.pack(side=LEFT, padx=5)

            fhour = ttk.Spinbox(
                master=fdate, from_=9, to=19, width=3)
            fhour.pack(side=LEFT, padx=5)
            fhour.set(date_object.hour)

            fmin = ttk.Combobox(
                master=fdate, values=l_minutes, width=3)
            fmin.pack(side=LEFT)
            m_idx = l_minutes.index(date_object.minute)
            fmin.current(m_idx)

            fdate.bind("<FocusOut>", lambda event,
                       sv=variable: self.one_date_change(sv, fdate.entry.get()))

            fhour.bind("<FocusOut>", lambda event,
                       sv=variable: self.on_hour_change(sv, fhour.get()))

            fmin.bind("<FocusOut>", lambda event,
                      sv=variable: self.on_min_change(sv, fmin.get()))

        elif type == 't_area':
            field = ScrolledText(
                master=container, height=8, width=45, wrap=WORD, autohide=TRUE)
            field.pack(side=LEFT, padx=3)

            field.bind("<FocusOut>", lambda event,
                       sv=variable: self.on_change(sv, field.get("1.0", "end-1c")))

        else:
            field = ttk.Entry(master=container, textvariable=variable)
            field.pack(side=LEFT, padx=5, fill=X)

    def on_change(self, variable, new_value):
        if variable.get() != new_value:
            variable.set(new_value)

    def one_date_change(self, variable, new_value):
        new_date = datetime.strptime(new_value, "%d/%m/%Y")
        date_obj = datetime.strptime(variable.get(), "%d/%m/%Y %H:%M:%S")
        updated_date_obj = date_obj.replace(
            day=new_date.day, month=new_date.month, year=new_date.year)
        variable.set(updated_date_obj.strftime("%d/%m/%Y %H:%M:%S"))

    def on_hour_change(self, variable, new_value):
        date_obj = datetime.strptime(variable.get(), "%d/%m/%Y %H:%M:%S")
        updated_date_obj = date_obj.replace(hour=int(new_value))
        variable.set(updated_date_obj.strftime("%d/%m/%Y %H:%M:%S"))

    def on_min_change(self, variable, new_value):
        date_obj = datetime.strptime(variable.get(), "%d/%m/%Y %H:%M:%S")
        updated_date_obj = date_obj.replace(minute=int(new_value))
        variable.set(updated_date_obj.strftime("%d/%m/%Y %H:%M:%S"))

    def create_buttonbox(self, container):
        container = ttk.Frame(container)
        container.pack(fill=X, expand=YES, pady=(15, 10))

        sub_btn = ttk.Button(
            master=container,
            text="Submit",
            command=self.on_submit,
            bootstyle=SUCCESS,
            width=6,
        )
        sub_btn.pack(side=LEFT, padx=5)
        sub_btn.focus_set()

        cnl_btn = ttk.Button(
            master=container,
            text="Cancel",
            command=self.on_cancel,
            bootstyle=DANGER,
            width=6,
        )
        cnl_btn.pack(side=LEFT, padx=5)

    def on_submit(self):
        name = self.name.get()
        start = self.start.get()
        end = self.end.get()
        detail = self.detail.get()

        # print("Name:", name)
        # print("Start:", start)
        # print("End:", end)
        # print("Detail:", detail)

        start = datetime.strptime(start, "%d/%m/%Y %H:%M:%S")
        start_formatted = start.strftime("%d-%b-%y %H.%M")

        end = datetime.strptime(end, "%d/%m/%Y %H:%M:%S")
        end_formatted = end.strftime("%d-%b-%y %H.%M")

        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active

        border_style = Border(left=Side(border_style='thin', color='000000'),
                              right=Side(border_style='thin', color='000000'),
                              top=Side(border_style='thin', color='000000'),
                              bottom=Side(border_style='thin', color='000000'))

        alignment_style = Alignment(
            wrapText=TRUE, horizontal=CENTER, vertical=CENTER)
        
        font_style = Font(size = "10")

        new_row_index = sheet.max_row + 1
        name_cell = sheet.cell(row=new_row_index, column=1)
        name_cell.value = name
        name_cell.border = border_style
        name_cell.alignment = alignment_style
        name_cell.font = font_style

        start_cell = sheet.cell(row=new_row_index, column=2)
        start_cell.value = start_formatted
        start_cell.border = border_style
        start_cell.alignment = alignment_style
        start_cell.font = font_style

        end_cell = sheet.cell(row=new_row_index, column=3)
        end_cell.value = end_formatted
        end_cell.border = border_style
        end_cell.alignment = alignment_style
        end_cell.font = font_style

        detail_cell = sheet.cell(row=new_row_index, column=4)
        detail_cell.value = detail
        detail_cell.border = border_style
        detail_cell.alignment = Alignment(
            wrapText=TRUE, horizontal=LEFT, vertical=TOP)
        detail_cell.font = font_style

        mandays_cell = sheet.cell(row=new_row_index, column=5)
        mandays_cell.value = f'=ROUND((HOUR(${end_cell.coordinate}-${start_cell.coordinate})/8), 2)'
        mandays_cell.border = border_style
        mandays_cell.alignment = alignment_style
        mandays_cell.font = font_style

        workbook.save(path)

        self.refresh_table()

    def on_cancel(self):
        self.quit()

    def on_delete(self):
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active

        s_rows = self.treeview.view.selection()
        for row_id in s_rows:
            value = self.treeview.view.item(row_id, 'values')
            for row in sheet.iter_rows(values_only=True):
                t_row = tuple(
                    'None' if value is None else value for value in row)
                if t_row == value:
                    index = list(sheet.iter_rows(values_only=True)).index(row)
                    sheet.delete_rows(index+1, 1)
            self.treeview.delete_row(iid=row_id)
        workbook.save(path)
        self.treeview.load_table_data()


if __name__ == "__main__":
    app = ttk.Window("Timesheet", "solar", resizable=(False, False))
    timesheet = TimesheetForm(app)
    timesheet.pack(padx=10, pady=10)
    app.mainloop()
