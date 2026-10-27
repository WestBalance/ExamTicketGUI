from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from excel_handler import get_group_names, get_students
from ticket_parser import TicketParseError
from ticket_service import (
    TicketAssignment,
    TicketAssignmentError,
    TicketSaveError,
    TicketService,
)


APPLICATION_TITLE = "Генератор экзаменационных билетов"


def create_window(
    students_path: str | Path | None = None,
    tickets_path: str | Path | None = None,
    results_path: str | Path | None = None,
) -> tk.Tk:
    project_directory = Path(__file__).resolve().parent
    students_file = (
        Path(students_path) if students_path else project_directory / "students.xlsx"
    )
    tickets_file = (
        Path(tickets_path) if tickets_path else project_directory / "tickets.docx"
    )
    results_file = (
        Path(results_path) if results_path else project_directory / "results.xlsx"
    )

    window = tk.Tk()
    window.title(APPLICATION_TITLE)
    window.geometry("650x400")
    window.minsize(520, 360)

    ticket_service = TicketService(tickets_file, results_file)
    students_by_display: dict[str, dict[str, str]] = {}

    content = ttk.Frame(window, padding=24)
    content.pack(fill="both", expand=True)
    content.columnconfigure(0, weight=1)

    heading = ttk.Label(
        content,
        text=APPLICATION_TITLE,
        font=("TkDefaultFont", 16, "bold"),
    )
    heading.grid(row=0, column=0, pady=(8, 28))

    ttk.Label(content, text="Выберите группу").grid(
        row=1, column=0, sticky="w", pady=(0, 6)
    )
    group_combobox = ttk.Combobox(content, state="readonly")
    group_combobox.grid(row=2, column=0, sticky="ew", pady=(0, 18))

    ttk.Label(content, text="Выберите студента").grid(
        row=3, column=0, sticky="w", pady=(0, 6)
    )
    student_combobox = ttk.Combobox(content, state="readonly")
    student_combobox.grid(row=4, column=0, sticky="ew", pady=(0, 22))

    generate_button = ttk.Button(
        content,
        text="Сгенерировать билет",
        state="disabled",
    )
    generate_button.grid(row=5, column=0)

    status_label = ttk.Label(
        content,
        text="",
        wraplength=580,
        justify="center",
    )
    status_label.grid(row=6, column=0, sticky="ew", pady=(18, 0))

    def update_generate_state() -> None:
        ready = bool(group_combobox.get() and student_combobox.get())
        generate_button.configure(state="normal" if ready else "disabled")

    def on_group_selected(_event: tk.Event) -> None:
        group_name = group_combobox.get()
        student_combobox.set("")
        students_by_display.clear()
        students = get_students(students_file, group_name)
        display_names = []
        for student in students:
            display_name = f"{student['last_name']} {student['first_name']}"
            display_names.append(display_name)
            students_by_display[display_name] = student
        student_combobox.configure(values=display_names)
        if students:
            status_label.configure(text="")
        else:
            status_label.configure(text="В выбранной группе нет студентов.")
        update_generate_state()

    def close_result(result_window: tk.Toplevel) -> None:
        if result_window.winfo_exists():
            result_window.grab_release()
            result_window.destroy()
        student_combobox.set("")
        update_generate_state()

    def show_result(assignment: TicketAssignment) -> None:
        result_window = tk.Toplevel(window)
        result_window.title("Экзаменационный билет")
        result_window.geometry("560x400")
        result_window.minsize(420, 300)
        result_window.transient(window)
        result_window.columnconfigure(0, weight=1)
        result_window.rowconfigure(1, weight=1)

        result_heading = ttk.Label(
            result_window,
            text=f"Ваш билет № {assignment.ticket_number}",
            font=("TkDefaultFont", 15, "bold"),
        )
        result_heading.grid(row=0, column=0, padx=24, pady=(22, 16))

        questions_frame = ttk.Frame(result_window, padding=(24, 0))
        questions_frame.grid(row=1, column=0, sticky="nsew")
        questions_frame.columnconfigure(0, weight=1)
        for index, question in enumerate(assignment.questions, start=1):
            question_label = ttk.Label(
                questions_frame,
                text=f"{index}. {question}",
                wraplength=500,
                justify="left",
                font=("TkDefaultFont", 11),
            )
            question_label.grid(
                row=index - 1, column=0, sticky="ew", pady=(0, 14)
            )

        close_button = ttk.Button(
            result_window,
            text="Закрыть",
            command=lambda: close_result(result_window),
        )
        close_button.grid(row=2, column=0, pady=20)
        result_window.bind(
            "<Escape>", lambda _event: close_result(result_window)
        )
        result_window.protocol(
            "WM_DELETE_WINDOW", lambda: close_result(result_window)
        )
        result_window.grab_set()
        close_button.focus_set()

    def save_with_retry(
        group_name: str, student: dict[str, str]
    ) -> TicketAssignment | None:
        try:
            return ticket_service.generate_ticket(
                group_name, student["last_name"], student["first_name"]
            )
        except TicketSaveError as error:
            while True:
                should_retry = messagebox.askretrycancel(
                    "Не удалось сохранить билет",
                    str(error),
                    parent=window,
                )
                if not should_retry:
                    status_label.configure(
                        text="Билет не сохранён. Можно повторить запрос позже."
                    )
                    return None

                try:
                    return ticket_service.retry_save(
                        group_name, student["last_name"], student["first_name"]
                    )
                except TicketSaveError as retry_error:
                    error = retry_error

    def on_generate() -> None:
        group_name = group_combobox.get()
        student = students_by_display.get(student_combobox.get())
        if not group_name or student is None:
            status_label.configure(text="Выберите группу и студента.")
            update_generate_state()
            return

        if not tickets_file.is_file():
            messagebox.showerror(
                "Файл не найден",
                f"Не найден файл с экзаменационными билетами:\n{tickets_file}",
                parent=window,
            )
            return

        generate_button.configure(state="disabled")
        status_label.configure(text="Сохранение билета...")
        window.update_idletasks()
        try:
            assignment = save_with_retry(group_name, student)
            if assignment is not None:
                status_label.configure(text="Билет успешно сохранён.")
                show_result(assignment)
        except (
            TicketParseError,
            TicketAssignmentError,
            OSError,
            ValueError,
        ) as error:
            messagebox.showerror("Ошибка", str(error), parent=window)
            status_label.configure(text="Не удалось сгенерировать билет.")
        finally:
            update_generate_state()

    group_combobox.bind("<<ComboboxSelected>>", on_group_selected)
    student_combobox.bind(
        "<<ComboboxSelected>>", lambda _event: update_generate_state()
    )
    generate_button.configure(command=on_generate)

    if not students_file.is_file():
        messagebox.showerror(
            "Файл не найден",
            f"Не найден файл со списком студентов:\n{students_file}",
            parent=window,
        )
        status_label.configure(text="Файл students.xlsx не найден.")
    else:
        group_names = get_group_names(students_file)
        group_combobox.configure(values=group_names)
        if not group_names:
            status_label.configure(
                text="Не удалось загрузить группы из файла students.xlsx."
            )

    return window
