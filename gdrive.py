from tkinter import *
import customtkinter  # <- import the CustomTkinter module
from PIL import Image, ImageTk
import tkinter.ttk as ttk

BG_GRAY = "#000000"
BG_COLOR = "#000000"
MSG_ENTRY_COLOR = "#000000"
TEXT_COLOR = "#EAECEE"

FONT = "Miriam 11"
FONT_BOLD = "Helvetica 13 bold"


class ChatApplication:

    def __init__(self):
        self.window = customtkinter.CTk()
        self._setup_main_window()

    def run(self):
        self.window.mainloop()

    def _setup_main_window(self):
        self.window.title("Vermera Virtual Assistant")
        self.window.resizable(width=False, height=False)
        self.window.configure(width=600, height=550, bg=BG_COLOR)

        # head label
        head_label = customtkinter.CTkLabel(self.window, bg=BG_COLOR, fg=TEXT_COLOR,text_font=FONT_BOLD,text="VERMEG", pady=10)
        head_label.place(relwidth=1)

        # text widget
        self.text_widget = Text(self.window, width=20, height=2, bg=BG_COLOR, fg=TEXT_COLOR,
                                font=FONT)
        self.text_widget.place(relheight=0.88, relwidth=1, rely=0.1)
        self.text_widget.configure(cursor="arrow", state=DISABLED)

        # scroll bar
        scrollbar = Scrollbar(self.text_widget)
        scrollbar.place(relheight=1, relx=0.974)
        scrollbar.configure(command=self.text_widget.yview)

        # bottom label
        bottom_label = Label(self.window, bg=BG_GRAY, height=40)
        bottom_label.place(relwidth=1, rely=0.9)

        # message entry box
        self.msg_entry = customtkinter.CTkEntry(bottom_label, bg=MSG_ENTRY_COLOR, fg=TEXT_COLOR)
        self.msg_entry.place(relwidth=0.74, relheight=0.05, rely=0.008, relx=0.011)
        self.msg_entry.focus()
        self.msg_entry.bind("<Return>", self._on_enter_pressed)

        # send button
        send_image = ImageTk.PhotoImage(Image.open("send.png").resize((20, 20), Image.ANTIALIAS))
        send_button = customtkinter.CTkButton(master=bottom_label, image=send_image, text="", width=50, height=50,
                                              corner_radius=10, fg_color="gray10", hover_color="gray25",
                                              command=lambda: self._on_enter_pressed(None))

        send_button.place(relx=0.76, rely=0.008, relheight=0.05, relwidth=0.11)

        micro_image = ImageTk.PhotoImage(Image.open("micro.png").resize((20, 20), Image.ANTIALIAS))
        micro_button = customtkinter.CTkButton(master=bottom_label, image=micro_image, text="", width=50, height=50,
                                               corner_radius=10, fg_color="#f54251", hover_color="#ed5562",
                                               command=lambda: self._on_enter_pressed(None))

        micro_button.place(relx=0.88, rely=0.008, relheight=0.05, relwidth=0.11)

    def _on_enter_pressed(self, event):
        msg = self.msg_entry.get()
        self._insert_message(msg, "Vermera")

    def _insert_message(self, msg, sender):
        if not msg:
            return

        self.msg_entry.delete(0, END)

        # Use created style in this frame
        # frame = customtkinter.CTkFrame(self.text_widget,width=200,height=40, )
        # frame.configure(fg_color='#428069',corner_radius=20)

        label_user = customtkinter.CTkLabel(master=self.text_widget, text=msg, fg_color="#6677d9", width=100, height=29,
                                       corner_radius=20, text_font=FONT)
        label_user.pack(padx=20, pady=5, anchor=E)
        label = customtkinter.CTkLabel(master=self.text_widget, text=msg, fg_color="#17215e", width=150, height=29,
                                       corner_radius=20, text_font=FONT)
        label.pack(padx=5, pady=5, anchor=W)


if __name__ == "__main__":
    app = ChatApplication()
    app.run()
