import tkinter as tk

window = tk.Tk()
# frame = tk.Frame(master=window, width=300, height=150)
# frame.pack()
window.rowconfigure(0, minsize=50)
window.columnconfigure([0, 1, 2, 3, 4], minsize=50)
greeting = tk.Label(text='RCS Reports')
# greeting.grid(row=0, column=0)
greeting.grid(row=0, column=2, sticky="ew")
lbl_client = tk.Label(text='Client ID')
lbl_client.grid(row=1, column=2, sticky='ew')
# lbl_client.place(x=75, y=25)
ent_client = tk.Entry()
# ent_client.place(x=75, y=50)
ent_client.grid(row=2, column=2, sticky='ew')
btn_meter = tk.Button(
    text='Meter',
    width=25,
    height=5,
    bg="blue",
    fg='yellow',
)
btn_hvac = tk.Button(
    text='HVAC',
    width=25,
    height=5,
    bg="blue",
    fg='yellow',
)

# btn_meter.pack()
btn_hvac.grid(row=3, column=0)
# btn_hvac.pack()
btn_meter.grid(row=3, column=3)
clientID = ent_client.get()


window.mainloop()
