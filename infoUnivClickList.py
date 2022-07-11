import tkinter

window = tkinter.Tk()
window.title("개별대학 테이블 목록")
window.geometry("320x480")
window.resizable(False, False)

iaif_list = {
    "IAIF5103US": "IAIF5103_14",
    "IAIF5111US": "IAIF5111_2014",
    "IAIF5352US": "IAIF5352_14",
    "IAIF5342US": "",
    "IAIF5343US": "IAIF5347_19",
    "IAIF5525US": "IAIF5525",
    "IAIF5515US": "IAIF5515_18",
    "IAIF5515US_20": "IAIF5515_20",
    "IAIF5515US_22": "IAIF5515_22",
    "IAIF7352US": "IAIF7352_14",
    "IAIF7525US": "IAIF7525",
    "IAIF5357US": "IAIF5357_14",
    "IAIF7357US": "IAIF7357_14",
}
iaif_key_list = list(iaif_list.keys())

check_us_name = iaif_key_list[0]
check_ref_name = iaif_list[check_us_name]

RadioVariety_1 = tkinter.StringVar(value=check_us_name)


def radio_click():
    global check_us_name, check_ref_name
    check_us_name = RadioVariety_1.get()
    check_ref_name = iaif_list[check_us_name]
    window.destroy()


for no, el in enumerate(iaif_key_list):
    tkinter.Radiobutton(window, text=el, value=el, variable=RadioVariety_1, command=radio_click).pack()

window.mainloop()
