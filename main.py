from tkinter import *
from assets.bg_loader import load_bg_image
from assets.bg_base64 import background_image

window = Tk()
window.title("Student Score Tracker by Khian Victory D. Calderon")
window.state("zoomed")

# Loads the background image
load_bg_image(window, background_image)

Label(window, text = "Hello World", padx = 10, pady = 10).pack()

window.mainloop()