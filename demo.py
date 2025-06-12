import tkinter as tk

root = tk.Tk()
root.geometry("800x600")

# Configure root to be responsive
root.rowconfigure(0, weight=1)
root.columnconfigure(0, weight=1)

# Create main frame
main_frame = tk.Frame(root)
main_frame.grid(row=0, column=0, sticky="nsew")

# Configure internal grid
main_frame.rowconfigure(0, weight=1)
main_frame.columnconfigure(0, weight=1)

# Add a canvas that resizes
canvas = tk.Canvas(main_frame, bg="lightblue")
canvas.grid(row=0, column=0, sticky="nsew")

# Resize the content inside canvas
def on_resize(event):
    canvas.delete("all")
    canvas.create_text(event.width/2, event.height/2, text="I'm Responsive!", font=("Arial", 20))

canvas.bind("<Configure>", on_resize)

root.mainloop()
