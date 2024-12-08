import tkinter as tk
from tkinter import ttk, filedialog
from tkinter import messagebox
from tkinter import PhotoImage
from PIL import Image, ImageTk  # Import PIL for handling images
import matplotlib.animation as animation
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import scipy.signal
import serial
from serial import SerialException
import peakutils
import xlwt
import threading
import time

# Variables to manage serial data and recording
recording = False
serialDataRecorded = []
serialOpen = False
global ser

OptionList = [
    "--Select a COM port--", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18"
] 

connected = False
serialData = [0] * 70

# Function to read from the serial port
def read_from_port(ser):
    global serialOpen
    global serialData
    global serialDataRecorded
    while serialOpen == True:
        reading = float(ser.readline().strip())
        serialData.append(reading)
        if recording == True:
            serialDataRecorded.append(reading)

# Function to start the serial connection
def startSerial():
    try:
        s = var.get()
        global ser
        ser = serial.Serial('COM' + s , 9600, timeout=20)
        ser.close()
        ser.open()
        global serialOpen
        serialOpen = True
        global thread
        thread = threading.Thread(target=read_from_port, args=(ser,))
        thread.start()
        connectText.set("Connected to COM" + s)
        labelConnect.config(fg="green")
        return serialOpen
        
    except SerialException:
        connectText.set("Error: wrong port?")
        labelConnect.config(fg="red")
    
# Function to close the serial connection
def kill_Serial():
    try:
        global ser
        global serialOpen
        serialOpen = False
        time.sleep(1)
        ser.close()
        connectText.set("Not connected")
        labelConnect.config(fg="red")
    except:
        connectText.set("Failed to end serial ")

# Function to animate the data
def animate(i):
    global serialData
    if len(serialData) > 70:
        serialData = serialData[-70:]
    data = serialData.copy()
    data = data[-70:]
    global ax
    x = np.linspace(0, 69, dtype='int', num=70)
    ax.clear()
    ax.plot(x, data)

# Function to start recording data
def startRecording():
    global recording
    if serialOpen:
        recording = True
        recordText.set("Recording . . . ")
        labelRecord.config(fg="red")
    else:
        messagebox.showinfo("Error", "Please start the serial monitor")

# Function to stop recording data
def stopRecording():
    global recording
    global serialDataRecorded
    if recording == True:
        recording = False
        recordText.set("Not Recording")
        labelRecord.config(fg="black")
        processRecording(serialDataRecorded)
        serialDataRecorded = []  # Clear recorded data
    else:
        messagebox.showinfo("Error", "You weren't recording!")

# Function to process the recorded data
def processRecording(data):
    z = scipy.signal.savgol_filter(data, 11, 3)
    data2 = np.asarray(z, dtype=np.float32)
    a = len(data)
    base = peakutils.baseline(data2, 2)
    y = data2 - base

    # Save data to Excel
    directory = filedialog.asksaveasfilename(defaultextension=".xls", filetypes=(("Excel Sheet", "*.xls"),("All Files", "*.*")))
    if directory is None:
        return

    # Create Excel sheet and write data
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Sheet 1")

    # Write heart rate data
    for i in range(a):
        sheet1.write(i + 3, 0, i)  # Row for time/index
        sheet1.write(i + 3, 1, y[i])  # Row for processed heart rate data
    
    book.save(directory)

# Function to show image in the popup
def show_image(image_path, text):
    # Open the image
    img = Image.open(image_path)
    img = img.resize((250, 250))  # Resize image for display
    img = ImageTk.PhotoImage(img)

    # Create a label to display the image
    label_img = tk.Label(popup, image=img)
    label_img.image = img  # Keep reference to avoid garbage collection
    label_img.grid(row=1, column=0, padx=10, pady=10)

    # Add text next to the image
    label_text = tk.Label(popup, text=text, font=("Helvetica", 12))
    label_text.grid(row=1, column=1, padx=10, pady=10, sticky="n")

# Function to open the placement popup with 3 image buttons
def open_popup():
    global popup
    popup = tk.Toplevel(window)
    popup.title("Image Placement")

    # Add buttons for each image with different texts
    btn_img1 = tk.Button(popup, text="Adult", command=lambda: show_image("image1.png", "This is the best placement in terms of monitoring, also used commonly in health care and also works on pediatrics"))
    btn_img1.grid(row=0, column=0, padx=10, pady=10)

    btn_img2 = tk.Button(popup, text="Image 2", command=lambda: show_image("image2.png", "This is the best placement in terms of recording data, has been scientifaclly proven in studies and works on both adults and pediatrics"))
    btn_img2.grid(row=0, column=1, padx=10, pady=10)

    btn_img3 = tk.Button(popup, text="Image 3", command=lambda: show_image("image3.jpg", "This is the triangle placement it forms a triangle via the electrodes it is also commonly used in other countries, and works on adults and pediatrics"))
    btn_img3.grid(row=0, column=2, padx=10, pady=10)

# Set up Tkinter window
window = tk.Tk()
window.title("tikarr v1")
window.iconbitmap("tikarv1icon.ico")  # Replace 'your_icon.ico' with your icon's filename

window.rowconfigure(0, minsize=800, weight=1)
window.columnconfigure(1, minsize=800, weight=1)

# Set up the matplotlib figure with larger dimensions
fig = Figure(figsize=(10, 8), dpi=50)  # Increase the figure size (width, height)
ax = fig.add_subplot(1, 1, 1)
ax.set_xlim([0, 10])
ax.set_ylim([0, 150])

# Create the canvas with the updated figure size
canvas = FigureCanvasTkAgg(fig, master=window)
canvas.draw()

# Ensure the canvas widget fills more space in the layout
canvas.get_tk_widget().grid(row=1, column=1, sticky="nsew")

# Frame for buttons and other controls
fr_buttons = tk.Frame(window, relief=tk.RAISED, bd=2)

# Placement button in sidebar
btn_placement = tk.Button(fr_buttons, text="            Placement          ", command=open_popup)
btn_placement.grid(row=0, column=0, padx=10, pady=5)

# Connect/Disconnect label
connectText = tk.StringVar(window)
connectText.set("Not connected")
labelConnect = tk.Label(fr_buttons, textvariable=connectText, font=('Helvetica', 12), fg='red')
labelConnect.grid(row=1, column=0, sticky="ew", padx=10)

# COM port selection dropdown
var = tk.StringVar(window)
var.set(OptionList[0])
opt_com = tk.OptionMenu(fr_buttons, var, *OptionList)
opt_com.config(width=20)
opt_com.grid(row=2, column=0, sticky="ew", padx=10)

# Buttons for starting and stopping the serial and recording
btn_st_serial = tk.Button(fr_buttons, text="Open Serial", command=startSerial)
btn_st_serial.grid(row=3, column=0, sticky="ew", padx=10, pady=5)

btn_stop_serial = tk.Button(fr_buttons, text="Close Serial", command=kill_Serial)
btn_stop_serial.grid(row=4, column=0, sticky="ew", padx=10, pady=5)

recordText = tk.StringVar(window)
recordText.set("Not Recording")
labelRecord = tk.Label(fr_buttons, textvariable=recordText, font=('Helvetica', 12), fg='black')
labelRecord.grid(row=5, column=0, sticky="ew", padx=10)

btn_st_rec = tk.Button(fr_buttons, text="Start Recording", command=startRecording)
btn_st_rec.grid(row=6, column=0, sticky="ew", padx=10, pady=5)

btn_stop_rec = tk.Button(fr_buttons, text="Stop Recording", command=stopRecording)
btn_stop_rec.grid(row=7, column=0, sticky="ew", padx=10, pady=5)

fr_buttons.grid(row=0, column=0, sticky="ns")

# Live Data Plot
lbl_live = tk.Label(text="Live Data:", font=('Helvetica', 12), fg='red')
lbl_live.grid(row=0, column=1, sticky="nsew")
canvas.get_tk_widget().grid(row=1, column=1, sticky="nsew")

# This closes the serial and destroys the window
def ask_quit():
    if tk.messagebox.askokcancel("Quit", "This will end the serial connection "):
        kill_Serial()
        window.destroy()

window.protocol("WM_DELETE_WINDOW", ask_quit)

# Animation for live data
ani = animation.FuncAnimation(fig, animate, interval=50)
window.mainloop()
