# Part 1: Importing Modules & Initializing Variables
import cv2
from xlwt import Workbook
from math import sqrt
from tkinter import *

count = 1
row = 1
col = 0
p = []  # Creating a list 'p'

# Creating a workbook, named 'wb'
wb = Workbook()

# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet 1')
sheet1.write(0, 0, "X-Coordinates")
sheet1.write(0, 1, "Y-Coordinates")


# Part 2: Function to calculate distance between two points.
def distance(p1, p2):
    x0 = p1[0] - p2[0]
    y0 = p1[1] - p2[1]
    return (x0*x0 + y0*y0)


# Part 3: Function to display coordinates of the points.
def click_event(event, x, y, flags, params):

    # checking for left mouse clicks
    if event == cv2.EVENT_LBUTTONDOWN:
        global count
        global row
        global col

        # displaying the coordinates on the Output Window
        print(x, y)

        # write(row, column, data)
        sheet1.write(count, 0, x)
        sheet1.write(count, 1, y)

        p.append([x, y])
        count = count + 1

        # Choosing the font
        font = cv2.FONT_HERSHEY_SIMPLEX

        # Putting text on the image
        cv2.putText(img, str(count-1), (x, y), font, .3, (0, 0, 0), 1)
        cv2.imshow('image', img)

        wb.save("Coordinates.xls")


# Part 4: Function to find the maximum distance between two points
def maxDistance(p):
    # Taking the length of list 'p'
    n = len(p)
    maxm = 0

    global far1
    global far2

    far1 = p[1]
    far2 = p[2]

    # Iterate over all possible pairs
    for i in range(n):
        for j in range(i + 1, n):
            # Update maxm
            maxm = max(maxm, distance(p[i], p[j]))

            # Updating the farthest points (far1, far2)
            if (maxm== distance(p[i], p[j])):
                far1= p[i]
                far2= p[j]

    # Return actual distance
    return sqrt(maxm)


# Part 5: Creating GUI Function.
def GUI():
    window = Tk()

    # specify size & title of the window.
    window.geometry("400x100")
    window.title('Hello Python')

    l = Label(window, text="This is the line between the two farthest points.", fg='red', font=("Helvetica", 14))
    l.pack()
    window.mainloop()


# Part 6: Creating the main function
if __name__ == "__main__":

    # reading the image
    img = cv2.imread('Balloon.jpg')

    # displaying the image
    cv2.imshow('image', img)

    # setting mouse handler for the image and calling the click_event() function
    cv2.setMouseCallback('image', click_event)

    # wait for a key to be pressed to exit
    cv2.waitKey(0)

    n= count
    print("\nDistance b/w two farthest points is: ", maxDistance(p))

    print("\nThe coordinates, which are at the maximum distance are: ")
    print(far1)
    print(far2)

    img = cv2.line(img, far1, far2, (0, 0, 255), 2)
    cv2.imshow("Image", img)

    GUI()
