# Module imports (tkinter and turtle modules)
import tkinter as tk
import turtle as tl


# Window class declaration
class Window:
    # Constructor
    def __init__(self):
        # Main window and details of window
        self.master = tk.Tk()
        self.master.title("Turtle Pathfinder")
        self.master.geometry("1200x600")
        self.master.resizable(False, False)

        # Frames in window
        self.main_frame = tk.Frame(self.master)
        self.main_frame.grid(column=0, row=0)

        # Main canvas in main frame
        self.canvas = tk.Canvas(self.main_frame, width=895, height=595)
        self.canvas.grid(column=0, row=0)

        self.block = Block(tl.RawTurtle(self.canvas), 0, 0, height=10, width=10)

        self.master.mainloop()


# Block object
class Block:
    def __init__(self, turtle, x, y, width, height):
        # turtle/object list and global variable declarations
        self.clone_list = []
        self.height = height * 10
        self.width = width * 10
        self.origin_x = (x - self.width, x + self.width)
        self.origin_y = (y - self.height, y + self.height)
        # turtle settings
        self.turtle = turtle
        self.turtle.turtlesize(height, width)
        self.turtle.shape("square")

        # turtle position
        self.turtle.up()
        self.turtle.goto(x, y)
        self.turtle.down()
        # Setting turtle commands
        self.turtle.onclick(self.onclick, btn=1)
        self.turtle.onrelease(self.release, btn=1)

    # When clicking the block
    def onclick(self, x, y):
        # clones turtle and stores clone's data in list
        clone = self.turtle.clone()
        self.clone_list.append(
            {"object": clone,
             "position": clone.position(),
             "xcoords": (),
             "ycoords": ()}
        )
        clone.up()
        clone.hideturtle()

    # On the release of clicking the block
    def release(self, x, y):
        # Places the turtle in the canvas and updates settings
        clone = self.clone_list[-1]
        clone_object = self.clone_list[-1]["object"]
        # Checks if the block is not in range of the original to make sure there's no bugs
        if not self.bounds_check(x, y):
            clone["object"].hideturtle()
            self.clone_list.remove(clone)
            del clone
            return
        clone_object.goto(x, y)
        # Sets the new coords for the clone
        clone["xcoords"] = (clone_object.xcor() - self.height, clone_object.xcor() + self.height)
        clone["ycoords"] = (clone_object.ycor() - self.width, clone_object.ycor() + self.width)
        clone_object.showturtle()
        # command for removing turtle
        clone_object.onclick(self.remove, btn=3)

    # Command for right click on clones
    def remove(self, x, y):
        # Searches through the list and deletes the clone selected
        for turtle in self.clone_list:
            if turtle["xcoords"][0] <= x <= turtle["xcoords"][1] and turtle["ycoords"][0] <= y <= turtle["ycoords"][1]:
                turtle["object"].hideturtle()
                self.clone_list.remove(turtle)
                del turtle
                break

    # Checks to see if object is in correct bounds
    # This is mainly to avoid clones covering the original
    def bounds_check(self, x, y):
        # 450 by 300
        if self.origin_x[0]*2 <= x <= self.origin_x[1]*2 and self.origin_y[0]*2 <= y <= self.origin_y[1]*2:
            return False
        # Returns false if not in bounds and true if in bounds
        if self.origin_x[0]*2 <= x <= self.origin_x[1]*2 and self.origin_y[0]*2 <= y <= self.origin_y[1]*2:
            return False
        return True


# Code run when script is run
if __name__ == '__main__':
    Window()
