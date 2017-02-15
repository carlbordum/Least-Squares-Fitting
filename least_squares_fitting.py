"""
Mindste kvadraters metode
~~~~~~~~~~~~~~~~~~~~~~~~~
    or Least Squares Fitting is part of the product from a school week with the subjects Math,
    Programming and Physics. It is just a simple graphical interface for drawing linear-,
    power- and exponential graphs on a dataset from excel.

    Written in early 2017 by Carl Bordum Hansen.
"""
import tkinter as tk
import openpyxl as pyxl
from functools import partial
import math
import matplotlib
matplotlib.use('TkAgg')
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure


def remove_zeros(xs, ys):
    """Remove all instances of zeros and its' "pair" in two lists."""
    new_xs, new_ys = [], []
    for x, y in zip(xs, ys):
        if x and y:
            new_xs.append(x)
            new_ys.append(y)
    return new_xs, new_ys


class Diagram:
    """Interface between matplotlib and tkinter."""
    def __init__(self, parent):
        self.parent = parent
        self.figure = Figure(figsize=(5, 4), dpi=100)
        self.points = self.figure.add_subplot(111)
        self.graph = self.figure.add_subplot(111)

    def draw_points(self, xs, ys):
        self.points.clear()
        self.points.plot(xs, ys, 'ro')
        canvas = FigureCanvasTkAgg(self.figure, master=self.parent)
        canvas.get_tk_widget().grid(column=1, row=0, rowspan=5)

    def draw_fit(self, equation, equation_str, xs, ys):
        self.graph.clear()
        self.points.plot(xs, ys, 'ro')
        ys = [equation(i) for i in range(int(min(xs)), int(max(xs)) + 1)]
        xs = [i for i in range(int(min(xs)), int(max(xs)) + 1)]
        self.graph.plot(xs, ys)
        self.graph.set_title(equation_str)
        canvas = FigureCanvasTkAgg(self.figure, master=self.parent)
        canvas.get_tk_widget().grid(column=1, row=0, rowspan=5)


class Regression:
    """Implements regression from math.

    Each fit returns the correspoding function and the mathematical expression as a string."""
    @staticmethod
    def _fit(xs, ys):
        a = (sum(x * y for x, y in zip(xs, ys)) - 1 / len(xs) * sum(xs) * sum(ys)) / \
            (sum(x**2 for x in xs) - 1 / len(xs) * (sum(xs))**2)
        b = 1 / len(xs) * (sum(ys) - a * sum(xs))
        return a, b

    @staticmethod
    def linear_fit(xs, ys):
        a, b = Regression._fit(xs, ys)
        return (lambda x: a * x + b, '$y = % .3f x + % .3f$' % (a, b))

    @staticmethod
    def power_fit(xs, ys):
        xs, ys = remove_zeros(xs, ys)
        a, b = Regression._fit(list(map(math.log10, xs)), map(math.log10, ys))  # list, need len()
        b = 10**b
        return (lambda x: x**a + b, '$y = x^{% .3f} + % .3f$' % (a, b))

    @staticmethod
    def exponential_fit(xs, ys):
        xs, ys = remove_zeros(xs, ys)
        a, b = Regression._fit(xs, map(math.log10, ys))
        a, b = 10**a, 10**b
        return (lambda x: a**x + b, '$y = % .3f^x + % .3f$' % (a, b))


class Application(tk.Frame):
    """GUI and user interaction."""
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.diagram = Diagram(self)
        self.diagram.draw_points([], [])
        self.datasets = {}
        self.linear_fit = partial(self._fit, Regression.linear_fit)
        self.power_fit = partial(self._fit, Regression.power_fit)
        self.exponential_fit = partial(self._fit, Regression.exponential_fit)
        self.create_widgets()

    def create_widgets(self):
        self.parent.title('Mindste kvadraters metode')
        self.pack(fill=tk.BOTH, expand=True)
        self.columnconfigure(0, weight=1, pad=40)
        fill = tk.N + tk.S + tk.E + tk.W

        buttons = [
            ('Browse', self.set_datasets),
            ('Linear Fit', self.linear_fit),
            ('Power Fit', self.power_fit),
            ('Exponential Fit', self.exponential_fit)
        ]

        for row, button in enumerate(buttons):
            text, command = button
            new_button = tk.Button(self, text=text, command=command)
            new_button.grid(row=row, sticky=fill)

        self.dropdown = tk.Listbox(self)
        self.dropdown.grid(row=4, sticky=fill)
        self.dropdown.bind('<<ListboxSelect>>', self.set_selection)

    def set_datasets(self):
        """Ask for a (n excel) file, then read it and update self.datasets accordingly."""
        path = tk.filedialog.askopenfilename()
        workbook = pyxl.load_workbook(path)

        for worksheet in workbook:
            xs = [cell.value for cell in worksheet['A']]
            ys = [cell.value for cell in worksheet['B']]
            self.datasets[worksheet.title] = [xs, ys]

        self.dropdown.delete(0, tk.END)  # clear
        for key in self.datasets.keys():
            self.dropdown.insert(tk.END, key)

    def set_selection(self, event):
        """Update self.xs, self.ys and draw new points when the user changes data sheet."""
        self.xs, self.ys = self.datasets[self.dropdown.get(self.dropdown.curselection()[0])]
        self.diagram.draw_points(self.xs, self.ys)  # draw points from new sheet

    def _fit(self, regression_func):
        """Draw <class 'Diagram'> based on *regression_func* (from <class 'Regression'>)."""
        equation, equation_str = regression_func(self.xs, self.ys)
        self.diagram.draw_fit(equation, equation_str, self.xs, self.ys)


if __name__ == '__main__':
    root = tk.Tk()
    app = Application(root)
    app.mainloop()
