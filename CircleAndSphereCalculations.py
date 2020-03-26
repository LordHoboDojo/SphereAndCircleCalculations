import numpy as np
import xlsxwriter as excel_writer

layer_thickness = 1.0
dwell_time = .25
velocity = 5.0


def calculate_circle_time(diameter):
    if diameter == 0:
        return 0
    res = 0.0
    range_var = int(diameter / 2.0 / layer_thickness) + 1
    for i in range(0, range_var, 1):
        arcsin_parameter = (i * layer_thickness) / (diameter / 2.0)
        sin_parameter = (np.pi - 2.0 * np.arcsin(arcsin_parameter)) / 2.0
        res += diameter * np.sin(sin_parameter) + layer_thickness
    return res / velocity * 2 + diameter * dwell_time / layer_thickness


def calculate_sphere_time(diameter):
    res = 0.0
    range_var = diameter / 2.0 / layer_thickness + 1
    for i in range(0, int(range_var), 1):
        arcsin_parameter = i * layer_thickness / (diameter / 2.0)
        sin_parameter = (np.pi - 2.0 * np.arcsin(arcsin_parameter)) / 2.0
        res += calculate_circle_time(diameter * np.sin(sin_parameter)) + layer_thickness / velocity

    return res * 2 + dwell_time * diameter / layer_thickness


data = excel_writer.Workbook('data.xlsx')
data_worksheet = data.add_worksheet()
circle_data = []
sphere_data = []
radii_list = range(10, 101, 10)
for x in range(10, 101, 10):
    circle_data.append(calculate_circle_time(x))
    sphere_data.append(calculate_sphere_time(x))
data_worksheet.write(0, 0, 'Diameter -Circle')
data_worksheet.write(0, 1, 'Time(s) -Circle')
data_worksheet.write(0, 2, 'Time(s) -Sphere')
data_worksheet.write_column(1, 0, radii_list)
data_worksheet.write_column(1, 1, circle_data)
data_worksheet.write_column(1, 2, sphere_data)
data_worksheet.set_column(0, 2, 25)
data.close()
