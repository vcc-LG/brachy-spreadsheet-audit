"""
A tool to gather date from the Brachytherapy Vienna spreadsheet, insert it into a database, and run some
basic queries and visualisations. For research purposes only.
"""

from pymongo import MongoClient
from bokeh.plotting import figure, output_file, show
from bokeh.models import Span, Label, DatetimeTickFormatter
import numpy as np
import re
import selenium.webdriver
from datetime import datetime
import matplotlib.pyplot as plt


def get_quantity(quantity_name):
    """fetches a given quantity from the collection"""
    data_list = []
    db_string = 'insertions.'+quantity_name
    for patient in db.patients.find({}, {db_string: 1, '_id': 0}):
        for insertion in patient['insertions']:
            try:
                if insertion[quantity_name]:
                    data_list.append(insertion[quantity_name])
            except KeyError:
                pass
    data_list_clean = []
    for el in data_list:
        if isinstance(el, str):
            try:
                data_list_clean.append(float(re.findall("\d+\.\d+", el)[0]))
            except IndexError:
                pass
        else:
            data_list_clean.append(el)
    return data_list_clean


def run_query(query):
    """fetches a given quantity from collection given two requirements"""
    for patient in db.patients.find({}, {db_string1: 1, db_string2: 1}):
        for insertion in patient['insertions']:
            try:
                if insertion[quantity_name]:
                    data_list.append(insertion[quantity_name])
            except KeyError:
                pass
    data_list_clean = []
    for el in data_list:
        if isinstance(el, str):
            try:
                data_list_clean.append(float(re.findall("\d+\.\d+", el)[0]))
            except IndexError:
                pass
        else:
            data_list_clean.append(el)
    return data_list_clean

#connect to mongodb
client = MongoClient()
db = client.patient_database
patients_data = db.patients

quantity = 'mean_point_a'
data_out = get_quantity(quantity)
output_file("output.html")
TOOLS = 'box_zoom,box_select,resize,reset'

p = figure(plot_width=1200, plot_height=800,
           title='IGBT audit: '+quantity,
           x_axis_label='patient #',
           y_axis_label='Mean Point A dose (Gy)',
           title_text_font_size='40pt',
           tools=TOOLS)

p.xaxis.axis_label_text_font_size = "32pt"
p.yaxis.axis_label_text_font_size = "32pt"

p.xaxis.major_label_text_font_size = "24pt"
p.yaxis.major_label_text_font_size = "24pt"

# add a line renderer with legend and line thickness
p.circle(range(len(data_out)), data_out, fill_color="red", line_color="red", size=6)
hline = Span(location=np.mean(data_out), dimension='width', line_color='green', line_width=1)
p.renderers.extend([hline])

hline = Span(location=np.mean(data_out)+np.std(data_out), dimension='width', line_color='blue', line_width=1,line_dash='dashed')
p.renderers.extend([hline])

hline = Span(location=np.mean(data_out)-np.std(data_out), dimension='width', line_color='blue', line_width=1,line_dash='dashed')
p.renderers.extend([hline])

mean_str = 'Mean = '+str(round(np.mean(data_out),2))
citation = Label(x=260, y=np.mean(data_out),
                 text=mean_str, render_mode='css',text_color = 'green')
p.add_layout(citation)
# show the results
show(p)

driver = selenium.webdriver.PhantomJS(executable_path=r'C:\Users\le165208\Apps\PhantomJS\phantomjs-2.1.1-windows\bin\phantomjs')
driver.get('file:///C:/Users/le165208/githubprojects/brachy_dose_audit/output.html')
save_str = 'screens/'+quantity+'.png'
driver.save_screenshot(save_str)

print(np.mean(data_out))
print(np.std(data_out))

date_list = []
d90_list = []
A = db.patients.find({'insertions.hr_ctv_d90_gy':{'$exists': True}},{'insertions.hr_ctv_d90_gy':1, 'insertions.insertion_date':1})
for patient in A:
    for insertion in patient['insertions']:
        try:
            d90_list.append(float(insertion['hr_ctv_d90_gy']))
            date_list.append(datetime.strptime(insertion['insertion_date'], '%Y-%m-%d'))
        except:
            pass

output_file("output.html")
TOOLS = 'box_zoom,box_select,resize,reset'
p = figure(plot_width=1200, plot_height=800,
           title='IGBT audit: HR-CTV D90',
           x_axis_label='Date',
           y_axis_label='HR-CTV D90 (Gy)',
           x_axis_type="datetime",
           title_text_font_size='40pt',
           tools=TOOLS)
p.xaxis.formatter = DatetimeTickFormatter(
    formats=dict(
        months=["%m/%Y"],
        years=["%m/%Y"],
    )
)
p.xaxis.axis_label_text_font_size = "32pt"
p.yaxis.axis_label_text_font_size = "32pt"
p.xaxis.major_label_text_font_size = "20pt"
p.yaxis.major_label_text_font_size = "24pt"
p.circle(date_list, d90_list, fill_color="red", line_color="red", size=6)
hline = Span(location=np.mean(d90_list), dimension='width', line_color='green', line_width=1)
p.renderers.extend([hline])
hline = Span(location=np.mean(d90_list) + np.std(d90_list), dimension='width', line_color='blue', line_width=1,
             line_dash='dashed')
p.renderers.extend([hline])
hline = Span(location=np.mean(d90_list) - np.std(d90_list), dimension='width', line_color='blue', line_width=1,
             line_dash='dashed')
p.renderers.extend([hline])
mean_str = 'Mean = ' + str(round(np.mean(d90_list), 2))
citation = Label(x=260, y=np.mean(d90_list),
                 text=mean_str, render_mode='css', text_color='green')
p.add_layout(citation)
show(p)
driver = selenium.webdriver.PhantomJS(
    executable_path=r'C:\Users\le165208\Apps\PhantomJS\phantomjs-2.1.1-windows\bin\phantomjs')
driver.get('file:///C:/Users/le165208/githubprojects/brachy_dose_audit/output.html')
save_str = 'screens/' + 'hr_ctv_d90' + '.png'
driver.save_screenshot(save_str)









volume_list = []
d90_list = []
A = db.patients.find({'insertions.hr_ctv_d90_gy':{'$exists': True}},{'insertions.hr_ctv_d90_gy':1, 'insertions.hr_ctv_volume_cm3':1})
for patient in A:
    for insertion in patient['insertions']:
        try:
            d90_list.append(float(insertion['hr_ctv_d90_gy']))
            volume_list.append(float(insertion['hr_ctv_volume_cm3']))
        except:
            pass

output_file("output.html")
TOOLS = 'box_zoom,box_select,resize,reset'
p = figure(plot_width=1200, plot_height=800,
           title='IGBT audit: HR-CTV D90',
           x_axis_label='HR-CTV volume (cm3)',
           y_axis_label='HR-CTV D90 (Gy)',
           # x_axis_type="datetime",
           title_text_font_size='40pt',
           tools=TOOLS)
p.xaxis.axis_label_text_font_size = "32pt"
p.yaxis.axis_label_text_font_size = "32pt"
p.xaxis.major_label_text_font_size = "24pt"
p.yaxis.major_label_text_font_size = "24pt"
p.circle(volume_list, d90_list, fill_color="red", line_color="red", size=6)
hline = Span(location=np.mean(d90_list), dimension='width', line_color='green', line_width=1)
p.renderers.extend([hline])
hline = Span(location=np.mean(d90_list) + np.std(d90_list), dimension='width', line_color='blue', line_width=1,
             line_dash='dashed')
p.renderers.extend([hline])
hline = Span(location=np.mean(d90_list) - np.std(d90_list), dimension='width', line_color='blue', line_width=1,
             line_dash='dashed')
p.renderers.extend([hline])
mean_str = 'Mean = ' + str(round(np.mean(d90_list), 2))
citation = Label(x=260, y=np.mean(d90_list),
                 text=mean_str, render_mode='css', text_color='green')
p.add_layout(citation)
show(p)
driver = selenium.webdriver.PhantomJS(
    executable_path=r'C:\Users\le165208\Apps\PhantomJS\phantomjs-2.1.1-windows\bin\phantomjs')
driver.get('file:///C:/Users/le165208/githubprojects/brachy_dose_audit/output.html')
save_str = 'screens/' + 'hr_ctv_d90_volume2' + '.png'
driver.save_screenshot(save_str)












client = MongoClient()
db = client.patient_database
volume_list = []
bladder_2cc_list = []
A = db.patients.find({'insertions.bladder_d2cc_gy':{'$exists': True}},{'insertions.bladder_d2cc_gy':1, 'insertions.bladder_volume_cm3':1})
for patient in A:
    for insertion in patient['insertions']:
        try:
            bladder_2cc_list.append(float(insertion['bladder_d2cc_gy']))
            volume_list.append(float(insertion['bladder_volume_cm3']))
        except:
            pass

output_file("output.html")
TOOLS = 'box_zoom,box_select,resize,reset'
p = figure(plot_width=1200, plot_height=800,
           title='IGBT audit: Bladder D2cc',
           x_axis_label='Bladder volume (cm3)',
           y_axis_label='Bladder D2cc (Gy)',
           # x_axis_type="datetime",
           title_text_font_size='40pt',
           tools=TOOLS)
p.xaxis.axis_label_text_font_size = "32pt"
p.yaxis.axis_label_text_font_size = "32pt"
p.xaxis.major_label_text_font_size = "24pt"
p.yaxis.major_label_text_font_size = "24pt"
p.circle(volume_list, bladder_2cc_list, fill_color="red", line_color="red", size=6)
hline = Span(location=np.mean(bladder_2cc_list), dimension='width', line_color='green', line_width=1)
p.renderers.extend([hline])
hline = Span(location=np.mean(bladder_2cc_list) + np.std(bladder_2cc_list), dimension='width', line_color='blue', line_width=1,
             line_dash='dashed')
p.renderers.extend([hline])
hline = Span(location=np.mean(bladder_2cc_list) - np.std(bladder_2cc_list), dimension='width', line_color='blue', line_width=1,
             line_dash='dashed')
p.renderers.extend([hline])
mean_str = 'Mean = ' + str(round(np.mean(bladder_2cc_list), 2))
citation = Label(x=260, y=np.mean(bladder_2cc_list),
                 text=mean_str, render_mode='css', text_color='green')
p.add_layout(citation)
show(p)
driver = selenium.webdriver.PhantomJS(
    executable_path=r'C:\Users\le165208\Apps\PhantomJS\phantomjs-2.1.1-windows\bin\phantomjs')
driver.get('file:///C:/Users/le165208/githubprojects/brachy_dose_audit/output.html')
save_str = 'screens/' + 'bladder_volume_2cc' + '.png'
driver.save_screenshot(save_str)








client = MongoClient()
db = client.patient_database
bowel_d2cc_list = []
rectum_d2cc_list = []
A = db.patients.find({'insertions.rectum_d2cc_gy':{'$exists': True}},{'insertions.rectum_d2cc_gy':1, 'insertions.bowel_d2cc_gy':1})
for patient in A:
    for insertion in patient['insertions']:
        try:
            rectum_d2cc_list.append(float(insertion['rectum_d2cc_gy']))
            bowel_d2cc_list.append(float(insertion['bowel_d2cc_gy']))
        except:
            pass

output_file("output.html")
TOOLS = 'box_zoom,box_select,resize,reset'
p = figure(plot_width=1200, plot_height=800,
           title='IGBT audit: bowel D2cc',
           x_axis_label='bowel D2cc (Gy)',
           y_axis_label='rectum D2cc (Gy)',
           # x_axis_type="datetime",
           title_text_font_size='40pt',
           tools=TOOLS)
p.xaxis.axis_label_text_font_size = "32pt"
p.yaxis.axis_label_text_font_size = "32pt"
p.xaxis.major_label_text_font_size = "24pt"
p.yaxis.major_label_text_font_size = "24pt"
p.circle(bowel_d2cc_list, rectum_d2cc_list, fill_color="red", line_color="red", size=6)
hline = Span(location=np.mean(rectum_d2cc_list), dimension='width', line_color='green', line_width=1)
p.renderers.extend([hline])
hline = Span(location=np.mean(rectum_d2cc_list) + np.std(rectum_d2cc_list), dimension='width', line_color='blue', line_width=1,
             line_dash='dashed')
p.renderers.extend([hline])
hline = Span(location=np.mean(rectum_d2cc_list) - np.std(rectum_d2cc_list), dimension='width', line_color='blue', line_width=1,
             line_dash='dashed')
p.renderers.extend([hline])
mean_str = 'Mean = ' + str(round(np.mean(rectum_d2cc_list), 2))
citation = Label(x=260, y=np.mean(rectum_d2cc_list),
                 text=mean_str, render_mode='css', text_color='green')
p.add_layout(citation)
show(p)
driver = selenium.webdriver.PhantomJS(
    executable_path=r'C:\Users\le165208\Apps\PhantomJS\phantomjs-2.1.1-windows\bin\phantomjs')
driver.get('file:///C:/Users/le165208/githubprojects/brachy_dose_audit/output.html')
save_str = 'screens/' + 'bowelvshrctv' + '.png'
driver.save_screenshot(save_str)















client = MongoClient()
db = client.patient_database
insertion_num_list = []
hr_ctv_d90_list = []
A = db.patients.find({'insertions.hr_ctv_d90_gy':{'$exists': True}},{'insertions.hr_ctv_d90_gy':1, 'insertions.insertion_number':1})
for patient in A:
    for insertion in patient['insertions']:
        try:
            hr_ctv_d90_list.append(float(insertion['hr_ctv_d90_gy']))
            insertion_num_list.append(int(insertion['insertion_number']))
        except:
            pass

output_file("output.html")
TOOLS = 'box_zoom,box_select,resize,reset'
p = figure(plot_width=1200, plot_height=800,
           title='IGBT audit: HR-CTV D90',
           x_axis_label='Insertion number',
           y_axis_label='HR-CTV D90 (Gy)',
           # x_axis_type="datetime",
           title_text_font_size='40pt',
           tools=TOOLS)
p.xaxis.axis_label_text_font_size = "32pt"
p.yaxis.axis_label_text_font_size = "32pt"
p.xaxis.major_label_text_font_size = "24pt"
p.yaxis.major_label_text_font_size = "24pt"
p.circle(insertion_num_list, hr_ctv_d90_list, fill_color="red", line_color="red", size=6)
hline = Span(location=np.mean(hr_ctv_d90_list), dimension='width', line_color='green', line_width=1)
p.renderers.extend([hline])
hline = Span(location=np.mean(hr_ctv_d90_list) + np.std(hr_ctv_d90_list), dimension='width', line_color='blue', line_width=1,
             line_dash='dashed')
p.renderers.extend([hline])
hline = Span(location=np.mean(hr_ctv_d90_list) - np.std(hr_ctv_d90_list), dimension='width', line_color='blue', line_width=1,
             line_dash='dashed')
p.xaxis[0].ticker=FixedTicker(ticks=[1, 2, 3])
p.renderers.extend([hline])
mean_str = 'Mean = ' + str(round(np.mean(hr_ctv_d90_list), 2))
citation = Label(x=260, y=np.mean(hr_ctv_d90_list),
                 text=mean_str, render_mode='css', text_color='green')
p.add_layout(citation)
show(p)
driver = selenium.webdriver.PhantomJS(
    executable_path=r'C:\Users\le165208\Apps\PhantomJS\phantomjs-2.1.1-windows\bin\phantomjs')
driver.get('file:///C:/Users/le165208/githubprojects/brachy_dose_audit/output.html')
save_str = 'screens/' + 'hr_ctv_insertion' + '.png'
driver.save_screenshot(save_str)


from bokeh.charts import BoxPlot
import pandas as pd
data = dict(insertion_number = insertion_num_list, hr_ctv_d90 = hr_ctv_d90_list)

data_to_plot = []
ins1 = [hr_ctv_d90_list[j] for j in [i for i in range(len(insertion_num_list)) if insertion_num_list[i]==1]]
ins2 = [hr_ctv_d90_list[j] for j in [i for i in range(len(insertion_num_list)) if insertion_num_list[i]==2]]
ins3 = [hr_ctv_d90_list[j] for j in [i for i in range(len(insertion_num_list)) if insertion_num_list[i]==3]]

data_to_plot = [ins1,ins2,ins3]
fig, ax1 = plt.subplots(figsize=(15, 10))
ax1.set_title('HR-CTV D90 (Gy) vs. insertion number',
              fontsize = 30)
fig.canvas.set_window_title('HR-CTV D90 (Gy)')
bp = plt.boxplot(data_to_plot, widths = 0.2,notch=0, sym='+', vert=1, whis=1.5,patch_artist=True)
xtickNames = plt.setp(ax1,xticklabels= ['Insertion 1', 'Insertion 2', 'Insertion 3'])
ax1.set_ylabel("HR-CTV D90 (Gy)", fontsize=26)
for tick in ax1.yaxis.get_major_ticks():
                tick.label.set_fontsize(24)
plt.setp(xtickNames, fontsize=26)
plt.setp(bp['boxes'],            # customise box appearance
         color='grey',       # outline colour
         linewidth=1.5,             # outline line width
         facecolor='SkyBlue')       # fill box with colour

plt.setp(bp['whiskers'],         # customise whisker appearence
         color='DarkMagenta',       # whisker colour
         linewidth=1.5)             # whisker thickness

plt.setp(bp['caps'],             # customize lines at the end of whiskers 
         color='DarkMagenta',       # cap colour
         linewidth=1.5)             # cap thickness

plt.setp(bp['fliers'],           # customize marks for extreme values
         color='Tomato',            # set mark colour
         marker='o',                # maker shape
         markersize=10)             # marker size

plt.setp(bp['medians'],          # customize median lines
         color='Tomato',            # line colour
         linewidth=1.5)             # line thickness
plt.show()






client = MongoClient()
db = client.patient_database
insertion_num_list = []
hr_ctv_volume_list = []
A = db.patients.find({'insertions.hr_ctv_volume_cm3':{'$exists': True}},{'insertions.hr_ctv_volume_cm3':1, 'insertions.insertion_number':1})
for patient in A:
    for insertion in patient['insertions']:
        try:
            hr_ctv_volume_list.append(float(insertion['hr_ctv_volume_cm3']))
            insertion_num_list.append(int(insertion['insertion_number']))
        except:
            pass
        
data_to_plot = []
ins1 = [hr_ctv_volume_list[j] for j in [i for i in range(len(insertion_num_list)) if insertion_num_list[i]==1]]
ins2 = [hr_ctv_volume_list[j] for j in [i for i in range(len(insertion_num_list)) if insertion_num_list[i]==2]]
ins3 = [hr_ctv_volume_list[j] for j in [i for i in range(len(insertion_num_list)) if insertion_num_list[i]==3]]

data_to_plot = [ins1,ins2,ins3]
fig, ax1 = plt.subplots(figsize=(15, 10))
ax1.set_title('HR-CTV volume (cm3) vs. insertion number',
              fontsize = 30)
fig.canvas.set_window_title('HR-CTV volume (cm3))')
bp = plt.boxplot(data_to_plot, widths = 0.2,notch=0, sym='+', vert=1, whis=1.5,patch_artist=True)
xtickNames = plt.setp(ax1,xticklabels= ['Insertion 1', 'Insertion 2', 'Insertion 3'])
ax1.set_ylabel("HR-CTV volume (cm3)", fontsize=26)
for tick in ax1.yaxis.get_major_ticks():
                tick.label.set_fontsize(24)
plt.setp(xtickNames, fontsize=26)
plt.setp(bp['boxes'],            # customise box appearance
         color='grey',       # outline colour
         linewidth=1.5,             # outline line width
         facecolor='SkyBlue')       # fill box with colour

plt.setp(bp['whiskers'],         # customise whisker appearence
         color='DarkMagenta',       # whisker colour
         linewidth=1.5)             # whisker thickness

plt.setp(bp['caps'],             # customize lines at the end of whiskers 
         color='DarkMagenta',       # cap colour
         linewidth=1.5)             # cap thickness

plt.setp(bp['fliers'],           # customize marks for extreme values
         color='Tomato',            # set mark colour
         marker='o',                # maker shape
         markersize=10)             # marker size

plt.setp(bp['medians'],          # customize median lines
         color='Tomato',            # line colour
         linewidth=1.5)             # line thickness
plt.show()



output_file("output.html")
TOOLS = 'box_zoom,box_select,resize,reset'
p = figure(plot_width=1200, plot_height=800,
           title='IGBT audit: HR-CTV D90',
           x_axis_label='Insertion number',
           y_axis_label='HR-CTV D90 (Gy)',
           # x_axis_type="datetime",
           title_text_font_size='40pt',
           tools=TOOLS)
p.xaxis.axis_label_text_font_size = "32pt"
p.yaxis.axis_label_text_font_size = "32pt"
p.xaxis.major_label_text_font_size = "24pt"
p.yaxis.major_label_text_font_size = "24pt"
p.circle(insertion_num_list, hr_ctv_volume_list, fill_color="red", line_color="red", size=6)
hline = Span(location=np.mean(hr_ctv_volume_list), dimension='width', line_color='green', line_width=1)
p.renderers.extend([hline])
hline = Span(location=np.mean(hr_ctv_volume_list) + np.std(hr_ctv_volume_list), dimension='width', line_color='blue', line_width=1,
             line_dash='dashed')
p.renderers.extend([hline])
hline = Span(location=np.mean(hr_ctv_volume_list) - np.std(hr_ctv_volume_list), dimension='width', line_color='blue', line_width=1,
             line_dash='dashed')
p.xaxis[0].ticker=FixedTicker(ticks=[1, 2, 3])
p.renderers.extend([hline])
mean_str = 'Mean = ' + str(round(np.mean(hr_ctv_volume_list), 2))
citation = Label(x=260, y=np.mean(hr_ctv_volume_list),
                 text=mean_str, render_mode='css', text_color='green')
p.add_layout(citation)
show(p)
