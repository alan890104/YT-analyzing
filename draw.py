import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import datetime

date = []
views = []
with open('data.txt','r') as f:
    for line in f.readlines():
        x = line.split()
        date.append(datetime.datetime.strptime(x[0],"%Y-%m-%d"))
        views.append(x[1])

ax = plt.gca()
formatter = mdates.DateFormatter("%Y-%m-%d")
ax.xaxis.set_major_formatter(formatter)
locator = mdates.DayLocator()
ax.xaxis.set_major_locator(locator)
plt.plot(date, views)
now_time = datetime.datetime.now().strftime("%Y-%m-%d")
plt.savefig(str(now_time)+'.png')
plt.show()
