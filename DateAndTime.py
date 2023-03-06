import datetime

e = datetime.datetime.now()

#Morning or Afternoon Variable Determination Logic
if e.strftime("%p") == "AM": 
    Time = "Morning"
elif e.strftime("%p") == "PM":
    Time = "Afternoon"
