monday, tuesday, wednesday, thursday, friday = [], [], [], [], []
crosstabinfo = [tuple(row) for row in csv.reader(open('c:/users/mcbridps/desktop/Crosstaboutputtemp.csv', 'rU'))]
for item in crosstabinfo:
    if item[0] == 'Monday':
        monday.append(item)
    elif item[0] == 'Tuesday':
        tuesday.append(item)
    elif item[0] == 'Wednesday':
        wednesday.append(item)
    elif item[0] == 'Thursday':
        thursday.append(item)
    elif item[0] == 'Friday':
        friday.append(item)
