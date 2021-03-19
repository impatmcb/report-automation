with open('I:/CrosstabData.csv') as csvfile:
    originalinfo = csv.reader((x.replace('\0', '').replace(',', ' ').replace('\t', ',') for x in csvfile), delimiter=',')
    next(originalinfo)
    myfile = open(I:/CrosstabOutputTemp.csv", "w", newline='')
    writer = csv.writer(myfile)
    for row in originalinfo:
        if len(row) > 0:
            if row[4] == 'Total':
                pass
            else:
                try:
                    classname = row[2]
                    date = row[0].replace('  ', ', ')
                    dayofweek = datetime.datetime.strptime(date, '%B %d, %Y').strftime('%A')
                    site = row[1].split(" ")[0]
                    enrolled = int(float(row[-1]))
                    starttime = row[3].lstrip()
                    endtime = row[4].lstrip()
                    data = [dayofweek, date, site, classname, starttime, endtime, enrolled]
                    writer.writerow(data)
                except ValueError:
                    print("Couldn't process a line")
    myfile.close()
