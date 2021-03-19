# Determine the date of the class
def classdate(number):
    return (datetime.date.today() + datetime.timedelta(days=(number-datetime.date.today().weekday()))).strftime("%m/%d")


nextmon, nexttue, nextwed, nextthu, nextfri = classdate(7), classdate(8), classdate(9), classdate(10), classdate(11)


# Get coaches for each day based on priority
primarycoachmon = {}
secondarycoachmon = {}
thirdcoachmon = {}
primarycoachtues = {}
secondarycoachtues = {}
thirdcoachtues = {}
primarycoachwed = {}
secondarycoachwed = {}
thirdcoachwed = {}
primarycoachthur = {}
secondarycoachthur = {}
thirdcoachthur = {}
primarycoachfri= {}
secondarycoachfri = {}
thirdcoachfri = {}

with open('I:/signupsheet.csv') as csvfile:
    reader = csv.reader(csvfile)
    for row in reader:
        datelist = row[2].split("/")
        dateofclass = datetime.datetime.strptime(f"{datelist[0]}/{datelist[1]}", "%m/%d")
        date = dateofclass.strftime("%m/%d")
        if date == nextmon:
            if row[6] in primarycoachmon:
                if row[6] in secondarycoachmon:
                    thirdcoachmon[row[6]] = row[7].title()
                else:
                    secondarycoachmon[row[6]] = row[7].title()
            else:
                primarycoachmon[row[6]] = row[7].title()
        if date == nexttue:
            if row[6] in primarycoachtues:
                if row[6] in secondarycoachtues:
                    thirdcoachtues[row[6]] = row[7].title()
                else:
                    secondarycoachtues[row[6]] = row[7].title()
            else:
                primarycoachtues[row[6]] = row[7].title()
        if date == nextwed:
            if row[6] in primarycoachwed:
                if row[6] in secondarycoachwed:
                    thirdcoachwed[row[6]] = row[7].title()
                else:
                    secondarycoachwed[row[6]] = row[7].title()
            else:
                primarycoachwed[row[6]] = row[7].title()
        if date == nextthu:
            if row[6] in primarycoachthur:
                if row[6] in secondarycoachthur:
                    thirdcoachthur[row[6]] = row[7].title()
                else:
                    secondarycoachthur[row[6]] = row[7].title()
            else:
                primarycoachthur[row[6]] = row[7].title()
        if date == nextfri:
            if row[6] in primarycoachfri:
                if row[6] in secondarycoachfri:
                    thirdcoachfri[row[6]] = row[7].title()
                else:
                    secondarycoachfri[row[6]] = row[7].title()
            else:
                primarycoachfri[row[6]] = row[7].title()
