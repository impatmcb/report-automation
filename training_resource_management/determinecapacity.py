monenrolled = {}
tuesenrolled = {}
wedsenrolled = {}
thurenrolled = {}
frienrolled = {}


def enrolleddict(day, dict):
    for value in day:
        if value[2] in dict:
            dict[value[2]] += int(value[6])
        else:
            dict[value[2]] = int(value[6])


enrolleddict(monday, monenrolled)
enrolleddict(tuesday, tuesenrolled)
enrolleddict(wednesday, wedsenrolled)
enrolleddict(thursday, thurenrolled)
enrolleddict(friday, frienrolled)

# Build dictionaries for capacity
primaryrooms = {}
secondaryrooms = {}
thirdrooms = {}
roomcapacity = {}

with open('I:/mastercap.csv') as csvfile:
    reader = csv.reader(csvfile)
    for line in reader:
        if line[0] in primaryrooms:
            if line[0] in secondaryrooms:
                thirdrooms[line[0]] = line[1]
                roomcapacity[line[1]] = int(float(line[2]))
            else:
                secondaryrooms[line[0]] = line[1]
                roomcapacity[line[1]] = int(float(line[2]))
        else:
            primaryrooms[line[0]] = line[1]
            roomcapacity[line[1]] = int(float(line[2]))
