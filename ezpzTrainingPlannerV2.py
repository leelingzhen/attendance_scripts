import requests
import xlsxwriter
import csv

ATTENDANCE_URL = '' #paste attendance url within parenthesis and ensure that the sheet is made public
PLAYER_PROFILES = 'player_profiles.csv'
SPLIT_GENDER = False
SPLIT_TEAM = True

def export_url_converter(url):
    editIndex = url.index('edit')
    export_url = f'{url[:editIndex]}export?format=csv&{url[editIndex+5:]}'
    return export_url 

def intialise_team_data(file):
    team_dict = {}
    
    if 'http' in file:
        print('Parsing download request for player profiles...')
        with requests.Session() as s:
            file = export_url_converter(file)
            download = s.get(file)
            decoded_content = download.content.decode('utf-8')
            cr = csv.reader(decoded_content.splitlines(), delimiter=',')
            data = list(cr)
            if data[0] == []:
                raise RuntimeError('Failed to open database, please ensure that your sheet sharing options are changed to "Anyone with the link" so that the script can access the sheet')
            team_category = data[1][1]
            team_dict['team categories'] =  {team_category}
    else:
        print(f'Accessing {file} locally')
        with open(file, newline = '') as f:
            data = list(csv.reader(f))
            team_category = data[1][1]
            team_dict['team categories'] =  {team_category}

    for player in data[1:]:
        if player[1] not in team_dict['team categories']:
            team_dict['team categories'].add(player[1])
        team_dict[player[0]]= {
            'team' : player[1],
            'gender': player[2]
        }
    return team_dict


def intialise_attendance_data(url):
    url = export_url_converter(url)
    rejectLst = ['Time','Location','Total', 'Guys', 'Girls','']
    cleanAttendance = {}
    
    print('Parsing download request for attendance sheets...')
    with requests.Session() as s:
        download = s.get(url)
        decoded_content = download.content.decode('utf-8')
        cr = csv.reader(decoded_content.splitlines(), delimiter=',')
        attendance = list(cr)
        if attendance[0] == []:
            raise RuntimeError('Failed to open database, please ensure that your sheet sharing options are changed to "Anyone with the link" so that the script can access the sheet')
        for entry in attendance:
            if entry[0] not in rejectLst:
                cleanAttendance[entry[0]] = entry[1:]
    print('Downloads complete.')
    return cleanAttendance

def training_attendance(data,date):
    attendance = {}
    dateIndex = data["Date"].index(date)
    for entry in data:
        if entry != "Date":
            attendance[entry] =  data[entry][dateIndex]
    return attendance

def training_attendance_sort(data):
    dict_status = {
        'Attending':[],
        'Cmi': [],
        'Injured': [],
        'Late': [],
        'Not Indicated' : [],
        'Invalid Input': []
    }
    for player in data:
        data[player] = data[player].strip() #removes all the spaces that people accidentally input
        if data[player] == '':
            dict_status["Not Indicated"].append((player,data[player]))
        elif '0.5' in data[player]:
            dict_status["Invalid Input"].append((player,data[player]))
        elif data[player][0] == "0":
            dict_status["Cmi"].append((player,data[player]))
        elif data[player][0] == '1':
            data[player].lower()
            if data[player] == '1':
                dict_status['Attending'].append((player,data[player]))
            elif 'inj' in data[player]:
                dict_status['Injured'].append((player,data[player]))
            else:
                dict_status["Late"].append((player,data[player]))
        else:
            dict_status["Invalid Input"].append((player,data[player]))
            
    return dict_status
    
def team_sort(player_lst,d,category):#sort into lists, teamLst for the list to be sorted for, d for dictionary
    if category == 'team':
        categories = d['team categories']
    else:
        categories = {"M","F"}
    init_dic = { c:[] for c in categories}
    for item in player_lst:
        player = item[0]
        if player != 'team categories':
            for c in categories:
                if d.get(player)[category] == c:
                    init_dic[c].append(item)

    
    return init_dic



def clean(attendance): #cleans dictionary to just a name with reasons (if reason exists)
    clean_attendance = {}
    for team in attendance:
        lst = attendance[team]
        clean_lst = []
        for entry in lst:
            entry = f'{entry[0]}{entry[1][1:]}' if len(entry[1])>2 else f'{entry[0]}'
            clean_lst.append(entry)
        clean_attendance[team] = clean_lst
    return clean_attendance

def invalid (attendance):
    clean_attendance = {}
    for team in attendance:
        lst = attendance[team]
        clean_lst = []
        for entry in lst:
            entry = f'{entry[0]} \'{entry[1]}\''
            clean_lst.append(entry)
        clean_attendance[team] = clean_lst
    return clean_attendance


def main():
    player_profiles = intialise_team_data(PLAYER_PROFILES)
    attendance_sheet = intialise_attendance_data(ATTENDANCE_URL)
    attendance_date = input('key in training date in the format d/m/yy:')
    print('Searching for date...')
    input_error_count = 0
    while attendance_date not in attendance_sheet["Date"]:
        input_error_count += 1
        print ("Training date not found.")
        print ('Please check if the input format is correct or if the date exists.')
        print('the input date must be exactly the same as the date on the attendance sheet')
        if input_error_count > 3:
             attendance_date = input('try keying in training date in the format d/m/yyyy:')
        else :
            attendance_date = input('key in training date in the format d/m/yy:')
    print(f'Training date found. selecting {attendance_date}')
    attendance_for_date = training_attendance(attendance_sheet,attendance_date)
    attendance_for_date = training_attendance_sort(attendance_for_date)

    sorted_attendance = {}
    absentees = {}
    if SPLIT_TEAM:
        sorted_by_team = team_sort(attendance_for_date['Attending'] + attendance_for_date["Late"] + attendance_for_date["Injured"],player_profiles,'team')
        absent_sorted_by_team = team_sort(attendance_for_date['Cmi'],player_profiles,'team')
        if SPLIT_GENDER:
            for team in sorted_by_team:
                label = f'Team {team}'
                sorted_by_gender = team_sort(sorted_by_team[team],player_profiles,'gender')
                for gender in sorted_by_gender:
                    sorted_attendance[label + f' ({gender}): {sorted_by_gender[gender]}'] = sorted_by_gender[gender]
            for absentTeam in absent_sorted_by_team:
                label = f'Absent from Team {absentTeam} traning'
                absentees[label+ f': {len(absent_sorted_by_team[absentTeam])}'] = absent_sorted_by_team[absentTeam]
        else:
            for team in sorted_by_team:
                label = f'Team {team}'
                sorted_attendance[label + f': {len(sorted_by_team[team])}'] = sorted_by_team[team]
            for absentTeam in absent_sorted_by_team:
                label = f'Absent from Team {absentTeam} traning'
                absentees[label + f': {len(absent_sorted_by_team[absentTeam])}'] = absent_sorted_by_team[absentTeam]
    elif SPLIT_GENDER:
        sorted_by_gender = team_sort(attendance_for_date['Attending'] + attendance_for_date["Late"] + attendance_for_date["Injured"],player_profiles,'gender')
        for gender in sorted_by_gender:
            sorted_attendance[gender + f': {len(sorted_by_gender[gender])}'] = sorted_by_gender[gender]
        absentees[f'Absent from training: {len(attendance_for_date["Cmi"])}'] = attendance_for_date["Cmi"]
            
    else:
        sorted_attendance[f'Attending: {len(attendance_for_date["Attending"] + attendance_for_date["Late"] + attendance_for_date["Injured"])}'] = attendance_for_date['Attending'] + attendance_for_date["Late"] + attendance_for_date["Injured"]
        absentees[f'Absent from training: {len(attendance_for_date["Cmi"])}'] = attendance_for_date["Cmi"]

    print('sorting complete')
    output = {}
    for i in [clean(sorted_attendance) , clean(absentees) , clean({f'Not Indicated: {len(attendance_for_date["Not Indicated"])}': attendance_for_date["Not Indicated"]}) , invalid({f'Invalid Inputs: {len(attendance_for_date["Invalid Input"])}' : attendance_for_date["Invalid Input"]})]:
        output.update(i)

    
    print('generating sheets')
    new_date = attendance_date.replace("/",'.')
    workbook = xlsxwriter.Workbook(f'Training_Attendance_{new_date}.xlsx')
    worksheet = workbook.add_worksheet()
    #writing headers
    row,col = 0,0
    for header in output:
        worksheet.write(row,col, header)
        col += 1
    #writing entire dataframe
    col =0
    for arr in output:
        row = 1
        for item in output[arr]:
            worksheet.write(row,col,(' '.join(str(element) for element in item)) if type(item) == list else (''.join(str(element) for element in item)) )
            row += 1
        col += 1
    workbook.close()
    print('sheets have been generated')



if __name__ == "__main__":
    main()