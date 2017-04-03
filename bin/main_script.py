import os, csv, sys, operator, math
from datetime import datetime

try:
    import xlrd
except:
    print "xlrd library must be installed to run this script"
    print "This library can eaisly be installed through pip by running the following command: pip install xlrd"

#.....................................................
# SETTING UP THE DIRECTORY'S
#.....................................................
INPUT_PATH = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(os.path.normpath(INPUT_PATH))
INPUT_DIR = os.path.join(BASE_DIR, "Input Data")
OUTPUT_DIR = os.path.join(BASE_DIR, "Output Data")
INPUT_PARAMETERS_WORKBOOK = os.path.join(INPUT_DIR, "input parameter.xlsm")
INPUT_PARAMETERS_FILE = os.path.join(INPUT_DIR, "input parameter.csv")
INPUT_CSV_FILE = os.path.join(INPUT_DIR, "input.csv")

COLS = [
    'Location ID',
    'Field Sample ID',
    'Start Depth',
    'End Depth',
    'Parameter Name',
    'Report Result',
    'Leached',
    'Detected',
    'Filtered',
    'Ground Elevation',
    'Location Group Name'
]
	# Fields that are not needed
    #'Lab Qualifier',
	#'Validation Qualifier',
	#'Sample Purpose',

#.....................................................
# Definitions
#.....................................................

def get_filters() :
    f = {
        'Convert BGS to Elevation':-1,
        'Interval Range':-1,
        'Max/Min Elevation Filter':-1,
        'Max/Min Depth Interval Filter':-1,
        'Max Elevation':-1,
        'Min Elevation':-1,
        'Max Depth':-1,
        'Min Depth':-1, 
        'Detect Values':-1,
        'Non-Detect Values':-1,
        'Filtered Results':-1,
        'Non-Filtered Results':-1,
        'Leached Results':-1,
        'Non-Leeched Results':-1}

    YesBoolean = {"YES", "yes", "y" ,"Y", "Yes"}
    NoBoolean = {"NO", "no", "n", "N", "No"}

    INPUT_PATH = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = os.path.dirname(os.path.normpath(INPUT_PATH))
    INPUT_DIR = os.path.join(BASE_DIR, "Input Data")
    INPUT_PARAMETERS_FILE = os.path.join(INPUT_DIR, "input parameter.csv")

    csvfile = open(INPUT_PARAMETERS_FILE, 'Ur')
    reader = csv.reader(csvfile)#, dialect='excel',delimiter = ',')
    
    # Create Dictionary from file
    FILE_FILTERS=dict()
    for row in reader:
        key = row[0]
        FILE_FILTERS[key] = row[1]
    
    # Compare main dictionary (its called f; set within program) and the file dictionary (its called FILE_FILTERS; created from items in input file)
    # and update values in the main dictionary.
    for key1,item1 in f.items():
        for key2, item2 in FILE_FILTERS.items():
            if key1 == key2:
                f[key1] = item2
            
    for key,item in f.items():
        if f[key] in YesBoolean:
            f[key] = True
        elif f[key] in NoBoolean:
            f[key] = False
        elif key == 'Max Depth' and item == '':
            f[key] = float(999)
        elif key == 'Min Depth' and item == '':
            f[key] = float(0)
        elif key == 'Max Elevation' and item == '':
            f[key] = 'null'
        elif key == 'Min Elevation' and item == '':
            f[key] = float(-999)
        else:
            f[key] = float(item)

    for key1,item1 in f.items():
        if item1 == -1:
            print "Found no matches for {}".format(key1)
            sys.exit()

    csvfile.close()
    return f

def add_data_elevation( listlist, data, filter ):
    start_el = float(data[index['Ground Elevation']]) - float(data[index['Start Depth']])
    end_el = float(data[index['Ground Elevation']]) - float(data[index['End Depth']])
    last_list = listlist.pop()
    if last_list[0] == data[index['Location ID']]:
        ls = last_list
    else:
        listlist.append(last_list)
        ls = list()
        ls.append( data[index['Location ID']] )
    if int(float(filter['Max Elevation'] - end_el)/filter['Interval Range']) >= (len(ls)-1)*filter['Interval Range']:#+ 1):
        more_empties = int(round((float(filter['Max Elevation'] - end_el)) / filter['Interval Range'] - len(ls) +1))
        ls.extend(['']*( more_empties + 1))
    value = data[index['Report Result']] + " " + (get_qualifier(data[index['Validation Qualifier']],data[index['Lab Qualifier']]))
    depth = start_el
    #This case solves for results where the start depth is larger than the max elevation but the end depth is lower.  Need to make the start depth elevation the same as the max elevation
    while depth > filter['Max Elevation']:
        print "{} | Start Elevation is larger than the Max Elevation".format(depth)
        depth -= filter['Interval Range']
    i = int((filter['Max Elevation'] - end_el) / filter['Interval Range'] ) + 1
    while depth >= end_el and depth >= filter['Min Elevation']:
        ls.pop(i)
        ls.insert(i,value)
        i -= 1
        depth -= filter['Interval Range']
    listlist.append(ls)
    return listlist


def add_data( listlist, data, filter ):
    start_depth = float(data[index['Start Depth']])
    end_depth = float(data[index['End Depth']])
    last_list = listlist.pop()
    if last_list[0] == data[index['Location ID']]:
        ls = last_list
    else:
        listlist.append(last_list)
        ls = list()
        ls.append( data[index['Location ID']] )
    if end_depth > (len(ls) - 2)*filter['Interval Range']:
        more_empties = int(round(end_depth / filter['Interval Range'] - len(ls) +1))
        ls.extend(['']*( more_empties + 1 ))
    value = data[index['Report Result']] + " " + (get_qualifier(data[index['Validation Qualifier']],data[index['Lab Qualifier']]))
    depth = start_depth
    while depth < filter['Min Depth']:
        depth += filter['Interval Range']
    i = int((depth - filter['Min Depth']) / filter['Interval Range'] ) + 1
    while depth <= end_depth and depth <= filter['Max Depth']:
        ls.pop(i)
        ls.insert(i,value)
        i += 1
        depth += filter['Interval Range']
    listlist.append(ls)
    return listlist

def get_qualifier(validator, lab):
    if validator != "":
        return validator
    elif lab != "":
        return lab
    else:
        return ""

def get_toggle_filters( filters ):
    toggle_filters= [
        ['', '', 'Detected'],
        ['', '', 'Filtered'],
        ['', '', 'Leached']
    ]

    for values in toggle_filters:
        if values[2] == 'Detected':
            values[0] = filters['Detect Values']
            values[1] = filters['Non-Detect Values']
        if values[2] == 'Filtered':
            values[0] = filters['Filtered Results']
            values[1] = filters['Non-Filtered Results']
        if values[2] == 'Leached':
            values[0] = filters['Leached Results']
            values[1] = filters['Non-Leeched Results']

    return toggle_filters

def is_of_interest( row, toggle_filters ):

    f = get_filters()
    #check for empty start and end depths
    if row[index['Start Depth']] == "" and row[index['End Depth']] == "":
        print ("{}:{} for the Constituent, {}, in {} does not have a Start and End Depth".format(row[index['Location ID']],row[index['Field Sample ID']],row[index['Parameter Name']],row[index['Location Group Name']]))
        return False
    if f['Max/Min Elevation Filter'] is True:
        start_el = float(row[index['Ground Elevation']]) - float(row[index['Start Depth']])
        end_el = float(row[index['Ground Elevation']]) - float(row[index['End Depth']])
        if start_el < f['Min Elevation']:
            return False
        elif end_el > f['Max Elevation']:
            return False
        
    if f['Max/Min Depth Interval Filter'] is True:
        start_depth = float(row[index['Start Depth']])
        end_depth = float(row[index['End Depth']])
        if end_depth < f['Min Depth']:
            return False
        elif start_depth > f['Max Depth']:
            return False
        
    for i in range(0,len(toggle_filters)):
        if toggle_filters[i][0] and toggle_filters[i][1]:
            pass
        elif toggle_filters[i][0]:
            if row[index[ toggle_filters[i][2] ]] == 'N':
                return False
        else:
            if row[index[ toggle_filters[i][2] ]] == 'Y':
                return False
    return True

def make_dictionary( ls  ):
    dic = dict()
    for col_title in COLS:
        try:
            dic[ col_title ] = ls.index(col_title)
        except ValueError:
            pass
    return dic
    
def fill_and_transpose_table( table, row_length ):
    for row in table:
        row.extend( ['']*(row_length-len(row)) )
    return [list(col) for col in zip(*table)]    

def ensure_dir(f):
    #d = os.path.dirname(os.path.abspath(f))
    if not os.path.exists(f):
        os.makedirs(f)
    else:
        return f

def header_check(main_header,file_header):
    missing_header=[]
    for header in main_header:
        if header not in file_header:
            missing_header.append(header)
    if len(missing_header) != 0:
        print "The following fields are missing from the input file: {}".format(missing_header)
        sys.exit()
    else:
        pass

def get_max_el(input, filters):
    if filters['Max Elevation'] == 'null':
        max_num = -999
        for f in input:
            ground_el = float(f[index['Ground Elevation']])
            if ground_el > max_num:
                max_num = ground_el
        filters['Max Elevation'] = math.ceil(max_num)

startTime = datetime.now()
print startTime
  
output = list()
index = dict()
max_depth_recorded = 0
min_depth_recorded = 9999

#Export Input Parameters worksheet to a .csv file
with xlrd.open_workbook(INPUT_PARAMETERS_WORKBOOK) as wb:
    sh = wb.sheet_by_name('Input Parameters')
    with open(INPUT_PARAMETERS_FILE, 'wb') as f:
        c = csv.writer(f)
        for r in range(sh.nrows):
            c.writerow(sh.row_values(r))

with open(INPUT_CSV_FILE, 'rb') as csvfile:
    reader = csv.reader(csvfile, dialect='excel',delimiter = ',')
    try:
        file_header = next(reader)
        header_check(COLS,file_header)
        index = make_dictionary( file_header )
        output = dict()

        #Prepping data to be sorted
        locID = index['Location ID']
        grp = index['Location Group Name']
        param = index['Parameter Name']
        lchd = index['Leached']
        input_list = []
        input_list.extend(reader)
        print "Sorting input.csv file for data extraction"
        input_list = sorted(input_list, key=operator.itemgetter(locID))
        input_list = sorted(input_list, key=operator.itemgetter(param))
        input_list = sorted(input_list, key=operator.itemgetter(lchd))
        input_list = sorted(input_list, key=operator.itemgetter(grp))
        x = 0
        filter = get_filters()
        y = 0

        #Add TCLP to front of Parameter Name if the result is Leached.
        for row_ in input_list:
            param_ = index['Parameter Name']
            lchd = row_[ index['Leached'] ]
            if lchd == "Y":
                input_list[y][param_] = "TCLP " + row_[ index['Parameter Name'] ]
            y+=1

        #Finding Max Elevation
        get_max_el(input_list, filter)

        #Going through each row from the sorted CSV file.    
        for row in input_list:
            if is_of_interest( row, get_toggle_filters(get_filters()) ):
                grp = row[ index['Location Group Name'] ]
                param = row[ index['Parameter Name'] ]
                if grp not in output:
                    output[grp]=dict()
                if param not in output[grp]:
                    output[grp][param] = list()
                    output[grp][param].append( ['a'] )
                if filter['Convert BGS to Elevation'] is False:
                    output[grp][param] = add_data( output[grp][param], row, filter )
                    if max_depth_recorded <= float(row[index['End Depth']]):
                        if float(row[index['End Depth']]) <= filter['Max Depth']:
                            max_depth_recorded = float(row[index['End Depth']])
                elif filter['Convert BGS to Elevation'] is True:
                    output[grp][param] = add_data_elevation( output[grp][param], row, filter )
                    row_end_depth_el = float(row[index['Ground Elevation']]) - float(row[index['End Depth']])
                    if min_depth_recorded > row_end_depth_el:
                        if row_end_depth_el >= filter['Min Elevation']:
                            min_depth_recorded = float(row_end_depth_el)
            x = x +1        

    except ValueError:
        import traceback, sys
        tb = sys.exc_info()[2]
        print "line %i" % tb.tb_lineno
        print( "\nFormat error." )
        print( "Column titles of the file %s does not match this "
               "program's specifications." % INPUT_CSV_FILE )
        print( "Please include the following column titles" )
        print( COLS )
csvfile.close()

depth_list = list()
depth_list.append('')
depth_interval = float(filter['Interval Range'])
if filter['Convert BGS to Elevation'] is False:
    i = float(filter['Min Depth'])
    while i <= max_depth_recorded:
        depth_list.append(i)
        i += depth_interval
    DepthRows = (1/depth_interval)*max_depth_recorded + (1/depth_interval) #Calculating the amount of rows based on the depth interval and the max depth 
elif filter['Convert BGS to Elevation'] is True:
    i = float(filter['Max Elevation']) - depth_interval
    while i >= min_depth_recorded:
        interval_list = str(i + depth_interval) + " to " + str(i)
        depth_list.append(interval_list)
        i -= depth_interval
    DepthRows = (1/depth_interval)*(float(filter['Max Elevation'] - min_depth_recorded)) + (1/depth_interval) #Calculating the amount of rows based on the depth interval and the max depth 
   

for grp in output:
    for param in output[grp]:
        output[grp][param].pop(0)
        output[grp][param].insert(0,depth_list)
        output[grp][param] = fill_and_transpose_table( output[grp][param], int(DepthRows) ) #max_depth_recorded
        grp_dir = os.path.join(OUTPUT_DIR,grp)
        ensure_dir(grp_dir)
        with open(os.path.join(grp_dir, param +'.csv'), 'wb') as ofile:
            writer = csv.writer(ofile, dialect='excel')
            for row in output[grp][param]:
                writer.writerow(row)

print "Finished processing the data.  You can find the final output located here: {}".format(OUTPUT_DIR)
print "......................................................................End Runtime: ", datetime.now()-startTime