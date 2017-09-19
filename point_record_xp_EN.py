import glob,os

def find_latest_file():                                         
    os.chdir(os.getcwd())
    All_of_origin_file = glob.glob('*.lvm')
    get_latest_file    = max(All_of_origin_file,key=os.path.getmtime)
    return get_latest_file

def read_file(filename=find_latest_file()):                     
    with open (filename,'r') as onlyread_data:
        latest_data = onlyread_data.readlines()[-1]
        get_x_point = int(float(latest_data.split('\t')[0]))    
        return get_x_point

def write_conc_with_x_point(filename=find_latest_file(),conc=""):                 
    point_record = filename.replace('.lvm','.txt')
    concentration = conc
    with open(point_record,'a') as record_point:
        record_point.write('\t'.join([conc,str(read_file()),'\n']))
    print('point record x_point---->',read_file())
    return concentration

def concentration_record(how_many_point = int(input('input how many concentration do you want to experiment (include blank):'))):
    All_concentration =[]
    for record_conc in range (how_many_point):
        if record_conc == 0: record_conc='blank'
        conc_input = input('input the concentration (EXCEL form) ,if blank please type blank _%s:' % record_conc)
        All_concentration.append(conc_input)
        print('point recorded ----->', All_concentration)
    return All_concentration

for each_conc in concentration_record():
    print('use the %s'% each_conc,' g/mL concentration record the x value, please confirm your concentration is correct ,and please enter any key to record x_point')
    os.system("pause")
    read_file()
    write_conc_with_x_point(conc=each_conc)
else:
    print('=*='*10,"Deisgn By Yuan-Yu Chen",'=*='*10)
    os.system("pause")