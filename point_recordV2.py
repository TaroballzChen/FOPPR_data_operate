import glob,os

def find_latest_file():                                         #獲取最新的lvm檔案
    os.chdir(os.getcwd())
    All_of_origin_file = glob.glob('*.lvm')
    get_latest_file    = max(All_of_origin_file,key=os.path.getmtime)
    return get_latest_file

def read_file(filename=find_latest_file()):                     #獲取最新lvm檔案中最新的數據(要記錄的x點)
    with open (filename,'r') as onlyread_data:
        latest_data = onlyread_data.readlines()[-1]
        get_x_point = int(float(latest_data.split('\t')[0]))    #將字符串-->float-->int
        return get_x_point

def write_conc_with_x_point(filename=find_latest_file(),conc=""):                 #將獲取的x值存在新檔案
    point_record = filename.replace('.lvm','.txt')
    concentration = conc
    with open(point_record,'a') as record_point:
        record_point.write('\t'.join([conc,str(read_file()),'\n']))
    print('以記錄x_point---->',read_file())
    return concentration

def concentration_record(how_many_point = int(input('請輸入要做幾個濃度(包含blank)：'))):
    All_concentration =[]
    for record_conc in range (how_many_point):
        if record_conc == 0: record_conc='blank'
        conc_input = input('請輸入濃度(EXCEL格式),blank輸入blank_%s：' % record_conc)
        All_concentration.append(conc_input)
        print('以記錄濃度 ----->', All_concentration)
    return All_concentration

for each_conc in concentration_record():
    print('將使用%s'% each_conc,'M 濃度紀錄數x值，請確認濃度無誤，並按任意鍵紀錄x_point')
    os.system("pause")
    read_file()
    write_conc_with_x_point(conc=each_conc)
else:
    print('=*='*10,"Deisgn By Yuan-Yu Chen",'=*='*10)
    os.system("pause")