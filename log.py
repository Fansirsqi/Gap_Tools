from datetime import datetime

def do_loger(context):
    current_now = datetime.now().strftime('%H:%M:%S.%f')
    print(current_now+':',context)
    with open(f'loger.log',mode='a+',encoding='utf-8')as log:
        log.writelines(current_now+' : '+str(context)+'\n')