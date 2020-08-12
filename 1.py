# -*- coding: UTF-8 -*-
import requests as req
import json,sys,time
#先注册azure应用,确保应用有以下权限:
#files:	Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
#user:	User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
#mail:  Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
#注册后一定要再点代表xxx授予管理员同意,否则outlook api无法调用

###################################################################
#把下方单引号内的内容改为你的应用id                                         #
id=r'8505ba2b-750f-4623-a3fc-4678ffecf4db'                         
#把下方单引号内的内容改为你的应用机密                                       #
secret=r'OAgABAAAAAAAGV_bv21oQQ4ROqh0_1-tAAQDs_wIA9P-qPdNGrRtV9oF2e88OqxOlHm33eVhLtcd2B2rq736-fFAKGIaj_h_q01XPJjGdv27_xtMIhaFTknEWM19J_RDLhqYnWgJUaTH3L5J6baL8yjvgHuz0uOJ0eUvCocReB-fzhHBAzN2Sms2JfNygtzrxQwUX3bnVIt-HkGevnnLdihEPU3FsE6syJ0Upq1tvRKEAtjKE09zeWlpDSCBqefzgtZDSoUp2nExp3J4yt3rdMZsas0BRZr4o8OfYoTBhPZ3_CPzQyJVitf0-qngjzIgB-8KcSYfHuEfvdUUeYm2qw13_jNLxF3cVACR-oOfyoaRNYLjvMduHGKuSQWGWHtp05FhPVtnZh03QWg6hBcMpei-98B0MJGxTfijJ3Aru5OJLqxoaCFmLUe9p63_DW6CkeksMZs5jCpx2huVnF-h4yUzOfMRMkoIix2qKcar5P_oz2PCVeJWzLL_qnNY3I6wFCLNwxEigYgf_OPbp_shf5JE33tvMDI83b0hJP94cgiyiqEatkx5KudIE-k75S_VWPSBlwnaYfNqcP-ECa1Jxkx56ixs-jO_VrtJie0KR0ePX13rr_9Y0OeBMvoAOlrNaDqwyswupaQ-Epr1gHSxpK7bsqlJDeU7duYWRsT83EMDWkwy63FyqDRhMqqjViPZGv9cjdGSXGpUnUS3kJbuhQIbFH4Q_FeAfK9WRH-_uvbKMvk3AeauOJO1co9YfTvi4ii_Lw6zDxjYybaiC2yhaDO3JcnC5ouxca4X2HR_ZSYDZgSzLXUhp6y6K0Ok6Dr7XG9YXYBPEcMiWXhH6_c9IcrQ1uk_i9Dgo-Z2doFwh_hA8oj9fdJstbr8AhXYfRRnpHK8PyQ8QaDpCs_NjzWtI76AeWOOSZXc-R4lXZKMPQ33cf8rPOOtxK1PZqq5ifWCTquNl0bc-FsZSmDkF5ekq'                                           
###################################################################

path=sys.path[0]+r'/1.txt'
num1 = 0

def gettoken(refresh_token):
    headers={'Content-Type':'application/x-www-form-urlencoded'
            }
    data={'grant_type': 'refresh_token',
          'refresh_token': refresh_token,
          'client_id':id,
          'client_secret':secret,
          'redirect_uri':'https://wangziyingwen.github.io/GetAutoApiToken/'
         }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    with open(path, 'w+') as f:
        f.write(refresh_token)
    return access_token
def main():
    fo = open(path, "r+")
    refresh_token = fo.read()
    fo.close()
    global num1
    access_token=gettoken(refresh_token)
    headers={
    'Authorization':access_token,
    'Content-Type':'application/json'
    }
    try:
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive',headers=headers).status_code == 200:
            num1+=1
            print("2调用成功"+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/drive/root',headers=headers).status_code == 200:
            num1+=1
            print('3调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/users ',headers=headers).status_code == 200:
            num1+=1
            print('4调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/messages',headers=headers).status_code == 200:
            num1+=1
            print('5调用成功'+str(num1)+'次')    
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',headers=headers).status_code == 200:
            num1+=1
            print('6调用成功'+str(num1)+'次')    
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',headers=headers).status_code == 200:
            num1+=1
            print('7调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root/children',headers=headers).status_code == 200:
            num1+=1
            print('8调用成功'+str(num1)+'次')
        if req.get(r'https://api.powerbi.com/v1.0/myorg/apps',headers=headers).status_code == 200:
            num1+=1
            print('8调用成功'+str(num1)+'次') 
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders',headers=headers).status_code == 200:
            num1+=1
            print('9调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories',headers=headers).status_code == 200:
            num1+=1
            print('10调用成功'+str(num1)+'次')
    except:
        print("pass")
        pass
for _ in range(3):
    main()
