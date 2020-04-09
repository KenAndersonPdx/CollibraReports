import xlsxwriter
import CollibraClasses as cc
import csv
import math
import CollibraConfig as ccf

#permfile = 'h:/Collibra/permissions.csv'

#so = cc.CollibraSession()
#sesh = sessionobject.s
#reportfile = 'h:/Collibra/UserReport.xlsx'

workbook = xlsxwriter.Workbook(ccf.reportfile)
header_format = workbook.add_format({'bold': True})
wrap = workbook.add_format()
wrap.set_text_wrap()
gray = workbook.add_format()
gray.set_bg_color('#C0C0C0')
#author_permissions = 
permissions = {}
roledict = {}
role_permissions = {}
user_roles = {} 
group_membership = {}

#errlog = open(ccf.errorfile,"w")

with open(ccf.permfile) as f:
    for line in csv.reader(f):
        permissions[line[4]] = line[6]

def user_license_required(username):
    if username in user_roles:
        roles = user_roles[username]
        lic = 'Consumer'
        for r in roles:
            if r[3][0] == 'Author':
                lic = 'Author'
                break
        return lic
    else:
        return 'Consumer'

users = cc.so.get_data('users')
for u in users:
    if 'createdBy' in u.keys():
            cb = u['createdBy']
    else:
        cb = 'Unknown'
    if 'lastModifiedBy' in u.keys():
        lmb = u['lastModifiedBy']
    else:
        lmb = 'Unknown'
    cc.User(u['id'],u['userName'],u['system'],cb,u['createdOn'],lmb,u['lastModifiedOn'],
             u['firstName'],u['lastName'],u['emailAddress'],u['activated'],u['enabled'],u['ldapUser'],u['licenseType'])
    

roles = cc.so.get_data('roles')
for r in roles:
    rolename = r['name']
    license = 'Consumer'
    roleperms = []
    for p in r['permissions']:
        roleperms.append(p)
        try:
            if permissions[p] == 'Author':
                license = 'Author'
                #break
        except KeyError as error:
            errmsg = 'Key error: ' + error + '\n'
            #print(error)
            cc.so.errlog.write(errmsg)
            
    roledict[rolename] = (license,r['global'],r['system'])
    role_permissions[rolename] = roleperms

#for k,v in roledict.items():
#    print(k,v)

responsibilities = cc.so.get_data('responsibilities')
for r in responsibilities:
    if r['owner']['resourceType'] == 'UserGroup':
        ownername = cc.so.get_user_group(r['owner']['id'])['name']
    elif r['owner']['resourceType'] == 'User':
        user = cc.User.get_by_key(r['owner']['id'])
        ownername = user.firstName + ' ' + user.lastName
    else:
        ownername = 'unknown'
    if 'baseResource' in r.keys():
        rolename = r['role']['name']
        resourcetype = r['baseResource']['resourceType']
        resourceid = r['baseResource']['id']
        cc.Responsibility(r['id'],'Responsibility',r['system'],r['createdBy'],r['createdOn'],r['lastModifiedBy'],
            r['lastModifiedOn'],'Responsibility',r['role']['id'],rolename,resourceid,resourcetype,
            r['owner']['resourceType'],ownername)
        if r['owner']['resourceType'] == 'User':
            owner = cc.User.get_by_key(r['owner']['id']) 
            if owner.name in user_roles:
                user_roles[owner.name].append((rolename,resourcetype,
                                               cc.so.get_resource_info(resourcetype,resourceid),roledict[rolename]))
            else:
                user_roles[owner.name] = [(rolename,resourcetype,
                                               cc.so.get_resource_info(resourcetype,resourceid),roledict[rolename])]
        #else:
            #print('user record not found for id: ',r['owner']['id'])
            #errmsg = 'User record not found for id: ' + r['owner']['id'] + 'owner type is ' + r['owner']['resourceType'] + '\n'
            #cc.so.errlog.write(errmsg)
    else:  #presumably it's global
        #print(r['role']['name'])
        cc.Responsibility(r['id'],'Responsibility',r['system'],r['createdBy'],r['createdOn'],r['lastModifiedBy'],
            r['lastModifiedOn'],'Global',r['role']['id'],r['role']['name'],'NA','NA',r['owner']['resourceType'],ownername)
        
        #else: #need to handle case where global role assigned to one or more users
                             
#write users
w = workbook.add_worksheet('Users')
headers = ['Username','First Name','Last Name','Email','Activated?','Enabled?','LDAP User?',
           'License Type Assigned','License Type Required','Groups','License Issue',]
for c,h in enumerate(headers):
    w.write(0,c,h,header_format)
ud = cc.User.get_dict()
row = 1
assigned_license_count = {'author':0,'consumer':0}
required_license_count = {'author':0,'consumer':0}
for u in ud.values():
    #build usergroup dictionary
    if u.groups != '':
        for ug in u.groups.split(','):
            if ug in group_membership:
                group_membership[ug].append(u.firstName + ' ' + u.lastName)
            else:
                group_membership[ug] = [u.firstName + ' ' + u.lastName]
    #write user row
    w.write(row,0,u.name)
    w.write(row,1,u.firstName)
    w.write(row,2,u.lastName)
    w.write(row,3,u.emailAddress)
    w.write(row,4,u.activated)
    w.write(row,5,u.enabled)
    w.write(row,6,u.ldapuser)
    w.write(row,7,u.licenseType)
    ulr = user_license_required(u.name)
    w.write(row,8,ulr)
    w.write(row,9,u.groups)
    if u.licenseType.lower() != ulr.lower():
        w.write(row,10,'X')
    assigned_license_count[u.licenseType.lower()] = assigned_license_count[u.licenseType.lower()] + 1
    required_license_count[ulr.lower()] = required_license_count[ulr.lower()] + 1
    row += 1
row += 1
w.write(row,0,'author licenses assigned: ' + str(assigned_license_count['author']))
row += 1
w.write(row,0,'author licenses required: ' + str(required_license_count['author']))
widths = [15,15,20,30,10,10,15,25,25,30,15]
for ix,wi in enumerate(widths):
    w.set_column(ix,ix,wi)

#write responsbilities
w = workbook.add_worksheet('Responsibilities')
headers = ['Object Name','Created By','Created On','Last Modified By','Last Modified On','Role Name',
           'Resource Type','Resource Name','Assigned User']
for c,h in enumerate(headers):
    w.write(0,c,h,header_format)
cc.Responsibility.write_responsibilities(w,header_format,wrap,gray)

#write global responsibilities
w = workbook.add_worksheet('GlobalResp')
headers = ['Object Name','Created By','Created On','Last Modified By','Last Modified On','Role Name',
           'Member Type','Member Name']
for c,h in enumerate(headers):
    w.write(0,c,h,header_format)
cc.Responsibility.write_global_responsibilities(w,header_format,wrap,gray)

#write user groups
for k in group_membership.keys():
    print(k,group_membership[k])

#print list of roles
w = workbook.add_worksheet('Role Definitions')
headers = ['Role','License Required','Type','System?','Permissions']
for c,h in enumerate(headers):
    w.write(0,c,h,header_format)
row = 1
for k,v in roledict.items():
    w.write(row,0,k)
    w.write(row,1,v[0])
    if v[1]:
        rtype = 'Global'
    else:
        rtype = 'Resource'
    w.write(row,2,rtype)
    if v[2]:
        sys = 'X'
    else:
        sys = ''
    w.write(row,3,sys)
    permlist = ''
    for perm in role_permissions[k]:
        permlist += perm + ','
    w.write(row,4,permlist[:-1])
    #print(k,v[0],v[1],role_permissions[k])
    row += 1

workbook.close()
#errlog.close()
cc.so.close_session()

###compare roles to see if any roles with exact same permissions
##
###get list of permissions
##p = permissions.keys()
##role_matrix = []
###for each role, add a list representing vector of permissions, true if role has it
###false otherwise
##for k,v in role_permissions.items():
##    #role_matrix[k] = [rp in p for rp in v]
##    #print(k,v)
##    row = [1 if rp in v else 0 for rp in p]
##    row.insert(0,k)
##    role_matrix.append(row)
##
##def distance(v1,v2):
##    if len(v1) != len(v2):
##        return None
##    s = 0
##    for i in range(0,len(v1)):
##        s += math.pow((v1[i]-v2[i]),2)
##    return math.sqrt(s)
##        
##
##for i in range(0,len(role_matrix)-1):
##    #print(role_matrix[i])
##    #print(role_matrix[i][0],roledict[role_matrix[i][0]][1])
##    if not roledict[role_matrix[i][0]][1]:
##        for j in range(i+1,len(role_matrix)):
##            #print(role_matrix[i][0],role_matrix[j][0],distance(role_matrix[i][1:],role_matrix[j][1:]))
##            #only print if distance is one or 0
##            #also should check some of the numbers look odd
##            dist = distance(role_matrix[i][1:],role_matrix[j][1:])
##            if dist == 0:
##                print(role_matrix[i][0],role_matrix[j][0],dist)
    
    

