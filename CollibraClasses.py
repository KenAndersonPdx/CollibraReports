import time
import requests
import json
import CollibraConfig as ccf
def get_date_from_epoch(epoch):
    return time.strftime("%d %b %Y %H:%M:%S", time.localtime(float(epoch)/1000))

class CollibraSession:
    def __init__(self):
        self.s = requests.Session()
        #url = 'https://multco.collibra.com/rest/2.0/auth/sessions'
        url = ccf.baseurl+'auth/sessions'
        data = {
            ##"username": "ApiUser",
            ##"password": "frododog"
            "username": ccf.username,
            "password": ccf.password
            }
##        data = {
##            "username": "integration_api_user",
##            "password": "aRe w# hav1ng fun yet"
##            }
        headers = {
            'Content-Type': 'application/json'
            }
        self.response = self.s.post(url,headers=headers,data=json.dumps(data))
        #self.url_base = 'https://multco.collibra.com/rest/2.0/'
        self.url_base = ccf.baseurl
        self.errlog = open(ccf.errorfile,"w")
    def get_data(self,object_type):
        url = self.url_base + object_type
        response = self.s.get(url)
        response_json = response.json()
        return response_json['results']
    def get_activities(self,obj_type='activities',limit=100):
        url = self.url_base + obj_type + '?limit=' + str(limit)
        response = self.s.get(url)
        response_json = response.json()
        return response_json['results']
    def get_resource_info(self,rtype,rid):
        url = self.url_base + rtype + '/' + rid
        response = self.s.get(url)
        return response.json()
    def get_user_group(self,id):
        url = self.url_base + 'userGroups/' + id
        response = self.s.get(url)
        return response.json()
    def get_user_groups_for_user(self,userid):
        url = self.url_base + 'userGroups?userId=' + userid
        response = self.s.get(url)
        grouplist = ''
        for g in response.json()['results']:
            if not g['system']:
                grouplist += g['name'] + ','
        return grouplist[:-1]
    def close_session(self):
        url = self.url_base + '/auth/sessions/current'
        close_response = self.s.delete(url)
        self.errlog.write('Session logout response: ' + str(close_response.status_code) + '\n')
        self.errlog.close()
        #print(response)
        
class CollibraObject:
    def __init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn):
        self.objectId = objectId
        self.name = name
        self.system = system
        self.createdBy = self.get_user_name(createdBy)
        #self.createdOn = self.get_date_from_epoch(createdOn)
        self.createdOn = get_date_from_epoch(createdOn)
        self.lastModifiedBy = self.get_user_name(lastModifiedBy)
        #self.lastModifiedOn = self.get_date_from_epoch(lastModifiedOn)
        self.lastModifiedOn = get_date_from_epoch(lastModifiedOn)
    def write_core(self,w,row):
            w.write(row,0,self.name)
            w.write(row,1,self.createdBy)
            w.write(row,2,self.createdOn)
            w.write(row,3,self.lastModifiedBy)
            w.write(row,4,self.lastModifiedOn)
    #def get_date_from_epoch(self,epoch):
        #return time.strftime("%d %b %Y %H:%M:%S", time.localtime(float(epoch)/1000))
    def get_name(self):
        return self.name
    def get_user_name(self,userid):
        if self.system:
            return 'System'
        else:
            u = User.get_by_key(userid)
            if u is None:
                return 'User Not Found'
            else:
                return u.getName()
        #will it have access to dictionary created at top level?

class User:
    _member_dict = {}
    @classmethod
    def get_by_key(cls,key):
        if key in cls._member_dict:
            return cls._member_dict[key]
        else:
            return None
    @classmethod
    def get_dict(cls):
        return cls._member_dict
    @classmethod
    def write_users(cls,w,header_format,wrap,gray):
        row = 1
        for u in cls._member_dict.values():
            #CollibraObject.write_core(u,w,row)
            w.write(row,0,u.name)
            w.write(row,1,u.firstName)
            w.write(row,2,u.lastName)
            w.write(row,3,u.emailAddress)
            w.write(row,4,u.activated)
            w.write(row,5,u.enabled)
            w.write(row,6,u.ldapuser)
            w.write(row,7,u.licenseType)
            row += 1
    def __init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn,firstName,lastName,emailAddress,
                 activated,enabled,ldapuser,licenseType):
        self.objectId = objectId
        self.name = name
        self.system = system
        self.createdBy = createdBy
        self.createdOn = self.get_date_from_epoch(createdOn)
        self.lastModifiedBy = lastModifiedBy
        self.lastModifiedOn = self.get_date_from_epoch(lastModifiedOn)
        self.firstName = firstName
        self.lastName = lastName
        self.emailAddress = emailAddress
        self.activated = activated
        self.enabled = enabled
        self.ldapuser = ldapuser
        self.licenseType = licenseType
        self.groups = so.get_user_groups_for_user(self.objectId)
        self.__class__._member_dict[objectId] = self
    def get_date_from_epoch(self,epoch):
        return time.strftime("%d %b %Y %H:%M:%S", time.localtime(float(epoch)/1000))
    def getName(self):
        return self.firstName + ' ' + self.lastName

class Responsibility(CollibraObject):
    _member_dict = {}
    @classmethod
    def get_by_key(cls,key):
        return cls._member_dict[key]
    @classmethod
    def get_by_userid(cls,userid):
        resps = []
        for r in cls._member_dict.values():
            if r.ownerID == userid:
                resps.append(r)
        return resps
    @classmethod
    def write_responsibilities(cls,w,header_format,wrap,gray):
        row = 1
        for r in cls._member_dict.values():
            if r.responsibilityType == 'Responsibility':
                CollibraObject.write_core(r,w,row)
                w.write(row,5,r.roleName)
                w.write(row,6,r.resourceType)
                w.write(row,7,r.resourceName)
                w.write(row,8,r.ownerName)
                row += 1
        widths = [15,15,20,20,25,15,50,25]
        for ix,wi in enumerate(widths):
            w.set_column(ix,ix,wi)
    @classmethod
    def write_global_responsibilities(cls,w,header_format,wrap,gray):
        row = 1
        for r in cls._member_dict.values():
            if r.responsibilityType == 'Global':
                CollibraObject.write_core(r,w,row)
                w.write(row,5,r.roleName)
                w.write(row,6,r.ownerType)
                w.write(row,7,r.ownerName)
                row += 1
        widths = [15,15,20,20,25,20,50]
        for ix,wi in enumerate(widths):
            w.set_column(ix,ix,wi)    
    def __init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn,
                 responsibilityType,roleId,roleName,resourceId,resourceType,ownerType,ownerName):
        CollibraObject.__init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,
                                lastModifiedOn)
        self.responsibilityType = responsibilityType
        self.roleId = roleId
        self.roleName = roleName
        self.resourceId = resourceId,
        self.resourceType = resourceType
        if responsibilityType == 'Responsibility':
            if resourceType == 'Asset':
                rtype = 'assets'
            elif resourceType == 'Community':
                rtype = 'communities'
            elif resourceType == 'Domain':
                rtype = 'domains'
            try:
                rinfo = so.get_resource_info(rtype,resourceId)
                self.resourceName = rinfo['name']
            except:
                self.resourceName = 'Unknown'
                msg = 'resource not found for: ' + rtype + resourceId + '\n'
                so.errlog.write(msg)
                for k in rinfo.keys():
                    msg = k + ': ' + rinfo[k] + '\n'
                    so.errlog.write(msg)       
        self.ownerType = ownerType
        self.ownerName = ownerName
        self.__class__._member_dict[objectId] = self
class AssetType(CollibraObject):
    _member_dict = {}
    _active_domain_types = []
    @classmethod
    def get_by_key(cls,key):
        return cls._member_dict[key]
    @classmethod
    def get_list_of_asset_type_names(cls):

        assetnames = {k:v.name for (k,v) in cls._member_dict.items()}

        return sorted(assetnames.values())
    @classmethod
    def build_family_tree(cls):
        for a in cls._member_dict.values():
            if a.parentName != 'Asset':
                cls.get_by_key(a.parentId).children.append(a)       
    @classmethod
    def get_levels(cls):
        maxlevel = 0
        nodes = []
        for at in cls._member_dict.values():
            if at.parentName == 'Asset':
                stack = [(0,at)]
                while stack:
                    cur_node = stack[0]
                    stack = stack[1:]
                    nodes.append(cur_node)
                    childrev = cur_node[1].children[:]
                    childrev.reverse()
                    for child in childrev:
                        curr_level = cur_node[0] + 1
                        if curr_level > maxlevel:
                            maxlevel = curr_level
                        stack.insert(0,(curr_level,child))
        return maxlevel,nodes 
    @classmethod
    def write_asset_types(cls,w,header_format,wrap,gray):  #takes workbook as parm
        maxlevel, nodes = cls.get_levels()
        for i in range(0,maxlevel+1):
            w.write(0,i,'Level ' + str(i),header_format)                    
        asset_type_headers = ['InUse?','Created By','Created On','Last Modified By','Last Modified On','Description',
                              'SymbolType','Acronym','Display Name Enabled','Rating Enabled']
        for i,h in enumerate(asset_type_headers,start=maxlevel+1):
            w.write(0,i,h,header_format)
        row = 1
        #for a in cls._member_dict.values():
        for n in nodes:
            #CollibraObject.write_core(a,w,row)
            if n[1].isInUse:
                w.write(row,n[0],n[1].name)
                w.write(row,maxlevel + 1,'X')
                bgformat = None
            else:
                bgformat = gray
                w.write(row,n[0],n[1].name,gray)
                for i in range(n[0]+1,maxlevel+2):
                    w.write(row,i,'',gray)
            w.write(row,maxlevel + 2,n[1].createdBy,bgformat)
            w.write(row,maxlevel + 3,n[1].createdOn,bgformat)
            w.write(row,maxlevel + 4,n[1].lastModifiedBy,bgformat)
            w.write(row,maxlevel + 5,n[1].lastModifiedOn,bgformat)
            w.write(row,maxlevel + 6,n[1].description,bgformat)
            w.write(row,maxlevel + 7,n[1].symbolType,bgformat)
            w.write(row,maxlevel + 8,n[1].acronymCode,bgformat)
            w.write(row,maxlevel + 9,n[1].displayNameEnabled,bgformat)
            w.write(row,maxlevel + 10,n[1].ratingEnabled,bgformat)
##            if not n[1].isInUse:
##                w.set_row(row,None,gray)
            row += 1
        for i in range(0,maxlevel+1):
            w.set_column(i,i,16,wrap)
        widths = [6,10,20,10,20,110,15,8,8,8]
        for ix,wi in enumerate(widths,start=maxlevel+1):
            if ix == maxlevel+6 or ix == maxlevel+9 or ix == maxlevel+10:
                w.set_column(ix,ix,wi,wrap)
            else:
                w.set_column(ix,ix,wi)
        w.freeze_panes(1,maxlevel+1)
    @classmethod
    def process_assignments(cls):
        assetmap = {}
        assetmap2 = {}
        for asset_type in cls._member_dict.values():
            dlist = asset_type.assignment['domainTypes']
            #print('asset type is ',asset_type.name)
            for d in dlist:
                #print(d['name'])
                dt = DomainType.get_by_key(d['id'])
                if dt:
                    dt.add_asset_type(asset_type.name)
                    if dt.is_in_use():
                        asset_type.isInUse = True
            clist = asset_type.assignment['characteristicTypes']
            for c in clist:
                if 'attributeType' in c:
                    att = c['attributeType']
                    attid = att['id']
                    att_type = AttributeType.get_by_key(attid) #get the attribute type instance
                    att_type.add_asset_type(asset_type.name) #add the asset type to the attribute type
                elif 'relationType' in c:
                    roleDirection = c['roleDirection']
                    rel = c['relationType']
                    relId = rel['id']
                    head = rel['sourceType']['name']         
                    tail = rel['targetType']['name']
                    role = rel['role']
                    coRole = rel['coRole']
                    if head in assetmap:
                        if tail in assetmap[head].keys():
                            assetmap[head][tail].append((asset_type.name,role,coRole))
                        else:
                            assetmap[head][tail] = [(asset_type.name,role,coRole)]
                    else:
                        assetmap[head] = {tail:[(asset_type.name,role,coRole)]}
                    if asset_type.name in assetmap2:
                        assetmap2[asset_type.name].append((roleDirection,relId))
                    else:
                        assetmap2[asset_type.name] = [(roleDirection,relId)]
        return assetmap,assetmap2         
       
    def __init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn,description,
                 parentId,parentName,symbolType,acronymCode,displayNameEnabled,ratingEnabled,assignment):
        CollibraObject.__init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn)
        self.description = description
        self.parentId = parentId
        self.parentName = parentName
        self.symbolType = symbolType
        self.acronymCode = acronymCode
        self.displayNameEnabled = displayNameEnabled
        self.ratingEnabled = ratingEnabled
        self.assignment = assignment
        self.children = []
        self.isInUse = False
        self.__class__._member_dict[objectId] = self

class AttributeType(CollibraObject):
    _member_dict = {}
    @classmethod
    def get_by_key(cls,key):
        return cls._member_dict[key]
    @classmethod
    def get_dict_of_asset_types_assigned_to(cls):
        return {v.name:v.assignedToAssetTypes for (v) in cls._member_dict.values()}
    @classmethod
    def write_attribute_types(cls,w,wrap):
        row = 1
        for at in cls._member_dict.values():
            #wd = []
            CollibraObject.write_core(at,w,row)
            w.write(row,5,at.resourceType)
            w.write(row,6,at.description)
            if at.stringType is None:
                st = ''
            else:
                st = at.stringType
            w.write(row,7,st)
            if at.isInteger is None:
                ii = ''
            else:
                ii = at.isInteger
            w.write(row,8,ii)
            if at.allowedValues is None:
                av = ''
            else:
                av = ''
                for val in at.allowedValues:
                    av += val +','
                av = av[:-1]
            w.write(row,9,av)
            row += 1
        widths = [28,15,20,15,20,15,60,10,8,60]
        for ix,wi in enumerate(widths):
            if ix == 6 or ix == 9:
                w.set_column(ix,ix,wi,wrap)
            else:
                w.set_column(ix,ix,wi)
        #w.set_column(6,6,wrap)
        #w.set_column(9,9,wrap)
    def __init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn,resourceType,
                 description,stringType,isInteger,allowedValues):
        CollibraObject.__init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn)
        self.resourceType = resourceType
        self.description = description
        self.stringType = stringType
        self.isInteger = isInteger
        self.allowedValues = allowedValues
        self.assignedToAssetTypes = []
        self.__class__._member_dict[objectId] = self
    def add_asset_type(self,at_name):
        self.assignedToAssetTypes.append(at_name)

class Community(CollibraObject):
    _member_dict = {}
    _cd = {}
    @classmethod
    def get_by_key(cls,key):
        return cls._member_dict[key]
    @classmethod
    def build_family_tree(cls):
        for c in cls._member_dict.values():
            if c.parentId is not None:
                #__class__.get_by_key(c.parentId).children.append(('Community',c.objectId,c.name))
                cls.get_by_key(c.parentId).children.append(('Community',c.objectId,c.name))
            domchildren = Domain.get_child_domains(c.name)
            c.children += domchildren
        for c in cls._member_dict.values():
            c.children.sort(key=cls.skey)
            #print(c.name,c.children)
        
    @classmethod
    def skey(cls,c):
        return c[0] + c[2]
    @classmethod
    def get_levels(cls):
        #amazing, this actually seems to be working
        #next step, return maxlevel, use that to format the
        #columns on spreadsheet (start attributes after the total # of levels
        #in the community domain tree
        #using the logic below, printing them out should be straightforward.
        maxlevel = 0
        nodes = []
        for c in cls._member_dict.values():
            #print(c.name)
            if c.topLevel:
                stack = [(0,c)]
                while stack:
                    cur_node = stack[0]
                    stack = stack[1:]
                    nodes.append(cur_node)
                    #need to deal with Community vs. domai
                    if isinstance(cur_node[1],Community):
                        childrev = cur_node[1].children[:]
                        childrev.reverse()
                        for child in childrev:
                            #problem is childe could be community or domain
                            #as long as function in superclass exists
                            curr_level = cur_node[0] + 1
                            if curr_level > maxlevel:
                                maxlevel = curr_level
                            if child[0] == 'Community':
                                #stack.insert(0,(curr_level,__class__.get_by_key(child[1])))
                                stack.insert(0,(curr_level,cls.get_by_key(child[1])))
                            else:
                                stack.insert(0,(curr_level,Domain.get_by_key(child[1])))  
        return maxlevel,nodes
                                     
##    @classmethod
##    def write_communities(cls,w):
##        for c in cls._member_dict.values():
##            if c.topLevel:
##                nodes = []
##                stack = [c.objectId]
##                level = 0
##                while stack:
##                    cur_node = stack[0]
##                    stack = stack[1:]          
    def __init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn,description,
                 parentId):
        CollibraObject.__init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn)
        self.description = description
        self.parentId = parentId
        self.children = []
        if parentId is None:
            self.topLevel = True
        else:
            self.topLevel = False
        self.__class__._member_dict[objectId] = self
    def get_type(self):
        return 'Community'
  
class DomainType(CollibraObject):
    _member_dict = {}
    @classmethod
    def get_by_key(cls,key):
        if key in cls._member_dict:
            return cls._member_dict[key]
        else:
            return False
    @classmethod
    def get_list_of_domain_type_names(cls):
        dtdict = {k:v.name for (k,v) in cls._member_dict.items()}
        return sorted(dtdict.values())
    @classmethod
    def get_dict_of_asset_types_assigned_to(cls):
        return {v.name:v.assignedToAssetTypes for (v) in cls._member_dict.values()}
        
    @classmethod
    def write_domain_types(cls,w,wrap):
        row = 1
        for dt in cls._member_dict.values():
            CollibraObject.write_core(dt,w,row)
            w.write(row,5,dt.description)
            w.write(row,6,dt.parent)
            row += 1
        widths = [25,15,20,15,20,130,24]
        for i,wi in enumerate(widths):
            if i ==5:
                w.set_column(i,i,wi,wrap)
            else:
                w.set_column(i,i,wi)
            w.set_column(i,i,wi)
        #w.set_column(5,5,wrap)
    def add_asset_type(self,at):
        self.assignedToAssetTypes.append(at)
    def set_in_use(self):
        self.isInUse = True
    def is_in_use(self):
        return self.isInUse
    def __init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn,description,
                 parent):
        CollibraObject.__init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn)
        self.description = description
        self.parent = parent
        self.assignedToAssetTypes = []
        self.isInUse = False
        self.__class__._member_dict[objectId] = self
    
    
class Domain(CollibraObject):
    _member_dict = {}
    @classmethod
    def get_by_key(cls,key):
        return cls._member_dict[key]
    @classmethod
    def get_child_domains(cls,parent):
        #returns a list of domain IDs belonging to a parent community
        domlist = []
        for d in cls._member_dict.values():
            if d.community == parent:
                domlist.append(('Domain',d.objectId,d.name))
        return domlist
    @classmethod
    def write_domains(cls,w):
##        for dt in DomainType._member_dict.values():
##            print(dt.name,dt.isInUse)
        row = 1
        for d in cls._member_dict.values():
            CollibraObject.write_core(d,w,row)
            w.write(row,5,d.domainTypeName)
            w.write(row,6,d.community)
            row += 1
        widths = [54,15,20,15,20,25,42]
        for i,wi in enumerate(widths):
            w.set_column(i,i,wi)
    def __init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn,domainTypeId,
                 domainTypeName,community,description):
        CollibraObject.__init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn)
        self.domainTypeId = domainTypeId
        self.domainTypeName = domainTypeName
        self.community = community
        self.description = description
        dt = DomainType.get_by_key(self.domainTypeId)
        if dt:
            dt.set_in_use()
        self.__class__._member_dict[objectId] = self
    def get_type(self):
        return 'Domain'

class RelationType(CollibraObject):
    _member_dict = {}
    @classmethod
    def get_by_key(cls,key):
        return cls._member_dict[key]
    @classmethod
    def write_relation_types(cls,w):
        row = 1
        for rt in cls._member_dict.values():
            CollibraObject.write_core(rt,w,row)
            w.write(row,5,rt.head)
            w.write(row,6,rt.role)
            w.write(row,7,rt.coRole)
            w.write(row,8,rt.tail)
            row += 1
        widths = [6,15,20,15,20,24,20,20,24]
        for i,wi in enumerate(widths):
            w.set_column(i,i,wi)
    def __init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn,
                 head,role,coRole,tail):
        CollibraObject.__init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn)
        self.head = head
        self.role = role
        self.coRole = coRole
        self.tail = tail
        self.__class__._member_dict[objectId] = self

class Asset(CollibraObject):
    _member_dict = {}
    @classmethod
    def write_assets(cls,w,wrap):
        row = 1
        assets = sorted(cls._member_dict.values(),key=lambda a: a.domain+a.assetType+a.name)
        #for a in assets.sort(key=lambda a: a.domain+a.assetType+a.name)
        for a in assets:
            #CollibraObject.write_core(a,w,row)
            w.write(row,0,a.domain)
            w.write(row,1,a.assetType)
            w.write(row,2,a.displayName)
            w.write(row,3,a.name)
            w.write(row,4,a.createdBy)
            w.write(row,5,a.createdOn)
            w.write(row,6,a.lastModifiedBy)
            w.write(row,7,a.lastModifiedOn)
            w.write(row,8,a.status)
            row += 1
        widths = [48,24,27,55,17,20,17,20,10]
        for i,wi in enumerate(widths):
            w.set_column(i,i,wi,wrap)
        w.freeze_panes(1,0)
    def __init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn,
                 displayName,excludedFromAutoHyperlinking,assetType,status,domain):
        CollibraObject.__init__(self,objectId,name,system,createdBy,createdOn,lastModifiedBy,lastModifiedOn)
        self.displayName = displayName
        self.excludedFromAutoHyperlinking = excludedFromAutoHyperlinking
        self.assetType = assetType
        self.status = status
        self.domain = domain
        self.__class__._member_dict[objectId] = self
    

if __name__ == '__main__':
    my_user = User(1,'usrname',True,'Me','1550128291934','Me','1550128291934','Frodo','Baggins',
                   'frodo@shire.com',True,False,False,'AUTHOR')
    print(my_user.getName())
    print(my_user.createdOn)
    
so = CollibraSession()
