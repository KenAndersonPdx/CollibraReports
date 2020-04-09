import xlsxwriter
import CollibraClasses as cc
import json

reportfile = 'h:/Collibra/activities.xlsx'
workbook = xlsxwriter.Workbook(reportfile)
w = workbook.add_worksheet('Activities')
header_format = workbook.add_format({'bold': True})
#keys = []
activities = cc.so.get_activities(limit=10000)
col_index = {
    'old:id':6,
    'old:name':7,
    'old:type':8,
    'kind':9,
    'field':10,
    'affected:id':11,
    'affected:name':12,
    'affected:type':13,
    'new:id':14,
    'new:name':15,
    'new:type':16,
    'role:id':17,
    'role:name':18,
    'role:type':19,
    'people:id':20,
    'people:name':21,
    'people:type':22,
    'resource:id':23,
    'resource:name':24,
    'resource:type':25,
    'role':26,
    'coRole':27,
    'source:id':28,
    'source:name':29,
    'source:type':30,
    'target:id':31,
    'target:name':32,
    'target:type':33,
    'new':34,
    'old':35,
    'textDifference':36,
    'businessItem:id':37,
    'businessItem:name':38,
    'businessItem:type':39,
    'rating':40,
    'salt':41,
    'attachmentFile':42}
    

headers1 = ['','','','','','','Old','','','','','affected','','','new','','','role','','',
            'people','','','resource','','','','','source','','','target','','','','',
            '','businessItem','','','','']
headers2 = ['UserName','Epoch', 'Timestamp', 'Cause','Call Id','Activity Type',
           'id','name','type','kind','field','id','name','type','id','name','type',
            'id','name','type','id','name','type','id','name','type','role','coRole',
            'id','name','type','id','name','type','new','old','textDifference','id','name','type',
            'rating','salt','attachmentFile']
for c,h in enumerate(headers1):
    w.write(0,c,h,header_format)
for c,h in enumerate(headers2):
    w.write(1,c,h,header_format)
row = 2
for a in activities:
    desc = json.loads(a['description'])
    
##    for d in desc:
##        if type(desc[d]) is dict:
##            for d2 in desc[d]:
##                #print(d + ':' + d2 + ': ' + desc[d][d2])
##                k = d+':'+d2
##                if k not in keys:
##                    keys.append(k)
##        else:
##            #print(d+': ' + desc[d])
##            if d not in keys:
##                keys.append(d)
        
    w.write(row,0,a['user']['userName'])
    w.write(row,1,a['timestamp'])
    w.write(row,2,cc.get_date_from_epoch(a['timestamp']))
    w.write(row,3,a['cause'])
    w.write(row,4,a['callId'])
    #w.write(row,4,a['callCount'])
    w.write(row,5,a['activityType'])
    #w.write(row,6,a['description'])
    for d in desc:
        if type(desc[d]) is dict:
            for d2 in desc[d]:
                k = d+':'+d2
                w.write(row,col_index[k],desc[d][d2])
        else:
            w.write(row,col_index[d],desc[d])
    row += 1

##for x in keys:
##    print(x)
    
workbook.close()
    
            
        
    
