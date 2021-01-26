# Function to set the name order (surname-firstname)
def SurnameFirst(namesDic,sn):
    oldnamelist=[]
    swap=0
    for indiv in namesDic:
        oldnamelist=oldnamelist+[indiv['name'].replace(',','').replace('.','')]
        #print(oldnamelist)
    for name in oldnamelist:
        if sn in name.split(' '):
            if name.split(' ').index(sn)==0:
                swap=0
                break;
            else:
                swap=1
                break;
    if swap:
        newnamelist=[]
        for name in oldnamelist:
            namesplit=name.split(' ')
            names=[namesplit[-1]]+namesplit[:-1]
            newnamelist=newnamelist+[' '.join(names)]
    else:
        newnamelist=oldnamelist
    return newnamelist

def ReturnDictWOerror(dictdata,key,nodata):
    if key in dictdata.keys():
        return dictdata[key]
    else:
        return nodata

def ReturnDictContent(dictdata,key,key1,nodata=''):
    d=ReturnDictWOerror(dictdata,key,nodata)
    d1=ReturnDictWOerror(dictdata,key1,nodata)
    if d!=nodata:
        return d
    else:
        return d1

def commaR(vol,spage):
    if (vol=='') & (spage==''):
        return ''
    elif (vol==''):
        return ' '
    elif (spage==''):
        return ' '
    else:
        return ', '