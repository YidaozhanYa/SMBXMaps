#迁移工具
import json
import leancloud
leancloud.init("wgcr3xHDSmfiaOJReHtlqD9z-MdYXbMMI", "7vDim8MYqChNNgt2D8NkFjtP")
DB = leancloud.Object.extend('DB')
lists=json.loads(open("qwq.json","r").read())['data']['files']
for singleitem in lists:
    if "[" in singleitem['name']:
        version=singleitem['name'].split("[")[1].split("]")[0]
        name=singleitem['name'].replace("."+singleitem['name'].split('.')[-1],"").replace("["+version+"] ","")
    else:
        version="Unknown"
        name=singleitem['name'].replace("."+singleitem['name'].split('.')[-1],"")
    newitem = DB()
    newitem.set('name', name)
    newitem.set('rel', "Initial")
    newitem.set('version', version)
    newitem.set('maker', "Unknown")
    newitem.set('desc', "")
    newitem.set('puburl', "")
    newitem.set('status', "public")
    newitem.set('repofolder', "S")
    newitem.set('ext', singleitem['name'].split('.')[-1])
    newitem.save()
    print(name)