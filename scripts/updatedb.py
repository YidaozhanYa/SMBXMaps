#迁移工具
import json
import leancloud ,requests
leancloud.init("wgcr3xHDSmfiaOJReHtlqD9z-MdYXbMMI", "7vDim8MYqChNNgt2D8NkFjtP")
DB = leancloud.Object.extend('DB')
prefix="K"
req=requests.get("https://pan.yidaozhan.ga/ali/SMBX地图仓库/"+prefix+"/?json").text
lists=json.loads(req)['list']
for singleitem in lists:
    if "[" in lists[singleitem]['name']:
        version=lists[singleitem]['name'].split("[")[1].split("]")[0]
        name=lists[singleitem]['name'].replace("."+lists[singleitem]['name'].split('.')[-1],"").replace("["+version+"] ","")
    else:
        version="Unknown"
        name=lists[singleitem]['name'].replace("."+lists[singleitem]['name'].split('.')[-1],"")
    newitem = DB()
    newitem.set('name', name)
    newitem.set('rel', "Initial")
    newitem.set('version', version)
    newitem.set('maker', "Unknown")
    newitem.set('desc', "")
    newitem.set('puburl', "")
    newitem.set('status', "public")
    newitem.set('repofolder', prefix)
    newitem.set('ext', lists[singleitem]['name'].split('.')[-1])
    newitem.save()
    print(name)