"""
migrate_journey.py
~~~~~~~~~~~~~~~~~

This script parses a directory containg json and jpeg files from a backup of the Journey Two App, prepare each post
and send them to a specific section in a OneNote notebook
- shrink images to specific width
- converts to UTF-8 and handles Umlaute
- creates a valid OneNote Page with references to the images

Copyright (c) 2016 Coffeemug13. All rights reserved. Licensed under the MIT license.
See LICENSE in the project root for license information.
"""

import json
import os
import traceback
from datetime import datetime
import markdown
import pyonenote

try:
    from config import client_id, client_secret, code, token, rtoken
except:
    pass

# This information is obtained upon registration of a new Outlook Application
# The values are just for information and not valid
# client_id = "cda3ffaa-2345-a122-3454-adadc556e7bf"
# client_secret = "AABfsafd6Q5d1VZmJQNsdac"
# code = "AcD5bcf9a-0fef-ca3a-1a3a-9v4543388572"
# token = "EAFSDTBRB$/UGCCXc8wU/zFu9QnLdZXy+YnElFkAAW......"
# rtoken = "MCKKgf55PCiM2aACbIYads*sdsa%*PWYNj436348v......"


def update_config(client_id, client_secret, code, token, rtoken):
    with open("config.py", mode="w") as credentials:
        credentials.write("'''autogenerated file at, just update client* and code'''\n")
        credentials.write("client_id = '{0}'\n".format(client_id))
        credentials.write("client_secret = '{0}'\n".format(client_secret))
        credentials.write("code = '{0}'\n".format(code))
        credentials.write("token = '{0}'\n".format(token))
        credentials.write("rtoken = '{0}'\n".format(rtoken))


# file exists

def resize_images(files: dict):
    from PIL import Image
    from resizeimage import resizeimage
    images = {}
    count = 0
    for name, file in files.items():
        try:
            with open(file, 'rb') as img:
                with Image.open(img) as tmpImg:
                    if tmpImg.width > 600:
                        tmp = resizeimage.resize_width(tmpImg, 600)
                    else:
                        tmp = tmpImg
                    tmpName = 'tmp{}.jpg'.format(count)
                    tmp.save(tmpName, 'JPEG')
                    images[name] = tmpName
                    count += 1
        except Exception as e:
            # Qucik hack, ignore images which can't be converted for whatever reason
            print(e)
    return images


def parse_page(file: str):
    conversion = {
        ord(u'ä'): u'&auml;',
        ord(u'ö'): u'&ouml;',
        ord(u'ü'): u'&uuml;',
        ord(u'Ä'): u'&Auml;',
        ord(u'Ö'): u'&Ouml;',
        ord(u'Ü'): u'&Uuml;',
        ord(u'ß'): u'&szlig;',
    }
    """Parse the JSON file from a Journey Page"""
    with open(file, 'r', encoding='utf8') as f:
        page = json.load(f)
        result = {}
        result['id'] = page.get('id')
        t = int(float(page['date_journal']) / 1000)
        td = datetime.fromtimestamp(t)
        result['created'] = td.isoformat()
        result['title'] =  td.date().isoformat() + " " + page['text'][:15] + ".."
        result['content'] = ''
        result['photos'] = {}
        for photo in page.get('photos'):
            result['photos'][photo] = ("journey/" + photo)
            result['content'] += '<p><img src="name:{0}"/></p>'.format(photo)
        x = page['text'].translate(conversion)
        result['content'] += markdown.markdown(page['text']).translate(conversion)

        tags = []
        for tag in page.get('tags'):
            tags.append("#" + tag)
        if len(tags) > 0:
            result['tags'] = "; ".join(tags)
        else:
            result['tags'] = ""
        result['address'] = page.get('address').translate(conversion)
        result['lat'] = page.get('lat')
        result['lon'] = page.get('lon')
        return result


def start(onenote):
    count = 0
    for dirName, subdirList, fileList in os.walk('journey'):
        for file in fileList:
            if (file.endswith('.json')):
                if count< 0:
                    # in case the import of a specific post breaks, you can restart at a specific count after cleaning
                    # some minor error in the page, e.g. special UTF-8 characters
                    print('skip file ({0}): {1}'.format(count, file))
                else:
                    # read the post and images and send to onenote
                    print('parsing file ({0}): {1}'.format(count,file))
                    post = parse_page('journey/' + file)
                    images = resize_images(post['photos'])
                    files = []
                    try:
                        for iName, iFile in images.items():
                            files.append((iName, (iName, open(iFile, 'rb'), 'image/jpeg', {'Content-Encoding': 'base64'})))

                        # files = [('dog.jpg',('d.jpg', img, 'image/jpeg', {'Content-Encoding':'base64'}))]
                        # result = onenote.list_notebooks()
                        content = post['content']
                        if post['lat'] and post['lon']:
                            content += '<p>Adresse: {2} <a href="http://www.openstreetmap.org/?mlat={0}&mlon={1}#map=15/{0}/{1}">Link</a></p>'.format(post['lat'],post['lon'],post['address'])
                        else:
                            content += "<p>Adresse: {0}</p>".format(post['address'])
                        content += "<p>{0}</p>".format(post['tags'])
                        result = onenote.post_page("0-C34ECF5A07994ACF!252063",
                                                   post['created'],
                                                   post['title'],
                                                   content, files)
                        print('posted')
                    except Exception as e:
                        print('EXCEPTION at count ' + str(count))
                        raise e
                count += 1

# only start when execute as script
if __name__ == '__main__':
    if not client_id:
        update_config("your client_id", "your client secret", "", "", "")
        print("Please update the config.py file with client credentials")
        exit()
    if not code:
        import webbrowser

        url = pyonenote.OneNote.get_authorize_url(client_id)
        print("Please open this URL to authorize this client")
        print(url)
        webbrowser.open(url)
        exit()
    print('####################################\nOneNote Client started\n')
    onenote = pyonenote.OneNote(client_id, client_secret, code, token, rtoken)
    try:
        start(onenote)
    except Exception as e:
        traceback.print_exc()
    update_config(*onenote.get_credentials())
    print("actual credentials/token for client saved")
    print('\n####################################\nOneNote Client stopped')
