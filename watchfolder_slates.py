#!/usr/bin/python
from PIL import Image, ImageDraw, ImageFont
import os
import subprocess
import sys
import openpyxl
import datetime


#/opt/local/bin/ffmpeg -loop 1 -framerate 23.976023976023978 -i /Users/e10/Desktop/watch/04_scripts/YVZW6108H\ Better\ Together\ 24\ GB\ Offer\ Generic\ HD\ 30_SLATE.png -i  /Users/e10/Desktop/watch/04_scripts/Countdown_2015_w_alpha.mov -filter_complex overlay -vcodec prores_ks -profile:v 3 -t 00:00:07.01 /Users/e10/Desktop/watch/03_done/test.mov
ffmpeg_path = '/opt/local/bin/ffmpeg'
excel_directory = '/Volumes/FIN_SHARE/0-FINI_JOBS/0000_ENGINEERING/WF_Slates/00_Slatenator'
source_directory = '/Volumes/FIN_SHARE/0-FINI_JOBS/0000_ENGINEERING/WF_Slates/01_drop_here'
compressed_directory = '/Volumes/FIN_SHARE/0-FINI_JOBS/0000_ENGINEERING/WF_Slates/02_compressed'
done_directory = '/Volumes/FIN_SHARE/0-FINI_JOBS/0000_ENGINEERING/WF_Slates/03_done'
countdown = '/Volumes/FIN_SHARE/0-FINI_JOBS/0000_ENGINEERING/WF_Slates/04_scripts/Countdown_2015_w_alpha.mov'
slate_starter = '/Volumes/FIN_SHARE/0-FINI_JOBS/0000_ENGINEERING/WF_Slates/04_scripts/slate_starter.tif'
source_extension = '.png'
destination_extension = '.mov'
excel_extension = '.xlsx'
rate = '23.976023976023978'


#master font to use, plus colors
font = '/System/Library/Fonts/HelveticaNeueDeskInterface.ttc'
gray = (200,200,200)
red = (200,77,82)
left_margin = 385
#set size and Bold, italic, etc
fnt1 = ImageFont.truetype(font, 55, 0)
fnt2 = ImageFont.truetype(font, 55, 1)
fnt3 = ImageFont.truetype(font, 55, 2)
fnt3b = ImageFont.truetype(font, 45, 2)
fnt3c = ImageFont.truetype(font, 35, 2)
fnt3d = ImageFont.truetype(font, 25, 2)
fnt3e = ImageFont.truetype(font, 15, 2)
fnt4 = ImageFont.truetype(font, 35, 1)
fnt5 = ImageFont.truetype(font, 25, 0)


png_slate_list = []
excel_list = []


def make_excel_list():
    os.chdir(excel_directory)
    files = os.listdir(os.getcwd())
    for i in files:
        if i.endswith(excel_extension):
            excel_list.append(i)

def open_excel():
    for i in excel_list:
        wb = openpyxl.load_workbook(i,read_only = True, data_only = True)
        ws = wb['Sheet1']
        return ws

def generate_slate_pngs(ws):
    for row in ws.rows:
        #skip header
        if row[0].value == 'Joint Jobs::Agency':
            pass
        else:
            #start blank black file
            txt2 = Image.open(slate_starter)
            txt = Image.new('RGB', (1920,1080), (0,0,0))
            txt.paste(txt2)
            #drawing instance
            d = ImageDraw.Draw(txt)
            #make blank list to store slate_contents
            slate_contents = []
            for cell in row:
                #format datetime to month/day/year
                if type(cell.value) == datetime.datetime:
                    new_date = cell.value.strftime('%m/%d/%Y')
                    slate_contents.append(new_date)
                #turn empty cells into blank strings
                elif cell.value is None:
                    new_value = ''
                    slate_contents.append(new_value)
                #write cell to list
                else:
                    slate_contents.append(cell.value)
            #Agency
            d.text((left_margin,200), slate_contents[0], font=fnt1, fill=gray)
            #Client
            d.text((left_margin,270), slate_contents[1], font=fnt1, fill=gray)
            #ISCI
            d.text((left_margin,360), slate_contents[2], font=fnt2, fill=red)
            #Spot Title / check to see if title is too long, scale down
            w,h = fnt3.getsize(slate_contents[3])
            w2,h2 = fnt3b.getsize(slate_contents[3])
            w3,h3 = fnt3c.getsize(slate_contents[3])
            w4,h4 = fnt3d.getsize(slate_contents[3])
            if w < 1150:
                d.text((left_margin,460), '{} {}'.format(slate_contents[3], ' '), font=fnt3, fill=gray)
            elif w2 < 1150:
                d.text((left_margin,460), '{} {}'.format(slate_contents[3], ' '), font=fnt3b, fill=gray)
            elif w3 < 1150:
                d.text((left_margin,460), '{} {}'.format(slate_contents[3], ' '), font=fnt3c, fill=gray)
            elif w4 < 1150:
                d.text((left_margin,460), '{} {}'.format(slate_contents[3], ' '), font=fnt3d, fill=gray)
            else:
                d.text((left_margin,460), '{} {}'.format(slate_contents[3], ' '), font=fnt3e, fill=gray)
            # TRT
            d.text((left_margin,540), slate_contents[4], font=fnt4, fill=gray)
            #audio
            d.text((left_margin,660), slate_contents[5], font=fnt4, fill=gray)
            #date
            d.text((left_margin,727), slate_contents[6], font=fnt4, fill=gray)
            #comments (aka NFA)
            d.text((left_margin,820), slate_contents[7], font=fnt4, fill=red)
            #legal / copyright
            d.text((left_margin,898), slate_contents[8], font=fnt5, fill=gray)
            #name file after ISCI
            outname = '{}{}'.format(slate_contents[2], '.png')
            outname_with_path = os.path.join(source_directory, outname)
            #Save file
            txt.save(outname_with_path)
            txt.close()

def move_excel_doc_to_done():
    for i in excel_list:
        excel_source = os.path.join(excel_directory, i)
        excel_done = os.path.join(done_directory, i)
        os.rename(excel_source, excel_done)



def make_png_slate_list():
    os.chdir(source_directory)
    files = os.listdir(os.getcwd())
    for i in files:
        png_slate_list.append(i)




def encode():
    for i in png_slate_list:
        print i
        ff_source = os.path.join(source_directory, i)
        done_path = os.path.join(done_directory, i)
        out_name_base = ''.join(i.split('.')[:-1])
        out_name = '{0}{1}'.format(out_name_base, destination_extension)
        ff_destination = os.path.join(compressed_directory, out_name)
        if i.endswith(source_extension):
            subprocess.call([ffmpeg_path, '-loop', '1', '-framerate', rate, '-i',
                            ff_source, '-i', countdown, '-filter_complex',
                            'overlay', '-vcodec', 'prores_ks', '-profile:v', '3',
                            '-t', '00:00:07.01', ff_destination])
            os.rename(ff_source, done_path)
        else:
            os.rename(ff_source, done_path)

make_excel_list()
generate_slate_pngs(open_excel())
move_excel_doc_to_done()
make_png_slate_list()
encode()
