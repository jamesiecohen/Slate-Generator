#!/usr/bin/python
# -*- coding: utf-8 -*-

from PIL import Image, ImageDraw, ImageFont
import time
import os
import subprocess
import sys
import openpyxl
import datetime
import re


#/opt/local/bin/ffmpeg -loop 1 -framerate 23.976023976023978 -i /Users/e10/Desktop/watch/04_scripts/YVZW6108H\ Better\ Together\ 24\ GB\ Offer\ Generic\ HD\ 30_SLATE.png -i  /Users/e10/Desktop/watch/04_scripts/Countdown_2015_w_alpha.mov -filter_complex overlay -vcodec prores_ks -profile:v 3 -t 00:00:07.01 /Users/e10/Desktop/watch/03_done/test.mov
ffmpeg_path = '/opt/local/bin/ffmpeg'
slate_compression_root_path = '/Volumes/FIN_SHARE/5-COMPRESSION/WF_Slates'
excel_directory = '{0}/00_drop_here'.format(slate_compression_root_path)
source_directory = '{0}/01_working_pngs'.format(slate_compression_root_path)
compressed_directory = '{0}/02_compressed'.format(slate_compression_root_path)
done_directory = '{0}/03_done'.format(slate_compression_root_path)
countdown = '{0}/04_scripts/Countdown_2015_w_alpha.mov'.format(slate_compression_root_path)
slate_starter = '{0}/04_scripts/slate_starter.tif'.format(slate_compression_root_path)
source_extension = '.png'
destination_extension = '.mov'
excel_extension = '.xlsx'
rate = '23.976023976023978'


#master font to use, plus colors
font = '/System/Library/Fonts/HelveticaNeueDeskInterface.ttc'
gray = (130,130,130)
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
        if not i.startswith('.'):
            if i.endswith(excel_extension):
                excel_list.append(i)

def check_if_excel_list_has_items():
    if len(excel_list) > 0:
        return True
    else:
        return False

def open_excel():
    for i in excel_list:
        print i
        wb = openpyxl.load_workbook(i,read_only = True, data_only = True)
        ws = wb['Sheet1']
        generate_slate_pngs(ws)
        #return ws

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
                    new_value = ' '
                    slate_contents.append(new_value)
                #write cell to list
                else:
                    #new_unicode_value = cell.value# maybe kill this
                    #new_unicode_value = new_unicode_value.encode('UTF-8')#maybe kill this
                    slate_contents.append(cell.value)
                    #slate_contents.append(new_unicode_value)
            if len(slate_contents) == 8:
                slate_contents.append(' ')
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
            pre_outname = '{0}_{1}_SLATE{2}'.format(slate_contents[2], slate_contents[3],'.png')
            regex = re.compile('[^a-zA-Z0-9 _.\-]')
            outname = regex.sub('', pre_outname)
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
        ff_source = os.path.join(source_directory, i)
        png_done_path = os.path.join(compressed_directory, i)
        out_name_base = ''.join(i.split('.')[:-1])
        out_name = '{0}{1}'.format(out_name_base, destination_extension)
        ff_destination = os.path.join(compressed_directory, out_name)
        if os.path.isfile(ff_destination):
            os.remove(ff_destination)
        if i.endswith(source_extension):
            subprocess.call([ffmpeg_path, '-loop', '1', '-framerate', rate, '-i',
                            ff_source, '-i', countdown, '-filter_complex',
                            'overlay', '-vcodec', 'prores_ks', '-profile:v', '3',
                            '-vendor', 'ap10', '-t', '00:00:07.01', ff_destination])
            os.rename(ff_source, png_done_path)
        else:
            os.rename(ff_source, png_done_path)


while True:
    excel_list = []
    png_slate_list = []
    make_excel_list()
    if check_if_excel_list_has_items() == True:
        print "Excel docs found"
        open_excel()
        move_excel_doc_to_done()
        make_png_slate_list()
        encode()
    else:
        print "No Excel docs found"
        pass
    time.sleep(30)
