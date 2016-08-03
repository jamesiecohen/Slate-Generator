#!/usr/bin/python
# -*- coding: utf-8 -*-

from PIL import Image, ImageDraw, ImageFont
import os
import subprocess
import sys
import openpyxl
import datetime

ffmpeg_path = '/usr/local/bin/ffmpeg'
spreadsheet_directory = '/Users/edit08/Desktop/WF_Slates/00_drop_spreadsheet_here'
source_directory = '/Users/edit08/Desktop/WF_Slates/01_pngs'
compressed_directory = '/Users/edit08/Desktop/WF_Slates/02_compressed'
done_directory = '/Users/edit08/Desktop/WF_Slates/03_done'
countdown = '/Users/edit08/Desktop/WF_Slates/04_scripts/Countdown_2015_w_alpha.mov'
slate_starter = '/Users/edit08/Desktop/WF_Slates/04_scripts/slate_starter.tif'
source_extension = '.png'
destination_extension = '.mov'
rate = '23.976023976023978'
png_slate_list = []
# Header row contents
header_row_cell_1 = 'Joint Jobs::Agency'


#### FONTS
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

def open_spreadsheet():
    os.chdir(spreadsheet_directory)
    spreadsheet = os.listdir(os.getcwd())
    spreadsheet_file = spreadsheet[0]
    wb = openpyxl.load_workbook(spreadsheet[0], read_only = True, data_only = True)
    ws = wb['Sheet1']
    return ws, spreadsheet_file

def make_pngs():
    worksheet, ss_file = open_spreadsheet()
    for row in worksheet.rows:
        #skip header
        if row[0].value == header_row_cell_1:
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
                d.text((left_margin,460), slate_contents[3], font=fnt3, fill=gray)
            elif w2 < 1150:
                d.text((left_margin,460), slate_contents[3], font=fnt3b, fill=gray)
            elif w3 < 1150:
                d.text((left_margin,460), slate_contents[3], font=fnt3c, fill=gray)
            elif w4 < 1150:
                d.text((left_margin,460), slate_contents[3], font=fnt3d, fill=gray)
            else:
                d.text((left_margin,460), slate_contents[3], font=fnt3e, fill=gray)
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
            png_out_name = slate_contents[2] + '.png'
            png_out_name = os.path.join(source_directory,png_out_name)
            #Save file
            txt.save(png_out_name)
            #txt.close()
    #move spreadsheet to done done directory
    ss_file_source_path = os.path.join(spreadsheet_directory, ss_file)
    ss_file_dest_path = os.path.join(done_directory, ss_file)
    os.rename(ss_file_source_path, ss_file_dest_path)

def make_png_slate_list():
    os.chdir(source_directory)
    files = os.listdir(os.getcwd())
    for i in files:
        if i.endswith(source_extension):
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



#open_spreadsheet()
make_pngs()
make_png_slate_list()
encode()
