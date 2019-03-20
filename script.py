from bs4 import BeautifulSoup as BS
import sys
import os
import re

# Get the current working directory

parent_dir = os.getcwd()
cwd = sys.argv[1]
file_MS = sys.argv[1]
dir_path = file_MS[:file_MS.rfind('\\')]
file_name = file_MS[file_MS.rfind('\\')+1:]

file_Libre = dir_path + '\\___FILES___\\' + file_name

# Store BeautifulSoup Objects of the MSOffice and Libre files
file_1 = open(file_MS, mode = 'r', encoding = 'utf-8')
file_2 = open(file_Libre, mode='r', encoding = 'utf-8')
msoffice = BS(file_1, features = 'html.parser')
libre = BS(file_2, features = 'html.parser')
file_1.close()
file_2.close()

# Creating and opening a log file
log_file = parent_dir + '\\conversion.log'  
log_file = open(log_file, mode = 'a', encoding = 'utf-8')

# List to store paths of equation images that will be deleted later
to_be_deleted = []

try:
    # Find the code to be replaced and the code to be replaced with
    original = msoffice.findAll('span', style = re.compile('Liberation|line-height'))
    replacement = libre.findAll('math')

    var = 0
    for i in range(len(original)):
        if original[i].attrs['style'].find('line-height') != -1:
            if original[i].attrs['style'].find('Calibri') == -1:
                del original[var]
                var -= 1
        var += 1

    # Remove the images of equations
    for j in range(len(replacement)):
        # Maintaining a list of images that need to be deleted
        del_img = original[j].find('img')['src']
        if '%20' in del_img:
            del_img = del_img.replace('%20', ' ')
        if '/' in del_img:
            del_img = del_img.replace('/','\\')
        to_be_deleted.append(dir_path + '\\' + del_img)
        original[j].clear()
        original[j].insert(0, BS(str(replacement[j]), features = 'html.parser'))

    # Remove the old HTML File
    os.remove(file_MS)
    os.remove(file_Libre)

    # Remove the equation images
    for j in set(to_be_deleted):
        os.remove(j)

    # Create modified HTML File with same name
    with open(file_MS, mode='w', encoding='utf-8') as file:
        msoffice.prettify()
        file.write(str(msoffice).replace('display="block"','display="flex"'))


    #Remove the ___FILES___ directory created to store Libre files
    try:
        os.rmdir(dir_path + '\\___FILES___')
    except Exception as e:
        pass

    print('Converted ', file_MS)
    log_file.write('\nConverted ' + file_MS)

except Exception as e:
    if hasattr(e, 'message'):
        log_file.write('\n' + e.message)
        log_file.write('***************' + file_MS + ' not converted ************** ')

log_file.close()
