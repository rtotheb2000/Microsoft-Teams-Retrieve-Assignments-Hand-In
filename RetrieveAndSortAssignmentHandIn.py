import os
from pathlib import Path
import pandas as pd
pd.set_option('display.max_rows', 999)
import img2pdf
from PIL import Image
import shutil
import time
from PyPDF2 import PdfFileMerger
import subprocess

debug = False

v='0.1 (beta)'

problems = ['Abgaben können übersehen werden, wenn sie als Link abgegebenen wurden.',
            'Symbolische Links können in Windows nur mit Administratorrechten erstellt werden (os.symlink -> \'OSError: symbolic link privilege not held\')']

# Information zur Version und bestehenden Problemen
print('\n' + '#'*77 + '\n')
print('Laden und Umbenennen von Abgaben der SchülerInnen von OneDrive (Version {})'.format(v))
print('\nListe von bekannten Problemen:\n')
lines = ''
for problem in problems:
    lines += '\n\t- {}\n'.format(problem)
print(lines)
print('#'*77 + '\n')

# Pfad der Datein von OneDrive
root_path = os.path.join(os.path.expanduser('~'),'Sophie Charlotte Gymnasium')
print('\nSuche Abgaben in {}\n'.format(root_path))

# Definition einer Klasse für abgegebene Dateien
class HandedInFile:
    def __init__(self, path):
        '''
        Die folgende Ordnerstruktur wird erwartet:

        "Kursname" + " - Aufgaben der Schüler"
            |
            -> "Gesendete Dateien"
                |
                -> "Name des Schülers" (Sonderzeichen sind enthalten)
                    |
                    -> "Titel des Aufgabe"
                        |
                        -> "Version x"
                            |
                            -> * (Dateiname der abgegebenen Datei)

        '''
        self.path = path
        self.course = (path.split(' - Aufgaben der Schüler')[0]).split('\\')[-1]
        self.student = path.split('\\')[-4]
        self.firstName = (path.split('\\')[-4]).split(' ')[0]
        self.lastName = (path.split('\\')[-4]).split(' ')[-1]
        self.assignment = path.split('\\')[-3]
        self.version = int((path.split('\\')[-2]).split('Version ')[-1])
        self.origFilename = path.split('\\')[-1]
        self.lastModified = os.path.getmtime(path)
        self.fileExtension = os.path.splitext(self.path)[-1]
        subject_dict = {
        'bi ':  'Biologie',
        'bb ':  'Biologie bilingual',
        'ch ':  'Chemie',
        'de ':  'Deutsch',
        'ds ':  'Darstellendes Spiel',
        'ek ':  'Erdkunde',
        'en ':  'Englisch',
        'frz ': 'Französisch',
        'ge ':  'Geschichte',
        'in ':  'Informatik',
        'ku ':  'Kunst',
        'ma ':  'Mathematik',
        'mu ':  'Musik',
        'pb ':  'Politische Bildung',
        'ph ':  'Physik',
        # this is out of place and only necessary because my naming was stupid
        'phy ': 'Physik',
        'phil ':'Philosophie',
        'pw ':  'Politische Weltkunde',
        'spa ': 'Spanisch'
        }
        list_of_keys = list(subject_dict.keys())
        for key in list_of_keys:
            subject_dict[key[:-1]+'-'] = subject_dict[key]
        test_string = self.course + ' ' + self.assignment
        test_list = subject_dict.keys()
        result = any(ele in test_string.lower() for ele in test_list)
        if result == True:
            self.subject =  [subject_dict[ele] for ele in test_list if ele in test_string.lower()][0]
        else:
            self.subject = "Undefiniert"
        filetype_dict = {
        '.pdf':  'Pdf',
        '.docx': 'Word',
        '.xlsx': 'Excel',
        '.pptx': 'PowerPoint',
        '.doc':  'Word',
        '.xls':  'Excel',
        '.ppt':  'PowerPoint',
        '.pages': 'Pages',
        '.jpg':  'Bild',
        '.jpeg': 'Bild',
        '.png':  'Bild',
        '.tiff': 'Bild',
        '.psd':  'Bild',
        '.eps':  'Bild',
        '.raw':  'Bild',
        '.heic': 'Bild'
        }
        try:
            self.filetype = filetype_dict[self.fileExtension.lower()]
        except:
            self.filetype = 'Unerkannt'
        self.newFilename = 'v{0} {1} {2}'.format(self.version, self.lastName, self.assignment)[:39-len(self.fileExtension)] + ' xxxxx' + '{}'.format(self.fileExtension)

    def DataFrame(self):
        input = [self.course, self.student, self.assignment, self.version, self.firstName, self.lastName, self.origFilename, self.lastModified, self.newFilename, self.filetype, self.subject, self.path]
        cols_dict = {
        0: 'Kurs',
        1: 'Student',
        2: 'Aufgabe',
        3: 'Version',
        4: 'Vorname',
        5: 'Nachname',
        6: 'Dateiname',
        7: 'zul. geändert',
        8: 'Dateiname intern',
        9: 'Dateityp',
        10: 'Fach',
        11: 'Dateipfad'
        }
        cols = [cols_dict[i] for i in range(len(input))]
        df = pd.DataFrame([input], columns=cols)
        return df

def FullDataFrame(root_path=root_path, selected=['Mathematik','Physik']):
    df = pd.DataFrame()
    for i in os.walk(root_path, topdown=True, onerror=None, followlinks=False):
        if 'Version ' in i[0]:
            for j in i[2]:
                file_path = HandedInFile(os.path.join(i[0],j))
                df = df.append(HandedInFile(os.path.join(i[0],j)).DataFrame())
    # drop subjects other than selected ones
    df = df.loc[(df['Fach']=='Mathematik')|(df['Fach']=='Physik')]
    # detect unneeded versions
    df_temp = df.set_index(['Kurs', 'Nachname', 'Aufgabe', 'Version']).sort_index()
    for mi in df_temp.index.unique():
        for i in range(mi[-1]):
            mi_temp = (mi[0], mi[1], mi[2], i)
            if mi_temp in df_temp.index:
                df_temp.drop(labels=mi_temp, inplace=True)
    df = df_temp.reset_index()
    df.index = range(len(df))
    for i in df.index:
        df.loc[i, 'Dateiname intern'] = df.loc[i, 'Dateiname intern'].replace('xxxxx', '{:05d}'.format(i))
    return df

def CreateFolderStructure(root_path=root_path, destination_path=os.path.join(os.path.expanduser('~'),'Schülerprodukte')):
    df = FullDataFrame(root_path=root_path)
    for Kurs in df.Kurs.unique():
        print('\nEs wird folgender Ordner erstellt, falls nicht vorhanden:\t{}'.format(os.path.join(destination_path, Kurs)))
        Path(os.path.join(destination_path, Kurs)).mkdir(parents=True, exist_ok=True)

def merger(output_path, input_paths):
    pdf_merger = PdfFileMerger()
    file_handles = []

    for path in input_paths:
        pdf_merger.append(path)

    with open(output_path, 'wb') as fileobj:
        pdf_merger.write(fileobj)

def DetectCoherentImages(root_path=root_path, destination_path=os.path.join(os.path.expanduser('~'),'Schülerprodukte')):
    df = FullDataFrame(root_path=root_path).set_index(['Kurs', 'Nachname', 'Aufgabe', 'Version']).sort_index()
    # Selecting only pictures (img2pdf can only handle JPEG, JPEG2000, PNG (non-interlaced), TIFF (CCITT Group 4))
    dfBild = df.loc[df['Dateityp']=='Bild']
    if debug:
        print(dfBild)
    for mi in dfBild.index.unique():
        if debug:
            print(mi)
        df_temp = dfBild.loc[mi,:].reset_index()
        imagelist = []
        new_filename = ''
        for index, row in df_temp.iterrows():
            if os.path.exists(df_temp.loc[index, 'Dateipfad']):
                im  = Image.open(df_temp.loc[index, 'Dateipfad'])
                imc = im.convert('RGB')
                imagelist.append(imc)
            else:
                raise ValueError('{} does not exist.'.format(df_temp.loc[index, 'Dateipfad']))
        new_filename = os.path.splitext(os.path.join(df_temp.loc[index, 'Dateiname intern']))[0] + '.pdf'
        Kurs = df_temp.loc[index, 'Kurs']
        writeTo = os.path.join(destination_path, Kurs, new_filename)
        if debug:
            print('type(imagelist)',type(imagelist))
            print('type(imagelist[0])',type(imagelist[0]))
        if not os.path.exists(writeTo):
            if len(imagelist)>1:
                imagelist[0].save(writeTo,save_all=True, append_images=imagelist[1:])
            else:
                imagelist[0].save(writeTo,save_all=True)
            print('\n[BILD]:\tFolgende zusammengeführte Datei wurde gespeichert:\t{}'.format(writeTo))
        else:
            print('\n[BILD]:\tFolgende Datei wurde übersprungen:\t\t\t{}'.format(writeTo))
    # Selecting only Pdfs
    dfPDF = df.loc[df['Dateityp']=='Pdf']
    for mi in dfPDF.index.unique():
        df_temp = dfPDF.loc[mi,:].reset_index()
        if len(df_temp)>1:
            df_temp = dfPDF.loc[mi,:].reset_index()
            pathslist = []
            for index, row in df_temp.iterrows():
                if os.path.exists(df_temp.loc[index, 'Dateipfad']):
                    pathslist.append(df_temp.loc[index, 'Dateipfad'])
                else:
                    raise ValueError('{} does not exist.'.format(df_temp.loc[index, 'Dateipfad']))
            new_filename = os.path.join(df_temp.loc[index, 'Dateiname intern'])
            Kurs = df_temp.loc[index, 'Kurs']
            writeTo = os.path.join(destination_path, Kurs, new_filename)
            if not os.path.exists(writeTo):
                merger(writeTo, pathslist)
                print('\n[PDF]:\tFolgende zusammengeführte Datei wurde gespeichert:\t{}'.format(writeTo))
            else:
                print('\n[PDF]:\tFolgende Datei wurde übersprungen:\t\t\t{}'.format(writeTo))
        else:
            df_temp = dfPDF.loc[mi,:]
            if debug:
                print(df_temp)
                print(mi)
            src = df_temp.loc[mi, 'Dateipfad']
            dst =  os.path.join(destination_path, Kurs, df_temp.loc[mi, 'Dateiname intern'])
            if not os.path.exists(dst):
                shutil.copy2(src, dst)
                print('\n[PDF]:\tFolgende Datei wurde als Kopie erstellt:\t\t{}'.format(writeTo))
            else:
                print('\n[PDF]:\tFolgende Datei wurde übersprungen:\t\t\t{}'.format(writeTo))
    # Selecting Office 365 filetypes
    dfO365 = df.loc[(df['Dateityp']=='Word')|(df['Dateityp']=='Excel')|(df['Dateityp']=='PowerPoint')]
    if debug:
        print(dfO365)
    df_temp = dfO365.reset_index()
    for index in df_temp.index:
        Kurs = df_temp.loc[index, 'Kurs']
        dst = os.path.join(destination_path, Kurs, df_temp.loc[index, 'Dateiname intern'])
        if not os.path.exists(dst):
            shutil.copy2(src, dst)
            print('\n[O365]:\tFolgende Datei wurde als Kopie erstellt:\t\t{}'.format(writeTo))
        else:
            print('\n[O365]:\tFolgende Datei wurde übersprungen:\t\t\t{}'.format(writeTo))
    # Selecting all other filetypes
    dfOTHER = df.loc[(df['Dateityp']!='Bild')&(df['Dateityp']!='Pdf')&(df['Dateityp']!='Word')&(df['Dateityp']!='Excel')&(df['Dateityp']!='PowerPoint')]
    if debug:
        print(dfOTHER)
    df_temp = dfOTHER.reset_index()
    for index in df_temp.index:
        Kurs = df_temp.loc[index, 'Kurs']
        dst = os.path.join(destination_path, Kurs, df_temp.loc[index, 'Dateiname intern'])
        if not os.path.exists(dst):
            shutil.copy2(src, dst)
            print('\n[OTHER]:\tFolgende Datei wurde als Kopie erstellt:\t\t{}'.format(writeTo))
        else:
            print('\n[OTHER]:\tFolgende Datei wurde übersprungen:\t\t\t{}'.format(writeTo))


def main(root_path=root_path, destination_path=os.path.join(os.path.expanduser('~'),'Schülerprodukte')):
    CreateFolderStructure(root_path=root_path)
    DetectCoherentImages(root_path=root_path, destination_path=destination_path)

main()
