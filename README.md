# Microsoft Teams Retrieve Assignments Hand In

## First steps

Download the OneDrive App <a href="https://upload.wikimedia.org/wikipedia/commons/thumb/d/d3/Microsoft_Office_OneDrive_%282018%E2%80%93present%29.svg/200px-Microsoft_Office_OneDrive_%282018%E2%80%93present%29.svg.png"><img src="https://upload.wikimedia.org/wikipedia/commons/thumb/d/d3/Microsoft_Office_OneDrive_%282018%E2%80%93present%29.svg/200px-Microsoft_Office_OneDrive_%282018%E2%80%93present%29.svg.png" width="30"/></a> and make sure you are logged in with your organizations OneDrive Account.

Navigate to the SharePoint folder <a href="https://upload.wikimedia.org/wikipedia/commons/thumb/3/31/Microsoft_Office_SharePoint_%282018%E2%80%93present%29.svg/200px-Microsoft_Office_SharePoint_%282018%E2%80%93present%29.svg.png"><img src="https://upload.wikimedia.org/wikipedia/commons/thumb/3/31/Microsoft_Office_SharePoint_%282018%E2%80%93present%29.svg/200px-Microsoft_Office_SharePoint_%282018%E2%80%93present%29.svg.png" width="30"/></a> of your organization and your class (which should be a closed team).

Navigate to "Websiteinhalte" and select "Aufgaben der Schüler".

Select "Synchronisieren" on top. The OneDrive App will automatically retrieve the folder and its subdirectories.

At this point you have the folder structure replicated on your computer.

This is where we can start using this script.

## Usage of the script

The `root_path` must be set such that it contains the all the team directories of type `"Kursname - Aufgaben der Schüler"`

The structure of subdirectories in SharePoint/OneDrive is as follows

`"Kursname - Aufgaben der Schüler"`

&nbsp;&nbsp;&nbsp;&#8627; `"Gesendete Dateien"`

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#8627; `"Name des Schülers"` (umlauts are included)

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#8627; `"Titel des Aufgabe"`

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#8627; `"Version x"`

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#8627; `"*"` (filename of the handed in assignment)

This subdirectory structure is used to create a `Padas Dataframe` which contains all Metadata information such as `student`, `assigment`, `course`, etc.

On this basis a new folder structure is created in `destination_path` where subdirectories for each `course` is created.

In this script only `courses` with the subject `"Mathematik"` or `"Physik"` are copied. However this behavior is easily adjusted by usage of the argument `selected` in `FullDataFrame(root_path=root_path, selected=['Mathematik','Physik'])`. Note that implemented subjects at this point are:

`'bi ':  'Biologie',`

`'bb ':  'Biologie bilingual',`

`'ch ':  'Chemie',`

`'de ':  'Deutsch',`

`'ds ':  'Darstellendes Spiel',`

`'ek ':  'Erdkunde',`

`'en ':  'Englisch',`

`'frz ': 'Französisch',`

`'ge ':  'Geschichte',`

`'in ':  'Informatik',`

`'ku ':  'Kunst',`

`'ma ':  'Mathematik',`

`'mu ':  'Musik',`

`'pb ':  'Politische Bildung',`

`'ph ':  'Physik',`

`# this is out of place and only necessary because my naming was stupid`

`'phy ': 'Physik',`

`'phil ':'Philosophie',`

`'pw ':  'Politische Weltkunde',`

`'spa ': 'Spanisch'`

Each handed in assignment is merged to a single file or simply copied to the respective `course` folder at `destionation_path`.

Multiple image files are merged to a single PDF-file by usage of `Pillow` in conjunction with `img2pdf`. Not all image filetypes are supported. To my knowledge `img2pdf` supports `JPEG`, `JPEG2000`, `PNG` (non-interlaced), `TIFF` (CCITT Group 4). Extensions like `HEIC`which come from Apple iPhones are probably not supported.

Single image files are converted to a single PDF-file.

At the current state this script will fail when it encounters an unsupported image filetype.

Multiple PDF-files are merged to a single PDF-file by usage of `PyPDF2`.

Single PDF-files are simply copied.

Word-, Excel- or PowerPoint-files are simply copied. A symbolic link would be nice here, because changes in the file are directly visible in Microsoft Teams <a href="https://upload.wikimedia.org/wikipedia/commons/thumb/c/c9/Microsoft_Office_Teams_%282018%E2%80%93present%29.svg/200px-Microsoft_Office_Teams_%282018%E2%80%93present%29.svg.png"><img src="https://upload.wikimedia.org/wikipedia/commons/thumb/c/c9/Microsoft_Office_Teams_%282018%E2%80%93present%29.svg/200px-Microsoft_Office_Teams_%282018%E2%80%93present%29.svg.png" width="30"/></a>. However, Windows 10 does not allow symbolic links without admin rights so I opted for simply copying the files.

All other filetypes are simply copied.

For each merged or copied file there will be a message that is has been created.

Copied files will have the following naming scheme:

`(version) (last name of the student) (name of the assignment) (some number to make sure there is no doubling of names)`

The end of the filename (without the extension) is removed in order to keep the the total number of characters equal to 45.

The output folder structure is

`"Kursname"`

&nbsp;&nbsp;&nbsp;&#8627; `"Aufgabe"`

![screenshot usage powershell](screenshot_powershell.png)
