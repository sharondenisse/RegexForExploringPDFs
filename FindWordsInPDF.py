#!/usr/bin/env python
# coding: utf-8

# In[2]:


import PyPDF2, os, re, sys, io
from io import StringIO


# Let's change the directory to the location where all the pdf files are located.
# Ideally, it is the location where this notebook is also sitting.  The work directory
# is the location where the text files where be created.

# In[3]:


get_ipython().run_line_magic('cd', 'C:\\Users\\Skenny02\\Desktop\\MapsForMatt')
workDir = (r'C:\Users\Skenny02\Desktop')


# In[4]:


pdfFiles_SDK = []
pdfName_SDK = []
directorySDK = [r'C:\Users\Skenny02\Desktop\MapsForMatt']

# If you would like to look inside many directories, you could add them here.
# This is only permissible if there is a small amount of pdfs in each directory
# CAUTION: THE CODE BELOW IS RECURSIVE


# In[8]:


i = 0
while i < ry:
    for root, dirs, files in os.walk(directorySDK[i]):
        for name in files:
            if name.endswith('.pdf'):
                pdfFiles_SDK.append(os.path.join(root, name))
                pdfName_SDK.append(os.path.join(name))
    i = i + 1


# In[12]:


for files in pdfName_SDK:
    print(files)


# In[13]:


print(len(pdfFiles_SDK))
#print(len(pdfName_SDK))


# In[24]:


tex = []
encrypted = []
textsOneP = []


# In[25]:


import shutil


# In[26]:


get_ipython().run_line_magic('pwd', '')


# In[27]:


# Create a directory where to place the encrpyted files
get_ipython().system('mkdir C:\\Users\\Skenny02\\Desktop\\xxxpdfs')


# In[28]:


# Confirm that the folder was created
get_ipython().system('dir C:\\Users\\Skenny02\\Desktop\\xxxpdfs')


# In[29]:


xxxpdfs = r'C:\Users\Skenny02\Desktop\xxxpdfs'


# In[30]:


i = 0
while i < len(pdfFiles_SDK):
    
    what = pdfFiles_SDK[i]
    when = pdfName_SDK[i]
    print(what)
    pdfFileObj =  open (what, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    
    if pdfReader.isEncrypted:
        pdfFileObj.close()
        print("Encrypted")
        encrypted.append(what)
        shutil.copy(what,xxxpdfs)
        print(what)
        pdfFiles_SDK.remove(what)
        pdfName_SDK.remove(when)
        
    else:
        pageThing = pdfReader.getPage(0)
        pageWord = pageThing.extractText()
        #cc = workDir  + '\\' + ff + '.txt'
        cc = workDir  + '\\' + when + str(i) + '.txt'
        print(cc)
        wfile = io.open(cc, "w", encoding="utf-8") 
        wfile.write(pageWord)
        wfile.close()
        textsOneP.append(cc)
    i = i + 1


# In[31]:


print(len(pdfFiles_SDK), len(textsOneP))
textsOneP


# In[37]:


# Searching for string "Map xxxx", where x are numbers from 0 to 9
z = 0
mapsName = []
mapsNumber = []
mapsLocation = []
while z < (len(pdfFiles_SDK)):
    regFile = textsOneP[z]
    regPDF = pdfFiles_SDK[z]  
    whatFile = io.open(regFile, 'r', encoding="utf-8")
    openFile = whatFile.read()
    #lineSpace = re.findall('[M]{1}[a-z]{2}\s[0-9]{4}', openFile)
    lineSpace = re.findall('S{1}[a-z]{4}[n]\s', openFile)
    mm = int(len(lineSpace))
    if mm != 0:
        print(openFile)
        print("The pdf file called   ", regPDF, "    contains the phrase    ", lineSpace[0]  )
        mapsLocation.append(pdfFiles_SDK[z])
        #apsLocation.append(regPDF)
        #mapsName.append(regFile)
        mapsNumber.append(lineSpace[0])
        
    z = z + 1
    whatFile.close()


# If searching for string "Mapxxxx", where there is no space before the numbers
# use the following regular expression 

#line = re.findall('[A-Za-z]{3}[\d]{4}', openFile)


# In[38]:


mapsLocation


# In[ ]:





# Below, we will create an Excel sheet names "For_SDK" in the desktop. It will have only one sheet, named "pdfFileName"

# In[39]:


from openpyxl import Workbook
titleSheet = "For_SDK"
filepath = workDir + "\\" + titleSheet + ".xlsx"
wBook = Workbook()

wSheet = wBook.active
wSheet['A1']= "pdfFileName"

wSheet.title = titleSheet
print (wSheet)
print(filepath)


# Below we are creating the titles for the first and second column, which will be placed in the first row.

# In[40]:


wSheet['A1']= "regFile"
wSheet['B1']= "DocPhrase"


# In[41]:


i = 0
rr = 2

while i < len(mapsLocation):
    j = str(rr)
    vv = mapsLocation[i]
    print(wSheet['A'+j])
    wSheet['A'+j] = vv
    print(wSheet['A'+j].value)
    i = i+1
    rr = rr + 1


# In[42]:


i = 0
rr = 2

while i < len(mapsLocation):
    j = str(rr)
    vv = mapsNumber[i]
    print(wSheet['B'+j])
    wSheet['B'+j] = vv
    print(wSheet['B'+j].value)
    i = i+1
    rr = rr + 1


# In[43]:


wBook.save(filepath)


# Below is a list of all the variables we used in this code

# In[44]:


get_ipython().run_line_magic('whos', '')


# In[ ]:





# In[ ]:




