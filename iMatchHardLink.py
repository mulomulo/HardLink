import os
import win32file
stitch_folder = r"C:\Users\Horst\Pictures\Output\Stitch"
location = r"C:\Users\Horst\Documents\Image Databases\HardLinks"
stitch_folder = "%s/Stitch" %location
def HardLink(src, dst):
  """Hard link function

  Args:
    src: Source filename to link to.
    dst: Destination link name to create.

  Raises:
    OSError: The link could not be created.
  """
  try:
    win32file.CreateHardLink(dst,src)
  except win32file.error, msg:
    # Translate errors into standard OSError
    raise OSError(msg)

def MakeLinks(target_folder):
  l = []
  rFile = open(r"%s\link.txt" %(location), 'r')
  files = rFile.readlines()
  rFile.close()
  for file in files:
    d = {}
    f = file.split("\t")

    name = f[0].split("\\")[-1:][0][:-4]
    
    d = getLocation(f, d)
    
    categories = ""
    categories_l = f[1].split(',')
    for category in categories_l:
      cat = category.split(".")
      #cat.reverse()
      cat = "|".join(cat)
      categories += "<rdf:li>%s</rdf:li>\n" %cat 
    d.setdefault('filefull',f[0])
    d.setdefault('filename', name)
    d.setdefault('categories',categories)
    d.setdefault('rating',f[2])
    d.setdefault('colour',f[3].title().strip())
    
    specials = ['Stitch','HDR']
    is_special = False
    for special in specials:
      if target_folder.endswith(special):
        p = f[1].split("_types.%s." %special)[1]
        p = p.split(",")[0]
        p = p.replace(".","\\")
        target_folder = "%s/%s" %(location, special)
        is_special = True
    if not is_special:
      try:
        p = f[1].split("Collections.")[1]
        p = p.split(",")[0]
        p = p.replace(".","\\")
        p = "Collections/%s" %p
      except:
        p = ""
      
    path = "%s\%s" %(target_folder, p)
    print "Output Location: %s" %path
    d.setdefault('outpath',path)
    l.append(("[%s]" %path,d))
  return l

def getLocation(f, d):
  loc = f[1].split(',')
  longest = 0
  ll = ""
  
  ids_l = ['continent', 'country', 'region', 'district', 'city', 'location', 'scene']
  for item in ids_l:
    d.setdefault(item,"")

  for item in loc:
    if item.startswith('Location.'):
      if len(item) > longest:
          longest = len(item)
          ll = item
  if not ll:
    return d
  
  location_l = ll.split('.')
  
  i = 0
  for item in location_l:
    if item != "Location":
      d[ids_l[i]] = item
      i += 1
  return d

def AllLinks():
  aFile = open(r"%s\AllLinks.txt" %(location), 'a')
  rFile = open(r"%s\link.txt" %(location), 'r')
  l = rFile.readlines()
  rFile.close()
  aFile.write("[%s]\n" %target_folder)
  for file_name in l:
    aFile.write("%s\n" %file_name.strip())
  i = 0
  aFile.close()
  rFile = open(r"%s\AllLinks.txt" %(location), 'r')
  l = rFile.readlines()
  aFile.close()
#  wFile = open("last_location.txt",'w')
#  wFile.write(target_folder)
#  wFile.close()
  return l

def createHardLinks(l):
  i = 0
  for line, d in l:
    if line.startswith("["):
      target_folder = line.lstrip("[").rstrip("]\n")
    i += 1
    
    if not os.path.exists(target_folder):
      os.makedirs(target_folder)
      
    src = d['filefull']
    name = src.split("\\")[-1:][0]
    dst = "%s/%s" %(target_folder, name)
    try:
      HardLink(src,dst)
    except:
      i -= 1
      pass
    try:
      writeMetaDataXMP(d)
      print "Written xmp"
    except:
      print "Failed with XMP writing"
      
  print "Copied %i HardLinks" %i
  
def deal_with_existing_xmp(path, xmp_d):
  tags = ['xmlns:photoshop', 'xmlns:Iptc4xmpCore', 'xmlns:lr', 'xmlns:xap']
  rFile = open(path, 'r')
  f = rFile.readlines()[2:-2]
  rFile.close()
  txt_l = []
  ignore = False
  i = 0
  for line in f:
    line = line
    if not line:
      continue
    i += 1
    
    if ignore and "</rdf:Description>" not in line:
      continue
    elif ignore and "</rdf:Description>":
      ignore = False
      continue
      
    for tag in tags:
      if tag in line:
        ignore = True
        txt_l = txt_l[:-1]
        continue
    if not ignore:
      txt_l.append(line)
  
  wFile = open(path, 'w')

  wFile.write(xmp_d['header'])
  for line in txt_l:
    wFile.write(line)
  wFile.write(xmp_d['add'])
  wFile.write(xmp_d['footer'])

def writeMetaDataXMP(d):
  xmp_header = '''<x:xmpmeta xmlns:x="adobe:ns:meta/" x:xmptk="Adobe XMP Core 4.2-c020 1.124078, Tue Sep 11 2007 23:21:40">
 <rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
 '''

  xmp_add = r'''<rdf:Description rdf:about=""
    xmlns:xap="http://ns.adobe.com/xap/1.0/">
   <xap:ModifyDate>2009-09-11T09:24:37.50+01:00</xap:ModifyDate>
   <xap:CreateDate>2009-09-11T09:24:37.50+01:00</xap:CreateDate>
   <xap:CreatorTool>Ver.1.10</xap:CreatorTool>
   <xap:Rating>%(rating)s</xap:Rating>
   <xap:MetadataDate>2009-10-20T15:47:46.144-01:00</xap:MetadataDate>
   <xap:Label>%(colour)s</xap:Label>
  </rdf:Description>
 
  <rdf:Description rdf:about=""
    xmlns:lr="http://ns.adobe.com/lightroom/1.0/">
   <lr:hierarchicalSubject>
    <rdf:Bag>
     %(categories)s
    </rdf:Bag>
   </lr:hierarchicalSubject>
  </rdf:Description>

  <rdf:Description rdf:about=""
    xmlns:Iptc4xmpCore="http://iptc.org/std/Iptc4xmpCore/1.0/xmlns/">
   <Iptc4xmpCore:Scene>
    <rdf:Bag>
     <rdf:li>%(scene)s</rdf:li>
    </rdf:Bag>
   </Iptc4xmpCore:Scene>
   <Iptc4xmpCore:Location>%(location)s</Iptc4xmpCore:Location>
  </rdf:Description>
  
  <rdf:Description rdf:about=""
    xmlns:photoshop="http://ns.adobe.com/photoshop/1.0/">
   <photoshop:City>%(city)s</photoshop:City>
   <photoshop:State>%(district)s</photoshop:State>
   <photoshop:Country>%(country)s</photoshop:Country>
  </rdf:Description>
  ''' %(d)

  xmp_footer = '''</rdf:RDF>
</x:xmpmeta>
'''

  xmp_d = {'header':xmp_header, 'add':xmp_add, 'footer':xmp_footer}
  
  path = "%(outpath)s/%(filename)s.xmp" %d
  if os.path.exists(path):
    deal_with_existing_xmp(path, xmp_d)
  
  else:
    wFile = open(path, 'w')
    wFile.write(xmp_header)
    wFile.write(xmp_add)
    wFile.write(xmp_footer)
    wFile.close()
  
  

target_folder = "%s\%s" %(location, raw_input("Folder Name: "))
if target_folder.strip('\\') == location:
  rFile = open("last_location.txt",'r')
  target_folder = rFile.read()

if not os.path.exists(target_folder):
  os.makedirs(target_folder)
  
l = MakeLinks(target_folder)
createHardLinks(l)  

print

