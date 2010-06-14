import os
import win32file

class CreateHardLinks():
  def __init__(self):
    DriveLetter = os.getcwd()[:1]
    working_directory =  r"%s:\Users\Horst\Pictures\DB\Scripts\HardLink" %DriveLetter
    print working_directory
    rFile = open(r"%s\link.txt" %working_directory , 'r')
    self.drive_letter = rFile.readline().strip()
    self.files = rFile.readlines()
    rFile.close()
    self.stitch_folder = r"%s:\Users\Horst\Pictures\Output\Stitch" %self.drive_letter
    self.location = r"%s:\Users\Horst\Pictures\HardLinks" %self.drive_letter
    self.stitch_folder = "%s\Stitch" %self.location
    self.workpath = "%s:\Users\Horst\Pictures\DB\Scripts" %self.drive_letter
    pass


  def open_log(self):
    self.log = open(r'%s\log.txt' %self.workpath, 'w')
  def close_log(self):
    self.log.close()

  def HardLink(self, src, dst):
    """Hard link function

    Args:
      src: Source filename to link to.
      dst: Destination link name to create.

    Raises:
      OSError: The link could not be created.
    """
    try:
      win32file.CreateHardLink(dst,src)
    except win32file.error as msg:
      # Translate errors into standard OSError
      raise OSError(msg)

  def MakeLinks(self, target_folder):
    self.log.write("Hardlink Export started")
    l = []
    #rFile = open(r"%s\link.txt" %(self.location), 'r')
    #rFile = open(r"link.txt", 'r')
    #files = rFile.readlines()
    #rFile.close()
    for file in self.files:
      d = {}
      f = file.split("\t")

      name = f[0].split("\\")[-1:][0][:-4]

      d = self.get_location(f, d)

      categories = ""
      categories_l = f[1].split(',')
      for category in categories_l:
        cat = category.split(".")
        #cat.reverse()
        if cat[0]:
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
          target_folder = "%s/%s" %(self.location, special)
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

      msg = "Output self.location: %s\n" %path
      print(msg)
      self.log.write(msg)

      d.setdefault('outpath',path)
      l.append(("[%s]" %path,d))
    return l

  def get_location(self, f, d):
    loc = f[1].split(',')
    longest = 0
    ll = ""

    ids_l = ['continent', 'country', 'region', 'district', 'city', 'self.location', 'scene']
    for item in ids_l:
      d.setdefault(item,"")

    for item in loc:
      if item.startswith('self.location.'):
        if len(item) > longest:
            longest = len(item)
            ll = item
    if not ll:
      return d

    self.location_l = ll.split('.')

    i = 0
    for item in self.location_l:
      if item != "self.location":
        d[ids_l[i]] = item
        i += 1
    return d

  def AllLinks(self, ):
    aFile = open(r"%s\AllLinks.txt" %(self.location), 'a')
    rFile = open(r"%s\link.txt" %(self.location), 'r')
    l = rFile.readlines()
    rFile.close()
    aFile.write("[%s]\n" %target_folder)
    for file_name in l:
      aFile.write("%s\n" %file_name.strip())
    i = 0
    aFile.close()
    rFile = open(r"%s\AllLinks.txt" %(self.location), 'r')
    l = rFile.readlines()
    aFile.close()
  #  wFile = open("last_self.location.txt",'w')
  #  wFile.write(target_folder)
  #  wFile.close()
    return l

  def createHardLinks(self, l):
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
        self.HardLink(src,dst)
      except:
        i -= 1
        pass
      try:
        self.writeMetaDataXMP(d)
        msg ="Written xmp\n"
        print(msg)
        self.log.write(msg)
      except:
        msg = "Failed with XMP writing\n"
        print(msg)
        self.log.write(msg)

    msg = "Copied %i HardLinks\n" %i
    print(msg)
    self.log.write(msg)
    self.log.close()


  def deal_with_existing_xmp(self, path, xmp_d):
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

  def writeMetaDataXMP(self, d):
    xmp_header = '''<x:xmpmeta xmlns:x="adobe:ns:meta/" x:xmptk="Adobe XMP Core 4.2-c020 1.124078, Tue Sep 11 2007 23:21:40">
   <rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
   '''

    xmp_add = r'''<rdf:Description rdf:about=""
      xmlns:xap="http://ns.adobe.com/xap/1.0/">
     <xap:CreatorTool>Ver.1.10</xap:CreatorTool>
     <xap:Rating>%(rating)s</xap:Rating>
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
     <Iptc4xmpCore:self.location>%(self.location)s</Iptc4xmpCore:self.location>
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


  def run(self):
    self.open_log()
    target_folder = "%s\%s" %(self.location, raw_input("Folder Name: "))
    if target_folder.strip('\\') == self.location:
      rFile = open("%s\last_location.txt" %self.workpath,'r')
      target_folder = rFile.read()
    else:
      wFile = open("%s\last_location.txt"  %self.workpath,'w')
      wFile.write(target_folder)
      wFile.close()
    if not os.path.exists(target_folder):
      os.makedirs(target_folder)
    l = self.MakeLinks(target_folder)
    self.createHardLinks(l)
    self.close_log()

if __name__ == "__main__":
  a = CreateHardLinks()
  a.run()
