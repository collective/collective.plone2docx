from DateTime import DateTime
import imghdr
from lxml import etree
import os
import shutil
import zipfile

from PIL import Image

import pynliner

from zope.component import getAdapters
from zope.interface import implementer
from zope.publisher.interfaces import IPublishTraverse

from plone.subrequest import subrequest
from plone.transformchain.interfaces import ITransform

from Products.CMFCore.utils import getToolByName
from Products.Five import BrowserView

import docx
from docx import makeelement

BANNED_IDS = ['edit-bar', 'contentActionMenus']

def sort_key(a, b):
    return cmp(a.order, b.order)

def get_attrs(element, nested_tags=[]):
    yield element
    # If an element has any children (nested elements) loop through them
    # unless they are specified to not loop
    if len(element) and element.tag.replace('{http://www.w3.org/1999/xhtml}', '') not in nested_tags:
        if element.attrib.has_key('id') and element.attrib['id'] not in BANNED_IDS:
            for node in element:
                # Recursively call this function, yielding each result:
                for attribute in get_attrs(node):
                    yield attribute

def add_header_and_footer(relationships, body):
    '''Add some content as a footer'''
    template_dir = os.path.dirname(docx.__file__)
    namespacemap = {}
    namespacemap['w'] = docx.nsprefixes['w']
    namespacemap['r'] = docx.nsprefixes['r']
    namespace = '{'+docx.nsprefixes['w']+'}'
    attributenamespace = '{'+docx.nsprefixes['r']+'}'
    footer_name = 'footer.xml'
    footer_rid = 'rId'+str(len(relationships)+1)
    relationships.append(['http://schemas.openxmlformats.org/officeDocument/2006/relationships/', footer_name])
    header_name = 'header.xml'
    header_rid = 'rId'+str(len(relationships)+1)
    relationships.append(['http://schemas.openxmlformats.org/officeDocument/2006/relationships/', header_name])
    p = docx.makeelement('p')
    pPr = docx.makeelement('pPr')
    sectPr = docx.makeelement('sectPr')
    # TODO can't use makeelement here as it only handles single namespace for attribs
    footerReference = etree.Element(namespace + 'footerReference', nsmap=namespacemap)
    footerReference.set(attributenamespace + 'id', footer_rid)
    footerReference.set(namespace + 'type', 'default')
    headerReference = etree.Element(namespace + 'headerReference', nsmap=namespacemap)
    headerReference.set(attributenamespace + 'id', header_rid)
    headerReference.set(namespace + 'type', 'default')
    sectPr.append(footerReference)
    sectPr.append(headerReference)
    # TODO page size should be customisable, hard code to A4 for now
    sectPr.append(makeelement('pgSz', nsprefix='w', attributes={'h':'16838', 'w':'11906'}))
    # TODO hard code margins
    sectPr.append(makeelement('pgMar', nsprefix='w', attributes={'bottom':'+1440',
                                                                 'footer':'720',
                                                                 'gutter':'0',
                                                                 'header':'720',
                                                                 'left':'1800',
                                                                 'right':'1800',
                                                                 'top':'+1800'}))
    pPr.append(sectPr)
    p.append(pPr)
    body.append(p)

def newdocument():
    document = docx.makeelement('document', nsprefix=['w', 'r', 'wp', 'a', 'pic'])
    document.append(docx.makeelement('body'))
    return document

def new_footer():
    footer = docx.makeelement('document', nsprefix=['w', 'r'])
    footer.append(docx.makeelement('ftr'))
    return footer

def new_header():
    header = docx.makeelement('document', nsprefix=['w', 'r', 'wp', 'a'])
    header.append(docx.makeelement('hdr'))
    return header

def fix_entities(html_string):
    """Entities are not defined in the html dtd"""
    html_string = html_string.replace('&nbsp;', u'\u00a0')
    return html_string

@implementer(IPublishTraverse)
class DocxView(BrowserView):
    """View a plone object in docx format"""

    def __init__(self, context, request):
        # don't put anything that needs plone api in here, as we may not have a context yet
        super(DocxView, self).__init__(context, request)
        self.image_count = 0

    def __call__(self):
        word_template_path = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir, os.pardir, 'docx-template'))
        # TODO get the var folder, to make sure user has write permissions and use a random for folder id
        destination_path = os.path.join(os.getcwd(), 'docx_temp')
        if os.path.exists(destination_path):
            shutil.rmtree(destination_path)
        shutil.copytree(word_template_path, destination_path)
        self.working_folder = destination_path
        self.create_the_docx()
        return self.set_the_response()

    def create_the_docx(self):
        self.relationships = docx.relationshiplist()
        self.content_types_list = {}
        document = newdocument()
        page = self.get_the_page()
        page = pynliner.fromString(page)
        page = fix_entities(page)
        tree = etree.fromstring(page)
        body = document.xpath('/w:document/w:body', namespaces=docx.nsprefixes)[0]
        self.write_the_docx(body, tree)
        self.write_the_header(tree)
        self.write_the_footer(tree)
        add_header_and_footer(self.relationships, body)
        self.zip_the_docx(document)
        return

    def write_the_header(self, tree):
        # TODO keep as a tree, rather than writing to the filesystem
        header_doc = new_header()
        header = header_doc.xpath('/w:document/w:hdr', namespaces=docx.nsprefixes)[0]
        header_content = self.get_header_content(tree)
        self.add_header_image(header_content, header)
        file_object = open(self.working_folder + '/word/header.xml', 'w')
        file_object.write(etree.tostring(header))
        file_object.close()

    def add_header_image(self, element, body):
        """Adding an image in the header is different to the body"""
        # TODO defensive coding
        src_url = element.attrib['src']
        urltool = getToolByName(self.context, "portal_url")
        portal = urltool.getPortalObject()
        base_url = portal.absolute_url()
        url = base_url + '/' + src_url
        media_path = self.working_folder + '/word/media'
        if not os.path.exists(media_path):
            os.makedirs(media_path)
        self.image_count += 1
        picid = str(self.image_count)
        url_parts = url.split('/')
        picname = url_parts[-1]
        picdescription = 'The header image'
        file_object = open(media_path + '/' + picname, 'w')
        image_string = subrequest(url).getBody()
        file_object.write(image_string)
        file_object.close()
        picrelid = 'rId'+str(len(self.relationships)+1)
        # TODO this should be moved to a separate method
        rels_content = types = etree.fromstring('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
        rels_content.append(makeelement('Relationship', nsprefix=None,
                                        attributes={'Id':picrelid,
                                                    'Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                                                    'Target':'media/%s' % picname}))
        rels_path = self.working_folder + '/word/_rels'
        if not os.path.exists(rels_path):
            os.makedirs(rels_path)
        # TODO hard code header name for now
        file_object = open(rels_path + '/header.xml.rels', 'w')
        file_object.write(etree.tostring(rels_content))
        file_object.close()
        self.relationships.append(['http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                         'media/'+picname])
        # TODO hard code the content_tpe entry for now
        self.content_types_list['/word/media/%s' % picname] = 'image/jpeg'
        # TODO hard code dimensions for now
        width = 7560310
        height = 1378585
        graphic = self.create_graphic_tag(width, height, picrelid, picid, picname, picdescription)
        # This needs to be in an anchor rather than a framelocks
        # TODO atrbibs shouldn't have a namespace
        anchor = docx.makeelement('anchor', nsprefix='wp',
                                  attributes={'allowOverlap':'1',
                                              'behindDoc':'1',
                                              'distB':'0',
                                              'distL':'0',
                                              'distR':'0',
                                              'distT':'0',
                                              'layoutInCell':'1',
                                              'locked':'0',
                                              'relativeHeight':'12',
                                              'simplePos':'0'})
        anchor.append(docx.makeelement('simplePos', nsprefix='wp', attributes={'x':'0', 'y':'0'}))
        positionH = docx.makeelement('positionH', nsprefix='wp', attributes={'relativeFrom':'character',})
        positionH.append(docx.makeelement('posOffset', tagtext='-1143000', nsprefix='wp'))
        anchor.append(positionH)
        positionV = docx.makeelement('positionV', nsprefix='wp', attributes={'relativeFrom':'line',})
        positionV.append(docx.makeelement('posOffset', tagtext='+457200', nsprefix='wp'))
        anchor.append(positionV)
        anchor.append(docx.makeelement('extent', nsprefix='wp', attributes={'cx':str(width), 'cy':str(height)}))
        anchor.append(docx.makeelement('effectExtent', nsprefix='wp', attributes={'b':'0', 'l':'0', 'r':'0', 't':'0'}))
        anchor.append(docx.makeelement('wrapNone', nsprefix='wp'))
        anchor.append(docx.makeelement('docPr', nsprefix='wp', attributes={'id': picid, 'name': 'Picture 1', 'descr': picdescription}))
        cNvGraphicFramePr = docx.makeelement('cNvGraphicFramePr', nsprefix='wp')
        cNvGraphicFramePr.append(docx.makeelement('graphicFrameLocks', nsprefix='a', attributes={'noChangeAspect':'1',}))
        anchor.append(cNvGraphicFramePr)
        # now we can append the actual graphic
        anchor.append(graphic)
        drawing = docx.makeelement('drawing', nsprefix='w')
        drawing.append(anchor)
        r = docx.makeelement('r', nsprefix='w')
        r.append(drawing)
        p = docx.makeelement('p', nsprefix='w')
        p.append(r)
        body.append(p)

    def get_header_content(self, tree):
        # TODO for now assume the header contains a single tag which is an image
        html_body = tree[1]
        content = html_body.xpath("//*[@id='docx_header']")
        if len(content) == 1:
            content = content[0]
        else:
            # TODO do something sensible
            return ''
        # TODO for now assume only a single image
        image_tag = content[0]
        return image_tag

    def write_the_footer(self, tree):
        # TODO keep as a tree, rather than writing to the filesystem
        namespacemap = {}
        namespacemap['w'] = docx.nsprefixes['w']
        namespacemap['r'] = docx.nsprefixes['r']
        footer_doc = new_footer()
        footer = footer_doc.xpath('/w:document/w:ftr', namespaces=docx.nsprefixes)[0]
        footer_text = self.get_footer_content(tree)
        self.add_page_number(footer, footer_text)
        file_object = open(self.working_folder + '/word/footer.xml', 'w')
        file_object.write(etree.tostring(footer))
        file_object.close()

    def add_page_number(self, footer, text):
        # TODO this needs tidying
        p = docx.makeelement('p')
        r1 = docx.makeelement('r')
        text = docx.makeelement('t', text)
        tab = docx.makeelement('tab')
        r2 = docx.makeelement('r')
        fldChar1 = docx.makeelement('fldChar', attributes={'fldCharType':'begin'})
        r3 = docx.makeelement('r')
        instrText = docx.makeelement('instrText', 'PAGE')
        r4 = docx.makeelement('r')
        fldChar2 = docx.makeelement('fldChar', attributes={'fldCharType':'separate'})
        r5 = docx.makeelement('r')
        text2 = docx.makeelement('t', '11')
        r6 = docx.makeelement('r')
        fldChar3 = docx.makeelement('fldChar', attributes={'fldCharType':'end'})
        r6.append(fldChar3)
        r5.append(text2)
        r4.append(fldChar2)
        r3.append(instrText)
        r2.append(fldChar1)
        r1.append(text)
        r1.append(tab)
        p.append(r1)
        p.append(r2)
        p.append(r3)
        p.append(r4)
        p.append(r5)
        p.append(r6)
        footer.append(p)

    def get_footer_content(self, tree):
        html_body = tree[1]
        content = html_body.xpath("//*[@id='docx_footer']")
        if len(content) == 1:
            content = content[0]
        else:
            # either no footer id or multiple, so return empty string
            return ''
        # TODO for now assume no nested tags
        return content.text.strip()

    def write_the_docx(self, body, tree):
        # TODO deal with docs that have no head element
        html_head = tree[0]
        html_body = tree[1]
        content = html_body.xpath("//*[@id='content']")
        if len(content) == 1:
            content = content[0]
        else:
            # either no content id or multiple, so just use the body
            content = html_body
        for item in get_attrs(content, nested_tags=['table', 'ul']):
            # get rid of the namespace
            try:
                tag = item.tag.replace('{http://www.w3.org/1999/xhtml}', '')
            except AttributeError:
                # if tag is callable, then it's probably a comment
                continue
            self.add_element(body, item, tag)

    def add_element(self, body, element, tag):
        """Add the element to the document"""
        if element.attrib.has_key('style'):
            if 'page-break-before' in element.attrib['style']:
                body.append(docx.pagebreak())
        if tag == 'h1':
            if element.text:
                body.append(docx.heading(element.text.strip(), 1))
        elif tag == 'h2':
            if element.text:
                body.append(docx.heading(element.text.strip(), 2))
        elif tag == 'h3':
            if element.text:
                body.append(docx.heading(element.text.strip(), 3))
        elif tag == 'p':
            if element.text:
                # TODO whitespace counts as text
                body.append(docx.paragraph(element.text.strip()))
            elif len(element) == 1:
                self.add_inline_image(element[0], body)
        elif tag == 'ul':
            self.add_a_list(element, body)
        elif tag == 'table':
            self.add_a_table(element, body)
        elif tag == 'img':
            self.add_anchor_image(element, body)

    def add_a_list(self, element, body):
        # TODO doesn't do nested lists
        for item in element:
            tag = item.tag.replace('{http://www.w3.org/1999/xhtml}', '')
            if tag == 'ul':
                continue
            if item.text:
                body.append(docx.paragraph(item.text.strip(), style='ListBullet'))

    def add_a_table(self, element, body):
        if len(element) == 0:
            # empty table so do nothing
            return
        table_content = []
        borders = self.set_table_borders(element)
        # TODO handle tables with a thead
        if element[0].tag.replace('{http://www.w3.org/1999/xhtml}', '') == 'tbody':
            table_rows = element[0]
        else:
            table_rows = element
        no_columns = len(table_rows[0])
        for table_row in table_rows:
            row_content = []
            for cell in table_row:
                if cell.text:
                    row_content.append(cell.text.strip())
            table_content.append(row_content)
        column_widths = []
        # TODO only twips are supported
        # 2390 is just over 4.22cm
        # page width is 14.6cm plus margins
        # so content width is just over 8260 twips
        cell_width = 8260/no_columns
        for i in range(no_columns):
            column_widths.append(cell_width)
        body.append(docx.table(table_content, heading=False, colw=column_widths, borders=borders))
        body.append(docx.paragraph(''))

    def set_table_borders(self, element):
        """Setup the table borders"""
        # TODO actually obey the css style
        borders = {'all': {}}
        if element.attrib.has_key('style'):
            style = element.attrib['style']
        else:
            return borders
        borders['all']['color'] = 'black'
        borders['all']['space'] = '8'
        borders['all']['sz'] = '8'
        borders['all']['val'] = 'single'
        return borders

    def download_image(self, src_url):
        """Download an image"""
        # TODO assume a relative link
        urltool = getToolByName(self.context, "portal_url")
        portal = urltool.getPortalObject()
        base_url = portal.absolute_url()
        url = base_url + '/' + src_url
        self.image_count += 1
        picid = str(self.image_count)
        # figure out what kind of image it is
        # TODO should check an image is actually returned
        image_response = subrequest(url)
        image_string = image_response.getBody()
        image_type = imghdr.what('ignore_this', h=image_string)
        picname = picid + '.' + image_type
        # TODO should check for an alt tag for the description
        picdescription = ''
        media_path = os.path.join(self.working_folder, 'word', 'media')
        if not os.path.exists(media_path):
            os.makedirs(media_path)
        image_path = os.path.join(media_path, picname)
        file_object = open(image_path, 'w')
        file_object.write(image_string)
        file_object.close()
        pil_image = Image.open(image_path)
        width, height = self.calculate_image_sizes(pil_image)
        # sizes should be in twips, and it's around 118dpi
        height = height*914400/118
        width = width*914400/118
        return picid, picname, picdescription, width, height

    def calculate_image_sizes(self, pil_image):
        """Get the image sizes"""
        # TODO this should look at the style on the img tag
        # resize the image so it fits on portrait page
        # A4 is just shy of 11906x16838
        # Taking out the margins
        # TODO since we are in portrait, just care about the width
        # which is 8306 twips, or 151,900,130,000 emus
        # which should be a little over 1400 pixels, but isn't
        width, height = pil_image.size
        if width > 800:
            ratio = float(800)/width
            width = int(width * ratio)
            height = int(height * ratio)
        return width, height

    def add_anchor_image(self, element, body):
        """Put an anchored image into the page"""
        # TODO defensive coding
        src_url = element.attrib['src']
        picid, picname, picdescription, width, height = self.download_image(src_url)
        picrelid = 'rId'+str(len(self.relationships)+1)
        self.relationships.append(['http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'media/'+picname])
        graphic = self.create_graphic_tag(width, height, picrelid, picid, picname, picdescription)
        anchor = docx.makeelement('anchor', nsprefix='wp',
                                  attributes={'allowOverlap':'1',
                                              'behindDoc':'0',
                                              'distB':'0',
                                              'distL':'0',
                                              'distR':'0',
                                              'distT':'0',
                                              'layoutInCell':'1',
                                              'locked':'0',
                                              'relativeHeight':'3',
                                              'simplePos':'0'})
        anchor.append(docx.makeelement('simplePos', nsprefix='wp', attributes={'x':'0', 'y':'0'}))
        positionH = docx.makeelement('positionH', nsprefix='wp', attributes={'relativeFrom':'column',})
        positionH.append(docx.makeelement('posOffset', tagtext='1506220', nsprefix='wp'))
        anchor.append(positionH)
        positionV = docx.makeelement('positionV', nsprefix='wp', attributes={'relativeFrom':'paragraph',})
        positionV.append(docx.makeelement('posOffset', tagtext='0', nsprefix='wp'))
        anchor.append(positionV)
        anchor.append(docx.makeelement('extent', nsprefix='wp', attributes={'cx':str(width), 'cy':str(height)}))
        anchor.append(docx.makeelement('effectExtent', nsprefix='wp', attributes={'b':'0', 'l':'0', 'r':'0', 't':'0'}))
        anchor.append(docx.makeelement('wrapNone', nsprefix='wp'))
        anchor.append(docx.makeelement('docPr', nsprefix='wp', attributes={'id': picid, 'name': 'Picture 1', 'descr': picdescription}))
        cNvGraphicFramePr = docx.makeelement('cNvGraphicFramePr', nsprefix='wp')
        cNvGraphicFramePr.append(docx.makeelement('graphicFrameLocks', nsprefix='a', attributes={'noChangeAspect':'1',}))
        anchor.append(cNvGraphicFramePr)
        # now we can append the actual graphic
        anchor.append(graphic)
        drawing = docx.makeelement('drawing', nsprefix='w')
        drawing.append(anchor)
        r = docx.makeelement('r', nsprefix='w')
        r.append(docx.makeelement('rPr', nsprefix='w'))
        r.append(drawing)
        p = docx.makeelement('p', nsprefix='w')
        pPr = docx.makeelement('pPr', nsprefix='w')
        pPr.append(docx.makeelement('pStyle', nsprefix='w', attributes={'val':'style0',}))
        p.append(pPr)
        p.append(r)
        body.append(p)

    def add_inline_image(self, element, body):
        """Put an inline image into the document"""
        # TODO defensive coding
        src_url = element.attrib['src']
        picid, picname, picdescription, width, height = self.download_image(src_url)
        picrelid = 'rId'+str(len(self.relationships)+1)
        self.relationships.append(['http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'media/'+picname])
        graphic = self.create_graphic_tag(width, height, picrelid, picid, picname, picdescription)
        inline = docx.makeelement('inline', nsprefix='wp', attributes={'distT':'0', 'distR':'0', 'distL':'0', 'distB':'0'})
        inline.append(docx.makeelement('extent', nsprefix='wp', attributes={'cy':str(height), 'cx':str(width)}))
        inline.append(docx.makeelement('effectExtent', nsprefix='wp', attributes={'r':'0', 'b':'0', 'l':'25400', 't':'0'}))
        inline.append(docx.makeelement('docPr', nsprefix='wp', attributes={'id':picid, 'descr':picdescription, 'name':picname}))
        cNvGraphicFramePr = docx.makeelement('cNvGraphicFramePr', nsprefix='wp')
        cNvGraphicFramePr.append(docx.makeelement('graphicFrameLocks', nsprefix='a', attributes={'noChangeAspect':'1',}))
        inline.append(cNvGraphicFramePr)
        inline.append(graphic)
        drawing = docx.makeelement('drawing', nsprefix='w')
        drawing.append(inline)
        r = docx.makeelement('r', nsprefix='w')
        r.append(drawing)
        p = docx.makeelement('p', nsprefix='w')
        p.append(r)
        body.append(p)

    def create_graphic_tag(self, width, height, picrelid, picid, picname, picdescription):
        """Create a graphic tag for an image"""
        # There are 3 main elements inside a picture
        nochangeaspect=True
        nochangearrowheads=True
        # 1. The Blipfill - specifies how the image fills the picture area (stretch, tile, etc.)
        blipfill = makeelement('blipFill', nsprefix='pic')
        blipfill.append(makeelement('blip', nsprefix='a', attrnsprefix='r',
                                    attributes={'embed': picrelid}))
        stretch = makeelement('stretch', nsprefix='a')
        stretch.append(makeelement('fillRect', nsprefix='a'))
        blipfill.append(makeelement('srcRect', nsprefix='a'))
        blipfill.append(stretch)
        # 2. The non visual picture properties
        nvpicpr = makeelement('nvPicPr', nsprefix='pic')
        cnvpr = makeelement('cNvPr', nsprefix='pic',
                            attributes={'id': picid, 'name': picname, 'descr': picdescription})
        nvpicpr.append(cnvpr)
        cnvpicpr = makeelement('cNvPicPr', nsprefix='pic')
        cnvpicpr.append(makeelement('picLocks', nsprefix='a',
                                    attributes={'noChangeAspect': str(int(nochangeaspect)),
                                    'noChangeArrowheads': str(int(nochangearrowheads))}))
        nvpicpr.append(cnvpicpr)
        # 3. The Shape properties
        sppr = makeelement('spPr', nsprefix='pic', attributes={'bwMode': 'auto'})
        xfrm = makeelement('xfrm', nsprefix='a')
        xfrm.append(makeelement('off', nsprefix='a', attributes={'x': '0', 'y': '0'}))
        xfrm.append(makeelement('ext', nsprefix='a', attributes={'cx': str(width), 'cy': str(height)}))
        prstgeom = makeelement('prstGeom', nsprefix='a', attributes={'prst': 'rect'})
        prstgeom.append(makeelement('avLst', nsprefix='a'))
        sppr.append(xfrm)
        sppr.append(prstgeom)
        ln = makeelement('ln', nsprefix='a', attributes={'w': '9525'})
        ln.append(makeelement('noFill', nsprefix='a'))
        ln.append(makeelement('miter', nsprefix='a', attributes={'lim': '800000'}))
        ln.append(makeelement('headEnd', nsprefix='a'))
        ln.append(makeelement('tailEnd', nsprefix='a'))
        sppr.append(ln)
        # Add our 3 parts to the picture element
        pic = makeelement('pic', nsprefix='pic')
        pic.append(nvpicpr)
        pic.append(blipfill)
        pic.append(sppr)
        # Now make the supporting elements
        # The following sequence is just: make element, then add its children
        graphicdata = makeelement('graphicData', nsprefix='a',
                                  attributes={'uri': 'http://schemas.openxmlforma'
                                  'ts.org/drawingml/2006/picture'})
        graphicdata.append(pic)
        graphic = makeelement('graphic', nsprefix='a')
        graphic.append(graphicdata)
        return graphic

    def fix_the_content_types(self):
        contenttypes = docx.contenttypes()
        for content_type in self.content_types_list:
            contenttypes.append(docx.makeelement('Override', nsprefix=None,
                                                 attributes={'PartName': content_type,
                                                             'ContentType': self.content_types_list[content_type]}))
        return contenttypes

    def zip_the_docx(self, document):
        title = 'foo'
        subject = 'foo'
        creator = 'foo'
        keywords = 'foo'
        coreprops = docx.coreproperties(title=title, subject=subject, creator=creator, keywords=keywords)
        appprops = docx.appproperties()
        contenttypes = self.fix_the_content_types()
        websettings = docx.websettings()
        wordrelationships = docx.wordrelationships(self.relationships)
        file_name = 'filename.docx'
        # Save our document
        # TODO all assets need to be in template_dir to work with this method
        # template_dir should be copied somewhere else and this method rewritten
        # as our use case is to create the same docx template this currently meets our needs
        self.savedocx(document, coreprops, appprops, contenttypes, websettings,wordrelationships, file_name)
        shutil.rmtree(self.working_folder)
        return

    def get_the_page(self):
        """Get the raw html page"""
        page = self.context()
        # check if diazo is enabled
        # TODO if diazo not enabled you'll get an mdash entity not defined
        if self.request.get('HTTP_X_THEME_ENABLED', None):
            page = self.transform_with_diazo(page)
        return page

    def transform_with_diazo(self, raw_html):
        published = self.request.get('PUBLISHED', None)
        handlers = [v[1] for v in getAdapters((published, self.request,), ITransform)]
        handlers.sort(sort_key)
        # The first handler is the diazo transform
        theme_handler = handlers[0]
        charset = self.context.portal_properties.site_properties.default_charset
        new_html = theme_handler.transformIterable([raw_html], charset)
        # If the theme is not enabled, transform returns None
        if new_html is not None:
            new_html = etree.tostring(new_html.tree)
        else:
            new_html = raw_html
        return new_html

    def set_the_response(self):
        nice_filename = 'filename.docx'
        file_object = open(nice_filename)
        stream = file_object.read()
        file_object.close()
        os.remove(nice_filename)

        self.request.response.setHeader("Content-Disposition",
                                        "attachment; filename=%s" %
                                        nice_filename)
        self.request.response.setHeader("Content-Type", "application/msword")
        self.request.response.setHeader("Content-Length", len(stream))
        self.request.response.setHeader('Last-Modified', DateTime.rfc822(DateTime()))
        self.request.response.setHeader("Cache-Control", "no-store")
        self.request.response.setHeader("Pragma", "no-cache")
        self.request.response.write(stream)

    def savedocx(self, document, coreprops, appprops, contenttypes, websettings, wordrelationships, output):
        '''Save a modified document'''
        # copied from docx so we can change the template_dir
        template_dir = self.working_folder
        assert os.path.isdir(template_dir)
        docxfile = zipfile.ZipFile(output, mode='w', compression=zipfile.ZIP_DEFLATED)

        # Move to the template data path
        prev_dir = os.path.abspath('.')  # save previous working dir
        os.chdir(template_dir)

        # Serialize our trees into out zip file
        treesandfiles = {document:     'word/document.xml',
            coreprops:    'docProps/core.xml',
            appprops:     'docProps/app.xml',
            contenttypes: '[Content_Types].xml',
            websettings:  'word/webSettings.xml',
            wordrelationships: 'word/_rels/document.xml.rels'}
        for tree in treesandfiles:
            docx.log.info('Saving: %s' % treesandfiles[tree])
            treestring = etree.tostring(tree, pretty_print=True)
            docxfile.writestr(treesandfiles[tree], treestring)

        # Add & compress support files
        files_to_ignore = ['.DS_Store']  # nuisance from some os's
        for dirpath, dirnames, filenames in os.walk('.'):
            for filename in filenames:
                if filename in files_to_ignore:
                    continue
                templatefile = os.path.join(dirpath, filename)
                archivename = templatefile[2:]
                docx.log.info('Saving: %s', archivename)
                docxfile.write(templatefile, archivename)
        docx.log.info('Saved new file to: %r', output)
        docxfile.close()
        os.chdir(prev_dir)  # restore previous working dir
        return
