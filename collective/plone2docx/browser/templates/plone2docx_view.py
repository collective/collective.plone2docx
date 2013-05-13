from DateTime import DateTime
from lxml import etree
import os
import zipfile

from zope.component import getAdapters
from zope.interface import implementer
from zope.publisher.interfaces import IPublishTraverse

from plone.transformchain.interfaces import ITransform

from Products.Five import BrowserView

import docx

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

@implementer(IPublishTraverse)
class DocxView(BrowserView):
    """View a plone object in docx format"""

    def __call__(self):
        self.create_the_docx()
        return self.set_the_response()

    def create_the_docx(self):
        relationships = docx.relationshiplist()
        document = docx.newdocument()
        page = self.get_the_page()
        tree = etree.fromstring(page)
        body = document.xpath('/w:document/w:body', namespaces=docx.nsprefixes)[0]
        self.write_the_docs(body, tree)
        self.zip_the_docx(relationships, document)
        return

    def write_the_docs(self, body, tree):
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
        if tag == 'h1':
            body.append(docx.heading(element.text.strip(), 1))
        elif tag == 'h2':
            body.append(docx.heading(element.text.strip(), 2))
        elif tag == 'h3':
            body.append(docx.heading(element.text.strip(), 3))
        elif tag == 'p':
            body.append(docx.paragraph(element.text.strip()))
        elif tag == 'ul':
            self.add_a_list(element, body)
        elif tag == 'table':
            self.add_a_table(element, body)

    def add_a_list(self, element, body):
        items = get_attrs(element)
        # TODO doesn't do nested lists
        for item in items:
            tag = item.tag.replace('{http://www.w3.org/1999/xhtml}', '')
            if tag == 'ul':
                continue
            if item.text:
                body.append(docx.paragraph(item.text.strip(), style='ListBullet'))

    def add_a_table(self, element, body):
        table_content = []
        table_rows = element[0]
        for table_row in table_rows:
            row_content = []
            for cell in table_row:
                if cell.text:
                    row_content.append(cell.text.strip())
            table_content.append(row_content)
        body.append(docx.table(table_content))

    def zip_the_docx(self, relationships, document):
        title = 'foo'
        subject = 'foo'
        creator = 'foo'
        keywords = 'foo'
        coreprops = docx.coreproperties(title=title, subject=subject, creator=creator, keywords=keywords)
        appprops = docx.appproperties()
        contenttypes = docx.contenttypes()
        websettings = docx.websettings()
        wordrelationships = docx.wordrelationships(relationships)
        file_name = 'filename.docx'
        # Save our document
        docx.savedocx(document, coreprops, appprops, contenttypes, websettings,wordrelationships, file_name)
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
        file = open(nice_filename)
        stream = file.read()
        file.close()
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
