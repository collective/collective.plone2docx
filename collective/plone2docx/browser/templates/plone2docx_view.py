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

def sort_key(a, b):
    return cmp(a.order, b.order)

@implementer(IPublishTraverse)
class DocxView(BrowserView):
    """View a plone object in docx format"""

    def __call__(self):
        page = self.get_the_page()
        self.create_the_docx()
        return self.set_the_response()

    def create_the_docx(self):
        relationships = docx.relationshiplist()
        document = docx.newdocument()
        page = self.get_the_page()
        body = document.xpath('/w:document/w:body', namespaces=docx.nsprefixes)[0]
        self.dummy_content(body)
        self.zip_the_docx(relationships, document)
        return

    def dummy_content(self, body):
        body.append(docx.heading('Editing documents', 2))
        body.append(docx.paragraph('Thanks to the awesomeness of the lxml module, '
                              'we can:'))

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
