from setuptools import setup, find_packages
import os

version = open(os.path.join("collective", "plone2docx", "version.txt")).read().strip()

setup(name='collective.plone2docx',
      version=version,
      description="Plone Plugin for rendering docx documents",
      long_description=open(os.path.join("README.md")).read() + "\n" +
                       open(os.path.join("docs", "HISTORY.txt")).read(),
      # Get more strings from
      # http://pypi.python.org/pypi?%3Aaction=list_classifiers
      classifiers=[
        "Framework :: Plone",
        "Programming Language :: Python",
        ],
      keywords='',
      author='Michael Davis',
      author_email='m.r.davis@pretaweb.com',
      url='http://pypi.python.org/pypi/collective.plone2docx',
      license='GPL',
      packages=find_packages(exclude=['ez_setup']),
      namespace_packages=['collective'],
      include_package_data=True,
      zip_safe=False,
      install_requires=[
          'setuptools',
          # -*- Extra requirements: -*-
          'docx',
      ],
      extras_require = {
      'test': [
               'plone.app.testing',
               ]
      },
      entry_points="""
      # -*- Entry points: -*-

      [z3c.autoinclude.plugin]
      target = plone
      """
      )
