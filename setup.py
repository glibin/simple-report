#coding: utf-8
from setuptools import setup, find_packages

requires = []
with open('src/requires.txt', 'r') as f:
    requires.extend(f.readlines())

setup(name='m3-simple-report',
      version='1.3.6.1',
      url='https://bitbucket.org/barsgroup/simple-report',
      license='MIT',
      author='BARS Group',
      author_email='bars@bars-open.ru',
      package_dir={'': 'src'},
      packages=find_packages('src'),
      description=u'Генератор отчетов',
      install_requires=requires,
      include_package_data=True,
      classifiers=[
        'Intended Audience :: Developers',
        'Environment :: Web Environment',
        'Natural Language :: Russian',
        'Natural Language :: English',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'License :: OSI Approved :: MIT License',
        'Development Status :: 5 - Production/Stable',
      ],
      )
