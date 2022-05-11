import sys
from pathlib import Path
from setuptools import setup


doc = Path(__file__).parent / 'README.md'

deps = [
    'pandas',
    'colorlover',
    'requests',
    'python-dateutil',
    'isodate',
    'tshistory',
    'tshistory_formula',
    'tshistory_supervision'
]

if sys.platform in ('darwin', 'win32'):
    deps.append('xlwings ~= 0.20')


setup(name='tshistory_xl',
      version='0.3.0',
      author='Pythonian',
      author_email='arnaud.campeas@pythonian.fr, aurelien.campeas@pythonian.fr',
      description='Light client for excel/tshistory',
      long_description=doc.read_text(),
      long_description_content_type='text/markdown',
      packages=['tshistory_xl'],
      zip_safe=False,
      package_data={'tshistory_xl': [
          'ZTSHISTORY.xlam',
      ]},
      install_requires=deps,
      entry_points={
          'tshistory.subcommands': [
              'xl-addin=tshistory_xl.cli:xl_addin',
              'xl=tshistory_xl.cli:xl',
          ],
      },
      classifiers=[
          'Development Status :: 4 - Beta',
          'Intended Audience :: Developers',
          'License :: OSI Approved :: GNU Lesser General Public License v3 (LGPLv3)',
          'Operating System :: OS Independent',
          'Programming Language :: Python :: 3',
          'Topic :: Database',
          'Topic :: Scientific/Engineering',
          'Topic :: Software Development :: Version Control'
      ]
)
