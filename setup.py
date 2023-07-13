import sys
from pathlib import Path
from setuptools import setup

from tshistory_xl import __version__


doc = Path(__file__).parent / 'README.md'

deps = [
    'pandas > 1.0.5, < 1.6',
    'colorlover',
    'requests',
    'python-dateutil',
    'isodate',
    'tshistory >= 0.18',
    'tshistory_formula >= 0.14',
    'tshistory_supervision >= 0.11'
]

dev_deps = [
    'rework',
    'pytest',
    'responses',
    'webtest',
    'pytest_sa_pg'
]


if sys.platform in ('darwin', 'win32'):
    deps.append('xlwings == 0.28.5')


setup(name='tshistory_xl',
      version=__version__,
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
      extras_require={
          'dev': dev_deps
      },
      entry_points={
          'tshistory.subcommands': [
              'xl-addin=tshistory_xl.cli:xl_addin',
              'xl=tshistory_xl.cli:xl',
          ],
          'tshclass': [
              'tshclass=tshistory_refinery.tsio:timeseries'
          ],
          'httpclient': [
              'httpclient=tshistory_xl.http:xl_httpclient'
          ]
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
