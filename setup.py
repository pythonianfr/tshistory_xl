import os
from setuptools import setup


deps = [
    'pandas',
    'colorlover',
    'requests',
    'python-dateutil',
    'isodate',
]

if os.name == 'nt':
    deps.append('xlwings ~= 0.20')


setup(name='tshistory_xl',
      version='0.1.0',
      author='Pythonian',
      author_email='arnaud.campeas@pythonian.fr, aurelien.campeas@pythonian.fr',
      description='Light client for excel/tshistory',
      packages=['tshistory_xl'],
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
