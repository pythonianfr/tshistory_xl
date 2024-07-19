import os
import sys
import shutil
from pathlib import Path
import time

import click

from tshistory_xl.excel_connect import (
    macro_pull_all,
    macro_pull_tab,
    macro_push_all,
    macro_push_tab,
)

# Directories/paths
XLAM = Path(__file__).parent / 'ZTSHISTORY.xlam'
PYXLL_XLAM = Path(__file__).parent / 'zts-refinery.xlam'

if sys.platform.startswith('win'):
    ADDINPATH = Path(os.getenv('APPDATA')) / 'Microsoft'/ 'Excel'/ 'XLSTART'


def addin_install(pyxll):
    if not pyxll:
        current_path = XLAM
        target_path = ADDINPATH / 'ZTSHISTORY.xlam'
    else:
        current_path = PYXLL_XLAM
        target_path = ADDINPATH / 'zts-refinery.xlam'
    if not sys.platform.startswith('win'):
        print('Cannot install the addin automatically on Mac. ')
        print('Install it via Tools > Excel Add-ins...')
        print(f'You find the addin here: {XLAM}')
    else:
        try:
            shutil.copyfile(current_path, target_path)
            print(
                'Successfully installed the tshistory_xl add-in! '
                'Please restart Excel.'
            )
        except IOError as e:
            if e.args[0] == 13:
                print('Error: Failed to install the add-in: '
                      'If Excel is running, quit Excel and try again.')
            else:
                print(str(e))
        except Exception as e:
            print(str(e))


def addin_remove(name, pyxll):
    if not pyxll:
        target_path = ADDINPATH / 'ZTSHISTORY.xlam'
    else:
        target_path = ADDINPATH / 'zts-refinery.xlam'
    if name is not None:
        target_path = ADDINPATH / name
    if not sys.platform.startswith('win'):
        print('Error: This command is not available on Mac. '
              'Please remove the addin manually.')
    else:
        try:
            os.remove(target_path)
            print('Successfully removed the tshistory_xl add-in!')
        except WindowsError as e:
            if e.args[0] == 32:
                print('Error: Failed to remove the add-in: '
                      'If Excel is running, quit Excel and try again.')
            elif e.args[0] == 2:
                print("Error: Could not remove the add-in. "
                      "The add-in doesn't seem to be installed.")
            else:
                print(str(e))
        except Exception as e:
            print(str(e))


@click.command('xl-addin')
@click.argument('action',
                type=click.Choice(['install', 'uninstall', "uninstall-any"]))
@click.option('--name')
@click.option('--pyxll', is_flag=True, default=False)
def xl_addin(action, name=None, pyxll=False):
    if action == 'install':
        addin_install(pyxll)
    elif action == 'uninstall':
        addin_remove(name=None, pyxll=pyxll)
    elif action == 'uninstall-any':
        if name is None:
            raise Exception('An excel addin name must be given')
        addin_remove(name)


@click.command('xl')
@click.argument('action',
                type=click.Choice(['push', 'pull']))
@click.argument('xl-path', type=click.Path())
@click.option('--tab')
def xl(action, xl_path, tab=None):
    start = time.time()
    if tab is None :
        if action == 'pull':
            macro_pull_all(xl_path)
        if action == 'push':
            macro_push_all(xl_path)
    else:
        if action == 'pull':
            macro_pull_tab(xl_path, tab)
        elif action == 'push':
            macro_push_tab(xl_path, tab)

    print('Operation done in %s seconds' %(time.time() - start))
