import os
from time import sleep
import pandas as pd
from colorama import Fore, Back, Style, init

init()

# Variables
projects_sheet = 'projects.xlsx'

_projects = pd.read_excel(projects_sheet)


def print_menu():
    print_header('Options')
    print('1. Open and Activate Projects')
    print('2. Open Projects Sheet')
    # print('3. Delete Project')
    print('\n(Type "x" to leave...)\n')


def print_projects(_projects: pd.DataFrame):
    for i in _projects[['ID', 'name']].values:
        print(f'{i[0]}. {i[1]}')

    print('\n0. Close')


def wait_user_input(placeholder: str):
    return input(Fore.BLUE + f'{placeholder} ' + Style.RESET_ALL)


def build_path(id, root = 'C:\\Code\\Python\\'):
    return root + _projects[['ID', 'folder']].query(f'ID == {id}').folder.values[0]


def print_header(placeholder: str):
    size = len(placeholder)+2
    print('╔' + '═'*size + '╗')
    print('║ ' + Fore.GREEN + placeholder + Style.RESET_ALL + ' ║')
    print('╚' + '═'*size + '╝')


def clear_terminal():
    """Clears the terminal screen."""
    # For Windows
    if os.name == 'nt':
        _ = os.system('cls')
    # For macOS and Linux
    else:
        _ = os.system('clear')


while True:
    clear_terminal()
    print_menu()
    op = wait_user_input('>> ')

    if op.lower() == 'x':
        print(Fore.RED + '\nLeaving...' + Style.RESET_ALL)
        break

    # Options
    if op == '1':
        # print(Fore.GREEN + '\nIn what project do you want to work today?\n' + Style.RESET_ALL)
        print_header('In what project do you want to work today?')
        print_projects(_projects)
        selected_project = wait_user_input('>> ')

        if selected_project ==  '0':
            continue

        print(Fore.YELLOW + 'Working on it' + Style.RESET_ALL)

        path = build_path(selected_project)
        os.system(f'start cmd /k "cd /d {path} && .\\.venv\\Scripts\\activate"')

        print('Done!')

    if op == '2':
        print(Fore.YELLOW + '\nOpening...' + Style.RESET_ALL)
        os.startfile(projects_sheet)
        sleep(5)