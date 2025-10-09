import os
import pandas as pd
from colorama import Fore, Back, Style, init

init()

_projects = pd.read_excel('projects.xlsx')


def print_menu():
    print('\n----- Options ----- ')
    print('1. Open and Activate Projects')
    # print('2. Save a new Project')
    # print('3. Delete Project')
    print('\n(Type "x" to leave...)\n')


def print_projects(_projects: pd.DataFrame):
    for i in _projects[['ID', 'name']].values:
        print(f'{i[0]}. {i[1]}')


def wait_user_input(placeholder: str):
    return input(Fore.BLUE + f'{placeholder} ' + Style.RESET_ALL)


def build_path(id, root = 'C:\\Code\\Python\\'):
    return root + _projects[['ID', 'folder']].query(f'ID == {id}').folder.values[0]


while True:
    print_menu()
    op = wait_user_input('>> ')

    if op.lower() == 'x':
        print(Fore.RED + '\nLeaving...' + Style.RESET_ALL)
        break

    # Options
    if op == '1':
        print(Fore.GREEN + '\nIn what project do you want to work today?\n' + Style.RESET_ALL)
        print_projects(_projects)
        selected_project = wait_user_input('>> ')

        path = build_path(selected_project)
        os.system(f'start cmd /k "cd /d {path} && .\\.venv\\Scripts\\activate"')

        print('Done!')