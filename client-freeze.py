#!/usr/bin/env python

from cx_Freeze import setup, Executable
from os import getcwd, path
from requests import certs
from shutil import rmtree
from sys import argv, platform


def SetupBuild():

    app = 'analytics'
    base = 'Win32GUI' if platform == 'win32' else None
    curpath = path.dirname(path.realpath(argv[0]))
    filpath = path.join(curpath, 'Others')
    apppath = path.join(curpath, app)

    build_options = {'include_files': [(certs.where(), 'cacerts.txt'), 'Images'], 'include_msvcr': True}

    setup(
        name = app.capitalize(),
        version = '1.0',
        description = 'Google Analytics statistics report generator',
        author = 'Frederico Martins',
        author_email = 'fredericomartins@outlook.com',
        options = {'build_exe': build_options},
        executables = [Executable(apppath + '.py', base = base, icon = filpath + '/icon.ico')])


if path.exists(getcwd() + '/build') == True:
    question = input("A build já está feita, pretende instalar outra vez? [S/N]: ")

    if question.upper() == 'S':
        rmtree(getcwd() + '/build')
        SetupBuild()

    else:
        raise SystemExit


if __name__ == '__main__':		
    SetupBuild()
